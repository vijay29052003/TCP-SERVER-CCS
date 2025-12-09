using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Linq; // Added for .ToList()

namespace TCP_Server
{
    public class MyTcpServer : IDisposable
    {
        private TcpListener _server;
        private readonly string _ipAddress;
        private readonly int _port;
        private readonly ConcurrentDictionary<Guid, TcpClient> _connectedClients = new ConcurrentDictionary<Guid, TcpClient>();
        private bool _isRunning;

        // Mapping from Guid to reported client ID
        private readonly ConcurrentDictionary<Guid, string> _clientIdMap = new ConcurrentDictionary<Guid, string>();
        private readonly ConcurrentDictionary<string, Guid> _idToGuidMap = new ConcurrentDictionary<string, Guid>();
        
        // Get all known client IDs (as reported by clients)
        public List<string> GetKnownClientIds()
        {
            return _idToGuidMap.Keys.ToList();
        }

        // Send a message to a specific client by reported client ID
        public async Task SendToClientByIdAsync(string clientId, string message)
        {
            if (_idToGuidMap.TryGetValue(clientId, out var guid) && _connectedClients.TryGetValue(guid, out var client))
            {
                var messageBytes = Encoding.ASCII.GetBytes(message);
                try
                {
                    var stream = client.GetStream();
                    await stream.WriteAsync(messageBytes, 0, messageBytes.Length);
                    OnMessageReceived($"Message sent to {clientId}: {message}");
                }
                catch (Exception ex)
                {
                    OnMessageReceived($"Failed to send to {clientId}: {ex.Message}");
                }
            }
            else
            {
                OnMessageReceived($"Client {clientId} not found.");
            }
        }

        public event EventHandler<string> MessageReceived;
        public event EventHandler<string> ClientMessageReceived;

        public int Port => _port;

        /// <summary>
        /// TCP Server helper class
        /// </summary>
        /// <param name="ipAddress">Default at loopback IP (Cannot be accessed from the outside world)</param>
        /// <param name="port"></param>
        public MyTcpServer(string ipAddress = "127.0.0.1", int port = 8083)
        {
            _ipAddress = ipAddress;
            _port = port;
        }

        public Task StartAsync(CancellationToken cancellationToken = default)
        {
            // Abort operation if server is already running
            if (_isRunning)
            {
                throw new InvalidOperationException("Server is already running!");
            }

            var tcs = new TaskCompletionSource<bool>();
            
            try
            {
                IPAddress localAddr = IPAddress.Parse(_ipAddress);
                _server = new TcpListener(localAddr, _port);
                _server.Start();
                _isRunning = true;

                OnMessageReceived($"Server started on {_ipAddress}:{_port}. Waiting for connections...");
                tcs.SetResult(true);

                // Start accepting clients in a background task
                _ = Task.Run(async () =>
                {
                    while (_isRunning && !cancellationToken.IsCancellationRequested)
                    {
                        try
                        {
                            var client = await _server.AcceptTcpClientAsync().ConfigureAwait(false);
                            var clientId = Guid.NewGuid();
                            _connectedClients[clientId] = client;
                            var clientEndpoint = client.Client?.RemoteEndPoint?.ToString() ?? "[unknown]";
                            OnMessageReceived($"Client connected: {clientEndpoint} (ID: {clientId})");
                            // Start a new task to handle the client without awaiting
                            _ = Task.Run(() => HandleClientAsync(client, clientId), cancellationToken);
                        }
                        catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
                        {
                            // Server is stopping
                            break;
                        }
                        catch (ObjectDisposedException) when (!_isRunning)
                        {
                            // Server is stopping
                            break;
                        }
                        catch (Exception ex)
                        {
                            OnMessageReceived($"Error accepting client: {ex.Message}");
                            // Small delay to prevent tight loop on error
                            try
                            {
                                await Task.Delay(1000, cancellationToken).ConfigureAwait(false);
                            }
                            catch (TaskCanceledException)
                            {
                                break;
                            }
                        }
                    }
                }, cancellationToken);

                return tcs.Task;
            }
            catch (Exception ex)
            {
                _isRunning = false;
                _server?.Stop();
                var errorMessage = $"Failed to start server: {ex.Message}";
                OnMessageReceived(errorMessage);
                tcs.SetException(new InvalidOperationException(errorMessage, ex));
                return tcs.Task;
            }
        }

        [Obsolete("Use StartAsync() instead.")]
        public Task Start()
        {
            return StartAsync(CancellationToken.None);
        }

        public void Stop()
        {
            _isRunning = false;
            try
            {
                // Stop accepting new connections
                _server?.Stop();

                // Close all connected clients
                foreach (var client in _connectedClients.Values)
                {
                    try { client.Close(); } catch { /* Ignore */ }
                }
                _connectedClients.Clear();

                OnMessageReceived("Server stopped.");
            }
            catch (Exception ex)
            {
                OnMessageReceived($"Error stopping server: {ex.Message}");
            }
        }

        public async Task BroadcastMessageAsync(string message)
        {
            if (string.IsNullOrEmpty(message)) return;

            var messageBytes = Encoding.ASCII.GetBytes(message);
            var disconnectedClients = new List<Guid>();

            foreach (var clientPair in _connectedClients)
            {
                try
                {
                    var stream = clientPair.Value.GetStream();
                    await stream.WriteAsync(messageBytes, 0, messageBytes.Length);
                }
                catch
                {
                    disconnectedClients.Add(clientPair.Key);
                }
            }

            // Clean up disconnected clients
            foreach (var clientId in disconnectedClients)
            {
                if (_connectedClients.TryRemove(clientId, out var client))
                {
                    try { client.Close(); } catch { /* Ignore */ }
                    OnMessageReceived($"Client {clientId} disconnected.");
                }
            }
        }

        public void Dispose()
        {
            Stop();
            _server = null;
        }

        private async Task HandleClientAsync(TcpClient client, Guid clientId)
        {
            try
            {
                var stream = client.GetStream();
                var buffer = new byte[1024];
                string reportedClientId = null;

                while (_isRunning && client.Connected)
                {
                    var bytesRead = await stream.ReadAsync(buffer, 0, buffer.Length);
                    if (bytesRead == 0) break; // Client disconnected

                    string message = Encoding.ASCII.GetString(buffer, 0, bytesRead);
                    // Check for $Sxxx# (client ID response)
                    if (message.StartsWith("$") && message.EndsWith("#") && message.Length > 3)
                    {
                        reportedClientId = message.Trim('$', '#', '\r', '\n');
                        _clientIdMap[clientId] = reportedClientId;
                        _idToGuidMap[reportedClientId] = clientId;
                        OnMessageReceived($"Mapped client {clientId} to reported ID {reportedClientId}");
                    }
                    OnClientMessageReceived($"[Client {clientId}] {message}");
                }
            }
            catch (Exception ex) when (ex is ObjectDisposedException || ex is InvalidOperationException)
            {
                // Expected during shutdown or client disconnect
            }
            catch (Exception ex)
            {
                OnMessageReceived($"Error with client {clientId}: {ex.Message}");
            }
            finally
            {
                if (_clientIdMap.TryRemove(clientId, out var repId))
                {
                    _idToGuidMap.TryRemove(repId, out _);
                }
                if (_connectedClients.TryRemove(clientId, out var _))
                {
                    try { client.Close(); } catch { /* Ignore */ }
                    OnMessageReceived($"Client {clientId} disconnected.");
                }
            }
        }

        public async Task SendToClientAsync(string message)
        {
            if (string.IsNullOrEmpty(message) || _connectedClients.Count == 0)
            {
                OnMessageReceived("No clients connected to send message to.");
                return;
            }

            var messageBytes = Encoding.ASCII.GetBytes(message);
            var disconnectedClients = new List<Guid>();
            int successfulSends = 0;

            foreach (var clientPair in _connectedClients)
            {
                try
                {
                    var stream = clientPair.Value.GetStream();
                    await stream.WriteAsync(messageBytes, 0, messageBytes.Length);
                    successfulSends++;
                }
                catch
                {
                    disconnectedClients.Add(clientPair.Key);
                }
                finally
                {
                    //
                }
            }

            // Clean up disconnected clients
            foreach (var clientId in disconnectedClients)
            {
                if (_connectedClients.TryRemove(clientId, out var client))
                {
                    try { client.Close(); } catch { /* Ignore */ }
                    OnMessageReceived($"Client {clientId} disconnected.");
                }
            }

            if (successfulSends > 0)
            {
                OnMessageReceived($"Message sent to {successfulSends} client(s): {message}");
            }
        }

        /// <summary>
        /// Send message to the UI thread
        /// </summary>
        /// <param name="message"></param>
        protected virtual void OnMessageReceived(string message)
        {
            MessageReceived?.Invoke(this, message);
        }

        protected virtual void OnClientMessageReceived(string message)
        {
            ClientMessageReceived?.Invoke(this, message);
        }

        internal object StopAsync()
        {
            throw new NotImplementedException();
        }
    }
}