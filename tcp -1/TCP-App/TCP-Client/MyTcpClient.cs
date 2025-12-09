
using System;
using System.Net.Sockets;
using System.Net;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TCP_Client
{
    public class MyTcpClient
    {
        private TcpClient _client;

        public event EventHandler<string> MessageReceived;

        private string _clientId; // Dynamically generated client ID

        /// <summary>
        /// Connect to the TCP server
        /// </summary>
        /// <param name="ipAddress"></param>
        /// <param name="port"></param>
        public void Connect(string ipAddress, int port)
        {
            _client = new TcpClient();
            _client.Connect(ipAddress, port);

            // Generate a dynamic client ID based on local IPv4 (fallback to machine hash)
            _clientId = GenerateClientId();

            // Start a separate task to receive data
            _ = ReceiveDataAsync();
        }

        /// <summary>
        /// Send message to the connected server
        /// </summary>
        /// <param name="message"></param>
        public void SendMessage(string message)
        {
            NetworkStream stream = _client.GetStream();
            byte[] buffer = Encoding.ASCII.GetBytes(message);
            stream.Write(buffer, 0, buffer.Length);
        }

        /// <summary>
        /// Handle incoming data from the server
        /// </summary>
        private async Task ReceiveDataAsync()
        {
            NetworkStream stream = _client.GetStream();
            byte[] buffer = new byte[1024];
            StringBuilder messageBuilder = new StringBuilder();

            while (true)
            {
                int bytesRead = await stream.ReadAsync(buffer, 0, buffer.Length);

                if (bytesRead == 0)
                {
                    break; // Server or client disconnected
                }

                string data = Encoding.ASCII.GetString(buffer, 0, bytesRead);
                messageBuilder.Append(data);

                if (data.Length > 0)
                {
                    string receivedMessage = messageBuilder.ToString();
                    OnMessageReceived(receivedMessage);
                    messageBuilder.Clear();
                }
            }

            _client.Close();
        }

        private static string GenerateClientId()
        {
            try
            {
                // Try get first IPv4 and use its last octet as numeric suffix
                var host = Dns.GetHostEntry(Dns.GetHostName());
                var ipv4 = host.AddressList.FirstOrDefault(a => a.AddressFamily == System.Net.Sockets.AddressFamily.InterNetwork);
                if (ipv4 != null)
                {
                    var octets = ipv4.ToString().Split('.');
                    var last = octets[octets.Length - 1];
                    return  last; 
                }
            }
            catch { }

            // Fallback: derive a stable 3-digit code from machine name
            var name = Environment.MachineName ?? "CLIENT";
            int hash = Math.Abs(name.GetHashCode());
            int suffix = (hash % 900) + 100; // 100-999
            return suffix.ToString();
        }

        public void Disconnect()
        {
            _client?.Close();
        }

        protected virtual void OnMessageReceived(string message)
        {
            // Handle server commands
            if (message.Contains("$GETID#"))
            {
                SendMessage($"${_clientId}#");
            }
            else if (message.Contains("$GETSTATUS#"))
            {
                // Example: $1,0,1,1,1$ (customize as needed)
                SendMessage("$1,1,1,1$"); // All OK, change as needed for demo
            }
            else if (message.Contains(":RESET#"))
            {
                // If reset targeted to this client, acknowledge
                // Pattern expected: $S<id>:RESET#
                int s = message.IndexOf('$');
                int hash = message.IndexOf('#', s + 1);
                if (s >= 0 && hash > s)
                {
                    string cmdCore = message.Substring(s + 1, hash - s - 1);
                    // cmdCore like S123:RESET
                    var parts = cmdCore.Split(':');
                    if (parts.Length == 2 && parts[0].Equals(_clientId, StringComparison.OrdinalIgnoreCase))
                    {
                        SendMessage("$OK#");
                    }
                }
            }
            // Raise event for UI
            MessageReceived?.Invoke(this, message);



        }
    }
}