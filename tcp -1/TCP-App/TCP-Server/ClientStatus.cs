using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

namespace TCP_Server
{
    public partial class ClientStatus : Form
    {
        private Dictionary<string, Panel> _clientPanels;
        private Dictionary<string, Label> _clientLabels;
        private readonly List<string> _connectedClients;
        private readonly object _lock = new object();

        public bool IsClientConnected(string clientId)
        {
            lock (_lock)
            {
                return _connectedClients.Contains(clientId);
            }
        }

        public ClientStatus()
        {
            InitializeComponent();
            _connectedClients = new List<string>();
            InitializeClientControls();
        }

        private void InitializeClientControls()
        {
            try
            {
                _clientPanels = new Dictionary<string, Panel>();
                _clientLabels = new Dictionary<string, Label>();

                // Map panels and labels
                var panels = new[] { panel1, panel2, panel3, panel4, panel5 };
                var labels = new[] { label1, label2, label3, label4, label5 };

                for (int i = 0; i < 5; i++)
                {
                    string clientId = $"C ID{i + 1}";
                    
                    if (i < panels.Length && panels[i] != null && !panels[i].IsDisposed)
                    {
                        _clientPanels[clientId] = panels[i];
                        panels[i].BackColor = Color.Red;
                        panels[i].Paint -= PanelPaint;
                        panels[i].Paint += PanelPaint;
                        Debug.WriteLine($"[Initialize] Panel {i + 1} initialized");
                    }
                    
                    if (i < labels.Length && labels[i] != null && !labels[i].IsDisposed)
                    {
                        _clientLabels[clientId] = labels[i];
                        labels[i].Text = clientId;
                        Debug.WriteLine($"[Initialize] Label {i + 1} initialized");
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[Initialize] Error: {ex.Message}");
            }
        }

        public void UpdateClientStatus(string clientId, bool isConnected)
        {
            if (string.IsNullOrEmpty(clientId)) return;

            if (InvokeRequired)
            {
                BeginInvoke(new Action(() => UpdateClientStatus(clientId, isConnected)));
                return;
            }
            
            // Update Excel with the status change
            var mainForm = Application.OpenForms.OfType<Form1>().FirstOrDefault();
            mainForm?.UpdateExcelStatus(clientId, isConnected ? "Connected" : "Disconnected");

            lock (_lock)
            {
                try
                {
                    // Update connected clients list
                    if (isConnected)
                    {
                        if (!_connectedClients.Contains(clientId))
                        {
                            _connectedClients.Add(clientId);
                            Debug.WriteLine($"[+] Client connected: {clientId}");
                        }
                    }
                    else
                    {
                        _connectedClients.RemoveAll(id => id == clientId);
                        Debug.WriteLine($"[-] Client disconnected: {clientId}");
                    }

                    // Update all panels
                    for (int i = 0; i < 5; i++)
                    {
                        string panelId = $"C ID{i + 1}";
                        bool isActive = i < _connectedClients.Count;

                        // Update panel color
                        if (_clientPanels.TryGetValue(panelId, out var panel) && panel != null && !panel.IsDisposed)
                        {
                            panel.BackColor = isActive ? Color.Green : Color.Red;
                            panel.Invalidate();
                        }

                        // Update label text
                        if (_clientLabels.TryGetValue(panelId, out var label) && label != null && !label.IsDisposed)
                        {
                            label.Text = isActive ? _connectedClients[i] : panelId;
                            label.Invalidate();
                        }
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"[ERROR] UpdateClientStatus: {ex.Message}");
                }
            }
        }

        private void PanelPaint(object sender, PaintEventArgs e)
        {
            try
            {
                if (!(sender is Panel panel) || panel.IsDisposed || e?.Graphics == null)
                    return;

                // Double buffering to prevent flicker
                using (var bufferedGraphics = BufferedGraphicsManager.Current.Allocate(e.Graphics, panel.ClientRectangle))
                {
                    var g = bufferedGraphics.Graphics;
                    g.SmoothingMode = SmoothingMode.AntiAlias;
                    g.Clear(panel.BackColor);

                    // Draw the background
                    using (var brush = new SolidBrush(panel.BackColor))
                    {
                        g.FillRectangle(brush, panel.ClientRectangle);
                    }

                    // Draw the border
                    using (var pen = new Pen(Color.Black, 1))
                    {
                        g.DrawRectangle(pen, 0, 0, panel.Width - 1, panel.Height - 1);
                    }

                    // Render to the screen
                    bufferedGraphics.Render(e.Graphics);
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[ERROR] PanelPaint: {ex.Message}");
            }
        }

        public void UpdateClientStatusById(string clientId, bool isConnected)
        {
            UpdateClientStatus(clientId, isConnected);
        }

        public void ResetClientStatus(string clientId)
        {
            UpdateClientStatus(clientId, false);
        }
    }
}
