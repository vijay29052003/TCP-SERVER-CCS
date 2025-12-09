using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Forms;
using TCP_Server;
using Excel = Microsoft.Office.Interop.Excel;

namespace TCP_Server
{
    public partial class Form1 : Form
    {
        private readonly MyTcpServer _server;
        private bool _serverConnected;

        private readonly HashSet<string> _knownClientIds = new HashSet<string>();
        private string _lastSelectedClientId;

        private readonly Button _resetButton;

        private readonly string _excelFilePath;
        private Excel.Application _excelApp;
        private Excel.Workbook _workbook;
        private Excel.Worksheet _worksheet;
        private bool _disposed;

        // ========= PictureBoxes for connected client status =========
        private readonly PictureBox[] _clientIndicators;
        // Mapping: clientId -> pictureBox index (0..4)
        private readonly Dictionary<string, int> _clientIndicatorMap = new Dictionary<string, int>();
        // ============================================================

        public Form1()
        {
            try
            {
                _excelFilePath = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                    "tcp status.xlsx");

                LogToFile("Initializing Form1 components...");
                InitializeComponent();

                // ====== initialize pictureBox1..pictureBox5 as status lights ======
                _clientIndicators = new PictureBox[]
                {
                    pictureBox1,
                    pictureBox2,
                    pictureBox3,
                    pictureBox4,
                    pictureBox5
                };

                if (_clientIndicators != null)
                {
                    foreach (var pb in _clientIndicators)
                    {
                        if (pb != null)
                        {
                            pb.BackColor = Color.Red;                 // default: disconnected
                            pb.BorderStyle = BorderStyle.FixedSingle; // just for visibility
                        }
                    }
                }
                // ===================================================================

                LogToFile("Creating TCP server instance...");
                _server = new MyTcpServer(IPAddress.Any.ToString());
                _server.MessageReceived += Server_MessageReceived;
                _server.ClientMessageReceived += Server_ClientMessageReceived;

                var found = Controls.Find("resetButton", true);
                if (found.Length > 0 && found[0] is Button btn)
                {
                    _resetButton = btn;
                    _resetButton.Click += ResetButton_Click;
                }

                comboBox1.SelectedIndexChanged += ComboBox1_SelectedIndexChanged;

                if (button2 != null)
                {
                    button2.Click += Button2_Click; // GET status
                }

                _ = Task.Run(async () =>
                {
                    try
                    {
                        await InitializeExcelAppAsync();
                        LogToFile("Excel initialization completed successfully.");
                    }
                    catch (Exception ex)
                    {
                        LogToFile($"Error initializing Excel: {ex.Message}");
                    }
                });

                LogToFile("Form1 initialization completed successfully.");
            }
            catch (Exception ex)
            {
                string error = $"Error initializing Form1: {ex}";
                LogToFile(error);
                MessageBox.Show(error, "Initialization Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                throw;
            }
        }

        private enum LogLevel
        {
            INFO,
            DEBUG,
            WARNING,
            ERROR
        }

        private void LogToFile(string message,
                               LogLevel level = LogLevel.INFO,
                               [CallerMemberName] string methodName = "",
                               [CallerLineNumber] int lineNumber = 0)
        {
            try
            {
                string timestamp = DateTime.Now.ToString("HH:mm:ss.fff");
                string levelStr = level.ToString().PadRight(7);
                string source = $"{methodName}:{lineNumber}";
                message = message?.Trim() ?? string.Empty;

                string logMessage = $"[{timestamp}] [{levelStr}] [{source,-20}] {message}";

                ConsoleColor originalColor = Console.ForegroundColor;
                switch (level)
                {
                    case LogLevel.ERROR:
                        Console.ForegroundColor = ConsoleColor.Red;
                        break;
                    case LogLevel.WARNING:
                        Console.ForegroundColor = ConsoleColor.Yellow;
                        break;
                    case LogLevel.DEBUG:
                        Console.ForegroundColor = ConsoleColor.Cyan;
                        break;
                    default:
                        Console.ForegroundColor = ConsoleColor.White;
                        break;
                }

                Console.WriteLine(logMessage);
                Console.ForegroundColor = originalColor;

                string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                string logFile = Path.Combine(desktopPath, "tcp_server.log");

                if (File.Exists(logFile) && new FileInfo(logFile).Length > 5 * 1024 * 1024)
                {
                    string backupFile = Path.Combine(desktopPath, $"tcp_server_{DateTime.Now:yyyyMMdd_HHmmss}.log");
                    File.Move(logFile, backupFile);
                }

                File.AppendAllText(logFile, logMessage + Environment.NewLine);
                Debug.WriteLine(logMessage);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[{DateTime.Now:HH:mm:ss.fff}] [ERROR] Failed to write log: {ex.Message}");
            }
        }

        private void LogInfo(string message, [CallerMemberName] string methodName = "", [CallerLineNumber] int lineNumber = 0)
            => LogToFile(message, LogLevel.INFO, methodName, lineNumber);

        private void LogDebug(string message, [CallerMemberName] string methodName = "", [CallerLineNumber] int lineNumber = 0)
            => LogToFile(message, LogLevel.DEBUG, methodName, lineNumber);

        private void LogWarning(string message, [CallerMemberName] string methodName = "", [CallerLineNumber] int lineNumber = 0)
            => LogToFile(message, LogLevel.WARNING, methodName, lineNumber);

        private void LogError(string message, [CallerMemberName] string methodName = "", [CallerLineNumber] int lineNumber = 0)
            => LogToFile(message, LogLevel.ERROR, methodName, lineNumber);

        private void LogMessage(string message, [CallerMemberName] string methodName = "")
            => LogInfo(message, methodName);

        private async Task<bool> InitializeExcelAppAsync()
        {
            if (!string.IsNullOrWhiteSpace(_excelFilePath) && _workbook != null && _excelApp != null)
                return true;

            try
            {
                LogToFile("InitializeExcelAppAsync: attempting to attach to running Excel instance...");

                try
                {
                    object runningExcel = null;
                    try
                    {
                        runningExcel = Marshal.GetActiveObject("Excel.Application");
                    }
                    catch (COMException)
                    {
                        runningExcel = null;
                    }

                    if (runningExcel != null)
                    {
                        _excelApp = runningExcel as Excel.Application;
                        LogToFile("Attached to running Excel instance.");
                    }
                }
                catch (Exception ex)
                {
                    LogToFile($"Attach to running Excel failed: {ex.Message}");
                }

                if (_excelApp == null)
                {
                    _excelApp = new Excel.Application
                    {
                        DisplayAlerts = false,
                        Visible = false
                    };
                    LogToFile("Created new Excel.Application instance.");
                }

                if (File.Exists(_excelFilePath))
                {
                    LogToFile($"Workbook exists. Trying to open workbook: {_excelFilePath}");

                    bool opened = false;
                    Excel.Workbook workbook = null;

                    try
                    {
                        workbook = _excelApp.Workbooks.Open(_excelFilePath, ReadOnly: false);
                        opened = true;
                        LogToFile("Opened workbook in read/write mode.");
                    }
                    catch (Exception ex)
                    {
                        LogToFile($"Could not open workbook read/write: {ex.Message}. Trying read-only open.");
                        try
                        {
                            workbook = _excelApp.Workbooks.Open(_excelFilePath, ReadOnly: true);
                            opened = true;
                            LogToFile("Opened workbook in read-only mode (file may be locked).");
                        }
                        catch (Exception ex2)
                        {
                            LogToFile($"Failed to open workbook even in read-only mode: {ex2.Message}");
                        }
                    }

                    if (opened && workbook != null)
                        _workbook = workbook;
                }
                else
                {
                    LogToFile("Workbook does not exist. Creating new workbook.");
                    _workbook = _excelApp.Workbooks.Add();
                    await Task.Run(() => _workbook.SaveAs(_excelFilePath)).ConfigureAwait(false);
                    LogToFile($"Created and saved new workbook at {_excelFilePath}");
                }

                if (_workbook != null)
                {
                    try
                    {
                        Excel.Worksheet ws = null;
                        try
                        {
                            ws = _workbook.Worksheets.Cast<Excel.Worksheet>()
                                .FirstOrDefault(s => s.Name.Equals("Client Status", StringComparison.OrdinalIgnoreCase));
                        }
                        catch
                        {
                            ws = null;
                        }

                        if (ws == null)
                        {
                            ws = _workbook.Worksheets.Count >= 1
                                ? (Excel.Worksheet)_workbook.Worksheets[1]
                                : (Excel.Worksheet)_workbook.Worksheets.Add();

                            try
                            {
                                ws.Name = "Client Status";
                            }
                            catch
                            {
                                // ignore if name already used
                            }
                        }

                        _worksheet = ws;

                        if (_worksheet.Cells[1, 1].Value2 == null)
                        {
                            InitializeExcelWorksheet(_worksheet);
                            try
                            {
                                _workbook.Save();
                            }
                            catch (Exception ex)
                            {
                                LogToFile($"Save after initializing worksheet failed: {ex.Message}");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        LogToFile($"Error preparing worksheet: {ex.Message}");
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                LogToFile($"Error initializing Excel: {ex.Message}");
                CleanupExcel();
                return false;
            }
        }

        private void CleanupExcel()
        {
            if (_disposed) return;

            try
            {
                if (_workbook != null)
                {
                    try
                    {
                        if (!_workbook.Saved)
                        {
                            try
                            {
                                _workbook.Save();
                            }
                            catch (Exception ex)
                            {
                                LogToFile($"Workbook save during cleanup failed: {ex.Message}");
                            }
                        }
                        _workbook.Close(false);
                        Marshal.ReleaseComObject(_workbook);
                    }
                    finally
                    {
                        _workbook = null;
                    }
                }

                if (_excelApp != null)
                {
                    try
                    {
                        _excelApp.Quit();
                        Marshal.ReleaseComObject(_excelApp);
                    }
                    finally
                    {
                        _excelApp = null;
                    }
                }
            }
            catch (Exception ex)
            {
                LogError($"Error during Excel cleanup: {ex.Message}");
            }
            finally
            {
                _disposed = true;
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        private void InitializeExcelWorksheet(Excel.Worksheet worksheet)
        {
            try
            {
                Excel.Range titleRange = worksheet.Range["A1"];
                titleRange.Value = "CLIENT STATUS";
                titleRange.Font.Bold = true;
                titleRange.Font.Size = 16;
                titleRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                worksheet.Range["A1:F1"].Merge();

                string[] headers = { "S.NO", "CLIENT ID", "RTC STATUS", "RS485", "GPS", "FLASH" };
                for (int i = 0; i < headers.Length; i++)
                {
                    worksheet.Cells[2, i + 1] = headers[i];
                }

                Excel.Range headerRange = worksheet.Range["A2:F2"];
                headerRange.Font.Bold = true;
                headerRange.Interior.Color = ColorTranslator.ToOle(Color.LightGray);
                headerRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                worksheet.Columns.AutoFit();
            }
            catch (Exception ex)
            {
                LogToFile($"InitializeExcelWorksheet error: {ex.Message}");
            }
        }

        private int GetNextAvailableRow(Excel.Worksheet worksheet)
        {
            if (worksheet == null)
            {
                LogToFile("GetNextAvailableRow: worksheet is null, returning default row 3.");
                return 3;
            }

            int row = 3;
            try
            {
                while (worksheet.Cells[row, 1].Value2 != null)
                    row++;
            }
            catch
            {
                // ignore COM errors
            }
            return row;
        }

        private async Task UpdateExcelWithCurrentStatusAsync()
        {
            if (string.IsNullOrEmpty(_excelFilePath))
            {
                LogToFile("Excel file path is not set");
                return;
            }

            string clientId = "";
            string rtcStatus = "";
            string rs485 = "";
            string gps = "";
            string flash = "";

            if (textBox1.InvokeRequired)
            {
                textBox1.Invoke((MethodInvoker)delegate
                {
                    clientId = textBox1.Text;
                    rtcStatus = textBox2.Text;
                    rs485 = textBox3.Text;
                    gps = textBox4.Text;
                    flash = textBox5.Text;
                });
            }
            else
            {
                clientId = textBox1.Text;
                rtcStatus = textBox2.Text;
                rs485 = textBox3.Text;
                gps = textBox4.Text;
                flash = textBox5.Text;
            }

            await Task.Run(() =>
            {
                Excel.Application localExcel = null;
                Excel.Workbook localWb = null;
                Excel.Worksheet localWs = null;
                bool localExcelCreated = false;
                bool openedReadOnly = false;
                string tempFile = null;

                try
                {
                    LogToFile("UpdateExcelWithCurrentStatusAsync: starting.");

                    try
                    {
                        object runningExcel = null;
                        try
                        {
                            runningExcel = Marshal.GetActiveObject("Excel.Application");
                        }
                        catch (COMException)
                        {
                            runningExcel = null;
                        }

                        if (runningExcel != null)
                        {
                            localExcel = runningExcel as Excel.Application;
                            LogToFile("Attached to running Excel for update.");
                        }
                    }
                    catch (Exception ex)
                    {
                        LogToFile($"Attach to running Excel failed: {ex.Message}");
                    }

                    if (localExcel == null)
                    {
                        localExcel = new Excel.Application { DisplayAlerts = false, Visible = false };
                        localExcelCreated = true;
                        LogToFile("Created new Excel instance for update.");
                    }

                    bool fileExists = File.Exists(_excelFilePath);

                    if (fileExists)
                    {
                        try
                        {
                            localWb = localExcel.Workbooks.Open(_excelFilePath, ReadOnly: false);
                            LogToFile("Opened workbook read/write.");
                        }
                        catch (Exception exOpen)
                        {
                            LogToFile($"Open in read/write failed: {exOpen.Message}. Trying read-only open.");
                            try
                            {
                                localWb = localExcel.Workbooks.Open(_excelFilePath, ReadOnly: true);
                                openedReadOnly = true;
                                LogToFile("Opened workbook read-only (file is locked by another process).");
                            }
                            catch (Exception exOpen2)
                            {
                                LogToFile($"Open read-only also failed: {exOpen2.Message}. Will create temp workbook.");
                                localWb = localExcel.Workbooks.Add();
                            }
                        }
                    }
                    else
                    {
                        localWb = localExcel.Workbooks.Add();
                        LogToFile("Workbook did not exist; created new workbook.");
                    }

                    try
                    {
                        localWs = null;
                        foreach (Excel.Worksheet sh in localWb.Worksheets)
                        {
                            try
                            {
                                if (string.Equals(sh.Name, "Client Status", StringComparison.OrdinalIgnoreCase))
                                {
                                    localWs = sh;
                                    break;
                                }
                            }
                            catch { }
                        }

                        if (localWs == null)
                        {
                            localWs = localWb.Worksheets.Count >= 1
                                ? (Excel.Worksheet)localWb.Worksheets[1]
                                : (Excel.Worksheet)localWb.Worksheets.Add();

                            try
                            {
                                localWs.Name = "Client Status";
                            }
                            catch { }

                            InitializeExcelWorksheet(localWs);
                        }
                    }
                    catch (Exception ex)
                    {
                        LogToFile($"Error locating/creating worksheet: {ex.Message}");
                    }

                    if (localWs == null)
                    {
                        LogToFile("UpdateExcelWithCurrentStatusAsync: worksheet is null, skipping row write.");
                    }
                    else
                    {
                        int lastRow = 3;
                        try
                        {
                            var rowsCount = localWs.Rows.Count;
                            var lastCell = localWs.Cells[rowsCount, 1].End(Excel.XlDirection.xlUp);
                            int candidate = 2;
                            try
                            {
                                candidate = lastCell != null ? lastCell.Row : 2;
                            }
                            catch
                            {
                                candidate = 2;
                            }

                            lastRow = Math.Max(candidate + 1, 3);
                        }
                        catch (Exception ex)
                        {
                            LogToFile($"End(xlUp) detection failed, falling back to GetNextAvailableRow: {ex.Message}");
                            try
                            {
                                lastRow = GetNextAvailableRow(localWs);
                            }
                            catch
                            {
                                lastRow = 3;
                            }
                        }

                        int rowNum = lastRow - 2;

                        try
                        {
                            localWs.Cells[lastRow, 1] = rowNum;
                            localWs.Cells[lastRow, 2] = clientId;
                            localWs.Cells[lastRow, 3] = rtcStatus;
                            localWs.Cells[lastRow, 4] = rs485;
                            localWs.Cells[lastRow, 5] = gps;
                            localWs.Cells[lastRow, 6] = flash;

                            Excel.Range dataRange = localWs.Range[$"A{lastRow}:F{lastRow}"];
                            dataRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                            dataRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        }
                        catch (Exception ex)
                        {
                            LogToFile($"Error writing data to worksheet: {ex.Message}");
                        }
                    }

                    if (!fileExists)
                    {
                        try
                        {
                            localWb.SaveAs(_excelFilePath);
                            LogToFile($"Saved new workbook to {_excelFilePath}");
                        }
                        catch (Exception ex)
                        {
                            LogToFile($"SaveAs failed for new workbook: {ex.Message}");
                        }
                    }
                    else
                    {
                        if (!openedReadOnly)
                        {
                            try
                            {
                                localWb.Save();
                                LogToFile("Workbook saved successfully.");
                            }
                            catch (Exception ex)
                            {
                                LogToFile($"Workbook.Save failed: {ex.Message}. Will try temp-save-and-copy method.");
                                openedReadOnly = true;
                            }
                        }

                        if (openedReadOnly)
                        {
                            try
                            {
                                tempFile = Path.Combine(
                                    Path.GetDirectoryName(_excelFilePath),
                                    $"temp_{DateTime.Now:yyyyMMddHHmmss}_{Path.GetFileName(_excelFilePath)}");

                                localWb.SaveAs(tempFile);
                                LogToFile($"Saved workbook to temp file: {tempFile}");

                                try
                                {
                                    File.Copy(tempFile, _excelFilePath, true);
                                    LogToFile("Copied temp file over original file successfully.");
                                    File.Delete(tempFile);
                                }
                                catch (Exception exCopy)
                                {
                                    LogToFile($"Could not copy temp file over original: {exCopy.Message}. Temp file kept at: {tempFile}");
                                }
                            }
                            catch (Exception exTemp)
                            {
                                LogToFile($"Temp SaveAs failed: {exTemp.Message}");
                            }
                        }
                    }

                    try
                    {
                        if (!localExcelCreated && localExcel != null)
                        {
                            try { localExcel.ScreenUpdating = true; } catch { }
                            try { localExcel.Calculate(); } catch { }

                            try
                            {
                                var win = localExcel.ActiveWindow;
                                if (win != null)
                                    win.ScrollRow = 3;
                            }
                            catch { }
                        }
                    }
                    catch (Exception ex)
                    {
                        LogToFile($"UI refresh attempt failed: {ex.Message}");
                    }

                    if (localExcelCreated)
                    {
                        try { localWb.Close(false); } catch { }
                        try { localExcel.Quit(); } catch { }
                    }

                    LogToFile("UpdateExcelWithCurrentStatusAsync: finished successfully.");
                }
                catch (Exception ex)
                {
                    LogToFile($"Error in UpdateExcelWithCurrentStatusAsync: {ex}");
                }
                finally
                {
                    try
                    {
                        if (localWs != null) Marshal.ReleaseComObject(localWs);
                    }
                    catch { }
                    try
                    {
                        if (localWb != null) Marshal.ReleaseComObject(localWb);
                    }
                    catch { }
                    try
                    {
                        if (localExcel != null && localExcel != _excelApp) Marshal.ReleaseComObject(localExcel);
                    }
                    catch { }

                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }
            });
        }

        private void SelectClientFromMessage(string message)
        {
            if (message == null) return;

            int dollarIndex = message.IndexOf('$');
            if (dollarIndex < 0) return;

            int hashIndex = message.IndexOf('#', dollarIndex + 1);
            if (hashIndex <= dollarIndex + 1) return;

            string clientId = message.Substring(dollarIndex + 1, hashIndex - dollarIndex - 1);
            if (clientId.Length == 0) return;

            if (comboBox1.InvokeRequired)
            {
                comboBox1.BeginInvoke(new Action<string>(SelectClientById), clientId);
            }
            else
            {
                SelectClientById(clientId);
            }
        }

        private void SelectClientById(string clientId)
        {
            if (string.IsNullOrWhiteSpace(clientId)) return;

            if (!_knownClientIds.Contains(clientId))
                _knownClientIds.Add(clientId);

            if (!comboBox1.Items.Contains(clientId))
                comboBox1.Items.Insert(0, clientId);

            comboBox1.SelectedItem = clientId;
            _lastSelectedClientId = clientId;
        }

        // ===================== STABLE INDICATOR MAPPING =====================

        /// <summary>
        /// Finds the first free picture box slot index (0..N-1) not used in _clientIndicatorMap.
        /// Returns -1 if all slots are occupied.
        /// </summary>
        private int FindFreeIndicatorSlot()
        {
            if (_clientIndicators == null || _clientIndicators.Length == 0)
                return -1;

            for (int i = 0; i < _clientIndicators.Length; i++)
            {
                if (!_clientIndicatorMap.ContainsValue(i))
                    return i;
            }
            return -1;
        }

        /// <summary>
        /// Update known clients and drive picture boxes WITHOUT reordering.
        /// Each client gets a fixed indicator slot until it disconnects.
        /// </summary>
        private void UpdateKnownClients(string clientId, bool isConnected)
        {
            LogToFile($"Updating known clients - Client: {clientId}, Connected: {isConnected}");
            try
            {
                if (InvokeRequired)
                {
                    BeginInvoke(new Action(() => UpdateKnownClients(clientId, isConnected)));
                    return;
                }

                if (isConnected)
                {
                    // Update set used for tracking / combo, if needed
                    if (!_knownClientIds.Contains(clientId))
                    {
                        _knownClientIds.Add(clientId);
                        Debug.WriteLine($"[+] Added client to known clients: {clientId}");
                    }

                    // Assign or re-use a slot
                    int slotIndex;
                    if (!_clientIndicatorMap.TryGetValue(clientId, out slotIndex))
                    {
                        slotIndex = FindFreeIndicatorSlot();
                        if (slotIndex >= 0)
                        {
                            _clientIndicatorMap[clientId] = slotIndex;
                            Debug.WriteLine($"[+] Mapped client {clientId} -> indicator slot {slotIndex}");
                        }
                        else
                        {
                            Debug.WriteLine($"[WARN] No free indicator slot for client {clientId}.");
                        }
                    }

                    // Turn GREEN for this client
                    if (slotIndex >= 0 && _clientIndicators != null &&
                        slotIndex < _clientIndicators.Length &&
                        _clientIndicators[slotIndex] != null)
                    {
                        _clientIndicators[slotIndex].BackColor = Color.LimeGreen;
                    }
                }
                else
                {
                    // Disconnect path
                    if (_knownClientIds.Contains(clientId))
                    {
                        _knownClientIds.Remove(clientId);
                        Debug.WriteLine($"[-] Removed client from known clients: {clientId}");
                    }

                    int slotIndex;
                    if (_clientIndicatorMap.TryGetValue(clientId, out slotIndex))
                    {
                        // Turn RED ONLY for this client
                        if (_clientIndicators != null &&
                            slotIndex < _clientIndicators.Length &&
                            _clientIndicators[slotIndex] != null)
                        {
                            _clientIndicators[slotIndex].BackColor = Color.Red;
                        }

                        _clientIndicatorMap.Remove(clientId);
                        Debug.WriteLine($"[-] Freed indicator slot {slotIndex} from client {clientId}");
                    }
                }

                UpdateClientComboBox();
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[ERROR] UpdateKnownClients failed: {ex}");
            }
        }

        // ====================================================================

        private void UpdateClientComboBox()
        {
            LogToFile("Updating client combo box with known client IDs");
            var ids = _server.GetKnownClientIds();

            comboBox1.Invoke((MethodInvoker)(() =>
            {
                var previous = comboBox1.SelectedItem as string;

                comboBox1.Items.Clear();
                foreach (var id in ids)
                {
                    comboBox1.Items.Add(id);
                }

                if (!string.IsNullOrEmpty(previous) && comboBox1.Items.Contains(previous))
                {
                    comboBox1.SelectedItem = previous;
                    _lastSelectedClientId = previous;
                }
                else if (comboBox1.Items.Count > 0 && comboBox1.SelectedIndex == -1)
                {
                    comboBox1.SelectedIndex = 0;
                    _lastSelectedClientId = comboBox1.SelectedItem?.ToString();
                }
            }));
        }

        // helper to extract GUID from "Client connected: ... (ID: xxxx)"
        private string ExtractGuidFromMessage(string message)
        {
            try
            {
                int idIndex = message.IndexOf("(ID:", StringComparison.OrdinalIgnoreCase);
                if (idIndex < 0) return null;
                idIndex += 4; // skip "(ID:"
                int end = message.IndexOf(')', idIndex);
                if (end < 0) end = message.Length;
                string guid = message.Substring(idIndex, end - idIndex).Trim();
                return guid;
            }
            catch
            {
                return null;
            }
        }

        // helper to extract GUID from "Client <guid> disconnected."
        private string ExtractGuidFromDisconnectMessage(string message)
        {
            try
            {
                const string prefix = "Client ";
                const string middle = " disconnected";

                int start = message.IndexOf(prefix, StringComparison.OrdinalIgnoreCase);
                if (start < 0) return null;
                start += prefix.Length;

                int end = message.IndexOf(middle, start, StringComparison.OrdinalIgnoreCase);
                if (end < 0) end = message.Length;

                string guid = message.Substring(start, end - start).Trim().TrimEnd('.');
                return guid;
            }
            catch
            {
                return null;
            }
        }

        private void Server_MessageReceived(object sender, string message)
        {
            LogMessage($"[Server] {message}");

            try
            {
                if (string.IsNullOrEmpty(message))
                    return;

                // CONNECT pattern (older form): "Client connected: ... (ID: <guid>)"
                if (message.IndexOf("Client connected:", StringComparison.OrdinalIgnoreCase) >= 0 &&
                    message.IndexOf("(ID:", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    string guid = ExtractGuidFromMessage(message);
                    if (!string.IsNullOrEmpty(guid))
                    {
                        UpdateKnownClients(guid, true);
                    }
                }

                // DISCONNECT pattern (your logs): "Client <guid> disconnected."
                if (message.IndexOf("disconnected", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    string guid = null;

                    if (message.IndexOf("(ID:", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        guid = ExtractGuidFromMessage(message);
                    }
                    else
                    {
                        guid = ExtractGuidFromDisconnectMessage(message);
                    }

                    if (!string.IsNullOrEmpty(guid))
                    {
                        UpdateKnownClients(guid, false);
                    }
                }
            }
            catch (Exception ex)
            {
                LogToFile($"Error handling Server_MessageReceived for indicators: {ex.Message}");
            }
        }

        private async void Server_ClientMessageReceived(object sender, string message)
        {
            try
            {
                if (IsDisposed || !IsHandleCreated || string.IsNullOrEmpty(message))
                {
                    LogToFile("Invalid message received - form disposed, handle not created, or empty message");
                    return;
                }

                SelectClientFromMessage(message);

                string clientInfo = sender != null ? $" from {sender}" : "";
                string safeMessage = message.Length > 500 ? message.Substring(0, 500) + "..." : message;
                LogMessage($"Received message{clientInfo}: {safeMessage}");

                LogToFile($"Message details - Length: {message.Length}, Contains $: {message.Contains("$")}, Contains #: {message.Contains("#")}");

                // old "Client ID:" text detection left here in case server uses it
                if (message.Contains("connected") && message.Contains("Client ID:"))
                {
                    await Task.Run(() =>
                    {
                        try
                        {
                            int start = message.IndexOf("Client ID:", StringComparison.OrdinalIgnoreCase);
                            if (start >= 0)
                            {
                                start += 10;
                                int end = message.IndexOf(' ', start);
                                if (end < 0) end = message.Length;

                                string clientId = end > start ? message.Substring(start, end - start).Trim() : string.Empty;
                                if (!string.IsNullOrWhiteSpace(clientId))
                                    UpdateKnownClients(clientId, true);
                            }
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine($"Error processing connected message: {ex.Message}");
                        }
                    });
                }
                else if (message.Contains("disconnected") && message.Contains("Client ID:"))
                {
                    await Task.Run(() =>
                    {
                        try
                        {
                            int start = message.IndexOf("Client ID:", StringComparison.OrdinalIgnoreCase);
                            if (start >= 0)
                            {
                                start += 10;
                                int end = message.IndexOf(' ', start);
                                if (end < 0) end = message.Length;

                                string clientId = end > start ? message.Substring(start, end - start).Trim() : string.Empty;
                                if (!string.IsNullOrWhiteSpace(clientId))
                                    UpdateKnownClients(clientId, false);
                            }
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine($"Error processing disconnected message: {ex.Message}");
                        }
                    });
                }
            }
            catch (Exception ex)
            {
                LogMessage($"Error in Server_ClientMessageReceived: {ex.Message}");
            }

            if (message.Contains("LED ON"))
            {
                await Task.Run(() =>
                {
                    ledPictureBox?.Invoke((MethodInvoker)(() =>
                    {
                        ledPictureBox.Image = Properties.Resources.on_led;
                    }));
                });
            }
            else if (message.Contains("LED OFF"))
            {
                await Task.Run(() =>
                {
                    ledPictureBox?.Invoke((MethodInvoker)(() =>
                    {
                        ledPictureBox.Image = Properties.Resources.off_led;
                    }));
                });
            }

            if (message.Contains("connected"))
            {
                await Task.Run(async () =>
                {
                    foreach (var id in _server.GetKnownClientIds())
                    {
                        await _server.SendToClientByIdAsync(id, "$GETID#");
                    }
                });
            }

            if (message.Contains("$1") || message.Contains("$0"))
            {
                await Task.Run(() =>
                {
                    var start = message.IndexOf('$');
                    var end = message.LastIndexOf('$');
                    if (start != -1 && end > start)
                    {
                        var statusStr = message.Substring(start + 1, end - start - 1);
                        var parts = statusStr.Split(',');
                        if (parts.Length == 4)
                        {
                            tableLayoutPanel1.Invoke((MethodInvoker)(() =>
                            {
                                textBox1.Text = _lastSelectedClientId ?? "";
                                textBox2.Text = parts[0] == "1" ? "OK" : "FAIL";
                                textBox3.Text = parts[1] == "1" ? "OK" : "FAIL";
                                textBox4.Text = parts[2] == "1" ? "OK" : "FAIL";
                                textBox5.Text = parts[3] == "1" ? "OK" : "FAIL";

                                _ = UpdateExcelWithCurrentStatusAsync().ContinueWith(t =>
                                {
                                    if (t.Exception != null)
                                    {
                                        LogToFile($"Error in automatic Excel update: {t.Exception.InnerException?.Message}");
                                    }
                                }, TaskScheduler.Default);
                            }));
                        }
                    }
                });
            }

            if (message.Contains("$RESET#"))
            {
                await Task.Run(() =>
                {
                    MessageBox.Show("Client reset successful!", "RESET", MessageBoxButtons.OK, MessageBoxIcon.Information);
                });
            }
        }

        private void ComboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox1.SelectedIndex != -1)
                _lastSelectedClientId = comboBox1.SelectedItem.ToString();
        }

        private async void Button2_Click(object sender, EventArgs e)
        {
            LogToFile($"GET Status button clicked for client: {_lastSelectedClientId ?? "None"}");
            if (_lastSelectedClientId != null)
            {
                try
                {
                    await _server.SendToClientByIdAsync(_lastSelectedClientId, "$GETSTATUS#");
                }
                catch (Exception ex)
                {
                    LogToFile($"Error sending GETSTATUS to client {_lastSelectedClientId}: {ex.Message}");
                    MessageBox.Show($"Failed to send GETSTATUS: {ex.Message}", "Error",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private async void ResetButton_Click(object sender, EventArgs e)
        {
            LogToFile($"RESET button clicked for client: {_lastSelectedClientId ?? "None"}");
            if (_lastSelectedClientId != null)
            {
                try
                {
                    string cmd = $"RESET; ${_lastSelectedClientId}:RESET#";
                    await _server.SendToClientByIdAsync(_lastSelectedClientId, cmd);
                    LogToFile($"RESET command sent to client: {_lastSelectedClientId}");
                }
                catch (Exception ex)
                {
                    LogToFile($"Error sending RESET to client {_lastSelectedClientId}: {ex.Message}");
                    MessageBox.Show($"Failed to send RESET: {ex.Message}", "Error",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private async void StartServerButton_Click(object sender, EventArgs e)
        {
            LogToFile("Start/Stop Server button clicked");
            button1.Enabled = false;

            if (!_serverConnected)
            {
                _ = Task.Run(async () =>
                {
                    try
                    {
                        await _server.StartAsync();
                        _serverConnected = true;

                        Invoke((MethodInvoker)delegate
                        {
                            LogMessage("Server started successfully.");
                            sendMessageButton.Enabled = true;
                            button1.Text = "Stop Server";
                            button1.Enabled = true;
                        });
                    }
                    catch (Exception ex)
                    {
                        LogMessage($"Error starting server: {ex.Message}");
                        button1.Enabled = true;
                    }
                });
            }
            else
            {
                await Task.Run(() =>
                {
                    try
                    {
                        _server.Stop();
                        _serverConnected = false;

                        Invoke((MethodInvoker)delegate
                        {
                            LogMessage("Server stopped.");
                            sendMessageButton.Enabled = false;
                            if (localNetworkLabel != null) localNetworkLabel.Text = "N/A";
                            if (lanNetworkLabel != null) lanNetworkLabel.Text = "N/A";
                            button1.Text = "Start Server";
                            button1.Enabled = true;

                            // reset indicators and maps when server stops
                            if (_clientIndicators != null)
                            {
                                foreach (var pb in _clientIndicators)
                                {
                                    if (pb != null) pb.BackColor = Color.Red;
                                }
                            }
                            _knownClientIds.Clear();
                            _clientIndicatorMap.Clear();
                        });
                    }
                    catch (Exception ex)
                    {
                        LogMessage($"Error stopping server: {ex.Message}");
                        button1.Enabled = true;
                    }
                });
            }
        }

        private async void SendMessageButton_Click(object sender, EventArgs e)
        {
            LogToFile("Send Message button clicked");
            if (!_serverConnected)
            {
                MessageBox.Show("Server is not running. Please start the server first.", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            sendMessageButton.Enabled = false;
            try
            {
                await ShowSendMessageDialog();
            }
            finally
            {
                sendMessageButton.Enabled = true;
            }
        }

        private async Task ShowSendMessageDialog()
        {
            LogToFile("Preparing to show send message dialog");
            await Task.Run(() =>
            {
                using (var inputDialog = new Form())
                {
                    inputDialog.Text = "Send Message to Client(s)";
                    inputDialog.Width = 400;
                    inputDialog.Height = 200;
                    inputDialog.FormBorderStyle = FormBorderStyle.FixedDialog;
                    inputDialog.MaximizeBox = false;
                    inputDialog.MinimizeBox = false;
                    inputDialog.StartPosition = FormStartPosition.CenterParent;

                    var messageLabel = new Label
                    {
                        Text = "Enter message to send to all connected clients:",
                        Left = 20,
                        Top = 20,
                        AutoSize = true
                    };

                    var messageBox = new TextBox
                    {
                        Left = 20,
                        Top = 50,
                        Width = 350,
                        Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
                        Multiline = true,
                        Height = 60
                    };

                    var sendButton = new Button
                    {
                        Text = "Send",
                        Left = 295,
                        Top = 120,
                        Width = 75,
                        Anchor = AnchorStyles.Bottom | AnchorStyles.Right
                    };

                    var cancelButton = new Button
                    {
                        Text = "Cancel",
                        Left = 210,
                        Top = 120,
                        Width = 75,
                        Anchor = AnchorStyles.Bottom | AnchorStyles.Right
                    };

                    sendButton.Click += async (s, args) =>
                    {
                        if (!string.IsNullOrWhiteSpace(messageBox.Text))
                        {
                            try
                            {
                                string messageToSend = $"[Server] {messageBox.Text}";
                                await _server.SendToClientAsync(messageToSend);
                                LogMessage($"Sent to client(s): {messageToSend}");
                                LogToFile($"Message sent to all connected clients. Length: {messageToSend.Length} chars");
                                inputDialog.DialogResult = DialogResult.OK;
                                inputDialog.Close();
                            }
                            catch (Exception ex)
                            {
                                LogMessage($"[Error] Failed to send message: {ex.Message}");
                            }
                        }
                        else
                        {
                            MessageBox.Show("Please enter a message to send.", "Error",
                                MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            return;
                        }
                    };

                    cancelButton.Click += (s, args) =>
                    {
                        inputDialog.DialogResult = DialogResult.Cancel;
                        inputDialog.Close();
                    };

                    inputDialog.Controls.Add(messageLabel);
                    inputDialog.Controls.Add(messageBox);
                    inputDialog.Controls.Add(sendButton);
                    inputDialog.Controls.Add(cancelButton);
                    inputDialog.AcceptButton = sendButton;
                    inputDialog.CancelButton = cancelButton;
                    inputDialog.ShowDialog(this);
                }
            });
        }

        private async void SendButton_Click(object sender, EventArgs e)
        {
            using (var inputDialog = new Form())
            {
                inputDialog.Text = "Send Message";
                inputDialog.Width = 300;
                inputDialog.Height = 150;
                inputDialog.FormBorderStyle = FormBorderStyle.FixedDialog;
                inputDialog.MaximizeBox = false;
                inputDialog.MinimizeBox = false;
                inputDialog.StartPosition = FormStartPosition.CenterParent;

                var messageBox = new TextBox
                {
                    Left = 20,
                    Top = 20,
                    Width = 240,
                    Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
                };

                var sendButton = new Button
                {
                    Text = "Send",
                    Left = 185,
                    Top = 60,
                    Width = 75,
                    Anchor = AnchorStyles.Top | AnchorStyles.Right
                };

                sendButton.Click += async (s, args) =>
                {
                    if (!string.IsNullOrWhiteSpace(messageBox.Text))
                    {
                        await Task.Run(async () =>
                        {
                            await _server.SendToClientAsync(messageBox.Text);
                        });
                        LogMessage($"[Server] {messageBox.Text}");
                        inputDialog.DialogResult = DialogResult.OK;
                    }
                };

                inputDialog.Controls.Add(messageBox);
                inputDialog.Controls.Add(sendButton);
                inputDialog.AcceptButton = sendButton;

                await Task.Run(() => inputDialog.ShowDialog(this));
            }
        }

        private void Label5_Click(object sender, EventArgs e) { }
        private void Label6_Click(object sender, EventArgs e) { }
        private void Label7_Click(object sender, EventArgs e) { }

        private void Button4_Click(object sender, EventArgs e)
        {
            LogToFile("Client Status button clicked");
            try
            {
                // placeholder for future Client_status form
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[{DateTime.Now:HH:mm:ss.fff}] [ERROR] Failed to open Client_status form: {ex}");
            }
        }

        private async void Button3_Click(object sender, EventArgs e)
        {
            Debug.WriteLine($"[DEBUG] Reset button clicked. _lastSelectedClientId: {_lastSelectedClientId}");

            if (string.IsNullOrWhiteSpace(_lastSelectedClientId))
            {
                Debug.WriteLine("[DEBUG] No client selected or _lastSelectedClientId is empty");
                MessageBox.Show("Select a client first.", "Reset", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            try
            {
                var knownClientIds = _server.GetKnownClientIds();
                Debug.WriteLine($"[DEBUG] Known client IDs: {string.Join(", ", knownClientIds)}");

                if (!knownClientIds.Contains(_lastSelectedClientId))
                {
                    Debug.WriteLine($"[DEBUG] Selected client ID '{_lastSelectedClientId}' is not connected.");
                    MessageBox.Show("Selected client is not connected.", "Reset", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                Debug.WriteLine($"[DEBUG] Attempting to send $RESET# to client: {_lastSelectedClientId}");

                await _server.SendToClientByIdAsync(_lastSelectedClientId, "$RESET#");

                Debug.WriteLine($"[DEBUG] Sent $RESET# to client: {_lastSelectedClientId}");
                MessageBox.Show($"Sent $RESET# to client {_lastSelectedClientId}.", "Reset", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[{DateTime.Now:HH:mm:ss.fff}] [ERROR] Failed to send $RESET#: {ex}");
                MessageBox.Show("Failed to send $RESET# to the client.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private async void Form1_Load(object sender, EventArgs e)
        {
            try
            {
                await Task.Run(() =>
                {
                    LogToFile("Form1 loaded and initializing components");
                    LogToFile($"Application version: {Application.ProductVersion}");
                    LogToFile($"OS Version: {Environment.OSVersion}");
                });

                if (Controls.Find("btnExportToExcel", true).FirstOrDefault() is Button btnExportToExcel)
                {
                    btnExportToExcel.Enabled = false;
                    btnExportToExcel.Visible = false;
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error in Form1_Load: {ex.Message}");
            }
        }

        private void ExportToExcel_Click(object sender, EventArgs e)
        {
            LogToFile("Exporting data to Excel");
            try
            {
                if (_workbook != null)
                {
                    try
                    {
                        _workbook.Save();
                    }
                    catch (Exception ex)
                    {
                        LogToFile($"ExportToExcel: Save failed: {ex.Message}");
                    }
                    Process.Start(_excelFilePath);
                }
                else
                {
                    Process.Start(_excelFilePath);
                }
            }
            catch (Exception ex)
            {
                LogToFile($"Error exporting to Excel: {ex.Message}");
                MessageBox.Show($"Error exporting to Excel: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (_worksheet != null)
                {
                    try { Marshal.ReleaseComObject(_worksheet); } catch { }
                    _worksheet = null;
                }

                if (_workbook != null)
                {
                    try
                    {
                        _workbook.Close(false);
                        Marshal.ReleaseComObject(_workbook);
                    }
                    catch (Exception ex)
                    {
                        LogToFile($"Error closing workbook in ExportToExcel_Click: {ex.Message}");
                    }
                    _workbook = null;
                }

                if (_excelApp != null)
                {
                    try
                    {
                        _excelApp.Quit();
                        Marshal.ReleaseComObject(_excelApp);
                    }
                    catch (Exception ex)
                    {
                        LogToFile($"Error quitting Excel in ExportToExcel_Click: {ex.Message}");
                    }
                    _excelApp = null;
                }

                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        private Task ExportToExcelAsync()
        {
            if (_disposed)
                return Task.CompletedTask;

            return UpdateExcelWithCurrentStatusAsync();
        }

        // Designer expects this name; forward to main handler if needed
        private void comboBox1_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            ComboBox1_SelectedIndexChanged(sender, e);
        }
    }
}
