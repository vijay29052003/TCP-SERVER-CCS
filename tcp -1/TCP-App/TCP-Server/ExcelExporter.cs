// using System;
// using System.IO;
// using System.Windows.Forms;
// using Excel = Microsoft.Office.Interop.Excel;
// using System.Drawing;

// namespace TCP_Server
// {
//     public class ExcelExporter
//     {
//         private readonly string _filePath;
//         private Excel.Application _excelApp;
//         private Excel.Workbook _workbook;
//         private Excel.Worksheet _worksheet;
//         private int _currentRow = 1;

//         public ExcelExporter()
//         {
//             string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
//             _filePath = Path.Combine(desktopPath, "tcp excel.xlsx");
//             InitializeExcel();
//         }

//         private void InitializeExcel()
//         {
//             try
//             {
//                 _excelApp = new Excel.Application { Visible = false };

//                 if (File.Exists(_filePath))
//                 {
//                     _workbook = _excelApp.Workbooks.Open(_filePath);
//                 }
//                 else
//                 {
//                     _workbook = _excelApp.Workbooks.Add();
//                     _workbook.SaveAs(_filePath);
//                 }

//                 if (_workbook.Worksheets.Count == 0)
//                 {
//                     _worksheet = (Excel.Worksheet)_workbook.Worksheets.Add();
//                 }
//                 else
//                 {
//                     _worksheet = (Excel.Worksheet)_workbook.Worksheets[1];
//                 }

//                 SetupWorksheet();
//             }
//             catch (Exception ex)
//             {
//                 throw new Exception($"Error initializing Excel: {ex.Message}");
//             }
//         }

//         private void SetupWorksheet()
//         {
//             // Clear existing content
//             _worksheet.Cells.Clear();

//             // Set title
//             _worksheet.Cells[1, 1] = "CLIENT STATUS";
//             Excel.Range titleRange = _worksheet.Range["A1"];
//             titleRange.Font.Bold = true;
//             titleRange.Font.Size = 16;
//             titleRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
//             _worksheet.Range["A1:F1"].Merge();

//             // Add client ID section
//             _worksheet.Cells[3, 1] = "CLIENT ID";
//             _worksheet.Range["A3"].Font.Bold = true;
//             _worksheet.Cells[3, 2] = "CLIENT001";

//             // Add status sections
//             string[] sections = { "RTC", "RS485", "GPS", "FLASH" };
//             for (int i = 0; i < sections.Length; i++)
//             {
//                 int row = 5 + i;
//                 _worksheet.Cells[row, 1] = $"{sections[i]} STATUS";
//                 _worksheet.Cells[row, 2] = "OK";
//                 _worksheet.Cells[row, 1].Font.Bold = true;
//             }

//             // Add table headers
//             string[] headers = { "S.NO", "CLIENT ID", "RTC STATUS", "RS485", "GPS", "FLASH" };
//             for (int i = 0; i < headers.Length; i++)
//             {
//                 _worksheet.Cells[10, i + 1] = headers[i];
//             }

//             // Format table headers
//             Excel.Range headerRange = _worksheet.Range["A10:F10"];
//             headerRange.Font.Bold = true;
//             headerRange.Interior.Color = ColorTranslator.ToOle(Color.LightGray);
//             headerRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

//             // Set column widths
//             _worksheet.Columns[1].ColumnWidth = 8;   // S.NO
//             _worksheet.Columns[2].ColumnWidth = 15;  // CLIENT ID
//             _worksheet.Columns[3].ColumnWidth = 15;  // RTC STATUS
//             _worksheet.Columns[4].ColumnWidth = 12;  // RS485
//             _worksheet.Columns[5].ColumnWidth = 10;  // GPS
//             _worksheet.Columns[6].ColumnWidth = 12;  // FLASH

//             _currentRow = 11; // Start data from row 11
//             _workbook.Save();
//         }

//         public void UpdateStatus(string clientId, string rtcStatus = "OK", string rs485 = "OK", string gps = "OK", string flash = "OK")
//         {
//             try
//             {
//                 // Update status section
//                 _worksheet.Cells[3, 2] = clientId;
//                 _worksheet.Cells[5, 2] = rtcStatus;
//                 _worksheet.Cells[6, 2] = rs485;
//                 _worksheet.Cells[7, 2] = gps;
//                 _worksheet.Cells[8, 2] = flash;

//                 // Add to table
//                 int rowNum = _currentRow - 10;
//                 _worksheet.Cells[_currentRow, 1] = rowNum;
//                 _worksheet.Cells[_currentRow, 2] = clientId;
//                 _worksheet.Cells[_currentRow, 3] = rtcStatus;
//                 _worksheet.Cells[_currentRow, 4] = rs485;
//                 _worksheet.Cells[_currentRow, 5] = gps;
//                 _worksheet.Cells[_currentRow, 6] = flash;

//                 // Format the new row
//                 Excel.Range dataRange = _worksheet.Range[$"A{_currentRow}:F{_currentRow}"];
//                 dataRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
//                 dataRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

//                 _currentRow++;
//                 _workbook.Save();
//             }
//             catch (Exception ex)
//             {
//                 throw new Exception($"Error updating Excel: {ex.Message}");
//             }
//         }

//         public void ShowExcel()
//         {
//             try
//             {
//                 _workbook.Save();
//                 _excelApp.Visible = true;
//             }
//             catch (Exception ex)
//             {
//                 throw new Exception($"Error showing Excel: {ex.Message}");
//             }
//         }

//         public void Close()
//         {
//             try
//             {
//                 _workbook?.Close(false);
//                 _excelApp?.Quit();
//                 ReleaseObject(_worksheet);
//                 ReleaseObject(_workbook);
//                 ReleaseObject(_excelApp);
//             }
//             catch (Exception ex)
//             {
//                 Console.WriteLine($"Error closing Excel: {ex.Message}");
//             }
//         }

//         private void ReleaseObject(object obj)
//         {
//             try
//             {
//                 if (obj != null)
//                 {
//                     System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
//                     obj = null;
//                 }
//             }
//             catch (Exception ex)
//             {
//                 obj = null;
//                 Console.WriteLine($"Error releasing object: {ex.Message}");
//             }
//             finally
//             {
//                 GC.Collect();
//             }
//         }
//     }
// }
