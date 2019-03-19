using System;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelWriter
{
    class ExcelWriterData : INotifyPropertyChanged
    {
        private string afield = string.Empty;
        private string bfield = string.Empty;
        private string cfield = string.Empty;
        private string dfield = string.Empty;
        private string efield = string.Empty;

        public string AField
        {
            get => afield;
            set { afield = value; NotifyPropertyChanged(); }
        }
        public string BField
        {
            get => bfield;
            set { bfield = value; NotifyPropertyChanged(); }
        }
        public string CField
        {
            get => cfield;
            set { cfield = value; NotifyPropertyChanged(); }
        }
        public string DField
        {
            get => dfield;
            set { dfield = value; NotifyPropertyChanged(); }
        }
        public string EField
        {
            get => efield;
            set { efield = value; NotifyPropertyChanged(); }
        }

        private string result;
        public string Result
        {
            get => result;
            set { result = value; NotifyPropertyChanged(); }
        }

        public ExcelWriterData()
        {
            Result = string.Empty;
        }

        public event PropertyChangedEventHandler PropertyChanged;

        private void NotifyPropertyChanged([CallerMemberName] string propertyName = "")
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        public void ClearValues()
        {
            AField = string.Empty;
            BField = string.Empty;
            CField = string.Empty;
            DField = string.Empty;
            EField = string.Empty;
        }

        public async Task AppendToFileAsync(string filename, CancellationToken cancellationToken)
        {
            try
            {
                await Task.Run(() =>
                {
                    //Result = "Sleeping in background...";
                    //System.Threading.Thread.Sleep(3000);
                    Result = "Appending...";

                    Excel.Application app = null;
                    Excel.Workbook workbook = null;
                    Excel.Worksheet worksheet = null;

                    try
                    {
                        // Open the new Excel instance
                        app = new Excel.Application();
                        if (app == null)
                            throw new Exception("Could not open the Excel application");

                        // Open the file
                        workbook = app.Workbooks.Open(filename);
                        if (workbook == null)
                            throw new Exception($"Could not open the file: {filename}");

                        // Get the first worksheet
                        worksheet = (Excel.Worksheet)workbook.Worksheets[1];
                        if (worksheet == null)
                            throw new Exception("Could not acquire the worksheet");

                        // Find the first empty row
                        int rowId = 0;
                        int numNotEmptyCells = 1;
                        while (numNotEmptyCells > 0)
                        {
                            string cell = "A" + (++rowId).ToString();
                            Excel.Range row = app.Range[cell, cell].EntireRow;
                            numNotEmptyCells = (int)app.WorksheetFunction.CountA(row);
                        }

                        // Write data
                        worksheet.Cells[rowId, 1].Value = AField;
                        worksheet.Cells[rowId, 2].Value = BField;
                        worksheet.Cells[rowId, 3].Value = CField;
                        worksheet.Cells[rowId, 4].Value = DField;
                        worksheet.Cells[rowId, 5].Value = EField;

                        if (cancellationToken.IsCancellationRequested)
                            throw new OperationCanceledException();

                        // Save the workbook
                        workbook.Save();

                        // Close the workbook
                        workbook.Close(true);
                        Result = "Append operation was successful";
                    }
                    catch (OperationCanceledException)
                    {
                        // cancellation is requested
                        // close workbook without saving it
                        if (workbook != null)
                            workbook.Close(false);

                        Result = "Append operation was canceled";
                    }
                    catch (Exception e)
                    {
                        Result = "Could not write data: " + e.Message;
                    }
                    finally
                    {
                        if (app != null)
                            app.Quit();
                    }
                },
                cancellationToken);
            }
            catch (OperationCanceledException)
            {
                Result = "Append operation was canceled";
            }
            catch (Exception)
            {
                Result = "Could not write data";
            }
        }
    }
}
