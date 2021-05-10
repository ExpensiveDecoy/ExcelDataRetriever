using System;
using System.Data;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelDataRetriever
{
    class Program
    {

        static void Main(string[] args)
        {
            Dictionary<string, DataTable> Sheets;
            Sheets = RetrieveDataFromXlsxFile("put_your_xlsx_filepath_here");
        }

        public static Dictionary<string, DataTable> RetrieveDataFromXlsxFile(string filePath)
        {
            Excel.Application excelApp = null;
            Excel.Workbook excelWorkbook = null;
            Excel.Worksheet excelSheet = null;
            Excel.Range excelRange = null;

            Dictionary<string, DataTable> Result = new Dictionary<string, DataTable>();
            try
            {
                excelApp = new Excel.Application();
                excelApp.Visible = false;
                excelWorkbook = excelApp.Workbooks.Open(filePath, Type.Missing, true);

                int numberOfSheets = excelWorkbook.Sheets.Count;
                for (int i = 1; i <= numberOfSheets; i++)
                {
                    excelSheet = excelWorkbook.Sheets[i];
                    string sheetName = excelSheet.Name;

                    excelRange = excelSheet.UsedRange;
                    int numberOfRows = excelRange.Rows.Count;
                    int numberOfColumns = excelRange.Columns.Count;

                    object[,] values = excelRange.Value2;

                    DataTable dataTable = new DataTable();
                    DataColumn dataColumn;
                    DataRow dataRow;

                    for (int j = 1; j <= numberOfColumns; j++)
                    {
                        dataColumn = new DataColumn();
                        dataColumn.DataType = Type.GetType("System.String");
                        dataColumn.ColumnName = j.ToString();
                        dataTable.Columns.Add(dataColumn);
                    }

                    for (int j = 1; j <= numberOfRows; j++)
                    {
                        dataRow = dataTable.NewRow();
                        for (int k = 1; k <= numberOfColumns; k++)
                        {
                            string value = values[j, k] as string;
                            if (!string.IsNullOrEmpty(value))
                            {
                                dataRow[k.ToString()] = value;
                            }
                            else
                            {
                                dataRow[k.ToString()] = string.Empty;
                            }
                        }
                        dataTable.Rows.Add(dataRow);
                    }

                    Result.Add(sheetName, dataTable);
                }
                excelWorkbook.Close();
                excelApp.Quit();
                return Result;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return null;
            }
            finally
            {
                if (excelRange != null) System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excelRange);
                if (excelSheet != null) System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excelSheet);
                if (excelWorkbook != null) System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excelWorkbook);
                if (excelApp != null) System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excelApp);
            }
        }
    }
}
