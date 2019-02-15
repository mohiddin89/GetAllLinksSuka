using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using excel = Microsoft.Office.Interop.Excel;

namespace GetList
{
    class ExportExcel
    {
        public void ExportToExcel(System.Data.DataSet dataSet, string outputPath)
        {

            try
            {
                excel._Application excelAppication = new excel.Application();
                excel.Workbook objWorkBook = excelAppication.Workbooks.Add(Type.Missing);


                int sheetIndex = 0;

                foreach (System.Data.DataTable dt in dataSet.Tables)
                {
                    object[,] rawData = new object[dt.Rows.Count + 1, dt.Columns.Count];

                    for (int column = 0; column < dt.Columns.Count; column++)
                    {
                        rawData[0, column] = dt.Columns[column].ColumnName;
                    }

                    for (int col = 0; col < dt.Columns.Count; col++)
                    {
                        try
                        {
                            for (int row = 0; row < dt.Rows.Count; row++)
                            {
                                try
                                {
                                    //if (col == 2)
                                    //{
                                    //    if (dt.Rows[row].ItemArray[col].ToString().Contains('/'))
                                    //    {

                                    //        rawData[row + 1, col] = Convert.ToDateTime(dt.Rows[row].ItemArray[col]).ToShortDateString();
                                    //    }
                                    //    else
                                    //    {
                                    //        double MyOADate = Convert.ToDouble(dt.Rows[row].ItemArray[col]);
                                    //        DateTime MyDate = DateTime.FromOADate(MyOADate);

                                    //        rawData[row + 1, col] = MyDate.ToShortDateString();
                                    //    }
                                    //}
                                    //else
                                    //{
                                    rawData[row + 1, col] = dt.Rows[row].ItemArray[col];
                                    //}

                                }
                                catch (Exception ex)
                                {

                                }
                                //rawData[row + 1, col] = dt.Rows[row].ItemArray[col];
                            }
                        }
                        catch (Exception ex)
                        {

                        }
                    }

                    // Calculate the final column letter
                    string finalColLetter = string.Empty;
                    string colCharset = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
                    int colCharsetLen = colCharset.Length;

                    if (dt.Columns.Count > colCharsetLen)
                    {
                        finalColLetter = colCharset.Substring(
                            (dt.Columns.Count - 1) / colCharsetLen - 1, 1);
                    }

                    finalColLetter += colCharset.Substring(
                            (dt.Columns.Count - 1) % colCharsetLen, 1);


                    excel.Worksheet objWorksheet = objWorkBook.Sheets.Add(objWorkBook.Sheets.get_Item(++sheetIndex), Type.Missing, 1, excel.XlSheetType.xlWorksheet);
                    objWorksheet.Name = dt.TableName;

                    // Fast data export to Excel
                    string excelRange = string.Format("A1:{0}{1}",
                        finalColLetter, dt.Rows.Count + 1);

                    objWorksheet.get_Range(excelRange, Type.Missing).Value2 = rawData;

                    ((excel.Range)objWorksheet.Rows[1, Type.Missing]).Font.Bold = true;
                }




                objWorkBook.SaveAs(outputPath, excel.XlFileFormat.xlOpenXMLWorkbook, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    excel.XlSaveAsAccessMode.xlNoChange, excel.XlSaveConflictResolution.xlUserResolution, Type.Missing, Type.Missing, Type.Missing);

                objWorkBook.Close(true, Type.Missing, Type.Missing);
                objWorkBook = null;

                excelAppication.Quit();
                excelAppication = null;

                GC.Collect();
                GC.WaitForPendingFinalizers();


            }
            catch (Exception ex)
            {
                Library.WriteLog("Error at writing excel file:- PageName:" , ex);
            }

        }
    }
}
