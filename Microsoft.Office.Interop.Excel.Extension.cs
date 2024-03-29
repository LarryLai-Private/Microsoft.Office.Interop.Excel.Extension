﻿using System;
using System.Data;
using System.Reflection;

namespace Microsoft.Office.Interop.Excel.Extension
{
    using Microsoft.Office.Interop.Excel;
    using System.Runtime.InteropServices;

    static public class ExcelFunc
    {
        /// <summary>
        /// CloseExcel
        /// </summary>
        /// <param name="excelApp"></param>
        static public void CloseExcel(Application excelApp)
        {
            if (excelApp != null)
            {
                try
                {
                    excelApp.Quit();
                }
                finally
                {
                    excelApp = null;

                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }
            }
        }

        /// <summary>
        /// OpenToPrintPreview
        /// </summary>
        /// <param name="FullPath"></param>
        static public void OpenToPrintPreview(string FullPath)
        {
            var excelApp = new Application();

            try
            {
                Workbook WorkBook = excelApp.Workbooks.Open(FullPath);
                WorkBook.PrintPreview();
            }
            finally
            {
                CloseExcel(excelApp);
            }
        }

        public class Export
        {
            #region DataTableToWorksheet
            /// <summary>
            /// DataTable轉Worksheet處理Columns的方式
            /// </summary>
            public enum DataTableToWorksheet_ColumnsType
            {
                /// <summary>
                /// 使用來源Table的欄位
                /// </summary>
                UseDataTableColumns,
                /// <summary>
                /// 使用目標Sheet的欄位
                /// </summary>
                UseWorksheetColumns,
            }
            /// <summary>
            /// DataTable匯入Worksheet
            /// </summary>
            /// <param name="dataTable">來源DataTable</param>
            /// <param name="worksheet">目的Worksheet</param>
            /// <param name="columnsType">DataTable轉Worksheet處理Columns的方式</param>
            /// <param name="ColumnsAutoFit">自動調整欄寬</param>
            static private void DataTableToWorksheet(System.Data.DataTable dataTable, Worksheet worksheet, DataTableToWorksheet_ColumnsType columnsType, bool ColumnsAutoFit = false)
            {

                int rowsCount = dataTable.Rows.Count;
                int colsCount = dataTable.Columns.Count;

                //使用來源標題列,目標需增加一列
                if (columnsType == DataTableToWorksheet_ColumnsType.UseDataTableColumns) rowsCount += 1;

                object[,] valuesArray = new object[rowsCount, colsCount];

                int rowIndex = 0;
                int colIndex = 0;

                //使用來源標題列,目標資料起始列+1
                if (columnsType == DataTableToWorksheet_ColumnsType.UseDataTableColumns)
                {
                    for (int c = colIndex; c < colsCount; c++)
                    {
                        valuesArray[0, c] = dataTable.Columns[c].ColumnName;
                    }

                    rowIndex += 1;
                }

                for (int r = rowIndex; r < rowsCount; r++)
                {
                    for (int c = colIndex; c < colsCount; c++)
                    {
                        int rIndex = r;

                        //使用來源標題列,目標資料列索引-1
                        if (columnsType == DataTableToWorksheet_ColumnsType.UseDataTableColumns)
                        {
                            rIndex -= 1;
                        }
                        //若DataTable欄資料為非string時有可能發生錯誤(如DateTime),統一使用toString()轉字串
                        valuesArray[r, c] = dataTable.Rows[rIndex][c].ToString();
                    }
                }

                rowIndex = 1;
                colIndex = 1;

                //目的已含標題列,列起始索引與總列數+1
                if (columnsType == DataTableToWorksheet_ColumnsType.UseWorksheetColumns)
                {
                    rowIndex += 1;
                    rowsCount += 1;
                }

                Range range;
                range = worksheet.Range[worksheet.Cells[rowIndex, colIndex], worksheet.Cells[rowsCount, colsCount]];

                range.Value2 = valuesArray;

                if (ColumnsAutoFit)
                {
                    range.Columns.AutoFit();
                }

                if (dataTable.TableName != string.Empty) worksheet.Name = dataTable.TableName;
            }
            /// <summary>
            /// DataTableToWorksheet_UseDataTableColumns
            /// </summary>
            /// <param name="dataTable">來源DataTable</param>
            /// <param name="worksheet">目的Worksheet</param>
            /// <param name="ColumnsAutoFit">自動調整欄寬</param>
            static public void DataTableToWorksheet_UseDataTableColumns(System.Data.DataTable dataTable, Worksheet worksheet, bool ColumnsAutoFit = false)
            {
                DataTableToWorksheet(dataTable, worksheet, DataTableToWorksheet_ColumnsType.UseDataTableColumns, ColumnsAutoFit);
            }
            /// <summary>
            /// DataTable匯入Worksheet 使用目的Worksheet的欄位
            /// </summary>
            /// <param name="dataTable">來源DataTable</param>
            /// <param name="worksheet">目的Worksheet</param>
            /// <param name="ColumnsAutoFit">自動調整欄寬</param>
            static public void DataTableToWorksheet_UseWorksheetColumns(System.Data.DataTable dataTable, Worksheet worksheet, bool ColumnsAutoFit = false)
            {
                DataTableToWorksheet(dataTable, worksheet, DataTableToWorksheet_ColumnsType.UseWorksheetColumns, ColumnsAutoFit);
            }
            #endregion

            #region CreateExcel
            static public Workbook CreateExcel(DataSet dSet, DataTableToWorksheet_ColumnsType columnsType, bool Visible = true, bool DisplayAlerts = true, bool ColumnsAutoFit = true, string DefFormat = "@")
            {
                Application excelApp;
                try
                {
                    excelApp = new Application
                    {
                        //一開始先關閉警告跟隱藏,等工作完成再套用參數
                        DisplayAlerts = false,
                        Visible = false
                    };


                    Workbook workbook = excelApp.Workbooks.Add();

                    Worksheet worksheet = null;
                    foreach (System.Data.DataTable dt in dSet.Tables)
                    {
                        //加到最後
                        worksheet = (Worksheet)workbook.Worksheets.Add(Missing.Value, workbook.Worksheets[workbook.Worksheets.Count]);

                        //全部預設為文字格式,否則很多字元會被轉為數值符號
                        worksheet.Columns.NumberFormat = DefFormat;

                        DataTableToWorksheet(dt, worksheet, columnsType, ColumnsAutoFit);
                    }


                    //刪掉多的預設資料表
                    int sunCount = workbook.Worksheets.Count - dSet.Tables.Count;
                    for (int i = 0; i < sunCount; i++)
                    {
                        Worksheet ws = (Worksheet)workbook.Worksheets[1];
                        ws.Delete();
                    }


                    excelApp.DisplayAlerts = DisplayAlerts;
                    excelApp.Visible = Visible;

                    return workbook;
                }
                catch (Exception ex)
                {
                    string msg = ex.Message;
                    throw ex;
                }
            }
            static public Workbook CreateExcel_UseDataTableColumns(DataSet dSet, bool Visible = true, bool DisplayAlerts = true, bool ColumnsAutoFit = true, string DefFormat = "@")
            {
                return CreateExcel(dSet, DataTableToWorksheet_ColumnsType.UseDataTableColumns, DisplayAlerts, Visible, ColumnsAutoFit, DefFormat);
            }

            static public Workbook CreateExcel(System.Data.DataTable dt, DataTableToWorksheet_ColumnsType columnsType, bool Visible = true, bool DisplayAlerts = true, bool ColumnsAutoFit = true, string DefFormat = "@")
            {
                DataSet ds = new DataSet();
                var dataTable = dt.Copy();
                ds.Tables.Add(dataTable);

                return CreateExcel(ds, columnsType, Visible, DisplayAlerts, ColumnsAutoFit, DefFormat);
            }
            static public Workbook CreateExcel_UseDataTableColumns(System.Data.DataTable dt, bool Visible = true, bool DisplayAlerts = true, bool ColumnsAutoFit = true, string DefFormat = "@")
            {
                return CreateExcel(dt, DataTableToWorksheet_ColumnsType.UseDataTableColumns, Visible, DisplayAlerts, ColumnsAutoFit, DefFormat);
            }
            #endregion

            #region SaveAsExcel
            /// <summary>
            /// 建立EXCEL並儲存
            /// </summary>
            /// <param name="dataTable"></param>
            /// <param name="FileName"></param>
            /// <param name="columnsType">DataTable轉Worksheet處理Columns的方式</param>
            /// <param name="ColumnsAutoFit"></param>
            /// <param name="Visible"></param>
            /// <param name="DisplayAlerts"></param>
            /// <returns></returns>
            static public bool SaveAsExcel(System.Data.DataTable dataTable, string FileName, DataTableToWorksheet_ColumnsType columnsType, bool ColumnsAutoFit = true, bool Visible = false, bool DisplayAlerts = false)
            {

                Workbook workbook = null;
                try
                {
                    workbook = CreateExcel(dataTable, columnsType, Visible, DisplayAlerts, ColumnsAutoFit);
                    workbook.SaveAs(FileName);
                    return true;
                }
                catch (Exception ex)
                {
                    string msg = ex.Message;
                    return false;
                }
                finally
                {
                    try
                    {
                        CloseExcel(workbook.Application);
                    }
                    catch { }
                }

            }
            /// <summary>
            /// 建立EXCEL並儲存 資料表欄位使用來源DataTable的欄位
            /// </summary>
            /// <param name="dataTable"></param>
            /// <param name="FileName"></param>
            /// <param name="ColumnsAutoFit"></param>
            /// <param name="Visible"></param>
            /// <param name="DisplayAlerts"></param>
            /// <returns></returns>
            static public bool SaveAsExcel_UseDataTableColumns(System.Data.DataTable dataTable, string FileName, bool ColumnsAutoFit = true, bool Visible = false, bool DisplayAlerts = false)
            {
                return SaveAsExcel(dataTable, FileName, DataTableToWorksheet_ColumnsType.UseDataTableColumns, ColumnsAutoFit, Visible, DisplayAlerts);
            }
            /// <summary>
            /// 建立EXCEL並儲存 資料表欄位使用來源DataTable的欄位
            /// </summary>
            /// <param name="dataTable"></param>
            /// <param name="FileName"></param>
            /// <param name="ColumnsAutoFit"></param>
            /// <param name="Visible"></param>
            /// <param name="DisplayAlerts"></param>
            /// <returns></returns>
            static public bool SaveAsExcel_UseWorksheetColumns(System.Data.DataTable dataTable, string FileName, bool ColumnsAutoFit = true, bool Visible = false, bool DisplayAlerts = false)
            {
                return SaveAsExcel(dataTable, FileName, DataTableToWorksheet_ColumnsType.UseWorksheetColumns, Visible, DisplayAlerts, ColumnsAutoFit);
            }
            #endregion
        }

        public class Import
        {
            static public Workbook GetWorkbook(string ExcelFilePath)
            {
                var application = new Application()
                {
                    DisplayAlerts = false,
                    Visible = false
                };

                try
                {
                    return application.Workbooks.Open(ExcelFilePath);
                }
                finally
                {
                    //注意: Excel是Unmanaged程式，要妥善結束才能乾淨不留痕跡
                    //否則，很容易留下一堆excel.exe在記憶體中
                    //所有用過的COM+物件都要使用Marshal.FinalReleaseComObject清掉
                    //COM+物件的Reference Counter，以利結束物件回收記憶體
                    //if (sheet != null)
                    //{
                    //    Marshal.FinalReleaseComObject(sheet);
                    //}
                    //if (wrkBook != null)
                    //{
                    //    wrkBook.Close(false); //忽略尚未存檔內容，避免跳出提示卡住
                    //    Marshal.FinalReleaseComObject(wrkBook);
                    //}
                    if (application != null)
                    {
                        application.Workbooks.Close();
                        application.Quit();
                        Marshal.FinalReleaseComObject(application);
                    }
                }
            }
            static public Worksheet GetWorkSheet(Workbook workbook, string SheetIndex) => (Worksheet)workbook.Sheets[SheetIndex];
            static public Worksheet GetWorkSheet(Workbook workbook, int SheetIndex = 1) => (Worksheet)workbook.Sheets[SheetIndex];
            static public System.Data.DataTable GetDataTable(Worksheet worksheet, bool hasTitleRow = true)
            {
                System.Data.DataTable dataTable = new System.Data.DataTable();

                //取得有資料的範圍
                Range range = worksheet.UsedRange.CurrentRegion;

                int index = 1;
                if (hasTitleRow)
                {

                    for (int j = 1; j <= range.Columns.Count; j++)
                    {
                        Range cell = (Range)range.Cells[1, j];
                        dataTable.Columns.Add(cell.Value2.ToString());
                    }
                    index++;
                }
                else
                {
                    for (int j = 1; j <= range.Columns.Count; j++)
                    {
                        string col = j.ToString();
                        dataTable.Columns.Add(col);
                    }
                }

                for (int row = index; row <= range.Rows.Count; row++)
                {
                    DataRow newRow = dataTable.NewRow();

                    for (int col = 0; col < range.Columns.Count; col++)
                    {
                        Range cell = (Range)range.Cells[row, col+1];
                        object value;
                        if (cell.Value2 == null)
                        {
                            value = string.Empty;
                        }
                        else
                        {
                            value = cell.Value2;
                        }
                        newRow[col] = value;
                    }

                    dataTable.Rows.Add(newRow);
                }

                return dataTable;
            }

            static public System.Data.DataTable TestGetDataTable(string ExcelFilePath)
            {
                var workbook = GetWorkbook(ExcelFilePath);
                var wrokSheet = GetWorkSheet(workbook);
                return WorkSheetToDataTable(wrokSheet);
            }

            /// <summary>
            /// WorkSheetToDataTable
            /// </summary>
            /// <param name="worksheet"></param>
            /// <returns>DataTable</returns>
            static public System.Data.DataTable WorkSheetToDataTable(Worksheet worksheet, bool hasTitleRow = true)
            {

                System.Data.DataTable dataTable = new System.Data.DataTable();
                Range range = worksheet.UsedRange.CurrentRegion;

                int index = 1;
                if (hasTitleRow)
                {
                    for (int j = 1; j <= range.Columns.Count; j++)
                    {
                        Range cell = (Range)range.Cells[1, j];
                        dataTable.Columns.Add(cell.Value2.ToString());
                    }
                    index = 2;
                }
                else
                {
                    for (int j = 1; j <= range.Columns.Count; j++)
                    {
                        string col = j.ToString();
                        dataTable.Columns.Add(col);
                    }
                }

                for (int i = index; i <= range.Rows.Count; i++)
                {
                    DataRow newRow = dataTable.NewRow();

                    for (int j = 1; j <= range.Columns.Count; j++)
                    {
                        Range cell = (Range)range.Cells[i, j];
                        object value;
                        if (cell.Value2 == null)
                        {
                            value = string.Empty;
                        }
                        else
                        {
                            value = cell.Value2;
                        }
                        newRow[j - 1] = value;
                    }

                    dataTable.Rows.Add(newRow);
                }

                return dataTable;

            }

            /// <summary>
            /// 從EXCEL檔案的工作表匯出DataTable
            /// </summary>
            /// <param name="ImportPath"></param>
            /// <param name="SheetIndex"></param>
            /// <returns>dataTable</returns>
            static public System.Data.DataTable ExcelToDataTableFromFileName(string ImportPath, int SheetIndex, bool hasTitleRow = true)
            {
                System.Data.DataTable dt = new System.Data.DataTable();

                var excelApp = new Application()
                {
                    DisplayAlerts = false,
                    Visible = false
                };

                try
                {
                    Workbook wb;
                    wb = excelApp.Workbooks.Open(ImportPath);
                    Worksheet ws = (Worksheet)wb.Sheets[SheetIndex];
                    dt = WorkSheetToDataTable(ws, hasTitleRow);
                }
                finally
                {
                    CloseExcel(excelApp);
                }

                return dt;
            }


            /// <summary>
            /// 從EXCEL檔案的工作表匯出DataTable
            /// </summary>
            /// <param name="ImportPath"></param>
            /// <param name="SheetIndex"></param>
            /// <returns>dataTable</returns>
            static public System.Data.DataTable ExcelSheetObjToDataTableFromFileName(string ImportPath, string SheetIndex, bool hasTitleRow = true)
            {
                return ExcelSheetObjToDataTableFromFileName(ImportPath, (object)SheetIndex, hasTitleRow);
            }

            /// <summary>
            /// 從EXCEL檔案的工作表匯出DataTable
            /// </summary>
            /// <param name="ImportPath"></param>
            /// <param name="SheetIndex"></param>
            /// <returns>dataTable</returns>
            static System.Data.DataTable ExcelSheetObjToDataTableFromFileName(string ImportPath, object SheetIndex, bool hasTitleRow = true)
            {
                System.Data.DataTable dt = new System.Data.DataTable();

                var excelApp = new Application()
                {
                    DisplayAlerts = false,
                    Visible = false
                };

                try
                {
                    Workbook wb;
                    wb = excelApp.Workbooks.Open(ImportPath);
                    Worksheet ws = (Worksheet)wb.Sheets[SheetIndex];
                    dt = WorkSheetToDataTable(ws, hasTitleRow);
                }
                finally
                {
                    CloseExcel(excelApp);
                }

                return dt;
            }

        }
    }


    public class ExcelTool
    {
        System.Data.DataTable GetDataTableFormWorkSheet(Worksheet worksheet, bool hasTitleRow = true)
        {
            System.Data.DataTable dataTable = new System.Data.DataTable();

            //取得有資料的範圍
            Range range = worksheet.UsedRange.CurrentRegion;

            int index = 1;
            if (hasTitleRow)
            {

                for (int j = 1; j <= range.Columns.Count; j++)
                {
                    Range cell = (Range)range.Cells[1, j];
                    dataTable.Columns.Add(cell.Value2.ToString());
                }
                index++;
            }
            else
            {
                for (int j = 1; j <= range.Columns.Count; j++)
                {
                    string col = j.ToString();
                    dataTable.Columns.Add(col);
                }
            }

            for (int row = index; row <= range.Rows.Count; row++)
            {
                DataRow newRow = dataTable.NewRow();

                for (int col = 0; col < range.Columns.Count; col++)
                {
                    Range cell = (Range)range.Cells[row, col + 1];
                    object value;
                    if (cell.Value2 == null)
                    {
                        value = string.Empty;
                    }
                    else
                    {
                        value = cell.Value2;
                    }
                    newRow[col] = value;
                }

                dataTable.Rows.Add(newRow);
            }

            return dataTable;
        }


        delegate object Delegate_ProcessApplication(Application application);
        object ProcessApplication(string excelFilePath, Delegate_ProcessApplication delegate_func)
        {
            Application application = null;
            try
            {
                application = new Application()
                {
                    DisplayAlerts = false,
                    Visible = false,
                };
                return delegate_func.Invoke(application);
            }
            finally
            {
                if (application != null)
                {
                    application.Workbooks.Close();
                    application.Quit();
                    Marshal.FinalReleaseComObject(application);
                }
            }
        }

        delegate object Delegate_ProcessWorkbook(Workbook workbook);
        object ProcessWorkbook(string excelFilePath, Delegate_ProcessWorkbook delegate_func)
        {
            return ProcessApplication(excelFilePath, new Delegate_ProcessApplication(delegate(Application application)
            {
                Workbook workbook = null;
                try
                {
                    workbook = application.Workbooks.Open(excelFilePath);
                    return delegate_func.Invoke(workbook);
                }
                finally
                {
                    if (workbook != null)
                    {
                        workbook.Close(false); //忽略尚未存檔內容，避免跳出提示卡住
                        Marshal.FinalReleaseComObject(workbook);
                    }
                }
            }
            ));
        }

        delegate object Delegate_ProcessWorksheet(Worksheet worksheet);
        object ProcessWorksheet(string excelFilePath, Delegate_ProcessWorksheet delegate_func, object SheetIndex)
        {
            return ProcessWorkbook(excelFilePath, new Delegate_ProcessWorkbook(delegate (Workbook workbook)
            {
                Worksheet worksheet = null;
                try
                {
                    worksheet = (Worksheet)workbook.Sheets[SheetIndex];
                    return delegate_func.Invoke(worksheet);
                }
                finally
                {
                    if (worksheet != null)
                    {
                        Marshal.FinalReleaseComObject(worksheet);
                    }
                }
            }));
        }


        public System.Data.DataTable GetTableFormExcel(string excelFilePath, object SheetIndex)
        {
            return (System.Data.DataTable)ProcessWorksheet(excelFilePath, new Delegate_ProcessWorksheet(delegate(Worksheet worksheet)
            {
                return GetDataTableFormWorkSheet(worksheet);
            }), SheetIndex);
        }
    }
}