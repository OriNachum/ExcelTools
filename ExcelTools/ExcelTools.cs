using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelTools
{
    static public class Convert
    {
        #region Public Methods
        static public string[,] ExcelTabToTable(string pFileAddress)
        {
            return ExcelTabToTable(pFileAddress, 1);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="pTable">data to inject</param>
        /// <param name="pExcelName">File name full address</param>
        /// <param name="pTabName">Tab name</param>
        /// <param name="pCellIndex">Letter + number, as: A12</param>
        public static void InjectTableToSpecificTabCells(string[,] pTable, Cell pTableStartPosition, int pTableWidth, int pTableHeight, string pExcelName, string pTabName, Cell pSheetStartPosition)
        {
            using (var app = new ExcelApplication())
            using (var sheet = new ExcelFile(app, pExcelName, pReadOnly: false, pEditable: true))
            {
                Excel.Worksheet excelTab;
                excelTab = sheet.Worksheet.Worksheets[pTabName];
                excelTab.ReplaceRange(pSheetStartPosition, pTable, pTableStartPosition, pTableWidth, pTableHeight);
                sheet.Worksheet.Save();
                excelTab.ReleaseObject();
            }
        }



        static public string[,] ExcelTabToTable(string pFileAddress, int pTabNumber)
        {
            int resultHeight; 
            int resultWidth;
            return ExcelTabToTable(pFileAddress, pTabNumber, out resultHeight, out resultWidth);
        }
        static public string[,] ExcelTabToTable(string pFileAddress, string pTabName)
        {
            if (string.IsNullOrEmpty(pFileAddress))
            {
                throw new Exception(pFileAddress + " is null");
            }
            using (var excelApplication = new ExcelApplication())
            using (var excelFile = new ExcelFile(excelApplication, pFileAddress, pReadOnly: true, pEditable: false))
            {
                Excel.Worksheet excelTab;
                try
                {
                    excelTab = excelFile.Worksheets[pTabName];
                }
                catch (Exception e)
                {
                    var tabNames = new StringBuilder();
                    foreach (var name in excelFile.GetWorksheetNames())
                    {
                        tabNames.Append(name).Append(", ");
                    }

                    throw new Exception("Could not convert tab to table." + Environment.NewLine + 
                                        "Excel file: " + pFileAddress + Environment.NewLine + 
                                        "Tab name: " + pTabName + Environment.NewLine + 
                                        tabNames.ToString() + Environment.NewLine +
                                        "Message: " + e.Message + Environment.NewLine + 
                                        "Full stack: " + e.StackTrace);
                }
                try
                {
                    var result = ExportTabToTable(excelTab);
                    return result;
                }
                catch (Exception e)
                {
                    throw new Exception("Could not convert tab to table." + Environment.NewLine +
                                        "Excel file: " + pFileAddress + Environment.NewLine +
                                        "Tab name: " + pTabName + Environment.NewLine +
                                        "Message: " + e.Message + Environment.NewLine +
                                        "Full stack: " + e.StackTrace);
                }
                finally
                {
                    excelTab.ReleaseObject();
                }
            }
        }
        static public string[,] ExcelTabToTable(string pFileAddress, int pTabNumber, out int pResultHeight, out int pResultWidth)
        {

            using (var excelApplication = new ExcelApplication())
            using (var excelFile = new ExcelFile(excelApplication, pFileAddress, pReadOnly: true, pEditable: false))
            {
                Excel.Worksheet excelTab;
                try
                {
                    excelTab = excelFile.Worksheets.get_Item(pTabNumber);
                }
                catch (Exception e)
                {
                    throw e;
                }
                try
                {
                    var result = ExportTabToTable(excelTab);
                    pResultHeight = result.GetLength(0);
                    pResultWidth = result.GetLength(1);
                    return result;
                }
                catch (Exception e)
                {
                    throw e;
                }
                finally
                {
                    excelTab.ReleaseObject();
                }
            }
        }

        private static string[,] ExportTabToTable(Excel.Worksheet excelTab)
        {
            Excel.Range range = excelTab.UsedRange;
            var tableHeight = range.Rows.Count;
            var tableWidth = range.Columns.Count;
            string[,] resultTable = new string[tableHeight, tableWidth];
            for (int height = 0; height < tableHeight; height++)
            for (int width = 0; width < tableWidth; width++)
            {
                if (range.Cells[height + 1, width + 1].value2 != null)
                {
                    resultTable[height, width] = range.Cells[height + 1, width + 1].value2.ToString();
                }
                else
                {
                    resultTable[height, width] = "";
                }
            }
            return resultTable;
        }

        static public void RunMacro(string pFileAddress, string pMacroName)
        {
            using (var excelApplication = new ExcelApplication())
            using (var excelFile = new ExcelFile(excelApplication, pFileAddress, pEditable: true, pReadOnly: false))
            {
                excelApplication.Application.Run(pMacroName);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="pSheet"></param>
        /// <param name="pLineNumber"></param>
        /// <param name="pColunNumber"></param>
        /// <param name="pLineCount"></param>
        /// <param name="pColumnCount"></param>
        /// <param name="pTable"></param>
        static public void ReplaceRange(this Excel.Worksheet pSheet, int pLineNumber, int pColumnNumber, int pTableLine, int pTableColumn, string[,] pTable, int pColumnCount, int pLineCount)
        {
            ReplaceRange(pSheet, new Cell(pLineNumber,pColumnNumber),pTable, new Cell(pTableLine, pTableColumn),pLineCount, pColumnCount);
        }
        static public void ReplaceRange(this Excel.Worksheet pSheet, Cell pSheetStartPosition , string[,] pTable, Cell pTableStartPosition, int pColumnCount, int pLineCount)
        {           
            var maxLineNumber = Math.Min(pLineCount, pTable.GetLength(0)-pTableStartPosition.Row);
            var maxColumnNumber = Math.Min(pColumnCount, pTable.GetLength(1) - pTableStartPosition.Column);
            for (int i = 0; i< maxLineNumber; i++)
            {
                for (int j = 0; j < maxColumnNumber; j++)
                {
                    pSheet.Cells[pSheetStartPosition.Row + i, pSheetStartPosition.Column + j].Value = pTable[i+ pTableStartPosition.Row, j+ pTableStartPosition.Column];
                }
            }
        }

        static public void TableToNewFile(string pFileAddress, string[,] pStringTable, bool pWarpText = true)
        {
            int pTableHeight = pStringTable.GetLength(0);
            int pTableWidth = pStringTable.GetLength(1);
            TableToNewFile(pFileAddress, pStringTable, pTableHeight, pTableWidth, pWarpText);
        }
        static public void TableToNewFile(string pFileAddress, string[,] pStringTable, int pTableHeight, int pTableWidth, bool pWarpText = true)
        {
            using (var excelApplication = new ExcelApplication())
            using (var excelFile = new ExcelFile(excelApplication, pFileAddress, pReadOnly: false, pEditable: true, pNew: true))
            {
                Excel.Worksheet excelTab = excelFile.Worksheets.get_Item(1);
                UnpreparedInflateTableToFile(excelTab, pStringTable, pWarpText);
                SaveFile(excelApplication, excelFile, pFileAddress);
                excelTab.ReleaseObject();
                excelTab = null;
            }
        }

        static public void TableToExistingFile(string pFileAddress, string[,] pStringTable, bool pWarpText = true)
        {
            int pTableHeight = pStringTable.GetLength(0);
            int pTableWidth = pStringTable.GetLength(1);
            TableToExistingFile(pFileAddress, pStringTable, pTableHeight, pTableWidth, pWarpText);
        }
        static public void TableToExistingFile(string pFileAddress, string[,] pStringTable, int pTableHeight, int pTableWidth, bool pWarpText = true)
        {
            using (var excelApplication = new ExcelApplication())
            using (var excelFile = new ExcelFile(excelApplication, pFileAddress, pReadOnly: false, pEditable: true))
            {
                excelFile.Worksheets.Add();
                Excel.Worksheet excelTab = excelFile.Worksheets.get_Item(1);
                UnpreparedInflateTableToFile(excelTab, pStringTable, pWarpText);
                SaveFile(excelApplication, excelFile, pFileAddress);
                excelTab.ReleaseObject();
                excelTab = null;
            }
        }

        static public void ClearExcelTab(String excelName, int tabIndex)
        {
            using (var xlApp = new ExcelApplication())
            using (var xlClearWorkBook = new ExcelFile(xlApp, excelName,pReadOnly: false, pEditable: true))
            {
                Excel.Worksheet xlClearWorkSheet;

                object misValue = System.Reflection.Missing.Value;
                xlClearWorkSheet = (Excel.Worksheet)xlClearWorkBook.Worksheets.get_Item(1);
                xlClearWorkSheet.Select(Type.Missing);
                Excel.Range range = xlClearWorkSheet.get_Range("A:Z");
                range.Delete(Excel.XlDeleteShiftDirection.xlShiftToLeft);

                range.ReleaseObject();
                xlClearWorkSheet.ReleaseObject();

                xlClearWorkBook.Save();
            }
        }  // clearExcelTab
        static public void appendExcelToTab(String sourceFileName, String destinationFileName)
        {
            using (var xlApp = new ExcelApplication())
            using (var xlSourceWorkBook = new ExcelFile(xlApp, sourceFileName, pReadOnly: true, pEditable: false))
            using (var xlDestinationWorkBook = new ExcelFile(xlApp, destinationFileName, pReadOnly: false, pEditable: true))
            {
                Excel.Worksheet xlSourceWorkSheet, xlDestinationWorkSheet;
                object misValue = System.Reflection.Missing.Value;
                /**
                    * Loading SaasUsersList.xls to xlSourceWorkbook 
                    * ===========================================
                    */
                xlSourceWorkSheet = (Excel.Worksheet)xlSourceWorkBook.Worksheets.get_Item(1);

                /**
                    * Loading SCOL_customer_addresses_export.xlsx to xlDestinationWorkbook 
                    * ==================================================
                    */
                xlDestinationWorkSheet = (Excel.Worksheet)xlDestinationWorkBook.Worksheets.get_Item(1);

                //MessageBox.Show("Appending " + sourceFileName + " to " + destinationFileName);

                /**
                    * Loading Source to Destination 
                    * ==================================================
                    */

                //xlSourceWorkSheet.Select(Type.Missing);
                xlDestinationWorkSheet.Select(Type.Missing);


                /**
                    * Get exact range for Inser to shift down properly
                    * ================================================
                    */
                String sRange = "A1:D" + xlSourceWorkSheet.UsedRange.Rows.Count.ToString();

                /**
                    * Copy the source and insert to Destination
                    * =========================================
                    */
                Excel.Range rangeSource = xlSourceWorkSheet.get_Range(sRange, Type.Missing);
                Excel.Range rangeDestination = xlDestinationWorkSheet.get_Range("A:D", Type.Missing);


                /**
                    * Copy the source and insert to Destination
                    * =========================================
                    */
                rangeSource.Copy();
                rangeDestination.Insert(Excel.XlInsertShiftDirection.xlShiftDown);
                xlDestinationWorkBook.Save();

                /**
                    * Release excel objects
                    * ========================
                    */
                rangeSource.ReleaseObject();
                rangeDestination.ReleaseObject();
                xlDestinationWorkSheet.ReleaseObject();
                xlSourceWorkSheet.ReleaseObject();
            }
        } // appendExcelToTab

        #endregion
        // excelFile.Worksheets.Add();
        #region Inner Methods
        private static void ReleaseExcelAndFile(Excel.Application excelApplication, Excel.Workbook excelFile, Excel.Worksheet excelTab)
        {
            excelTab.ReleaseObject();
            excelFile.Close();
            excelFile.ReleaseObject();
            excelApplication.Quit();
            excelApplication.ReleaseObject();
            return;
        }

        static private void UnpreparedInflateTableToFile(Excel.Worksheet excelTab, string[,] pStringTable, bool pWarpText = true)
        {
            for (int width = 0; width < pStringTable.GetLength(1); width++)
            {
                for (int height = 0; height < pStringTable.GetLength(0); height++)
                {
                    excelTab.Cells[height + 1, width + 1] = pStringTable[height, width];
                }
            }
            excelTab.Cells.WrapText = pWarpText;
        }

        private static void SaveFile(Excel.Application excelApplication, Excel.Workbook excelFile, string pFileAddress)
        {
            bool failedSave = true;
            while (failedSave)
            {
                try
                {
                    excelApplication.DisplayAlerts = false;
                    
                    excelFile.SaveAs(pFileAddress, AddToMru: false, ConflictResolution: Excel.XlSaveConflictResolution.xlLocalSessionChanges);
                    failedSave = false;
                }
                catch (Exception ex)
                {
                    throw new ExcelFileSaveException(string.Format("Couldn't save file {0}. Please close the file and try again. Exception: {1}\n", pFileAddress,ex));
                }
            }
        }

        private static IEnumerable<string> GetWorksheetNames(this ExcelFile pExcelFile)
        {
            if (pExcelFile.Worksheets.Count!=0)
            {
                foreach (Excel.Worksheet worksheet in pExcelFile.Worksheets)
                {
                    yield return worksheet.Name;
                }
            }
        }
        #endregion
    }
}
