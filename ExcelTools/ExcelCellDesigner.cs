using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelTools
{
    public class ExcelCellDesigner
    {
        /// <summary>
        /// Clears each cell in range
        /// </summary>
        /// <param name="pExcelFile"></param>
        /// <param name="pTabName"></param>
        /// <param name="pTopLeftCell">First cell in the range, from top left to bottom right</param>
        /// <param name="pBottomRightCell">Last cell in the range, from top left to bottom right</param>
        public static void ClearRange(string pExcelFile, string pTabName ,Cell pTopLeftCell, Cell pBottomRightCell)
        {
            if (pTopLeftCell == null || pBottomRightCell == null)
            {
                throw new Exception("Cells can't be null");
            }
            using (var excelApplication = new ExcelApplication())
            {
                using (var excelFile = new ExcelFile(excelApplication, pExcelFile, pReadOnly: false, pEditable: true))
                {
                    Excel.Worksheet excelTab;
                    try
                    {
                        excelTab = excelFile.Worksheets[pTabName];
                    }
                    catch (Exception e)
                    {
                        throw e;
                    }
                    try
                    {

                        var range = excelTab.Cells.Range[pTopLeftCell.ToIndex(), pBottomRightCell.ToIndex()];
                        range.Clear();
                        excelFile.Worksheet.Save();
                        excelTab.ReleaseObject();
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
        }

        /// <summary>
        /// Set color of each cell in range to white
        /// </summary>
        /// <param name="pExcelFile"></param>
        /// <param name="pTabName"></param>
        /// <param name="pTopLeftCell">First cell in the range, from top left to bottom right</param>
        /// <param name="pBottomRightCell">Last cell in the range, from top left to bottom right</param>
        public static void RemoveRangeColor(string pExcelFile, string pTabName, Cell pTopLeftCell, Cell pBottomRightCell)
        {
            if (pTopLeftCell == null || pBottomRightCell == null)
            {
                throw new Exception("Cells can't be null");
            }
            ChangeRangeFormat(pExcelFile, pTabName, pTopLeftCell, pBottomRightCell, null);
        }

        /// <summary>
        /// Colors a range in requested file/table, 
        /// using a cell from the table, as an example.
        /// </summary>
        /// <param name="pExcelFile"></param>
        /// <param name="pTabName"></param>
        /// <param name="pTopLeftCell">First cell in the range, from top left to bottom right</param>
        /// <param name="pBottomRightCell">Last cell in the range, from top left to bottom right</param>
        /// <param name="pCellBaseColor">Coordinates of a cell as base color.</param>
        public static void BrushRangeWithFormat(string pExcelFile, string pTabName, Cell pTopLeftCell, Cell pBottomRightCell, Cell pCellBaseColor)
        {
            if (pTopLeftCell==null || pBottomRightCell == null || pCellBaseColor==null)
            {
                throw new Exception("Cells can't be null");
            }
            ChangeRangeFormat(pExcelFile, pTabName, pTopLeftCell, pBottomRightCell, pCellBaseColor);
        }


        private static void ChangeRangeFormat(string pExcelFile, string pTabName, Cell pTopLeftCell, Cell pBottomRightCell, Cell pCellBaseColor)
        {
            using (var excelApplication = new ExcelApplication())
            {
                using (var excelFile = new ExcelFile(excelApplication, pExcelFile, pReadOnly: false, pEditable: true))
                {
                    Excel.Worksheet excelTab;
                    try
                    {
                        excelTab = excelFile.Worksheets[pTabName];
                    }
                    catch (Exception e)
                    {
                        throw e;
                    }
                    try
                    {

                        var range = excelTab.Cells.Range[pTopLeftCell.ToIndex(), pBottomRightCell.ToIndex()];

                        if (pCellBaseColor!=null)
                        {
                            var baseCell = excelTab.Cells.Range[pCellBaseColor.ToIndex(), pCellBaseColor.ToIndex()];
                            baseCell.Copy();
                            range.PasteSpecial(Excel.XlPasteType.xlPasteFormats);
                        }
                        else
                        {
                            range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
                        }
                        excelFile.Worksheet.Save();
                        excelTab.ReleaseObject();
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
        }
       
    }
}
