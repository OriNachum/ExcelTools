using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelTools
{
    public class Cell
    {
        public Cell(int pRow, int pColumn)
        {
            Row = pRow;
            Column = pColumn;
        }

        public string ToIndex()
        {
            var columnIndexByNumbers = new List<int>();
            var divisionResult = Column;
            const int LettersCount = 22;
            while (divisionResult != 0)
            {
                var leftover = divisionResult % LettersCount;
                columnIndexByNumbers.Add(leftover);
                divisionResult = divisionResult / LettersCount;
            }
            var cellIndex = string.Empty;
            foreach (var number in columnIndexByNumbers)
            {
                cellIndex = cellIndex + ((char)('A' + number - 1)).ToString();
            }
            cellIndex = cellIndex + Row.ToString();
            return cellIndex;
        }

        public string value;
        public int Row;
        public int Column;   
    }
}
