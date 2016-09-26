using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelTools
{

    public class ExcelException : Exception
    {
        public ExcelException() : base()
        { }

        public ExcelException(string pMessage) : base(pMessage) { }
    }

    public class ExcelFileLoadException : ExcelException
    {
        public ExcelFileLoadException() : base()
        { }

        public ExcelFileLoadException(string pMessage) : base(pMessage) { }   
    }

    public class ExcelFileSaveException : ExcelException
    {
        public ExcelFileSaveException() : base()
        { }

        public ExcelFileSaveException(string pMessage) : base(pMessage) { }
    }
}
