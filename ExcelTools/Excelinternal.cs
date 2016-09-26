using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelTools
{
    internal static class ExcelInternal
    {
        static internal int ReleaseObject(this object obj)
        {
            int result;
            try
            {
                result = System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                throw new Exception("Exception Occured while releasing object: " + ex.ToString());
            }
            return result;
        }
    }
}
