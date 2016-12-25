using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenExcel
{
    public enum ExcelDataType : short
    {
        Boolean,
        Date,
        Error,
        Number,
    }

    public enum ExcelDataSourceType : short
    {
        InlineString,
        SharedString,
        Formula
    }
}
