using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelOp
{
    public static internal class ExcelFormat
    {

        /// <summary>
        /// Excelの列番号をアルファベット（例えば、1ならA、2ならB、27ならAA）に変換する
        /// </summary>
        /// <param name="columnNumber"></param>
        /// <returns></returns>
        static string ColumnNumberToName(int columnNumber)
        {
            string columnName = string.Empty;
            while (columnNumber > 0)
            {
                int modulo = (columnNumber - 1) % 26;
                columnName = Convert.ToChar('A' + modulo) + columnName;
                columnNumber = (columnNumber - modulo) / 26;
            }
            return columnName;
        }
    }
}
