using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NamLongGroup
{
    class CurrencyFormater
    {
        public static String format(Double rowValue)
        {
            return rowValue.ToString("C", CultureInfo.CurrentCulture).Split('$')[1].Split('.')[0];
        }
    }
}
