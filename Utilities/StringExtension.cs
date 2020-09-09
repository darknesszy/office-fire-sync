using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeFireSync.Utilities
{
    public static class StringExtension
    {
        public static string ToCamel(this String ob)
        {
            return Char.ToLowerInvariant(ob[0]) + ob.Substring(1).Replace(" ", "");
        }
    }
}
