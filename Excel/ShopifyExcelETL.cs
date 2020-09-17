using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;

namespace OfficeFireSync.Excel
{
    public class ShopifyExcelETL : SheetBasedExcelETL
    {
        protected override Regex Separators => separaters;
        private readonly Regex separaters = new Regex(@"[//\d]");
        protected override string PrimaryKey => "handle";

        public ShopifyExcelETL(ImagePreprocessor imagePreprocessor) : base(imagePreprocessor)
        {

        }
    }
}
