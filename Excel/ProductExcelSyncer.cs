using ClosedXML.Excel;
using OfficeFireSync.Utilities;
using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeFireSync.Excel
{
    public class ProductExcelSyncer : RelationalTableExcelSyncer
    {
        private string category;

        public ProductExcelSyncer(ImagePreprocessor imagePreprocessor) : base(imagePreprocessor)
        {

        }

        protected override IDictionary<string, object> RowToDocument(IXLRangeRow row, IList<string> headers)
        {
            var document = base.RowToDocument(row, headers);
            document.Add("category", category);

            return document;
        }

        protected override void SyncTable(IXLTable table, string primaryKey)
        {
            category = table.Name.ToCamel();
            base.SyncTable(table, primaryKey);
        }
    }
}
