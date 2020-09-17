using ClosedXML.Excel;
using OfficeFireSync.Utilities;
using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeFireSync.Excel
{
    public class ProductExcelETL : TableBasedExcelETL
    {
        private string category;
        protected override string PrimaryKey { get => "name"; }

        public ProductExcelETL(ImagePreprocessor imagePreprocessor) : base(imagePreprocessor)
        {

        }

        protected override IDictionary<string, object> RowToDocument(IXLRangeRow row, IList<string> headers)
        {
            var document = base.RowToDocument(row, headers);
            document.Add("category", category);

            return document;
        }

        protected override void TableToCollection(IXLTable table, string primaryKey)
        {
            category = table.Name.ToCamel();
            base.TableToCollection(table, primaryKey);
        }
    }
}
