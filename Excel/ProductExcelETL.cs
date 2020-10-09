using ClosedXML.Excel;
using Google.Cloud.Firestore;
using OfficeFireSync.Utilities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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

        protected override string MapDocumentId(DocumentSnapshot document)
        {
            document.TryGetValue(PrimaryKey, out string primaryValue);
            document.TryGetValue("category", out string categoryValue);
            return primaryValue + categoryValue;
        }

        protected override string MapDocumentId(IDictionary<string, object> document)
        {
            return document[PrimaryKey] as string + document["category"] as string;
        }
    }
}
