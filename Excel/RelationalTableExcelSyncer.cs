using ClosedXML.Excel;
using OfficeFireSync.Utilities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeFireSync.Excel
{
    public class RelationalTableExcelSyncer : ExcelSyncer
    {
        private string primaryKey;
        private string foreignKey = "FK";
        private string sheetName;
        private Dictionary<string, IXLTable> relatedTables;

        public RelationalTableExcelSyncer(ImagePreprocessor imagePreprocessor) : base(imagePreprocessor)
        {

        }

        protected override void OnWorksheetSync(IXLWorksheet worksheet, string primaryKey)
        {
            this.primaryKey = primaryKey;
            sheetName = worksheet.Name.Replace(" ", "");
            var primaryTable = worksheet.Tables.First(el => el.Name == sheetName);
            relatedTables = worksheet.Tables
                .Where(el => el.Name != primaryTable.Name)
                .ToDictionary(el => el.Name, el => el);

            SyncTable(primaryTable, primaryKey);
        }

        protected override IDictionary<string, object> RowToDocument(IXLRangeRow row, IList<string> headers)
        {
            var document = base.RowToDocument(row, headers);

            foreach (var relatedTable in relatedTables)
            {
                var fieldName = relatedTable.Key.Replace(sheetName, "").ToCamel();
                var fieldValues = SyncRelatedTable(relatedTable.Value, (string)row.Cells().First().Value);

                document.Add(
                    fieldName,
                    fieldValues.Count() == 1 ? fieldValues.First() : fieldValues
                );
            }

            return document;
        }

        protected virtual IDictionary<string, object> RowToField(IXLRangeRow row, IList<string> headers)
        {
            IDictionary<string, object> document = new Dictionary<string, object>();
            var count = 0;

            foreach (var cell in row.Cells())
            {
                if (headers[count] != foreignKey)
                {
                    document.Add(headers[count], CellToField(document, cell));
                }

                count++;
            }

            return document;
        }

        protected virtual IList<object> SyncRelatedTable(IXLTable table, string keyValue)
        {
            var result = new List<object>();
            var headers = table.Rows().First().Cells().Select(el => (string)el.Value).ToList();
            var rows = table.Rows().Skip(1).Where(el => (string)el.Cell(1).Value == keyValue);

            return rows.Select(el => RowToField(el, headers)).ToList<object>();
        }
    }
}
