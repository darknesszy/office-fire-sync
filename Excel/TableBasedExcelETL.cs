using ClosedXML.Excel;
using OfficeFireSync.Utilities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeFireSync.Excel
{
    public abstract class TableBasedExcelETL : ExcelETL
    {
        private string primaryKey;
        private string foreignKey = "fK";
        private string sheetName;
        private Dictionary<string, IXLTable> relatedTables;

        public TableBasedExcelETL(ImagePreprocessor imagePreprocessor) : base(imagePreprocessor)
        {

        }

        protected override void SyncWorksheet(IXLWorksheet worksheet, string primaryKey)
        {
            this.primaryKey = primaryKey;
            sheetName = worksheet.Name.Replace(" ", "");
            var primaryTable = worksheet.Tables.First(el => el.Name == sheetName);
            relatedTables = worksheet.Tables
                .Where(el => el.Name != primaryTable.Name)
                .ToDictionary(el => el.Name, el => el);

            TableToCollection(primaryTable, primaryKey);
        }

        protected virtual void TableToCollection(IXLTable table, string primaryKey)
        {
            var rows = table.Rows().Skip(1);
            var columnHeads = table.Rows()
                .First()
                .Cells()
                .Select(el => ((string)el.Value).ToCamel())
                .ToList();

            foreach (var row in rows)
            {
                var document = RowToDocument(row, columnHeads);

                if (documentIds.ContainsKey((string)document[primaryKey]))
                {
                    var id = documentIds[(string)document[primaryKey]];
                    batch.Update(collectionRef.Document(id), document);
                    documentIds.Remove((string)document[primaryKey]);
                    Console.WriteLine($"Updating {document[primaryKey]}");
                }
                else
                {
                    batch.Create(collectionRef.Document(), document);
                    Console.WriteLine($"Creating {document[primaryKey]}");
                }
            }
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

        protected virtual IDictionary<string, object> RowToField(IXLRangeRow row, IList<string> fieldNames)
        {
            IDictionary<string, object> document = new Dictionary<string, object>();
            var count = 0;

            foreach (var cell in row.Cells())
            {
                if (fieldNames[count] != foreignKey)
                {
                    document.Add(fieldNames[count], CellToField(document, cell));
                }

                count++;
            }

            return document;
        }

        protected virtual IList<object> SyncRelatedTable(IXLTable table, string primaryKeyValue)
        {
            var columnHeads = table
                .Rows()
                .First()
                .Cells()
                .Select(el => ((string)el.Value).ToCamel())
                .ToList();
            
            return table
                .Rows()
                .Skip(1)
                .Where(el => (string)el.Cell(1).Value == primaryKeyValue)
                .Select(el => RowToField(el, columnHeads))
                .ToList<object>();
        }
    }
}
