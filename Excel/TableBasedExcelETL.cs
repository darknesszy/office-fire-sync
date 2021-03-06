﻿using ClosedXML.Excel;
using Google.Cloud.Firestore;
using OfficeFireSync.Utilities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeFireSync.Excel
{
    public abstract class TableBasedExcelETL : ExcelETL
    {
        private string sheetName;
        private string foreignKey;
        private IXLTable primaryTable;
        private Dictionary<string, IXLTable> relatedTables;

        public TableBasedExcelETL(ImagePreprocessor imagePreprocessor) : base(imagePreprocessor)
        {

        }

        protected async override Task<IDictionary<string, string>> GetDocumentIds(string collectionName, string primaryKey)
        {
            collectionRef = db.Collection(collectionName);
            try
            {
                QuerySnapshot snapshot = await collectionRef.GetSnapshotAsync();
                return snapshot.Documents.ToDictionary(
                    el => MapDocumentId(el),
                    el => el.Id
                );
            }
            catch (Exception caughtEx)
            {
                throw new Exception("Unknown Exception Thrown: "
                       + "\n  Type:    " + caughtEx.GetType().Name
                       + "\n  Message: " + caughtEx.Message);
            }
        }

        protected abstract string MapDocumentId(DocumentSnapshot document);
        protected abstract string MapDocumentId(IDictionary<string, object> document);

        protected override void SyncWorksheet(IXLWorksheet worksheet, string primaryKey)
        {
            foreignKey = worksheet.Name.ToCamel();
            sheetName = worksheet.Name.Replace(" ", "");

            primaryTable = worksheet.Tables.First(el => el.Name == sheetName);
            relatedTables = worksheet.Tables
                .Where(el => el.Name != primaryTable.Name)
                .ToDictionary(el => el.Name, el => el);

            TableToCollection(primaryTable, primaryKey);
        }

        protected virtual void TableToCollection(IXLTable table, string primaryKey)
        {
            var rows = table.Rows().Skip(1); // First row is the Column Head.
            var columnHeads = table.Rows()
                .First()
                .Cells()
                .Select(el => (el.Value as string).ToCamel())
                .ToList();

            foreach (var row in rows)
            {
                var document = RowToDocument(row, columnHeads);
                var key = MapDocumentId(document);

                if (existingDocumentIds.ContainsKey(key))
                {
                    var id = existingDocumentIds[key];
                    batch.Update(collectionRef.Document(id), document);
                    existingDocumentIds.Remove(key);
                    Console.WriteLine($"Updating {key}");
                }
                else
                {
                    batch.Create(collectionRef.Document(), document);
                    Console.WriteLine($"Creating {key}");
                }
            }
        }

        protected override IDictionary<string, object> RowToDocument(IXLRangeRow row, IList<string> fieldNames)
        {
            var document = base.RowToDocument(row, fieldNames);

            foreach (var relatedTable in relatedTables)
            {
                var fieldName = relatedTable.Key.Replace(sheetName, "").ToCamel();
                var fieldValues = SyncRelatedTable(relatedTable.Value, row.Cells().First().Value as string);

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
            var columnHeads = table.Rows()
                .First()
                .Cells()
                .Select(el => (el.Value as string).ToCamel())
                .ToList();
            
            return table.Rows()
                .Skip(1)
                .Where(el => el.Cell(1).Value as string == primaryKeyValue)
                .Select(el => RowToField(el, columnHeads))
                .ToList<object>();
        }
    }
}
