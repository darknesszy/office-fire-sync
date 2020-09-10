using ClosedXML.Excel;
using Google.Cloud.Firestore;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeFireSync.Utilities;
using System.IO;
using System.Collections;

namespace OfficeFireSync.Excel
{
    public abstract class ExcelSyncer
    {
        private readonly string mediaPath = Environment.GetEnvironmentVariable("MEDIA_PATH");
        private readonly FirestoreDb db;
        private readonly ImagePreprocessor imagePreprocessor;
        private WriteBatch batch;
        private CollectionReference collectionRef;
        protected IDictionary<string, string> documentIds;

        public ExcelSyncer(ImagePreprocessor imagePreprocessor)
        {
            this.imagePreprocessor = imagePreprocessor;

            var project = Environment.GetEnvironmentVariable("PROJECT_ID");
            db = FirestoreDb.Create(project);
        }

        public async Task SyncToFireStore()
        {
            var primaryKey = "name";
            batch = db.StartBatch();
            documentIds = await GetDocumentIds("product", primaryKey);
            var workbook = new XLWorkbook(Environment.GetEnvironmentVariable("MANIFEST_PATH"));

            foreach (var worksheet in workbook.Worksheets)
            {
                OnWorksheetSync(worksheet, primaryKey);
            }

            RemoveDanglingDocument();
            await batch.CommitAsync();
            Console.WriteLine("# Synchronise Complete...");
        }

        protected async virtual Task<IDictionary<string, string>> GetDocumentIds(string collectionName, string primaryKey)
        {
            collectionRef = db.Collection(collectionName);
            QuerySnapshot snapshot = await collectionRef.GetSnapshotAsync();
            return snapshot.Documents.ToDictionary(
                el => {
                    el.TryGetValue(primaryKey, out string keyValue);
                    return keyValue;
                },
                el => el.Id
            );
        }

        protected abstract void OnWorksheetSync(IXLWorksheet worksheet, string primaryKey);

        protected virtual void SyncTable(IXLTable table, string primaryKey)
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

        protected virtual IDictionary<string, object> RowToDocument(IXLRangeRow row, IList<string> fieldNames)
        {
            IDictionary<string, object> document = new Dictionary<string, object>();
            var count = 0;

            foreach (var cell in row.Cells())
            {
                document.Add(fieldNames[count], CellToField(document, cell));
                count++;
            }

            return document;
        }

        protected virtual object CellToField(IDictionary<string, object> document, IXLCell cell)
        {
            if (cell.DataType == XLDataType.DateTime)
            {
                return ((DateTime)cell.Value).ToUniversalTime();
            }
            else if (cell.DataType == XLDataType.Number && cell.Style.NumberFormat.Format == "_-\"$\"* #,##0.00_-;\\-\"$\"* #,##0.00_-;_-\"$\"* \"-\"??_-;_-@_-")
            {
                return (double)cell.Value * 100;
            }
            else
            {
                if (cell.DataType == XLDataType.Text && Uri.IsWellFormedUriString((string)cell.Value, UriKind.RelativeOrAbsolute))
                {
                    PreprocessMedia(document, (string)cell.Value);
                }
                return cell.Value;
            }
        }

        private void RemoveDanglingDocument()
        {
            foreach (var keyValue in documentIds)
            {
                batch.Delete(collectionRef.Document(keyValue.Value));
                Console.WriteLine($"Deleting {keyValue.Key}");
            }
        }

        private void PreprocessMedia(IDictionary<string, object> dict, string uri)
        {
            var path = mediaPath + uri;
            if (Path.GetExtension(uri) == ".jpg" && File.Exists(path))
            {
                var image = imagePreprocessor.ResizeImage(path, 45);
                dict.Add("base64", imagePreprocessor.ConvertToBase64(image));
            }
        }
    }
}
