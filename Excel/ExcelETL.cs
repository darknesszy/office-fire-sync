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
    public abstract class ExcelETL
    {
        private readonly string mediaPath = Environment.GetEnvironmentVariable("MEDIA_PATH");
        private readonly FirestoreDb db;
        private readonly ImagePreprocessor imagePreprocessor;
        protected WriteBatch batch;
        protected CollectionReference collectionRef;
        protected IDictionary<string, string> documentIds;
        protected abstract string PrimaryKey { get; }

        public ExcelETL(ImagePreprocessor imagePreprocessor)
        {
            this.imagePreprocessor = imagePreprocessor;

            var project = Environment.GetEnvironmentVariable("PROJECT_ID");
            db = FirestoreDb.Create(project);
        }

        public async Task SyncToFireStore(string filePath, string collectionName)
        {
            batch = db.StartBatch();
            documentIds = await GetDocumentIds(collectionName, PrimaryKey);

            var workbook = new XLWorkbook(filePath);

            foreach (var worksheet in workbook.Worksheets)
            {
                SyncWorksheet(worksheet, PrimaryKey);
            }

            RemoveDanglingDocument();
            await batch.CommitAsync();
            Console.WriteLine("# Synchronise Complete...");
        }

        protected async virtual Task<IDictionary<string, string>> GetDocumentIds(string collectionName, string primaryKey)
        {
            collectionRef = db.Collection(collectionName);
            try
            {
                QuerySnapshot snapshot = await collectionRef.GetSnapshotAsync();
                return snapshot.Documents.ToDictionary(
                    el => {
                        el.TryGetValue(primaryKey, out string keyValue);
                        return keyValue;
                    },
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

        protected abstract void SyncWorksheet(IXLWorksheet worksheet, string primaryKey);

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
