using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
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

            await batch.CommitAsync();

            //var products = new List<object>();

            //foreach (var worksheet in workbook.Worksheets)
            //{
            //    var documents = ExtractDocumentsFromWorkSheet(workbook.Worksheet(worksheet.Name));
            //    products.AddRange(documents);
            //}

            //await SendDocuments(new Stack(products), "product");
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
            var fieldNames = table.Rows()
                .First()
                .Cells()
                .Select(el => ((string)el.Value).ToCamel())
                .ToList();
            
            foreach (var row in rows)
            {
                var document = RowToDocument(row, fieldNames);
   
                if (documentIds.ContainsKey((string)document[primaryKey]))
                {
                    var id = documentIds[(string)document[primaryKey]];
                    batch.Set(collectionRef.Document(id), document);
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


        //public async Task SendDocuments(Stack documents, string collectionName)
        //{
        //    WriteBatch batch = db.StartBatch();
        //    var collection = db.Collection(collectionName);

        //    foreach (var document in documents)
        //    {
        //        batch.Set(collection.Document(), document);
        //    }

        //    await batch.CommitAsync();
        //}

        //public IList<object> ExtractDocumentsFromWorkSheet(IXLWorksheet worksheet)
        //{
        //    var sheetName = worksheet.Name.Replace(" ", "");
        //    var primaryTable = worksheet.Tables.First(el => el.Name == sheetName);
        //    var subTables = worksheet.Tables
        //        .Where(el => el.Name != primaryTable.Name)
        //        .ToDictionary(el => el.Name, el => el);

        //    return TableToList(
        //        primaryTable,
        //        (dict, row) => {
        //            foreach (var subTable in subTables)
        //            {
        //                var fieldName = subTable.Key.Replace(sheetName, "").ToCamel();
        //                var fieldValues = TableToList(subTable.Value, (string)row.Cells().First().Value);

        //                dict.Add(
        //                    fieldName,
        //                    fieldValues.Count() == 1 ? fieldValues[0] : fieldValues
        //                );
        //            }
        //        }
        //    );
        //}

        //private IList<object> TableToList(IXLTable table, string primaryKey)
        //{
        //    return TableToList(table, (dict, row) => { }, primaryKey);
        //}

        //private IList<object> TableToList(IXLTable table, Action<IDictionary<string, object>, IXLRangeRow> perRow, string primaryKey)
        //{
        //    var result = new List<object>();
        //    var headers = table.Rows().First().Cells().Select(el => (string)el.Value).ToList();
        //    var rows = table.Rows().Skip(1).Where(el => (string)el.Cell(1).Value == primaryKey);

        //    foreach (var row in rows)
        //    {
        //        var dict = RowToDictionary(row, headers, perRow);
        //        result.Add(dict);
        //    }

        //    return result;
        //}

        //private IList<object> TableToList(IXLTable table, Action<IDictionary<string, object>, IXLRangeRow> perRow)
        //{
        //    var category = table.Name.ToCamel();
        //    var result = new List<object>();
        //    var headers = table.Rows().First().Cells().Select(el => (string)el.Value).ToList();
        //    var rows = table.Rows().Skip(1);

        //    foreach (var row in rows)
        //    {
        //        var dict = RowToDictionary(row, headers, perRow);
        //        dict.Add("category", category);
        //        result.Add(dict);
        //    }

        //    return result;
        //}

        //private IDictionary<string, object> RowToDictionary(IXLRangeRow row, IList<string> headers, Action<IDictionary<string, object>, IXLRangeRow> perRow)
        //{
        //    IDictionary<string, object> dict = new Dictionary<string, object>();
        //    var count = 0;

        //    foreach (var cell in row.Cells())
        //    {
        //        if(headers[count] != foreignKey)
        //        {
        //            var header = headers[count].ToCamel();
        //            if (cell.DataType == XLDataType.DateTime)
        //            {
        //                dict.Add(header, ((DateTime)cell.Value).ToUniversalTime());
        //            }
        //            else if (cell.DataType == XLDataType.Number && cell.Style.NumberFormat.Format == "_-\"$\"* #,##0.00_-;\\-\"$\"* #,##0.00_-;_-\"$\"* \"-\"??_-;_-@_-")
        //            {
        //                dict.Add(header, (double)cell.Value * 100);
        //            }
        //            else
        //            {
        //                if (cell.DataType == XLDataType.Text && Uri.IsWellFormedUriString((string)cell.Value, UriKind.RelativeOrAbsolute))
        //                {
        //                    PreprocessMedia(dict, (string)cell.Value);
        //                }
        //                dict.Add(header, cell.Value);
        //            }
        //        }

        //        count++;
        //    }

        //    perRow(dict, row);
        //    return dict;
        //}

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
