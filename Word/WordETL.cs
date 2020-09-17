using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Google.Cloud.Firestore;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeFireSync.Word
{
    public class Section
    {
        public string paragraph;
        public IList<Section> body;
    }

    public class WordETL
    {
        private readonly FirestoreDb db;
        private WriteBatch batch;
        private CollectionReference collectionRef;
        private string documentPath;
        protected IDictionary<string, string> documentIds;

        public WordETL()
        {
            var project = Environment.GetEnvironmentVariable("PROJECT_ID");
            db = FirestoreDb.Create(project);
        }

        public IList<Section> Extract()
        {
            Stream stream = File.Open(documentPath, FileMode.Open);
            using WordprocessingDocument wordDoc = WordprocessingDocument.Open(stream, false);

            Body body = wordDoc.MainDocumentPart.Document.Body;
            var documents = new List<Section>();

            foreach (var paragraph in body.OfType<Paragraph>())
            {
                if(paragraph.ParagraphProperties == null ||
                    paragraph.ParagraphProperties.ParagraphStyleId == null)
                {
                    if(paragraph.InnerText != null)
                    {
                        var heading1 = documents.LastOrDefault();
                        if (heading1 != null && heading1.body != null)
                        {
                            var heading2 = heading1.body.LastOrDefault();
                            if (heading2 != null && heading2.body != null)
                            {
                                var heading3 = heading2.body.LastOrDefault();
                                if (heading3 != null && heading3.body != null)
                                {
                                    var heading4 = heading3.body.LastOrDefault();
                                    if (heading4 != null && heading4.body != null)
                                    {
                                        heading4.body.Add(new Section { paragraph = paragraph.InnerText });
                                    } else
                                    {
                                        heading3.body.Add(new Section { paragraph = paragraph.InnerText });
                                    }
                                } else
                                {
                                    heading2.body.Add(new Section { paragraph = paragraph.InnerText });
                                }
                            } else
                            {
                                heading1.body.Add(new Section { paragraph = paragraph.InnerText });
                            }
                        } else
                        {
                            documents.Add(new Section { paragraph = paragraph.InnerText });
                        }
                    } else
                    {
                        continue;
                    }
                } else
                {
                    var paragraphStyle = paragraph.ParagraphProperties.ParagraphStyleId.Val.Value;

                    if (paragraphStyle.Contains("Heading1"))
                    {
                        documents.Add(new Section
                        {
                            paragraph = paragraph.InnerText,
                            body = new List<Section>()
                        });
                    }
                    else if (paragraphStyle.Contains("Heading2"))
                    {
                        documents.Last().body.Add(new Section
                        {
                            paragraph = paragraph.InnerText,
                            body = new List<Section>()
                        });
                    }
                    else if (paragraphStyle.Contains("Heading3"))
                    {
                        documents.Last().body.Last().body.Add(new Section
                        {
                            paragraph = paragraph.InnerText,
                            body = new List<Section>()
                        });
                    }
                    else if (paragraphStyle.Contains("Heading4"))
                    {
                        documents.Last().body.Last().body.Last().body.Add(new Section
                        {
                            paragraph = paragraph.InnerText,
                            body = new List<Section>()
                        });
                    } else
                    {
                        throw new NotImplementedException("Unsupported style found!");
                    }
                }
            }

            return documents;
            //var paragraphs = body.OfType<Paragraph>()
            //    .Where(p => p.ParagraphProperties != null &&
            //                p.ParagraphProperties.ParagraphStyleId != null &&
            //                p.ParagraphProperties.ParagraphStyleId.Val.Value.Contains("Heading1"))
            //    .ToList()
            //    .Select(el => el.InnerText);
        }

        public void Transform()
        {

        }

        public async Task SyncToFirebase(string filePath)
        {
            //var primaryKey = "name";
            //batch = db.StartBatch();
            //documentIds = await GetDocumentIds("content", primaryKey);

            this.documentPath = filePath;
            var documents = Extract();
            var primaryKey = "name";
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
    }
}
