using ClosedXML.Excel;
using OfficeFireSync.Utilities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace OfficeFireSync.Excel
{
    public class SheetBasedExcelETL : ExcelETL
    {
        private readonly Regex separators = new Regex(@"[\W\d]");
        protected virtual Regex Separators => separators;
        protected override string PrimaryKey => "color";

        public SheetBasedExcelETL(ImagePreprocessor imagePreprocessor) : base(imagePreprocessor)
        {

        }

        protected override void SyncWorksheet(IXLWorksheet worksheet, string primaryKey)
        {
            var rows = worksheet
                .Rows()
                .Skip(1);
            var columnHeads = worksheet.Rows()
                .First()
                .Cells()
                .Select(el => (el.Value as string).ToCamel())
                .ToList();

            foreach (var row in rows)
            {
                var r = worksheet.Range(row.FirstCell(), row.Cell(columnHeads.Count() - 1)).FirstRow();
                var document = RowToDocument(r, columnHeads);

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

        private string NameSection(string name, int section)
        {
            var sections = Separators.Split(name);
            return sections[section];
        }

        private int SectionIndex(string name, int section)
        {
            var matches = Separators.Matches(name);
            return Int32.Parse(matches[section].Value);
        }

        protected override IDictionary<string, object> RowToDocument(IXLRangeRow row, IList<string> fieldNames)
        {
            var tester = "Option1Name";
            var a = separators.Matches(tester);
            var b = a[0].Value;

            var cellGroups = fieldNames
                .Select((value, index) => new { index, value = value.Replace(" ", "") })
                .ToDictionary(
                    el => el.value,
                    el => row.Cell(el.index + 1)
                );
            return ColumnsToMap(cellGroups, 0);
        }

        protected virtual IDictionary<string, object> ColumnsToMap(IDictionary<string, IXLCell> columns, int section)
        {
            var map = new Dictionary<string, object>();
            var cellGroups = columns.GroupBy(el => NameSection(el.Key, section));

            foreach (var cellGroup in cellGroups)
            {
                var startIndex = cellGroup.First().Key.IndexOf(cellGroup.Key);
                var endIndex = startIndex + cellGroup.Key.Length;

                if (cellGroup.First().Key.Length > endIndex) // Not the last section
                {
                    var segment = cellGroup.First().Key.Substring(endIndex, 1);
                    if (Int32.TryParse(segment, out int result)) // List
                    {
                        map.Add(
                            cellGroup.Key.ToCamel(),
                            ColumnsToList(cellGroup.ToDictionary(el => el.Key, el => el.Value), section)
                        );
                    }
                    else // Map
                    {
                        map.Add(
                            cellGroup.Key.ToCamel(),
                            ColumnsToMap(cellGroup.ToDictionary(el => el.Key, el => el.Value), section + 1)
                        );
                    }
                }
                else if (cellGroup.Count() == 1)
                {
                    map.Add(
                        cellGroup.Key.ToCamel(),
                        CellToField(map, cellGroup.First().Value)
                    );
                }
                else
                {
                    throw new FormatException("ERROR: Heading column format unsupported");
                }
            }

            return map;
        }

        protected virtual IList<object> ColumnsToList(IDictionary<string, IXLCell> columns, int section)
        {
            var list = new List<object>();
            var cellGroups = columns.GroupBy(el => SectionIndex(el.Key, section)).OrderBy(el => el.Key);

            foreach (var cellGroup in cellGroups)
            {
                if (cellGroup.Count() != 1) // Not the last section
                {
                    list.Add(
                        ColumnsToMap(cellGroup.ToDictionary(el => el.Key, el => el.Value), section + 1)
                    );
                }
                else
                {
                    list.Add(CellToField(new Dictionary<string, object>(), cellGroup.First().Value));
                }
            }

            return list;
        }
    }
}
