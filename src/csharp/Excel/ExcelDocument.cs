namespace Office.Excel
{
    using System.ComponentModel;
    using System.IO;

    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Wordprocessing;

    public class ExcelDocument : IMappableDocument
    {
        private ExcelFile ExcelFile { get; set; }
        private ExcelWorkbook Workbook { get; set; }

        public ExcelDocument(string path)
        {
            this.ExcelFile = ExcelFile.Create(path, false);
            this.Workbook = new ExcelWorkbook(SpreadsheetDocument.Open(this.ExcelFile.FilePath, true));
        }

        public ExcelDocument(Stream stream, string path)
        {
            this.ExcelFile = ExcelFile.Create(stream, path);
            this.Workbook = new ExcelWorkbook(SpreadsheetDocument.Open(this.ExcelFile.FilePath, true));
        }


        public void MapNamedRangesTo<T>(T instance)
        {
            var props = TypeDescriptor.GetProperties(instance);
            foreach (var nr in this.Workbook.DefinedNames)
            {

                var namedRange = nr;
                var instanceNumberIndex = namedRange.Key.LastIndexOf('?');
                var key = namedRange.Key;

                if (instanceNumberIndex != -1)
                {
                    key = namedRange.Key.Remove(instanceNumberIndex, namedRange.Key.Length - instanceNumberIndex);
                }

                var matchingProperty = props.Find(key, true);

                var sheet = this.Workbook.Sheets[namedRange.Value.Address.SheetName];

                if (matchingProperty != null)
                {
                    var value = matchingProperty.GetValue(instance);
                    if (matchingProperty.PropertyType == typeof(ExcelPagedArray))
                    {
                        sheet.WritePagedArray(namedRange.Value, (ExcelPagedArray)value);
                    }
                    else if (matchingProperty.PropertyType == typeof(string[][]))
                    {
                        sheet.WriteArray2D(namedRange.Value, (string[][])value);
                    } 
                    else if (matchingProperty.PropertyType == typeof(string[]))
                    {
                        sheet.WriteArray(namedRange.Value, (string[])value);
                    }
                    else
                    {
                        sheet.WriteValue(namedRange.Value.Address.StartCell, value);
                    }
                }
            }
        }

        public void Commit(string outputPath)
        {
            this.Workbook.Close();
            this.ExcelFile.Commit(outputPath);
        }

        public void Dispose()
        {
            this.ExcelFile.Dispose();
        }
    }
}