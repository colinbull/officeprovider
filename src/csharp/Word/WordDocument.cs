namespace EonUk.Retail.Ice.NonSupplyCustomerServicingModule.Helpers.Office.Word
{
    using System.Collections;
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.IO;
    using System.Linq;

    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Wordprocessing;

    public class WordDocument : IMappableDocument
    {
        public WordFile WordFile { get; private set; }

        private WordprocessingDocument Document { get; set; }

        private ILookup<string, SdtElement> MappableElements { get; set; }

        public WordDocument(string wordFilePath)
        {
            this.WordFile = WordFile.Create(wordFilePath, false);
            this.Document = WordprocessingDocument.Open(this.WordFile.FilePath, true);

            this.MappableElements = this.GetMappableElements().ToLookup(r => r.SdtProperties.GetFirstChild<Tag>().Val.Value, r => r);
        }

        public WordDocument(Stream stream, string path)
        {
            this.WordFile = WordFile.Create(stream, path);
            this.Document = WordprocessingDocument.Open(this.WordFile.FilePath, true);

            this.MappableElements = this.GetMappableElements().ToLookup(r => r.SdtProperties.GetFirstChild<Tag>().Val.Value, r => r);
        }

        private IEnumerable<SdtElement> GetMappableElements()
        {
            foreach (var descendant in this.Document.MainDocumentPart.Document.Descendants<SdtElement>())
            {
                yield return descendant;
            }

            foreach (var descendant in this.Document.MainDocumentPart.HeaderParts.SelectMany(headerPart => headerPart.Header.Descendants<SdtElement>()))
            {
                yield return descendant;
            }

            foreach (var descendant in this.Document.MainDocumentPart.FooterParts.SelectMany(footerPart => footerPart.Footer.Descendants<SdtElement>()))
            {
                yield return descendant;
            }

            if (this.Document.MainDocumentPart.FootnotesPart != null)
            {
                foreach (var cc in this.Document.MainDocumentPart.FootnotesPart.Footnotes.Descendants<SdtElement>())
                {
                    yield return cc;
                }
            }

            if (this.Document.MainDocumentPart.EndnotesPart != null)
            {
                foreach (var cc in this.Document.MainDocumentPart.EndnotesPart.Endnotes.Descendants<SdtElement>())
                {
                    yield return cc;
                }
            }
        }

        public void MapNamedRangesTo<T>(T instance)
        {
            var props = TypeDescriptor.GetProperties(instance);
            foreach (var element in this.MappableElements)
            {
                var matchingProperty = props.Find(element.Key, true);

                if (matchingProperty != null)
                {
                    var value = matchingProperty.GetValue(instance);
                    if (matchingProperty.PropertyType.IsArray || matchingProperty.PropertyType == typeof(WordTable))
                    {
                        var table = 
                            matchingProperty.PropertyType.IsArray
                                ? new WordTable(((IEnumerable)value).Cast<object>().ToArray()) 
                                : ((WordTable)value);

                        foreach (var target in element)
                        {
                            var t = table.ToWordTable();
                            target.AppendChild(t);
                        }
                    }
                    else
                    {
                        if (value != null)
                        {
                            foreach (var target in element)
                            {
                                var para = target.Descendants<Paragraph>().First();
                                if (para != null)
                                {
                                    // simple text, assume 1 run to handle
                                    Run r = para.Elements<Run>().First();
                                    if (r != null)
                                    {
                                        Text t = r.Elements<Text>().First();
                                        if (t == null)
                                        {
                                            t = new Text(value.ToString());
                                            r.Append(t);
                                        }
                                        else
                                        {
                                            t.Text = value.ToString();
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }

        public void Commit(string outputPath)
        {
            this.Document.MainDocumentPart.Document.Save();
            this.Document.Close();
            this.WordFile.Commit(outputPath);
        }

        public void Dispose()
        {
            this.WordFile.Dispose();
        }
    }
}