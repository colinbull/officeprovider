namespace Office.Word
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.Linq;

    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Wordprocessing;

    public class WordTable
    {
        public string[] Headers { get; private set; }
        public object[] Rows { get; private set; }

        public WordTable(string[] headers, object[] data)
        {
            this.Headers = headers;
            this.Rows = data;
        }

        public WordTable(object[] data)
        {
            this.Headers = this.GetHeaders(data).ToArray();
            this.Rows = data;
        }

        private IEnumerable<string> GetHeaders(object[] data)
        {
            if (data.Length > 0)
            {
                foreach (PropertyDescriptor prop in TypeDescriptor.GetProperties(data[0].GetType()))
                {
                    if (String.IsNullOrEmpty(prop.DisplayName))
                    {
                        yield return prop.Name;
                    }
                    else
                    {
                        yield return prop.DisplayName;
                    }
                }
            }
        }

        private TableRow CreateTableRow(object rowData)
        {
            var tr = new TableRow();

            var cells = 
                TypeDescriptor.GetProperties(rowData.GetType())
                    .OfType<PropertyDescriptor>()
                    .SelectMany(props =>
                        {
                            var value = props.GetValue(rowData);
                            if (value == null)
                            {
                                return new OpenXmlElement[]{};
                            }

                            var tc = new TableCell();
                            var content = new Paragraph(new Run(new Text(value.ToString())));
                            tc.Append(new OpenXmlElement[]{content});
                            return new []{(OpenXmlElement)tc};
                        }).ToArray();
            
            tr.Append(cells);
            return tr;
        }

        public Table ToWordTable()
        {
            var table = new Table();

            var props = new TableProperties(
                new TableBorders(
                    new TopBorder
                        {
                            Val = new EnumValue<BorderValues>(BorderValues.Single),
                            Size = 12
                        },
                    new BottomBorder
                        {
                            Val = new EnumValue<BorderValues>(BorderValues.Single),
                            Size = 12
                        },
                    new LeftBorder
                        {
                            Val = new EnumValue<BorderValues>(BorderValues.Single),
                            Size = 12
                        },
                    new RightBorder
                        {
                            Val = new EnumValue<BorderValues>(BorderValues.Single),
                            Size = 12
                        },
                    new InsideHorizontalBorder
                        {
                            Val = new EnumValue<BorderValues>(BorderValues.Single),
                            Size = 12
                        },
                    new InsideVerticalBorder
                        {
                            Val = new EnumValue<BorderValues>(BorderValues.Single),
                            Size = 12
                        }));
            table.AppendChild<TableProperties>(props);

            foreach (var row in this.Rows)
            {
                var tr = this.CreateTableRow(row);
                table.Append(tr);
            }

            return table;
        }
        
    }
}