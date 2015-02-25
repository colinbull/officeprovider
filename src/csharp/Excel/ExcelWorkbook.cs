namespace EonUk.Retail.Ice.NonSupplyCustomerServicingModule.Helpers.Office.Excel
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;

    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Spreadsheet;

    public class ExcelWorkbook
    {
        public SpreadsheetDocument Document { get; private set; }
        public IDictionary<string, ExcelWorksheet> Sheets { get; private set; }
        public IDictionary<string, Addressable<DefinedName, ExcelAddress.NamedRange>> DefinedNames { get; private set; } 


        public static readonly HashSet<Type> NumericTypes = new HashSet<Type>
                                                                {
                                                                    typeof(decimal), typeof(byte), typeof(sbyte),
                                                                    typeof(short), typeof(ushort), typeof(uint), typeof(ulong),
                                                                    typeof(int), typeof(long), typeof(decimal),
                                                                    typeof(double), typeof(float)
                                                                };

        private readonly WorkbookPart workbookPart;

        public ExcelWorkbook(SpreadsheetDocument doc)
        {
            this.Document = doc;
            this.workbookPart = doc.WorkbookPart;
            this.DefinedNames = this.GetDefinedNames();
            this.Sheets = this.workbookPart
                        .Workbook
                        .Descendants<Sheet>()
                        .ToDictionary(s => s.Name.Value, s => new ExcelWorksheet(this, s.Name.Value, (WorksheetPart)this.Document.WorkbookPart.GetPartById(s.Id)));
        }

        public Addressable<DefinedName, ExcelAddress.NamedRange> TryGetDefinedName(string name, Addressable<DefinedName, ExcelAddress.NamedRange> defaultValue)
        {
            if (string.IsNullOrEmpty(name))
            {
                return defaultValue;
            }

            Addressable<DefinedName, ExcelAddress.NamedRange> retValue;
            if (this.DefinedNames.TryGetValue(name, out retValue))
            {
                return retValue;
            }
            else
            {
                return defaultValue;
            }
        }

        private IDictionary<string, Addressable<DefinedName, ExcelAddress.NamedRange>> GetDefinedNames()
        {
            var names = new Dictionary<string, Addressable<DefinedName, ExcelAddress.NamedRange>>();
            var dns = this.workbookPart.Workbook.DefinedNames;

            if (dns != null)
            {
                foreach (var openXmlElement in dns)
                {
                    var dn = (DefinedName)openXmlElement;
                    if (!dn.Name.Value.Contains("Print_Area"))
                    {
                        names.Add(dn.Name.Value, Addressable<DefinedName, ExcelAddress.NamedRange>.Lift(dn));
                    }
                }
            }

            return names;
        }

        public int InsertSharedStringItem(string value)
        {
            int index = 0;
            bool found = false;
            var stringTablePart = this.workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();

            if (stringTablePart == null)
            {
                stringTablePart = this.workbookPart.AddNewPart<SharedStringTablePart>();
            }

            var stringTable = stringTablePart.SharedStringTable;
            if (stringTable == null)
            {
                stringTable = new SharedStringTable();
            }

            foreach (SharedStringItem item in stringTable.Elements<SharedStringItem>())
            {
                if (item.InnerText == value)
                {
                    found = true;
                    break;
                }
                index += 1;
            }

            if (!found)
            {
                stringTable.AppendChild(new SharedStringItem(new Text(value)));
                stringTable.Save();
            }

            return index;
        }

        public ExcelWorksheet CopySheet(string sheetName, string copiedSheetName)
        {
            SpreadsheetDocument tempdoc = SpreadsheetDocument.Create(new MemoryStream(), this.Document.DocumentType);
            var tempWorkbook = tempdoc.AddWorkbookPart();
            var sourceSheetPart = this.Sheets[sheetName].WorksheetPart;
            var tempWorksheet = tempWorkbook.AddPart<WorksheetPart>(sourceSheetPart);
            var clonedSheet = this.workbookPart.AddPart(tempWorksheet);

            int numTableDefParts = sourceSheetPart.GetPartsCountOfType<TableDefinitionPart>();
            int tableId = numTableDefParts;

            //Clean up table definition parts (tables need unique ids)
            if (numTableDefParts != 0)
            {
                foreach (TableDefinitionPart tableDefPart in clonedSheet.TableDefinitionParts)
                {
                    tableId++;
                    tableDefPart.Table.Id = (uint)tableId;
                    tableDefPart.Table.DisplayName = "CopiedTable" + tableId;
                    tableDefPart.Table.Name = "CopiedTable" + tableId;
                    tableDefPart.Table.Save();
                }
            }

            //There should only be one sheet that has focus
            var views = clonedSheet.Worksheet.GetFirstChild<SheetViews>();
            if (views != null)
            {
                views.Remove();
                clonedSheet.Worksheet.Save();
            }

            var sheets = this.workbookPart.Workbook.GetFirstChild<Sheets>();
            var copiedSheet = new Sheet { Name = copiedSheetName, Id = this.workbookPart.GetIdOfPart(clonedSheet), SheetId = (uint)sheets.ChildElements.Count + 1 };
            sheets.Append(copiedSheet);
            
            //Save Changes
            this.workbookPart.Workbook.Save();

            var result =  new ExcelWorksheet(this, copiedSheetName, clonedSheet);

            this.Sheets.Add(result.Name, result);

            return result;
        }

        public void Close()
        {
            this.workbookPart.Workbook.Save();
            this.Document.Close();
        }
    }
}