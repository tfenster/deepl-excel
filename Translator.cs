using DeepL;
using DeepL.Model;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace DeepLExcel;

public class ExcelTranslator
{
    private Translator translator;
    private TextTranslateOptions textTranslateOptions = new TextTranslateOptions { Formality = Formality.More };
    private string path;
    private WorkbookPart? workbookPart;

    public ExcelTranslator(string authKey, string path)
    {
        this.translator = new Translator(authKey);
        this.path = path;
    }

    public async Task<GlossaryInfo> CreateGlossaryAsync(string name, string sourceLanguageCode, string targetLanguageCode, GlossaryEntries entries)
    {
        return await translator.CreateGlossaryAsync(name, sourceLanguageCode, targetLanguageCode, entries);
    }

    public async Task TranslateFile(bool skipHeader, string targetColumn, string sourceLanguage, string targetLanguage, string glossaryId)
    {
        using (var spreadsheetDocument = SpreadsheetDocument.Open(path, true))
        {
            this.workbookPart = spreadsheetDocument.WorkbookPart;
            if (workbookPart == null || workbookPart.Workbook == null || workbookPart.Workbook.Sheets == null || workbookPart.SharedStringTablePart == null)
                return;
            //var outerCount = 0;
            foreach (Sheet sheet in workbookPart.Workbook.Sheets.ChildElements)
            {
                /*outerCount++;
                if (outerCount > 5)
                    break;*/
                if (sheet.Id == null)
                    continue;
                var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id!);
                var worksheet = worksheetPart.Worksheet;
                var sheetData = (SheetData?)worksheet.ChildElements.Where(s => s is SheetData).FirstOrDefault();
                if (sheetData == null)
                {
                    Console.WriteLine($"WARN: Can't read worksheet data!");
                }
                else
                {
                    //var count = 0;
                    var isFirst = true;
                    foreach (Row currentRow in sheetData.ChildElements)
                    {
                        if (isFirst)
                        {
                            isFirst = false;
                            if (skipHeader)
                                continue;
                        }
                        /*count++;
                        if (count > 5)
                            break;*/
                        if (currentRow.ChildElements.Count == 0)
                            break;
                        var currentcell = (Cell)currentRow.ChildElements.GetItem(0);
                        var currentcellvalue = GetCellValue(currentcell);
                        var currenttranslatedcellvalue = string.Empty;
                        if (currentRow.ChildElements.Count > 2)
                        {
                            var currenttranslatedcell = (Cell)currentRow.ChildElements.GetItem(2);
                            currenttranslatedcellvalue = GetCellValue(currenttranslatedcell);
                        }

                        if (!(string.IsNullOrWhiteSpace(currentcellvalue) || currenttranslatedcellvalue.ToLower().StartsWith("is already on another page")))
                        {
                            var textTranslateOptions = new TextTranslateOptions();
                            if (targetLanguage == "de") textTranslateOptions.Formality = Formality.More;
                            if (glossaryId != null) textTranslateOptions.GlossaryId = glossaryId;
                            var translated = await translator.TranslateTextAsync(
                                currentcellvalue,
                                sourceLanguage,
                                targetLanguage,
                                textTranslateOptions
                            );

                            int index = InsertSharedStringItem(translated.Text);
                            Cell cell = InsertCellInRow(currentRow, targetColumn, sheetData);
                            worksheet.Save();

                            cell.CellValue = new CellValue(index.ToString());
                            cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
                            worksheet.Save();
                            Console.WriteLine(".");
                        }
                    }
                }
            }
        }
    }

    private string GetCellValue(Cell currentcell)
    {
        string currentcellvalue = string.Empty;
        if (currentcell.DataType != null)
        {
            if (currentcell.DataType == CellValues.SharedString)
            {
                int id = -1;

                if (Int32.TryParse(currentcell.InnerText, out id))
                {
                    SharedStringItem item = GetSharedStringItemById(id);

                    if (item.Text != null)
                    {
                        currentcellvalue = item.Text.Text;
                    }
                    else if (item.InnerText != null)
                    {
                        currentcellvalue = item.InnerText;
                    }
                    else if (item.InnerXml != null)
                    {
                        currentcellvalue = item.InnerXml;
                    }
                }
            }
        }
        return currentcellvalue;
    }

    private SharedStringItem GetSharedStringItemById(int id)
    {
        return workbookPart!.SharedStringTablePart!.SharedStringTable.Elements<SharedStringItem>().ElementAt(id);
    }

    private int InsertSharedStringItem(string text)
    {
        // If the part does not contain a SharedStringTable, create one.
        if (workbookPart!.SharedStringTablePart!.SharedStringTable == null)
        {
            workbookPart.SharedStringTablePart.SharedStringTable = new SharedStringTable();
        }

        int i = 0;

        // Iterate through all the items in the SharedStringTable. If the text already exists, return its index.
        foreach (SharedStringItem item in workbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>())
        {
            if (item.InnerText == text)
            {
                return i;
            }

            i++;
        }

        // The text does not exist in the part. Create the SharedStringItem and return its index.
        workbookPart.SharedStringTablePart.SharedStringTable.AppendChild(new SharedStringItem(new DocumentFormat.OpenXml.Spreadsheet.Text(text)));
        workbookPart.SharedStringTablePart.SharedStringTable.Save();

        return i;
    }

    private Cell InsertCellInRow(Row row, string columnName, SheetData sheetData)
    {
        var cellReference = columnName + row.RowIndex;
        // If there is not a cell with the specified column name, insert one.  
        if (row.Elements<Cell>().Where(c => c.CellReference?.Value == cellReference).Count() > 0)
        {
            return row.Elements<Cell>().Where(c => c.CellReference?.Value == cellReference).First();
        }
        else
        {
            // Cells must be in sequential order according to CellReference. Determine where to insert the new cell.
            Cell? refCell = null;
            foreach (Cell cell in row.Elements<Cell>())
            {
                if (cell.CellReference?.Value?.Length == cellReference.Length)
                {
                    if (string.Compare(cell.CellReference.Value, cellReference, true) > 0)
                    {
                        refCell = cell;
                        break;
                    }
                }
            }

            Cell newCell = new Cell() { CellReference = cellReference };
            row.InsertBefore(newCell, refCell);
            return newCell;
        }
    }

    public GlossaryEntries ReadGlossaryFromExcel(bool skipHeader = false)
    {
        var entries = new Dictionary<string, string>();
        using (var spreadsheetDocument = SpreadsheetDocument.Open(path, false))
        {
            this.workbookPart = spreadsheetDocument.WorkbookPart;
            if (workbookPart == null || workbookPart.Workbook == null || workbookPart.Workbook.Sheets == null || workbookPart.SharedStringTablePart == null)
                return new GlossaryEntries(entries);
            WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
            SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();
            foreach (Row r in sheetData.Elements<Row>().Skip((skipHeader ? 1 : 0)))
            {
                (int currentPositionleft, int currentPositionTop) = Console.GetCursorPosition();
                Console.Write(String.Format("Processing row {0} / {1}", r.RowIndex, sheetData.Elements<Row>().Count()));
                Console.SetCursorPosition(currentPositionleft, currentPositionTop);
                var cell1 = (Cell)r.ElementAt(0);
                var cell2 = (Cell)r.ElementAt(1);
                if (r.Elements().Count() == 3)
                {
                    cell1 = (Cell)r.ElementAt(1);
                    cell2 = (Cell)r.ElementAt(2);
                }
                entries.Add(GetCellValue(cell1), GetCellValue(cell2));
            }
            Console.WriteLine();
        }
        return new GlossaryEntries(entries);
    }
}