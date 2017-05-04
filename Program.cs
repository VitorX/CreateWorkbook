using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace CreateWorkbook
{
    class Program
    {
        static void Main(string[] args)
        {
            createsheetFromTXT("test.xlsx", "output10.txt", "");
        }

        public static void createsheetFromTXT(string filepath, string txtfilePath, string ErrMsg)
        {
            SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Create(filepath, SpreadsheetDocumentType.Workbook);
            WorkbookPart workbookpart = spreadsheetDocument.AddWorkbookPart();
            workbookpart.Workbook = new Workbook();

            SharedStringTablePart shareStringPart = workbookpart.AddNewPart<SharedStringTablePart>();
            SharedStringTable sharedStringTable = shareStringPart.SharedStringTable = new SharedStringTable();

            int counter = 1;
            string line;
            System.IO.StreamReader file = new System.IO.StreamReader(txtfilePath);
            while ((line = file.ReadLine()) != null)
            {
                string[] words = line.Split(new string[] { "^&*" }, System.StringSplitOptions.RemoveEmptyEntries);

                if (words.Length > 1)
                {
                    ErrMsg = ValSheetName(words[1]);

                    if (ErrMsg == "" && words[0] == "SN")//
                    {
                        WorksheetPart worksheetPart = InsertWorksheetCus(workbookpart, words[1]);
                        if (words.Length >7)
                        {
                            Cell cell = InsertCellInWorksheetFun(words[3], Convert.ToUInt32(words[5]), worksheetPart);
                            int index = InsertSharedStringItemFun(words[7], shareStringPart);
                            cell.CellValue = new CellValue(index.ToString());
                            cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
                        }
                    }
                }
                counter++;
            }
            workbookpart.Workbook.Save();
            spreadsheetDocument.Close();
        }

        private static bool isValid(String value)
        {
            var matches = Regex.Matches(value, @"^[A-Za-z_-][A-Za-z0-9_-]*$");
            return (matches.Count > 0);
        }

        private static string ValSheetName(string worksheetName)
        {
            if (worksheetName.Trim() == "")
                return "Sheet name cannot be empty!";

            if (worksheetName.Trim().Length > 31)
                return "Worksheet name cannot have more than 31 characters.";

            if (!isValid(worksheetName))
                return "You should have only alphanumeric characters with either _ or - to the Worksheet name.";

            List<string> List0 = new List<string>(); List0.Add("Database"); List0.Add("Criteria"); List0.Add("Extract"); List0.Add("Print_Area"); List0.Add("Print_Titles");

            bool b = List0.Any(worksheetName.Contains);
            if (b)
                return "Worksheet name cannot be one value of these: Database, Criteria, Extract, Print_Area or Print_Titles.";

            return "";
        }

        private static WorksheetPart InsertWorksheetCus(WorkbookPart workbookPart, string worksheetName)
        {
            Sheets sheet = workbookPart.Workbook.GetFirstChild<Sheets>();
            if (sheet == null)
            {
                sheet = workbookPart.Workbook.AppendChild<Sheets>(new Sheets());
            }
            IEnumerable<Sheet> sheets = workbookPart.Workbook.Descendants<Sheet>().Where(s => s.Name == worksheetName);
            if (sheets.Count() == 0)
            {
                WorksheetPart newWorksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                newWorksheetPart.Worksheet = new Worksheet(new SheetData());

                Sheets newsheets = workbookPart.Workbook.GetFirstChild<Sheets>();
                string relationshipId = workbookPart.GetIdOfPart(newWorksheetPart);
                uint sheetId = 1;
                if (newsheets.Elements<Sheet>().Count() > 0)
                {
                    sheetId = newsheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1;
                }
                Sheet newsheet = new Sheet() { Id = relationshipId, SheetId = sheetId, Name = worksheetName };
                newsheets.Append(newsheet);
                return newWorksheetPart;
            }
            WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheets.First().Id);
            return worksheetPart;
        }
        private static Cell InsertCellInWorksheetFun(string columnName, uint rowIndex, WorksheetPart worksheetPart)
        {
            DocumentFormat.OpenXml.Spreadsheet.Worksheet worksheet = worksheetPart.Worksheet;
            SheetData sheetData = worksheet.GetFirstChild<SheetData>();
            string cellReference = columnName + rowIndex;
            Row row;
            if (sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).Count() != 0)
            {
                row = sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).First();
            }
            else
            {
                row = new Row() { RowIndex = rowIndex };
                sheetData.Append(row);
            }
            if (row.Elements<Cell>().Where(c => c.CellReference.Value == columnName + rowIndex).Count() > 0)
            {
                return row.Elements<Cell>().Where(c => c.CellReference.Value == cellReference).First();
            }
            else
            {
                Cell cell = new Cell() { CellReference = cellReference };
                row.Append(cell);
                worksheet.Save();
                return cell;
            }

        }
        static int InsertSharedStringItemFun(string text, SharedStringTablePart shareStringPart)
        {
            if (shareStringPart.SharedStringTable == null)
            {
                shareStringPart.SharedStringTable = new SharedStringTable();
            }
            int i = shareStringPart.SharedStringTable.Elements<SharedStringItem>().Count();
            if (i == 0)
            {
                shareStringPart.SharedStringTable.AppendChild(new SharedStringItem(new DocumentFormat.OpenXml.Spreadsheet.Text(text)));
                return i;
            }
            shareStringPart.SharedStringTable.AppendChild(new SharedStringItem(new DocumentFormat.OpenXml.Spreadsheet.Text(text)));
            shareStringPart.SharedStringTable.Save();
            return i;
        }
    }
}
