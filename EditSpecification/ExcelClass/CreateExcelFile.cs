using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace EditSpecification.ExcelClass
{
    public class CreateExcelFile
    {
        public void GetExcelFile()
        {
            //Создается файл с расширением .xlsx
            string filePath = @"C:\Example\Add.xlsx";
            SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook);
            WorkbookPart workbookPart = spreadsheetDocument.AddWorkbookPart();
            workbookPart.Workbook = new Workbook();

            WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());

            Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());
            //В созданый файл добавляется новая книга
            //Sheet sheet = new Sheet()
            //{
            //    Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart),
            //    SheetId = 1,
            //    Name = "mySheet"
            //};
            //sheets.Append(sheet);

            workbookPart.Workbook.Save();

            spreadsheetDocument.Close();
        }
    }
}
