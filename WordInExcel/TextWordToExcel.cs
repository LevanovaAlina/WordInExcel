using Excel = Microsoft.Office.Interop.Excel;
namespace WordInExcel
{
    static public class TextWordToExcel
    {
        private static dynamic textInDoc;
        private static dynamic workBook;
        private static dynamic excelApp;
        public static void ReadTextWordFile(string fileName)
        {
            var wordType = Type.GetTypeFromProgID("Word.Application");
            dynamic wordapp = Activator.CreateInstance(wordType);
            var worddoc = wordapp.Documents.Add(fileName);
            wordapp.Application.Documents.Open(fileName);
            textInDoc = worddoc.Content.Text;
        }

        public static void CreateExcelFile()
        {
            var excelType = Type.GetTypeFromProgID("Excel.Application");
            excelApp = Activator.CreateInstance(excelType);
            excelApp.SheetsInNewWorkbook = 1;
            workBook = excelApp.Workbooks.Add();
        }

        public static void RecordTextWordToExcelFile(string fileName)
        {
            Excel.Worksheet workSheet = workBook.ActiveSheet;
            workSheet.Cells[2, "B"] = textInDoc;
            workSheet.Cells[2, "B"].Font.Bold = true;
            workSheet.Columns.AutoFit();
            workBook.Close(true, fileName);
            excelApp.Quit();
            Console.WriteLine("Файл успешно сохранён!");
        }
    }
}
