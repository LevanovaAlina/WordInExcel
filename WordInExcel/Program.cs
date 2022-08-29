namespace WordInExcel
{
    class Program
    {
        static void Main(string[] args)
        {
            const string fileWordName = @"D:\Games\doc.docx";
            const string fileExcelName = @"D:\Games\table.xlsx";
            TextWordToExcel.ReadTextWordFile(fileWordName);
            TextWordToExcel.CreateExcelFile();
            TextWordToExcel.RecordTextWordToExcelFile(fileExcelName);
        }
    }
}