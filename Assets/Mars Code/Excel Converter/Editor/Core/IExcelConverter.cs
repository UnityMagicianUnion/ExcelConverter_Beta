namespace MarsCode.ExcelConverter
{
    public interface IExcelConverter
    {
        void ConvertSingleWorksheet(WorkSheetData worksheet);

        void ConvertAllWorkbooks(WorkBookData[] workbooks);
    }
}