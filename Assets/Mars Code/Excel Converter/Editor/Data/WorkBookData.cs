namespace MarsCode.ExcelConverter
{
    using System;
    using System.Globalization;
    using System.IO;
    using NPOI.SS.UserModel;
    using NPOI.XSSF.UserModel;
    using UnityEngine;

    public class WorkBookData
    {

        public WorkBookData(string path)
        {
            XSSFWorkbook workbook;

            try
            {
                using(FileStream stream = File.Open(path, FileMode.Open, FileAccess.Read))
                {
                    workbook = new XSSFWorkbook(stream);
                }
            }
            catch(Exception)
            {
                var msg = string.Format("檔案不存在或開啟中.\n{0}", path);
                Debug.LogError(msg);
                return;
            }

            m_fileName = Path.GetFileNameWithoutExtension(path);

            var size = workbook.NumberOfSheets;
            m_WorkSheetTabs = new string[size];
            m_WorksheetData = new WorkSheetData[size];

            var df = new DataFormatter(CultureInfo.CurrentCulture);
            ISheet worksheet;
            for(int i = 0; i < size; i++)
            {
                m_WorkSheetTabs[i] = workbook.GetSheetName(i);
                worksheet = workbook.GetSheetAt(i);
                ConvertWorkSheetData(worksheet, ref m_WorksheetData[i], df);
            }
        }


        void ConvertWorkSheetData(ISheet sheet, ref WorkSheetData target, DataFormatter df)
        {
            var rows = sheet.PhysicalNumberOfRows;
            var columns = sheet.GetRow(0).PhysicalNumberOfCells;

            var tab = sheet.SheetName;
            var data = new string[rows, columns];
            var rowHeights = new int[rows];
            var columnWidths = new int[columns];

            IRow rowData;
            for(int r = 0; r < rows; r++)
            {
                rowData = sheet.GetRow(r);

                rowHeights[r] = Mathf.CeilToInt(rowData.HeightInPoints + 10);

                for(int c = 0; c < columns; c++)
                {
                    data[r, c] = df.FormatCellValue(rowData.GetCell(c));

                    if(r == 0)
                    {
                        columnWidths[c] = Mathf.CeilToInt(sheet.GetColumnWidthInPixels(c) + 10);
                    }
                }
            }

            target = new WorkSheetData(tab, data, rowHeights, columnWidths);
        }


        string m_fileName;

        string[] m_WorkSheetTabs;

        WorkSheetData[] m_WorksheetData;


        /// <summary>
        /// 回傳活頁簿檔案名稱(無副檔名).
        /// </summary>
        public string FileName { get { return m_fileName; } }


        /// <summary>
        /// 回傳工作頁的數量. 
        /// </summary>
        public int WorkSheetCount { get { return m_WorkSheetTabs.Length; } }


        /// <summary>
        /// 回傳所有工作頁的名稱.
        /// </summary>
        public string[] WorkSheetTabs { get { return m_WorkSheetTabs; } }


        /// <summary>
        /// 依據索引回傳工作頁資料.
        /// </summary>
        public WorkSheetData GetWorkSheetData(int id)
        {
            return m_WorksheetData[id];
        }


        /// <summary>
        /// 依據索引回傳工作頁資料.
        /// </summary>
        public WorkSheetData GetWorkSheetData(string id)
        {
            foreach(var sheet in m_WorksheetData)
            {
                if(sheet.TabName == id)
                    return sheet;
            }

            throw new Exception("cannot find the worksheet named: " + id);
        }


        /// <summary>
        /// 回傳所有的工作頁資料.
        /// </summary>
        public WorkSheetData[] WorksheetData { get { return m_WorksheetData; } }

    }

}