namespace MarsCode.ExcelConverter
{
    public class WorkSheetData
    {

        public WorkSheetData(string index, string[,] data, int[] rowHeightsws, int[] columnWidthsmns)
        {
            m_tabName = index;

            m_data = data;

            m_RowHeights = rowHeightsws;

            m_ColumnWidths = columnWidthsmns;
        }


        string m_tabName;

        string[,] m_data;

        int[] m_RowHeights;

        int[] m_ColumnWidths;


        /// <summary>
        /// 回傳工作頁名稱.
        /// </summary>
        public string TabName { get { return m_tabName; } }


        /// <summary>
        /// 回傳 row-base 資料.
        /// </summary>
        public string[,] Data { get { return m_data; } }


        /// <summary>
        /// 回傳 row 的數量.
        /// </summary>
        public int Rows { get { return m_RowHeights.Length; } }


        /// <summary>
        /// 回傳 coulumn 的數量.
        /// </summary>
        public int Columns { get { return m_ColumnWidths.Length; } }


        /// <summary>
        /// 回傳工作頁的 row cell 高度.
        /// </summary>
        public int[] RowHeights { get { return m_RowHeights; } }


        /// <summary>
        /// 回傳工作頁的 column cell 寬度.
        /// </summary>
        public int[] ColumnWidths { get { return m_ColumnWidths; } }

    }
}