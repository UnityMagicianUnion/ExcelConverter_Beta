namespace MarsCode.ExcelConverter
{
    using UnityEditor;
    using UnityEngine;

    public class ExcelPreviewWindow : EditorWindow
    {

        public static void Launch(IExcelConverter input, WorkBookData[] data)
        {
            var window = GetWindow<ExcelPreviewWindow>();
            window.target = input;
            window.workbooks = data;
            window.book = data[0];
            window.sheet = data[0].GetWorkSheetData(0);

            var len = data.Length;
            window.workbookTabs = new string[len];
            for(int i = 0; i < len; i++)
            {
                window.workbookTabs[i] = window.workbooks[i].FileName;
            }

            window.SwitchSheet();
        }


        IExcelConverter target;

        WorkBookData[] workbooks;

        WorkBookData book;

        WorkSheetData sheet;

        int bookID, sheetID;

        string[] workbookTabs;

        bool onMaximizeWindow;

        GUIStyle style_title, style_index, style_data, style_popup;

        Texture2D[] bg = new Texture2D[2];

        Vector2 scrollPos, windowMaxSize, windowMinSize;


        void OnEnable()
        {
            style_title = new GUIStyle(EditorStyles.whiteBoldLabel);
            style_title.fontSize = 16;
            style_title.alignment = TextAnchor.MiddleCenter;
            style_title.normal.background = ExcelConverterUtility.GenerateColorBlockTexture(ExcelConverterUtility.ExcelStandardColor);

            style_index = new GUIStyle();
            style_index.fontSize = 14;
            style_index.alignment = TextAnchor.MiddleRight;
            style_index.normal.textColor = new Color(0, 0, 0, 0.5f);
            style_index.normal.background = ExcelConverterUtility.GenerateColorBlockTexture(new Color(1, 1, 1, 0.5f));

            style_data = new GUIStyle();
            style_data.fontSize = 14;
            style_data.alignment = TextAnchor.MiddleLeft;
            bg[0] = ExcelConverterUtility.GenerateColorBlockTexture(new Color(0.7f, 0.7f, 0.7f));
            bg[1] = ExcelConverterUtility.GenerateColorBlockTexture(new Color(0.75f, 0.75f, 0.75f));

            style_popup = new GUIStyle(GUI.skin.button);
            style_popup.focused.textColor = Color.black;
        }


        void OnGUI()
        {
            DrawHeaderUI();

            DrawWorkSheetUI();

            DrawToolButtonUI();
        }


        void DrawHeaderUI()
        {
            EditorGUILayout.LabelField("Excel Viewer - " + workbooks[bookID].FileName, style_title, GUILayout.Height(30));

            GUILayout.Space(20);
        }


        void DrawWorkSheetUI()
        {
            scrollPos = EditorGUILayout.BeginScrollView(scrollPos, false, false);
            {
                for(int row = 0; row < sheet.Data.GetLength(0); row++)
                {
                    style_data.normal.background = bg[row % 2];

                    EditorGUILayout.BeginHorizontal();
                    {
                        GUILayout.Label((row + 1).ToString() + ". ", style_index, GUILayout.Height(sheet.RowHeights[row]), GUILayout.Width(50));

                        GUILayout.Label("", style_data, GUILayout.Height(sheet.RowHeights[row]), GUILayout.Width(20));

                        for(int col = 0; col < sheet.Data.GetLength(1); col++)
                        {
                            EditorGUILayout.SelectableLabel(sheet.Data[row, col], style_data, GUILayout.Height(sheet.RowHeights[row]), GUILayout.Width(sheet.ColumnWidths[col]));
                        }
                    }
                    EditorGUILayout.EndHorizontal();
                }
            }
            EditorGUILayout.EndScrollView();
        }


        void DrawToolButtonUI()
        {
            GUILayout.Space(5);

            DrawWorkBookSwitchUI();

            EditorGUILayout.BeginHorizontal();
            {
                DrawWorksheetSwitchUI();

                DrawWindowMaximizeButton();

                DrawWindowCloseButton();
            }
            EditorGUILayout.EndVertical();

            GUILayout.Space(5);

            DrawConvertSheetButton();

            GUILayout.Space(5);

            DrawConvertAllButton();

            GUILayout.Space(10);
        }


        #region >>  Toollist Functions

        void DrawWorkBookSwitchUI()
        {
            var id = EditorGUILayout.Popup(bookID, workbookTabs, style_popup, GUILayout.Height(20));

            if(bookID != id)
            {
                sheetID = 0;

                bookID = id;

                book = workbooks[bookID];

                SwitchSheet();
            }
        }


        void DrawWorksheetSwitchUI()
        {
            var id = EditorGUILayout.Popup(sheetID, workbooks[bookID].WorkSheetTabs, style_popup, GUILayout.Height(20));

            if(sheetID != id)
            {
                sheetID = id;

                SwitchSheet();
            }
        }


        void SwitchSheet()
        {
            sheet = book.GetWorkSheetData(sheetID);

            var heights = sheet.RowHeights;
            var widths = sheet.ColumnWidths;

            var totalHeight = 180;
            var totalWidth = 80;

            for(int r = 0; r < heights.Length; r++)
                totalHeight += heights[r];

            for(int c = 0; c < widths.Length; c++)
                totalWidth += widths[c];

            var minWidth = totalWidth > 1000 ? 1000 : totalWidth;
            var minHeight = totalHeight > 800 ? 800 : totalHeight;
            windowMinSize = new Vector2(minWidth, minHeight);
            windowMaxSize = new Vector2(totalWidth, totalHeight);

            SwitchWindowSize();
        }


        void SwitchWindowSize()
        {
            if(onMaximizeWindow)
            {
                minSize = windowMaxSize;
                maxSize = windowMaxSize;
                minSize = windowMinSize;
            }
            else
            {
                minSize = windowMinSize;
                maxSize = windowMinSize;
                maxSize = windowMaxSize;
            }
        }


        void DrawWindowMaximizeButton()
        {
            var on = GUILayout.Toggle(onMaximizeWindow, "視窗最大化", "button", GUILayout.Height(20), GUILayout.Width(100));
            if(onMaximizeWindow != on)
            {
                onMaximizeWindow = on;
                SwitchWindowSize();
            }
        }


        void DrawWindowCloseButton()
        {
            if(GUILayout.Button(new GUIContent("關閉視窗"), GUILayout.Height(20), GUILayout.Width(100)))
            {
                Close();
            }
        }


        void DrawConvertSheetButton()
        {
            GUI.color = ExcelConverterUtility.StandardYellow;
            if(GUILayout.Button(new GUIContent("轉換檢視的工作頁"), GUILayout.Height(25)))
            {
                var sheetData = book.GetWorkSheetData(sheetID);
                target.ConvertSingleWorksheet(sheetData);

                Close();
            }
            GUI.color = Color.white;
        }


        void DrawConvertAllButton()
        {
            GUI.color = ExcelConverterUtility.StandardYellow;
            if(GUILayout.Button(new GUIContent("轉換所有的活頁簿"), GUILayout.Height(25)))
            {
                target.ConvertAllWorkbooks(workbooks);

                Close();
            }
            GUI.color = Color.white;
        }

        #endregion

    }
}