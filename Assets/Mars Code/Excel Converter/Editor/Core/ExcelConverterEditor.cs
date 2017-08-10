namespace MarsCode.ExcelConverter
{
    using System;
    using UnityEditor;
    using UnityEditorInternal;
    using UnityEngine;

    public class ExcelConverterEditor : Editor, IExcelConverter
    {

        protected virtual string DefaultPath { get { return "Assets/"; } }

        protected SerializedProperty filePaths, autoRefresh;

        protected ReorderableList filePathList;

        bool onMainUI = true;

        int workbookID = 0;

        string[] workbookList;

        WorkBookData[] workbooks;

        Vector2 scrollPos;

        bool datachanged = true;


        protected void ExcelConverterInitialize(string filePathsFieldName = "filePaths", string autoRefreshFieldName = "autoRefresh")
        {
            filePaths = serializedObject.FindProperty(filePathsFieldName);
            autoRefresh = serializedObject.FindProperty(autoRefreshFieldName);

            if(filePaths == null)
                Debug.LogError("filePaths is null.");

            if(autoRefresh == null)
                Debug.LogError("autoRefresh is null.");

            filePathList = new ReorderableList(serializedObject, filePaths, true, false, true, true)
            {
                drawElementCallback = (rect, index, isActive, isFocused) =>
                {
                    var element = filePaths.GetArrayElementAtIndex(index);
                    rect.y += 2;
                    rect.height = EditorGUIUtility.singleLineHeight;

                    EditorGUI.PropertyField(rect, element, GUIContent.none);
                },

                onReorderCallback = (list) =>
                {
                    if(autoRefresh.boolValue)
                        Refresh();
                    else
                        datachanged = true;
                },

                onAddCallback = (list) =>
                {
                    var path = EditorUtility.OpenFilePanel("Excele Converter", DefaultPath, "xlsx");

                    if(path != "")
                    {
                        filePaths.arraySize = filePaths.arraySize + 1;
                        filePaths.GetArrayElementAtIndex(filePaths.arraySize - 1).stringValue = path;

                        if(autoRefresh.boolValue)
                            Refresh();
                        else
                            datachanged = true;
                    }
                },

                onRemoveCallback = (list) =>
                {
                    filePaths.DeleteArrayElementAtIndex(list.index);

                    if(autoRefresh.boolValue)
                        Refresh();
                    else
                        datachanged = true;
                }
            };

            if(autoRefresh.boolValue)
                Refresh();
        }


        #region >>  UI Drawing

        public override void OnInspectorGUI()
        {
            serializedObject.Update();

            GUILayout.Space(10);

            if(filePaths == null)
            {
                EditorGUILayout.HelpBox("property \"filePaths\" is null", MessageType.Info);
                return;
            }

            EditorGUILayout.BeginVertical(GUI.skin.box);
            {
                DrawHeaderUI();

                if(onMainUI)
                {
                    DrawCoreUI();
                }
                else
                {
                    DrawWorksheetListUI();
                }
            }
            EditorGUILayout.EndVertical();

            serializedObject.ApplyModifiedProperties();
        }


        void DrawHeaderUI()
        {
            var style_title = new GUIStyle(EditorStyles.whiteBoldLabel);
            style_title.normal.background = ExcelConverterUtility.GenerateColorBlockTexture(ExcelConverterUtility.ExcelStandardColor);

            GUILayout.Label(new GUIContent("Excel Converter"), style_title);
        }


        void DrawCoreUI()
        {
            filePathList.DoLayoutList();

            GUILayout.Space(10);

            if(filePaths.arraySize > 0)
            {
                DrawRefreshUI();

                GUILayout.Space(5);

                DrawMainButtons();
            }
            else
            {
                EditorGUILayout.HelpBox("點選 \"+\" 按鈕以新增檔案", MessageType.Info);
            }

            GUILayout.Space(5);
        }


        void DrawRefreshUI()
        {
            if(autoRefresh.boolValue)
            {
                autoRefresh.boolValue = GUILayout.Toggle(autoRefresh.boolValue, "自動更新", "button");
            }
            else
            {
                if(datachanged)
                {
                    GUI.color = ExcelConverterUtility.StandardRed;

                    if(GUILayout.Button(new GUIContent("立即更新!")))
                        Refresh();

                    GUI.color = Color.white;
                }
                else
                {
                    autoRefresh.boolValue = GUILayout.Toggle(autoRefresh.boolValue, "自動更新", "button");
                }
            }
        }


        void DrawMainButtons()
        {
            if(workbooks != null)
            {
                GUI.color = ExcelConverterUtility.StandardBlue;
                if(GUILayout.Button(new GUIContent("開啟預覽視窗")))
                {
                    ExcelPreviewWindow.Launch(this, workbooks);
                }

                GUILayout.Space(5);

                GUI.color = ExcelConverterUtility.StandardBlue;
                if(GUILayout.Button(new GUIContent("選擇工作表")))
                {
                    DisplayWorkSheetList();
                }

                GUILayout.Space(5);

                GUI.color = ExcelConverterUtility.StandardYellow;
                if(GUILayout.Button(new GUIContent("轉換所有活頁簿")))
                {
                    ConvertAllWorkbooks(workbooks);
                }

                GUI.color = Color.white;
            }
        }


        void DrawWorksheetListUI()
        {
            workbookID = EditorGUILayout.Popup(workbookID, workbookList);

            GUILayout.Space(5);

            var sheetLen = workbooks[workbookID].WorkSheetCount;

            if(sheetLen > 20)
            {
                scrollPos = EditorGUILayout.BeginScrollView(scrollPos, false, false, GUILayout.Height(420));
                {
                    DrawSheetButtons(sheetLen);
                }
                EditorGUILayout.EndScrollView();
            }
            else
            {
                DrawSheetButtons(sheetLen);
            }

            GUILayout.Space(5);

            GUI.color = ExcelConverterUtility.StandardBlue;

            if(GUILayout.Button(new GUIContent(" ◄ 返回")))
            {
                onMainUI = true;
            }

            GUI.color = Color.white;
        }


        void DrawSheetButtons(int sheetLen)
        {
            for(int i = 0; i < sheetLen; i++)
            {
                var sheetName = workbooks[workbookID].WorkSheetTabs[i];

                GUI.color = (sheetName == target.name) ? ExcelConverterUtility.StandardYellow : Color.white;

                if(GUILayout.Button(new GUIContent(sheetName)))
                {
                    ConvertSingleWorksheet(workbooks[workbookID].GetWorkSheetData(i));
                }

                GUI.color = Color.white;
            }
        }

        #endregion


        #region >>  Convert Data

        void DisplayWorkSheetList()
        {
            scrollPos = Vector2.zero;
            workbookID = 0;

            var len = workbooks.Length;
            workbookList = new string[len];

            for(int i = 0; i < len; i++)
            {
                workbookList[i] = workbooks[i].FileName;
            }

            onMainUI = false;
        }


        void Refresh()
        {
            var len = filePaths.arraySize;
            workbooks = new WorkBookData[len];

            for(int i = 0; i < len; i++)
            {
                var path = filePaths.GetArrayElementAtIndex(i).stringValue;
                workbooks[i] = new WorkBookData(path);
            }

            datachanged = false;
        }


        protected WorkBookData GetWorkBook(string id, WorkBookData[] data)
        {
            foreach(var d in data)
            {
                if(d.FileName == id)
                    return d;
            }

            throw new Exception("找不到活頁簿: " + id);
        }


        public virtual void ConvertSingleWorksheet(WorkSheetData worksheet) { }


        public virtual void ConvertAllWorkbooks(WorkBookData[] workbooks) { }

        #endregion

    }
}
