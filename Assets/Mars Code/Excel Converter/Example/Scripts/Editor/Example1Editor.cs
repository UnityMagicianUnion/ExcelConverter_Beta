using MarsCode.ExcelConverter;          // ** 引用 MarsCode.ExcelConverter **
using UnityEditor;
using UnityEngine;

[CustomEditor(typeof(Example1)), CanEditMultipleObjects]
public class Example1Editor : ExcelConverterEditor      // ** 繼承 ExcelConverterEditor **
{

    // ** 覆寫預設的檔案路徑 **
    protected override string DefaultPath { get { return "Assets/Mars Code/Excel Converter/Example/Data"; } }


    void OnEnable()
    {
        // ** 必須執行初始化 **
        ExcelConverterInitialize();

        // filePahs and autoRefresh 的屬性名稱是可以變更的.
        // ex-> ExcelConverterInitialize( "myFilePathsPropertyName" , "myAutoRefreshPropertyName" );
    }


    public override void OnInspectorGUI()
    {
        // 自訂功能起始
        DrawDefaultInspector();

        serializedObject.Update();

        ClearData();

        serializedObject.ApplyModifiedProperties();
        // 自訂功能結尾


        // ** ExcelConverterEditor 功能 **
        base.OnInspectorGUI();
    }


    // ** 在這個範例我們只需要覆寫 ConvertAllWorkbooks **
    public override void ConvertAllWorkbooks(WorkBookData[] workbooks)
    {
        // 取得 "角色" 活頁簿
        var book1 = GetWorkBook("角色", workbooks);
        ConvertCharacterData(book1);

        // 取得 "物品" 活頁簿
        var book2 = GetWorkBook("物品", workbooks);

        // 取得 "武器" 工作頁
        var sheet1 = book2.GetWorkSheetData("武器");
        ConvertWeaponData(sheet1);

        // 取得 "裝備" 工作頁
        var sheet2 = book2.GetWorkSheetData("裝備");
        ConvertEquipmentData(sheet2);

        serializedObject.ApplyModifiedProperties();
    }


    void ClearData()
    {
        if(GUILayout.Button("清除檔案"))
        {
            var property = serializedObject.FindProperty("jobs");
            property.ClearArray();

            property = serializedObject.FindProperty("weapons");
            property.ClearArray();

            property = serializedObject.FindProperty("equipments");
            property.ClearArray();
        }
    }


    void ConvertCharacterData(WorkBookData data)
    {
        // 取得工作頁的數量
        var len1 = data.WorkSheetCount;

        // 指派 jobs 屬性
        var property = serializedObject.FindProperty("jobs");
        property.ClearArray();
        property.arraySize = len1;

        for(int i = 0; i < len1; i++)
        {
            // 取得工作頁內容
            var sheet = data.GetWorkSheetData(i);

            // 指派 m_index 參數
            property.GetArrayElementAtIndex(i).FindPropertyRelative("m_index").stringValue = sheet.TabName;

            // 取得工作頁的 row 數量
            var len2 = sheet.Rows;

            // 指派 m_level 元素寬度，排除掉工作頁的第一行
            var levels = property.GetArrayElementAtIndex(i).FindPropertyRelative("m_level");
            levels.ClearArray();
            levels.arraySize = len2 - 1;

            // 在此迴圈內賦值
            for(int j = 1; j < len2; j++)
            {
                levels.GetArrayElementAtIndex(j - 1).FindPropertyRelative("m_level").intValue = int.Parse(sheet.Data[j, 0]);
                levels.GetArrayElementAtIndex(j - 1).FindPropertyRelative("m_attack").intValue = int.Parse(sheet.Data[j, 1]);
                levels.GetArrayElementAtIndex(j - 1).FindPropertyRelative("m_defense").intValue = int.Parse(sheet.Data[j, 2]);
                levels.GetArrayElementAtIndex(j - 1).FindPropertyRelative("m_luck").intValue = int.Parse(sheet.Data[j, 3]);
            }
        }
    }


    void ConvertWeaponData(WorkSheetData data)
    {
        var len = data.Rows;

        // 指派 weapons 元素寬度，排除掉工作頁的第一行
        var property = serializedObject.FindProperty("weapons");
        property.ClearArray();
        property.arraySize = len - 1;

        // 在此迴圈內賦值
        for(int i = 1; i < len; i++)
        {
            property.GetArrayElementAtIndex(i - 1).FindPropertyRelative("m_index").intValue = int.Parse(data.Data[i, 0]);
            property.GetArrayElementAtIndex(i - 1).FindPropertyRelative("m_name").stringValue = data.Data[i, 1];
            property.GetArrayElementAtIndex(i - 1).FindPropertyRelative("m_attack").intValue = int.Parse(data.Data[i, 2]);
        }
    }


    void ConvertEquipmentData(WorkSheetData data)
    {
        var len = data.Rows;

        // 指派 equipments 元素寬度，排除掉工作頁的第一行
        var property = serializedObject.FindProperty("equipments");
        property.ClearArray();
        property.arraySize = len - 1;

        // 在此迴圈內賦值
        for(int i = 1; i < len; i++)
        {
            property.GetArrayElementAtIndex(i - 1).FindPropertyRelative("m_index").intValue = int.Parse(data.Data[i, 0]);
            property.GetArrayElementAtIndex(i - 1).FindPropertyRelative("m_name").stringValue = data.Data[i, 1];
            property.GetArrayElementAtIndex(i - 1).FindPropertyRelative("m_defense").intValue = int.Parse(data.Data[i, 2]);
        }
    }

}
