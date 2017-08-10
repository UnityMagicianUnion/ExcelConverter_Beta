# 當前版本
v1.1.0

<br/>

# 操作影片

- 核心功能

[![Excel Converter](https://img.youtube.com/vi/UgRWwWcC2zM/0.jpg)](https://www.youtube.com/watch?v=UgRWwWcC2zM)

<br/>

# 用法

### ExcelConverterEditor   

- 自訂型別

   - 建立一個繼承自 MonoBeheviour 或 ScriptableObj 的型別
   
   - 分別宣告兩個屬性 filePaths 和 autoRefresh

```csharp
/*********************************************
範例
*********************************************/

using UnityEngine;

public class Example1 : MonoBehaviour
{
    [SerializeField] CharacterJobs[] jobs;
    [SerializeField] Weapons[] weapons;
    [SerializeField] Equipments[] equipments;

    #if UNITY_EDITOR
    [SerializeField, HideInInspector] string[] filePaths;   // ** 宣告 filePaths 屬性 **
    [SerializeField, HideInInspector] bool autoRefresh;     // ** 宣告 autoRefresh 屬性 **
    #endif
}
```

<br/>

- 自訂編輯器型別

  - 建立一個繼承自 ExcelConverterEditor 的型別

  - 在 OnEnable 函式中呼叫 ExcelConverterInitialize

  - 在 OnInspectorGUI 函式中呼叫呼叫 base.OnInspectorGUI

  - 如果需要轉換單一工作頁的功能，覆寫 ConvertSingleWorksheet 函式 

  - 如果需要轉換所有活頁簿的功能，覆寫 ConvertAllWorkbooks 函式
```csharp
/*********************************************
範例
*********************************************/

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
        // ** 初始化 **
        ExcelConverterInitialize();
    }

    public override void OnInspectorGUI()
    {
        // ** ExcelConverter 功能 **
        base.OnInspectorGUI();
    }
    
    
    // ** 覆寫此函式可以自訂轉換單個工作頁資料的功能 **
    public override void ConvertSingleWorksheet(WorkSheetData worksheet)
    {
        /********************************
        自定格式轉換邏輯
        ********************************/
        
        serializedObject.ApplyModifiedProperties();
    }

    // ** 覆寫此函式可以自訂轉換全部活頁簿資料的功能 **
    public override void ConvertAllWorkbooks(WorkBookData[] workbooks)
    {
        /********************************
        自定格式轉換邏輯
        ********************************/

        serializedObject.ApplyModifiedProperties();
    }
}
```

<br/>

# API 參考

### WorkBookData 型別

- **public string FileName**  
*回傳活頁簿的檔案名稱(沒有副檔名)。*

- **public int WorkSheetCount**  
*回傳工作頁的數量。*
 
- **public string[] WorkSheetTabs**  
*回傳所有的工作頁名稱。*

- **public WorkSheetData[] WorksheetData**  
*回傳所有的工作頁資料。*

- **public WorkSheetData GetWorkSheetData(int id)**  
*依據代入的索引回傳指定的工作頁資料。*

- **public WorkSheetData GetWorkSheetData(string id)**  
*依據代入的索引回傳指定的工作頁資料。*

<br/>

### WorkSheetData 型別

- **public string TabName**  
*回傳工作頁的名稱。*

- **public string[,] Data**  
*回傳 row-base 數據。*

- **public int Rows**  
*回傳工作頁 row 的數量。*

- **public int Columns**  
*回傳工作頁 coulumn 的數量。*

- **public int[] RowHeights**  
*回傳工作頁所有 row cell 的高度。*

- **public int[] ColumnWidths**  
*回傳工作頁所有 column cell 的寬度。*

<br/>

# 版本注釋

### v1.1.0
- 新增自動更新的特性
- 不需透過預覽視窗可以直接呼叫格式轉換的功能
- 在工作頁列表下，如果工作頁的名稱與 GameObject 一致將會變色提醒
