# EasyToAccess VBA Library

**EasyToAccess** is a lightweight and user-friendly VBA library that simplifies interactions between Microsoft Excel and Microsoft Access.  
It allows you to easily perform database operations such as SELECT, INSERT, UPDATE, DELETE, and even table/field exploration — all within Excel using Dictionary-based syntax.

---

## 🌍 Available Languages

- 🇯🇵 [日本語版はこちら](./README.ja.md)

---

## 📦 Features

- Easily connect from Excel to Access (.accdb)
- Intuitive data operations using `Dictionary` objects
- Optional type conversion feature to avoid common type mismatch issues (e.g., numeric-looking Excel cells vs. Access text fields)
- Basic transaction support (Begin, Commit, Rollback)
- Output results directly to Excel ranges
- Minimal and practical code, ideal for business use

---

## 📁 Files Included

- `EasyToAccess.cls` – Class module for database interactions
- `ETAUtil.bas` – Wrapper module for simplified function calls

---

## 🔧 Setup

1. **Import the files into your VBA project**  
   Open the VBA editor (`Alt + F11`), then go to  
   `File > Import File...` and import the following:
   - `EasyToAccess.cls`
   - `ETAUtil.bas`
---
## 🚀 Getting Started
### Common Functions
Before using any database operations, you must first **connect to the Access file** using `ETAUtil.ConnectDB`, and finally **disconnect** with `ETAUtil.DisconnectDB`.

Each function uses a **user-defined name** (`dbName`) as the first argument.  
This allows you to manage multiple connections by assigning different names to each.
```
ETAUtil.ConnectDB "myDb", "C:\Path\To\Your\Database.accdb"
```
-   `"myDb"` is an arbitrary name for the database connection.
    
-   The second argument is the full path to your `.accdb` file.
    
-   Optional third argument: password (if your Access file is password-protected).

```
ETAUtil.DisconnectDB "myDb"
```
Use the same `dbName` ("myDb" in this example) to close the connection.

### Insert Examples
```
Sub Example_Insert()
    Dim dbName As String, dbPath As String
    dbName = "myDb"
    dbPath = "C:\Path\To\Your\Database.accdb"
    
    With ThisWorkbook.Sheets(1)
        Dim lastRow As Long
        lastRow = .Cells(.Rows.Count, 1).End(xlUp).Row
        
        Dim fieldCount As Long
        fieldCount = .Cells(1, Columns.Count).End(xlToLeft).Column
        
        Dim i As Long, c As Long
        Dim errorOccurred As Boolean
        
        ETAUtil.ConnectDB dbName, dbPath
        
        ETAUtil.BeginTransaction dbName
        
        For i = 2 To lastRow
            Dim dic As New Dictionary
            
            For c = 1 To fieldCount
                dic.Add .Cells(1, c).Value, .Cells(i, c).Value
            Next c
            
            If Not ETAUtil.ExecInsert(dbName, "YourTableName", dic, True) Then
                errorOccurred = True
                Exit For
            End If
            
            Set dic = Nothing
        Next i
    End With
    
    If errorOccurred Then
        ETAUtil.RollbackTransaction dbName
    Else
        ETAUtil.CommitTransaction dbName
    End If
    
    ETAUtil.DisconnectDB dbName
End Sub
```
### Update Examples
```
Sub Example_Update()
    Dim dbName As String, dbPath As String
    dbName = "myDb"
    dbPath = "C:\Path\To\Your\Database.accdb"
    
    Dim dataDic As New Dictionary
    Dim whereDic As New Dictionary
    
    dataDic.Add "Name", "Yamada Taro"
    dataDic.Add "Email", "taro@example.com"
    
    whereDic.Add "UserID", "00001"
    
    ETAUtil.ConnectDB dbName, dbPath
    ETAUtil.BeginTransaction dbName
    
    If ETAUtil.ExecUpdate(dbName, "YourTableName", dataDic, whereDic, True) Then
        ETAUtil.CommitTransaction dbName
    Else
        ETAUtil.RollbackTransaction dbName
    End If
    
    ETAUtil.DisconnectDB dbName
End Sub
```

## 🔒 License

This project is licensed under the MIT License.

---

## 📫 Contact

Feel free to open an issue for any questions or suggestions. Contributions are welcome!