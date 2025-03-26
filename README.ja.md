# EasyToAccess VBAライブラリ

**EasyToAccess** は、Microsoft Excel から Microsoft Access を簡単に操作できるようにする軽量かつ直感的なVBAライブラリです。  
`Dictionary` を用いたシンプルな記述で、`SELECT`・`INSERT`・`UPDATE`・`DELETE` などのデータベース操作や、テーブル・フィールドの取得も可能です。

----------

## 🌍 対応言語

🇬🇧 [English version here](./README.md)

----------

## 📦 主な特徴

-   Excel VBA から Access（.accdb）ファイルへの接続が簡単
    
-   `Dictionary` を使ってデータを柔軟に操作
    
-   ExcelとAccessでよくある**型の不一致**を避けるための「型自動変換機能」あり（任意で有効化可能）
    
-   トランザクションの基本的な操作（開始・コミット・ロールバック）に対応
    
-   SQL実行結果をExcelシートに出力可能
    
-   実務向けに無駄を削ぎ落としたシンプルな構成
    

----------

## 📁 含まれるファイル

-   `EasyToAccess.cls`：Access操作のためのクラスモジュール
    
-   `ETAUtil.bas`：上記クラスを簡単に呼び出すためのラッパーモジュール
    

----------

## 🔧 導入手順

1.  **VBAプロジェクトへファイルをインポート**  
    Excelで `Alt + F11` を押してVBAエディタを開き、  
    メニューから  
    `ファイル > ファイルのインポート...` を選択し、以下の2つを読み込んでください：
    
    -   `EasyToAccess.cls`
        
    -   `ETAUtil.bas`
        

----------

## 🚀 はじめに

### 共通関数（ConnectDB / DisconnectDB）

すべての操作の前に、まずは Access ファイルへ接続し、最後に切断します。

接続時の第1引数には、自分で好きな**接続名（dbName）**を設定します。  
これは複数のAccessファイルを同時に扱う際に識別するためのものです。
```
ETAUtil.ConnectDB "myDb", "C:\Path\To\Your\Database.accdb"
```

-   `"myDb"`：任意の接続名（識別用）
    
-   `"C:\..."`：Accessファイルのフルパス
    
-   第3引数（省略可）：パスワード（必要な場合のみ）
    

```
ETAUtil.DisconnectDB "myDb"
``` 

-  切断時にも、接続名を使います：


## 🧩 操作サンプル（INSERT / UPDATE）

### Insertの例

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
### Updateの例
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

## 🔒 ライセンス

このプロジェクトは MITライセンス のもとで公開されています。  
（商用利用、改変、再配布が可能ですが、著作権表示とライセンス文の同梱が必要です）

----------

## 📫 お問い合わせ

バグ報告・機能追加のご提案などは GitHub の [Issue](https://github.com/%E3%81%82%E3%81%AA%E3%81%9F%E3%81%AE%E3%83%A6%E3%83%BC%E3%82%B6%E3%83%BC%E5%90%8D/EasyToAccess/issues) よりお気軽にご連絡ください。  
コントリビューションも歓迎します！