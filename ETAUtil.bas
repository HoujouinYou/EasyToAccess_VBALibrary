Attribute VB_Name = "ETAUtil"
Option Explicit

Private ETADIC As Object

Private Const LOG_PREFIX As String = "Connection state check failed at: "
Private Const LOG_SPLITTER As String = "------------------------------"

Public Enum PrintBorderType
    PrintBorderType_None = 0
    PrintBorderType_Solid = 1
    PrintBorderType_Dash = 2
    PrintBorderType_VerticalSolidAndHorizontalDashed = 3
    PrintBorderType_VerticalDashedAndHorizontalSolid = 4
End Enum

Public Enum ShowDirection
    Direction_Vertical = 0
    Direction_Horizontal = 1
End Enum

Private Type CellsRange
    rTop As Long
    rLeft As Long
    rBottom As Long
    rRight As Long
End Type

Public Sub ConnectDB(ByVal dbName As String, ByVal accessFilePath As String, Optional ByVal dbPass As String = "")
    If ETADIC Is Nothing Then
        Set ETADIC = CreateObject("Scripting.Dictionary")
    End If
    
    Dim ETA As New EasyToAccess
    If ETADIC.exists(dbName) Then
        Debug.Print "exists err"
        Exit Sub
    End If
    
    Dim result As ETAInitResult
    result = ETA.init(accessFilePath, dbPass)
    
    Dim dbgStr As String
    Select Case result
        Case InitSuccess:
            dbgStr = "InitSuccess"
            ETADIC.Add dbName, ETA
        Case FileNotFound:
            dbgStr = "FileNotFound"
            Set ETA = Nothing
        Case ConnectionError:
            dbgStr = "ConnectionError"
            Set ETA = Nothing
    End Select
    
    Debug.Print "DBConnection - " + dbgStr
End Sub

Private Function getETAInstance(ByVal dbName As String) As EasyToAccess
    If Not ETADIC Is Nothing Then
        If ETADIC.exists(dbName) Then
            Set getETAInstance = ETADIC.Item(dbName)
        Else
            Set getETAInstance = Nothing
        End If
    End If
End Function

Public Sub DisconnectDB(ByVal dbName As String)
    Dim ETA As EasyToAccess
    Set ETA = getETAInstance(dbName)
    If ETA Is Nothing Then
        Debug.Print LOG_PREFIX + "disconnectDB"
        Exit Sub
    End If
    
    If isConnectedDB(ETA) Then
        ETA.closeDBConnection
        Debug.Print "DBConnection - Disconnected"
    Else
        Debug.Print LOG_PREFIX + "disconnectDB"
    End If
    ETADIC.Remove dbName
End Sub

Private Function isConnectedDB(ByRef ETA As EasyToAccess) As Boolean
    If ETA Is Nothing Then Exit Function
    If ETA.getCON Is Nothing Then Exit Function
    isConnectedDB = (ETA.getCON.State = 1)
End Function

Public Function ExecSelect(ByVal dbName As String, ByVal sql As String, Optional ByVal includeFieldNames As SelectResultFormat = withoutFieldName) As Variant
    If UCase(left(sql, 6)) <> "SELECT" Then
        Debug.Print "Invalid query: Only SELECT statements are allowed"
        Exit Function
    End If
    
    Dim ETA As EasyToAccess
    Set ETA = getETAInstance(dbName)
    If ETA Is Nothing Then
        Debug.Print LOG_PREFIX + "disconnectDB"
        Exit Function
    End If

    If Not isConnectedDB(ETA) Then
        Debug.Print LOG_PREFIX + "execSelect"
        Exit Function
    End If
    
    Dim res As Variant
    res = ETA.executeSelect(sql, includeFieldNames)
    If res(1) <> 0 Then
        Select Case res(1)
            Case SqlError:
                Debug.Print "SQL Error: " & res(2)
            Case DbError:
                Debug.Print "DataBase Error: " & res(2)
            Case Else:
                Debug.Print "Other Error: " & res(2)
        End Select
    Else
        Debug.Print "Record Count: " & res(3)
    End If
    ExecSelect = res(0)
End Function

Public Function PrintD(ByVal arr2D, ByRef pRng As Range, Optional ByVal borderType As PrintBorderType = PrintBorderType_None)
    If Not IsArray(arr2D) Then
        Debug.Print "printD error: input is not a valid 2D array"
        Exit Function
    End If
    
    Dim cr As CellsRange
    cr = getCellsRange(pRng, arr2D)
    
    With Workbooks(pRng.Parent.Parent.Name).Sheets(pRng.Parent.Name)
        With .Range(.Cells(cr.rTop, cr.rLeft), .Cells(cr.rBottom, cr.rRight))
            .Value = arr2D
            
            Select Case borderType
                Case PrintBorderType_Solid:
                    .Borders.LineStyle = xlContinuous
                Case PrintBorderType_Dash:
                    .Borders.LineStyle = xlContinuous
                    .Borders(xlInsideVertical).LineStyle = xlDash
                    .Borders(xlInsideHorizontal).LineStyle = xlDash
                Case PrintBorderType_VerticalSolidAndHorizontalDashed:
                    .Borders.LineStyle = xlContinuous
                    .Borders(xlInsideVertical).LineStyle = xlContinuous
                    .Borders(xlInsideHorizontal).LineStyle = xlDash
                Case PrintBorderType_VerticalDashedAndHorizontalSolid:
                    .Borders.LineStyle = xlContinuous
                    .Borders(xlInsideVertical).LineStyle = xlDash
                    .Borders(xlInsideHorizontal).LineStyle = xlContinuous
            End Select
        End With
    End With
End Function

Private Function getCellsRange(ByRef pRng As Range, ByRef dat As Variant) As CellsRange
    Dim ret As CellsRange
    
    Dim sCellRow As Long, sCellColumn As Long
    ret.rTop = pRng.Cells(1, 1).Row
    ret.rLeft = pRng.Cells(1, 1).Column
    
    
    Dim rowCount As Long, colCount As Long
    rowCount = UBound(dat, 1) - LBound(dat, 1) + 1
    colCount = UBound(dat, 2) - LBound(dat, 2) + 1
    
    ret.rBottom = ret.rTop + rowCount - 1
    ret.rRight = ret.rLeft + colCount - 1
    
    getCellsRange = ret
End Function

Public Function ShowFields(ByVal accessFilePath As String, ByVal tblName As String, Optional ByVal dbPass As String = "")
    Dim tmpNm As String
    tmpNm = "tmp"
    
    ConnectDB tmpNm, accessFilePath, dbPass
    
    Dim ETA As EasyToAccess
    Set ETA = getETAInstance(tmpNm)
    If ETA Is Nothing Then
        Debug.Print LOG_PREFIX + "showFields"
        Exit Function
    End If

    If Not isConnectedDB(ETA) Then
        Debug.Print LOG_PREFIX + "showFields"
        Exit Function
    End If
    
    Dim res As Variant
    res = ETA.GetFields(tblName)
    
    If res(1) <> 0 Then
        Select Case res(1)
            Case DbError:
                Debug.Print "DataBase Error: " & res(2)
            Case Else:
                Debug.Print "Other Error: " & res(2)
        End Select
    Else
        Dim dat() As String, d
        dat = res(0)
        Debug.Print LOG_SPLITTER
        Debug.Print "AccessFilePath：" + accessFilePath
        Debug.Print "TableName：" + tblName
        Debug.Print "FieldList："
        For Each d In dat
            Debug.Print d
        Next
        Debug.Print LOG_SPLITTER
        
        Debug.Print "Field Count: " & res(3)
        Debug.Print LOG_SPLITTER
    End If
    
    DisconnectDB tmpNm
End Function

Public Function ShowTables(ByVal accessFilePath As String, Optional dbPass As String = "", Optional ByVal tableType As TableObjectType = TableOnly)
    Dim tmpNm As String
    tmpNm = "tmp"
    
    ConnectDB tmpNm, accessFilePath, dbPass
    
    Dim ETA As EasyToAccess
    Set ETA = getETAInstance(tmpNm)
    If ETA Is Nothing Then
        Debug.Print LOG_PREFIX + "showTables"
        Exit Function
    End If
    
    If Not isConnectedDB(ETA) Then
        Debug.Print LOG_PREFIX + "showTables"
        Exit Function
    End If
    
    Dim res As Variant
    res = ETA.getTables(tableType)
    
    If res(1) <> 0 Then
        Select Case res(1)
            Case DbError:
                Debug.Print "DataBase Error: " & res(2)
            Case Else:
                Debug.Print "Other Error: " & res(2)
        End Select
    Else
        Dim dat() As String, d
        dat = res(0)
        Debug.Print LOG_SPLITTER
        Debug.Print "AccessFilePath：" + accessFilePath
        Debug.Print "TableList："
        For Each d In dat
            Debug.Print d
        Next
        Debug.Print LOG_SPLITTER
        Debug.Print "Table Count: " & res(3)
        Debug.Print LOG_SPLITTER
    End If
    
    DisconnectDB tmpNm
End Function

Public Function GetFields(ByVal dbName As String, ByVal tblName As String) As String()
    Dim ETA As EasyToAccess
    Set ETA = getETAInstance(dbName)
    If ETA Is Nothing Then
        Debug.Print LOG_PREFIX + "getTables"
        Exit Function
    End If
    
    If Not isConnectedDB(ETA) Then
        Debug.Print LOG_PREFIX + "getFields"
        Exit Function
    End If
    
    Dim res As Variant
    res = ETA.GetFields(tblName)
    
    If res(1) <> 0 Then
        Select Case res(1)
            Case DbError:
                Debug.Print "DataBase Error: " & res(2)
            Case Else:
                Debug.Print "Other Error: " & res(2)
        End Select
    Else
        Dim dat() As String, d
        dat = res(0)
        Debug.Print "getFields() Result: " & res(3)
    End If
    GetFields = dat
End Function

Public Function ExecSql(ByVal dbName As String, ByVal sql As String) As Boolean
    Dim ETA As EasyToAccess
    Set ETA = getETAInstance(dbName)
    If ETA Is Nothing Then
        Debug.Print "execSql error: DB instance not found"
        Exit Function
    End If

    If UCase(left(Trim(sql), 6)) = "SELECT" Then
        Debug.Print "execSql error: SELECT is not allowed in this method"
        Exit Function
    End If
    
    If Not isConnectedDB(ETA) Then
        Debug.Print LOG_PREFIX + "execSql"
        Exit Function
    End If
    
    Dim ret As Boolean
    Dim res As Variant
    res = ETA.ExecSql(sql)
    
    If Not res(0) Then
        Select Case res(1)
            Case SqlError:
                Debug.Print "SQL Error: " & res(2)
            Case DbError:
                Debug.Print "DataBase Error: " & res(2)
            Case Else:
                Debug.Print "Other Error: " & res(2)
        End Select
    Else
        Debug.Print "execSql() Result: " & res(3)
        ret = True
    End If
    ExecSql = ret
End Function

Private Function areAllTypesValid(ByRef ETA As EasyToAccess, ByVal tableName As String, ByVal dataDic As Object, Optional ByRef rowIndex As Variant = Empty) As Boolean
    Dim k As Variant
    For Each k In dataDic.keys
        Dim expectedType As Variant
        expectedType = ETA.getFieldType(tableName, k)

        If IsError(expectedType) Then
            Debug.Print "[Type Check Error] Field '" & k & "' not found in table '" & tableName & "'"
            Exit Function
        End If

        Dim val As Variant
        val = dataDic(k)

        If Not isTypeCompatible(val, expectedType) Then
            Dim idxStr As String
            If Not IsEmpty(rowIndex) Then idxStr = "[Row: " & rowIndex & "] "
            Debug.Print idxStr & "[Field: " & k & "] Type mismatch: Expected=" & expectedType & ", Actual=" & TypeName(val)
            Exit Function
        End If
    Next

    areAllTypesValid = True
End Function

Private Function isTypeCompatible(ByVal val As Variant, ByVal adoType As Long) As Boolean
    Select Case adoType
        Case 3, 20
            isTypeCompatible = IsNumeric(val)
        Case 7
            isTypeCompatible = IsDate(val)
        Case 202, 203
            isTypeCompatible = VarType(val) = vbString
        Case 11
            isTypeCompatible = (VarType(val) = vbBoolean Or val = 0 Or val = 1)
        Case Else
            Debug.Print "isTypeCompatible: Unhandled ADO type " & adoType & ". Assuming compatible."
            isTypeCompatible = True
    End Select
End Function

Private Function sqlFormatValue(ByVal val As Variant) As String
    If IsNull(val) Then
        sqlFormatValue = "NULL"
    ElseIf IsDate(val) Then
        sqlFormatValue = "#" & Format(val, "yyyy/mm/dd") & "#"
    ElseIf VarType(val) = vbString Then
        sqlFormatValue = "'" & Replace(val, "'", "''") & "'"
    ElseIf VarType(val) = vbBoolean Then
        sqlFormatValue = IIf(val, "True", "False")
    ElseIf IsNumeric(val) Then
        sqlFormatValue = val
    Else
        sqlFormatValue = "'" & Replace(CStr(val), "'", "''") & "'"
    End If
End Function

Private Function areAllKeysValid(ByVal dbName As String, ByVal tableName As String, ByVal dataDic As Object) As Boolean
    Dim allFields() As String
    allFields = GetFields(dbName, tableName)
    If Not IsArray(allFields) Then Exit Function
    
    Dim fldDic As Object
    Set fldDic = CreateObject("Scripting.Dictionary")
    
    Dim f As Variant
    For Each f In allFields
        fldDic.Add f, True
    Next
    
    Dim k As Variant
    For Each k In dataDic.keys
        If Not fldDic.exists(CStr(k)) Then
            Debug.Print "areAllKeysValid Error: [" & k & "] is not a field in table [" & tableName & "]"
            Exit Function
        End If
    Next
    
    areAllKeysValid = True
End Function

Public Function ExecInsert(ByVal dbName As String, ByVal tableName As String, ByVal dataDic As Object, Optional ByVal autoConvert As Boolean = True, Optional ByVal rowIndex As Variant = Empty) As Boolean
    Dim ETA As EasyToAccess
    Set ETA = getETAInstance(dbName)
    If ETA Is Nothing Then
        Debug.Print "execInsert Error: DB instance not found"
        Exit Function
    End If

    If Not isConnectedDB(ETA) Then
        Debug.Print "execInsert Error: DB is not connected"
        Exit Function
    End If

    If Not areAllKeysValid(dbName, tableName, dataDic) Then Exit Function
    If Not autoConvert Then
        If Not areAllTypesValid(ETA, tableName, dataDic, rowIndex) Then Exit Function
    End If

    Dim sql As String
    Dim colList As String, valList As String
    Dim key As Variant

    For Each key In dataDic.keys
        Dim val As Variant
        val = dataDic(key)

        If autoConvert Then
            Dim expectedType As Variant
            expectedType = ETA.getFieldType(tableName, key)
            If IsError(expectedType) Then
                Debug.Print "[AutoConvert Error] Field '" & key & "' not found in table '" & tableName & "'"
                Exit Function
            End If
        
            val = convertToExpectedType(val, expectedType)
            If IsError(val) Then
                Debug.Print "[AutoConvert Error] Conversion failed for field '" & key & "'"
                Exit Function
            End If
        
            If Not isTypeCompatible(val, expectedType) Then
                Debug.Print "[AutoConvert Error] Type still mismatched after conversion at field '" & key & "' → Value: " & val
                Exit Function
            End If
        End If

        colList = colList & "[" & key & "], "
        valList = valList & sqlFormatValue(val) & ", "
    Next key

    If right(colList, 2) = ", " Then colList = left(colList, Len(colList) - 2)
    If right(valList, 2) = ", " Then valList = left(valList, Len(valList) - 2)

    sql = "INSERT INTO [" & tableName & "] (" & colList & ") VALUES (" & valList & ")"
    Dim result As Variant
    result = ETA.ExecSql(sql)

    If Not result(0) Then
        Debug.Print "execInsert error: SQL execution failed"
        Debug.Print "SQL: " & sql
        Select Case result(1)
            Case SqlError: Debug.Print "SQL Error: " & result(2)
            Case DbError: Debug.Print "DB Error: " & result(2)
            Case Else: Debug.Print "Unknown Error: " & result(2)
        End Select
        ExecInsert = False
        Exit Function
    End If

    ExecInsert = True
End Function

Public Function ExecUpdate(ByVal dbName As String, ByVal tableName As String, ByVal dataDic As Object, ByVal whereDic As Object, Optional ByVal autoConvert As Boolean = True, Optional ByVal rowIndex As Variant = Empty) As Boolean
    Dim ETA As EasyToAccess
    Set ETA = getETAInstance(dbName)
    If ETA Is Nothing Then
        Debug.Print "execUpdate Error: DB instance not found"
        Exit Function
    End If

    If Not isConnectedDB(ETA) Then
        Debug.Print "execUpdate Error: DB is not connected"
        Exit Function
    End If

    If Not areAllKeysValid(dbName, tableName, dataDic) Then Exit Function
    If Not areAllKeysValid(dbName, tableName, whereDic) Then Exit Function

    Dim expectedType As Variant
    Dim val As Variant
    Dim key As Variant
    
    For Each key In dataDic.keys
        val = dataDic(key)
        If autoConvert Then
            expectedType = ETA.getFieldType(tableName, key)
            If IsError(expectedType) Then
                Debug.Print "[AutoConvert Error] Field '" & key & "' not found in table '" & tableName & "'"
                Exit Function
            End If
            val = convertToExpectedType(val, expectedType)
            If IsError(val) Then
                Debug.Print "[AutoConvert Error] Conversion failed for field '" & key & "'"
                Exit Function
            End If
            If Not isTypeCompatible(val, expectedType) Then
                Debug.Print "[AutoConvert Error] Type still mismatched after conversion at field '" & key & "' → Value: " & val
                Exit Function
            End If
            dataDic(key) = val
        Else
            If Not areAllTypesValid(ETA, tableName, dataDic, rowIndex) Then Exit Function
        End If
    Next key

    For Each key In whereDic.keys
        val = whereDic(key)
        If autoConvert Then
            expectedType = ETA.getFieldType(tableName, key)
            If IsError(expectedType) Then
                Debug.Print "[AutoConvert Error] Field '" & key & "' not found in table '" & tableName & "'"
                Exit Function
            End If
            val = convertToExpectedType(val, expectedType)
            If IsError(val) Then
                Debug.Print "[AutoConvert Error] Conversion failed for WHERE field '" & key & "'"
                Exit Function
            End If
            If Not isTypeCompatible(val, expectedType) Then
                Debug.Print "[AutoConvert Error] Type still mismatched in WHERE after conversion at field '" & key & "' → Value: " & val
                Exit Function
            End If
            whereDic(key) = val
        Else
            If Not areAllTypesValid(ETA, tableName, whereDic, rowIndex) Then Exit Function
        End If
    Next key

    Dim setClause As String, whereClause As String

    For Each key In dataDic.keys
        setClause = setClause & "[" & key & "] = " & sqlFormatValue(dataDic(key)) & ", "
    Next key
    If right(setClause, 2) = ", " Then setClause = left(setClause, Len(setClause) - 2)

    For Each key In whereDic.keys
        whereClause = whereClause & "[" & key & "] = " & sqlFormatValue(whereDic(key)) & " AND "
    Next key
    If right(whereClause, 5) = " AND " Then whereClause = left(whereClause, Len(whereClause) - 5)

    Dim sql As String
    sql = "UPDATE [" & tableName & "] SET " & setClause & " WHERE " & whereClause

    Dim result As Variant
    result = ETA.ExecSql(sql)

    If Not result(0) Then
        Debug.Print "execUpdate error: SQL execution failed"
        Debug.Print "SQL: " & sql
        Select Case result(1)
            Case SqlError: Debug.Print "SQL Error: " & result(2)
            Case DbError: Debug.Print "DB Error: " & result(2)
            Case Else: Debug.Print "Unknown Error: " & result(2)
        End Select
        ExecUpdate = False
        Exit Function
    End If

    ExecUpdate = True
End Function

Public Function ExecDelete(ByVal dbName As String, ByVal tableName As String, ByVal whereDic As Object, Optional ByVal autoConvert As Boolean = True) As Boolean
    Dim ETA As EasyToAccess
    Set ETA = getETAInstance(dbName)
    If ETA Is Nothing Then
        Debug.Print "execDelete Error: DB instance not found"
        Exit Function
    End If

    If Not isConnectedDB(ETA) Then
        Debug.Print "execDelete Error: DB is not connected"
        Exit Function
    End If

    If whereDic Is Nothing Or whereDic.Count = 0 Then
        Debug.Print "execDelete Error: execDelete Error: WHERE clause is missing (preventing full deletion)"
        Exit Function
    End If

    If Not areAllKeysValid(dbName, tableName, whereDic) Then Exit Function
    
    Dim key As Variant
    Dim expectedType As Variant
    Dim val As Variant

    For Each key In whereDic.keys
        val = whereDic(key)
        If autoConvert Then
            expectedType = ETA.getFieldType(tableName, key)
            If IsError(expectedType) Then
                Debug.Print "[AutoConvert Error] Field '" & key & "' not found in table '" & tableName & "'"
                Exit Function
            End If
            val = convertToExpectedType(val, expectedType)
            If IsError(val) Then
                Debug.Print "[AutoConvert Error] Conversion failed for WHERE field '" & key & "'"
                Exit Function
            End If
            If Not isTypeCompatible(val, expectedType) Then
                Debug.Print "[AutoConvert Error] Type mismatch in WHERE clause at field '" & key & "' after conversion → Value: " & val
                Exit Function
            End If
            whereDic(key) = val
        Else
            If Not areAllTypesValid(ETA, tableName, whereDic) Then Exit Function
        End If
    Next

    Dim whereClause As String
    For Each key In whereDic.keys
        whereClause = whereClause & "[" & key & "] = " & sqlFormatValue(whereDic(key)) & " AND "
    Next

    If right(whereClause, 5) = " AND " Then
        whereClause = left(whereClause, Len(whereClause) - 5)
    End If

    Dim sql As String
    sql = "DELETE FROM [" & tableName & "] WHERE " & whereClause

    Dim result As Variant
    result = ETA.ExecSql(sql)

    If Not result(0) Then
        Debug.Print "execDelete error: SQL execution failed"
        Debug.Print "SQL: " & sql
        Select Case result(1)
            Case SqlError: Debug.Print "SQL Error: " & result(2)
            Case DbError: Debug.Print "DB Error: " & result(2)
            Case Else: Debug.Print "Unknown Error: " & result(2)
        End Select
        ExecDelete = False
        Exit Function
    End If

    ExecDelete = True
End Function

Public Function Compact_and_Repair(ByVal accessFilePath As String, Optional ByVal pass As String) As Boolean
    If ETADIC Is Nothing Then GoTo main
    
    Dim keys
    keys = ETADIC.keys
    
    If UBound(keys) >= 0 Then
        Dim k As Variant
        Dim p As String
        For Each k In keys
            p = ETADIC.Item(k).getAccessFilePath
            If accessFilePath = p Then
                Debug.Print "The database is currently open. Please close the database before Compact_and_Repair."
                Exit Function
            End If
        Next
    End If
    
main:
    Dim dbe As Object
    Set dbe = CreateObject("DAO.DBEngine.120")
    
    Dim fileName As String, sp, dirName As String
    Dim tempName As String
    tempName = "_tempAccdb.accdb"
    
    sp = Split(accessFilePath, "\")
    fileName = sp(UBound(sp))
    dirName = Split(accessFilePath, fileName)(0)
    
    On Error GoTo compacting_err
    dbe.CompactDatabase accessFilePath, dirName + tempName
    
    Kill accessFilePath
    
    Name dirName + tempName As accessFilePath
    
    Compact_and_Repair = True
    Exit Function
    
compacting_err:
    Debug.Print "Cannot Complete Compact_and_Repair()"
End Function

Public Sub BeginTransaction(ByVal dbName As String)
    Dim ETA As EasyToAccess
    Set ETA = getETAInstance(dbName)
    If Not ETA Is Nothing Then
        If ETA.getCON.State = 1 Then
            ETA.BeginTransaction
        End If
    End If
End Sub

Public Sub CommitTransaction(ByVal dbName As String)
    Dim ETA As EasyToAccess
    Set ETA = getETAInstance(dbName)
    If Not ETA Is Nothing Then
        If ETA.getCON.State = 1 Then
            ETA.CommitTransaction
            Debug.Print "Transaction has been committed"
        End If
    End If
End Sub

Public Sub RollbackTransaction(ByVal dbName As String)
    Dim ETA As EasyToAccess
    Set ETA = getETAInstance(dbName)
    If Not ETA Is Nothing Then
        If ETA.getCON.State = 1 Then
            ETA.RollbackTransaction
            Debug.Print "Transaction has been rolled back"
        End If
    End If
End Sub

Private Function convertToExpectedType(ByVal val As Variant, ByVal adoType As Long) As Variant
    On Error GoTo conversion_failed
    Select Case adoType
        Case 3, 20 ' Integer, Long
            convertToExpectedType = CLng(val)
        Case 7 ' Date
            convertToExpectedType = CDate(val)
        Case 202, 203 ' Text, Memo
            convertToExpectedType = CStr(val)
        Case 11 ' Boolean
            convertToExpectedType = CBool(val)
        Case Else
            convertToExpectedType = val
    End Select
    Exit Function

conversion_failed:
    convertToExpectedType = CVErr(xlErrValue)
End Function
