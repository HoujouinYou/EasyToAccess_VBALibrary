VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EasyToAccess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private accessFilePath As String
Private DBPassWord As String
Private CON As Object

Private Const adOpenKeyset = 1

Public Enum ETAInitResult
    InitSuccess = 0
    FileNotFound = -1
    ConnectionError = -2
End Enum

Public Enum SelectResultFormat
    withFieldName = 1
    withoutFieldName = 2
End Enum

Public Enum SQLErrorList
    SqlError = 1
    DbError = 2
End Enum

Public Enum TableObjectType
    TableOnly = 0
    TableAndQuery = 1
End Enum

Public Property Get getAccessFilePath() As String
    getAccessFilePath = accessFilePath
End Property

Public Property Get getCON() As Object
    Set getCON = CON
End Property

Private Sub Class_Initialize()
    Set CON = CreateObject("ADODB.Connection")
End Sub

Public Function init(ByVal fPath As String, ByVal dbPass As String) As ETAInitResult
    Dim ret As ETAInitResult
    
    accessFilePath = fPath
    DBPassWord = dbPass
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(accessFilePath) Then
        ret = FileNotFound
        GoTo ext
    End If
    
    ret = openDBConnection
    
ext:
    init = ret
End Function

Private Function openDBConnection() As ETAInitResult
    On Error Resume Next
    If DBPassWord = "" Then
        CON.Open "Provider=Microsoft.ACE.OLEDB.12.0;" & _
                 "Data Source=" + accessFilePath + ";"
    Else
        CON.Open "Provider=Microsoft.ACE.OLEDB.12.0;" & _
                 "Data Source=" + accessFilePath + ";" & _
                 "Jet OLEDB:Database Password=" + DBPassWord + ";"
    End If
    
    If err.Number = 0 Then
        openDBConnection = InitSuccess
    Else
        openDBConnection = ConnectionError
    End If
    On Error GoTo 0
End Function

Public Sub closeDBConnection()
    If Not CON Is Nothing Then
        If CON.State = 1 Then CON.Close
        Set CON = Nothing
    End If
End Sub

Private Function chkConnection() As Boolean
    If CON Is Nothing Then Exit Function
    chkConnection = (CON.State = 1)
End Function

Private Sub closeRecordSet(ByRef rs As Object)
    If TypeName(rs) = "Recordset" Then
        If rs.State = 1 Then rs.Close
    End If
    Set rs = Nothing
End Sub

Public Function executeSelect(ByVal sql As String, ByVal includeFieldNames As SelectResultFormat) As Variant
    Dim retVal As Variant, errNum As SQLErrorList, errDesc As String, recordCnt As Long
    
    If Not chkConnection Then
        errNum = DbError
        errDesc = "Database connection is not open."
        GoTo ext
    End If
    
    Dim rs As Object
    Set rs = CreateObject("ADODB.RecordSet")
    
    On Error Resume Next
    rs.Open sql, CON, adOpenKeyset
    If err.Number <> 0 Then
        On Error GoTo 0
        errNum = SqlError
        errDesc = err.Description
        GoTo ext
    End If
    On Error GoTo 0
    
    recordCnt = rs.RecordCount
    
    If rs.EOF Then
        GoTo skipLoop
    End If
    
    Dim r As Long, c As Long
    If includeFieldNames = withFieldName Then
        ReDim retVal(rs.RecordCount, rs.fields.Count - 1)
        For c = 0 To rs.fields.Count - 1
            retVal(0, c) = rs.fields(c).Name
        Next c
        r = 1
    ElseIf includeFieldNames = withoutFieldName Then
        ReDim retVal(rs.RecordCount - 1, rs.fields.Count - 1)
        r = 0
    End If

skipLoop:
    Do Until rs.EOF
        For c = 0 To rs.fields.Count - 1
            retVal(r, c) = rs.fields(c)
        Next c
        r = r + 1
        rs.MoveNext
    Loop
    
    closeRecordSet rs
    
ext:
    Dim ret(3)
    If recordCnt > 0 Then
        ret(0) = retVal
    Else
        ret(0) = Empty
    End If
    ret(1) = errNum
    ret(2) = errDesc
    ret(3) = recordCnt
    executeSelect = ret
End Function

Public Function getTables(ByVal tblType As TableObjectType) As Variant
    Dim retVal() As String, errNum As SQLErrorList, errDesc As String
    If Not chkConnection Then
        errNum = DbError
        errDesc = "Database connection is not open."
        GoTo ext
        Exit Function
    End If
    
    Dim cat As Object
    Set cat = CreateObject("ADOX.Catalog")
    cat.ActiveConnection = CON
    
    Dim tbl As Object, cnt As Integer
    cnt = 0
    For Each tbl In cat.tables
        Select Case tblType
            Case TableOnly:
                If tbl.Type = "TABLE" Then
                    ReDim Preserve retVal(cnt)
                    retVal(cnt) = tbl.Name + " - TABLE"
                    cnt = cnt + 1
                End If
            Case TableAndQuery:
                If tbl.Type = "TABLE" Or tbl.Type = "VIEW" Then
                    ReDim Preserve retVal(cnt)
                    retVal(cnt) = tbl.Name + " - " + tbl.Type
                    cnt = cnt + 1
                End If
        End Select
    Next
    
    Set cat = Nothing

ext:
    Dim ret(3)
    If cnt > 0 Then
        ret(0) = retVal
    Else
        ret(0) = Empty
    End If
    ret(1) = errNum
    ret(2) = errDesc
    ret(3) = cnt
    getTables = ret
End Function

Public Function GetFields(ByVal tblName As String) As Variant
    Dim retVal() As String, errNum As SQLErrorList, errDesc As String, recordCnt As Long
    If Not chkConnection Then
        errNum = DbError
        errDesc = "Database connection is not open."
        GoTo ext
        Exit Function
    End If
    
    Dim sql As String
    sql = "SELECT * FROM " + tblName
    
    Dim rs As Object
    Set rs = CreateObject("ADODB.RecordSet")
    
    On Error Resume Next
    rs.Open sql, CON, adOpenKeyset
    If err.Number <> 0 Then
        On Error GoTo 0
        errNum = SqlError
        errDesc = err.Description
        GoTo ext
    End If
    On Error GoTo 0
    
    Dim cnt As Integer
    cnt = 0
    
    For cnt = 0 To rs.fields.Count - 1
        ReDim Preserve retVal(cnt)
        retVal(cnt) = rs.fields(cnt).Name
    Next cnt
    
ext:
    Dim ret(3)
    If cnt > 0 Then
        ret(0) = retVal
    Else
        ret(0) = Empty
    End If
    ret(1) = errNum
    ret(2) = errDesc
    ret(3) = cnt
    GetFields = ret
End Function

Public Function ExecSql(ByVal sql As String) As Variant
    Dim errNum As SQLErrorList, errDesc As String
    If Not chkConnection Then
        errNum = DbError
        errDesc = "Database connection is not open."
        GoTo errorHandler
    End If
    
    CON.BeginTrans
    
    Dim ret(3) As Variant
    Dim affectedRows As Long
    On Error Resume Next
    CON.Execute sql, affectedRows
    If err.Number <> 0 Then
        errNum = SqlError
        errDesc = err.Description
        On Error GoTo 0
        GoTo errorHandler
    End If
    On Error GoTo 0
    CON.CommitTrans
    ret(0) = True
    ret(3) = affectedRows
    ExecSql = ret
    Exit Function
    
errorHandler:
    On Error Resume Next
    CON.RollbackTrans
    ret(0) = False
    ret(1) = errNum
    ret(2) = errDesc
    ret(3) = 0
    ExecSql = ret
    On Error GoTo 0
End Function

Public Function getFieldType(ByVal tblName As String, ByVal fieldName As String) As Variant
    If Not chkConnection Then
        getFieldType = CVErr(xlErrNA)
        Exit Function
    End If

    Dim rs As Object
    Set rs = CreateObject("ADODB.RecordSet")

    On Error Resume Next
    rs.Open "SELECT * FROM [" & tblName & "] WHERE 1=0", CON, adOpenKeyset
    If err.Number <> 0 Then
        getFieldType = CVErr(xlErrValue)
        Exit Function
    End If
    On Error GoTo 0

    Dim fldType As Variant
    fldType = rs.fields(fieldName).Type

    rs.Close
    Set rs = Nothing

    getFieldType = fldType
End Function

Public Sub BeginTransaction()
    CON.BeginTrans
End Sub

Public Sub CommitTransaction()
    CON.CommitTrans
End Sub

Public Sub RollbackTransaction()
    CON.RollbackTrans
End Sub
