' =====================================================================
'   Table Constant
' =====================================================================

'   Load a Table into a 2D Array
'
Public Function Misc_TableConstLoad(ByVal Table As String) As String()

    Dim Rows() As String
    Rows = Split(Table, vbLf)
    Dim Cols() As String
    Cols = Split(Rows(0), "|")
    Dim TableArray() As String
    ReDim TableArray(0 To UBound(Rows), 0 To UBound(Cols))
    
    Dim RowsIndex As Long
    For RowsIndex = 0 To UBound(Rows)
        Cols = Split(Rows(RowsIndex), "|")
        Dim ColsIndex As Long
        For ColsIndex = 0 To UBound(Cols)
            TableArray(RowsIndex, ColsIndex) = Trim(Cols(ColsIndex))
        Next ColsIndex
    Next RowsIndex
    
    Misc_TableConstLoad = TableArray()

End Function

'   Get a Table data Value from the first matching RowKey and ColIndex
'
'   False   <-  If RowKey not found
'
Public Function Misc_TableConstFind(ByVal Table As String, ByVal RowKey As String, ByVal ColIndex As Long, ByRef Value As String) As Boolean
Misc_TableConstFind = False

    Dim TableArray() As String
    TableArray = Misc_TableConstLoad(Table)
    
    Dim RowIndex As Long
    For RowIndex = 1 To UBound(TableArray, 1)
        If StrComp(TableArray(RowIndex, 0), RowKey, vbTextCompare) = 0 Then
            Value = TableArray(RowIndex, ColIndex)
            Misc_TableConstFind = True
            Exit Function
        End If
    Next RowIndex
    Value = ""

End Function

'   Does RowKey exist in Table data rows?
'
'   False   <-  If RowKey not found
'
Public Function Misc_TableConstExist(ByVal Table As String, ByVal RowKey As String) As Boolean

    Dim Value As String
    Misc_TableConstExist = Misc_TableConstFind(Table, RowKey, 0, Value)

End Function

'   Get a Table data Value from the first matching RowKey and Header ColKey
'
'   False   <-  If RowKey not found
'
Public Function Misc_TableConstFindCol(ByVal Table As String, ByVal RowKey As String, ByVal ColKey As String, ByRef Value As String) As Boolean

    Dim Rows() As String
    Rows = Split(Table, vbLf)
    Dim Cols() As String
    Cols = Split(Rows(0), "|")
    
    Dim ColIndex As Long
    For ColIndex = 0 To UBound(Cols)
    
        If StrComp(Trim(Cols(ColIndex)), ColKey, vbTextCompare) = 0 Then
            Misc_TableConstFindCol = Misc_TableConstFind(Table, RowKey, ColIndex, Value)
            Exit Function
        End If
    
    Next ColIndex

    Debug.Print "ColKey not found.": Stop: Exit Function

End Function

' =====================================================================
'	Test Cases
' =====================================================================

Private Const CustForm_Card                     As String = "IPM.Post.Card"
Private Const CustForm_WipProject               As String = "IPM.Post.ProjectV3"
Private Const CustForm_WipActivity              As String = "IPM.Post.ActivityV3"
Private Const CustForm_WipProjectV4             As String = "IPM.Post.ProjectV4"

Private Const CustFormType_Card                 As String = "Card"
Private Const CustFormType_WIP                  As String = "WIP"

Private Const TEST_Table As String = _
 _
  " Class                           |    Type                       |    Name                      " & vbLf & _
 _
    CustForm_Card & "               |" & CustFormType_Card & "      |    Card                      " & vbLf & _
    CustForm_WipProject & "         |" & CustFormType_WIP & "       |    WIP Project               " & vbLf & _
    CustForm_WipActivity & "        |" & CustFormType_WIP & "       |    WIP Activity              " & vbLf & _
    CustForm_WipProjectV4 & "       |" & CustFormType_WIP & "       |    WIP Project V4            " & vbLf & _
 "" & _
    "RowA                           |    RowA Type                  |    RowA Name                 " & vbLf & _
    "RowB                           |    RowB Type                  |    RowB Name                 " & vbLf & _
    "RowC                           |    RowC Type                  |    RowC Name                 "
    '
    Private Const TEST_TableColClass    As Integer = 0
    Private Const TEST_TableColType     As Integer = 1
    Private Const TEST_TableColName     As Integer = 2
'
Public Sub TEST_TableConst()

    If Misc_TableConstExist(TEST_Table, "RowA") Then Debug.Print "RowA Exist"
    If Not Misc_TableConstExist(TEST_Table, "RowX") Then Debug.Print "RowX Does Not Exist"
    If Misc_TableConstExist(TEST_Table, CustForm_WipActivity) Then Debug.Print CustForm_WipActivity & " found."
    
    Dim MyValue As String
    If Misc_TableConstFind(TEST_Table, "RowA", TEST_TableColName, MyValue) Then Debug.Print "(RowA, 2) = " & MyValue
    If Misc_TableConstFindCol(TEST_Table, "RowA", "Name", MyValue) Then Debug.Print "(RowA, Name) = " & MyValue
    
    If Not Misc_TableConstFind(TEST_Table, "RowX", TEST_TableColName, MyValue) Then Debug.Print "RowX not found"
    If Misc_TableConstFind(TEST_Table, CustForm_WipProject, TEST_TableColName, MyValue) Then Debug.Print "(" & CustForm_WipProject & ", Name) = " & MyValue
    
    Dim MyArray() As String
    MyArray = Misc_TableConstLoad(TEST_Table)
    Dim RowIndex As Long
    For RowIndex = 0 To UBound(MyArray, 1)
        Dim ColIndex As Long
        For ColIndex = 0 To UBound(MyArray, 2)
            Debug.Print "(" & RowIndex & ", " & ColIndex & ") = " & MyArray(RowIndex, ColIndex)
        Next ColIndex
    Next RowIndex

End Sub

