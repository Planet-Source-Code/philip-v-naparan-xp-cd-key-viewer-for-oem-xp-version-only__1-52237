Attribute VB_Name = "modCustomes"
Option Explicit
Global rs_customer_f As New ADODB.Recordset
Public Sub open_rs_cus_f(ByRef sDGrid As DataGrid, ByVal sTable As String)
Call DB1(DBPath)
rs_customer_f.CursorLocation = adUseClient
rs_customer_f.Open "Select * From " & sTable, CONN1, adOpenKeyset, adLockPessimistic
Set sDGrid.DataSource = rs_customer_f
'remove variable from memory
Set CONN1 = Nothing
Set rs_customer_f = Nothing
End Sub
Public Function getRightString(ByRef sString As String) As String
Dim i    As Byte
Dim strH As String
For i = 1 To Len(sString)
    strH = Left(Right(sString, i), 1)
    If strH = " " Then Exit For
Next i
getRightString = Right(sString, i - 1)
strH = ""
i = 0
End Function
Public Sub getAuto(ByRef sNum As Double, ByVal sTable As String, ByVal sfield As String)
Call DB1(DBPath)
Dim rsRec As New ADODB.Recordset
With rsRec
    .Open "Select * From " & sTable, CONN1, adOpenStatic, adLockOptimistic
        sNum = .Fields(sfield) + 1
End With
Set rsRec = Nothing
Set CONN1 = Nothing
End Sub
Public Sub incrementAuto(ByVal sTable As String, ByVal sfield As String)
Call DB1(DBPath)
Dim rsRec As New ADODB.Recordset
With rsRec
    .Open "Select * From " & sTable, CONN1, adOpenStatic, adLockOptimistic
        .Fields(sfield) = Val(.Fields(sfield)) + 1
        .Update
End With
Set rsRec = Nothing
Set CONN1 = Nothing
End Sub
Public Sub FillListView(ByRef sListView As ListView, ByRef sDB As ADODB.Connection, ByRef sRecordSource As ADODB.Recordset, ByVal sNumOfFields As Byte, ByVal sNumIco As Byte)
Dim x
Dim i As Byte
sRecordSource.MoveFirst
Do While Not sRecordSource.EOF
    Set x = sListView.ListItems.Add(, , sRecordSource.Fields(0), , sNumIco)
        For i = 1 To sNumOfFields - 1
            x.SubItems(i) = sRecordSource.Fields(Val(i))
        Next i
    sRecordSource.MoveNext
Loop
i = 0
Set x = Nothing
Set sRecordSource = Nothing
Set sDB = Nothing
End Sub
Public Function non_millitary_time() As String
Dim H As Byte
Dim T As String
H = Val(Format(Time, "hh"))
If H > 12 Then
    H = H - 12
    T = H
End If
If H < 10 Then
    T = "0" & H
Else
    T = H
End If
non_millitary_time = T & ":" & Format(Time, "mm")
H = 0
T = ""
End Function
Public Function dateFromDB(ByVal sString As String)
Dim d As String
Dim H As String
Dim i As Byte
For i = 1 To Len(sString)
    d = Left(Right(sString, i), 1)
    If d = "-" Then Exit For
    H = d & H
Next i
dateFromDB = H
i = 0
d = ""
H = ""
End Function
Public Function HaveWebAcc(ByVal sCusId As String) As Boolean
HaveWebAcc = True
Dim CustomersWebAccount As New ADODB.Recordset
Call DB1(DBPath)
With CustomersWebAccount
    .Open "Select * From CustomersWebAccount Where CustomerID ='" & sCusId & "'", CONN1, adOpenStatic, adLockPessimistic
    If .RecordCount < 1 Then HaveWebAcc = False
End With
Set CustomersWebAccount = Nothing
Set CONN1 = Nothing
End Function
