Attribute VB_Name = "ADOConnections"
Option Explicit


Global CONN1    As New ADODB.Connection
Global DataPath As String
Global DBPath   As String
Global CompName As String
Global CompPath As String
Global CurrName As String
Public Sub DB1(ByVal DataName As String)
On Error GoTo err
CONN1.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\" & DataName & ";Persist Security Info=False;Jet OLEDB:Database Password=!@#$%philipnaparan^&"
Exit Sub
err:
MsgBox "Error Number: " & err.Number & vbCrLf & _
           "Error Source: " & err.Source & vbCrLf & _
           "Description: " & err.Description, vbCritical, "Terminate NaparanSoft"
End
End Sub
