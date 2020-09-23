Attribute VB_Name = "Initialization"
Option Explicit
Public Sub fillTreeView(sTV As TreeView)
Call sTV.Nodes.Add(, , "p1", "Program Menu", 1, 1)
    Call sTV.Nodes.Add("p1", tvwChild, "1-c1", "Sales", 2, 2)
        Call sTV.Nodes.Add("1-c1", tvwChild, "1-cc1", "Invoices", 12, 12)
        Call sTV.Nodes.Add("1-c1", tvwChild, "1-cc2", "Orders", 13, 13)
        Call sTV.Nodes.Add("1-c1", tvwChild, "1-cc3", "Customers", 14, 14)
        Call sTV.Nodes.Add("1-c1", tvwChild, "1-cc4", "Payments", 15, 15)
        Call sTV.Nodes.Add("1-c1", tvwChild, "1-cc5", "Sales Reps", 16, 16)
        Call sTV.Nodes.Add("1-c1", tvwChild, "1-cc6", "Sales Tax", 17, 17)
    Call sTV.Nodes.Add("p1", tvwChild, "1-c2", "Purchasing", 3, 3)
        Call sTV.Nodes.Add("1-c2", tvwChild, "2-cc1", "Show Invoices", 12, 12)
    Call sTV.Nodes.Add("p1", tvwChild, "1-c3", "Inventory", 4, 4)
        Call sTV.Nodes.Add("1-c3", tvwChild, "3-cc1", "Show Invoices", 12, 12)
    Call sTV.Nodes.Add("p1", tvwChild, "1-c4", "Accounting", 9, 9)
        Call sTV.Nodes.Add("1-c4", tvwChild, "4-cc1", "Show Invoices", 12, 12)
    Call sTV.Nodes.Add("p1", tvwChild, "1-c5", "Payroll", 6, 6)
        Call sTV.Nodes.Add("1-c5", tvwChild, "5-cc1", "Show Invoices", 12, 12)
    Call sTV.Nodes.Add("p1", tvwChild, "1-c6", "Contacts", 7, 7)
        Call sTV.Nodes.Add("1-c6", tvwChild, "6-cc1", "Show Invoices", 12, 12)
    Call sTV.Nodes.Add("p1", tvwChild, "1-c7", "System Manager", 8, 8)
        Call sTV.Nodes.Add("1-c7", tvwChild, "7-cc1", "RecycleBin", 16, 16)
    Call sTV.Nodes.Add("p1", tvwChild, "1-c8", "Sales Reports", 11, 11)
        Call sTV.Nodes.Add("1-c8", tvwChild, "8-cc1", "Show Invoices", 12, 12)
    Call sTV.Nodes.Add("p1", tvwChild, "1-c9", "Help", 10, 10)
        Call sTV.Nodes.Add("1-c9", tvwChild, "9-cc1", "Show Invoices", 12, 12)
sTV.Nodes("p1").Expanded = True
End Sub
Public Sub centerForm(ByRef sForm As Form, ByVal sHeight As Integer, ByVal sWidth As Integer)
sForm.Move (sWidth - sForm.Width) / 2, (sHeight - sForm.Height) / 2
End Sub
Public Sub write_get_INI(ByRef sboolean As Boolean, OptionWrite As Boolean)
Dim tmpStr As String
On Error GoTo err:
'if true then write
If OptionWrite = True Then
    Open App.Path & "\resources\BM.ini" For Output As #1
        Print #1, sboolean
    Close #1
'if false then get
Else
    Open App.Path & "\resources\BM.ini" For Input As #2
        Input #2, tmpStr
        sboolean = tmpStr
    Close #2
End If
Exit Sub
err:
    MsgBox err.Description, vbExclamation, "Business Magic"
End Sub
