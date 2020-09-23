Attribute VB_Name = "Declaration_CompSetup"
Option Explicit
Dim rs_update_comp_info As New ADODB.Recordset
Dim rs_update_comp_user As New ADODB.Recordset
Public Sub CLEAR_VARIABLE_COMPSETUP()
With CurrAndCountry
    .CountryName = ""
    .CurrencySymbol = ""
End With
With UserInfo
    .FullName = ""
    .Password = ""
    .UserName = ""
End With
With CompanyInfo
    .CompanyName = ""
    .ContactName = ""
    .StreetAdd = ""
    .City = ""
    .ZipCode = 0
    .Phone = ""
    .Fax = ""
    .EAdd = ""
    .WebSite = ""
    .BusinessType = ""
End With
End Sub
Public Sub FOCUS_STEXT(ByRef sText As TextBox)
With sText
    .SelStart = 0
    .SelLength = Len(.Text)
End With
End Sub
Public Sub INFORM_VIA_MSG()
MsgBox "Some important field/s is/are empty.", vbExclamation, "Company Setup"
End Sub
Public Function CHECK_COMP_RC_ZERO() As Boolean
CHECK_COMP_RC_ZERO = False
Dim rs_chk_comp_rc As New ADODB.Recordset
Call DB1("Nsap.nsdb")
rs_chk_comp_rc.Open "Select * From Companies", CONN1, adOpenStatic, adLockOptimistic
If Val(rs_chk_comp_rc.RecordCount) < 1 Then CHECK_COMP_RC_ZERO = True
Set rs_chk_comp_rc = Nothing
Set CONN1 = Nothing
End Function
Public Function CORRECT_EMAIL_ADD(ByVal EAdd As String) As Boolean
Dim x As Byte
Dim tmpStr As String
CORRECT_EMAIL_ADD = False
For x = 1 To Len(EAdd)
    tmpStr = Mid(EAdd, x, 1)
    If tmpStr = "@" Then
        CORRECT_EMAIL_ADD = True
        Exit For
    End If
Next x
'Clear variable from computer memory
x = 0
tmpStr = ""
End Function
Public Sub ADD_COMPANY()
Dim rs_add_company As New ADODB.Recordset
Call DB1("Nsap.nsdb")
With rs_add_company
    .Open "Select * From Companies", CONN1, adOpenStatic, adLockOptimistic
        .AddNew
            .Fields(0) = CompanyInfo.CompanyName
            .Fields(1) = "\" & CompanyInfo.CompanyName & "\DataBase\MasterDataBase.nsdb"
        .Update
End With
Set rs_add_company = Nothing
Set CONN1 = Nothing
End Sub
Public Sub CREATE_DATA()
MkDir App.Path & "\" & CompanyInfo.CompanyName
MkDir App.Path & "\" & CompanyInfo.CompanyName & "\DataBase"
FileCopy App.Path & "\Src_data.nsdb", App.Path & "\" & CompanyInfo.CompanyName & "\DataBase\MasterDataBase.nsdb"
Call DB1(CompanyInfo.CompanyName & "\DataBase\MasterDataBase.nsdb")
CONN1.BeginTrans
On Error GoTo err
With rs_update_comp_info
    .Open "Select * From CompanyInfo", CONN1, adOpenStatic, adLockOptimistic
    .AddNew
        .Fields(0) = CompanyInfo.CompanyName
        .Fields(1) = CompanyInfo.ContactName
        .Fields(2) = CompanyInfo.StreetAdd
        .Fields(3) = CompanyInfo.City
        .Fields(4) = CompanyInfo.ZipCode
        .Fields(5) = CurrAndCountry.CountryName
        .Fields(6) = CurrAndCountry.CurrencySymbol
        .Fields(7) = CompanyInfo.Phone
        .Fields(8) = CompanyInfo.Fax
        .Fields(9) = CompanyInfo.WebSite
        .Fields(10) = CompanyInfo.EAdd
        .Fields(11) = CompanyInfo.BusinessType
    .Update
End With
With rs_update_comp_user
    .Open "Select * From Users", CONN1, adOpenStatic, adLockOptimistic
    .AddNew
        .Fields(1) = UserInfo.FullName
        .Fields(2) = UserInfo.UserName
        .Fields(3) = UserInfo.Password
        .Fields(4) = "Administrator"
    .Update
End With
CONN1.CommitTrans
Call CREATE_DATA_clearMemo
Exit Sub
err:
    CONN1.RollbackTrans
    MsgBox "Error Number: " & err.Number & vbCrLf & _
           "Error Source: " & err.Source & vbCrLf & _
           "Description: " & err.Description, vbCritical, "Terminate NaparanSoft"
    End
End Sub
Private Sub CREATE_DATA_clearMemo()
Set rs_update_comp_info = Nothing
Set CONN1 = Nothing
End Sub
Public Function COMPANY_NAME_EXIST(ByVal srcCOMP_NAME As String) As Boolean
COMPANY_NAME_EXIST = False
Dim rs_check_company As New ADODB.Recordset
Call DB1("Nsap.nsdb")
With rs_check_company
    .Open "Select * From Companies", CONN1, adOpenStatic, adLockOptimistic
    .Filter = "CompanyName ='" & srcCOMP_NAME & "'"
    If .RecordCount > 0 Then COMPANY_NAME_EXIST = True
End With
Set rs_check_company = Nothing
Set CONN1 = Nothing
End Function
