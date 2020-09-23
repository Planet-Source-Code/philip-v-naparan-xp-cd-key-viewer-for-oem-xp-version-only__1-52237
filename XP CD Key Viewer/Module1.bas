Attribute VB_Name = "Module1"
'********************************************************
'Programmer: Philip V. Naparan
'Website: www.philipnaparan.com,www.philipnaparan.cjb.net
'E-mail Address: philipnaparan@yahoo.com
'Contact Number: 639186443161
'
'WARNING: Do not distribute the code without asking
'         permission from the author and donot used
'         this product for illegal purpose.
'********************************************************
Option Explicit


Public Const BIF_RETURNONLYFSDIRS = 1
Public Const MAX_PATH = 260

Public Type Locating
     hwndOwner As Long
     pIDLRoot As Long
     pszDisplayName As Long
     lpszTitle As Long
     ulFlags As Long
     lpfnCallback As Long
     lParam As Long
     iImage As Long
End Type
Declare Function GetPrivateProfileString Lib "kernel32" Alias _
      "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal _
      lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As _
      String, ByVal nSize As Long, ByVal lpFileName As String) As Long: Global tmpF As String
Declare Function MciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As Any, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Public Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Public Declare Function SHBrowseForFolder Lib "shell32" (lpbi As Locating) As Long
Public Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long

Public Function LocateInstaller(hwndOwner As Long, sPrompt As String) As String
'-----------------------------------------------
'A function that return the path of the selected
'directory.
'-----------------------------------------------
     Dim iNull As Integer
     Dim lpIDList As Long
     Dim lResult As Long
     Dim spath As String
     Dim udtBI As Locating

     With udtBI
        .hwndOwner = hwndOwner
        .lpszTitle = lstrcat(sPrompt, "")
        .ulFlags = BIF_RETURNONLYFSDIRS
     End With

     lpIDList = SHBrowseForFolder(udtBI)
     
     If lpIDList Then
        spath = String$(MAX_PATH, 0)
        lResult = SHGetPathFromIDList(lpIDList, spath)
        Call CoTaskMemFree(lpIDList)
        iNull = InStr(spath, vbNullChar)
        If iNull Then spath = Left$(spath, iNull - 1)
     End If

     LocateInstaller = spath
'-----------------------------------------------
'End function
'-----------------------------------------------
End Function

Public Function file_from_path(ByVal spath As String) As String
'--------------------------------------------
'Get only the file
'--------------------------------------------
Dim c As Integer
Dim tmp_h As String
For c = 1 To Len(spath)
    tmp_h = Left(Right(spath, c), 1)
    If tmp_h = "\" Or tmp_h = "/" Then Exit For
    file_from_path = tmp_h & file_from_path
Next c
tmp_h = ""
c = 0
'---------------------------------------------
'End getting
'---------------------------------------------
End Function
Function ReadIniFile(ByVal sIniFileName As String, ByVal sSection As String, ByVal sItem As String, ByVal sDefault As String) As String
'-----------------------------------------------
'Read initialization file
'-----------------------------------------------
   Dim iRetAmount As Integer
   Dim sTemp As String

   sTemp = String$(50, 0)
   iRetAmount = GetPrivateProfileString(sSection, sItem, sDefault, sTemp, 50, sIniFileName)
   sTemp = Left$(sTemp, iRetAmount)
   ReadIniFile = sTemp
'-----------------------------------------------
'End Reading
'-----------------------------------------------
End Function
