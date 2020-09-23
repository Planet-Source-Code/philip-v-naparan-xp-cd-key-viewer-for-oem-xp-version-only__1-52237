Attribute VB_Name = "modAplication"
Option Explicit
Public Sub getURL(urlADD As String, sourceHWND As String)
Call ShellExecute(sourceHWND, vbNullString, urlADD, "", vbNullString, 1)
End Sub

Public Sub prompt_err(ByVal sErrorDescription As String)
MsgBox sErrorDescription, vbExclamation, "Media Tracker"
End Sub
Public Sub select_color_type(ByVal sColorOption As Byte)
Select Case sColorOption
    Case 0: '[ XP Default ]
            New_System_Color.SelectColor(4) = RGB(239, 238, 224)  'Menu
            New_System_Color.SelectColor(15) = RGB(240, 240, 224) 'Button
            New_System_Color.SelectColor(16) = RGB(216, 210, 189) 'Button Shadow
            New_System_Color.SelectColor(20) = RGB(255, 255, 255) 'Button Highlight

            Call change_system_color
            
    Case 1: '[ Mac Grey ]
            New_System_Color.SelectColor(4) = RGB(235, 235, 235)  'Menu
            New_System_Color.SelectColor(15) = RGB(235, 235, 235) 'Button
            New_System_Color.SelectColor(16) = RGB(186, 186, 186) 'Button Shadow
            New_System_Color.SelectColor(20) = RGB(255, 255, 255) 'Button Highlight
            
            Call change_system_color
            
    Case 2: '[ XP Blue ]
            New_System_Color.SelectColor(4) = RGB(211, 229, 251)    'Menu
            New_System_Color.SelectColor(15) = RGB(211, 229, 251) 'Button
            New_System_Color.SelectColor(16) = RGB(139, 188, 254)  'Button Shadow
            New_System_Color.SelectColor(20) = RGB(255, 255, 255) 'Button Highlight
            
            Call change_system_color
    
    Case 3: '[ Cool Green ]
            New_System_Color.SelectColor(4) = RGB(217, 238, 205)   'Menu
            New_System_Color.SelectColor(15) = RGB(217, 238, 205) 'Button
            New_System_Color.SelectColor(16) = RGB(149, 207, 114)  'Button Shadow
            New_System_Color.SelectColor(20) = RGB(255, 255, 255) 'Button Highlight
            
            Call change_system_color
            
    Case 4: '[ Light Violet ]
            New_System_Color.SelectColor(4) = RGB(220, 220, 223)  'Menu
            New_System_Color.SelectColor(15) = RGB(220, 220, 223) 'Button
            New_System_Color.SelectColor(16) = RGB(185, 191, 199)  'Button Shadow
            New_System_Color.SelectColor(20) = RGB(235, 244, 255) 'Button Highlight
            
            Call change_system_color
            
    Case 5: '[ Light Brown ]
            New_System_Color.SelectColor(4) = RGB(218, 214, 206)   'Menu
            New_System_Color.SelectColor(15) = RGB(218, 214, 206) 'Button
            New_System_Color.SelectColor(16) = RGB(167, 163, 155)  'Button Shadow
            New_System_Color.SelectColor(20) = RGB(235, 231, 223)  'Button Highlight
            
            Call change_system_color
        
    Case 6: '[ Win Classic ]
            New_System_Color.SelectColor(4) = RGB(212, 208, 200)    'Menu
            New_System_Color.SelectColor(15) = RGB(212, 208, 200) 'Button
            New_System_Color.SelectColor(16) = RGB(128, 128, 128) 'Button Shadow
            New_System_Color.SelectColor(20) = RGB(255, 255, 255)  'Button Highlight
            
            Call change_system_color
            
End Select
End Sub
Public Sub change_system_color()

Call SetSysColors(1, 4, New_System_Color.SelectColor(4))   'Menu
Call SetSysColors(1, 15, New_System_Color.SelectColor(15)) 'Button
Call SetSysColors(1, 16, New_System_Color.SelectColor(16)) 'Button Shadow
Call SetSysColors(1, 20, New_System_Color.SelectColor(20)) 'Button Highlight

End Sub
Public Function folder_from_path(ByVal spath As String) As String
Dim c As Integer
Dim tmp_h As String
For c = 1 To Len(spath)
    tmp_h = Left(Right(spath, c), 1)
    If tmp_h = "\" Then Exit For
    folder_from_path = tmp_h & folder_from_path
Next c
tmp_h = ""
c = 0
End Function
Public Function file_from_path(ByVal spath As String) As String
Dim c As Integer
Dim tmp_h As String
For c = 1 To Len(spath)
    tmp_h = Left(Right(spath, c), 1)
    If tmp_h = "\" Or tmp_h = "/" Then Exit For
    file_from_path = tmp_h & file_from_path
Next c
tmp_h = ""
c = 0
End Function
Public Function file_type(ByVal sFile As String) As String
Dim c As Integer
Dim tmp_h As String
For c = 1 To Len(sFile)
    tmp_h = Left(Right(sFile, c), 1)
    If tmp_h = "." Then Exit For
    file_type = tmp_h & file_type
Next c
file_type = "." & file_type
tmp_h = ""
c = 0
End Function
Public Function correct_email(ByRef sEntryField As Variant, ByRef sSSTab As SSTab, ByVal caption As Boolean, ByVal have_sstab As Boolean, ByVal sTab_num As Byte) As Boolean
Dim c As Integer
Dim at_finder As String
correct_email = False
If caption = True Then
    For c = 1 To Len(sEntryField.caption)
        If Mid(sEntryField.caption, c, 1) = "@" Then
            correct_email = True
            Exit Function
        End If
    Next c
Else
    For c = 1 To Len(sEntryField.Text)
        If Mid(sEntryField.Text, c, 1) = "@" Then
            correct_email = True
            Exit Function
        End If
    Next c
End If
MsgBox "Invalid e-mail address.Please check it!", vbExclamation, "Business Magic"
If have_sstab = True Then sSSTab.Tab = sTab_num
sEntryField.SetFocus
'Clear variable
at_finder = ""
c = 0
End Function
Public Function parent_dir(ByVal sDir As String) As String
Dim c As Integer
Dim tmp_h As String
For c = 1 To Len(sDir)
    tmp_h = Left(Right(sDir, c), 1)
    If tmp_h = "\" Then Exit For
    parent_dir = Left(sDir, Len(sDir) - (c + 1))
Next c
tmp_h = ""
c = 0
End Function
Public Sub open_help_file(ByVal sHelpFileLocation As String, ByVal sCommandNum As Long, ByVal sHlpPageNum As Long)
WinHelp MDI_MAIN.hwnd, sHelpFileLocation, sCommandNum, sHlpPageNum
End Sub
Public Sub search_in_listview(ByRef sListView As ListView, ByVal sFindText As String)
Dim tmp_listtview As ListItem
Set tmp_listtview = sListView.FindItem(sFindText, lvwSubItem, , lvwPartial)
If Not tmp_listtview Is Nothing Then
    tmp_listtview.EnsureVisible
    tmp_listtview.Selected = True
End If
End Sub
Public Function pic_path_from_html_tag(ByVal sHTMLTag As String) As String
Dim lc1, lc2 As Integer 'lc stand for letter counter
Dim tmpStr1 As String

For lc1 = 1 To Len(sHTMLTag)
    tmpStr1 = Mid(sHTMLTag, lc1, 1)
    If tmpStr1 = """" Then
        For lc2 = lc1 To Len(sHTMLTag)
            tmpStr1 = Mid(sHTMLTag, (lc2 + 1), 1)
            If tmpStr1 = """" Then Exit Function Else pic_path_from_html_tag = pic_path_from_html_tag & tmpStr1
        Next lc2
    End If
Next lc1

lc1 = 0
lc2 = 0
tmpStr1 = ""
End Function
Public Function write_html_tag(ByVal sFilePath As String) As String
write_html_tag = "<img src=" & Chr(34) & sFilePath & Chr(34) & ">"
End Function
