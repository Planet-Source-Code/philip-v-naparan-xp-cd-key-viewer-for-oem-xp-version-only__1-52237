VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "XP CD-Key Viewer by Philip V. Naparan"
   ClientHeight    =   1980
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4890
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1980
   ScaleWidth      =   4890
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   1800
      Top             =   600
   End
   Begin VB.CommandButton Command3 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "About"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1400
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1400
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Browse the XP Installer"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1400
      Width           =   2295
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   4455
      Begin VB.TextBox Text1 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Left            =   480
         Locked          =   -1  'True
         TabIndex        =   3
         Text            =   "CD-Key of XP will be display here!"
         Top             =   240
         Width           =   3855
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   -120
         Picture         =   "Form1.frx":1042
         Top             =   150
         Width           =   480
      End
   End
   Begin VB.DirListBox Dir1 
      Height          =   315
      Left            =   4920
      TabIndex        =   1
      Top             =   2400
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.FileListBox File1 
      Height          =   285
      Left            =   4920
      TabIndex        =   0
      Top             =   2760
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Pls. don't forget to vote !!!!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   1080
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "xp"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3880
      TabIndex        =   8
      Top             =   870
      Width           =   255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Windows "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2640
      TabIndex        =   7
      Top             =   960
      Width           =   1335
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   4200
      Picture         =   "Form1.frx":190C
      Top             =   840
      Width           =   480
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H8000000F&
      BorderWidth     =   2
      Height          =   1965
      Left            =   15
      Top             =   15
      Width           =   4875
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   100
      Picture         =   "Form1.frx":25EE
      Top             =   270
      Width           =   480
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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


Dim IniFiles As String
 
Private Sub Command1_Click()
'-----------------------------------
'Browse the the XP Installer
'-----------------------------------
Dim strResFolder As String
strResFolder = LocateInstaller(hWnd, "Please select the folder where the XP Installer is located."): tmpF = "Winnt.sif"
If strResFolder <> "" Then
    Screen.MousePointer = vbHourglass
    IniFiles = ""
    Call DirWalk("*.*", strResFolder)
    If IniFiles = "" Then
        MsgBox "Cannot find the required files.", vbExclamation, "XP CD-Key Viewer"
    Else
        Text1.Text = ReadIniFile(IniFiles, "UserData", "ProductID", "")
        If Text1.Text = "" Then
            Text1.Text = ReadIniFile(IniFiles, "UserData", "ProductKey", "")
        Else
            MsgBox "Cannot open the required files.", vbExclamation, "XP CD-Key Viewer"
        End If
        MsgBox "Product key has been successfully captured.", vbInformation, "XP CD-Key Viewer"
    End If
    Screen.MousePointer = vbDefault
End If
'-----------------------------------
'End browsing
'-----------------------------------
End Sub
Sub DirWalk(ByVal sPattern As String, ByVal CurrDir As String)
'------------------------------------------------
'The code bellow was taken from file search
'application which can be found in
'www.a1vbcode.com.
'
'Note: I crated a drive crawler by my my own but
'it is  a little bit slow compare to this one.
'--------------------------------------------------
On Error GoTo Err
Dim i As Integer
Dim sCurrPath As String
Dim sFile As String
Dim ii As Integer
Dim iFiles As Integer
Dim iLen As Integer
Dim tmpH As String

If Right$(CurrDir, 1) <> "\" Then
    Dir1.Path = CurrDir & "\"
Else
    Dir1.Path = CurrDir
End If
For i = 0 To Dir1.ListCount
    If Dir1.List(i) <> "" Then
        DoEvents
        Call DirWalk(sPattern, Dir1.List(i))
    Else
        If Right$(Dir1.Path, 1) = "\" Then
            sCurrPath = Left(Dir1.Path, Len(Dir1.Path) - 1)
        Else
            sCurrPath = Dir1.Path
        End If
        File1.Path = sCurrPath: File1.Pattern = sPattern
        If File1.ListCount > 0 Then
            For ii = 0 To File1.ListCount - 1: tmpH = sCurrPath & "\" & File1.List(ii): If LCase(file_from_path(tmpH)) = LCase(tmpF) Then IniFiles = tmpH: tmpH = ""
            Next ii
        End If
        iLen = Len(Dir1.Path)
        Do While Mid(Dir1.Path, iLen, 1) <> "\"
            iLen = iLen - 1
        Loop
        Dir1.Path = Mid(Dir1.Path, 1, iLen)
    End If
Next i
Exit Sub
Err:
    MsgBox "Cannot read files from the selected path.", vbCritical, "Error Occur"
    Exit Sub
'-----------------------------------------------
'End the barrowed code
'-----------------------------------------------
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
Form2.Show vbModal
End Sub

Private Sub Form_Load()
Text1.ForeColor = RGB(255, 255, 18)
Me.BackColor = RGB(228, 230, 240)
Frame1.BackColor = RGB(228, 230, 240)
End Sub

Private Sub Timer1_Timer()
Label3.Visible = Not Label3.Visible
End Sub
