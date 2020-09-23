VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About the Programmer"
   ClientHeight    =   6270
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4785
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6270
   ScaleWidth      =   4785
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   $"Form2.frx":1042
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   120
      TabIndex        =   9
      Top             =   5640
      Width           =   4695
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Contact #:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   8
      Top             =   5400
      Width           =   975
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "E-mail:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   7
      Top             =   5160
      Width           =   975
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "639186443161"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2280
      TabIndex        =   6
      Top             =   5400
      Width           =   1215
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "philipnaparan@yahoo.com"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2280
      TabIndex        =   5
      Top             =   5160
      Width           =   1935
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "www.philipnaparan.cjb.net"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2280
      TabIndex        =   4
      Top             =   4920
      Width           =   2055
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "www.philipnaparan.com"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2280
      TabIndex        =   3
      Top             =   4680
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Website:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   2
      Top             =   4680
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Philip V. Naparan"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2280
      TabIndex        =   1
      Top             =   4440
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Programmer:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   0
      Top             =   4440
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C00000&
      BorderWidth     =   2
      Height          =   4210
      Left            =   120
      Top             =   120
      Width           =   4530
   End
   Begin VB.Image Image1 
      Height          =   4200
      Left            =   120
      Picture         =   "Form2.frx":10EF
      Top             =   120
      Width           =   4500
   End
   Begin VB.Image Image5 
      Height          =   960
      Left            =   120
      Picture         =   "Form2.frx":315D
      Top             =   4440
      Width           =   960
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Me.BackColor = RGB(228, 230, 240)
End Sub

