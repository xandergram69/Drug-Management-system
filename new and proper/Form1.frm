VERSION 5.00
Begin VB.Form loginn 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DRUG MANAGEMENT AND INFORMATION SYSTEM"
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5820
   FillColor       =   &H00FFFFFF&
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   1  'Arrow
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   3.833
   ScaleMode       =   5  'Inch
   ScaleWidth      =   4.042
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text4 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   36
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   735
      Left            =   2280
      TabIndex        =   8
      Top             =   4200
      Width           =   975
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      IMEMode         =   3  'DISABLE
      Left            =   2640
      PasswordChar    =   "*"
      TabIndex        =   6
      Top             =   2760
      Width           =   2655
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   2640
      TabIndex        =   5
      Top             =   1920
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   2640
      TabIndex        =   4
      Top             =   960
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C000&
      Caption         =   "LOGIN"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      MaskColor       =   &H00808080&
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3600
      Width           =   3495
   End
   Begin VB.Label Label4 
      Caption         =   "PASSWORD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   2880
      Width           =   2175
   End
   Begin VB.Label Label3 
      Caption         =   "EMAIL ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "NAME"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "LOGIN "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   23.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Left            =   1920
      TabIndex        =   0
      Top             =   240
      Width           =   1455
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000009&
      BorderStyle     =   3  'Dot
      FillStyle       =   0  'Solid
      Height          =   5055
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   5415
   End
End
Attribute VB_Name = "loginn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
 

Dim a, b, c, d
a = Text1.Text
b = Text2.Text
c = Text3.Text

If (a = "admin" And b = "admin@gmail.com" And c = "admin") Then
MsgBox ("welcome")
CreateObject("sapi.SPvoice").speak ("welcome")
MDIForm1.Show
Else
MsgBox ("wrong credientals")
End If
End Sub


Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command1.BackColor = RGB(230, 73, 25)
End Sub

Private Sub Form_Load()
Command1.BackColor = RGB(96, 3, 252)

Label1.BackColor = RGB(6, 12, 33)
Label1.ForeColor = RGB(135, 254, 4)
Label2.BackColor = RGB(6, 12, 33)
Label3.BackColor = RGB(6, 12, 33)
Label4.BackColor = RGB(6, 12, 33)
Text4.BackColor = RGB(6, 12, 33)
Text1.BackColor = RGB(6, 12, 33)
Text2.BackColor = RGB(6, 12, 33)
Text3.BackColor = RGB(6, 12, 33)
Text1.ForeColor = RGB(46, 204, 113)
Text2.ForeColor = RGB(46, 204, 113)
Text3.ForeColor = RGB(46, 204, 113)
Shape1.FillColor = RGB(6, 12, 33)
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
 Select Case KeyAscii
  Case 32 To 64, 91 To 96, 123 To 126
     MsgBox ("must be a letter ! please try again !")
     KeyAscii = 0
   Exit Sub
 End Select
End Sub



Private Sub Text4_Change()
If (Text4.Text = "a" Or Text4.Text = "A") Then
Text1.Text = "admin"
Text2.Text = "mail@admin"
Text3.Text = "#access"
End If


End Sub
