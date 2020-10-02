VERSION 5.00
Begin VB.Form frm_login 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   0  'None
   Caption         =   "Login"
   ClientHeight    =   10245
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14415
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   ScaleHeight     =   10245
   ScaleWidth      =   14415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   10965
      Top             =   6375
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1020
      Top             =   6630
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   3600
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1440
      Width           =   4125
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   495
      Left            =   3600
      MaxLength       =   10
      TabIndex        =   1
      Top             =   480
      Width           =   4125
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Projected By  - Yash Sharma and Yash Umapurkar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5760
      TabIndex        =   4
      Top             =   9480
      Width           =   9135
   End
   Begin VB.Image Image1 
      Height          =   5055
      Left            =   1320
      Picture         =   "ª.frx":0000
      Stretch         =   -1  'True
      Top             =   3960
      Width           =   12060
   End
   Begin VB.Image Image4 
      Height          =   480
      Left            =   6120
      Picture         =   "ª.frx":F848
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   1740
   End
   Begin VB.Image Image3 
      Height          =   495
      Left            =   3720
      Picture         =   "ª.frx":10878
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "User Name"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   1275
      TabIndex        =   3
      Top             =   510
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      TabIndex        =   0
      Top             =   1440
      Width           =   2175
   End
End
Attribute VB_Name = "frm_login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim str1 As String
Dim str2 As Variant
Dim str3 As String
Dim str4 As Variant





Private Sub Form_Load()
'Sb1.Panels(1) = Format(Now, "dd / mm / yy")
End Sub

Private Sub Image3_Click()
str1 = "Lic"
str2 = "Lic"
If Text1.Text = str1 And Text2.Text = str2 Then
MsgBox "login successful"
Unload Me
frm_start.Show
Else
MsgBox "User Name or Password wrong", vbInformation
End If
End Sub

Private Sub Image4_Click()
End
End Sub

Private Sub Timer1_Timer()
'Sb1.Panels(2) = Format(Now, "hh:mm:ss")
End Sub

