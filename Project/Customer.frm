VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_customer 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Customer"
   ClientHeight    =   10950
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   14820
   DrawStyle       =   6  'Inside Solid
   BeginProperty Font 
      Name            =   "MT Extra"
      Size            =   12
      Charset         =   2
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   10950
   ScaleWidth      =   14820
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   660
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   7800
      Width           =   2775
   End
   Begin VB.CommandButton Command2 
      Caption         =   "save"
      DownPicture     =   "Customer.frx":0000
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   5160
      Picture         =   "Customer.frx":0279
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "save record"
      Top             =   9960
      Width           =   2655
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   495
      Left            =   3840
      TabIndex        =   6
      Top             =   6480
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   114753537
      CurrentDate     =   40805
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   495
      Left            =   3825
      TabIndex        =   4
      Top             =   3960
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   114753537
      CurrentDate     =   40805
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   3840
      TabIndex        =   5
      Top             =   5160
      Width           =   2775
   End
   Begin VB.CommandButton add 
      Caption         =   "Add Record"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1920
      Picture         =   "Customer.frx":04F2
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Add New Record"
      Top             =   9960
      Width           =   2655
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   3825
      TabIndex        =   3
      Top             =   2880
      Width           =   2775
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   3840
      MaxLength       =   10
      TabIndex        =   2
      Top             =   1800
      Width           =   2775
   End
   Begin VB.TextBox Txtcname 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   3825
      TabIndex        =   1
      Top             =   720
      Width           =   2775
   End
   Begin VB.Image Image2 
      Height          =   9015
      Left            =   10320
      Picture         =   "Customer.frx":063C
      Stretch         =   -1  'True
      Top             =   480
      Width           =   6420
   End
   Begin VB.Image Image1 
      Height          =   615
      Left            =   0
      Picture         =   "Customer.frx":85C8
      Stretch         =   -1  'True
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Customer Number"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   720
      TabIndex        =   15
      Top             =   8040
      Width           =   2655
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Alert Date"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      TabIndex        =   14
      Top             =   6600
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Type"
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
      Index           =   5
      Left            =   720
      TabIndex        =   13
      Top             =   5280
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Date Of Birth"
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
      Index           =   4
      Left            =   840
      TabIndex        =   12
      Top             =   4200
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Email Id"
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
      Index           =   3
      Left            =   840
      TabIndex        =   11
      Top             =   3000
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Contact Number"
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
      Index           =   2
      Left            =   765
      TabIndex        =   10
      Top             =   1920
      Width           =   2775
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Customer Name"
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
      Index           =   1
      Left            =   765
      TabIndex        =   0
      Top             =   960
      Width           =   3015
   End
End
Attribute VB_Name = "frm_customer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim symbol1 As Integer
Dim TestString As String

Private Sub add_Click()
Txtcname.Text = ""
Text4.Text = ""
Text3.Text = ""
Text5.Text = ""
Text1.Text = ""


End Sub

Private Sub Command2_Click()
If Len(Txtcname.Text) = 0 Then
MsgBox "Please Enter Customer Name", vbInformation
Txtcname.SetFocus
Exit Sub
End If
If Len(Text3.Text) = 0 Then
MsgBox "Please Enter Number", vbInformation
Text3.SetFocus
Exit Sub
End If
TestString = Text4.Text
symbol1 = InStr(TestString, ".")
a = 0
For i = 2 To Len(Text4.Text) - 1
    If Mid(Text4.Text, i, 1) = "@" And (symbol1 <> 0) And Mid(Text4.Text, Len(Text4.Text)) <> "." Then
        a = 1
        GoTo aaaa
    End If
Next i
aaaa:
    If a = 0 Then
        MsgBox "Please,Enter Valid Email-id", vbInformation, "Message"
        Text4.SetFocus
        Exit Sub
    End If
  If Len(Text5.Text) = 0 Then
MsgBox "Please Enter Type", vbInformation
Text5.SetFocus
Exit Sub
End If
 
 rs.Open "select * from customer", con, 2, 2
 rs.AddNew
 rs.Fields(0) = Val(Text1.Text)
 rs.Fields(1) = Txtcname.Text
 rs.Fields(2) = Text3.Text
 rs.Fields(3) = Text4.Text
 rs.Fields(4) = DTPicker1.Value
 rs.Fields(5) = Text5.Text
 rs.Fields(6) = DTPicker2.Value
rs.Update
rs.Close
MsgBox "Record Saved successfully", vbInformation
Txtcname.Text = ""
Text4.Text = ""
Text3.Text = ""
Text5.Text = ""
Text1.Text = ""
DTPicker1.Value = Now()
DTPicker2.Value = Now()

End Sub

Private Sub DTPicker1_Change()
DTPicker1.MaxDate = Now

End Sub



Private Sub DTPicker2_Change()
DTPicker2.MinDate = Now
End Sub

Private Sub Form_Load()
Set con = New ADODB.Connection
con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\lic.mdb"
con.Open
Set rs = New ADODB.Recordset
'rs.Open "select * from Customer", con, adOpenDynamic, adLockOptimistic
'rs.Close
'Set rs = New ADODB.Recordset
'If rs.State = 0 Then
rs.Open "select max(Customer_Number) from Customer", con, adOpenDynamic, adLockOptimistic
'MsgBox "ok"
'End If
Txtcname.Text = ""

Text3.Text = ""
Text4.Text = ""
Text5.Text = ""

'If rs.BOF = False And rs.EOF = False Then
'rs.MoveLast
'If rs.RecordCount <> -1 Then
''Text1.Text = rs.Fields(0) + 1
'Text1.Text = 1
'Else
''Text1.Text = 1
'Text1.Text = rs.Fields(0) + 1
'End If
''Else:
''Text1.Text = 1
''End If
''rs.AddNew
rs.Close
DTPicker1.Value = Now()
DTPicker2.Value = Now()

End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Or KeyAscii = 27 Then
ElseIf KeyAscii = 13 Then
Text3.SetFocus
ElseIf (KeyAscii >= 48 And KeyAscii <= 57) Then

Else
MsgBox "Enter Number Only", vbInformation
KeyAscii = 0
End If
End Sub

Private Sub Text3_LostFocus()
'Dim a As Integer
'
'a = Len(Txtcname.Text)
'If a = 0 Then
'MsgBox "Enter the data", vbQuestion
'Text3.SetFocus
'End If

End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Or KeyAscii = 27 Then
ElseIf KeyAscii = 13 Then
Text5.SetFocus
ElseIf (KeyAscii >= 65 And KeyAscii <= 91) Or (KeyAscii >= 97 And KeyAscii <= 122) Then

Else
MsgBox "Enter Character Only"
KeyAscii = 0
End If
End Sub

Private Sub Text5_LostFocus()
rs.Open "select max(Customer_Number) from customer", con, 2, 2
On Error GoTo x:
Text1.Text = rs.Fields(0).Value + 1

rs.Close
Exit Sub
x:
Text1.Text = 1
rs.Close
End Sub
Private Sub Txtcname_KeyPress(KeyAscii As Integer)
If KeyAscii <= 64 Or KeyAscii >= 122 Then
If KeyAscii <> 32 Then
If KeyAscii <> 8 Then
KeyAscii = 0
End If
End If
End If
End Sub

Private Sub Txtcname_LostFocus()
'Dim s As Integer
'
's = Len(Txtcname.Text)
'If s = 0 Then
'MsgBox "Enter the data", vbQuestion
'Txtcname.SetFocus
'End If

End Sub
