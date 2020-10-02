VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_calc 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Premium Caculator"
   ClientHeight    =   11430
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   18960
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   11430
   ScaleWidth      =   18960
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command3 
      Caption         =   "Final Premium"
      Height          =   735
      Left            =   6240
      TabIndex        =   33
      Top             =   9480
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Back"
      Height          =   735
      Left            =   600
      TabIndex        =   31
      Top             =   9840
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      Height          =   900
      Left            =   5520
      Picture         =   "frm_calc.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   30
      ToolTipText     =   "Refresh"
      Top             =   7680
      Width           =   1035
   End
   Begin VB.TextBox Txtprm 
      Height          =   450
      Left            =   0
      TabIndex        =   29
      Top             =   8760
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.TextBox finalprem 
      Height          =   1035
      Left            =   9480
      Locked          =   -1  'True
      TabIndex        =   28
      Top             =   9240
      Width           =   3840
   End
   Begin VB.ComboBox Cspecial 
      Height          =   360
      ItemData        =   "frm_calc.frx":0621
      Left            =   10800
      List            =   "frm_calc.frx":062B
      TabIndex        =   21
      Text            =   "Special Plan"
      Top             =   720
      Width           =   2055
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   780
      Left            =   0
      Top             =   8640
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   1376
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Riders"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   1320
      TabIndex        =   8
      Top             =   7080
      Width           =   5655
      Begin VB.CheckBox Check4 
         BackColor       =   &H00FFC0C0&
         Caption         =   "TERM RIDER"
         Height          =   375
         Left            =   1920
         TabIndex        =   20
         Top             =   1200
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00FFC0C0&
         Caption         =   "CIR"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   1320
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "PWB"
         Height          =   255
         Left            =   1920
         TabIndex        =   18
         Top             =   480
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "DAB"
         Height          =   255
         Left            =   255
         TabIndex        =   17
         Top             =   480
         Visible         =   0   'False
         Width           =   1455
      End
   End
   Begin VB.ComboBox Combo15 
      Height          =   360
      ItemData        =   "frm_calc.frx":064B
      Left            =   4560
      List            =   "frm_calc.frx":065E
      TabIndex        =   7
      Text            =   "Occupation"
      Top             =   6120
      Width           =   2295
   End
   Begin VB.ComboBox Combo14 
      Height          =   360
      ItemData        =   "frm_calc.frx":0699
      Left            =   4560
      List            =   "frm_calc.frx":06AC
      TabIndex        =   6
      Text            =   "Health"
      Top             =   5160
      Width           =   2295
   End
   Begin VB.ComboBox Combo9 
      Height          =   360
      ItemData        =   "frm_calc.frx":06E7
      Left            =   4560
      List            =   "frm_calc.frx":06E9
      TabIndex        =   5
      Text            =   "Select"
      Top             =   4200
      Width           =   2310
   End
   Begin VB.ComboBox Combo8 
      Height          =   360
      ItemData        =   "frm_calc.frx":06EB
      Left            =   4560
      List            =   "frm_calc.frx":06FB
      TabIndex        =   4
      Text            =   "Select"
      Top             =   3240
      Width           =   2310
   End
   Begin VB.ComboBox Combo7 
      Height          =   360
      ItemData        =   "frm_calc.frx":0737
      Left            =   4560
      List            =   "frm_calc.frx":074A
      TabIndex        =   3
      Text            =   "Select"
      Top             =   2400
      Width           =   2310
   End
   Begin VB.ComboBox cage 
      Height          =   360
      ItemData        =   "frm_calc.frx":0777
      Left            =   4560
      List            =   "frm_calc.frx":0779
      TabIndex        =   2
      Text            =   "Select"
      Top             =   1560
      Width           =   2310
   End
   Begin VB.ComboBox Ctrmplan 
      Height          =   360
      ItemData        =   "frm_calc.frx":077B
      Left            =   13320
      List            =   "frm_calc.frx":0788
      TabIndex        =   15
      Text            =   "Term Inssurance Plan"
      Top             =   720
      Width           =   2655
   End
   Begin VB.ComboBox Cendow 
      ForeColor       =   &H00000000&
      Height          =   360
      ItemData        =   "frm_calc.frx":07D0
      Left            =   2640
      List            =   "frm_calc.frx":07E0
      TabIndex        =   13
      Text            =   "Endowment Plan"
      Top             =   720
      Width           =   2175
   End
   Begin VB.ComboBox Cmnybck 
      Height          =   360
      ItemData        =   "frm_calc.frx":083C
      Left            =   5400
      List            =   "frm_calc.frx":084C
      TabIndex        =   11
      Text            =   "Money Back Plan"
      Top             =   765
      Width           =   2310
   End
   Begin VB.ComboBox Cpension 
      Height          =   360
      ItemData        =   "frm_calc.frx":08B5
      Left            =   8280
      List            =   "frm_calc.frx":08C5
      TabIndex        =   10
      Text            =   "Penssion Plan"
      Top             =   720
      Width           =   2055
   End
   Begin VB.ComboBox Cchildplan 
      Height          =   360
      ItemData        =   "frm_calc.frx":092A
      Left            =   120
      List            =   "frm_calc.frx":093A
      TabIndex        =   1
      Text            =   "Children Plan"
      Top             =   720
      Width           =   2055
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Occupation Extra"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   32
      Top             =   6240
      Width           =   2295
   End
   Begin VB.Image Image1 
      Height          =   6885
      Left            =   11400
      Picture         =   "frm_calc.frx":0995
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   5325
   End
   Begin VB.Label Label13 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Children  Plan"
      Height          =   270
      Left            =   120
      TabIndex        =   27
      Top             =   360
      Width           =   2310
   End
   Begin VB.Label Label12 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Money Back Plan"
      Height          =   270
      Left            =   5400
      TabIndex        =   26
      Top             =   360
      Width           =   2310
   End
   Begin VB.Label Label11 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Term Inssurance Plan"
      Height          =   270
      Left            =   13320
      TabIndex        =   25
      Top             =   360
      Width           =   2820
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Special Plan"
      Height          =   270
      Left            =   10800
      TabIndex        =   24
      Top             =   360
      Width           =   1800
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Penssion Plan"
      Height          =   270
      Left            =   8280
      TabIndex        =   23
      Top             =   360
      Width           =   2055
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Endowment Plan"
      Height          =   270
      Left            =   2640
      TabIndex        =   22
      Top             =   360
      Width           =   1800
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Health Extra"
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
      Index           =   0
      Left            =   1320
      TabIndex        =   16
      Top             =   5280
      Width           =   2415
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFC0C0&
      Caption         =   "PPT"
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
      Left            =   1320
      TabIndex        =   14
      Top             =   4320
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Sum Assured"
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
      Left            =   1320
      TabIndex        =   12
      Top             =   3240
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Age"
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
      Left            =   1320
      TabIndex        =   9
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Mode"
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
      Index           =   0
      Left            =   1320
      TabIndex        =   0
      Top             =   2400
      Width           =   2055
   End
End
Attribute VB_Name = "frm_calc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset
Dim con As ADODB.Connection
Dim Premium As Double

Dim cp1 As String
Dim cp2 As String
Dim cp3 As String
Dim cp4 As String
Dim ep1 As String
Dim ep2 As String
Dim ep3 As String
Dim ep4 As String
Dim mp1 As String
Dim mp2 As String
Dim mp3 As String
Dim mp4 As String
Dim tp1 As String
Dim tp2 As String
Dim tp3 As String
Dim sp1 As String
Dim sp2 As String
Dim pp1 As String
Dim pp2 As String
Dim pp3 As String
Dim pp4 As String
 

Private Sub cage_click()
Set con = New ADODB.Connection
con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\lic.mdb"
con.Open
Set rs = New ADODB.Recordset
If Cchildplan.Text = cp1 Or Cchildplan.Text = cp2 Or Cchildplan.Text = cp3 Or Cchildplan.Text = cp4 Then
 rs.Open "select * from Children_Premium where age=" & cage.Text, con, adOpenKeyset, adLockOptimistic
 Premium = rs.Fields(1)
rs.Close
End If

If Cendow.Text = ep1 Or Cendow.Text = ep2 Or Cendow.Text = ep3 Or Cendow.Text = ep4 Then
rs.Open "select * from Endowment_premium where age=" & cage.Text, con, adOpenKeyset, adLockOptimistic
Premium = rs.Fields(1)
rs.Close
End If

If Cmnybck.Text = mp1 Or Cmnybck.Text = mp2 Or Cmnybck.Text = mp3 Or Cmnybck.Text = mp4 Then
 rs.Open "select * from Money_back where age=" & cage.Text, con, adOpenKeyset, adLockOptimistic
Premium = rs.Fields(1)
rs.Close
End If

If Cpension.Text = pp1 Or Cpension.Text = pp2 Or Cpension.Text = pp3 Or Cpension.Text = pp4 Then
rs.Open "select * from Pension_Premium where age=" & cage.Text, con, adOpenKeyset, adLockOptimistic
Premium = rs.Fields(1)
rs.Close
End If

If Cspecial.Text = sp1 Or Cspecial.Text = sp2 Then
rs.Open "select * from Special_premium where age=" & cage.Text, con, adOpenKeyset, adLockOptimistic
Premium = rs.Fields(1)
rs.Close
End If

If Ctrmplan = tp1 Or Ctrmplan = tp2 Or Ctrmplan = tp3 Then
rs.Open "select * from Term_premium where age=" & cage.Text, con, adOpenKeyset, adLockOptimistic
Premium = rs.Fields(1)
rs.Close
End If
con.Close

MsgBox Premium
End Sub

Private Sub Cchildplan_click()
cage.Clear
Set con = New ADODB.Connection
con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\lic.mdb"
Dim s As String
s = "select * from Children_Premium "
con.Open
Set rs = New ADODB.Recordset
rs.Open s, con
cage.Text = rs.Fields(0)
While Not rs.EOF
cage.AddItem (rs.Fields(0))
rs.MoveNext
Wend
If cage.Text = Selected Then
Txtprm.Text = rs.Fields(1)
End If
con.Close
Check1.Visible = True
Check2.Visible = True
Check3.Visible = True
Check4.Visible = True
End Sub



Private Sub Cendow_Click()
cage.Clear
Set con = New ADODB.Connection
con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\lic.mdb"
Dim s As String
s = "select * from Endowment_premium"

con.Open
Set rs = New ADODB.Recordset
rs.Open s, con
cage.Text = rs.Fields(0)
While Not rs.EOF
cage.AddItem (rs.Fields(0))
rs.MoveNext
Wend


Check1.Visible = True
Check3.Visible = True
Check4.Visible = True

End Sub

Private Sub Check1_Click()
If Check1.Value = True Then
Premium = Premium + (Premium * 0.02)
End If
MsgBox Premium
End Sub

Private Sub Check2_Click()
If Check2.Value = True Then
Premium = Premium + (Premium * 0.05)
End If
MsgBox Premium

End Sub

Private Sub Check3_Click()
If Check3.Value = True Then
Premium = Premium + (Premium * 0.03)
End If
MsgBox Premium
End Sub

Private Sub Check4_Click()
If Check4.Value = True Then
Premium = Premium + (Premium * 0.04)
End If
MsgBox Premium
End Sub

Private Sub Cmnybck_Click()
cage.Clear
Set con = New ADODB.Connection
con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\lic.mdb"
Dim s As String
s = "select * from Money_back "

con.Open
Set rs = New ADODB.Recordset
rs.Open s, con
cage.Text = rs.Fields(0)
While Not rs.EOF
cage.AddItem (rs.Fields(0))
rs.MoveNext
Wend

Check1.Visible = True
Check3.Visible = True
Check4.Visible = True



End Sub
Private Sub Combo14_click()
If Combo14.Text = "category1" Then
Premium = Premium + (Premium * 0.01)
ElseIf Combo14.Text = "category2" Then
Premium = Premium + (Premium * 0.02)
ElseIf Combo14.Text = "category3" Then
Premium = Premium + (Premium * 0.035)
ElseIf Combo14.Text = "category4" Then
Premium = Premium + (Premium * 0.04)
ElseIf Combo14.Text = "category5" Then
Premium = Premium + (Premium * 0.06)
Else
End If

MsgBox Premium
End Sub
Private Sub Combo15_Click()
If Combo15.Text = "category1" Then
Premium = Premium + (Premium * 0.01)
ElseIf Combo15.Text = "category2" Then
Premium = Premium + (Premium * 0.02)
ElseIf Combo15.Text = "category3" Then
Premium = Premium + (Premium * 0.035)
ElseIf Combo15.Text = "category4" Then
Premium = Premium + (Premium * 0.04)
ElseIf Combo15.Text = "category5" Then
Premium = Premium + (Premium * 0.06)
Else
End If

MsgBox Premium
End Sub
Private Sub Combo7_Click()
If Combo7.Text = "SSS" Or Combo7.Text = "ECS" Or Combo7.Text = "Quaterly" Then
Premium = Premium + (Premium * 0.03)
End If
If Combo7.Text = "Half Yearly" Then
Premium = Premium + (Premium * 0.015)
MsgBox Premium
End If
End Sub
Private Sub Combo8_Click()
If Combo8.Text = "50000-100000" Then
Premium = (Premium * 0.2) + Premium
ElseIf Combo8.Text = "100000-300000" Then
Premium = (Premium * 0.3) + Premium
ElseIf Combo8.Text = "300000-5000000" Then
Premium = (Premium * 0.5) + pPremium
ElseIf Combo8.Text = "500000-any" Then
Premium = (Premium * 0.7) + Premium
Else
End If
MsgBox Premium
End Sub
Private Sub Combo9_Click()
If Val(Combo9.Text) > 13 Then
Premium = Premium + Premium * 0.05
ElseIf Val(Combo9.Text) > 18 Then
Premium = Premium + Premium * 0.05
ElseIf Val(Combo9.Text) > 25 Then
Premium = Premium + Premium * 0.07
Else
End If
MsgBox Premium
End Sub

Private Sub Command1_Click()
frm_tc.Show
End Sub

Private Sub Command2_Click()
Check1.Visible = False
Check2.Visible = False
Check3.Visible = False
Check4.Visible = False
End Sub

Private Sub Command3_Click()
finalprem.Text = Premium
Premium = 0
End Sub

Private Sub Cpension_Click()
cage.Clear
Set con = New ADODB.Connection
con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\lic.mdb"
Dim s As String
s = "select * from Pension_Premium"
con.Open
Set rs = New ADODB.Recordset
rs.Open s, con
cage.Text = rs.Fields(0)
While Not rs.EOF
cage.AddItem (rs.Fields(0))
rs.MoveNext
Wend


Check1.Visible = True


End Sub
Private Sub Cspecial_Click()
cage.Clear
Set con = New ADODB.Connection
con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\lic.mdb"
Dim s As String
s = "select * from Special_premium"
con.Open
Set rs = New ADODB.Recordset
rs.Open s, con
cage.Text = rs.Fields(0)
While Not rs.EOF
cage.AddItem (rs.Fields(0))
rs.MoveNext
Wend

Check1.Visible = True
Check3.Visible = True
Check4.Visible = True


End Sub
Private Sub Ctrmplan_Click()
cage.Clear
Set con = New ADODB.Connection
con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\lic.mdb"
Dim s As String
s = "select * from Term_premium"
con.Open
Set rs = New ADODB.Recordset
rs.Open s, con
cage.Text = rs.Fields(0)
While Not rs.EOF
cage.AddItem (rs.Fields(0))
rs.MoveNext
Wend

Check1.Visible = True

End Sub

Private Sub Final_Click()

End Sub

Private Sub Form_Load()

Set con = New ADODB.Connection
con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\lic.mdb"
'con.Open
Set rs = New ADODB.Recordset
For i = 11 To 35
Combo9.AddItem (i)
Next
cp1 = "Komal Jeevan(159)"
cp2 = "Jeevan Kishor(102)"
cp3 = "Child Career Plan(184)"
cp4 = "Child Future Plan(185)"
ep1 = "Endowment with profit(14)"
ep2 = "Jeevan Chhaya(103)"
ep3 = "Jeevan Pramukh(167)"
ep4 = "Jeevan Anurag(168)"
mp1 = "Bima Bachat (175)"
mp2 = "Money Back(20 years)(7)"
mp3 = "Money Back (25 years)(93)"
mp4 = "Jeevan Surbhi(20 Years)(108)"
sp1 = "Jeevan Anand"
sp2 = "Jeevan Saral"
pp1 = "New Jeevan Suraksha - 1(147)"
pp2 = "New Jeevan Akshay - 1(148)"
pp3 = "Jeevan Nidhi(169)"
tp1 = "Temporary Assurance(43)"
tp2 = "Anmol Jeevan-1(164)"
tp3 = "Amulya Jeevan-1(190)"

End Sub


