VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_calc 
   Caption         =   "Tabular Premium"
   ClientHeight    =   7425
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14730
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
   ScaleHeight     =   7425
   ScaleWidth      =   14730
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox finalpre 
      Height          =   1035
      Left            =   8160
      TabIndex        =   33
      Text            =   "Text1"
      Top             =   6120
      Width           =   3840
   End
   Begin VB.ComboBox Cspecial 
      Height          =   360
      ItemData        =   "Form1.frx":0000
      Left            =   9435
      List            =   "Form1.frx":000A
      TabIndex        =   24
      Text            =   "Special Plan"
      Top             =   765
      Width           =   2055
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   780
      Left            =   510
      Top             =   5610
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
   Begin VB.TextBox Txtprm 
      Height          =   1290
      Left            =   1785
      TabIndex        =   23
      Top             =   5865
      Width           =   1800
   End
   Begin VB.Frame Frame1 
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
      Left            =   8640
      TabIndex        =   18
      Top             =   1800
      Width           =   4575
      Begin VB.CheckBox Check4 
         Caption         =   "TERM RIDER"
         Height          =   375
         Left            =   2520
         TabIndex        =   22
         Top             =   1320
         Width           =   1935
      End
      Begin VB.CheckBox Check3 
         Caption         =   "CIR"
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   1320
         Width           =   1935
      End
      Begin VB.CheckBox Check2 
         Caption         =   "PWB"
         Height          =   255
         Left            =   2520
         TabIndex        =   20
         Top             =   480
         Width           =   1695
      End
      Begin VB.CheckBox Check1 
         Caption         =   "DAB"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   480
         Width           =   1455
      End
   End
   Begin VB.ComboBox Combo15 
      Height          =   360
      ItemData        =   "Form1.frx":002A
      Left            =   12240
      List            =   "Form1.frx":003D
      TabIndex        =   17
      Text            =   "Occupation"
      Top             =   4440
      Width           =   2175
   End
   Begin VB.ComboBox Combo14 
      Height          =   360
      ItemData        =   "Form1.frx":0078
      Left            =   9600
      List            =   "Form1.frx":008B
      TabIndex        =   16
      Text            =   "Health"
      Top             =   4440
      Width           =   2175
   End
   Begin VB.ComboBox Combo9 
      Height          =   360
      Left            =   4335
      TabIndex        =   14
      Text            =   "Select"
      Top             =   5100
      Width           =   2310
   End
   Begin VB.ComboBox Combo8 
      Height          =   360
      Left            =   4335
      TabIndex        =   13
      Text            =   "Select"
      Top             =   4080
      Width           =   2310
   End
   Begin VB.ComboBox Combo7 
      Height          =   360
      ItemData        =   "Form1.frx":00C6
      Left            =   4335
      List            =   "Form1.frx":00D9
      TabIndex        =   12
      Text            =   "Select"
      Top             =   3060
      Width           =   2310
   End
   Begin VB.ComboBox cage 
      Height          =   360
      ItemData        =   "Form1.frx":0106
      Left            =   4335
      List            =   "Form1.frx":0108
      TabIndex        =   11
      Text            =   "Select"
      Top             =   2040
      Width           =   2310
   End
   Begin VB.ComboBox Ctrmplan 
      Height          =   360
      ItemData        =   "Form1.frx":010A
      Left            =   11730
      List            =   "Form1.frx":0117
      TabIndex        =   10
      Text            =   "Term Inssurance Plan"
      Top             =   765
      Width           =   2655
   End
   Begin VB.ComboBox Cendow 
      ForeColor       =   &H00000000&
      Height          =   360
      ItemData        =   "Form1.frx":015F
      Left            =   2295
      List            =   "Form1.frx":016F
      TabIndex        =   8
      Text            =   "Endowment Plan"
      Top             =   765
      Width           =   2175
   End
   Begin VB.ComboBox Cmnybck 
      Height          =   360
      ItemData        =   "Form1.frx":01CB
      Left            =   4590
      List            =   "Form1.frx":01DB
      TabIndex        =   6
      Text            =   "Money Back Plan"
      Top             =   765
      Width           =   2310
   End
   Begin VB.ComboBox Cpension 
      Height          =   360
      ItemData        =   "Form1.frx":0244
      Left            =   7140
      List            =   "Form1.frx":0254
      TabIndex        =   5
      Text            =   "Penssion Plan"
      Top             =   765
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Next"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   12120
      TabIndex        =   4
      Top             =   7320
      Width           =   2295
   End
   Begin VB.ComboBox Cchildplan 
      Height          =   360
      ItemData        =   "Form1.frx":02B9
      Left            =   0
      List            =   "Form1.frx":02C9
      TabIndex        =   0
      Text            =   "Children Plan"
      Top             =   765
      Width           =   2055
   End
   Begin VB.Label Label13 
      Caption         =   "Children  Plan"
      Height          =   270
      Left            =   0
      TabIndex        =   32
      Top             =   255
      Width           =   2310
   End
   Begin VB.Label Label12 
      Caption         =   "Money Back Plan"
      Height          =   270
      Left            =   4590
      TabIndex        =   31
      Top             =   255
      Width           =   2310
   End
   Begin VB.Label Label11 
      Caption         =   "Term Inssurance Plan"
      Height          =   270
      Left            =   11730
      TabIndex        =   30
      Top             =   255
      Width           =   2820
   End
   Begin VB.Label Label10 
      Caption         =   "Label10"
      Height          =   15
      Left            =   12240
      TabIndex        =   29
      Top             =   510
      Width           =   1800
   End
   Begin VB.Label Label9 
      Caption         =   "Special Plan"
      Height          =   270
      Left            =   9435
      TabIndex        =   28
      Top             =   255
      Width           =   1800
   End
   Begin VB.Label Label8 
      Caption         =   "Penssion Plan"
      Height          =   270
      Left            =   7140
      TabIndex        =   27
      Top             =   255
      Width           =   2055
   End
   Begin VB.Label Label7 
      Caption         =   "Label7"
      Height          =   15
      Left            =   4845
      TabIndex        =   26
      Top             =   510
      Width           =   1800
   End
   Begin VB.Label Label5 
      Caption         =   "Endowment Plan"
      Height          =   270
      Left            =   2295
      TabIndex        =   25
      Top             =   255
      Width           =   1800
   End
   Begin VB.Label Label6 
      Caption         =   "Extra"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7560
      TabIndex        =   15
      Top             =   4440
      Width           =   1695
   End
   Begin VB.Label Label3 
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
      Left            =   1530
      TabIndex        =   9
      Top             =   5100
      Width           =   2055
   End
   Begin VB.Label Label2 
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
      Left            =   1530
      TabIndex        =   7
      Top             =   4080
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Premium"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   3
      Left            =   5355
      TabIndex        =   3
      Top             =   6375
      Width           =   3015
   End
   Begin VB.Label Label1 
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
      Left            =   1530
      TabIndex        =   2
      Top             =   2040
      Width           =   2175
   End
   Begin VB.Label Label1 
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
      Left            =   1530
      TabIndex        =   1
      Top             =   3060
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


Private Sub cage_click()
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



Set con = New ADODB.Connection
con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\lic.mdb"
con.Open
Set rs = New ADODB.Recordset
If Cchildplan.Text = cp1 Or Cchildplan.Text = cp2 Or Cchildplan.Text = cp3 Or Cchildplan.Text = cp4 Then
 rs.Open "select * from Children_Premium where age=" & cage.Text, con, adOpenKeyset, adLockOptimistic
Txtprm = rs.Fields(1)
rs.Close
End If

If Cendow.Text = ep1 Or Cendow.Text = ep2 Or Cendow.Text = ep3 Or Cendow.Text = ep4 Then
rs.Open "select * from Endowment_premium where age=" & cage.Text, con, adOpenKeyset, adLockOptimistic
Txtprm = rs.Fields(1)
rs.Close
End If

If Cmnybck.Text = mp1 Or Cmnybck.Text = mp2 Or Cmnybck.Text = mp3 Or Cmnybck.Text = mp4 Then
 rs.Open "select * from Money_back where age=" & cage.Text, con, adOpenKeyset, adLockOptimistic
Txtprm = rs.Fields(1)
rs.Close
End If

If Cpension.Text = pp1 Or Cpension.Text = pp2 Or Cpension.Text = pp3 Or Cpension.Text = pp4 Then
rs.Open "select * from Pension_Premium where age=" & cage.Text, con, adOpenKeyset, adLockOptimistic
Txtprm = rs.Fields(1)
rs.Close
End If

If Cspecial.Text = sp1 Or Cspecial.Text = sp2 Then
rs.Open "select * from Special_premium where age=" & cage.Text, con, adOpenKeyset, adLockOptimistic
Txtprm = rs.Fields(1)
rs.Close
End If

If Ctrmplan = tp1 Or Ctrmplan = tp2 Or Ctrmplan = tp3 Then
rs.Open "select * from Term_premium where age=" & cage.Text, con, adOpenKeyset, adLockOptimistic
Txtprm = rs.Fields(1)
rs.Close
End If


con.Close


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
End Sub




Private Sub Combo14_Change()

If Combo14.Text = "catogory1" Then Premium = Premium + (Premium * 0.01)
ElseIf Combo14.Text = "catogory2" Then Premium = Premium + (Premium * 0.02)
ElseIf Combo14.Text = "catogory3" Then Premium = Premium + (Premium * 0.035)
ElseIf Combo14.Text = "catogory4" Then Premium = Premium + (Premium * 0.04)
ElseIf Combo14.Text = "catogory5" Then Premium = Premium + (Premium * 0.06)
End If
End Sub

Private Sub Combo15_Change()

If Combo15.Text = "catogory1" Then Premium = Premium + (Premium * 0.01)
ElseIf Combo15.Text = "catogory2" Then Premium = Premium + (Premium * 0.02)
ElseIf Combo15.Text = "catogory3" Then Premium = Premium + (Premium * 0.035)
ElseIf Combo15.Text = "catogory4" Then Premium = Premium + (Premium * 0.04)
ElseIf Combo15.Text = "catogory5" Then Premium = Premium + (Premium * 0.06)
End If
End Sub

Private Sub Combo7_Change()
If Combo7.Text = "SSS" Or "ECS" Or "Quaterly" Then
 Premium = Premium + (Premium * 0.03)
End If
If Combo7.Text = "Half Yearly" Then
Premium = Premium + (Premium * 0.015)
End If
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
End Sub

Private Sub Form_Load()
Set con = New ADODB.Connection
con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\lic.mdb"
'con.Open
Set rs = New ADODB.Recordset
End Sub

