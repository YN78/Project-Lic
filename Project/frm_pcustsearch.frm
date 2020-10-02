VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_pcustsearch 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Policy Customer Search"
   ClientHeight    =   9705
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15060
   LinkTopic       =   "Form1"
   ScaleHeight     =   9705
   ScaleWidth      =   15060
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   8280
      TabIndex        =   5
      Top             =   1320
      Width           =   2310
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   720
      TabIndex        =   4
      Top             =   2160
      Width           =   2535
   End
   Begin VB.CommandButton AddData 
      Caption         =   "Add Data"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   -360
      TabIndex        =   3
      Top             =   8040
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   3480
      TabIndex        =   2
      Top             =   1320
      Width           =   4335
   End
   Begin VB.CommandButton Search 
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   720
      TabIndex        =   1
      Top             =   1320
      Width           =   2535
   End
   Begin VB.PictureBox DataGrid1 
      Height          =   1935
      Left            =   510
      ScaleHeight     =   1875
      ScaleWidth      =   12915
      TabIndex        =   0
      Top             =   3315
      Width           =   12975
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   615
      Left            =   0
      Top             =   7650
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   1085
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\Project\lic.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\Project\lic.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from policy_customer"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Image Image2 
      Height          =   4380
      Left            =   600
      Picture         =   "frm_pcustsearch.frx":0000
      Stretch         =   -1  'True
      Top             =   6480
      Width           =   12240
   End
   Begin VB.Image Image1 
      Height          =   615
      Left            =   0
      Picture         =   "frm_pcustsearch.frx":4234
      Stretch         =   -1  'True
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "frm_pcustsearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset

Private Sub Command1_Click()
Adodc1.RecordSource = "select * from policy_customer where Customer_Name like'" & Text1.Text & "'"
Adodc1.Refresh
End Sub

Private Sub Command2_Click()

Adodc1.RecordSource = "select * from policy_customer"
Adodc1.Refresh
hEnd Sub


Private Sub Command3_Click()
Dim id As Integer
id = Val(InputBox("Enter id"))
rs2.Open "select * from policy_customer  where  PolicyCustomer_Number = " & id
rs2.Delete

End Sub


Private Sub Form_Load()
Set con = New ADODB.Connection
con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\lic.mdb"
con.Open

Set rs = New ADODB.Recordset
rs.Open "select * from policy_customer", con
Combo1.Text = rs.Fields(0)
While Not rs.EOF
Combo1.AddItem (rs.Fields(0))
rs.MoveNext
Wend
 rs.Close

End Sub


Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii <= 64 Or KeyAscii >= 122 Then
If KeyAscii <> 32 Then
If KeyAscii <> 8 Then
KeyAscii = 0
End If
End If
End If
End Sub

Private Sub Text1_LostFocus()
Dim s As Integer

s = Len(Txtcname.Text)
If s = 0 Then
MsgBox "Enter the data", vbQuestion
End If
End Sub
