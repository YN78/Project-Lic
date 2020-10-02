VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form2 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Policy Customer"
   ClientHeight    =   9855
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14415
   LinkTopic       =   "Form2"
   ScaleHeight     =   9855
   ScaleWidth      =   14415
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   840
      TabIndex        =   1
      Top             =   1440
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   495
      Left            =   4080
      TabIndex        =   2
      Top             =   1440
      Width           =   4215
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   8760
      Locked          =   -1  'True
      TabIndex        =   4
      Text            =   "Select"
      Top             =   1320
      Width           =   3015
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   840
      TabIndex        =   0
      Top             =   2280
      Width           =   2805
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   6360
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   4320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   735
      Left            =   8280
      Top             =   4635
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1296
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
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3975
      Left            =   120
      TabIndex        =   5
      Top             =   3000
      Width           =   13335
      _ExtentX        =   23521
      _ExtentY        =   7011
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   22
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Image Image2 
      Height          =   2535
      Left            =   720
      Picture         =   "frm_search2.frx":0000
      Stretch         =   -1  'True
      Top             =   7320
      Width           =   12015
   End
   Begin VB.Image Image1 
      Height          =   1095
      Left            =   0
      Picture         =   "frm_search2.frx":4234
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2655
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset

Private Sub Combo1_Click()
Adodc1.RecordSource = "select * from policy_customer where Customer_Name  like '" & Combo1.Text & "'"
Adodc1.Refresh
Set DataGrid1.DataSource = Adodc1

End Sub

Private Sub Command1_Click()
Adodc1.RecordSource = "select * from policy_customer where Customer_Name  like '" & Text1.Text & "'"
Adodc1.Refresh
Set DataGrid1.DataSource = Adodc1
End Sub



Private Sub Command3_Click()
Dim id As Integer
Dim rs1 As New ADODB.Recordset
id = Val(InputBox("Enter id"))
rs.Open "delete * from policy_customer where  Customer_Number=" & id
'rs.Delete
'rs.Update
Adodc1.Refresh
DataGrid1.Refresh
'rs.Close
'Combo1.Refresh
rs1.Open "select * from policy_customer", con, 2, 2
'Combo1.text = rs.Fields(0)
'txtcustname.Text = rs.Fields(1)
Combo1.Clear
While Not rs1.EOF
Combo1.AddItem (rs1.Fields(1))
rs1.MoveNext
Wend
End Sub

Private Sub Form_Load()
Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\lic.mdb"
Adodc1.RecordSource = "select Customer_Name from policy_customer"
Set Text2.DataSource = Adodc1
While Adodc1.Recordset.EOF <> True
Combo1.AddItem (Adodc1.Recordset.Fields(0))
Adodc1.Recordset.MoveNext
Wend

Set con = New ADODB.Connection
con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\lic.mdb"
con.Open

Set rs = New ADODB.Recordset
rs.Open "select * from policy_customer", con, adOpenKeyset, adLockOptimistic
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


