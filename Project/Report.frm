VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frm_alert1 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Report"
   ClientHeight    =   10050
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   16545
   LinkTopic       =   "Form1"
   ScaleHeight     =   10050
   ScaleWidth      =   16545
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      BackColor       =   &H00FF8080&
      Caption         =   "FOLLOW UP ALERT"
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   3135
      Left            =   9240
      TabIndex        =   2
      Top             =   480
      Width           =   7095
      Begin VB.CommandButton txtshow 
         Caption         =   "Show Alert"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2040
         TabIndex        =   7
         Top             =   2280
         Width           =   1935
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   4560
         TabIndex        =   4
         Top             =   600
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   393216
         Format          =   103088129
         CurrentDate     =   40811
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FF8080&
         Caption         =   "Up to Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   6
         Top             =   600
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      Caption         =   "ALERT FOR POLICY CUSTOMER"
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   840
      TabIndex        =   1
      Top             =   360
      Width           =   7095
      Begin VB.CommandButton txtshow1 
         Caption         =   "Show Alert"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1920
         TabIndex        =   8
         Top             =   2400
         Width           =   1935
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   3600
         TabIndex        =   3
         Top             =   600
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   393216
         CalendarForeColor=   16711680
         Format          =   103088129
         CurrentDate     =   40811
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FF8080&
         Caption         =   "Up To Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         TabIndex        =   5
         Top             =   600
         Width           =   1935
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   735
      Left            =   240
      Top             =   4080
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
      Bindings        =   "Report.frx":0000
      Height          =   4935
      Left            =   720
      TabIndex        =   0
      Top             =   3960
      Width           =   12855
      _ExtentX        =   22675
      _ExtentY        =   8705
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
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
   Begin VB.Image Image1 
      Height          =   615
      Left            =   11880
      Picture         =   "Report.frx":0015
      Stretch         =   -1  'True
      Top             =   9240
      Width           =   2055
   End
End
Attribute VB_Name = "frm_alert1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset




Private Sub DTPicker1_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)
DTPicker1.MinDate = Now
End Sub

Private Sub DTPicker2_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)
DTPicker1.MinDate = Now
End Sub

Private Sub Frame1_DragDrop(Source As Control, x As Single, Y As Single)

End Sub

Private Sub txtshow1_Click()
Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\lic.mdb"
Adodc1.RecordSource = "select * from policy_customer where Maturity_Date>= " & Format(DateValue(DTPicker1.Value), "yyyy-mm-dd") & " AND Maturity_Date>=" & Format(Now, "yyyy-mm-dd")
Set DataGrid1.DataSource = Adodc1
End Sub




Private Sub txtshow_Click()
Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\lic.mdb"
Adodc1.RecordSource = "select * from customer where Alert_Days>= " & Format(DateValue(DTPicker2.Value), "yyyy-mm-dd") & " AND Alert_days>=" & Format(Now, "yyyyf-mm-dd")
Set DataGrid1.DataSource = Adodc1
End Sub


