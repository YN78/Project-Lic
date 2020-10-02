VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_start 
   BackColor       =   &H00FFC0C0&
   Caption         =   "LIC "
   ClientHeight    =   9255
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   14340
   LinkTopic       =   "Form2"
   ScaleHeight     =   9255
   ScaleWidth      =   14340
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   855
      Left            =   240
      Top             =   7920
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1508
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
      Enabled         =   0
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
   Begin VB.Image Image1 
      Height          =   9255
      Left            =   1080
      Picture         =   "menu.frx":0000
      Stretch         =   -1  'True
      Top             =   960
      Width           =   17175
   End
   Begin VB.Menu Presentation 
      Caption         =   "  Plan Presentation"
   End
   Begin VB.Menu policy_info 
      Caption         =   "  Policy Information"
      Begin VB.Menu tc 
         Caption         =   "Terms And Condition"
      End
      Begin VB.Menu calc 
         Caption         =   "Premium Calculator"
      End
   End
   Begin VB.Menu Customer1 
      Caption         =   "  Customer"
      Begin VB.Menu Prospactor 
         Caption         =   "Prospactor"
      End
      Begin VB.Menu policy_customer 
         Caption         =   "Policy Customer"
      End
   End
   Begin VB.Menu Search 
      Caption         =   "  Search"
      Begin VB.Menu pcsearch 
         Caption         =   "Policy Customer Search"
      End
      Begin VB.Menu custsearch 
         Caption         =   "Customer Search"
      End
   End
   Begin VB.Menu commission 
      Caption         =   "  Commission"
   End
   Begin VB.Menu alert 
      Caption         =   "  Alert"
   End
   Begin VB.Menu Reports 
      Caption         =   "  Reports"
      Begin VB.Menu policycustomer 
         Caption         =   "policy_customer"
      End
      Begin VB.Menu cust_report 
         Caption         =   "Customer"
      End
      Begin VB.Menu comission 
         Caption         =   "Commission"
      End
      Begin VB.Menu bday 
         Caption         =   "Birthday's"
      End
   End
   Begin VB.Menu about 
      Caption         =   "  About"
   End
   Begin VB.Menu logout 
      Caption         =   "  Logout"
   End
End
Attribute VB_Name = "frm_start"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub about_Click()
frm_about.Show
End Sub

Private Sub alert_Click()
frm_alert1.Show
End Sub

Private Sub bday_Click()
DataReport4.Show
End Sub

Private Sub calc_Click()
frm_calc.Show
End Sub

Private Sub comission_Click()
DataReport1.Show
End Sub

Private Sub commission_Click()
frm_commission.Show
End Sub

Private Sub cust_report_Click()
Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\lic.mdb"
Adodc1.RecordSource = "customer"
DataReport2.Show

End Sub

Private Sub custsearch_Click()
Form1.Show
End Sub

Private Sub exit_Click()
End
End Sub

Private Sub logout_Click()
'frm_login.Show
End
End Sub

Private Sub pcsearch_Click()
Form2.Show
End Sub

Private Sub policy_customer_Click()
frm_policycust.Show
End Sub

Private Sub policycustomer_Click()
DataReport3.Show
End Sub

Private Sub Presentation_Click()
frm_presentasion.Show
End Sub
Private Sub Prospactor_Click()
frm_customer.Show
End Sub



Private Sub tc_Click()
frm_tc.Show
End Sub
