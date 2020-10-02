VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_policycust 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Policy Customer"
   ClientHeight    =   10110
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   18960
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   13.5
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FF0000&
   LinkTopic       =   "Form2"
   ScaleHeight     =   10110
   ScaleWidth      =   18960
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox DTPicker1 
      Height          =   495
      Left            =   3360
      ScaleHeight     =   435
      ScaleWidth      =   2835
      TabIndex        =   16
      Top             =   7680
      Width           =   2895
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   840
      Top             =   9120
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
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
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.ComboBox Combo4 
      ForeColor       =   &H00000000&
      Height          =   480
      ItemData        =   "frm_policycus.frmt.frx":0000
      Left            =   3315
      List            =   "frm_policycus.frmt.frx":0010
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   6600
      Width           =   2910
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   645
      Left            =   11040
      TabIndex        =   1
      Top             =   135
      Width           =   2415
   End
   Begin VB.TextBox txtplan_name 
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   3315
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   3720
      Width           =   2895
   End
   Begin VB.ComboBox Combo3 
      Height          =   480
      Left            =   3315
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   4680
      Width           =   2925
   End
   Begin VB.ComboBox Combo2 
      ForeColor       =   &H00000000&
      Height          =   480
      ItemData        =   "frm_policycus.frmt.frx":003B
      Left            =   3315
      List            =   "frm_policycus.frmt.frx":007E
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   2760
      Width           =   2940
   End
   Begin VB.PictureBox DTPicker2 
      Height          =   495
      Left            =   3315
      ScaleHeight     =   435
      ScaleWidth      =   2835
      TabIndex        =   7
      Top             =   5640
      Width           =   2895
   End
   Begin VB.TextBox txtcustname 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   525
      Left            =   3360
      TabIndex        =   3
      Top             =   1800
      Width           =   2775
   End
   Begin VB.TextBox txtpcnumber 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   540
      Left            =   3360
      TabIndex        =   5
      Top             =   240
      Width           =   2895
   End
   Begin VB.CommandButton add_new 
      Caption         =   "Add New Record"
      Height          =   735
      Left            =   4680
      Picture         =   "frm_policycus.frmt.frx":00E6
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Add New Record"
      Top             =   9120
      Width           =   2535
   End
   Begin VB.CommandButton save 
      Caption         =   "Save Record"
      Height          =   735
      Left            =   7560
      Picture         =   "frm_policycus.frmt.frx":0230
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Save Record"
      Top             =   9120
      Width           =   2415
   End
   Begin VB.ComboBox Combo1 
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   3360
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   1080
      Width           =   2895
   End
   Begin VB.TextBox txtaddress 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1395
      Left            =   11040
      TabIndex        =   12
      Top             =   6480
      Width           =   3735
   End
   Begin VB.TextBox txtcontact_no 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   735
      Left            =   11040
      TabIndex        =   11
      Top             =   5400
      Width           =   2415
   End
   Begin VB.TextBox txtnominee_name 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   735
      Left            =   11040
      TabIndex        =   10
      Top             =   4200
      Width           =   2415
   End
   Begin VB.TextBox txtbenifit 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   735
      Left            =   11040
      TabIndex        =   9
      Top             =   3000
      Width           =   2415
   End
   Begin VB.TextBox txtamt 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   735
      Left            =   11040
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   1920
      Width           =   2415
   End
   Begin VB.TextBox txtpremium 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   735
      Left            =   11040
      TabIndex        =   15
      Top             =   960
      Width           =   2415
   End
   Begin VB.Image Image1 
      Height          =   8535
      Left            =   15000
      Picture         =   "frm_policycus.frmt.frx":04A9
      Stretch         =   -1  'True
      Top             =   360
      Width           =   3780
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFC0C0&
      Caption         =   "DOB"
      Height          =   495
      Left            =   360
      TabIndex        =   33
      Top             =   7800
      Width           =   2295
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Mode"
      Height          =   525
      Left            =   360
      TabIndex        =   32
      Top             =   6600
      Width           =   2565
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Inssured Person"
      Height          =   735
      Left            =   360
      TabIndex        =   31
      Top             =   1800
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Customer Number"
      Height          =   495
      Index           =   13
      Left            =   360
      TabIndex        =   30
      Top             =   1080
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Plan Number"
      Height          =   495
      Index           =   12
      Left            =   360
      TabIndex        =   29
      Top             =   2760
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Plan Name"
      Height          =   495
      Index           =   11
      Left            =   360
      TabIndex        =   28
      Top             =   3720
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Term"
      Height          =   495
      Index           =   10
      Left            =   360
      TabIndex        =   27
      Top             =   4680
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Premium"
      ForeColor       =   &H00000000&
      Height          =   495
      Index           =   9
      Left            =   7440
      TabIndex        =   26
      Top             =   1080
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Maturity Date"
      ForeColor       =   &H00000000&
      Height          =   495
      Index           =   8
      Left            =   7440
      TabIndex        =   25
      Top             =   240
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Maturity Amount"
      ForeColor       =   &H00000000&
      Height          =   495
      Index           =   7
      Left            =   7440
      TabIndex        =   24
      Top             =   2160
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Nominee Name"
      ForeColor       =   &H00000000&
      Height          =   495
      Index           =   6
      Left            =   7560
      TabIndex        =   23
      Top             =   4200
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Nominee Contact NO"
      ForeColor       =   &H00000000&
      Height          =   495
      Index           =   5
      Left            =   7440
      TabIndex        =   22
      Top             =   5400
      Width           =   3135
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Customer Address"
      ForeColor       =   &H00000000&
      Height          =   495
      Index           =   4
      Left            =   7440
      TabIndex        =   21
      Top             =   6480
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Survival Benifit"
      ForeColor       =   &H00000000&
      Height          =   495
      Index           =   3
      Left            =   7440
      TabIndex        =   20
      Top             =   3240
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "DOC"
      Height          =   495
      Index           =   1
      Left            =   360
      TabIndex        =   19
      Top             =   5640
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "P.C Number"
      Height          =   495
      Index           =   0
      Left            =   480
      TabIndex        =   4
      Top             =   240
      Width           =   2655
   End
End
Attribute VB_Name = "frm_policycust"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset


Private Sub add_new_Click()
txtpcnumber = ""
txtcustname = ""
txtplan_name = ""
txtpremium = ""
txtamt = ""
txtbenifit = ""
txtnominee_name = ""
txtcontact_no = ""
txtaddress = ""
End Sub

Private Sub Combo1_Click()
Set con = New ADODB.Connection
con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\lic.mdb"
con.Open
Set rs = New ADODB.Recordset
rs.Open "select * from Customer where Customer_Number=" & Val(Combo1.Text), con, 2, 2
'Combo1.Text = rs.Fields(0)
txtcustname.Text = rs.Fields(1)
'While Not rs.EOF
'Combo1.AddItem (rs.Fields(0))
'rs.MoveNext
'Wend
rs.Close

End Sub

Private Sub Combo2_click()
If (Combo2.Text = "2") Then txtplan_name.Text = "Endowment with profit"
If (Combo2.Text = "43") Then txtplan_name.Text = "Temporary Assurance"
If (Combo2.Text = "75") Then txtplan_name.Text = "Money Back(20 years)"
If (Combo2.Text = "93") Then txtplan_name.Text = "Money Back (25 years)"
If (Combo2.Text = "102") Then txtplan_name.Text = "Jeevan Kishor"
If (Combo2.Text = "103") Then txtplan_name.Text = "Jeevan Chhaya"
If (Combo2.Text = "107") Then txtplan_name.Text = "jeevan Surbhi(20 Years)"
If (Combo2.Text = "147") Then txtplan_name.Text = "New Jeevan Suraksha - 1"
If (Combo2.Text = "148") Then txtplan_name.Text = "New Jeevan Dhara -1"
If (Combo2.Text = "149") Then txtplan_name.Text = "Jeevan Anand"
If (Combo2.Text = "159") Then txtplan_name.Text = "Komal Jeevan"
If (Combo2.Text = "164") Then txtplan_name.Text = "Anmol Jeevan - 1"
If (Combo2.Text = "165") Then txtplan_name.Text = "Jeevan Saral"
If (Combo2.Text = "167") Then txtplan_name.Text = "Jeevan Pramukh"
If (Combo2.Text = "168") Then txtplan_name.Text = "Jeevan Anurag"
If (Combo2.Text = "169") Then txtplan_name.Text = "Jeevan Nidhi"
If (Combo2.Text = "175") Then txtplan_name.Text = "Bima Bachat"
If (Combo2.Text = "184") Then txtplan_name.Text = "Child Career Plan"
If (Combo2.Text = "185") Then txtplan_name.Text = "Child Future Plan"
If (Combo2.Text = "189") Then txtplan_name.Text = "Jeevan Akshay - v1"
If (Combo2.Text = "190") Then txtplan_name.Text = "Amulya Jeevan - 1"


End Sub

Private Sub DTPicker1_Change()
DTPicker1.MaxDate = Now
End Sub



Private Sub DTPicker2_Change()
a = Val(Combo3.List(Combo3.ListIndex))
Text1.Text = DateAdd("yyyy", a, Format(DTPicker2.Value, "mm/dd/yyyy"))
DTPicker2.MinDate = Now

End Sub

Private Sub DTPicker2_Click()
a = Val(Combo3.List(Combo3.ListIndex))
Text1.Text = DateAdd("yyyy", a, Format(DTPicker2.Value, "mm/dd/yyyy"))
DTPicker2.MinDate = Now
End Sub

Private Sub DTPicker2_KeyPress(KeyAscii As Integer)
a = Val(Combo3.List(Combo3.ListIndex))
Text1.Text = DateAdd("yyyy", a, Format(DTPicker2.Value, "mm/dd/yyyy"))
DTPicker2.MinDate = Now
End Sub

Private Sub Form_Load()
Set con = New ADODB.Connection
Set rs = New ADODB.Recordset
con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\lic.mdb"
con.Open
rs.Open "select * from Customer", con, 2, 2
'Combo1.text = rs.Fields(0)
'txtcustname.Text = rs.Fields(1)
While Not rs.EOF
Combo1.AddItem (rs.Fields(0))
rs.MoveNext
Wend
'rs.Close

Dim i As Integer
For i = 5 To 80
Combo3.AddItem (i)
Next
DTPicker1.Value = Now()
DTPicker2.Value = Now()
End Sub

Private Sub save_Click()

If Len(txtpcnumber.Text) = 0 Then
MsgBox "Enter Policy Number", vbInformation
txtpcnumber.SetFocus
Exit Sub
End If

If Len(Combo1.Text) = 0 Then
MsgBox ("Enter Customer Number"), vbInformation
Combo1.SetFocus
Exit Sub
End If

If Len(Combo2.Text) = 0 Then
MsgBox ("Enter Plan Number"), vbInformation
Combo2.SetFocus
Exit Sub
End If
If Len(Combo3.Text) = 0 Then
MsgBox ("Enter Term"), vbInformation
Combo3.SetFocus
Exit Sub
End If




If Len(Combo4.Text) = 0 Then
MsgBox ("Enter Mode"), vbInformation
Combo4.SetFocus
Exit Sub
End If

If Len(txtpremium.Text) = 0 Then
MsgBox ("Enter Premium"), vbInformation
txtpremium.SetFocus
Exit Sub
End If


If Len(txtbenifit.Text) = 0 Then
MsgBox ("Enter Benifit Yes/No"), vbInformation
txtbenifit.SetFocus
Exit Sub
End If

If Len(txtnominee_name.Text) = 0 Then
MsgBox ("Enter Nominee name"), vbInformation
txtnominee_name.SetFocus
Exit Sub
End If

If Len(txtcontact_no.Text) = 0 Then
MsgBox ("Enter Contact Number"), vbInformation
txtcontact_no.SetFocus
Exit Sub
End If

If Len(txtaddress.Text) = 0 Then
MsgBox ("Enter Address")
txtaddress.SetFocus
Exit Sub
End If

Set con = New ADODB.Connection
con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\lic.mdb"
con.Open
Set rs = New ADODB.Recordset
rs.Open "select * from policy_customer", con, 2, 2
rs.AddNew
rs.Fields("PolicyCustomer_Number") = Val(txtpcnumber.Text)
rs.Fields("Customer_Name") = txtcustname.Text
rs.Fields("Customer_Number") = Val(Combo1.Text)
rs.Fields("Policy_Number") = Val(txtpcnumber.Text)
rs.Fields("Policy_Name") = txtplan_name.Text
rs.Fields("Term") = Val(Combo3.Text)
rs.Fields("Maturity_Date") = Val(Text1.Text)
rs.Fields("Premium") = Val(txtpremium.Text)
rs.Fields("DOC") = DTPicker2.Value
rs.Fields("Mode") = Combo4.Text
rs.Fields("SurvivalBenifit") = txtbenifit.Text
rs.Fields("Nominee_Name") = txtnominee_name.Text
rs.Fields("Nominee_Contact_NO") = Val(txtcontact_no.Text)
rs.Fields("Customer_Add") = txtaddress.Text
rs.Fields("plan_no") = Val(Combo2.Text)
rs.Fields("dob") = DTPicker1.Value
rs.Update

MsgBox "Record Saved successfully", vbInformation

rs.Close
txtpcnumber = ""
txtcustname = ""
txtplan_name = ""
txtpremium = ""
txtamt = ""
txtbenifit = ""
txtnominee_name = ""
txtcontact_no = ""
txtaddress = ""
End Sub









'Private Sub txtaddress_GotFocus()
'Dim s As Integer
'
's = Len(txtaddress.Text)
'If s = 0 Then
'MsgBox "Enter the data", vbQuestion
'End If
'End Sub

Private Sub txtamt_GotFocus()
'Dim sa1 As Double
'Dim b1 As Integer
'Dim ma1 As Double
'Dim Bonus1 As Double
'Dim x As Double
'b1 = Val(InputBox("Enter Bonus Rate As per Plan"))
'
'
'
'sa1 = Val(txtpremium) * Val(Combo3.Text)
'x = sa1
'Bunus1 = Val((sa1 / 1000)) * b1 * Val(Combo3.Text)
'ma1 = Bonus1 + x
'txtamt.Text = ma1
End Sub




Private Sub txtbenifit_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Or KeyAscii = 27 Then
ElseIf KeyAscii = 13 Then
Txtbeniifit.SetFocus
ElseIf (KeyAscii >= 65 And KeyAscii <= 91) Or (KeyAscii >= 97 And KeyAscii <= 122) Then

Else
MsgBox "Enter Character Only", vbInformation
KeyAscii = 0
End If
End Sub





Private Sub txtcontact_no_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Or KeyAscii = 27 Then
ElseIf KeyAscii = 13 Then
Txtcname.SetFocus
ElseIf (KeyAscii >= 48 And KeyAscii <= 57) Then

Else
MsgBox "Enter Number Only", vbInformation
KeyAscii = 0
End If
End Sub



Private Sub txtcustname_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Or KeyAscii = 27 Then
ElseIf KeyAscii = 13 Then
txtcustname.SetFocus
ElseIf (KeyAscii >= 65 And KeyAscii <= 91) Or (KeyAscii >= 97 And KeyAscii <= 122) Then

Else
MsgBox "Enter Character Only", vbInformation
KeyAscii = 0
End If
End Sub
'
'Private Sub txtcustname_LostFocus()
'Dim s As Integer
'
's = Len(txtcustname.Text)
'If s = 0 Then
'MsgBox "Enter the data", vbQuestion
'End If
'End Sub

Private Sub txtnominee_name_KeyPress(KeyAscii As Integer)
If KeyAscii <= 64 Or KeyAscii >= 122 Then
If KeyAscii <> 32 Then
If KeyAscii <> 8 Then
KeyAscii = 0
End If
End If
End If
End Sub



'Private Sub txtnominee_name_LostFocus()
'Dim w As Integer
'
'w = Len(txtnominee_name.Text)
'If w = 0 Then
'MsgBox "Enter the data", vbQuestion
'End If
'End Sub

Private Sub txtpcnumber_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Or KeyAscii = 27 Then
ElseIf KeyAscii = 13 Then
txtpcnumber.SetFocus
ElseIf (KeyAscii >= 48 And KeyAscii <= 57) Then

Else
MsgBox "Enter Number Only", vbInformation
KeyAscii = 0
End If
End Sub

'Private Sub txtpcnumber_LostFocus()
'Dim x As Integer
'
'x = Len(txtpcnumber.Text)
'If x = 0 Then
'MsgBox "Enter the data", vbQuestion
'End If
'End Sub

'Private Sub txtplan_name_KeyPress(KeyAscii As Integer)
'If KeyAscii = 8 Or KeyAscii = 27 Then
'ElseIf KeyAscii = 13 Then
'txtplan_name.SetFocus
'ElseIf (KeyAscii >= 65 And KeyAscii <= 91) Or (KeyAscii >= 97 And KeyAscii <= 122) Then
'
'Else
'MsgBox "Enter Character Only"
'KeyAscii = 0
'End If
'End Sub



Private Sub txtpremium_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Or KeyAscii = 27 Then
ElseIf KeyAscii = 13 Then
txtpremium.SetFocus
ElseIf (KeyAscii >= 48 And KeyAscii <= 57) Then

Else
MsgBox "Enter Number Only"
KeyAscii = 0
End If
End Sub


Private Sub txtpremium_LostFocus()
Dim sa1 As Double
Dim b1 As Integer
Dim ma1 As Double
Dim Bonus1 As Double
Dim x As Double
b1 = Val(InputBox("Enter Bonus Rate As per Plan"))
sa1 = Val(txtpremium) * Val(Combo3.Text)
x = sa1
Bunus1 = Val((sa1 / 1000)) * b1 * Val(Combo3.Text)
ma1 = Bonus1 + x
txtamt.Text = ma1
End Sub
