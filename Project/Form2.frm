VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_policycust 
   BackColor       =   &H00FFFFFF&
   Caption         =   "POLICYCUSTOMER"
   ClientHeight    =   8430
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15555
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   13.5
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   8430
   ScaleWidth      =   15555
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtplan_name 
      Height          =   495
      Left            =   3315
      Locked          =   -1  'True
      TabIndex        =   30
      Top             =   3720
      Width           =   2895
   End
   Begin VB.ComboBox Combo3 
      Height          =   480
      Left            =   3315
      TabIndex        =   29
      Text            =   "Select"
      Top             =   4680
      Width           =   3045
   End
   Begin VB.ComboBox Combo2 
      Height          =   480
      ItemData        =   "Form2.frx":0000
      Left            =   3315
      List            =   "Form2.frx":0043
      TabIndex        =   28
      Text            =   "Select"
      Top             =   2760
      Width           =   2820
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   495
      Left            =   3240
      TabIndex        =   27
      Top             =   5640
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   95551489
      CurrentDate     =   40805
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   495
      Left            =   10920
      TabIndex        =   26
      Top             =   240
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   95551489
      CurrentDate     =   40805
   End
   Begin VB.TextBox txtcustname 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   3360
      TabIndex        =   25
      Top             =   1800
      Width           =   2895
   End
   Begin VB.TextBox txtpcnumber 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3360
      TabIndex        =   24
      Top             =   240
      Width           =   2775
   End
   Begin VB.CommandButton add_new 
      Caption         =   "Add New Record"
      Height          =   735
      Left            =   3000
      TabIndex        =   23
      Top             =   7320
      Width           =   2535
   End
   Begin VB.CommandButton save 
      Caption         =   "Save Record"
      Height          =   735
      Left            =   5880
      TabIndex        =   22
      Top             =   7320
      Width           =   2415
   End
   Begin VB.ComboBox Combo1 
      Height          =   480
      Left            =   3360
      TabIndex        =   21
      Text            =   "Select"
      Top             =   1080
      Width           =   2775
   End
   Begin VB.TextBox txtaddress 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   10680
      TabIndex        =   19
      Top             =   5760
      Width           =   3735
   End
   Begin VB.TextBox txtcontact_no 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10680
      TabIndex        =   18
      Top             =   4920
      Width           =   2535
   End
   Begin VB.TextBox txtnominee_name 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   10920
      TabIndex        =   17
      Top             =   3840
      Width           =   2055
   End
   Begin VB.TextBox txtbenifit 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10920
      TabIndex        =   16
      Top             =   2880
      Width           =   2055
   End
   Begin VB.TextBox txtamt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10920
      TabIndex        =   15
      Top             =   1920
      Width           =   2055
   End
   Begin VB.TextBox txtpremium 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10920
      TabIndex        =   14
      Top             =   960
      Width           =   2055
   End
   Begin VB.CommandButton next 
      Caption         =   "NEXT"
      Height          =   735
      Left            =   8640
      TabIndex        =   13
      Top             =   7320
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Customer name"
      Height          =   735
      Left            =   360
      TabIndex        =   20
      Top             =   1800
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Customer Number"
      Height          =   495
      Index           =   13
      Left            =   360
      TabIndex        =   12
      Top             =   1080
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Plan Number"
      Height          =   495
      Index           =   12
      Left            =   360
      TabIndex        =   11
      Top             =   2760
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Plan Name"
      Height          =   495
      Index           =   11
      Left            =   360
      TabIndex        =   10
      Top             =   3720
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Term"
      Height          =   495
      Index           =   10
      Left            =   360
      TabIndex        =   9
      Top             =   4680
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Premium"
      Height          =   495
      Index           =   9
      Left            =   7440
      TabIndex        =   8
      Top             =   1080
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Maturity Date"
      Height          =   495
      Index           =   8
      Left            =   7440
      TabIndex        =   7
      Top             =   240
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Maturity Amount"
      Height          =   495
      Index           =   7
      Left            =   7440
      TabIndex        =   6
      Top             =   2040
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Nominee Name"
      Height          =   495
      Index           =   6
      Left            =   7440
      TabIndex        =   5
      Top             =   3960
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Nominee Contact NO"
      Height          =   495
      Index           =   5
      Left            =   7440
      TabIndex        =   4
      Top             =   4920
      Width           =   3135
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Customer Address"
      Height          =   495
      Index           =   4
      Left            =   7440
      TabIndex        =   3
      Top             =   5880
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Survival Benifit"
      Height          =   495
      Index           =   3
      Left            =   7440
      TabIndex        =   2
      Top             =   2880
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "DOC"
      Height          =   495
      Index           =   1
      Left            =   480
      TabIndex        =   1
      Top             =   5640
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "P.C Number"
      Height          =   495
      Index           =   0
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   2655
   End
End
Attribute VB_Name = "frm_policycust"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


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

Private Sub Combo1_Change()
Set con = New ADODB.Connection
con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\lic.mdb"
con.Open
Set rs = New ADODB.Recordset
rs.Open "select * from Customer", con
Combo1.Text = rs.Fields(0)
txtcustname.Text = rs.Fields(1)
While Not rs.EOF
Combo1.AddItem (rs.Fields(0))
rs.MoveNext
Wend
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





Private Sub Form_Load()
Set con = New ADODB.Connection
con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\lic.mdb"
con.Open
Set rs = New ADODB.Recordset

rs.Open "select * from Customer", con
Combo1.Text = rs.Fields(0)
txtcustname.Text = rs.Fields(1)
While Not rs.EOF
Combo1.AddItem (rs.Fields(0))
rs.MoveNext
Wend
rs.Close

Dim i As Integer
For i = 5 To 80
Combo3.AddItem (i)
Next
End Sub

Private Sub save_Click()

Set con = New ADODB.Connection
con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\lic.mdb"
con.Open
Set rs = New ADODB.Recordset
rs.Open "select * from policy_customer", con, 2, 2
rs.AddNew
rs.Fields("PolicyCustomer_Number") = txtpcnumber.Text
rs.Fields("Customer_Name") = txtcustname.Text
rs.Fields("Customer_Number") = Combo1.Text
rs.Fields("Policy_Number") = txtpcnumber.Text
rs.Fields("Policy_Name") = txtplan_name.Text
rs.Fields("Term") = Combo3.Text

rs.Fields("Maturity_Date") = DTPicker1.Value
rs.Fields("Premium") = txtpremium.Text
rs.Fields("DOC") = DTPicker2.Value
rs.Fields("SurvivalBenifit") = txtbenifit.Text
rs.Fields("Nominee_Name") = txtnominee_name.Text
rs.Fields("Nominee_Contact_NO") = txtcontact_no.Text
rs.Fields("Customer_Add") = txtaddress.Text

rs.Update


MsgBox "Record Saved successfully", vbInformation

rs.Close

End Sub



Private Sub txtamt_GotFocus()
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

