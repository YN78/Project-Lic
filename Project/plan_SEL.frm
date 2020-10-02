VERSION 5.00
Begin VB.Form frm_presentasion 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Plan Presentation"
   ClientHeight    =   10935
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   15120
   LinkTopic       =   "Form1"
   ScaleHeight     =   10935
   ScaleWidth      =   15120
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox Plandetail 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   5400
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   7320
      Width           =   5535
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
      Height          =   735
      Left            =   14640
      TabIndex        =   4
      Top             =   10200
      Width           =   3015
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   2670
      ItemData        =   "plan_SEL.frx":0000
      Left            =   10080
      List            =   "plan_SEL.frx":0002
      TabIndex        =   2
      Top             =   3360
      Width           =   6615
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00FFFFFF&
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
      Height          =   480
      ItemData        =   "plan_SEL.frx":0004
      Left            =   10455
      List            =   "plan_SEL.frx":001A
      TabIndex        =   1
      Text            =   "PLANS"
      Top             =   255
      Width           =   4920
   End
   Begin VB.Image Image2 
      Height          =   3240
      Left            =   2040
      Picture         =   "plan_SEL.frx":0081
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   4320
   End
   Begin VB.Image Image1 
      Height          =   1425
      Left            =   0
      Picture         =   "plan_SEL.frx":1867
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3330
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6720
      TabIndex        =   5
      Top             =   2640
      Width           =   3615
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "SELECT PLAN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6120
      TabIndex        =   3
      Top             =   255
      Width           =   3135
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "PLAN DETAILS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   840
      TabIndex        =   0
      Top             =   7200
      Width           =   3255
   End
End
Attribute VB_Name = "frm_presentasion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Click()
Label3.Caption = Combo1.Text

If Combo1.Text = "Children Plan" Then
List1.Clear
Image2.Picture = LoadPicture("" & App.Path & "\child.jpg")
List1.AddItem ("Komal Jeevan")
List1.AddItem ("Jeevan Kishor")
List1.AddItem ("Child Career Plan")
List1.AddItem ("Child Future Plan")
End If


If Combo1.Text = "Penssion Plan" Then
List1.Clear
 Image2.Picture = LoadPicture("" & App.Path & "\penssion.jpg")
List1.AddItem ("Jeevan Akshay")
List1.AddItem ("New Jeevan Suraksha - 1")
List1.AddItem ("New Jeevan Dhara - 1")
List1.AddItem ("Jeevan Nidhi")
End If


If Combo1.Text = "Money Back Plan" Then
List1.Clear
 Image2.Picture = LoadPicture("" & App.Path & "\moneyback.gif")
List1.AddItem ("Bima Bachat")
List1.AddItem ("Money Back(20 years)")
List1.AddItem ("Money Back (25 years)")
List1.AddItem ("Jeevan Surbhi(20 Years)")
End If


If Combo1.Text = "Endowment Plan" Then
List1.Clear
 Image2.Picture = LoadPicture("" & App.Path & "\endo.gif")
List1.AddItem ("Endowment with profit")
List1.AddItem ("Jeevan Chhaya")
List1.AddItem ("Jeevan Pramukh")
List1.AddItem ("Jeevan Anurag")
End If


If Combo1.Text = "Term Inssurance Plan" Then
List1.Clear
 Image2.Picture = LoadPicture("" & App.Path & "\term.jpg")
List1.AddItem ("Temporary Assurance")
List1.AddItem ("Anmol Jeevan-1")
List1.AddItem ("Amulya Jeevan-1")
End If


If Combo1.Text = "Special Plan" Then
List1.Clear
Image2.Picture = LoadPicture("" & App.Path & "\special.jpg")
 
List1.AddItem ("Jeevan Anand")
List1.AddItem ("Jeevan Saral")
End If
End Sub

Private Sub Command1_Click()
frm_tc.Show
End Sub

Private Sub List1_Click()
Plandetail.Text = ""
Dim str As String
If List1.Text = "Komal Jeevan" Then
Open "" & App.Path & "\Detail\159.txt" For Input As #1
Do Until EOF(1)
Input #1, str
Plandetail.Text = Plandetail.Text & str & vbCrLf
Loop
Close #1
End If

If List1.Text = "Jeevan Kishor" Then
Open "" & App.Path & "\Detail\102.txt" For Input As #1
Do Until EOF(1)
Input #1, str
Plandetail.Text = Plandetail.Text & str & vbCrLf
Loop
Close #1
End If

If List1.Text = "Child Career Plan" Then
Open "" & App.Path & "\Detail\184.txt" For Input As #1
Do Until EOF(1)
Input #1, str
Plandetail.Text = Plandetail.Text & str & vbCrLf
Loop
Close #1
End If

If List1.Text = "Child Future Plan" Then
Open "" & App.Path & "\Detail\159.txt" For Input As #1
Do Until EOF(1)
Input #1, str
Plandetail.Text = Plandetail.Text & str & vbCrLf
Loop
Close #1
End If

If List1.Text = "New Jeevan Suraksha - 1" Then
Open "" & App.Path & "\Detail\147.txt" For Input As #1
Do Until EOF(1)
Input #1, str
Plandetail.Text = Plandetail.Text & str & vbCrLf
Loop
Close #1
End If

If List1.Text = "New Jeevan Dhara - 1" Then
Open "" & App.Path & "\Detail\148.txt" For Input As #1
Do Until EOF(1)
Input #1, str
Plandetail.Text = Plandetail.Text & str & vbCrLf
Loop
Close #1
End If


If List1.Text = "Jeevan Nidhi" Then
Open "" & App.Path & "\Detail\169.txt" For Input As #1
Do Until EOF(1)
Input #1, str
Plandetail.Text = Plandetail.Text & str & vbCrLf
Loop
Close #1
End If


If List1.Text = "Jeevan Akshay" Then
Open "" & App.Path & "\Detail\189.txt" For Input As #1
Do Until EOF(1)
Input #1, str
Plandetail.Text = Plandetail.Text & str & vbCrLf
Loop
Close #1
End If


If List1.Text = "Temporary Assurance" Then
Open "" & App.Path & "\Detail\43.txt" For Input As #1
Do Until EOF(1)
Input #1, str
Plandetail.Text = Plandetail.Text & str & vbCrLf
Loop
Close #1
End If

If List1.Text = "Anmol Jeevan-1" Then
Open "" & App.Path & "\Detail\164.txt" For Input As #1
Do Until EOF(1)
Input #1, str
Plandetail.Text = Plandetail.Text & str & vbCrLf
Loop
Close #1
End If

If List1.Text = "Amulya Jeevan-1" Then
Open "" & App.Path & "\Detail\190.txt" For Input As #1
Do Until EOF(1)
Input #1, str
Plandetail.Text = Plandetail.Text & str & vbCrLf
Loop
Close #1
End If


If List1.Text = "Money Back(20 years)" Then
Open "" & App.Path & "\Detail\75.txt" For Input As #1
Do Until EOF(1)
Input #1, str
Plandetail.Text = Plandetail.Text & str & vbCrLf
Loop
Close #1
End If

If List1.Text = "Money Back (25 years)" Then
Open "" & App.Path & "\Detail\93.txt" For Input As #1
Do Until EOF(1)
Input #1, str
Plandetail.Text = Plandetail.Text & str & vbCrLf
Loop
Close #1
End If

If List1.Text = "Jeevan Surbhi(20 Years)" Then
Open "" & App.Path & "\Detail\107.txt" For Input As #1
Do Until EOF(1)
Input #1, str
Plandetail.Text = Plandetail.Text & str & vbCrLf
Loop
Close #1
End If

If List1.Text = "Bima Bachat" Then
Open "" & App.Path & "\Detail\175.txt" For Input As #1
Do Until EOF(1)
Input #1, str
Plandetail.Text = Plandetail.Text & str & vbCrLf
Loop
Close #1
End If

If List1.Text = "Jeevan Anurag" Then
Open "" & App.Path & "\Detail\168.txt" For Input As #1
Do Until EOF(1)
Input #1, str
Plandetail.Text = Plandetail.Text & str & vbCrLf
Loop
Close #1
End If


If List1.Text = "Jeevan Pramukh" Then
Open "" & App.Path & "\Detail\167.txt" For Input As #1
Do Until EOF(1)
Input #1, str
Plandetail.Text = Plandetail.Text & str & vbCrLf
Loop
Close #1
End If


If List1.Text = "Jeevan Chhaya" Then
Open "" & App.Path & "\Detail\103.txt" For Input As #1
Do Until EOF(1)
Input #1, str
Plandetail.Text = Plandetail.Text & str & vbCrLf
Loop
Close #1
End If


If List1.Text = "Endowment with profit" Then
Open "" & App.Path & "\Detail\2.txt" For Input As #1
Do Until EOF(1)
Input #1, str
Plandetail.Text = Plandetail.Text & str & vbCrLf
Loop
Close #1
End If


If List1.Text = "Jeevan Saral" Then
Open "" & App.Path & "\Detail\165.txt" For Input As #1
Do Until EOF(1)
Input #1, str
Plandetail.Text = Plandetail.Text & str & vbCrLf
Loop
Close #1
End If


If List1.Text = "Jeevan Anand" Then
Open "" & App.Path & "\Detail\149.txt" For Input As #1
Do Until EOF(1)
Input #1, str
Plandetail.Text = Plandetail.Text & str & vbCrLf
Loop
Close #1
End If


End Sub

