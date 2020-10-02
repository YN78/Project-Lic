VERSION 5.00
Begin VB.MDIForm Main 
   BackColor       =   &H8000000C&
   Caption         =   "Parent Form"
   ClientHeight    =   8340
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   16005
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu Login 
      Caption         =   "Login"
      Index           =   1
      NegotiatePosition=   1  'Left
   End
   Begin VB.Menu PlanSelection 
      Caption         =   "        Plan Selection"
      Index           =   2
   End
   Begin VB.Menu Plan_Information 
      Caption         =   "        Plan Information"
      Index           =   3
   End
   Begin VB.Menu Premium_calc 
      Caption         =   "        Premium Calculator"
      Index           =   4
   End
   Begin VB.Menu Customer 
      Caption         =   "        Customer"
      Index           =   4
   End
   Begin VB.Menu Prospactor 
      Caption         =   "        Prospactor"
      Index           =   5
   End
   Begin VB.Menu Alert 
      Caption         =   "        Alert"
      Index           =   6
   End
   Begin VB.Menu Commssion 
      Caption         =   "       Commssion"
      Index           =   7
   End
   Begin VB.Menu Repoert 
      Caption         =   "       Repoert"
      Index           =   8
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Login_Click(Index As Integer)
Login.Show
End Sub

