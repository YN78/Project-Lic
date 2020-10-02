VERSION 5.00
Begin VB.Form frm_custsearch 
   Caption         =   "Search"
   ClientHeight    =   7935
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14010
   LinkTopic       =   "Form1"
   ScaleHeight     =   7935
   ScaleWidth      =   14010
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
End
Attribute VB_Name = "frm_custsearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset

Private Sub Command1_Click()
Adodc1.RecordSource = "select * from Customer where Customer_Name  like '" & Text1.Text & "'"
Adodc1.Refresh
End Sub

Private Sub Command2_Click()
Adodc1.RecordSource = "select * from Customer"
Adodc1.Refresh
End Sub

Private Sub Command3_Click()
Dim id As Integer
id = Val(InputBox("Enter id"))
rs.Open "select * from Customer where  Customer_Number=" & id
rs.Delete
rs.Close
End Sub

Private Sub Form_Load()
Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\lic.mdb"
Set con = New ADODB.Connection
con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\lic.mdb"
con.Open

Set rs = New ADODB.Recordset
rs.Open "select * from Customer", con, adOpenKeyset, adLockOptimistic
rs.Close
End Sub


Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Or KeyAscii = 27 Then
ElseIf KeyAscii = 13 Then
Text1.SetFocus
ElseIf (KeyAscii >= 65 And KeyAscii <= 91) Or (KeyAscii >= 97 And KeyAscii <= 122) Then

Else
MsgBox "Enter Character Only"
KeyAscii = 0
End If
End Sub
