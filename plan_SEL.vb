Option Strict Off
Option Explicit On
Friend Class frm_presentasion
	Inherits System.Windows.Forms.Form
	'UPGRADE_WARNING: Event Combo1.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub Combo1_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Combo1.SelectedIndexChanged
		Label3.Text = Combo1.Text
		
		If Combo1.Text = "Children Plan" Then
			List1.Items.Clear()
			Image2.Image = System.Drawing.Image.FromFile(&"\child.jpg")
			List1.Items.Add(("Komal Jeevan"))
			List1.Items.Add(("Jeevan Kishor"))
			List1.Items.Add(("Child Career Plan"))
			List1.Items.Add(("Child Future Plan"))
		End If
		
		
		If Combo1.Text = "Penssion Plan" Then
			List1.Items.Clear()
			Image2.Image = System.Drawing.Image.FromFile("" & My.Application.Info.DirectoryPath & "\penssion.jpg")
			List1.Items.Add(("Jeevan Akshay"))
			List1.Items.Add(("New Jeevan Suraksha - 1"))
			List1.Items.Add(("New Jeevan Dhara - 1"))
			List1.Items.Add(("Jeevan Nidhi"))
		End If
		
		
		If Combo1.Text = "Money Back Plan" Then
			List1.Items.Clear()
			Image2.Image = System.Drawing.Image.FromFile("" & My.Application.Info.DirectoryPath & "\moneyback.gif")
			List1.Items.Add(("Bima Bachat"))
			List1.Items.Add(("Money Back(20 years)"))
			List1.Items.Add(("Money Back (25 years)"))
			List1.Items.Add(("Jeevan Surbhi(20 Years)"))
		End If
		
		
		If Combo1.Text = "Endowment Plan" Then
			List1.Items.Clear()
			Image2.Image = System.Drawing.Image.FromFile("" & My.Application.Info.DirectoryPath & "\endo.gif")
			List1.Items.Add(("Endowment with profit"))
			List1.Items.Add(("Jeevan Chhaya"))
			List1.Items.Add(("Jeevan Pramukh"))
			List1.Items.Add(("Jeevan Anurag"))
		End If
		
		
		If Combo1.Text = "Term Inssurance Plan" Then
			List1.Items.Clear()
			Image2.Image = System.Drawing.Image.FromFile("" & My.Application.Info.DirectoryPath & "\term.jpg")
			List1.Items.Add(("Temporary Assurance"))
			List1.Items.Add(("Anmol Jeevan-1"))
			List1.Items.Add(("Amulya Jeevan-1"))
		End If
		
		
		If Combo1.Text = "Special Plan" Then
			List1.Items.Clear()
			Image2.Image = System.Drawing.Image.FromFile("" & My.Application.Info.DirectoryPath & "\special.jpg")
			
			List1.Items.Add(("Jeevan Anand"))
			List1.Items.Add(("Jeevan Saral"))
		End If
	End Sub
	
	Private Sub Command1_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command1.Click
        'frm_tc.Show()
	End Sub
	
	'UPGRADE_WARNING: Event List1.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub List1_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles List1.SelectedIndexChanged
		Plandetail.Text = ""
		'UPGRADE_NOTE: str was upgraded to str_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim str_Renamed As String
		If List1.Text = "Komal Jeevan" Then
			FileOpen(1, "" & My.Application.Info.DirectoryPath & "\Detail\159.txt", OpenMode.Input)
			Do Until EOF(1)
				Input(1, str_Renamed)
				Plandetail.Text = Plandetail.Text & str_Renamed & vbCrLf
			Loop 
			FileClose(1)
		End If
		
		If List1.Text = "Jeevan Kishor" Then
			FileOpen(1, "" & My.Application.Info.DirectoryPath & "\Detail\102.txt", OpenMode.Input)
			Do Until EOF(1)
				Input(1, str_Renamed)
				Plandetail.Text = Plandetail.Text & str_Renamed & vbCrLf
			Loop 
			FileClose(1)
		End If
		
		If List1.Text = "Child Career Plan" Then
			FileOpen(1, "" & My.Application.Info.DirectoryPath & "\Detail\184.txt", OpenMode.Input)
			Do Until EOF(1)
				Input(1, str_Renamed)
				Plandetail.Text = Plandetail.Text & str_Renamed & vbCrLf
			Loop 
			FileClose(1)
		End If
		
		If List1.Text = "Child Future Plan" Then
			FileOpen(1, "" & My.Application.Info.DirectoryPath & "\Detail\159.txt", OpenMode.Input)
			Do Until EOF(1)
				Input(1, str_Renamed)
				Plandetail.Text = Plandetail.Text & str_Renamed & vbCrLf
			Loop 
			FileClose(1)
		End If
		
		If List1.Text = "New Jeevan Suraksha - 1" Then
			FileOpen(1, "" & My.Application.Info.DirectoryPath & "\Detail\147.txt", OpenMode.Input)
			Do Until EOF(1)
				Input(1, str_Renamed)
				Plandetail.Text = Plandetail.Text & str_Renamed & vbCrLf
			Loop 
			FileClose(1)
		End If
		
		If List1.Text = "New Jeevan Dhara - 1" Then
			FileOpen(1, "" & My.Application.Info.DirectoryPath & "\Detail\148.txt", OpenMode.Input)
			Do Until EOF(1)
				Input(1, str_Renamed)
				Plandetail.Text = Plandetail.Text & str_Renamed & vbCrLf
			Loop 
			FileClose(1)
		End If
		
		
		If List1.Text = "Jeevan Nidhi" Then
			FileOpen(1, "" & My.Application.Info.DirectoryPath & "\Detail\169.txt", OpenMode.Input)
			Do Until EOF(1)
				Input(1, str_Renamed)
				Plandetail.Text = Plandetail.Text & str_Renamed & vbCrLf
			Loop 
			FileClose(1)
		End If
		
		
		If List1.Text = "Jeevan Akshay" Then
			FileOpen(1, "" & My.Application.Info.DirectoryPath & "\Detail\189.txt", OpenMode.Input)
			Do Until EOF(1)
				Input(1, str_Renamed)
				Plandetail.Text = Plandetail.Text & str_Renamed & vbCrLf
			Loop 
			FileClose(1)
		End If
		
		
		If List1.Text = "Temporary Assurance" Then
			FileOpen(1, "" & My.Application.Info.DirectoryPath & "\Detail\43.txt", OpenMode.Input)
			Do Until EOF(1)
				Input(1, str_Renamed)
				Plandetail.Text = Plandetail.Text & str_Renamed & vbCrLf
			Loop 
			FileClose(1)
		End If
		
		If List1.Text = "Anmol Jeevan-1" Then
			FileOpen(1, "" & My.Application.Info.DirectoryPath & "\Detail\164.txt", OpenMode.Input)
			Do Until EOF(1)
				Input(1, str_Renamed)
				Plandetail.Text = Plandetail.Text & str_Renamed & vbCrLf
			Loop 
			FileClose(1)
		End If
		
		If List1.Text = "Amulya Jeevan-1" Then
			FileOpen(1, "" & My.Application.Info.DirectoryPath & "\Detail\190.txt", OpenMode.Input)
			Do Until EOF(1)
				Input(1, str_Renamed)
				Plandetail.Text = Plandetail.Text & str_Renamed & vbCrLf
			Loop 
			FileClose(1)
		End If
		
		
		If List1.Text = "Money Back(20 years)" Then
			FileOpen(1, "" & My.Application.Info.DirectoryPath & "\Detail\75.txt", OpenMode.Input)
			Do Until EOF(1)
				Input(1, str_Renamed)
				Plandetail.Text = Plandetail.Text & str_Renamed & vbCrLf
			Loop 
			FileClose(1)
		End If
		
		If List1.Text = "Money Back (25 years)" Then
			FileOpen(1, "" & My.Application.Info.DirectoryPath & "\Detail\93.txt", OpenMode.Input)
			Do Until EOF(1)
				Input(1, str_Renamed)
				Plandetail.Text = Plandetail.Text & str_Renamed & vbCrLf
			Loop 
			FileClose(1)
		End If
		
		If List1.Text = "Jeevan Surbhi(20 Years)" Then
			FileOpen(1, "" & My.Application.Info.DirectoryPath & "\Detail\107.txt", OpenMode.Input)
			Do Until EOF(1)
				Input(1, str_Renamed)
				Plandetail.Text = Plandetail.Text & str_Renamed & vbCrLf
			Loop 
			FileClose(1)
		End If
		
		If List1.Text = "Bima Bachat" Then
			FileOpen(1, "" & My.Application.Info.DirectoryPath & "\Detail\175.txt", OpenMode.Input)
			Do Until EOF(1)
				Input(1, str_Renamed)
				Plandetail.Text = Plandetail.Text & str_Renamed & vbCrLf
			Loop 
			FileClose(1)
		End If
		
		If List1.Text = "Jeevan Anurag" Then
			FileOpen(1, "" & My.Application.Info.DirectoryPath & "\Detail\168.txt", OpenMode.Input)
			Do Until EOF(1)
				Input(1, str_Renamed)
				Plandetail.Text = Plandetail.Text & str_Renamed & vbCrLf
			Loop 
			FileClose(1)
		End If
		
		
		If List1.Text = "Jeevan Pramukh" Then
			FileOpen(1, "" & My.Application.Info.DirectoryPath & "\Detail\167.txt", OpenMode.Input)
			Do Until EOF(1)
				Input(1, str_Renamed)
				Plandetail.Text = Plandetail.Text & str_Renamed & vbCrLf
			Loop 
			FileClose(1)
		End If
		
		
		If List1.Text = "Jeevan Chhaya" Then
			FileOpen(1, "" & My.Application.Info.DirectoryPath & "\Detail\103.txt", OpenMode.Input)
			Do Until EOF(1)
				Input(1, str_Renamed)
				Plandetail.Text = Plandetail.Text & str_Renamed & vbCrLf
			Loop 
			FileClose(1)
		End If
		
		
		If List1.Text = "Endowment with profit" Then
			FileOpen(1, "" & My.Application.Info.DirectoryPath & "\Detail\2.txt", OpenMode.Input)
			Do Until EOF(1)
				Input(1, str_Renamed)
				Plandetail.Text = Plandetail.Text & str_Renamed & vbCrLf
			Loop 
			FileClose(1)
		End If
		
		
		If List1.Text = "Jeevan Saral" Then
			FileOpen(1, "" & My.Application.Info.DirectoryPath & "\Detail\165.txt", OpenMode.Input)
			Do Until EOF(1)
				Input(1, str_Renamed)
				Plandetail.Text = Plandetail.Text & str_Renamed & vbCrLf
			Loop 
			FileClose(1)
		End If
		
		
		If List1.Text = "Jeevan Anand" Then
			FileOpen(1, "" & My.Application.Info.DirectoryPath & "\Detail\149.txt", OpenMode.Input)
			Do Until EOF(1)
				Input(1, str_Renamed)
				Plandetail.Text = Plandetail.Text & str_Renamed & vbCrLf
			Loop 
			FileClose(1)
		End If
		
		
	End Sub

    Private Sub frm_presentasion_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub
End Class