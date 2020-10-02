<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frm_presentasion
#Region "Windows Form Designer generated code "
	<System.Diagnostics.DebuggerNonUserCode()> Public Sub New()
		MyBase.New()
		'This call is required by the Windows Form Designer.
		InitializeComponent()
	End Sub
	'Form overrides dispose to clean up the component list.
	<System.Diagnostics.DebuggerNonUserCode()> Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
		If Disposing Then
			If Not components Is Nothing Then
				components.Dispose()
			End If
		End If
		MyBase.Dispose(Disposing)
	End Sub
	'Required by the Windows Form Designer
	Private components As System.ComponentModel.IContainer
	Public ToolTip1 As System.Windows.Forms.ToolTip
	Public WithEvents Command1 As System.Windows.Forms.Button
	Public WithEvents Plandetail As System.Windows.Forms.RichTextBox
	Public WithEvents List1 As System.Windows.Forms.ListBox
	Public WithEvents Combo1 As System.Windows.Forms.ComboBox
	Public WithEvents Image2 As System.Windows.Forms.PictureBox
	Public WithEvents Image1 As System.Windows.Forms.PictureBox
	Public WithEvents Label3 As System.Windows.Forms.Label
	Public WithEvents Label2 As System.Windows.Forms.Label
	Public WithEvents Label1 As System.Windows.Forms.Label
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frm_presentasion))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.Command1 = New System.Windows.Forms.Button
        Me.Plandetail = New System.Windows.Forms.RichTextBox
        Me.List1 = New System.Windows.Forms.ListBox
        Me.Combo1 = New System.Windows.Forms.ComboBox
        Me.Image2 = New System.Windows.Forms.PictureBox
        Me.Image1 = New System.Windows.Forms.PictureBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        CType(Me.Image2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Image1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Command1
        '
        Me.Command1.BackColor = System.Drawing.SystemColors.Control
        Me.Command1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Command1.Font = New System.Drawing.Font("Arial", 13.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Command1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Command1.Location = New System.Drawing.Point(976, 680)
        Me.Command1.Name = "Command1"
        Me.Command1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Command1.Size = New System.Drawing.Size(201, 49)
        Me.Command1.TabIndex = 4
        Me.Command1.Text = "Next"
        Me.Command1.UseVisualStyleBackColor = False
        '
        'Plandetail
        '
        Me.Plandetail.Font = New System.Drawing.Font("Arial", 13.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Plandetail.Location = New System.Drawing.Point(336, 472)
        Me.Plandetail.Name = "Plandetail"
        Me.Plandetail.ReadOnly = True
        Me.Plandetail.Size = New System.Drawing.Size(577, 225)
        Me.Plandetail.TabIndex = 6
        Me.Plandetail.Text = ""
        '
        'List1
        '
        Me.List1.BackColor = System.Drawing.Color.White
        Me.List1.Cursor = System.Windows.Forms.Cursors.Default
        Me.List1.Font = New System.Drawing.Font("Arial", 18.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.List1.ForeColor = System.Drawing.Color.Blue
        Me.List1.ItemHeight = 27
        Me.List1.Location = New System.Drawing.Point(672, 224)
        Me.List1.Name = "List1"
        Me.List1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.List1.Size = New System.Drawing.Size(441, 166)
        Me.List1.TabIndex = 2
        '
        'Combo1
        '
        Me.Combo1.BackColor = System.Drawing.Color.White
        Me.Combo1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Combo1.Font = New System.Drawing.Font("Arial", 13.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Combo1.ForeColor = System.Drawing.Color.Blue
        Me.Combo1.Items.AddRange(New Object() {"Children Plan", "Endowment Plan", "Money Back Plan", "Penssion Plan", "Term Inssurance Plan", "Special Plan"})
        Me.Combo1.Location = New System.Drawing.Point(697, 17)
        Me.Combo1.Name = "Combo1"
        Me.Combo1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Combo1.Size = New System.Drawing.Size(328, 29)
        Me.Combo1.TabIndex = 1
        Me.Combo1.Text = "PLANS"
        '
        'Image2
        '
        Me.Image2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Image2.Image = CType(resources.GetObject("Image2.Image"), System.Drawing.Image)
        Me.Image2.Location = New System.Drawing.Point(136, 144)
        Me.Image2.Name = "Image2"
        Me.Image2.Size = New System.Drawing.Size(288, 216)
        Me.Image2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.Image2.TabIndex = 7
        Me.Image2.TabStop = False
        '
        'Image1
        '
        Me.Image1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Image1.Image = CType(resources.GetObject("Image1.Image"), System.Drawing.Image)
        Me.Image1.Location = New System.Drawing.Point(0, 0)
        Me.Image1.Name = "Image1"
        Me.Image1.Size = New System.Drawing.Size(222, 95)
        Me.Image1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.Image1.TabIndex = 8
        Me.Image1.TabStop = False
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label3.Font = New System.Drawing.Font("Arial", 18.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label3.Location = New System.Drawing.Point(448, 176)
        Me.Label3.Name = "Label3"
        Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label3.Size = New System.Drawing.Size(241, 25)
        Me.Label3.TabIndex = 5
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label2.Font = New System.Drawing.Font("Arial", 13.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label2.Location = New System.Drawing.Point(408, 17)
        Me.Label2.Name = "Label2"
        Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label2.Size = New System.Drawing.Size(209, 41)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "SELECT PLAN"
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.Font = New System.Drawing.Font("Arial", 18.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(56, 480)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(217, 57)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "PLAN DETAILS"
        '
        'frm_presentasion
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(1264, 746)
        Me.Controls.Add(Me.Command1)
        Me.Controls.Add(Me.Plandetail)
        Me.Controls.Add(Me.List1)
        Me.Controls.Add(Me.Combo1)
        Me.Controls.Add(Me.Image2)
        Me.Controls.Add(Me.Image1)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Location = New System.Drawing.Point(8, 30)
        Me.Name = "frm_presentasion"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text = "Plan Presentation"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        CType(Me.Image2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Image1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
#End Region 
End Class