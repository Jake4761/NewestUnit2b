<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Details
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.cmdSave = New System.Windows.Forms.Button()
        Me.lblName = New System.Windows.Forms.Label()
        Me.TxtLevel = New System.Windows.Forms.TextBox()
        Me.txtLocation = New System.Windows.Forms.TextBox()
        Me.txtDay = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.txtSession = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.txtType = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.cmdCount = New System.Windows.Forms.Button()
        Me.btnSwapFormCustomers = New System.Windows.Forms.Button()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.Panel1.Controls.Add(Me.btnSwapFormCustomers)
        Me.Panel1.Controls.Add(Me.cmdSave)
        Me.Panel1.Controls.Add(Me.lblName)
        Me.Panel1.Controls.Add(Me.TxtLevel)
        Me.Panel1.Controls.Add(Me.txtLocation)
        Me.Panel1.Controls.Add(Me.txtDay)
        Me.Panel1.Controls.Add(Me.Label6)
        Me.Panel1.Controls.Add(Me.txtSession)
        Me.Panel1.Controls.Add(Me.Label5)
        Me.Panel1.Controls.Add(Me.txtType)
        Me.Panel1.Controls.Add(Me.Label4)
        Me.Panel1.Controls.Add(Me.Label3)
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Controls.Add(Me.cmdCount)
        Me.Panel1.Location = New System.Drawing.Point(8, 8)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(254, 236)
        Me.Panel1.TabIndex = 19
        '
        'cmdSave
        '
        Me.cmdSave.BackColor = System.Drawing.Color.Gray
        Me.cmdSave.Location = New System.Drawing.Point(11, 161)
        Me.cmdSave.Name = "cmdSave"
        Me.cmdSave.Size = New System.Drawing.Size(96, 32)
        Me.cmdSave.TabIndex = 18
        Me.cmdSave.Text = "Save"
        Me.cmdSave.UseVisualStyleBackColor = False
        '
        'lblName
        '
        Me.lblName.AutoSize = True
        Me.lblName.BackColor = System.Drawing.Color.Gray
        Me.lblName.Font = New System.Drawing.Font("Microsoft Sans Serif", 18.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblName.Location = New System.Drawing.Point(0, 0)
        Me.lblName.Name = "lblName"
        Me.lblName.Size = New System.Drawing.Size(255, 29)
        Me.lblName.TabIndex = 9
        Me.lblName.Text = "Fitness Leisure Centre"
        '
        'TxtLevel
        '
        Me.TxtLevel.BackColor = System.Drawing.Color.Silver
        Me.TxtLevel.Location = New System.Drawing.Point(72, 135)
        Me.TxtLevel.Name = "TxtLevel"
        Me.TxtLevel.Size = New System.Drawing.Size(112, 20)
        Me.TxtLevel.TabIndex = 6
        '
        'txtLocation
        '
        Me.txtLocation.BackColor = System.Drawing.Color.Silver
        Me.txtLocation.Location = New System.Drawing.Point(72, 112)
        Me.txtLocation.Name = "txtLocation"
        Me.txtLocation.Size = New System.Drawing.Size(112, 20)
        Me.txtLocation.TabIndex = 5
        '
        'txtDay
        '
        Me.txtDay.BackColor = System.Drawing.Color.Silver
        Me.txtDay.Location = New System.Drawing.Point(72, 88)
        Me.txtDay.Name = "txtDay"
        Me.txtDay.Size = New System.Drawing.Size(112, 20)
        Me.txtDay.TabIndex = 3
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(37, 138)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(33, 13)
        Me.Label6.TabIndex = 15
        Me.Label6.Text = "Level"
        '
        'txtSession
        '
        Me.txtSession.BackColor = System.Drawing.Color.Silver
        Me.txtSession.Location = New System.Drawing.Point(72, 64)
        Me.txtSession.Name = "txtSession"
        Me.txtSession.Size = New System.Drawing.Size(112, 20)
        Me.txtSession.TabIndex = 2
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(24, 115)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(48, 13)
        Me.Label5.TabIndex = 14
        Me.Label5.Text = "Location"
        '
        'txtType
        '
        Me.txtType.BackColor = System.Drawing.Color.Silver
        Me.txtType.Location = New System.Drawing.Point(72, 40)
        Me.txtType.Name = "txtType"
        Me.txtType.Size = New System.Drawing.Size(112, 20)
        Me.txtType.TabIndex = 1
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(8, 112)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(0, 13)
        Me.Label4.TabIndex = 13
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(45, 91)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(26, 13)
        Me.Label3.TabIndex = 12
        Me.Label3.Text = "Day"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(26, 67)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(44, 13)
        Me.Label2.TabIndex = 11
        Me.Label2.Text = "Session"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(40, 40)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(31, 13)
        Me.Label1.TabIndex = 10
        Me.Label1.Text = "Type"
        '
        'cmdCount
        '
        Me.cmdCount.BackColor = System.Drawing.Color.Gray
        Me.cmdCount.Location = New System.Drawing.Point(144, 161)
        Me.cmdCount.Name = "cmdCount"
        Me.cmdCount.Size = New System.Drawing.Size(96, 32)
        Me.cmdCount.TabIndex = 0
        Me.cmdCount.Text = "Count"
        Me.cmdCount.UseVisualStyleBackColor = False
        '
        'btnSwapFormCustomers
        '
        Me.btnSwapFormCustomers.Location = New System.Drawing.Point(72, 199)
        Me.btnSwapFormCustomers.Name = "btnSwapFormCustomers"
        Me.btnSwapFormCustomers.Size = New System.Drawing.Size(97, 34)
        Me.btnSwapFormCustomers.TabIndex = 19
        Me.btnSwapFormCustomers.Text = "Swap To Customers Form"
        Me.btnSwapFormCustomers.UseVisualStyleBackColor = True
        '
        'Details
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(270, 242)
        Me.Controls.Add(Me.Panel1)
        Me.Name = "Details"
        Me.Text = "Details"
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents lblName As System.Windows.Forms.Label
    Friend WithEvents TxtLevel As System.Windows.Forms.TextBox
    Friend WithEvents txtLocation As System.Windows.Forms.TextBox
    Friend WithEvents txtDay As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents txtSession As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents txtType As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents cmdCount As System.Windows.Forms.Button
    Friend WithEvents cmdSave As System.Windows.Forms.Button
    Friend WithEvents btnSwapFormCustomers As System.Windows.Forms.Button

End Class
