<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Customers
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
        Me.btnInput = New System.Windows.Forms.Button()
        Me.lblName = New System.Windows.Forms.Label()
        Me.lblLastName = New System.Windows.Forms.Label()
        Me.lblEmail = New System.Windows.Forms.Label()
        Me.txtFirstName = New System.Windows.Forms.TextBox()
        Me.txtLastName = New System.Windows.Forms.TextBox()
        Me.txtEmail = New System.Windows.Forms.TextBox()
        Me.lblPhone = New System.Windows.Forms.Label()
        Me.lblDoB = New System.Windows.Forms.Label()
        Me.lblAddress = New System.Windows.Forms.Label()
        Me.lblID = New System.Windows.Forms.Label()
        Me.txtPhone = New System.Windows.Forms.TextBox()
        Me.dtpDoB = New System.Windows.Forms.DateTimePicker()
        Me.txtAddress = New System.Windows.Forms.TextBox()
        Me.txtSession = New System.Windows.Forms.TextBox()
        Me.btnSearch = New System.Windows.Forms.Button()
        Me.lblSearchID = New System.Windows.Forms.Label()
        Me.cmbID = New System.Windows.Forms.ComboBox()
        Me.btnSwapFormActivities = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'btnInput
        '
        Me.btnInput.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnInput.Location = New System.Drawing.Point(28, 294)
        Me.btnInput.Name = "btnInput"
        Me.btnInput.Size = New System.Drawing.Size(514, 84)
        Me.btnInput.TabIndex = 0
        Me.btnInput.Text = "INPUT"
        Me.btnInput.UseVisualStyleBackColor = True
        '
        'lblName
        '
        Me.lblName.AutoSize = True
        Me.lblName.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblName.Location = New System.Drawing.Point(24, 21)
        Me.lblName.Name = "lblName"
        Me.lblName.Size = New System.Drawing.Size(59, 20)
        Me.lblName.TabIndex = 1
        Me.lblName.Text = "Name :"
        '
        'lblLastName
        '
        Me.lblLastName.AutoSize = True
        Me.lblLastName.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblLastName.Location = New System.Drawing.Point(24, 56)
        Me.lblLastName.Name = "lblLastName"
        Me.lblLastName.Size = New System.Drawing.Size(94, 20)
        Me.lblLastName.TabIndex = 2
        Me.lblLastName.Text = "Last Name :"
        '
        'lblEmail
        '
        Me.lblEmail.AutoSize = True
        Me.lblEmail.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblEmail.Location = New System.Drawing.Point(24, 93)
        Me.lblEmail.Name = "lblEmail"
        Me.lblEmail.Size = New System.Drawing.Size(56, 20)
        Me.lblEmail.TabIndex = 4
        Me.lblEmail.Text = "Email :"
        '
        'txtFirstName
        '
        Me.txtFirstName.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtFirstName.Location = New System.Drawing.Point(342, 18)
        Me.txtFirstName.Name = "txtFirstName"
        Me.txtFirstName.Size = New System.Drawing.Size(200, 26)
        Me.txtFirstName.TabIndex = 9
        '
        'txtLastName
        '
        Me.txtLastName.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtLastName.Location = New System.Drawing.Point(342, 53)
        Me.txtLastName.Name = "txtLastName"
        Me.txtLastName.Size = New System.Drawing.Size(200, 26)
        Me.txtLastName.TabIndex = 10
        '
        'txtEmail
        '
        Me.txtEmail.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtEmail.Location = New System.Drawing.Point(342, 93)
        Me.txtEmail.Name = "txtEmail"
        Me.txtEmail.Size = New System.Drawing.Size(200, 26)
        Me.txtEmail.TabIndex = 12
        '
        'lblPhone
        '
        Me.lblPhone.AutoSize = True
        Me.lblPhone.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!)
        Me.lblPhone.Location = New System.Drawing.Point(26, 133)
        Me.lblPhone.Name = "lblPhone"
        Me.lblPhone.Size = New System.Drawing.Size(119, 20)
        Me.lblPhone.TabIndex = 13
        Me.lblPhone.Text = "Phone Number:"
        '
        'lblDoB
        '
        Me.lblDoB.AutoSize = True
        Me.lblDoB.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!)
        Me.lblDoB.Location = New System.Drawing.Point(25, 169)
        Me.lblDoB.Name = "lblDoB"
        Me.lblDoB.Size = New System.Drawing.Size(106, 20)
        Me.lblDoB.TabIndex = 14
        Me.lblDoB.Text = "Date Of Birth:"
        '
        'lblAddress
        '
        Me.lblAddress.AutoSize = True
        Me.lblAddress.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!)
        Me.lblAddress.Location = New System.Drawing.Point(26, 205)
        Me.lblAddress.Name = "lblAddress"
        Me.lblAddress.Size = New System.Drawing.Size(72, 20)
        Me.lblAddress.TabIndex = 15
        Me.lblAddress.Text = "Address:"
        '
        'lblID
        '
        Me.lblID.AutoSize = True
        Me.lblID.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!)
        Me.lblID.Location = New System.Drawing.Point(24, 237)
        Me.lblID.Name = "lblID"
        Me.lblID.Size = New System.Drawing.Size(87, 20)
        Me.lblID.TabIndex = 16
        Me.lblID.Text = "Session ID"
        '
        'txtPhone
        '
        Me.txtPhone.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!)
        Me.txtPhone.Location = New System.Drawing.Point(342, 130)
        Me.txtPhone.Name = "txtPhone"
        Me.txtPhone.Size = New System.Drawing.Size(200, 26)
        Me.txtPhone.TabIndex = 17
        '
        'dtpDoB
        '
        Me.dtpDoB.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!)
        Me.dtpDoB.Location = New System.Drawing.Point(342, 164)
        Me.dtpDoB.Name = "dtpDoB"
        Me.dtpDoB.Size = New System.Drawing.Size(200, 26)
        Me.dtpDoB.TabIndex = 18
        '
        'txtAddress
        '
        Me.txtAddress.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!)
        Me.txtAddress.Location = New System.Drawing.Point(342, 202)
        Me.txtAddress.Name = "txtAddress"
        Me.txtAddress.Size = New System.Drawing.Size(200, 26)
        Me.txtAddress.TabIndex = 19
        '
        'txtSession
        '
        Me.txtSession.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!)
        Me.txtSession.Location = New System.Drawing.Point(342, 234)
        Me.txtSession.Name = "txtSession"
        Me.txtSession.ReadOnly = True
        Me.txtSession.Size = New System.Drawing.Size(200, 26)
        Me.txtSession.TabIndex = 20
        '
        'btnSearch
        '
        Me.btnSearch.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSearch.Location = New System.Drawing.Point(30, 384)
        Me.btnSearch.Name = "btnSearch"
        Me.btnSearch.Size = New System.Drawing.Size(514, 84)
        Me.btnSearch.TabIndex = 21
        Me.btnSearch.Text = "Search And Recieve"
        Me.btnSearch.UseVisualStyleBackColor = True
        '
        'lblSearchID
        '
        Me.lblSearchID.AutoSize = True
        Me.lblSearchID.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSearchID.Location = New System.Drawing.Point(26, 475)
        Me.lblSearchID.Name = "lblSearchID"
        Me.lblSearchID.Size = New System.Drawing.Size(312, 20)
        Me.lblSearchID.TabIndex = 23
        Me.lblSearchID.Text = "Please Enter The ID you want to search for"
        '
        'cmbID
        '
        Me.cmbID.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbID.FormattingEnabled = True
        Me.cmbID.Location = New System.Drawing.Point(359, 474)
        Me.cmbID.Name = "cmbID"
        Me.cmbID.Size = New System.Drawing.Size(183, 21)
        Me.cmbID.TabIndex = 24
        '
        'btnSwapFormActivities
        '
        Me.btnSwapFormActivities.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSwapFormActivities.Location = New System.Drawing.Point(28, 521)
        Me.btnSwapFormActivities.Name = "btnSwapFormActivities"
        Me.btnSwapFormActivities.Size = New System.Drawing.Size(514, 84)
        Me.btnSwapFormActivities.TabIndex = 25
        Me.btnSwapFormActivities.Text = "Swap To Activities Form"
        Me.btnSwapFormActivities.UseVisualStyleBackColor = True
        '
        'Customers
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(584, 615)
        Me.Controls.Add(Me.btnSwapFormActivities)
        Me.Controls.Add(Me.cmbID)
        Me.Controls.Add(Me.lblSearchID)
        Me.Controls.Add(Me.btnSearch)
        Me.Controls.Add(Me.txtSession)
        Me.Controls.Add(Me.txtAddress)
        Me.Controls.Add(Me.dtpDoB)
        Me.Controls.Add(Me.txtPhone)
        Me.Controls.Add(Me.lblID)
        Me.Controls.Add(Me.lblAddress)
        Me.Controls.Add(Me.lblDoB)
        Me.Controls.Add(Me.lblPhone)
        Me.Controls.Add(Me.txtEmail)
        Me.Controls.Add(Me.txtLastName)
        Me.Controls.Add(Me.txtFirstName)
        Me.Controls.Add(Me.lblEmail)
        Me.Controls.Add(Me.lblLastName)
        Me.Controls.Add(Me.lblName)
        Me.Controls.Add(Me.btnInput)
        Me.Name = "Customers"
        Me.Text = "Customers"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnInput As System.Windows.Forms.Button
    Friend WithEvents lblName As System.Windows.Forms.Label
    Friend WithEvents lblLastName As System.Windows.Forms.Label
    Friend WithEvents lblEmail As System.Windows.Forms.Label
    Friend WithEvents txtFirstName As System.Windows.Forms.TextBox
    Friend WithEvents txtLastName As System.Windows.Forms.TextBox
    Friend WithEvents txtEmail As System.Windows.Forms.TextBox
    Friend WithEvents lblPhone As System.Windows.Forms.Label
    Friend WithEvents lblDoB As System.Windows.Forms.Label
    Friend WithEvents lblAddress As System.Windows.Forms.Label
    Friend WithEvents lblID As System.Windows.Forms.Label
    Friend WithEvents txtPhone As System.Windows.Forms.TextBox
    Friend WithEvents dtpDoB As System.Windows.Forms.DateTimePicker
    Friend WithEvents txtAddress As System.Windows.Forms.TextBox
    Friend WithEvents txtSession As System.Windows.Forms.TextBox
    Friend WithEvents btnSearch As System.Windows.Forms.Button
    Friend WithEvents lblSearchID As System.Windows.Forms.Label
    Friend WithEvents cmbID As System.Windows.Forms.ComboBox
    Friend WithEvents btnSwapFormActivities As System.Windows.Forms.Button
End Class
