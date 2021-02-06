<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmCustomer
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmCustomer))
        Me.Label1 = New System.Windows.Forms.Label()
        Me.B_name = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.CheckBox1 = New System.Windows.Forms.CheckBox()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.B_Address = New System.Windows.Forms.TextBox()
        Me.B_Landmark = New System.Windows.Forms.TextBox()
        Me.B_City = New System.Windows.Forms.TextBox()
        Me.B_State = New System.Windows.Forms.ComboBox()
        Me.S_State = New System.Windows.Forms.ComboBox()
        Me.S_City = New System.Windows.Forms.TextBox()
        Me.S_Landmark = New System.Windows.Forms.TextBox()
        Me.S_Address = New System.Windows.Forms.TextBox()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.Label17 = New System.Windows.Forms.Label()
        Me.S_name = New System.Windows.Forms.TextBox()
        Me.Label18 = New System.Windows.Forms.Label()
        Me.Label19 = New System.Windows.Forms.Label()
        Me.Label20 = New System.Windows.Forms.Label()
        Me.Label21 = New System.Windows.Forms.Label()
        Me.Label22 = New System.Windows.Forms.Label()
        Me.Label23 = New System.Windows.Forms.Label()
        Me.txtCustomerNo = New System.Windows.Forms.TextBox()
        Me.txtEmail = New System.Windows.Forms.TextBox()
        Me.txtNotes = New System.Windows.Forms.TextBox()
        Me.txtMobileNo = New System.Windows.Forms.MaskedTextBox()
        Me.B_ZipCode = New System.Windows.Forms.MaskedTextBox()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.Delete = New System.Windows.Forms.Button()
        Me.Update_Record = New System.Windows.Forms.Button()
        Me.Save = New System.Windows.Forms.Button()
        Me.NewRecord = New System.Windows.Forms.Button()
        Me.S_ZipCode = New System.Windows.Forms.MaskedTextBox()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.txtPhone = New System.Windows.Forms.TextBox()
        Me.txtFaxNo = New System.Windows.Forms.TextBox()
        Me.GroupBox2.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.BackColor = System.Drawing.Color.LightGray
        Me.Label1.Font = New System.Drawing.Font("Palatino Linotype", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(43, 37)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(116, 22)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Billing Details"
        '
        'B_name
        '
        Me.B_name.Location = New System.Drawing.Point(173, 73)
        Me.B_name.Name = "B_name"
        Me.B_name.Size = New System.Drawing.Size(273, 20)
        Me.B_name.TabIndex = 0
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Palatino Linotype", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(44, 74)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(46, 17)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "*Name"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Palatino Linotype", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(44, 110)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(57, 17)
        Me.Label5.TabIndex = 5
        Me.Label5.Text = "*Address"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("Palatino Linotype", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(44, 177)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(36, 17)
        Me.Label6.TabIndex = 6
        Me.Label6.Text = "*City"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Font = New System.Drawing.Font("Palatino Linotype", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.Location = New System.Drawing.Point(44, 144)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(65, 17)
        Me.Label7.TabIndex = 7
        Me.Label7.Text = "Landmark"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Font = New System.Drawing.Font("Palatino Linotype", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.Location = New System.Drawing.Point(44, 211)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(40, 17)
        Me.Label8.TabIndex = 8
        Me.Label8.Text = "*State"
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Font = New System.Drawing.Font("Palatino Linotype", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label9.Location = New System.Drawing.Point(44, 246)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(91, 17)
        Me.Label9.TabIndex = 9
        Me.Label9.Text = "*Zip/Post Code"
        '
        'CheckBox1
        '
        Me.CheckBox1.AutoSize = True
        Me.CheckBox1.Location = New System.Drawing.Point(657, 41)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.Size = New System.Drawing.Size(138, 17)
        Me.CheckBox1.TabIndex = 7
        Me.CheckBox1.Text = "Same as Billing Address"
        Me.CheckBox1.UseVisualStyleBackColor = True
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.BackColor = System.Drawing.Color.LightGray
        Me.Label11.Font = New System.Drawing.Font("Palatino Linotype", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label11.Location = New System.Drawing.Point(509, 37)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(135, 22)
        Me.Label11.TabIndex = 12
        Me.Label11.Text = "Shipping Details"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Palatino Linotype", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(510, 344)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(86, 17)
        Me.Label4.TabIndex = 13
        Me.Label4.Text = "Distributor ID"
        '
        'B_Address
        '
        Me.B_Address.Location = New System.Drawing.Point(173, 107)
        Me.B_Address.Name = "B_Address"
        Me.B_Address.Size = New System.Drawing.Size(273, 20)
        Me.B_Address.TabIndex = 2
        '
        'B_Landmark
        '
        Me.B_Landmark.Location = New System.Drawing.Point(173, 141)
        Me.B_Landmark.Name = "B_Landmark"
        Me.B_Landmark.Size = New System.Drawing.Size(273, 20)
        Me.B_Landmark.TabIndex = 3
        '
        'B_City
        '
        Me.B_City.Location = New System.Drawing.Point(173, 174)
        Me.B_City.Name = "B_City"
        Me.B_City.Size = New System.Drawing.Size(144, 20)
        Me.B_City.TabIndex = 4
        '
        'B_State
        '
        Me.B_State.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest
        Me.B_State.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems
        Me.B_State.FormattingEnabled = True
        Me.B_State.Items.AddRange(New Object() {"Andaman & Nicobar", "Andhra Pradesh", "Arunachal Pradesh", "Assam", "Bihar", "Chandigarh", "Chhattisgarh", "Goa", "Gujarat", "Haryana", "Himachal Pradesh", "Jammu and Kashmir", "Jharkhand", "Karnataka", "Kerala", "Madhya Pradesh", "Maharashtra", "Manipur", "Meghalaya", "Mizoram", "Nagaland", "New Delhi", "Orissa", "Puducherry", "Rajasthan", "Sikkim", "Tamil Nadu", "Tripura", "Uttar Pradesh", "Uttarakhand", "West Bengal", "Dadra and Nagar Haveli", "Daman and Diu", "Lakshadweep"})
        Me.B_State.Location = New System.Drawing.Point(173, 207)
        Me.B_State.Name = "B_State"
        Me.B_State.Size = New System.Drawing.Size(192, 21)
        Me.B_State.TabIndex = 5
        '
        'S_State
        '
        Me.S_State.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest
        Me.S_State.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems
        Me.S_State.FormattingEnabled = True
        Me.S_State.Items.AddRange(New Object() {"Andaman & Nicobar", "Andhra Pradesh", "Arunachal Pradesh", "Assam", "Bihar", "Chandigarh", "Chhattisgarh", "Goa", "Gujarat", "Haryana", "Himachal Pradesh", "Jammu and Kashmir", "Jharkhand", "Karnataka", "Kerala", "Madhya Pradesh", "Maharashtra", "Manipur", "Meghalaya", "Mizoram", "Nagaland", "New Delhi", "Orissa", "Puducherry", "Rajasthan", "Sikkim", "Tamil Nadu", "Tripura", "Uttar Pradesh", "Uttarakhand", "West Bengal", "Dadra and Nagar Haveli", "Daman and Diu", "Lakshadweep"})
        Me.S_State.Location = New System.Drawing.Point(640, 207)
        Me.S_State.Name = "S_State"
        Me.S_State.Size = New System.Drawing.Size(201, 21)
        Me.S_State.TabIndex = 13
        '
        'S_City
        '
        Me.S_City.Location = New System.Drawing.Point(640, 174)
        Me.S_City.Name = "S_City"
        Me.S_City.Size = New System.Drawing.Size(144, 20)
        Me.S_City.TabIndex = 12
        '
        'S_Landmark
        '
        Me.S_Landmark.Location = New System.Drawing.Point(640, 141)
        Me.S_Landmark.Name = "S_Landmark"
        Me.S_Landmark.Size = New System.Drawing.Size(273, 20)
        Me.S_Landmark.TabIndex = 11
        '
        'S_Address
        '
        Me.S_Address.Location = New System.Drawing.Point(640, 107)
        Me.S_Address.Name = "S_Address"
        Me.S_Address.Size = New System.Drawing.Size(273, 20)
        Me.S_Address.TabIndex = 10
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Font = New System.Drawing.Font("Palatino Linotype", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label10.Location = New System.Drawing.Point(511, 246)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(91, 17)
        Me.Label10.TabIndex = 27
        Me.Label10.Text = "*Zip/Post Code"
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Font = New System.Drawing.Font("Palatino Linotype", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label12.Location = New System.Drawing.Point(511, 211)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(40, 17)
        Me.Label12.TabIndex = 26
        Me.Label12.Text = "*State"
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Font = New System.Drawing.Font("Palatino Linotype", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label13.Location = New System.Drawing.Point(511, 144)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(65, 17)
        Me.Label13.TabIndex = 25
        Me.Label13.Text = "Landmark"
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Font = New System.Drawing.Font("Palatino Linotype", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label14.Location = New System.Drawing.Point(511, 177)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(36, 17)
        Me.Label14.TabIndex = 24
        Me.Label14.Text = "*City"
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.Font = New System.Drawing.Font("Palatino Linotype", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label15.Location = New System.Drawing.Point(511, 110)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(57, 17)
        Me.Label15.TabIndex = 23
        Me.Label15.Text = "*Address"
        '
        'Label17
        '
        Me.Label17.AutoSize = True
        Me.Label17.Font = New System.Drawing.Font("Palatino Linotype", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label17.Location = New System.Drawing.Point(511, 74)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(46, 17)
        Me.Label17.TabIndex = 21
        Me.Label17.Text = "*Name"
        '
        'S_name
        '
        Me.S_name.Location = New System.Drawing.Point(640, 73)
        Me.S_name.Name = "S_name"
        Me.S_name.Size = New System.Drawing.Size(273, 20)
        Me.S_name.TabIndex = 8
        '
        'Label18
        '
        Me.Label18.AutoSize = True
        Me.Label18.BackColor = System.Drawing.Color.LightGray
        Me.Label18.Font = New System.Drawing.Font("Palatino Linotype", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label18.Location = New System.Drawing.Point(510, 304)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(146, 22)
        Me.Label18.TabIndex = 34
        Me.Label18.Text = "Additional Details"
        '
        'Label19
        '
        Me.Label19.AutoSize = True
        Me.Label19.Font = New System.Drawing.Font("Palatino Linotype", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label19.Location = New System.Drawing.Point(510, 380)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(43, 17)
        Me.Label19.TabIndex = 35
        Me.Label19.Text = "Phone"
        '
        'Label20
        '
        Me.Label20.AutoSize = True
        Me.Label20.Font = New System.Drawing.Font("Palatino Linotype", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label20.Location = New System.Drawing.Point(510, 414)
        Me.Label20.Name = "Label20"
        Me.Label20.Size = New System.Drawing.Size(38, 17)
        Me.Label20.TabIndex = 36
        Me.Label20.Text = "Email"
        '
        'Label21
        '
        Me.Label21.AutoSize = True
        Me.Label21.Font = New System.Drawing.Font("Palatino Linotype", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label21.Location = New System.Drawing.Point(508, 452)
        Me.Label21.Name = "Label21"
        Me.Label21.Size = New System.Drawing.Size(74, 17)
        Me.Label21.TabIndex = 37
        Me.Label21.Text = "*Mobile No."
        '
        'Label22
        '
        Me.Label22.AutoSize = True
        Me.Label22.Font = New System.Drawing.Font("Palatino Linotype", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label22.Location = New System.Drawing.Point(510, 490)
        Me.Label22.Name = "Label22"
        Me.Label22.Size = New System.Drawing.Size(49, 17)
        Me.Label22.TabIndex = 38
        Me.Label22.Text = "Fax No."
        '
        'Label23
        '
        Me.Label23.AutoSize = True
        Me.Label23.Font = New System.Drawing.Font("Palatino Linotype", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label23.Location = New System.Drawing.Point(511, 524)
        Me.Label23.Name = "Label23"
        Me.Label23.Size = New System.Drawing.Size(40, 17)
        Me.Label23.TabIndex = 39
        Me.Label23.Text = "Notes"
        '
        'txtCustomerNo
        '
        Me.txtCustomerNo.Location = New System.Drawing.Point(640, 344)
        Me.txtCustomerNo.Name = "txtCustomerNo"
        Me.txtCustomerNo.ReadOnly = True
        Me.txtCustomerNo.Size = New System.Drawing.Size(108, 20)
        Me.txtCustomerNo.TabIndex = 15
        '
        'txtEmail
        '
        Me.txtEmail.Location = New System.Drawing.Point(640, 414)
        Me.txtEmail.Name = "txtEmail"
        Me.txtEmail.Size = New System.Drawing.Size(226, 20)
        Me.txtEmail.TabIndex = 17
        '
        'txtNotes
        '
        Me.txtNotes.Location = New System.Drawing.Point(640, 524)
        Me.txtNotes.Multiline = True
        Me.txtNotes.Name = "txtNotes"
        Me.txtNotes.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtNotes.Size = New System.Drawing.Size(251, 67)
        Me.txtNotes.TabIndex = 20
        '
        'txtMobileNo
        '
        Me.txtMobileNo.Location = New System.Drawing.Point(640, 449)
        Me.txtMobileNo.Mask = "0000000000"
        Me.txtMobileNo.Name = "txtMobileNo"
        Me.txtMobileNo.Size = New System.Drawing.Size(100, 20)
        Me.txtMobileNo.TabIndex = 18
        '
        'B_ZipCode
        '
        Me.B_ZipCode.Location = New System.Drawing.Point(173, 243)
        Me.B_ZipCode.Mask = "000000"
        Me.B_ZipCode.Name = "B_ZipCode"
        Me.B_ZipCode.Size = New System.Drawing.Size(74, 20)
        Me.B_ZipCode.TabIndex = 6
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.Delete)
        Me.GroupBox2.Controls.Add(Me.Update_Record)
        Me.GroupBox2.Controls.Add(Me.Save)
        Me.GroupBox2.Controls.Add(Me.NewRecord)
        Me.GroupBox2.Location = New System.Drawing.Point(959, 32)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(134, 178)
        Me.GroupBox2.TabIndex = 21
        Me.GroupBox2.TabStop = False
        '
        'Delete
        '
        Me.Delete.Enabled = False
        Me.Delete.Font = New System.Drawing.Font("Palatino Linotype", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Delete.Location = New System.Drawing.Point(16, 133)
        Me.Delete.Name = "Delete"
        Me.Delete.Size = New System.Drawing.Size(99, 29)
        Me.Delete.TabIndex = 3
        Me.Delete.Text = "&Delete"
        Me.Delete.UseVisualStyleBackColor = True
        '
        'Update_Record
        '
        Me.Update_Record.Enabled = False
        Me.Update_Record.Font = New System.Drawing.Font("Palatino Linotype", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Update_Record.Location = New System.Drawing.Point(16, 98)
        Me.Update_Record.Name = "Update_Record"
        Me.Update_Record.Size = New System.Drawing.Size(99, 29)
        Me.Update_Record.TabIndex = 2
        Me.Update_Record.Text = "&Update"
        Me.Update_Record.UseVisualStyleBackColor = True
        '
        'Save
        '
        Me.Save.Font = New System.Drawing.Font("Palatino Linotype", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Save.Location = New System.Drawing.Point(16, 63)
        Me.Save.Name = "Save"
        Me.Save.Size = New System.Drawing.Size(99, 29)
        Me.Save.TabIndex = 1
        Me.Save.Text = "&Save"
        Me.Save.UseVisualStyleBackColor = True
        '
        'NewRecord
        '
        Me.NewRecord.Font = New System.Drawing.Font("Palatino Linotype", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.NewRecord.Location = New System.Drawing.Point(16, 28)
        Me.NewRecord.Name = "NewRecord"
        Me.NewRecord.Size = New System.Drawing.Size(99, 29)
        Me.NewRecord.TabIndex = 0
        Me.NewRecord.Text = "&New"
        Me.NewRecord.UseVisualStyleBackColor = True
        '
        'S_ZipCode
        '
        Me.S_ZipCode.Location = New System.Drawing.Point(640, 243)
        Me.S_ZipCode.Mask = "000000"
        Me.S_ZipCode.Name = "S_ZipCode"
        Me.S_ZipCode.Size = New System.Drawing.Size(74, 20)
        Me.S_ZipCode.TabIndex = 14
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(766, 342)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(29, 21)
        Me.Button1.TabIndex = 51
        Me.Button1.Text = "<"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'txtPhone
        '
        Me.txtPhone.Location = New System.Drawing.Point(640, 380)
        Me.txtPhone.Name = "txtPhone"
        Me.txtPhone.Size = New System.Drawing.Size(119, 20)
        Me.txtPhone.TabIndex = 16
        '
        'txtFaxNo
        '
        Me.txtFaxNo.Location = New System.Drawing.Point(640, 487)
        Me.txtFaxNo.Name = "txtFaxNo"
        Me.txtFaxNo.Size = New System.Drawing.Size(119, 20)
        Me.txtFaxNo.TabIndex = 19
        '
        'frmCustomer
        '
        Me.AcceptButton = Me.Save
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.ClientSize = New System.Drawing.Size(1108, 605)
        Me.Controls.Add(Me.txtFaxNo)
        Me.Controls.Add(Me.txtPhone)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.S_ZipCode)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.B_ZipCode)
        Me.Controls.Add(Me.txtMobileNo)
        Me.Controls.Add(Me.txtNotes)
        Me.Controls.Add(Me.txtEmail)
        Me.Controls.Add(Me.txtCustomerNo)
        Me.Controls.Add(Me.Label23)
        Me.Controls.Add(Me.Label22)
        Me.Controls.Add(Me.Label21)
        Me.Controls.Add(Me.Label20)
        Me.Controls.Add(Me.Label19)
        Me.Controls.Add(Me.Label18)
        Me.Controls.Add(Me.S_State)
        Me.Controls.Add(Me.S_City)
        Me.Controls.Add(Me.S_Landmark)
        Me.Controls.Add(Me.S_Address)
        Me.Controls.Add(Me.Label10)
        Me.Controls.Add(Me.Label12)
        Me.Controls.Add(Me.Label13)
        Me.Controls.Add(Me.Label14)
        Me.Controls.Add(Me.Label15)
        Me.Controls.Add(Me.Label17)
        Me.Controls.Add(Me.S_name)
        Me.Controls.Add(Me.B_State)
        Me.Controls.Add(Me.B_City)
        Me.Controls.Add(Me.B_Landmark)
        Me.Controls.Add(Me.B_Address)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label11)
        Me.Controls.Add(Me.CheckBox1)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.B_name)
        Me.Controls.Add(Me.Label1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Name = "frmCustomer"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Super Distributor Details"
        Me.GroupBox2.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents B_name As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents CheckBox1 As System.Windows.Forms.CheckBox
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents B_Address As System.Windows.Forms.TextBox
    Friend WithEvents B_Landmark As System.Windows.Forms.TextBox
    Friend WithEvents B_City As System.Windows.Forms.TextBox
    Friend WithEvents B_State As System.Windows.Forms.ComboBox
    Friend WithEvents S_State As System.Windows.Forms.ComboBox
    Friend WithEvents S_City As System.Windows.Forms.TextBox
    Friend WithEvents S_Landmark As System.Windows.Forms.TextBox
    Friend WithEvents S_Address As System.Windows.Forms.TextBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents S_name As System.Windows.Forms.TextBox
    Friend WithEvents Label18 As System.Windows.Forms.Label
    Friend WithEvents Label19 As System.Windows.Forms.Label
    Friend WithEvents Label20 As System.Windows.Forms.Label
    Friend WithEvents Label21 As System.Windows.Forms.Label
    Friend WithEvents Label22 As System.Windows.Forms.Label
    Friend WithEvents Label23 As System.Windows.Forms.Label
    Friend WithEvents txtCustomerNo As System.Windows.Forms.TextBox
    Friend WithEvents txtEmail As System.Windows.Forms.TextBox
    Friend WithEvents txtNotes As System.Windows.Forms.TextBox
    Friend WithEvents txtMobileNo As System.Windows.Forms.MaskedTextBox
    Friend WithEvents B_ZipCode As System.Windows.Forms.MaskedTextBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents Delete As System.Windows.Forms.Button
    Friend WithEvents Update_Record As System.Windows.Forms.Button
    Friend WithEvents Save As System.Windows.Forms.Button
    Friend WithEvents NewRecord As System.Windows.Forms.Button
    Friend WithEvents S_ZipCode As System.Windows.Forms.MaskedTextBox
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents txtPhone As System.Windows.Forms.TextBox
    Friend WithEvents txtFaxNo As System.Windows.Forms.TextBox
End Class
