<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmSales
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
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmSales))
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.btnUpdate = New System.Windows.Forms.Button()
        Me.NewRecord = New System.Windows.Forms.Button()
        Me.Delete = New System.Windows.Forms.Button()
        Me.Save = New System.Windows.Forms.Button()
        Me.Button4 = New System.Windows.Forms.Button()
        Me.txtCartons = New System.Windows.Forms.TextBox()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.txtWeight = New System.Windows.Forms.TextBox()
        Me.txtAvailableCartons = New System.Windows.Forms.TextBox()
        Me.txtPrice = New System.Windows.Forms.TextBox()
        Me.txtCustomerNo = New System.Windows.Forms.TextBox()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.txtProductName = New System.Windows.Forms.TextBox()
        Me.txtProductCode = New System.Windows.Forms.TextBox()
        Me.dtpTransactionDate = New System.Windows.Forms.DateTimePicker()
        Me.txtInvoiceNo = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.txtTotalAmount = New System.Windows.Forms.TextBox()
        Me.txtPackets = New System.Windows.Forms.TextBox()
        Me.txtAvailablePackets = New System.Windows.Forms.TextBox()
        Me.Label18 = New System.Windows.Forms.Label()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Button7 = New System.Windows.Forms.Button()
        Me.btnPrint = New System.Windows.Forms.Button()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.txtPaymentDue = New System.Windows.Forms.TextBox()
        Me.txtTotalPayment = New System.Windows.Forms.TextBox()
        Me.Label19 = New System.Windows.Forms.Label()
        Me.payment = New System.Windows.Forms.Label()
        Me.txtTotal = New System.Windows.Forms.TextBox()
        Me.Label16 = New System.Windows.Forms.Label()
        Me.txtTaxAmt = New System.Windows.Forms.TextBox()
        Me.Label24 = New System.Windows.Forms.Label()
        Me.txtTaxPer = New System.Windows.Forms.TextBox()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.txtSubTotal = New System.Windows.Forms.TextBox()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.txtCustomerName = New System.Windows.Forms.TextBox()
        Me.ListView1 = New System.Windows.Forms.ListView()
        Me.ColumnHeader1 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ProductCode = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader2 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader3 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader7 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader4 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader5 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader6 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.btnRemove = New System.Windows.Forms.Button()
        Me.Label17 = New System.Windows.Forms.Label()
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.txtOrderNo = New System.Windows.Forms.TextBox()
        Me.BtnSubmit = New System.Windows.Forms.Button()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.lblCartons = New System.Windows.Forms.Label()
        Me.ColumnHeader8 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.Panel1.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.Controls.Add(Me.btnUpdate)
        Me.Panel1.Controls.Add(Me.NewRecord)
        Me.Panel1.Controls.Add(Me.Delete)
        Me.Panel1.Controls.Add(Me.Save)
        Me.Panel1.Location = New System.Drawing.Point(831, 58)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(127, 152)
        Me.Panel1.TabIndex = 10
        '
        'btnUpdate
        '
        Me.btnUpdate.Enabled = False
        Me.btnUpdate.Font = New System.Drawing.Font("Palatino Linotype", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnUpdate.Location = New System.Drawing.Point(14, 113)
        Me.btnUpdate.Name = "btnUpdate"
        Me.btnUpdate.Size = New System.Drawing.Size(99, 29)
        Me.btnUpdate.TabIndex = 100
        Me.btnUpdate.Text = "&Update"
        Me.btnUpdate.UseVisualStyleBackColor = True
        '
        'NewRecord
        '
        Me.NewRecord.Font = New System.Drawing.Font("Palatino Linotype", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.NewRecord.Location = New System.Drawing.Point(14, 8)
        Me.NewRecord.Name = "NewRecord"
        Me.NewRecord.Size = New System.Drawing.Size(99, 29)
        Me.NewRecord.TabIndex = 0
        Me.NewRecord.Text = "&New"
        Me.NewRecord.UseVisualStyleBackColor = True
        '
        'Delete
        '
        Me.Delete.Enabled = False
        Me.Delete.Font = New System.Drawing.Font("Palatino Linotype", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Delete.Location = New System.Drawing.Point(14, 78)
        Me.Delete.Name = "Delete"
        Me.Delete.Size = New System.Drawing.Size(99, 29)
        Me.Delete.TabIndex = 3
        Me.Delete.Text = "&Delete"
        Me.Delete.UseVisualStyleBackColor = True
        '
        'Save
        '
        Me.Save.Font = New System.Drawing.Font("Palatino Linotype", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Save.Location = New System.Drawing.Point(14, 43)
        Me.Save.Name = "Save"
        Me.Save.Size = New System.Drawing.Size(99, 29)
        Me.Save.TabIndex = 1
        Me.Save.Text = "&Save"
        Me.Save.UseVisualStyleBackColor = True
        '
        'Button4
        '
        Me.Button4.Location = New System.Drawing.Point(355, 56)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(29, 21)
        Me.Button4.TabIndex = 90
        Me.Button4.Text = "<"
        Me.Button4.UseVisualStyleBackColor = True
        '
        'txtCartons
        '
        Me.txtCartons.Location = New System.Drawing.Point(159, 157)
        Me.txtCartons.Name = "txtCartons"
        Me.txtCartons.Size = New System.Drawing.Size(89, 20)
        Me.txtCartons.TabIndex = 4
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Font = New System.Drawing.Font("Palatino Linotype", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label12.Location = New System.Drawing.Point(280, 89)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(61, 17)
        Me.Label12.TabIndex = 87
        Me.Label12.Text = "Unit Price"
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(355, 122)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(29, 21)
        Me.Button1.TabIndex = 86
        Me.Button1.Text = "<"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(369, 24)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(29, 21)
        Me.Button2.TabIndex = 85
        Me.Button2.Text = "<"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'txtWeight
        '
        Me.txtWeight.Location = New System.Drawing.Point(159, 88)
        Me.txtWeight.Name = "txtWeight"
        Me.txtWeight.ReadOnly = True
        Me.txtWeight.Size = New System.Drawing.Size(77, 20)
        Me.txtWeight.TabIndex = 2
        '
        'txtAvailableCartons
        '
        Me.txtAvailableCartons.Location = New System.Drawing.Point(159, 123)
        Me.txtAvailableCartons.Name = "txtAvailableCartons"
        Me.txtAvailableCartons.ReadOnly = True
        Me.txtAvailableCartons.Size = New System.Drawing.Size(89, 20)
        Me.txtAvailableCartons.TabIndex = 5
        '
        'txtPrice
        '
        Me.txtPrice.Location = New System.Drawing.Point(415, 87)
        Me.txtPrice.Name = "txtPrice"
        Me.txtPrice.ReadOnly = True
        Me.txtPrice.Size = New System.Drawing.Size(89, 20)
        Me.txtPrice.TabIndex = 3
        '
        'txtCustomerNo
        '
        Me.txtCustomerNo.Location = New System.Drawing.Point(152, 123)
        Me.txtCustomerNo.Name = "txtCustomerNo"
        Me.txtCustomerNo.ReadOnly = True
        Me.txtCustomerNo.Size = New System.Drawing.Size(185, 20)
        Me.txtCustomerNo.TabIndex = 2
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Font = New System.Drawing.Font("Palatino Linotype", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label10.Location = New System.Drawing.Point(28, 195)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(85, 17)
        Me.Label10.TabIndex = 76
        Me.Label10.Text = "Total Amount"
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Font = New System.Drawing.Font("Palatino Linotype", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label9.Location = New System.Drawing.Point(26, 125)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(103, 17)
        Me.Label9.TabIndex = 75
        Me.Label9.Text = "Available Cartons"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Palatino Linotype", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(26, 57)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(87, 17)
        Me.Label5.TabIndex = 74
        Me.Label5.Text = "Product Name"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("Palatino Linotype", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(28, 28)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(83, 17)
        Me.Label6.TabIndex = 73
        Me.Label6.Text = "Product Code"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Font = New System.Drawing.Font("Palatino Linotype", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.Location = New System.Drawing.Point(26, 89)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(100, 17)
        Me.Label7.TabIndex = 72
        Me.Label7.Text = "Weight (Per Qty)"
        '
        'txtProductName
        '
        Me.txtProductName.Location = New System.Drawing.Point(159, 57)
        Me.txtProductName.Name = "txtProductName"
        Me.txtProductName.ReadOnly = True
        Me.txtProductName.Size = New System.Drawing.Size(366, 20)
        Me.txtProductName.TabIndex = 1
        '
        'txtProductCode
        '
        Me.txtProductCode.Location = New System.Drawing.Point(159, 25)
        Me.txtProductCode.Name = "txtProductCode"
        Me.txtProductCode.ReadOnly = True
        Me.txtProductCode.Size = New System.Drawing.Size(185, 20)
        Me.txtProductCode.TabIndex = 0
        '
        'dtpTransactionDate
        '
        Me.dtpTransactionDate.CustomFormat = "dd/MMM/yyyy"
        Me.dtpTransactionDate.Font = New System.Drawing.Font("Palatino Linotype", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtpTransactionDate.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpTransactionDate.Location = New System.Drawing.Point(152, 88)
        Me.dtpTransactionDate.Name = "dtpTransactionDate"
        Me.dtpTransactionDate.Size = New System.Drawing.Size(120, 24)
        Me.dtpTransactionDate.TabIndex = 1
        '
        'txtInvoiceNo
        '
        Me.txtInvoiceNo.Location = New System.Drawing.Point(152, 57)
        Me.txtInvoiceNo.Name = "txtInvoiceNo"
        Me.txtInvoiceNo.ReadOnly = True
        Me.txtInvoiceNo.Size = New System.Drawing.Size(185, 20)
        Me.txtInvoiceNo.TabIndex = 0
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Palatino Linotype", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(36, 92)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(77, 17)
        Me.Label3.TabIndex = 66
        Me.Label3.Text = "Invoice Date"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Palatino Linotype", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(36, 60)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(71, 17)
        Me.Label4.TabIndex = 65
        Me.Label4.Text = "Invoice No."
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Palatino Linotype", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(36, 157)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(104, 17)
        Me.Label2.TabIndex = 64
        Me.Label2.Text = "Distributor Name"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Palatino Linotype", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(36, 124)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(86, 17)
        Me.Label1.TabIndex = 63
        Me.Label1.Text = "Distributor ID"
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.BackColor = System.Drawing.Color.LightGray
        Me.Label13.Font = New System.Drawing.Font("Palatino Linotype", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label13.Location = New System.Drawing.Point(35, 16)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(60, 22)
        Me.Label13.TabIndex = 95
        Me.Label13.Text = "Billing"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.txtTotalAmount)
        Me.GroupBox1.Controls.Add(Me.txtPackets)
        Me.GroupBox1.Controls.Add(Me.txtAvailablePackets)
        Me.GroupBox1.Controls.Add(Me.Label18)
        Me.GroupBox1.Controls.Add(Me.Label11)
        Me.GroupBox1.Controls.Add(Me.Label8)
        Me.GroupBox1.Controls.Add(Me.Button7)
        Me.GroupBox1.Controls.Add(Me.Label6)
        Me.GroupBox1.Controls.Add(Me.txtProductCode)
        Me.GroupBox1.Controls.Add(Me.Label5)
        Me.GroupBox1.Controls.Add(Me.txtProductName)
        Me.GroupBox1.Controls.Add(Me.Button2)
        Me.GroupBox1.Controls.Add(Me.Label7)
        Me.GroupBox1.Controls.Add(Me.txtWeight)
        Me.GroupBox1.Controls.Add(Me.Label12)
        Me.GroupBox1.Controls.Add(Me.txtCartons)
        Me.GroupBox1.Controls.Add(Me.Label9)
        Me.GroupBox1.Controls.Add(Me.txtPrice)
        Me.GroupBox1.Controls.Add(Me.Label10)
        Me.GroupBox1.Controls.Add(Me.txtAvailableCartons)
        Me.GroupBox1.Location = New System.Drawing.Point(38, 192)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(601, 234)
        Me.GroupBox1.TabIndex = 9
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Product Details"
        '
        'txtTotalAmount
        '
        Me.txtTotalAmount.Location = New System.Drawing.Point(159, 193)
        Me.txtTotalAmount.Name = "txtTotalAmount"
        Me.txtTotalAmount.ReadOnly = True
        Me.txtTotalAmount.Size = New System.Drawing.Size(94, 20)
        Me.txtTotalAmount.TabIndex = 93
        '
        'txtPackets
        '
        Me.txtPackets.Location = New System.Drawing.Point(415, 157)
        Me.txtPackets.Name = "txtPackets"
        Me.txtPackets.ReadOnly = True
        Me.txtPackets.Size = New System.Drawing.Size(89, 20)
        Me.txtPackets.TabIndex = 92
        '
        'txtAvailablePackets
        '
        Me.txtAvailablePackets.Location = New System.Drawing.Point(415, 122)
        Me.txtAvailablePackets.Name = "txtAvailablePackets"
        Me.txtAvailablePackets.ReadOnly = True
        Me.txtAvailablePackets.Size = New System.Drawing.Size(89, 20)
        Me.txtAvailablePackets.TabIndex = 91
        '
        'Label18
        '
        Me.Label18.AutoSize = True
        Me.Label18.Font = New System.Drawing.Font("Palatino Linotype", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label18.Location = New System.Drawing.Point(282, 160)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(81, 17)
        Me.Label18.TabIndex = 90
        Me.Label18.Text = "Total Packets"
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Font = New System.Drawing.Font("Palatino Linotype", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label11.Location = New System.Drawing.Point(28, 160)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(51, 17)
        Me.Label11.TabIndex = 89
        Me.Label11.Text = "Cartons"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Font = New System.Drawing.Font("Palatino Linotype", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.Location = New System.Drawing.Point(280, 125)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(101, 17)
        Me.Label8.TabIndex = 88
        Me.Label8.Text = "Available Packets"
        '
        'Button7
        '
        Me.Button7.Font = New System.Drawing.Font("Palatino Linotype", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button7.Location = New System.Drawing.Point(497, 19)
        Me.Button7.Name = "Button7"
        Me.Button7.Size = New System.Drawing.Size(87, 23)
        Me.Button7.TabIndex = 6
        Me.Button7.Text = "&Add To Cart"
        Me.Button7.UseVisualStyleBackColor = True
        '
        'btnPrint
        '
        Me.btnPrint.Enabled = False
        Me.btnPrint.Font = New System.Drawing.Font("Palatino Linotype", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnPrint.Location = New System.Drawing.Point(1033, 630)
        Me.btnPrint.Name = "btnPrint"
        Me.btnPrint.Size = New System.Drawing.Size(99, 29)
        Me.btnPrint.TabIndex = 13
        Me.btnPrint.Text = "&Print Invoice"
        Me.btnPrint.UseVisualStyleBackColor = True
        '
        'Panel2
        '
        Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel2.Controls.Add(Me.txtPaymentDue)
        Me.Panel2.Controls.Add(Me.txtTotalPayment)
        Me.Panel2.Controls.Add(Me.Label19)
        Me.Panel2.Controls.Add(Me.payment)
        Me.Panel2.Controls.Add(Me.txtTotal)
        Me.Panel2.Controls.Add(Me.Label16)
        Me.Panel2.Controls.Add(Me.txtTaxAmt)
        Me.Panel2.Controls.Add(Me.Label24)
        Me.Panel2.Controls.Add(Me.txtTaxPer)
        Me.Panel2.Controls.Add(Me.Label15)
        Me.Panel2.Controls.Add(Me.txtSubTotal)
        Me.Panel2.Controls.Add(Me.Label14)
        Me.Panel2.Location = New System.Drawing.Point(934, 432)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(323, 189)
        Me.Panel2.TabIndex = 11
        '
        'txtPaymentDue
        '
        Me.txtPaymentDue.Location = New System.Drawing.Point(116, 150)
        Me.txtPaymentDue.Name = "txtPaymentDue"
        Me.txtPaymentDue.ReadOnly = True
        Me.txtPaymentDue.Size = New System.Drawing.Size(109, 20)
        Me.txtPaymentDue.TabIndex = 98
        '
        'txtTotalPayment
        '
        Me.txtTotalPayment.Location = New System.Drawing.Point(116, 116)
        Me.txtTotalPayment.Name = "txtTotalPayment"
        Me.txtTotalPayment.Size = New System.Drawing.Size(109, 20)
        Me.txtTotalPayment.TabIndex = 1
        '
        'Label19
        '
        Me.Label19.AutoSize = True
        Me.Label19.Font = New System.Drawing.Font("Palatino Linotype", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label19.Location = New System.Drawing.Point(18, 150)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(83, 17)
        Me.Label19.TabIndex = 96
        Me.Label19.Text = "Payment Due"
        '
        'payment
        '
        Me.payment.AutoSize = True
        Me.payment.Font = New System.Drawing.Font("Palatino Linotype", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.payment.Location = New System.Drawing.Point(18, 116)
        Me.payment.Name = "payment"
        Me.payment.Size = New System.Drawing.Size(89, 17)
        Me.payment.TabIndex = 95
        Me.payment.Text = "Total Payment"
        '
        'txtTotal
        '
        Me.txtTotal.Location = New System.Drawing.Point(116, 82)
        Me.txtTotal.Name = "txtTotal"
        Me.txtTotal.ReadOnly = True
        Me.txtTotal.Size = New System.Drawing.Size(111, 20)
        Me.txtTotal.TabIndex = 3
        '
        'Label16
        '
        Me.Label16.AutoSize = True
        Me.Label16.Font = New System.Drawing.Font("Palatino Linotype", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label16.Location = New System.Drawing.Point(19, 83)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(75, 17)
        Me.Label16.TabIndex = 94
        Me.Label16.Text = "Grand Total"
        '
        'txtTaxAmt
        '
        Me.txtTaxAmt.Location = New System.Drawing.Point(220, 47)
        Me.txtTaxAmt.Name = "txtTaxAmt"
        Me.txtTaxAmt.ReadOnly = True
        Me.txtTaxAmt.Size = New System.Drawing.Size(80, 20)
        Me.txtTaxAmt.TabIndex = 2
        '
        'Label24
        '
        Me.Label24.AutoSize = True
        Me.Label24.BackColor = System.Drawing.Color.White
        Me.Label24.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label24.ForeColor = System.Drawing.Color.Black
        Me.Label24.Location = New System.Drawing.Point(174, 47)
        Me.Label24.Name = "Label24"
        Me.Label24.Size = New System.Drawing.Size(23, 20)
        Me.Label24.TabIndex = 92
        Me.Label24.Text = "%"
        '
        'txtTaxPer
        '
        Me.txtTaxPer.Location = New System.Drawing.Point(116, 48)
        Me.txtTaxPer.Name = "txtTaxPer"
        Me.txtTaxPer.Size = New System.Drawing.Size(52, 20)
        Me.txtTaxPer.TabIndex = 0
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.Font = New System.Drawing.Font("Palatino Linotype", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label15.Location = New System.Drawing.Point(19, 50)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(27, 17)
        Me.Label15.TabIndex = 90
        Me.Label15.Text = "Tax"
        '
        'txtSubTotal
        '
        Me.txtSubTotal.Location = New System.Drawing.Point(116, 14)
        Me.txtSubTotal.Name = "txtSubTotal"
        Me.txtSubTotal.ReadOnly = True
        Me.txtSubTotal.Size = New System.Drawing.Size(184, 20)
        Me.txtSubTotal.TabIndex = 7
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Font = New System.Drawing.Font("Palatino Linotype", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label14.Location = New System.Drawing.Point(19, 16)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(61, 17)
        Me.Label14.TabIndex = 77
        Me.Label14.Text = "Sub Total"
        '
        'txtCustomerName
        '
        Me.txtCustomerName.Location = New System.Drawing.Point(152, 154)
        Me.txtCustomerName.Name = "txtCustomerName"
        Me.txtCustomerName.ReadOnly = True
        Me.txtCustomerName.Size = New System.Drawing.Size(353, 20)
        Me.txtCustomerName.TabIndex = 3
        '
        'ListView1
        '
        Me.ListView1.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader1, Me.ProductCode, Me.ColumnHeader2, Me.ColumnHeader3, Me.ColumnHeader7, Me.ColumnHeader4, Me.ColumnHeader8, Me.ColumnHeader5, Me.ColumnHeader6})
        Me.ListView1.Font = New System.Drawing.Font("Palatino Linotype", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ListView1.FullRowSelect = True
        Me.ListView1.GridLines = True
        Me.ListView1.Location = New System.Drawing.Point(38, 432)
        Me.ListView1.Name = "ListView1"
        Me.ListView1.Size = New System.Drawing.Size(884, 227)
        Me.ListView1.TabIndex = 96
        Me.ListView1.UseCompatibleStateImageBehavior = False
        Me.ListView1.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader1
        '
        Me.ColumnHeader1.Width = 0
        '
        'ProductCode
        '
        Me.ProductCode.Text = "Product Code"
        Me.ProductCode.Width = 90
        '
        'ColumnHeader2
        '
        Me.ColumnHeader2.Text = "Product Name"
        Me.ColumnHeader2.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.ColumnHeader2.Width = 250
        '
        'ColumnHeader3
        '
        Me.ColumnHeader3.Text = "Weight/Qty"
        Me.ColumnHeader3.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.ColumnHeader3.Width = 90
        '
        'ColumnHeader7
        '
        Me.ColumnHeader7.Text = "Unit Price"
        Me.ColumnHeader7.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.ColumnHeader7.Width = 80
        '
        'ColumnHeader4
        '
        Me.ColumnHeader4.Text = "Cartons"
        Me.ColumnHeader4.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.ColumnHeader4.Width = 85
        '
        'ColumnHeader5
        '
        Me.ColumnHeader5.Text = "Total Packets"
        Me.ColumnHeader5.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.ColumnHeader5.Width = 90
        '
        'ColumnHeader6
        '
        Me.ColumnHeader6.Text = "Total Amount"
        Me.ColumnHeader6.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.ColumnHeader6.Width = 92
        '
        'btnRemove
        '
        Me.btnRemove.Enabled = False
        Me.btnRemove.Font = New System.Drawing.Font("Palatino Linotype", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnRemove.Location = New System.Drawing.Point(934, 630)
        Me.btnRemove.Name = "btnRemove"
        Me.btnRemove.Size = New System.Drawing.Size(86, 29)
        Me.btnRemove.TabIndex = 12
        Me.btnRemove.Text = "&Remove"
        Me.btnRemove.UseVisualStyleBackColor = True
        '
        'Label17
        '
        Me.Label17.AutoSize = True
        Me.Label17.Font = New System.Drawing.Font("Palatino Linotype", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label17.Location = New System.Drawing.Point(434, 58)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(64, 17)
        Me.Label17.TabIndex = 98
        Me.Label17.Text = "Order No."
        '
        'Timer1
        '
        '
        'txtOrderNo
        '
        Me.txtOrderNo.Location = New System.Drawing.Point(505, 57)
        Me.txtOrderNo.Name = "txtOrderNo"
        Me.txtOrderNo.ReadOnly = True
        Me.txtOrderNo.Size = New System.Drawing.Size(144, 20)
        Me.txtOrderNo.TabIndex = 99
        '
        'BtnSubmit
        '
        Me.BtnSubmit.Font = New System.Drawing.Font("Palatino Linotype", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtnSubmit.Location = New System.Drawing.Point(694, 55)
        Me.BtnSubmit.Name = "BtnSubmit"
        Me.BtnSubmit.Size = New System.Drawing.Size(75, 23)
        Me.BtnSubmit.TabIndex = 100
        Me.BtnSubmit.Text = "&Submit"
        Me.BtnSubmit.UseVisualStyleBackColor = True
        '
        'Button3
        '
        Me.Button3.Location = New System.Drawing.Point(658, 57)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(29, 21)
        Me.Button3.TabIndex = 101
        Me.Button3.Text = "<"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'lblCartons
        '
        Me.lblCartons.AutoSize = True
        Me.lblCartons.Location = New System.Drawing.Point(1116, 211)
        Me.lblCartons.Name = "lblCartons"
        Me.lblCartons.Size = New System.Drawing.Size(43, 13)
        Me.lblCartons.TabIndex = 102
        Me.lblCartons.Text = "Cartons"
        Me.lblCartons.Visible = False
        '
        'ColumnHeader8
        '
        Me.ColumnHeader8.Text = "Packets/Carton"
        Me.ColumnHeader8.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.ColumnHeader8.Width = 100
        '
        'frmSales
        '
        Me.AcceptButton = Me.Save
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.ClientSize = New System.Drawing.Size(1269, 669)
        Me.Controls.Add(Me.lblCartons)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.BtnSubmit)
        Me.Controls.Add(Me.txtOrderNo)
        Me.Controls.Add(Me.Label17)
        Me.Controls.Add(Me.btnRemove)
        Me.Controls.Add(Me.ListView1)
        Me.Controls.Add(Me.btnPrint)
        Me.Controls.Add(Me.txtCustomerName)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.Label13)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.Button4)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.txtInvoiceNo)
        Me.Controls.Add(Me.dtpTransactionDate)
        Me.Controls.Add(Me.txtCustomerNo)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Name = "frmSales"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Sales"
        Me.Panel1.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents NewRecord As System.Windows.Forms.Button
    Friend WithEvents Delete As System.Windows.Forms.Button
    Friend WithEvents Save As System.Windows.Forms.Button
    Friend WithEvents Button4 As System.Windows.Forms.Button
    Friend WithEvents txtCartons As System.Windows.Forms.TextBox
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents txtWeight As System.Windows.Forms.TextBox
    Friend WithEvents txtAvailableCartons As System.Windows.Forms.TextBox
    Friend WithEvents txtPrice As System.Windows.Forms.TextBox
    Friend WithEvents txtCustomerNo As System.Windows.Forms.TextBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents txtProductName As System.Windows.Forms.TextBox
    Friend WithEvents txtProductCode As System.Windows.Forms.TextBox
    Friend WithEvents dtpTransactionDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents txtInvoiceNo As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents Button7 As System.Windows.Forms.Button
    Friend WithEvents btnPrint As System.Windows.Forms.Button
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents txtTotal As System.Windows.Forms.TextBox
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents txtTaxAmt As System.Windows.Forms.TextBox
    Friend WithEvents Label24 As System.Windows.Forms.Label
    Friend WithEvents txtTaxPer As System.Windows.Forms.TextBox
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents txtSubTotal As System.Windows.Forms.TextBox
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents txtCustomerName As System.Windows.Forms.TextBox
    Friend WithEvents ListView1 As System.Windows.Forms.ListView
    Friend WithEvents btnRemove As System.Windows.Forms.Button
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents ColumnHeader1 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ProductCode As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader2 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader3 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader4 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader5 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader6 As System.Windows.Forms.ColumnHeader
    Friend WithEvents Label18 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents txtTotalAmount As System.Windows.Forms.TextBox
    Friend WithEvents txtPackets As System.Windows.Forms.TextBox
    Friend WithEvents txtAvailablePackets As System.Windows.Forms.TextBox
    Friend WithEvents txtPaymentDue As System.Windows.Forms.TextBox
    Friend WithEvents txtTotalPayment As System.Windows.Forms.TextBox
    Friend WithEvents Label19 As System.Windows.Forms.Label
    Friend WithEvents payment As System.Windows.Forms.Label
    Friend WithEvents ColumnHeader7 As System.Windows.Forms.ColumnHeader
    Friend WithEvents btnUpdate As System.Windows.Forms.Button
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    Friend WithEvents txtOrderNo As System.Windows.Forms.TextBox
    Friend WithEvents BtnSubmit As System.Windows.Forms.Button
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents lblCartons As System.Windows.Forms.Label
    Friend WithEvents ColumnHeader8 As System.Windows.Forms.ColumnHeader
End Class
