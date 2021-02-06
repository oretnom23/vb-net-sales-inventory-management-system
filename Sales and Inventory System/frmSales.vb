Imports System.Data.OleDB
Imports System.Security.Cryptography
Imports System.Text

Public Class frmSales
    Dim rdr As OleDbDataReader = Nothing
    Dim dtable As DataTable
    Dim con As OleDbConnection = Nothing
    Dim adp As OleDbDataAdapter
    Dim ds As DataSet
    Dim cmd As OleDbCommand = Nothing
    Dim dt As New DataTable
    Dim sSql As String

    Dim cs As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\SI_DB.accdb;Persist Security Info=False;"
    Private Sub auto()
        txtInvoiceNo.Text = "INV-" & GetUniqueKey(8)

    End Sub
    Public Shared Function GetUniqueKey(ByVal maxSize As Integer) As String
        Dim chars As Char() = New Char(61) {}
        chars = "123456789".ToCharArray()
        Dim data As Byte() = New Byte(0) {}
        Dim crypto As New RNGCryptoServiceProvider()
        crypto.GetNonZeroBytes(data)
        data = New Byte(maxSize - 1) {}
        crypto.GetNonZeroBytes(data)
        Dim result As New StringBuilder(maxSize)
        For Each b As Byte In data
            result.Append(chars(b Mod (chars.Length)))
        Next
        Return result.ToString()
    End Function
    Sub clear()
        txtInvoiceNo.Text = ""
        txtCustomerNo.Text = ""
        txtCustomerName.Text = ""
        dtpTransactionDate.Text = Today
        txtProductCode.Text = ""
        txtProductName.Text = ""
        txtWeight.Text = ""
        txtAvailableCartons.Text = ""
        txtCartons.Text = ""
        txtPrice.Text = ""
        txtAvailablePackets.Text = ""
        txtPackets.Text = ""
        txtTotalAmount.Text = ""
        txtSubTotal.Text = ""
        txtTaxPer.Text = ""
        txtTaxAmt.Text = ""
        txtTotal.Text = ""
        txtTotalPayment.Text = ""
        txtPaymentDue.Text = ""
        txtOrderNo.Text = ""
    End Sub
    Private Sub NewRecord_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NewRecord.Click
        clear()
        btnPrint.Enabled = False
        Save.Enabled = True
        Delete.Enabled = False
        btnUpdate.Enabled = False
        ListView1.Items.Clear()
        btnRemove.Enabled = False
    End Sub

    Private Sub Save_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Save.Click

        If Len(Trim(txtCustomerNo.Text)) = 0 Then
            MessageBox.Show("Select Distributor id", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtCustomerNo.Focus()
            Exit Sub
        End If
        If Len(Trim(txtCustomerName.Text)) = 0 Then
            MessageBox.Show("Select Distributor name", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Button1.Focus()
            Exit Sub
        End If

        If Len(Trim(txtTaxPer.Text)) = 0 Then
            MessageBox.Show("Please enter tax percentage", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtTaxPer.Focus()
            Exit Sub
        End If
        If Len(Trim(txtTaxAmt.Text)) = 0 Then
            MessageBox.Show("Please enter tax amount", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtTaxAmt.Focus()
            Exit Sub
        End If
        If Len(Trim(txtTotalPayment.Text)) = 0 Then
            MessageBox.Show("Please enter total payment", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtTotalPayment.Focus()
            Exit Sub
        End If
        If ListView1.Items.Count = 0 Then
            MessageBox.Show("sorry no product added", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If
        Try
            For j = 0 To ListView1.Items.Count - 1
                Dim con As New OleDbConnection(cs)
                con.Open()
                Dim cmd As New OleDbCommand("SELECT cartons  from stock where ProductCode= '" & ListView1.Items(j).SubItems(1).Text & "' and Packets =" & ListView1.Items(j).SubItems(6).Text & "", con)
                Dim da As New OleDbDataAdapter(cmd)
                Dim ds As DataSet = New DataSet()
                da.Fill(ds)
                If ds.Tables(0).Rows.Count > 0 Then
                    lblCartons.Text = ds.Tables(0).Rows(0)("Cartons")
                    If Val(ListView1.Items(j).SubItems(5).Text) > Val(lblCartons.Text) Then
                        MessageBox.Show("added cartons to cart are more than" & vbCrLf & "available cartons of product code '" & ListView1.Items(j).SubItems(1).Text & "' and Packets = " & ListView1.Items(j).SubItems(6).Text & "", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        txtTaxAmt.Text = ""
                        txtTaxPer.Text = ""
                        txtTotal.Text = ""
                        txtTotalPayment.Text = ""
                        txtPaymentDue.Text = ""
                        Exit Sub
                    End If
                End If
                con.Close()
            Next
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        auto()
        con = New OleDbConnection(cs)
        con.Open()
        Dim ct As String = "select invoiceno from billinfo where invoiceno=@find"

        cmd = New OleDbCommand(ct)
        cmd.Connection = con
        cmd.Parameters.Add(New OleDbParameter("@find", System.Data.OleDb.OleDbType.VarChar, 20, "invoiceno"))
        cmd.Parameters("@find").Value = txtInvoiceNo.Text
        rdr = cmd.ExecuteReader()

        If rdr.Read Then
            MessageBox.Show("Invoice No. Already Exists", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)

            If Not rdr Is Nothing Then
                rdr.Close()
            End If

        Else



            con = New OleDbConnection(cs)
            con.Open()

            Dim cb As String = "insert Into billinfo(InvoiceNo,BillingDate,CustomerNo,CustomerName,SubTotal,TaxPercentage,TaxAmount,GrandTotal,TotalPayment,PaymentDue) VALUES ('" & txtInvoiceNo.Text & "','" & dtpTransactionDate.Text & "','" & txtCustomerNo.Text & "','" & txtCustomerName.Text & "','" & CInt(txtSubTotal.Text) & "','" & CDbl(txtTaxPer.Text) & "','" & CInt(txtTaxAmt.Text) & "','" & CInt(txtTotal.Text) & "','" & CInt(txtTotalPayment.Text) & "','" & CInt(txtPaymentDue.Text) & "')"

            cmd = New OleDbCommand(cb)

            cmd.Connection = con


            cmd.ExecuteReader()

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            con.Close()

            For i = 0 To ListView1.Items.Count - 1

                con = New OleDbConnection(cs)

                Dim cd As String = "insert Into ProductSold(InvoiceNo,ProductCode,ProductName,Weight,Price,Cartons,Packets,TotalPackets,TotalAmount) VALUES (@InvoiceNo,@ProductCode,@ProductName,@Weight,@Price,@Cartons,@Packets,@TotalPackets,@Totalamount)"

                cmd = New OleDbCommand(cd)

                cmd.Connection = con



                cmd.Parameters.AddWithValue("InvoiceNo", txtInvoiceNo.Text)
                cmd.Parameters.AddWithValue("ProductCode", ListView1.Items(i).SubItems(1).Text)
                cmd.Parameters.AddWithValue("ProductName", ListView1.Items(i).SubItems(2).Text)
                cmd.Parameters.AddWithValue("Weight", ListView1.Items(i).SubItems(3).Text)
                cmd.Parameters.AddWithValue("Price", ListView1.Items(i).SubItems(4).Text)
                cmd.Parameters.AddWithValue("Cartons", ListView1.Items(i).SubItems(5).Text)
                cmd.Parameters.AddWithValue("Packets", ListView1.Items(i).SubItems(6).Text)
                cmd.Parameters.AddWithValue("TotalPackets", ListView1.Items(i).SubItems(7).Text)
                cmd.Parameters.AddWithValue("TotalAmount", ListView1.Items(i).SubItems(8).Text)
                con.Open()
                cmd.ExecuteNonQuery()
                con.Close()
            Next

            For i = 0 To ListView1.Items.Count - 1

                con = New OleDbConnection(cs)
                con.Open()

                Dim cb1 As String = "update stock set Cartons = Cartons - '" & ListView1.Items(i).SubItems(5).Text & "' where productcode= '" & ListView1.Items(i).SubItems(1).Text & "' and packets = " & ListView1.Items(i).SubItems(6).Text & ""

                cmd = New OleDbCommand(cb1)

                cmd.Connection = con


                cmd.ExecuteNonQuery()
                con.Close()
            Next

            For i = 0 To ListView1.Items.Count - 1

                con = New OleDbConnection(cs)
                con.Open()

                Dim cb2 As String = "update stock set TotalPackets = TotalPackets - '" & ListView1.Items(i).SubItems(7).Text & "' where productcode= '" & ListView1.Items(i).SubItems(1).Text & "' and packets = " & ListView1.Items(i).SubItems(6).Text & ""

                cmd = New OleDbCommand(cb2)

                cmd.Connection = con


                cmd.ExecuteReader()
                con.Close()
            Next
            con = New OleDbConnection(cs)
            con.Open()

            Dim cb3 As String = "update orderinfo set orderstatus = 'Completed' where orderNo ='" & txtOrderNo.Text & "'"

            cmd = New OleDbCommand(cb3)

            cmd.Connection = con


            cmd.ExecuteReader()
            con.Close()
            Save.Enabled = False
            btnPrint.Enabled = True
            MessageBox.Show("Successfully saved", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information)

        End If




    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        frmCustomersRecord.Show()
        txtCustomerName.Text = ""
        txtCustomerNo.Text = ""

    End Sub




    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        Try
            If Len(Trim(txtProductCode.Text)) = 0 Then
                MessageBox.Show("Please select product code", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Button1.Focus()
                Exit Sub
            End If
            If Len(Trim(txtCartons.Text)) = 0 Then
                MessageBox.Show("Please enter no. of cartons", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                txtCartons.Focus()
                Exit Sub
            End If
            If Val(txtCartons.Text) = 0 Then
                MessageBox.Show("no. of cartons can not be zero", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                txtCartons.Focus()
                Exit Sub
            End If
            Dim temp As Integer
            temp = ListView1.Items.Count()
            If temp = 0 Then
                Dim i As Integer
                Dim lst As New ListViewItem(i)
                lst.SubItems.Add(txtProductCode.Text)
                lst.SubItems.Add(txtProductName.Text)
                lst.SubItems.Add(txtWeight.Text)
                lst.SubItems.Add(txtPrice.Text)
                lst.SubItems.Add(txtCartons.Text)
                lst.SubItems.Add(CInt(Val(txtAvailablePackets.Text) / Val(txtAvailableCartons.Text)))
                lst.SubItems.Add(txtPackets.Text)
                lst.SubItems.Add(txtTotalAmount.Text)
                ListView1.Items.Add(lst)
                i = i + 1
                txtSubTotal.Text = subtot()
                txtProductCode.Text = ""
                txtProductName.Text = ""
                txtCartons.Text = ""
                txtWeight.Text = ""
                txtPrice.Text = ""
                txtAvailableCartons.Text = ""
                txtAvailablePackets.Text = ""
                txtPackets.Text = ""
                txtTotalAmount.Text = ""
                Exit Sub
            End If

            For j = 0 To temp - 1
                If (ListView1.Items(j).SubItems(1).Text = txtProductCode.Text And ListView1.Items(j).SubItems(6).Text = Val(txtAvailablePackets.Text) / Val(txtAvailableCartons.Text)) Then
                    ListView1.Items(j).SubItems(1).Text = txtProductCode.Text
                    ListView1.Items(j).SubItems(2).Text = txtProductName.Text
                    ListView1.Items(j).SubItems(3).Text = txtWeight.Text
                    ListView1.Items(j).SubItems(4).Text = txtPrice.Text
                    ListView1.Items(j).SubItems(5).Text = Val(ListView1.Items(j).SubItems(5).Text) + Val(txtCartons.Text)
                    ListView1.Items(j).SubItems(6).Text = Val(txtAvailablePackets.Text) / Val(txtAvailableCartons.Text)
                    ListView1.Items(j).SubItems(7).Text = Val(ListView1.Items(j).SubItems(7).Text) + Val(txtPackets.Text)
                    ListView1.Items(j).SubItems(8).Text = Val(ListView1.Items(j).SubItems(8).Text) + Val(txtTotalAmount.Text)
                    txtSubTotal.Text = subtot()
                    txtProductCode.Text = ""
                    txtProductName.Text = ""
                    txtCartons.Text = ""
                    txtWeight.Text = ""
                    txtPrice.Text = ""
                    txtAvailableCartons.Text = ""
                    txtAvailablePackets.Text = ""
                    txtPackets.Text = ""
                    txtTotalAmount.Text = ""
                    Exit Sub

                End If
            Next j
            Dim k As Integer
            Dim lst1 As New ListViewItem(k)

            lst1.SubItems.Add(txtProductCode.Text)
            lst1.SubItems.Add(txtProductName.Text)
            lst1.SubItems.Add(txtWeight.Text)
            lst1.SubItems.Add(txtPrice.Text)
            lst1.SubItems.Add(txtCartons.Text)
            lst1.SubItems.Add(Val(txtAvailablePackets.Text) / Val(txtAvailableCartons.Text))
            lst1.SubItems.Add(txtPackets.Text)
            lst1.SubItems.Add(txtTotalAmount.Text)
            ListView1.Items.Add(lst1)
            k = k + 1
            txtSubTotal.Text = subtot()
            txtProductCode.Text = ""
            txtProductName.Text = ""
            txtCartons.Text = ""
            txtWeight.Text = ""
            txtPrice.Text = ""
            txtAvailableCartons.Text = ""
            txtAvailablePackets.Text = ""
            txtPackets.Text = ""
            txtTotalAmount.Text = ""
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try


    End Sub
    Public Function subtot() As Double

        Dim i, j, k As Integer
        i = 0
        j = 0
        k = 0

        Try
            j = ListView1.Items.Count
            For i = 0 To j - 1
                k = k + CInt(ListView1.Items(i).SubItems(8).Text)
            Next
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        Return k

    End Function

    Private Sub frmSales_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Me.Hide()
        FrmMain.Show()
    End Sub

    Private Sub frmSales_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRemove.Click
        Try

            If ListView1.Items.Count = 0 Then
                MsgBox("No items to remove", MsgBoxStyle.Critical, "Error")
            Else
                Dim itmCnt, i, t As Integer

                ListView1.FocusedItem.Remove()
                itmCnt = ListView1.Items.Count
                t = 1
                For i = 1 To itmCnt + 1

                    'Dim lst1 As New ListViewItem(i)
                    'ListView1.Items(i).SubItems(0).Text = t
                    t = t + 1

                Next
                txtSubTotal.Text = subtot()
            End If

            btnRemove.Enabled = False
            If ListView1.Items.Count = 0 Then
                txtSubTotal.Text = ""
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub txtCartons_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtCartons.KeyPress
        If (e.KeyChar < Chr(48) Or e.KeyChar > Chr(57)) And e.KeyChar <> Chr(8) Then

            e.Handled = True

        End If
    End Sub

    Private Sub txtQty_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCartons.TextChanged

        Try
            If Val(txtAvailableCartons.Text) = 0 Then
                txtPackets.Text = 0
                txtTotalAmount.Text = 0
                Exit Sub
            End If
            txtPackets.Text = CInt(Val(txtCartons.Text) * (Val(txtAvailablePackets.Text) / Val(txtAvailableCartons.Text)))
            txtTotalAmount.Text = CInt(Val(txtPackets.Text) * Val(txtPrice.Text))
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub txtTaxPer_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtTaxPer.KeyPress
        Dim keyChar = e.KeyChar

        If Char.IsControl(keyChar) Then
            'Allow all control characters.
        ElseIf Char.IsDigit(keyChar) OrElse keyChar = "."c Then
            Dim text = Me.txtTaxPer.Text
            Dim selectionStart = Me.txtTaxPer.SelectionStart
            Dim selectionLength = Me.txtTaxPer.SelectionLength

            text = text.Substring(0, selectionStart) & keyChar & text.Substring(selectionStart + selectionLength)

            If Integer.TryParse(text, New Integer) AndAlso text.Length > 16 Then
                'Reject an Integereger that is longer than 16 digits.
                e.Handled = True
            ElseIf Double.TryParse(text, New Double) AndAlso text.IndexOf("."c) < text.Length - 3 Then
                'Reject a real number with two many decimal places.
                e.Handled = False
            End If
        Else
            'Reject all other characters.
            e.Handled = True
        End If
    End Sub

    Private Sub txtTaxPer_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtTaxPer.TextChanged
        Try
            If txtTaxPer.Text = "" Then
                txtTaxAmt.Text = ""
                txtTotal.Text = ""
                Exit Sub
            End If
            txtTaxAmt.Text = CInt((Val(txtSubTotal.Text) * Val(txtTaxPer.Text)) / 100)
            txtTotal.Text = Val(txtSubTotal.Text) + Val(txtTaxAmt.Text)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        Cursor = Cursors.WaitCursor
        Timer1.Enabled = True
        frmBillingReport.Show()
    End Sub



    Private Sub txtQty_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtCartons.Validating
        Try
            If Val(txtCartons.Text) > Val(txtAvailableCartons.Text) Then
                MessageBox.Show("selling quanties of cartons are more than available", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                txtCartons.Text = ""
                txtCartons.Focus()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub delete_records()
        Try



            Dim RowsAffected As Integer = 0




            con = New OleDbConnection(cs)

            con.Open()


            Dim cq1 As String = "delete from productsold where invoiceno=@DELETE1;"


            cmd = New OleDbCommand(cq1)

            cmd.Connection = con

            cmd.Parameters.Add(New OleDbParameter("@DELETE1", System.Data.OleDB.OleDBType.VarChar, 20, "InvoiceNo"))


            cmd.Parameters("@DELETE1").Value = Trim(txtInvoiceNo.Text)
            cmd.ExecuteNonQuery()
            con.Close()
            con = New OleDbConnection(cs)

            con.Open()


            Dim cq As String = "delete from billinfo where invoiceno=@DELETE1;"


            cmd = New OleDbCommand(cq)

            cmd.Connection = con

            cmd.Parameters.Add(New OleDbParameter("@DELETE1", System.Data.OleDB.OleDBType.VarChar, 20, "InvoiceNo"))


            cmd.Parameters("@DELETE1").Value = Trim(txtInvoiceNo.Text)
            RowsAffected = cmd.ExecuteNonQuery()
            If RowsAffected > 0 Then

                MessageBox.Show("Successfully deleted", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information)

                clear()
                Delete.Enabled = False
                btnUpdate.Enabled = False
                btnPrint.Enabled = False
            Else
                MessageBox.Show("No record found", "Sorry", MessageBoxButtons.OK, MessageBoxIcon.Information)

                clear()
                Delete.Enabled = False
                btnUpdate.Enabled = False
                btnPrint.Enabled = False


                If con.State = ConnectionState.Open Then

                    con.Close()
                End If

                con.Close()
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub Delete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Delete.Click
        Try



            If MessageBox.Show("Do you really want to delete the record?", "Sales Record", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) = Windows.Forms.DialogResult.Yes Then
                delete_records()



            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Me.clear()
        frmSalesRecord1.fillCustomerName()
        frmSalesRecord1.fillInvoiceNo()
        frmSalesRecord1.DataGridView4.DataSource = Nothing
        frmSalesRecord1.cmbInvoiceNo.Text = ""
        frmSalesRecord1.GroupBox5.Visible = False
        frmSalesRecord1.DataGridView3.DataSource = Nothing
        frmSalesRecord1.cmbCustomerName.Text = ""
        frmSalesRecord1.GroupBox4.Visible = False
        frmSalesRecord1.DateTimePicker1.Text = Today
        frmSalesRecord1.DateTimePicker2.Text = Today
        frmSalesRecord1.DataGridView2.DataSource = Nothing
        frmSalesRecord1.GroupBox10.Visible = False
        frmSalesRecord1.DataGridView1.DataSource = Nothing
        frmSalesRecord1.dtpInvoiceDateFrom.Text = Today
        frmSalesRecord1.dtpInvoiceDateTo.Text = Today
        frmSalesRecord1.GroupBox3.Visible = False
        frmSalesRecord1.Show()
    End Sub




    Public Sub FillList()
        With ListView1
            .Clear()
            .Columns.Add("Column11", 0)
            .Columns.Add("Product Code", 90)
            .Columns.Add("Product Name", 250, HorizontalAlignment.Center)
            .Columns.Add("Weight/Qty", 90, HorizontalAlignment.Center)
            .Columns.Add("Unit Price", 80, HorizontalAlignment.Center)
            .Columns.Add("Cartons", 85, HorizontalAlignment.Center)
            .Columns.Add("Packets/Carton", 100, HorizontalAlignment.Center)
            .Columns.Add("Total packets", 90, HorizontalAlignment.Center)
            .Columns.Add("Total Amount", 92, HorizontalAlignment.Center)
            FillListView(ListView1, GetData(sSql))

        End With
    End Sub
    Public Function GetData(ByVal sSQL As String)

        Dim sqlCmd As OleDbCommand = New OleDbCommand(sSQL)
        Dim myData As OleDbDataReader

        con = New OleDbConnection(cs)

        Try
            con.Open()

            sqlCmd.Connection = con

            myData = sqlCmd.ExecuteReader

            Return myData
        Catch ex As Exception
            Return ex
        End Try
    End Function
    Sub tot()
        Dim dblTotal As Double = 0
        Dim dblTemp As Double

        For Each lvItem As ListViewItem In ListView1.Items
            If Double.TryParse(lvItem.SubItems(5).Text, dblTemp) Then
                dblTotal += dblTemp
            End If
        Next

        txtSubTotal.Text = dblTotal
    End Sub




    Private Sub ListView1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ListView1.SelectedIndexChanged
        btnRemove.Enabled = True

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        frmSearchProduct1.fillCategory()
        frmSearchProduct1.fillProductName()
        frmSearchProduct1.fillWeight()
        frmSearchProduct1.txtProductName.Text = ""
        frmSearchProduct1.cmbProductName.Text = ""
        frmSearchProduct1.GroupBox2.Visible = False
        frmSearchProduct1.DataGridView1.DataSource = Nothing
        frmSearchProduct1.GroupBox5.Visible = False
        frmSearchProduct1.cmbWeight.Text = ""
        frmSearchProduct1.DataGridView2.DataSource = Nothing
        frmSearchProduct1.GroupBox7.Visible = False
        frmSearchProduct1.DataGridView3.DataSource = Nothing
        frmSearchProduct1.ComboBox1.Text = ""
        frmSearchProduct1.ComboBox2.Text = ""
        frmSearchProduct1.DataGridView4.DataSource = Nothing
        frmSearchProduct1.GroupBox13.Visible = False
        frmSearchProduct1.GroupBox16.Visible = False
        frmSearchProduct1.cmbCategory.Text = ""
        frmSearchProduct1.DataGridView5.DataSource = Nothing
        frmSearchProduct1.Show()
    End Sub

    Private Sub txtTotalPayment_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtTotalPayment.KeyPress
        If (e.KeyChar < Chr(48) Or e.KeyChar > Chr(57)) And e.KeyChar <> Chr(8) Then

            e.Handled = True

        End If
    End Sub

    Private Sub txtTotalPayment_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtTotalPayment.TextChanged
        Try
            txtPaymentDue.Text = Val(txtTotal.Text) - Val(txtTotalPayment.Text)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub txtTotalPayment_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtTotalPayment.Validating
        If Val(txtTotalPayment.Text) > Val(txtTotal.Text) Then
            MessageBox.Show("Total payment can not be more than grand total", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtTotalPayment.Text = ""
            txtTotalPayment.Focus()
        End If
    End Sub

    Private Sub btnUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdate.Click
        Try

            con = New OleDbConnection(cs)
            con.Open()

            Dim cb As String = "update billinfo set GrandTotal= '" & CInt(txtTotal.Text) & "',TotalPayment= '" & CInt(txtTotalPayment.Text) & "',PaymentDue= '" & CInt(txtPaymentDue.Text) & "' where Invoiceno= '" & txtInvoiceNo.Text & "'"

            cmd = New OleDbCommand(cb)

            cmd.Connection = con

            cmd.ExecuteReader()

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            con.Close()
            MessageBox.Show("Successfully updated", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information)

            btnUpdate.Enabled = False
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Cursor = Cursors.Default
        Timer1.Enabled = False
    End Sub

    Private Sub BtnSubmit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnSubmit.Click
        Try
            If (txtOrderNo.Text = "") Then
                MessageBox.Show("Retrieve order no.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Button3.Focus()
                Exit Sub
            End If
            sSql = "SELECT Cartons,ProductCode,ProductName,Weight,Price,Cartons,Packets,TotalPackets,TotalAmount from orderedProduct  where orderno = '" & txtOrderNo.Text & "'"
            Call FillList()

            con = New OleDbConnection(cs)

            con.Open()


            Dim ct As String = "select CustomerNo,CustomerName,subtotal,TaxPercentage,TaxAmount,TotalAmount from orderinfo where orderno=@find"


            cmd = New OleDbCommand(ct)
            cmd.Connection = con
            cmd.Parameters.Add(New OleDbParameter("@find", System.Data.OleDB.OleDBType.VarChar, 20, "orderno"))
            cmd.Parameters("@find").Value = Trim(txtOrderNo.Text)
            rdr = cmd.ExecuteReader()
            If rdr.Read Then


                txtCustomerNo.Text = Trim(rdr.GetString(0))
                txtCustomerName.Text = Trim(rdr.GetString(1))
                txtSubTotal.Text = Trim(rdr.GetInt32(2))
                txtTaxPer.Text = Trim(rdr.GetDouble(3))
                txtTaxAmt.Text = Trim(rdr.GetInt32(4))
                txtTotal.Text = Trim(rdr.GetInt32(5))
            End If


            If con.State = ConnectionState.Open Then
                con.Close()
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        frmOrders.DateTimePicker1.Text = Today
        frmOrders.DateTimePicker2.Text = Today
        frmOrders.DataGridView1.DataSource = Nothing
        frmOrders.Show()
    End Sub


End Class