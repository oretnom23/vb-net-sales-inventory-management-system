Imports System.Data.OleDB
Imports System.Security.Cryptography
Imports System.Text

Public Class frmOrder
    Dim rdr As OleDbDataReader = Nothing
    Dim dtable As DataTable
    Dim con As OleDbConnection = Nothing
    Dim adp As OleDbDataAdapter
    Dim ds As DataSet
    Dim cmd As OleDbCommand = Nothing
    Dim dt As New DataTable

    Dim cs As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\SI_DB.accdb;Persist Security Info=False;"
    Private Sub auto()
        txtOrderNo.Text = "OD-" & GetUniqueKey(8)

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
        txtOrderNo.Text = ""
        txtCustomerNo.Text = ""
        txtCustomerName.Text = ""
        dtpOrderDate.Text = Today
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
        cmbOrderStatus.Text = ""

    End Sub
    Private Sub NewRecord_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NewRecord.Click
        clear()
        ListView1.Items.Clear()
        Save.Enabled = True
        Delete.Enabled = False
        cmbOrderStatus.Enabled = False
        btnRemove.Enabled = False
    End Sub

    Private Sub frmOrder_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Me.Hide()
        FrmMain.Show()
    End Sub

    Private Sub frmOrder_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Save_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Save.Click

        If Len(Trim(txtCustomerNo.Text)) = 0 Then
            MessageBox.Show("Select Distributor id", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Button2.Focus()
            Exit Sub
        End If
        If Len(Trim(txtCustomerName.Text)) = 0 Then
            MessageBox.Show("Select Distributor name", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtCustomerName.Focus()
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

                        Exit Sub
                    End If
                End If
                con.Close()
            Next
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        Try
            auto()
            con = New OleDbConnection(cs)
            con.Open()
            Dim ct As String = "select OrderNo from Orderinfo where Orderno=@find"

            cmd = New OleDbCommand(ct)
            cmd.Connection = con
            cmd.Parameters.Add(New OleDbParameter("@find", System.Data.OleDB.OleDBType.VarChar, 20, "OrderNo"))
            cmd.Parameters("@find").Value = txtOrderNo.Text
            rdr = cmd.ExecuteReader()

            If rdr.Read Then
                MessageBox.Show("Order No. Already Exists", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)

                If Not rdr Is Nothing Then
                    rdr.Close()
                End If

            Else



                con = New OleDbConnection(cs)
                con.Open()

                Dim cb As String = "insert Into orderinfo(OrderNo,OrderDate,OrderStatus,CustomerNo,CustomerName,SubTotal,TaxPercentage,TaxAmount,TotalAmount) VALUES ('" & txtOrderNo.Text & "','" & dtpOrderDate.Text & "','Uncompleted','" & txtCustomerNo.Text & "','" & txtCustomerName.Text & "','" & CInt(txtSubTotal.Text) & "','" & CDbl(txtTaxPer.Text) & "','" & CInt(txtTaxAmt.Text) & "','" & CInt(txtTotal.Text) & "')"

                cmd = New OleDbCommand(cb)

                cmd.Connection = con

                cmd.ExecuteReader()

                If con.State = ConnectionState.Open Then
                    con.Close()
                End If

                con.Close()





                For i = 0 To ListView1.Items.Count - 1
                    con = New OleDbConnection(cs)

                    Dim cd As String = "insert Into OrderedProduct(OrderNo,ProductCode,ProductName,Weight,Price,Cartons,Packets,TotalPackets,TotalAmount) VALUES (@OrderNo,@ProductCode,@ProductName,@Weight,@Price,@Cartons,@Packets,@TotalPackets,@Totalamount)"

                    cmd = New OleDbCommand(cd)

                    cmd.Connection = con
                    cmd.Parameters.AddWithValue("OrderNo", txtOrderNo.Text)
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
                Save.Enabled = False
                MessageBox.Show("Successfully placed", "Order", MessageBoxButtons.OK, MessageBoxIcon.Information)
                cmbOrderStatus.Text = "Uncompleted"
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


            Dim cq1 As String = "delete from orderedproduct where orderno=@DELETE1;"


            cmd = New OleDbCommand(cq1)

            cmd.Connection = con

            cmd.Parameters.Add(New OleDbParameter("@DELETE1", System.Data.OleDB.OleDBType.VarChar, 20, "orderNo"))


            cmd.Parameters("@DELETE1").Value = Trim(txtOrderNo.Text)
            cmd.ExecuteNonQuery()
            con.Close()
            con = New OleDbConnection(cs)

            con.Open()


            Dim cq As String = "delete from orderinfo where orderno=@DELETE1;"


            cmd = New OleDbCommand(cq)

            cmd.Connection = con

            cmd.Parameters.Add(New OleDbParameter("@DELETE1", System.Data.OleDB.OleDBType.VarChar, 20, "orderNo"))


            cmd.Parameters("@DELETE1").Value = Trim(txtOrderNo.Text)
            RowsAffected = cmd.ExecuteNonQuery()
            If RowsAffected > 0 Then

                MessageBox.Show("Successfully deleted", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information)

                clear()
                cmbOrderStatus.Enabled = False
                Delete.Enabled = False
                btnUpdate.Enabled = False
            Else
                MessageBox.Show("No record found", "Sorry", MessageBoxButtons.OK, MessageBoxIcon.Information)


                clear()
                cmbOrderStatus.Enabled = False
                Delete.Enabled = False
                btnUpdate.Enabled = False


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



            If MessageBox.Show("Do you really want to delete the record?", "Order Record", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) = Windows.Forms.DialogResult.Yes Then
                delete_records()



            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click

        frmCustomersRecord2.Show()
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




    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click

        Me.clear()
        frmOrderRecord1.fillorderNo()
        frmOrderRecord1.fillProductName()
        frmOrderRecord1.fillCustomerName()
        frmOrderRecord1.DataGridView1.DataSource = Nothing
        frmOrderRecord1.dtpOrderDateFrom.Text = Today
        frmOrderRecord1.dtpOrderDateTo.Text = Today
        frmOrderRecord1.DataGridView6.DataSource = Nothing
        frmOrderRecord1.cmbCustomerName.Text = ""
        frmOrderRecord1.DataGridView3.DataSource = Nothing
        frmOrderRecord1.cmbStatus.Text = ""
        frmOrderRecord1.DataGridView2.DataSource = Nothing
        frmOrderRecord1.cmbOrderNo.Text = ""
        frmOrderRecord1.cmbProductName.Text = ""
        frmOrderRecord1.DataGridView5.DataSource = Nothing
        frmOrderRecord1.Show()
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

    Private Sub txtCartons_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtCartons.KeyPress
        If (e.KeyChar < Chr(48) Or e.KeyChar > Chr(57)) And e.KeyChar <> Chr(8) Then

            e.Handled = True

        End If
    End Sub



    Private Sub txtCartons_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCartons.TextChanged
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

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        frmSearchProduct.txtProductName.Text = ""
        frmSearchProduct.cmbProductName.Text = ""
        frmSearchProduct.GroupBox2.Visible = False
        frmSearchProduct.DataGridView1.DataSource = Nothing
        frmSearchProduct.GroupBox5.Visible = False
        frmSearchProduct.cmbWeight.Text = ""
        frmSearchProduct.DataGridView2.DataSource = Nothing
        frmSearchProduct.GroupBox7.Visible = False
        frmSearchProduct.DataGridView3.DataSource = Nothing
        frmSearchProduct.ComboBox1.Text = ""
        frmSearchProduct.ComboBox2.Text = ""
        frmSearchProduct.DataGridView4.DataSource = Nothing
        frmSearchProduct.GroupBox13.Visible = False
        frmSearchProduct.GroupBox16.Visible = False
        frmSearchProduct.cmbCategory.Text = ""
        frmSearchProduct.DataGridView5.DataSource = Nothing
        frmSearchProduct.Show()
    End Sub

    Private Sub btnUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdate.Click
        Try

            con = New OleDbConnection(cs)
            con.Open()

            Dim cb As String = "update orderinfo set orderstatus = '" & cmbOrderStatus.Text & "' where Orderno ='" & txtOrderNo.Text & "'"

            cmd = New OleDbCommand(cb)

            cmd.Connection = con


            cmd.ExecuteReader()

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            con.Close()
            MessageBox.Show("Successfully updated", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information)
            cmbOrderStatus.Enabled = False
            btnUpdate.Enabled = False
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub ListView1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListView1.SelectedIndexChanged
        btnRemove.Enabled = True
    End Sub

    Private Sub btnRemove_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRemove.Click
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
End Class