Imports System.Data.OleDB
Imports System.Security.Cryptography
Imports System.Text
Public Class frmStock
    Dim rdr As OleDbDataReader = Nothing
    Dim dtable As DataTable
    Dim con As OleDbConnection = Nothing
    Dim adp As OleDbDataAdapter
    Dim ds As DataSet
    Dim cmd As OleDbCommand = Nothing
    Dim dt As New DataTable
    Dim cs As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\SI_DB.accdb;Persist Security Info=False;"

    Sub clear()
        txtStockID.Text = ""
        txtCartons.Text = ""
        txtCategory.Text = ""
        txtPackets.Text = ""
        txtWeight.Text = ""
        txtProductCode.Text = ""
        txtProductName.Text = ""
        txtTotalPackets.Text = ""
        dtpStockDate.Text = Today
        Button2.Focus()
    End Sub
    Private Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\SI_DB.accdb;Persist Security Info=False;"
    Private ReadOnly Property Connection() As OleDbConnection
        Get
            Dim ConnectionToFetch As New OleDbConnection(ConnectionString)
            ConnectionToFetch.Open()
            Return ConnectionToFetch
        End Get
    End Property
    Public Function GetData() As DataView
        Dim SelectQry = "SELECT (ProductCode) as [Product Code],(ProductName) as [Product Name],(Weight) as [Weight],sum(Cartons) as [Cartons],Packets,Sum(TotalPackets) as [Total Packets] FROM stock where Cartons > 0 and TotalPackets > 0   group by ProductCode,ProductName,Weight,Packets order by ProductName "
        Dim SampleSource As New DataSet
        Dim TableView As DataView
        Try
            Dim SampleCommand As New OleDbCommand()
            Dim SampleDataAdapter = New OleDbDataAdapter()
            SampleCommand.CommandText = SelectQry
            SampleCommand.Connection = Connection
            SampleDataAdapter.SelectCommand = SampleCommand
            SampleDataAdapter.Fill(SampleSource)
            TableView = SampleSource.Tables(0).DefaultView
        Catch ex As Exception
            Throw ex
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        Return TableView
    End Function

    Private Sub txtPackets_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtPackets.KeyPress
        If (e.KeyChar < Chr(48) Or e.KeyChar > Chr(57)) And e.KeyChar <> Chr(8) Then

            e.Handled = True

        End If
    End Sub
    Private Sub txtPackets_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtPackets.TextChanged
        txtTotalPackets.Text = CInt(Val(txtCartons.Text) * Val(txtPackets.Text))
    End Sub

    Private Sub txtCartons_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtCartons.KeyPress
        If (e.KeyChar < Chr(48) Or e.KeyChar > Chr(57)) And e.KeyChar <> Chr(8) Then

            e.Handled = True

        End If
    End Sub

    Private Sub txtCartons_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCartons.TextChanged
        txtTotalPackets.Text = CInt(Val(txtCartons.Text) * Val(txtPackets.Text))
    End Sub


    Private Sub NewRecord_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NewRecord.Click
        clear()
        Save.Enabled = True
        Update_Record.Enabled = False
        Delete.Enabled = False
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
    Sub auto()
        txtStockID.Text = "ST-" & GetUniqueKey(6)
    End Sub
    Private Sub Save_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Save.Click
        If Len(Trim(txtProductCode.Text)) = 0 Then
            MessageBox.Show("Please select product code", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtProductCode.Focus()
            Exit Sub
        End If
        If Len(Trim(txtProductName.Text)) = 0 Then
            MessageBox.Show("Please select product name", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtProductName.Focus()
            Exit Sub
        End If
        If Len(Trim(txtCategory.Text)) = 0 Then
            MessageBox.Show("Please select category", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtCategory.Focus()
            Exit Sub
        End If
        If Len(Trim(txtWeight.Text)) = 0 Then
            MessageBox.Show("Please select weight", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtWeight.Focus()
            Exit Sub
        End If
        If Len(Trim(txtCartons.Text)) = 0 Then
            MessageBox.Show("Please enter cartons", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtCartons.Focus()
            Exit Sub
        End If
        If Len(Trim(txtPackets.Text)) = 0 Then
            MessageBox.Show("Please enter packets", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtPackets.Focus()
            Exit Sub
        End If
        Try
            con = New OleDbConnection(cs)
            con.Open()
            Dim ct As String = "select ProductCode and Packets  from stock where ProductCode=@find and packets=@find1"

            cmd = New OleDbCommand(ct)
            cmd.Connection = con
            cmd.Parameters.Add(New OleDbParameter("@find", System.Data.OleDB.OleDBType.VarChar, 20, "ProductCode"))
            cmd.Parameters("@find").Value = txtProductCode.Text
            cmd.Parameters.Add(New OleDbParameter("@find1", System.Data.OleDb.OleDbType.Integer, 10, "Packets"))
            cmd.Parameters("@find1").Value = CInt(txtPackets.Text)
            rdr = cmd.ExecuteReader()

            If rdr.Read Then
                MessageBox.Show("Record already exists" & vbCrLf & "please update the stock of product", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)

                If Not rdr Is Nothing Then
                    rdr.Close()
                End If
                Exit Sub
            End If


            auto()
            con = New OleDbConnection(cs)
            con.Open()
            Dim ct1 As String = "select stockid from stock where stockid=@find"

            cmd = New OleDbCommand(ct1)
            cmd.Connection = con
            cmd.Parameters.Add(New OleDbParameter("@find", System.Data.OleDB.OleDBType.VarChar, 20, "stockid"))
            cmd.Parameters("@find").Value = txtStockID.Text
            rdr = cmd.ExecuteReader()

            If rdr.Read Then
                MessageBox.Show("Stock ID Already Exists", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)

                If Not rdr Is Nothing Then
                    rdr.Close()
                End If

            Else



                con = New OleDbConnection(cs)
                con.Open()

                Dim cb As String = "insert into stock(StockID,productcode,productname,category,weight,stockdate,Cartons,Packets,TotalPackets) VALUES ('" & txtStockID.Text & "','" & txtProductCode.Text & "','" & txtProductName.Text & "','" & txtCategory.Text & "','" & txtWeight.Text & "','" & dtpStockDate.Text & "','" & CInt(txtCartons.Text) & "','" & CInt(txtPackets.Text) & "','" & CInt(txtTotalPackets.Text) & "')"

                cmd = New OleDbCommand(cb)

                cmd.Connection = con
              

                cmd.ExecuteReader()
                MessageBox.Show("Successfully saved", "Stock Details", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Save.Enabled = False
                DataGridView1.DataSource = GetData()
                If con.State = ConnectionState.Open Then
                    con.Close()
                End If

                con.Close()
            End If



        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub frmStock_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Me.Hide()
        FrmMain.Show()
    End Sub

    Private Sub frmStock_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        DataGridView1.DataSource = GetData()
    End Sub

    Private Sub Update_Record_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Update_Record.Click
        Try
            con = New OleDbConnection(cs)
            con.Open()

            Dim cb As String = "update stock set productcode = '" & txtProductCode.Text & "',productname='" & txtProductName.Text & "',category='" & txtCategory.Text & "',weight='" & txtWeight.Text & "',stockdate='" & dtpStockDate.Text & "',Cartons='" & CInt(txtCartons.Text) & "',Packets='" & CInt(txtPackets.Text) & "',TotalPackets='" & CInt(txtTotalPackets.Text) & "' where stockid='" & txtStockID.Text & "'"

            cmd = New OleDbCommand(cb)

            cmd.Connection = con
          

            cmd.ExecuteReader()
            MessageBox.Show("Successfully updated", "Stock Details", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Update_Record.Enabled = False
            DataGridView1.DataSource = GetData()
            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            con.Close()



        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Delete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Delete.Click
        Try



            If MessageBox.Show("Do you really want to delete the record?", "Stock Record", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) = Windows.Forms.DialogResult.Yes Then
                delete_records()



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


            Dim cq As String = "delete from stock where stockid=@DELETE1;"


            cmd = New OleDbCommand(cq)

            cmd.Connection = con

            cmd.Parameters.Add(New OleDbParameter("@DELETE1", System.Data.OleDB.OleDBType.VarChar, 20, "stockid"))


            cmd.Parameters("@DELETE1").Value = Trim(txtStockID.Text)
            RowsAffected = cmd.ExecuteNonQuery()
            If RowsAffected > 0 Then

                MessageBox.Show("Successfully deleted", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information)
                DataGridView1.DataSource = GetData()
                clear()

                Update_Record.Enabled = False
                Delete.Enabled = False

            Else
                MessageBox.Show("No record found", "Sorry", MessageBoxButtons.OK, MessageBoxIcon.Information)
                DataGridView1.DataSource = GetData()
                clear()
                Update_Record.Enabled = False
                Delete.Enabled = False

                If con.State = ConnectionState.Open Then

                    con.Close()
                End If

                con.Close()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.clear()
        frmProductsRecord.DataGridView4.DataSource = Nothing
        frmProductsRecord.cmbCategory.Text = ""
        frmProductsRecord.cmbWeight.Text = ""
        frmProductsRecord.DataGridView3.DataSource = Nothing
        frmProductsRecord.cmbProductName.Text = ""
        frmProductsRecord.txtProduct.Text = ""
        frmProductsRecord.DataGridView2.DataSource = Nothing
        frmProductsRecord.DataGridView1.DataSource = Nothing
        frmProductsRecord.Show()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.clear()
        frmStockDetails1.fillCategory()
        frmStockDetails1.fillProduct()
        frmStockDetails1.fillWeight()
        frmStockDetails1.cmbProductName.Text = ""
        frmStockDetails1.txtProduct.Text = ""
        frmStockDetails1.DataGridView2.DataSource = Nothing
        frmStockDetails1.cmbCategory.Text = ""
        frmStockDetails1.DataGridView3.DataSource = Nothing
        frmStockDetails1.cmbWeight.Text = ""
        frmStockDetails1.DataGridView4.DataSource = Nothing
        frmStockDetails1.DataGridView1.DataSource = Nothing
        frmStockDetails1.Show()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        System.Diagnostics.Process.Start("Calc.exe")
    End Sub

   
End Class