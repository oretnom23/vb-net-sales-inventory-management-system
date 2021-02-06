Imports System.Data.OleDB
Public Class frmCustomersRecord

    Dim rdr As OleDbDataReader = Nothing
    Dim dtable As DataTable
    Dim con As OleDbConnection = Nothing
    Dim adp As OleDbDataAdapter
    Dim ds As DataSet
    Dim cmd As OleDbCommand = Nothing
    Dim dt As New DataTable
    Dim cs As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\SI_DB.accdb;Persist Security Info=False;"
    Private ReadOnly Property Connection() As OleDbConnection
        Get
            Dim ConnectionToFetch As New OleDbConnection(cs)
            ConnectionToFetch.Open()
            Return ConnectionToFetch
        End Get
    End Property
    Public Function GetData() As DataView
        Dim SelectQry = "SELECT (customerNo) as [Distributor ID],(B_name) as [B_Name],(b_address) as [B_Address],(b_landmark) as [B_LandMark],(b_city) as [B_City],(b_state) as [B_State],(b_zipcode) as [B_Zip/Post Code],(s_name) as [S_Name],(s_address) as [S_Address],(s_landmark) as [S_LandMark],(s_city) as [S_City],(s_state) as [S_State],(s_zipcode) as [S_Zip/Post Code],(Phone) as [Phone],(email)as [Email],(mobileno) as [Mobile No],(faxno) as [Fax No],(notes)as [Notes] from Customer order by customerno"

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









    Private Sub frmCustomersRecord_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        fillName()
        DataGridView1.DataSource = GetData()
    End Sub









    Private Sub TabControl1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabControl1.Click

        txtCustomer.Text = ""
        txtName.Text = ""
        DataGridView2.DataSource = Nothing
    End Sub


    Private Sub DataGridView2_RowHeaderMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView2.RowHeaderMouseClick
        Try
            Dim dr As DataGridViewRow = DataGridView2.SelectedRows(0)
            Me.Hide()
            frmSales.Show()
            ' or simply use column name instead of index
            'dr.Cells["id"].Value.ToString();
            frmSales.txtCustomerNo.Text = dr.Cells(0).Value.ToString()
            frmSales.txtCustomerName.Text = dr.Cells(1).Value.ToString()

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub DataGridView1_RowHeaderMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView1.RowHeaderMouseClick
        Try
            Dim dr As DataGridViewRow = DataGridView1.SelectedRows(0)
            Me.Hide()
            frmSales.Show()
            ' or simply use column name instead of index
            'dr.Cells["id"].Value.ToString();
            frmSales.txtCustomerNo.Text = dr.Cells(0).Value.ToString()
            frmSales.txtCustomerName.Text = dr.Cells(1).Value.ToString()

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub




    Sub fillName()

        Try

            Dim CN As New OleDbConnection(cs)

            CN.Open()
            adp = New OleDbDataAdapter()
            adp.SelectCommand = New OleDbCommand("SELECT distinct  (B_Name) FROM Customer", CN)
            ds = New DataSet("ds")

            adp.Fill(ds)
            dtable = ds.Tables(0)
            txtName.Items.Clear()

            For Each drow As DataRow In dtable.Rows
                txtName.Items.Add(drow(0).ToString())
                'DocName.SelectedIndex = -1
            Next
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCustomer.TextChanged
        Try
            con = New OleDbConnection(cs)
            con.Open()
            cmd = New OleDbCommand("SELECT (customerNo)[Distributor ID],(B_name)[B_Name],(b_address)[B_Address],(b_landmark)[B_LandMark],(b_city)[B_City],(b_state)[B_State],(b_zipcode)[B_Zip/Post Code],(s_name)[S_Name],(s_address)[S_Address],(s_landmark)[S_LandMark],(s_city)[S_City],(s_state)[S_State],(s_zipcode)[S_Zip/Post Code],(Phone)[Phone],(email)[Email],(mobileno)[Mobile No.],(faxno)[Fax No.],(notes)[Notes] from Customer where B_Name like '" & txtCustomer.Text & "%'  order by CustomerNo", con)



            Dim myDA As OleDbDataAdapter = New OleDbDataAdapter(cmd)

            Dim myDataSet As DataSet = New DataSet()

            myDA.Fill(myDataSet, "Customer")

            DataGridView2.DataSource = myDataSet.Tables("Customer").DefaultView



            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        txtCustomer.Text = ""
        txtName.Text = ""
        DataGridView2.DataSource = Nothing
    End Sub

    Private Sub txtName_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtName.SelectedIndexChanged
        Try
            con = New OleDbConnection(cs)
            con.Open()
            cmd = New OleDbCommand("SELECT (customerNo) as [Distributor ID],(B_name) as [B_Name],(b_address) as [B_Address],(b_landmark) as [B_LandMark],(b_city) as [B_City],(b_state) as [B_State],(b_zipcode) as [B_Zip/Post Code],(s_name) as [S_Name],(s_address) as [S_Address],(s_landmark) as [S_LandMark],(s_city) as [S_City],(s_state) as [S_State],(s_zipcode) as [S_Zip/Post Code],(Phone) as [Phone],(email)as [Email],(mobileno) as [Mobile No],(faxno) as [Fax No],(notes)as [Notes] from Customer where B_Name = '" & txtName.Text & "'  order by CustomerNo", con)



            Dim myDA As OleDbDataAdapter = New OleDbDataAdapter(cmd)

            Dim myDataSet As DataSet = New DataSet()

            myDA.Fill(myDataSet, "Customer")

            DataGridView2.DataSource = myDataSet.Tables("Customer").DefaultView



            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
End Class