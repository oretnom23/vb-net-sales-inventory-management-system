Imports System.Data.OleDB
Public Class frmOrders
    Dim rdr As OleDbDataReader = Nothing
    Dim dtable As DataTable
    Dim con As OleDbConnection = Nothing
    Dim adp As OleDbDataAdapter
    Dim ds As DataSet
    Dim cmd As OleDbCommand = Nothing
    Dim dt As New DataTable
    Dim cs As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\SI_DB.accdb;Persist Security Info=False;"

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        DateTimePicker1.Text = Today
        DateTimePicker2.Text = Today
        DataGridView1.DataSource = Nothing
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try


            con = New OleDbConnection(cs)

            con.Open()
            cmd = New OleDbCommand("SELECT (OrderNo) as [Order No],(OrderDate) as [Order Date],(OrderStatus) as [Order Status],(CustomerNo) as [Distributor ID],(CustomerName) as [Customer Name],(TotalAmount)as [Total Amount] from Orderinfo where OrderStatus='Uncompleted' and OrderDate between #" & DateTimePicker2.Value & "# And #" & DateTimePicker1.Value & "# order by orderinfo.OrderNo,OrderDate", con)
           
            Dim myDA As OleDbDataAdapter = New OleDbDataAdapter(cmd)

            Dim myDataSet As DataSet = New DataSet()

            myDA.Fill(myDataSet, "OrderInfo")

            DataGridView1.DataSource = myDataSet.Tables("Orderinfo").DefaultView

            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub frmOrders_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private Sub DataGridView1_RowHeaderMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView1.RowHeaderMouseClick
        Try
            Dim dr As DataGridViewRow = DataGridView1.SelectedRows(0)
            Me.Hide()
            frmSales.Show()
            ' or simply use column name instead of index
            'dr.Cells["id"].Value.ToString();
            frmSales.txtOrderNo.Text = dr.Cells(0).Value.ToString()


        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
End Class