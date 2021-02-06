Imports System.Data.OleDB
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Public Class frmOrdersReport
    Dim dtable As DataTable
    Dim adp As OleDbDataAdapter
    Dim ds As DataSet
    Dim dt As New DataTable
    Dim cs As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\SI_DB.accdb;Persist Security Info=False;"


    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        dtpOrderDateFrom.Text = Today
        dtpOrderDateFrom.Text = Today
        CrystalReportViewer1.ReportSource = Nothing
    End Sub

    Private Sub Button12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button12.Click
        cmbStatus.Text = ""
        CrystalReportViewer2.ReportSource = Nothing
    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        cmbCustomerName.Text = ""
        CrystalReportViewer3.ReportSource = Nothing
    End Sub



    Private Sub TabControl1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabControl1.Click

        cmbCustomerName.Text = ""
        CrystalReportViewer3.ReportSource = Nothing
        dtpOrderDateFrom.Text = Today
        dtpOrderDateFrom.Text = Today
        CrystalReportViewer1.ReportSource = Nothing
        cmbStatus.Text = ""
        CrystalReportViewer2.ReportSource = Nothing
    End Sub

    Sub fillCustomerName()

        Try

            Dim CN As New OleDbConnection(cs)

            CN.Open()
            adp = New OleDbDataAdapter()
            adp.SelectCommand = New OleDbCommand("SELECT distinct  (CustomerName) FROM orderinfo", CN)
            ds = New DataSet("ds")

            adp.Fill(ds)
            dtable = ds.Tables(0)
            cmbCustomerName.Items.Clear()

            For Each drow As DataRow In dtable.Rows
                cmbCustomerName.Items.Add(drow(0).ToString())
                'DocName.SelectedIndex = -1
            Next
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub frmOrdersReport_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Me.Hide()
        FrmMain.Show()
    End Sub
    Private Sub frmOrdersReport_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        fillCustomerName()

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try
            Cursor = Cursors.WaitCursor
            Timer1.Enabled = True
            Dim rpt As New rptOrder() 'The report you created.
            Dim myConnection As OleDbConnection
            Dim MyCommand As New OleDbCommand()
            Dim myDA As New OleDbDataAdapter()
            Dim myDS As New SI_DBDataSet 'The DataSet you created.
            myConnection = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\SI_DB.accdb;Persist Security Info=False;")
            MyCommand.Connection = myConnection
            MyCommand.CommandText = "select * from Orderinfo where OrderDate between #" & dtpOrderDateFrom.Value & "# And #" & dtpOrderDateTo.Value & "# order by orderinfo.OrderNo,OrderDate"
           
            MyCommand.CommandType = CommandType.Text
            myDA.SelectCommand = MyCommand
            myDA.Fill(myDS, "Orderinfo")
            rpt.SetDataSource(myDS)
            CrystalReportViewer1.ReportSource = rpt
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click
        Try
            Cursor = Cursors.WaitCursor
            Timer1.Enabled = True
            Dim rpt As New rptOrder() 'The report you created.
            Dim myConnection As OleDbConnection
            Dim MyCommand As New OleDbCommand()
            Dim myDA As New OleDbDataAdapter()
            Dim myDS As New SI_DBDataSet 'The DataSet you created.


            myConnection = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\SI_DB.accdb;Persist Security Info=False;")
            MyCommand.Connection = myConnection
            MyCommand.CommandText = "select * from Orderinfo where OrderStatus = '" & cmbStatus.Text & "' order by OrderDate"
            MyCommand.CommandType = CommandType.Text
            myDA.SelectCommand = MyCommand
            myDA.Fill(myDS, "Orderinfo")
            rpt.SetDataSource(myDS)
            CrystalReportViewer2.ReportSource = rpt
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        Try
            Cursor = Cursors.WaitCursor
            Timer1.Enabled = True
            Dim rpt As New rptOrder() 'The report you created.
            Dim myConnection As OleDbConnection
            Dim MyCommand As New OleDbCommand()
            Dim myDA As New OleDbDataAdapter()
            Dim myDS As New SI_DBDataSet 'The DataSet you created.


            myConnection = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\SI_DB.accdb;Persist Security Info=False;")
            MyCommand.Connection = myConnection
            MyCommand.CommandText = "select * from Orderinfo where CustomerName = '" & cmbCustomerName.Text & "' order by OrderDate"
            MyCommand.CommandType = CommandType.Text
            myDA.SelectCommand = MyCommand
            myDA.Fill(myDS, "Orderinfo")
            rpt.SetDataSource(myDS)
            CrystalReportViewer3.ReportSource = rpt
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Cursor = Cursors.Default
        Timer1.Enabled = False
    End Sub
End Class