Imports System.Data.OleDb
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Public Class frmInVoicesReport
    Dim dtable As DataTable
    Dim adp As OleDbDataAdapter
    Dim ds As DataSet
    Dim dt As New DataTable
    Dim cs As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\SI_DB.accdb;Persist Security Info=False;"

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        fillCustomerName()
     
    End Sub

  


    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        DateTimePicker1.Text = Today
        DateTimePicker2.Text = Today
        CrystalReportViewer3.ReportSource = Nothing
    End Sub


    Private Sub TabControl1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TabControl1.Click
        CrystalReportViewer1.ReportSource = Nothing
        dtpInvoiceDateFrom.Text = Today
        dtpInvoiceDateTo.Text = Today
        CrystalReportViewer2.ReportSource = Nothing
        DateTimePicker1.Text = Today
        DateTimePicker2.Text = Today
        CrystalReportViewer3.ReportSource = Nothing
        cmbCustomerName.Text = ""
    End Sub

  

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Try
            Cursor = Cursors.WaitCursor
            Timer1.Enabled = True
            Dim rpt As New rptInvoices() 'The report you created.
            Dim myConnection As OleDbConnection
            Dim MyCommand As New OleDbCommand()
            Dim myDA As New OleDbDataAdapter()
            Dim myDS As New DataSet 'The DataSet you created.


            myConnection = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\SI_DB.accdb;Persist Security Info=False;")
            MyCommand.Connection = myConnection
            MyCommand.CommandText = "SELECT Billinfo.InvoiceNo,BillingDate,Customername,productCode,ProductName,Weight,Price,Cartons,TotalPackets,TotalAmount,TaxPercentage,TaxAmount,GrandTotal,TotalPayment,PaymentDue,Subtotal,count(CustomerNo),Packets from BillInfo,ProductSold where BillInfo.InvoiceNo = ProductSold.InvoiceNo and PaymentDue > 0 and BillingDate between #" & DateTimePicker2.Value & "# And #" & DateTimePicker1.Value & "#  group by Billinfo.InvoiceNo,Customername,BillingDate,productCode,ProductName,Weight,Price,Cartons,TotalPackets,TotalAmount,TotalPayment,GrandTotal,PaymentDue,TaxAmount,TaxPercentage,subtotal,Packets order by BillingDate"
            MyCommand.CommandType = CommandType.Text
            myDA.SelectCommand = MyCommand
            myDA.Fill(myDS, "BillInfo")
            myDA.Fill(myDS, "ProductSold")
            rpt.SetDataSource(myDS)

            CrystalReportViewer3.ReportSource = rpt
            myConnection.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Sub fillCustomerName()

        Try

            Dim CN As New OleDbConnection(cs)

            CN.Open()
            adp = New OleDbDataAdapter()
            adp.SelectCommand = New OleDbCommand("SELECT distinct  (CustomerName) FROM Billinfo", CN)
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

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Cursor = Cursors.Default
        Timer1.Enabled = False
    End Sub

    Private Sub frmInVoicesReport_FormClosing(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles MyBase.FormClosing
        Me.Hide()
        FrmMain.Show()
    End Sub


    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try
            Cursor = Cursors.WaitCursor
            Timer1.Enabled = True
            Dim rpt As New rptInvoices() 'The report you created.
            Dim myConnection As OleDbConnection
            Dim MyCommand As New OleDbCommand()
            Dim myDA As New OleDbDataAdapter()
            Dim myDS As New DataSet 'The DataSet you created.


            myConnection = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\SI_DB.accdb;Persist Security Info=False;")
            MyCommand.Connection = myConnection
            MyCommand.CommandText = "SELECT Billinfo.InvoiceNo,BillingDate,Customername,productCode,ProductName,subtotal,Weight,Price,Cartons,TotalPackets,TotalAmount,TaxPercentage,TaxAmount,GrandTotal,TotalPayment,PaymentDue,count(CustomerNo),Packets from BillInfo,ProductSold where BillInfo.InvoiceNo = ProductSold.InvoiceNo and BillingDate between #" & dtpInvoiceDateFrom.Value & "# And #" & dtpInvoiceDateTo.Value & "# group by Billinfo.InvoiceNo,Customername,BillingDate,productCode,ProductName,Weight,Price,Cartons,TotalPackets,TotalAmount,TotalPayment,GrandTotal,PaymentDue,TaxAmount,TaxPercentage,subtotal,Packets order by BillingDate"
            MyCommand.CommandType = CommandType.Text
            myDA.SelectCommand = MyCommand
            myDA.Fill(myDS, "BillInfo")
            myDA.Fill(myDS, "ProductSold")
            rpt.SetDataSource(myDS)

            CrystalReportViewer2.ReportSource = rpt
            myConnection.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        dtpInvoiceDateFrom.Text = Today
        dtpInvoiceDateTo.Text = Today
        CrystalReportViewer2.ReportSource = Nothing
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        CrystalReportViewer1.ReportSource = Nothing
        cmbCustomerName.Text = ""
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Try
            If cmbCustomerName.Text = "" Then
                MessageBox.Show("Please select Distributor Name", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                cmbCustomerName.Focus()
                Exit Sub
            End If
            Cursor = Cursors.WaitCursor
            Timer1.Enabled = True
            Dim rpt As New rptInvoices() 'The report you created.
            Dim myConnection As OleDbConnection
            Dim MyCommand As New OleDbCommand()
            Dim myDA As New OleDbDataAdapter()
            Dim myDS As New DataSet 'The DataSet you created.


            myConnection = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\SI_DB.accdb;Persist Security Info=False;")
            MyCommand.Connection = myConnection
            MyCommand.CommandText = "SELECT distinct Billinfo.InvoiceNo,BillingDate,Customername,productCode,ProductName,Weight,Price,Cartons,TotalPackets,TotalAmount,TaxPercentage,TaxAmount,GrandTotal,TotalPayment,PaymentDue,subtotal,Packets from BillInfo,ProductSold where BillInfo.InvoiceNo = ProductSold.InvoiceNo and CustomerName= '" & cmbCustomerName.Text & "'order by BillingDate"
            MyCommand.CommandType = CommandType.Text
            myDA.SelectCommand = MyCommand
            myDA.Fill(myDS, "BillInfo")
            myDA.Fill(myDS, "ProductSold")
            rpt.SetDataSource(myDS)

            CrystalReportViewer1.ReportSource = rpt
            myConnection.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
End Class