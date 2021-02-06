Imports System.Data.OleDB
Public Class frmInVoiceNo
    Dim rdr As OleDbDataReader = Nothing
    Dim dtable As DataTable
    Dim con As OleDbConnection = Nothing
    Dim adp As OleDbDataAdapter
    Dim ds As DataSet
    Dim cmd As OleDbCommand = Nothing
    Dim dt As New DataTable
    Dim cs As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\SI_DB.accdb;Persist Security Info=False;"

    Private Sub frmInVoiceNo_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load


        Try

            Dim CN As New OleDbConnection(cs)

            CN.Open()
            adp = New OleDbDataAdapter()
            adp.SelectCommand = New OleDbCommand("SELECT distinct (InvoiceNo) FROM billinfo", CN)
            ds = New DataSet("ds")

            adp.Fill(ds)
            dtable = ds.Tables(0)
            cmbInvoiceNo.Items.Clear()

            For Each drow As DataRow In dtable.Rows
                cmbInvoiceNo.Items.Add(drow(0).ToString())
                'DocName.SelectedIndex = -1
            Next
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub


    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Close()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Cursor = Cursors.WaitCursor
        Timer1.Enabled = True
        Try
          
            Dim rpt As New rptInvoice() 'The report you created.
            Dim myConnection As OleDbConnection
            Dim MyCommand As New OleDbCommand()
            Dim myDA As New OleDbDataAdapter()
            Dim myDS As New DataSet 'The DataSet you created.


            myConnection = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\SI_DB.accdb;Persist Security Info=False;")
            myConnection.Open()
            MyCommand.Connection = myConnection
            MyCommand.CommandText = "select Billinfo.Invoiceno,BillingDate,ProductCode,ProductName,Weight,Price,Cartons,Packets,TotalPackets,TotalAmount,SubTotal,TaxPercentage,TaxAmount,GrandTotal,TotalPayment,PaymentDue,B_name,B_address,B_Landmark,B_city,B_state,B_zipcode,S_name,S_address,S_landmark,S_city,S_state,S_zipcode,Phone,email,MobileNo from Billinfo,ProductSold,Customer where Billinfo.InvoiceNo=ProductSold.InvoiceNo and Customer.CustomerNo=Billinfo.CustomerNo and Billinfo.Invoiceno= '" & cmbInvoiceNo.Text & "'"
            MyCommand.CommandType = CommandType.Text
            myDA.SelectCommand = MyCommand
            myDA.Fill(myDS, "BillInfo")
            myDA.Fill(myDS, "ProductSold")
            myDA.Fill(myDS, "Customer")
            rpt.SetDataSource(myDS)
            frmInvoice.CrystalReportViewer1.ReportSource = rpt
            frmInvoice.Show()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Cursor = Cursors.Default
        Timer1.Enabled = False
    End Sub
End Class