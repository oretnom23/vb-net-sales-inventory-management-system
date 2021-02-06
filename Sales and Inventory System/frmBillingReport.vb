Imports System.Data.OleDb
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Public Class frmBillingReport

    Private Sub frmBillingReport_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try

            Dim rpt As New rptInvoice() 'The report you created.
            Dim myConnection As OleDbConnection
            Dim MyCommand As New OleDbCommand()
            Dim myDA As New OleDbDataAdapter()
            Dim myDS As New DataSet 'The DataSet you created.


            myConnection = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\SI_DB.accdb;Persist Security Info=False;")
            myConnection.Open()
            MyCommand.Connection = myConnection
            MyCommand.CommandText = "select Billinfo.Invoiceno,BillingDate,ProductCode,ProductName,Weight,Price,Cartons,Packets,TotalPackets,TotalAmount,SubTotal,TaxPercentage,TaxAmount,GrandTotal,TotalPayment,PaymentDue,B_name,B_address,B_Landmark,B_city,B_state,B_zipcode,S_name,S_address,S_landmark,S_city,S_state,S_zipcode,Phone,email,MobileNo from Billinfo,ProductSold,Customer where Billinfo.InvoiceNo=ProductSold.InvoiceNo and Customer.CustomerNo=Billinfo.CustomerNo and Billinfo.Invoiceno= '" & frmSales.txtInvoiceNo.Text & "'"
            MyCommand.CommandType = CommandType.Text
            myDA.SelectCommand = MyCommand
            myDA.Fill(myDS, "BillInfo")
            myDA.Fill(myDS, "ProductSold")
            myDA.Fill(myDS, "Customer")
            rpt.SetDataSource(myDS)
            CrystalReportViewer1.ReportSource = rpt
            myConnection.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub


End Class
