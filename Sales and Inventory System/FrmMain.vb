Imports System.Data.OleDb
Imports System.IO
Public Class FrmMain
    Dim sSql As String
    Dim rdr As OleDbDataReader = Nothing
    Dim con As OleDbConnection = Nothing
    Dim cmd As OleDbCommand = Nothing
    Dim cs As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\SI_DB.accdb;Persist Security Info=False;"

    Private Sub CalculatorToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CalculatorToolStripMenuItem.Click
        Try
            System.Diagnostics.Process.Start("Calc.exe")
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub NotepadToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NotepadToolStripMenuItem.Click
        Try
            System.Diagnostics.Process.Start("Notepad.exe")
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub TaskManagerToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TaskManagerToolStripMenuItem.Click
        Try
            System.Diagnostics.Process.Start("TaskMgr.exe")
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub MSWordToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MSWordToolStripMenuItem.Click
        Try
            System.Diagnostics.Process.Start("Winword.exe")
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub SystemInfoToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SystemInfoToolStripMenuItem.Click
        frmSystemInfo.Show()
    End Sub

    Private Sub FrmMain_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        End
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

        Dim SelectQry = "SELECT (StockID) as [Stock ID],(ProductCode) as [Product Code],(ProductName) as [Product Name],(Weight) as [Weight],sum(Cartons) as [Cartons],Packets,Sum(TotalPackets) as [Total Packets] FROM stock where Cartons > 0 and TotalPackets > 0   group by StockID, ProductCode,ProductName,Weight,Packets order by ProductName "
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
    Private Sub FrmMain_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
     
        ToolStripStatusLabel4.Text = Now

        ToolStripStatusLabel2.Text = frmLogin.UserName.Text
        Me.Refresh()
        Button1.PerformClick()
        Timer2.Start()
        Timer2.Interval = 1000

    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Timer1.Start()
        ToolStripStatusLabel4.Text = Now
    End Sub

    Private Sub RegistrationToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RegistrationToolStripMenuItem.Click
        Me.Hide()
        frmRegistration.Show()
    End Sub

    Private Sub LoginDetailsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LoginDetailsToolStripMenuItem.Click
        frmLoginDetails.Show()
    End Sub

    Private Sub AboutToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AboutToolStripMenuItem.Click
        frmAbout.Show()
    End Sub

    Private Sub WordpadToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles WordpadToolStripMenuItem.Click
        Try
            System.Diagnostics.Process.Start("Wordpad.exe")
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub VendorToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles VendorToolStripMenuItem.Click
        Me.Hide()
        frmVendor.Show()
    End Sub

    Private Sub CustomerToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CustomerToolStripMenuItem.Click
        Me.Hide()
        frmCustomer.Show()
    End Sub

    Private Sub InvoiceToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles InvoiceToolStripMenuItem.Click
        Me.Hide()
        frmSales.Show()
    End Sub

    Private Sub ProductToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ProductToolStripMenuItem.Click
        Me.Hide()
        frmProduct.Show()
    End Sub

    Private Sub ProfileEntryToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ProfileEntryToolStripMenuItem.Click
        Me.Hide()
        frmCustomer.Show()
    End Sub

    Private Sub ProfileEntryToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ProfileEntryToolStripMenuItem1.Click
        Me.Hide()
        frmVendor.Show()
    End Sub

    Private Sub CategoryToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CategoryToolStripMenuItem.Click
        Me.Hide()
        frmInventoryCategory.Show()
    End Sub


    Private Sub StockToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles StockToolStripMenuItem1.Click
        Me.Hide()
        frmStock.Show()
    End Sub

    Private Sub SalesToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SalesToolStripMenuItem1.Click
        Me.Hide()
        frmSales.Show()
    End Sub



    Private Sub VendorsToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles VendorsToolStripMenuItem1.Click
        Me.Hide()
        frmVendorRecords1.Show()
    End Sub

    Private Sub CustomersToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CustomersToolStripMenuItem1.Click
        Me.Hide()
        frmCustomersRecord1.Show()
    End Sub



    Private Sub RegistrationToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RegistrationToolStripMenuItem1.Click
        Me.Hide()
        frmRegistration.Show()
    End Sub




    Private Sub SalesToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SalesToolStripMenuItem.Click
        Me.Hide()
        frmSalesRecord.Show()
    End Sub

    Private Sub LogoutToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LogoutToolStripMenuItem.Click
        Me.Hide()
        frmLogin.Show()
        frmLogin.UserName.Text = ""
        frmLogin.Password.Text = ""
        frmLogin.UserName.Focus()
        frmLogin.ProgressBar1.Visible = False
    End Sub

    Private Sub RegistrationToolStripMenuItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RegistrationToolStripMenuItem2.Click
        frmRegistrationReport.Show()
    End Sub

    Private Sub SalesToolStripMenuItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SalesToolStripMenuItem2.Click
        Me.Hide()
        frmSalesReport.CrystalReportViewer1.ReportSource = Nothing
        frmSalesReport.dtpInvoiceDateFrom.Text = Today
        frmSalesReport.dtpInvoiceDateTo.Text = Today
        frmSalesReport.CrystalReportViewer2.ReportSource = Nothing
        frmSalesReport.DateTimePicker1.Text = Today
        frmSalesReport.DateTimePicker2.Text = Today
        frmSalesReport.CrystalReportViewer3.ReportSource = Nothing


        frmSalesReport.Show()
    End Sub



    Private Sub OrderToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OrderToolStripMenuItem.Click
        Me.Hide()
        frmOrder.Show()
    End Sub


    Private Sub InvoiceToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles InvoiceToolStripMenuItem1.Click
        frmInVoiceNo.Show()
    End Sub



    Private Sub OrderToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OrderToolStripMenuItem1.Click
        Me.Hide()
        frmOrderRecord.DataGridView1.DataSource = Nothing
        frmOrderRecord.dtpOrderDateFrom.Text = Today
        frmOrderRecord.dtpOrderDateTo.Text = Today
        frmOrderRecord.DataGridView6.DataSource = Nothing
        frmOrderRecord.cmbCustomerName.Text = ""
        frmOrderRecord.DataGridView3.DataSource = Nothing
        frmOrderRecord.cmbOrderNo.Text = ""
        frmOrderRecord.DataGridView2.DataSource = Nothing
        frmOrderRecord.cmbOrderNo.Text = ""
        frmOrderRecord.cmbProductName.Text = ""
        frmOrderRecord.DataGridView5.DataSource = Nothing
        frmOrderRecord.Show()
    End Sub



    Private Sub OrderToolStripMenuItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OrderToolStripMenuItem2.Click
        Me.Hide()

        frmOrdersReport.cmbCustomerName.Text = ""
        frmOrdersReport.CrystalReportViewer3.ReportSource = Nothing
        frmOrdersReport.dtpOrderDateFrom.Text = Today
        frmOrdersReport.dtpOrderDateFrom.Text = Today
        frmOrdersReport.CrystalReportViewer1.ReportSource = Nothing
        frmOrdersReport.cmbStatus.Text = ""
        frmOrdersReport.CrystalReportViewer2.ReportSource = Nothing
        frmOrdersReport.Show()
    End Sub
  
  




    Private Sub Timer2_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer2.Tick
        Button1.PerformClick()

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        DataGridView1.DataSource = GetData()
    End Sub

    Private Sub ProductMasterEntryToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ProductMasterEntryToolStripMenuItem.Click
        Me.Hide()
        frmProduct.Show()
    End Sub

    Private Sub OrdersToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OrdersToolStripMenuItem.Click
        Me.Hide()
        frmOrder.Show()
    End Sub

    Private Sub StockToolStripMenuItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles StockToolStripMenuItem2.Click
        Me.Hide()
        frmStock.Show()
    End Sub


    Private Sub MasterToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MasterToolStripMenuItem.Click
        Me.Hide()

        frmProductsRecord1.DataGridView4.DataSource = Nothing
        frmProductsRecord1.cmbCategory.Text = ""
        frmProductsRecord1.cmbWeight.Text = ""
        frmProductsRecord1.DataGridView3.DataSource = Nothing
        frmProductsRecord1.cmbProductName.Text = ""
        frmProductsRecord1.txtProduct.Text = ""
        frmProductsRecord1.DataGridView2.DataSource = Nothing
        frmProductsRecord1.DataGridView1.DataSource = Nothing
        frmProductsRecord1.Show()
    End Sub

    Private Sub StockToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles StockToolStripMenuItem.Click
        Me.Hide()

        frmStockDetails.cmbProductName.Text = ""
        frmStockDetails.txtProduct.Text = ""
        frmStockDetails.DataGridView2.DataSource = Nothing
        frmStockDetails.cmbCategory.Text = ""
        frmStockDetails.DataGridView3.DataSource = Nothing
        frmStockDetails.cmbWeight.Text = ""
        frmStockDetails.DataGridView4.DataSource = Nothing

        frmStockDetails.Show()
    End Sub



    Private Sub InvoicesToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles InvoicesToolStripMenuItem.Click
        Me.Hide()
        frmInVoicesReport.Show()
    End Sub
End Class