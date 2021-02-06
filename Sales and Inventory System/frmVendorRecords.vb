Imports System.Data.OleDB
Imports Excel = Microsoft.Office.Interop.Excel
Public Class frmVendorRecords
    Dim rdr As OleDbDataReader = Nothing
    Dim con As OleDbConnection = Nothing
    Dim cmd As OleDbCommand = Nothing
    Dim cs As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\SI_DB.accdb;Persist Security Info=False;"
    Private ReadOnly Property Connection() As OleDbConnection
        Get
            Dim ConnectionToFetch As New OleDbConnection(cs)
            ConnectionToFetch.Open()
            Return ConnectionToFetch
        End Get
    End Property
    Public Function GetData() As DataView
        Dim SelectQry = "SELECT (vendorID)as [Vendor ID],(name) as [Name],(address) as [Address],(landmark) as [Landmark],(city) as [City],(state) as [State],(zipcode) as [Zip/Post Code],(Phone) as [Phone],(email) as [Email],(mobileno) as [Mobile No],(faxno) as [Fax No],(notes) as [Notes] from vendor order by vendorid"

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



    Private Sub frmVendorRecords_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        DataGridView1.DataSource = GetData()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        DataGridView1.DataSource = Nothing
    End Sub


    Private Sub DataGridView1_RowHeaderMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView1.RowHeaderMouseClick
        Try
            Dim dr As DataGridViewRow = DataGridView1.SelectedRows(0)
            Me.Hide()
            frmVendor.Show()
            ' or simply use column name instead of index
            'dr.Cells["id"].Value.ToString();
            frmVendor.txtVendorID.Text = dr.Cells(0).Value.ToString()
            frmVendor.txtName.Text = dr.Cells(1).Value.ToString()

            frmVendor.txtAddress.Text = dr.Cells(2).Value.ToString()
            frmVendor.txtCity.Text = dr.Cells(4).Value.ToString()
            frmVendor.txtLandmark.Text = dr.Cells(3).Value.ToString()
            frmVendor.cmbState.Text = dr.Cells(5).Value.ToString()
            frmVendor.txtZipCode.Text = dr.Cells(6).Value.ToString()
            frmVendor.txtPhone.Text = dr.Cells(7).Value.ToString()
            frmVendor.txtEmail.Text = dr.Cells(8).Value.ToString()
            frmVendor.txtMobileNo.Text = dr.Cells(9).Value.ToString()
            frmVendor.txtFaxNo.Text = dr.Cells(10).Value.ToString()
            frmVendor.txtNotes.Text = dr.Cells(11).Value.ToString()
            frmVendor.Update_Record.Enabled = True
            frmVendor.Delete.Enabled = True
            frmVendor.Save.Enabled = False
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        If DataGridView1.RowCount = Nothing Then
            MessageBox.Show("Sorry nothing to export into excel sheet.." & vbCrLf & "Please retrieve data in datagridview", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If
        Dim rowsTotal, colsTotal As Short
        Dim I, j, iC As Short
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        Dim xlApp As New Excel.Application

        Try
            Dim excelBook As Excel.Workbook = xlApp.Workbooks.Add
            Dim excelWorksheet As Excel.Worksheet = CType(excelBook.Worksheets(1), Excel.Worksheet)
            xlApp.Visible = True

            rowsTotal = DataGridView1.RowCount - 1
            colsTotal = DataGridView1.Columns.Count - 1
            With excelWorksheet
                .Cells.Select()
                .Cells.Delete()
                For iC = 0 To colsTotal
                    .Cells(1, iC + 1).Value = DataGridView1.Columns(iC).HeaderText
                Next
                For I = 0 To rowsTotal - 1
                    For j = 0 To colsTotal
                        .Cells(I + 2, j + 1).value = DataGridView1.Rows(I).Cells(j).Value
                    Next j
                Next I
                .Rows("1:1").Font.FontStyle = "Bold"
                .Rows("1:1").Font.Size = 12

                .Cells.Columns.AutoFit()
                .Cells.Select()
                .Cells.EntireColumn.AutoFit()
                .Cells(1, 1).Select()
            End With
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            'RELEASE ALLOACTED RESOURCES
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
            xlApp = Nothing
        End Try
    End Sub
End Class