Imports System.Data.OleDB
Imports Excel = Microsoft.Office.Interop.Excel
Public Class frmSearchProduct1
    Dim rdr As OleDbDataReader = Nothing
    Dim dtable As DataTable
    Dim con As OleDbConnection = Nothing
    Dim adp As OleDbDataAdapter
    Dim ds As DataSet
    Dim cmd As OleDbCommand = Nothing
    Dim dt As New DataTable
    Dim cs As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\SI_DB.accdb;Persist Security Info=False;"

    Private Sub cmbProductName_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbProductName.SelectedIndexChanged
        Try
            GroupBox2.Visible = True
            con = New OleDbConnection(cs)
            con.Open()
            cmd = New OleDbCommand("SELECT (StockID) as [Stock ID],(Product.ProductCode) as [Product Code],(Product.ProductName) as [Product Name],(Product.category) as [Category],(Product.Weight) as [Weight],(Price) as [Unit Price],sum(Cartons) as [Cartons],Packets as [Packets Per Carton],sum(TotalPackets) as [Total Packets]  from product,Stock  where product.productcode=stock.productcode and Product.ProductName= '" & cmbProductName.Text & "' group by Product.ProductCode,Product.Productname,Product.Weight,Product.Category,Product.Price,stockID,Packets order by Product.ProductName", con)
            Dim myDA As OleDbDataAdapter = New OleDbDataAdapter(cmd)
            Dim myDataSet As DataSet = New DataSet()
            myDA.Fill(myDataSet, "Product")
            myDA.Fill(myDataSet, "Stock")
            DataGridView1.DataSource = myDataSet.Tables("Product").DefaultView
            DataGridView1.DataSource = myDataSet.Tables("Stock").DefaultView
            Dim sum As Int64 = 0
            Dim sum1 As Int64 = 0
            For Each r As DataGridViewRow In Me.DataGridView1.Rows
                sum = sum + r.Cells(6).Value
                sum1 = sum1 + r.Cells(8).Value

            Next
            txtC1.Text = sum
            txtP1.Text = sum1

            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub


    Private Sub frmSearchProduct_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        fillCategory()
        fillProductName()
        fillWeight()
    End Sub
    Sub fillWeight()

        Try

            Dim CN As New OleDbConnection(cs)

            CN.Open()
            adp = New OleDbDataAdapter()
            adp.SelectCommand = New OleDbCommand("SELECT distinct  (weight) FROM Stock", CN)
            ds = New DataSet("ds")

            adp.Fill(ds)
            dtable = ds.Tables(0)
            cmbWeight.Items.Clear()
            ComboBox1.Items.Clear()
            For Each drow As DataRow In dtable.Rows
                cmbWeight.Items.Add(drow(0).ToString())
                ComboBox1.Items.Add(drow(0).ToString())
                'DocName.SelectedIndex = -1
            Next
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Sub fillProductName()

        Try

            Dim CN As New OleDbConnection(cs)

            CN.Open()
            adp = New OleDbDataAdapter()
            adp.SelectCommand = New OleDbCommand("SELECT distinct  (ProductName) FROM Stock", CN)
            ds = New DataSet("ds")

            adp.Fill(ds)
            dtable = ds.Tables(0)
            cmbProductName.Items.Clear()
            ComboBox2.Items.Clear()
            For Each drow As DataRow In dtable.Rows
                cmbProductName.Items.Add(drow(0).ToString())
                ComboBox2.Items.Add(drow(0).ToString())
                'DocName.SelectedIndex = -1
            Next
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Sub fillCategory()

        Try

            Dim CN As New OleDbConnection(cs)

            CN.Open()
            adp = New OleDbDataAdapter()
            adp.SelectCommand = New OleDbCommand("SELECT distinct  (Category) FROM stock", CN)
            ds = New DataSet("ds")

            adp.Fill(ds)
            dtable = ds.Tables(0)
            cmbCategory.Items.Clear()

            For Each drow As DataRow In dtable.Rows
                cmbCategory.Items.Add(drow(0).ToString())
                'DocName.SelectedIndex = -1
            Next
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click

        txtProductName.Text = ""
        cmbProductName.Text = ""
        DataGridView1.DataSource = Nothing
        GroupBox2.Visible = False
    End Sub

    Private Sub TabControl1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabControl1.Click

        txtProductName.Text = ""
        cmbProductName.Text = ""
        GroupBox2.Visible = False
        DataGridView1.DataSource = Nothing
        GroupBox5.Visible = False
        cmbWeight.Text = ""
        DataGridView2.DataSource = Nothing
        GroupBox7.Visible = False
        DataGridView3.DataSource = Nothing
        ComboBox1.Text = ""
        ComboBox2.Text = ""
        DataGridView4.DataSource = Nothing
        GroupBox13.Visible = False
        GroupBox16.Visible = False
        cmbCategory.Text = ""
        DataGridView5.DataSource = Nothing
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
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

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        GroupBox5.Visible = False
        cmbWeight.Text = ""
        DataGridView2.DataSource = Nothing
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If DataGridView2.RowCount = Nothing Then
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

            rowsTotal = DataGridView2.RowCount - 1
            colsTotal = DataGridView2.Columns.Count - 1
            With excelWorksheet
                .Cells.Select()
                .Cells.Delete()
                For iC = 0 To colsTotal
                    .Cells(1, iC + 1).Value = DataGridView2.Columns(iC).HeaderText
                Next
                For I = 0 To rowsTotal - 1
                    For j = 0 To colsTotal
                        .Cells(I + 2, j + 1).value = DataGridView2.Rows(I).Cells(j).Value
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

    Private Sub cmbWeight_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbWeight.SelectedIndexChanged
        Try
            GroupBox5.Visible = True
            con = New OleDbConnection(cs)
            con.Open()
            cmd = New OleDbCommand("SELECT (StockID) as [Stock ID],(Product.ProductCode) as [Product Code],(Product.ProductName) as [Product Name],(Product.category) as [Category],(Product.Weight) as [Weight],(Price) as [Unit Price],sum(Cartons) as [Cartons],Packets as [Packets Per Carton],sum(TotalPackets) as [Total Packets]  from product,Stock  where product.productcode=stock.productcode and Product.Weight= '" & cmbWeight.Text & "' group by Product.ProductCode,Product.Productname,Product.Weight,Product.Category,Product.Price,stockID,Packets order by Product.ProductName,Product.Weight", con)
            Dim myDA As OleDbDataAdapter = New OleDbDataAdapter(cmd)
            Dim myDataSet As DataSet = New DataSet()
            myDA.Fill(myDataSet, "Product")
            myDA.Fill(myDataSet, "Stock")
            DataGridView2.DataSource = myDataSet.Tables("Product").DefaultView
            DataGridView2.DataSource = myDataSet.Tables("Stock").DefaultView
            Dim sum As Int64 = 0
            Dim sum1 As Int64 = 0
            For Each r As DataGridViewRow In Me.DataGridView2.Rows
                sum = sum + r.Cells(6).Value
                sum1 = sum1 + r.Cells(8).Value

            Next
            txtC2.Text = sum
            txtP2.Text = sum1

            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub


    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        GroupBox7.Visible = False
        DataGridView3.DataSource = Nothing
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        If DataGridView3.RowCount = Nothing Then
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

            rowsTotal = DataGridView3.RowCount - 1
            colsTotal = DataGridView3.Columns.Count - 1
            With excelWorksheet
                .Cells.Select()
                .Cells.Delete()
                For iC = 0 To colsTotal
                    .Cells(1, iC + 1).Value = DataGridView3.Columns(iC).HeaderText
                Next
                For I = 0 To rowsTotal - 1
                    For j = 0 To colsTotal
                        .Cells(I + 2, j + 1).value = DataGridView3.Rows(I).Cells(j).Value
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

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Try
            GroupBox7.Visible = True
            con = New OleDbConnection(cs)
            con.Open()
            cmd = New OleDbCommand("SELECT (StockID) as [Stock ID],(Product.ProductCode) as [Product Code],(Product.ProductName) as [Product Name],(Product.category) as [Category],(Product.Weight) as [Weight],(Price) as [Unit Price],sum(Cartons) as [Cartons],Packets as [Packets Per Carton],sum(TotalPackets) as [Total Packets]  from product,Stock  where product.productcode=stock.productcode and cartons > 0  group by Product.ProductCode,Product.Productname,Product.Weight,Product.Category,Product.Price,stockID,Packets order by Product.ProductName", con)
            Dim myDA As OleDbDataAdapter = New OleDbDataAdapter(cmd)
            Dim myDataSet As DataSet = New DataSet()
            myDA.Fill(myDataSet, "Product")
            myDA.Fill(myDataSet, "Stock")
            DataGridView3.DataSource = myDataSet.Tables("Product").DefaultView
            DataGridView3.DataSource = myDataSet.Tables("Stock").DefaultView
            Dim sum As Int64 = 0
            Dim sum1 As Int64 = 0
            For Each r As DataGridViewRow In Me.DataGridView3.Rows
                sum = sum + r.Cells(6).Value
                sum1 = sum1 + r.Cells(8).Value

            Next
            txtC3.Text = sum
            txtP3.Text = sum1

            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub txtProductName_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtProductName.TextChanged
        Try
            GroupBox2.Visible = True
            con = New OleDbConnection(cs)
            con.Open()
            cmd = New OleDbCommand("SELECT (StockID) as [Stock ID],(Product.ProductCode) as [Product Code],(Product.ProductName) as [Product Name],(Product.category) as [Category],(Product.Weight) as [Weight],(Price) as [Unit Price],sum(Cartons) as [Cartons],Packets as [Packets Per Carton],sum(TotalPackets) as [Total Packets]  from product,Stock  where product.productcode=stock.productcode and Product.ProductName like '" & txtProductName.Text & "%' group by Product.ProductCode,Product.Productname,Product.Weight,Product.Category,Product.Price,stockID,Packets order by Product.ProductName", con)
            Dim myDA As OleDbDataAdapter = New OleDbDataAdapter(cmd)
            Dim myDataSet As DataSet = New DataSet()
            myDA.Fill(myDataSet, "Product")
            myDA.Fill(myDataSet, "Stock")
            DataGridView1.DataSource = myDataSet.Tables("Product").DefaultView
            DataGridView1.DataSource = myDataSet.Tables("Stock").DefaultView
            Dim sum As Int64 = 0
            Dim sum1 As Int64 = 0
            For Each r As DataGridViewRow In Me.DataGridView1.Rows
                sum = sum + r.Cells(6).Value
                sum1 = sum1 + r.Cells(8).Value

            Next
            txtC1.Text = sum
            txtP1.Text = sum1

            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Try
            GroupBox13.Visible = True
            con = New OleDbConnection(cs)
            con.Open()
            cmd = New OleDbCommand("SELECT (StockID) as [Stock ID],(Product.ProductCode) as [Product Code],(Product.ProductName) as [Product Name],(Product.category) as [Category],(Product.Weight) as [Weight],(Price) as [Unit Price],sum(Cartons) as [Cartons],Packets as [Packets Per Carton],sum(TotalPackets) as [Total Packets]  from product,Stock  where product.productcode=stock.productcode and Product.ProductName = '" & ComboBox2.Text & "' and Product.weight= '" & ComboBox1.Text & "' group by Product.ProductCode,Product.Productname,Product.Weight,Product.Category,Product.Price,stockID,Packets order by Product.ProductName,Product.Weight", con)
            Dim myDA As OleDbDataAdapter = New OleDbDataAdapter(cmd)
            Dim myDataSet As DataSet = New DataSet()
            myDA.Fill(myDataSet, "Product")
            myDA.Fill(myDataSet, "Stock")
            DataGridView4.DataSource = myDataSet.Tables("Product").DefaultView
            DataGridView4.DataSource = myDataSet.Tables("Stock").DefaultView
            Dim sum As Int64 = 0
            Dim sum1 As Int64 = 0
            For Each r As DataGridViewRow In Me.DataGridView4.Rows
                sum = sum + r.Cells(6).Value
                sum1 = sum1 + r.Cells(8).Value

            Next
            txtC4.Text = sum
            txtP4.Text = sum1

            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        If DataGridView4.RowCount = Nothing Then
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

            rowsTotal = DataGridView4.RowCount - 1
            colsTotal = DataGridView4.Columns.Count - 1
            With excelWorksheet
                .Cells.Select()
                .Cells.Delete()
                For iC = 0 To colsTotal
                    .Cells(1, iC + 1).Value = DataGridView4.Columns(iC).HeaderText
                Next
                For I = 0 To rowsTotal - 1
                    For j = 0 To colsTotal
                        .Cells(I + 2, j + 1).value = DataGridView4.Rows(I).Cells(j).Value
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

    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
        ComboBox1.Text = ""
        ComboBox2.Text = ""
        DataGridView4.DataSource = Nothing
        GroupBox13.Visible = False
    End Sub

    Private Sub Button12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button12.Click
        GroupBox16.Visible = False
        cmbCategory.Text = ""
        DataGridView5.DataSource = Nothing

    End Sub

    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click
        If DataGridView5.RowCount = Nothing Then
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

            rowsTotal = DataGridView5.RowCount - 1
            colsTotal = DataGridView5.Columns.Count - 1
            With excelWorksheet
                .Cells.Select()
                .Cells.Delete()
                For iC = 0 To colsTotal
                    .Cells(1, iC + 1).Value = DataGridView5.Columns(iC).HeaderText
                Next
                For I = 0 To rowsTotal - 1
                    For j = 0 To colsTotal
                        .Cells(I + 2, j + 1).value = DataGridView5.Rows(I).Cells(j).Value
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

    Private Sub cmbCategory_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbCategory.SelectedIndexChanged
        Try
            GroupBox16.Visible = True
            con = New OleDbConnection(cs)
            con.Open()
            cmd = New OleDbCommand("SELECT (StockID) as [Stock ID],(Product.ProductCode) as [Product Code],(Product.ProductName) as [Product Name],(Product.category) as [Category],(Product.Weight) as [Weight],(Price) as [Unit Price],sum(Cartons) as [Cartons],Packets as [Packets Per Carton],sum(TotalPackets) as [Total Packets]  from product,Stock  where product.productcode=stock.productcode and Product.Category= '" & cmbCategory.Text & "' group by Product.ProductCode,Product.Productname,Product.Weight,Product.Category,Product.Price,stockID,Packets order by Product.ProductName,Product.Category", con)
            Dim myDA As OleDbDataAdapter = New OleDbDataAdapter(cmd)
            Dim myDataSet As DataSet = New DataSet()
            myDA.Fill(myDataSet, "Product")
            myDA.Fill(myDataSet, "Stock")
            DataGridView5.DataSource = myDataSet.Tables("Product").DefaultView
            DataGridView5.DataSource = myDataSet.Tables("Stock").DefaultView
            Dim sum As Int64 = 0
            Dim sum1 As Int64 = 0
            For Each r As DataGridViewRow In Me.DataGridView5.Rows
                sum = sum + r.Cells(6).Value
                sum1 = sum1 + r.Cells(8).Value

            Next
            txtC5.Text = sum
            txtP5.Text = sum1

            con.Close()
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
            frmSales.txtProductCode.Text = dr.Cells(1).Value()
            frmSales.txtProductName.Text = dr.Cells(2).Value()
            frmSales.txtWeight.Text = dr.Cells(4).Value()
            frmSales.txtPrice.Text = dr.Cells(5).Value()
            frmSales.txtAvailableCartons.Text = dr.Cells(6).Value()
            frmSales.txtAvailablePackets.Text = dr.Cells(8).Value()
            frmSales.txtCartons.Focus()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub DataGridView2_RowHeaderMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView2.RowHeaderMouseClick
        Try
            Dim dr As DataGridViewRow = DataGridView2.SelectedRows(0)
            Me.Hide()
            frmSales.Show()
            ' or simply use column name instead of index
            'dr.Cells["id"].Value.ToString();
            frmSales.txtProductCode.Text = dr.Cells(1).Value()
            frmSales.txtProductName.Text = dr.Cells(2).Value()
            frmSales.txtWeight.Text = dr.Cells(4).Value()
            frmSales.txtPrice.Text = dr.Cells(5).Value()
            frmSales.txtAvailableCartons.Text = dr.Cells(6).Value()
            frmSales.txtAvailablePackets.Text = dr.Cells(8).Value()
            frmSales.txtCartons.Focus()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub DataGridView3_RowHeaderMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView3.RowHeaderMouseClick
        Try
            Dim dr As DataGridViewRow = DataGridView3.SelectedRows(0)
            Me.Hide()
            frmSales.Show()
            ' or simply use column name instead of index
            'dr.Cells["id"].Value.ToString();
            frmSales.txtProductCode.Text = dr.Cells(1).Value()
            frmSales.txtProductName.Text = dr.Cells(2).Value()
            frmSales.txtWeight.Text = dr.Cells(4).Value()
            frmSales.txtPrice.Text = dr.Cells(5).Value()
            frmSales.txtAvailableCartons.Text = dr.Cells(6).Value()
            frmSales.txtAvailablePackets.Text = dr.Cells(8).Value()
            frmSales.txtCartons.Focus()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub DataGridView4_RowHeaderMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView4.RowHeaderMouseClick
        Try
            Dim dr As DataGridViewRow = DataGridView4.SelectedRows(0)
            Me.Hide()
            frmSales.Show()
            ' or simply use column name instead of index
            'dr.Cells["id"].Value.ToString();
            frmSales.txtProductCode.Text = dr.Cells(1).Value()
            frmSales.txtProductName.Text = dr.Cells(2).Value()
            frmSales.txtWeight.Text = dr.Cells(4).Value()
            frmSales.txtPrice.Text = dr.Cells(5).Value()
            frmSales.txtAvailableCartons.Text = dr.Cells(6).Value()
            frmSales.txtAvailablePackets.Text = dr.Cells(8).Value()
            frmSales.txtCartons.Focus()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub DataGridView5_RowHeaderMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView5.RowHeaderMouseClick
        Try
            Dim dr As DataGridViewRow = DataGridView5.SelectedRows(0)
            Me.Hide()
            frmSales.Show()
            ' or simply use column name instead of index
            'dr.Cells["id"].Value.ToString();
            frmSales.txtProductCode.Text = dr.Cells(1).Value()
            frmSales.txtProductName.Text = dr.Cells(2).Value()
            frmSales.txtWeight.Text = dr.Cells(4).Value()
            frmSales.txtPrice.Text = dr.Cells(5).Value()
            frmSales.txtAvailableCartons.Text = dr.Cells(6).Value()
            frmSales.txtAvailablePackets.Text = dr.Cells(8).Value()
            frmSales.txtCartons.Focus()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
End Class