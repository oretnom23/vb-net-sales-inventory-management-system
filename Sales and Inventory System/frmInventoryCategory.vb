Imports System.Data.OleDB
Public Class frmInventoryCategory
    Dim rdr As OleDbDataReader = Nothing
    Dim con As OleDbConnection = Nothing
    Dim cmd As OleDbCommand = Nothing
    Dim cs As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\SI_DB.accdb;Persist Security Info=False;"

    Sub clear()
        txtCategoryID.Text = ""
        txtCategoryName.Text = ""
    End Sub
    Private Sub btnNewRecord_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNewRecord.Click
        txtCategoryID.Text = ""
        txtCategoryName.Text = ""
        btnSave.Enabled = True
        btnDelete.Enabled = False
        btnUpdate_record.Enabled = False
        txtCategoryName.Focus()
    End Sub
    Private Sub auto()

        con = New OleDbConnection(cs)

        con.Open()

        Dim ct As String = "select Max(CategoryID) as CatID from InventoryCategory"

        cmd = New OleDbCommand(ct)
        cmd.Connection = con
        rdr = cmd.ExecuteReader()
        rdr.Read()

        If rdr("CatID").ToString() <> "" Then
            txtCategoryID.Text = Integer.Parse(rdr("CatID").ToString()) + 1
        Else
            txtCategoryID.Text = 1
        End If
        con.Close()
        rdr.Close()
    End Sub
    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        If Len(Trim(txtCategoryName.Text)) = 0 Then
            MessageBox.Show("Please enter category name", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtCategoryName.Focus()
            Exit Sub
        End If


        Try
            auto()
            con = New OleDbConnection(cs)
            con.Open()
            Dim ct As String = "select CategoryName from InventoryCategory where CategoryName=@find"

            cmd = New OleDbCommand(ct)
            cmd.Connection = con
            cmd.Parameters.Add(New OleDbParameter("@find", System.Data.OleDb.OleDbType.VarChar, 150, "CategoryName"))
            cmd.Parameters("@find").Value = txtCategoryName.Text
            rdr = cmd.ExecuteReader()

            If rdr.Read Then
                MessageBox.Show("Category name Already Exists", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                txtCategoryName.Text = ""
                txtCategoryName.Focus()
                If Not rdr Is Nothing Then
                    rdr.Close()
                End If

            Else



                con = New OleDbConnection(cs)
                con.Open()

                Dim cb As String = "insert into InventoryCategory(CategoryID,CategoryName) VALUES (@d1,@d2)"

                cmd = New OleDbCommand(cb)

                cmd.Connection = con

                cmd.Parameters.Add(New OleDbParameter("@d1", System.Data.OleDB.OleDBType.VarChar, 10, "VendorID"))
                cmd.Parameters.Add(New OleDbParameter("@d2", System.Data.OleDB.OleDBType.VarChar, 150, "CategoryName"))

                cmd.Parameters("@d1").Value = txtCategoryID.Text
                cmd.Parameters("@d2").Value = txtCategoryName.Text


                cmd.ExecuteReader()
                MessageBox.Show("Successfully saved", "Category Details", MessageBoxButtons.OK, MessageBoxIcon.Information)
                btnSave.Enabled = False
                autocomplete()
                If con.State = ConnectionState.Open Then
                    con.Close()
                End If

                con.Close()
            End If



        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnUpdate_record_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdate_record.Click
        Try

            con = New OleDbConnection(cs)
            con.Open()

            Dim cb As String = "update InventoryCategory set CategoryName='" & txtCategoryName.Text & "' where CategoryID= '" & txtCategoryID.Text & "'"

            cmd = New OleDbCommand(cb)
            cmd.Connection = con
          
            cmd.ExecuteReader()
            MessageBox.Show("Successfully updated", "Category Details", MessageBoxButtons.OK, MessageBoxIcon.Information)

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            con.Close()

            btnUpdate_record.Enabled = False
            autocomplete()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub delete_records()
        Try



            Dim RowsAffected As Integer = 0
            con = New OleDbConnection(cs)

            con.Open()
            Dim ct As String = "select Category from Product where Category=@find"


            cmd = New OleDbCommand(ct)

            cmd.Connection = con
            cmd.Parameters.Add(New OleDbParameter("@find", System.Data.OleDB.OleDBType.VarChar, 150, "Category"))


            cmd.Parameters("@find").Value = txtCategoryName.Text


            rdr = cmd.ExecuteReader()

            If rdr.Read Then
                MessageBox.Show("Unable to delete..Already in use", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)

                clear()
                txtCategoryName.Focus()
                btnUpdate_record.Enabled = False
                btnDelete.Enabled = False
                If Not rdr Is Nothing Then
                    rdr.Close()
                End If
                Exit Sub
            End If

            con = New OleDbConnection(cs)

            con.Open()


            Dim cq As String = "delete from InventoryCategory where CategoryID=@DELETE1;"


            cmd = New OleDbCommand(cq)

            cmd.Connection = con

            cmd.Parameters.Add(New OleDbParameter("@DELETE1", System.Data.OleDB.OleDBType.VarChar, 10, "CategoryID"))


            cmd.Parameters("@DELETE1").Value = Trim(txtCategoryID.Text)
            RowsAffected = cmd.ExecuteNonQuery()
            If RowsAffected > 0 Then

                MessageBox.Show("Successfully deleted", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information)

                clear()
                txtCategoryName.Focus()
                btnUpdate_record.Enabled = False
                btnDelete.Enabled = False
                autocomplete()
            Else
                MessageBox.Show("No record found", "Sorry", MessageBoxButtons.OK, MessageBoxIcon.Information)

                clear()
                txtCategoryName.Focus()
                btnUpdate_record.Enabled = False
                btnDelete.Enabled = False
                autocomplete()


                If con.State = ConnectionState.Open Then

                    con.Close()
                End If

                con.Close()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Try



            If MessageBox.Show("Do you really want to delete the record?", "Inventory Record", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) = Windows.Forms.DialogResult.Yes Then
                delete_records()



            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnGetDetails_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGetDetails.Click
        frmInventoryCategoryRecord.DataGridView1.DataSource = Nothing
        frmInventoryCategoryRecord.Show()
    End Sub



    Private Sub Button1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Hide()
        clear()
        frmInventoryCategoryRecord.DataGridView1.DataSource = Nothing
        frmInventoryCategoryRecord.Show()
    End Sub

    Private Sub frmInventoryCategory_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Me.Hide()
        FrmMain.Show()
    End Sub


    Sub autocomplete()
        Try
            con = New OleDbConnection(cs)
            con.Open()

            Dim cmd As New OleDbCommand("SELECT CategoryName FROM InventoryCategory", con)
            Dim ds As New DataSet
            Dim da As New OleDbDataAdapter(cmd)
            da.Fill(ds, "My List") 'list can be any name u want

            Dim col As New AutoCompleteStringCollection
            Dim i As Integer
            For i = 0 To ds.Tables(0).Rows.Count - 1
                col.Add(ds.Tables(0).Rows(i)("Categoryname").ToString())

            Next
            txtCategoryName.AutoCompleteSource = AutoCompleteSource.CustomSource
            txtCategoryName.AutoCompleteCustomSource = col
            txtCategoryName.AutoCompleteMode = AutoCompleteMode.Suggest

            con.Close()

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub frmInventoryCategory_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        autocomplete()
        txtCategoryName.Focus()
    End Sub
End Class