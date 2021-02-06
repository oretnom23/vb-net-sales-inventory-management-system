Imports System.Data.OleDB
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Public Class frmRegistrationReport

    Private Sub frmRegistrationReport_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try

            Dim rpt As New rptRegistration() 'The report you created.
            Dim myConnection As OleDbConnection
            Dim MyCommand As New OleDbCommand()
            Dim myDA As New OleDbDataAdapter()
            Dim myDS As New SI_DBDataSet 'The DataSet you created.


            myConnection = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\SI_DB.accdb;Persist Security Info=False;")
            MyCommand.Connection = myConnection
            MyCommand.CommandText = "SELECT * from Registration "
            MyCommand.CommandType = CommandType.Text
            myDA.SelectCommand = MyCommand
            myDA.Fill(myDS, "Registration")

            rpt.SetDataSource(myDS)
            CrystalReportViewer1.ReportSource = rpt
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
End Class