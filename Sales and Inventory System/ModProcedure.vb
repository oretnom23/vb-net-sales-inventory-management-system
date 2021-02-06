Imports System.Data.OleDB

Module ModProcedure
    Public con As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\SI_DB.accdb;Persist Security Info=False;")

    Public Sub FillListView(ByRef lvList As ListView, ByRef myData As OleDbDataReader)
        Dim itmListItem As ListViewItem

        Dim strValue As String

        Do While myData.Read
            itmListItem = New ListViewItem()
            strValue = IIf(myData.IsDBNull(0), "", myData.GetValue(0))
            itmListItem.Text = strValue

            For shtCntr = 1 To myData.FieldCount() - 1
                If myData.IsDBNull(shtCntr) Then
                    itmListItem.SubItems.Add("")
                Else
                    itmListItem.SubItems.Add(myData.GetValue(shtCntr))
                End If
            Next shtCntr

            lvList.Items.Add(itmListItem)
        Loop
    End Sub
    Public Sub SetConnection(Optional ByVal Firstcomp As Boolean = False)
        Dim str As String
        str = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\SI_DB.accdb;Persist Security Info=False;"

        Try
            If IsNothing(con) = False Then
                If con.State = ConnectionState.Closed Then
                    con.Close()
                End If
            End If
            con = New OleDbConnection(str)
            con.Open()
        Catch ex As System.Exception
            MsgBox(ex.Message)
            MsgBox("Not Connecting to Database Server.Please check your network.")
        End Try
    End Sub
End Module

