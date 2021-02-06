Imports System.Data.OleDB
Public Class frmLogin

    Private Sub frmLogin_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        End
    End Sub

    ' TODO: Insert code to perform custom authentication using the provided username and password 
    ' (See http://go.microsoft.com/fwlink/?LinkId=35339).  
    ' The custom principal can then be attached to the current thread's principal as follows: 
    '     My.User.CurrentPrincipal = CustomPrincipal
    ' where CustomPrincipal is the IPrincipal implementation used to perform authentication. 
    ' Subsequently, My.User will return identity information encapsulated in the CustomPrincipal object
    ' such as the username, display name, etc.









    Private Sub frmLogin_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub



    Private Sub LinkLabel2_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs)

    End Sub

    Private Sub login_button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles login_button.Click
        If Len(Trim(UserName.Text)) = 0 Then
            MessageBox.Show("Please enter user name", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            UserName.Focus()
            Exit Sub
        End If
        If Len(Trim(Password.Text)) = 0 Then
            MessageBox.Show("Please enter password", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Password.Focus()
            Exit Sub
        End If
        Try
            Dim myConnection As OleDbConnection
            myConnection = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\SI_DB.accdb;Persist Security Info=False;")

            Dim myCommand As OleDbCommand

            myCommand = New OleDbCommand("SELECT Username,user_password FROM Users WHERE username = @username AND user_password = @UserPassword", myConnection)

            Dim uName As New OleDbParameter("@username", SqlDbType.VarChar)

            Dim uPassword As New OleDbParameter("@UserPassword", SqlDbType.VarChar)

            uName.Value = UserName.Text

            uPassword.Value = Password.Text

            myCommand.Parameters.Add(uName)

            myCommand.Parameters.Add(uPassword)

            myCommand.Connection.Open()

            Dim myReader As OleDbDataReader = myCommand.ExecuteReader(CommandBehavior.CloseConnection)

            Dim Login As Object = 0

            If myReader.HasRows Then

                myReader.Read()

                Login = myReader(Login)

            End If

            If Login = Nothing Then

                MsgBox("Login is Failed...Try again !", MsgBoxStyle.Critical, "Login Denied")
                UserName.Clear()
                Password.Clear()
                UserName.Focus()

            Else

                ProgressBar1.Visible = True
                ProgressBar1.Maximum = 5000
                ProgressBar1.Minimum = 0
                ProgressBar1.Value = 4
                ProgressBar1.Step = 1

                For i = 0 To 5000
                    ProgressBar1.PerformStep()
                Next

                FrmMain.ToolStripStatusLabel2.Text = UserName.Text
                Me.Hide()

                FrmMain.Show()



            End If
            myCommand.Dispose()
            myConnection.Close()




        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        End
    End Sub

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Me.Hide()
        frmChangePassword.Show()
        frmChangePassword.UserName.Text = ""
        frmChangePassword.OldPassword.Text = ""
        frmChangePassword.NewPassword.Text = ""
        frmChangePassword.ConfirmPassword.Text = ""
        frmChangePassword.UserName.Focus()
    End Sub

    Private Sub LinkLabel2_LinkClicked_1(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel2.LinkClicked
        Me.Hide()
        frmPasswordRecovery.Show()
    End Sub
End Class
