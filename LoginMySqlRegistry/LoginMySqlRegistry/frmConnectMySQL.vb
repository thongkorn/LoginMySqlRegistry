#Region "About"
' / --------------------------------------------------------------------------------
' / Developer : Mr.Surapon Yodsanga (Thongkorn Tubtimkrob)
' / eMail : thongkorn@hotmail.com
' / URL: http://www.g2gnet.com (Khon Kaen - Thailand)
' / Facebook: https://www.facebook.com/g2gnet (For Thailand)
' / Facebook: https://www.facebook.com/commonindy (Worldwide)
' / More Info: http://www.g2gnet.com/webboard
' / 
' / Purpose: Design Login Form with Microsoft Visual Basic PowerPacks.
' /             Connection test MySQL Server and registry with VB.NET (2010)
' /
' / Microsoft Visual Basic .NET (2010)
' /
' / This is open source code under @CopyLeft by Thongkorn Tubtimkrob.
' / You can modify and/or distribute without to inform the developer.
' / --------------------------------------------------------------------------------
#End Region

Imports MySql.Data.MySqlClient
'// For Registry System.
Imports Microsoft.Win32

Public Class frmConnectMySQL

    ' / --------------------------------------------------------------------------------
    ' / S T A R T ... H E R E
    Private Sub frmConnectMySQL_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Me.CenterToScreen()
        Try
            Call ReadRegistry()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            End
        End Try
    End Sub

    ' / --------------------------------------------------------------------------------
    Private Sub ReadRegistry()
        '// Retrieve data from the registry.
        '// Computer\HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\MyProgram\LoginMySQL
        Dim oRegKey As RegistryKey = Registry.LocalMachine.OpenSubKey("SOFTWARE\MyProgram", True)
        '// If nothing then create subkey registry.
        If oRegKey Is Nothing Then
            '/ First time to create SubKey.
            oRegKey = Registry.LocalMachine.CreateSubKey("SOFTWARE\MyProgram\LoginMySQL")
            '/ Default Value.
            With oRegKey
                .SetValue("Host", "localhost")
                .SetValue("DataBaseName", "contact")
                .SetValue("DBUsername", "admin")
                .SetValue("DBPassword", "admin")
            End With
            'MsgBox("Registry key added.")
        Else
            '/ Get value from SubKey.
            oRegKey = Registry.LocalMachine.OpenSubKey("SOFTWARE\MyProgram\LoginMySQL")
            txtHost.Text = oRegKey.GetValue("Host", 0)
            txtDBName.Text = oRegKey.GetValue("DataBaseName", 0)
            txtDBUserName.Text = oRegKey.GetValue("DBUsername", 0)
            txtDBPassword.Text = oRegKey.GetValue("DBPassword", 0)
        End If
        oRegKey.Close()
        oRegKey.Dispose()
    End Sub

    ' / --------------------------------------------------------------------------------
    Private Sub SaveRegistry()
        '// Retrieve data from other part of the registry.
        Dim oRegKey As RegistryKey = Registry.LocalMachine.OpenSubKey("SOFTWARE\MyProgram", True)
        If oRegKey Is Nothing Then oRegKey = Registry.LocalMachine.CreateSubKey("SOFTWARE\MyProgram\LoginMySQL")
        '/ Save Registry.
        oRegKey = Registry.LocalMachine.OpenSubKey("SOFTWARE\MyProgram\LoginMySQL", True)
        With oRegKey
            .SetValue("Host", txtHost.Text.Trim)
            .SetValue("DataBaseName", txtDBName.Text.Trim)
            .SetValue("DBUsername", txtDBUserName.Text.Trim)
            .SetValue("DBPassword", txtDBPassword.Text.Trim)
        End With
        oRegKey.Close()
        oRegKey.Dispose()
    End Sub

    '/ Check MySQL Server connection.
    Private Sub btnConnect_Click(sender As System.Object, e As System.EventArgs) Handles btnConnect.Click
        If Trim(txtHost.Text.Trim.Length) = 0 Then
            MessageBox.Show("Enter your DNS or IP Address.", "Report status", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            txtHost.Focus()
            Return
        ElseIf Trim(txtDBName.Text.Trim.Length) = 0 Then
            MessageBox.Show("Enter your DataBase Name.", "Report status", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            txtDBName.Focus()
            Return
        ElseIf Trim(txtDBUserName.Text.Trim.Length) = 0 Then
            MessageBox.Show("Enter your DataBase Username.", "Report status", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            txtDBUserName.Focus()
            Return
        ElseIf Trim(txtDBPassword.Text.Trim.Length) = 0 Then
            MessageBox.Show("Enter your DataBase Password.", "Report status", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            txtDBPassword.Focus()
            Return
        End If

        '// Call to ConnectMySQLServer ... modDatabase.vb
        '// Send Host, DataBasename, DataBase UserName/Password for Login to MySQL Server.
        If ConnectMySQLServer(Trim(txtHost.Text), txtDBName.Text.Trim, txtDBUserName.Text.Trim, txtDBPassword.Text.Trim) Then
            MsgBox("Connection to MySQL Server successful.")
            '// Save them to registry.
            Call SaveRegistry()
            '// Open Main Form and hidden login form.
            '/ frmMain.ShowDialog()
            '/ Me.Hide()
            If Conn.State = ConnectionState.Open Then Conn.Close()
            Me.Close()
        End If
    End Sub

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Me.Close()
    End Sub

    Private Sub frmConnectDB_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Me.Dispose()
        GC.SuppressFinalize(Me)
        Application.Exit()
    End Sub

    ' / --------------------------------------------------------------------------------
    ' / Not use now.
    Private Function LoginSystem() As Boolean
        If Conn.State = ConnectionState.Closed Then Conn.Open()
        '//
        Cmd = New MySqlCommand( _
            " SELECT * FROM tbluser WHERE " & _
            " UserName = @UNAME AND Password = @PWD ", Conn)

        Dim UserNameParam As New MySqlParameter("@UNAME", Me.txtUsername.Text)
        Dim PasswordParam As New MySqlParameter("@PWD", Me.txtPassword.Text)

        Cmd.Parameters.Add(UserNameParam)
        Cmd.Parameters.Add(PasswordParam)

        DR = Cmd.ExecuteReader()
        '// Found data
        If DR.HasRows Then
            MessageBox.Show("You can logged into system.")
            LoginSystem = True
        Else
            LoginSystem = False
            MessageBox.Show("Enter your User Name, Password is incorrect.")
            txtUsername.Focus()
        End If
        DR.Close()
        Cmd.Connection.Close()
    End Function

End Class
