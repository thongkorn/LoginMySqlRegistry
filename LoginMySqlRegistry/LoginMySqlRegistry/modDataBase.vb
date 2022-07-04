' / --------------------------------------------------------------------------------
' / Developer : Mr.Surapon Yodsanga (Thongkorn Tubtimkrob)
' / eMail : thongkorn@hotmail.com
' / URL: http://www.g2gnet.com (Khon Kaen - Thailand)
' / Facebook: https://www.facebook.com/g2gnet (For Thailand)
' / Facebook: https://www.facebook.com/commonindy (Worldwide)
' / More Info: http://www.g2gnet.com/webboard
' / 
' / Purpose: This module is part of the database system and declare the variable as public.
' /
' / Microsoft Visual Basic .NET (2010)
' /
' / This is open source code under @CopyLeft by Thongkorn Tubtimkrob.
' / You can modify and/or distribute without to inform the developer.
' / --------------------------------------------------------------------------------

Imports MySql.Data.MySqlClient
Imports Microsoft.VisualBasic

Module modDataBase
    '// Declare variable one time but use many times.
    Public Conn As MySqlConnection
    Public Cmd As MySqlCommand
    Public DS As DataSet
    Public DR As MySqlDataReader
    Public DA As MySqlDataAdapter
    Public strSQL As String '// Major SQL
    Public strStmt As String    '// Minor SQL
    '//
    ' / --------------------------------------------------------------------------------
    '// Connect to MySQL Server and return true/false for successful or not.
    Public Function ConnectMySQLServer(Server As String, ByVal DBName As String, ByVal UID As String, PWD As String) As Boolean
        Dim strCon As String = _
            " Server=" & Server & "; " & _
            " Database=" & DBName & "; " & _
            " User ID=" & UID & "; " & _
            " Password=" & PWD & "; " & _
            " Port = 3306;" & _
            " CharSet = utf8; " & _
            " Connect Timeout = 90; " & _
            " Pooling = True; " & _
            " Persist Security Info = False; "
        Conn = New MySqlConnection
        Conn.ConnectionString = strCon
        Try
            Conn.Open()
            Return True
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Report Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return False
            'End
        End Try
    End Function

    ' / --------------------------------------------------------------------------------
    ' / Get my project path
    ' / AppPath = C:\My Project\bin\debug
    ' / Replace "\bin\debug" with "\"
    ' / Return : C:\My Project\
    Function MyPath(AppPath As String) As String
        '/ Return Value
        MyPath = AppPath.ToLower.Replace("\bin\debug", "\").Replace("\bin\release", "\").Replace("\bin\x86\debug", "\")
        '// If not found folder then put the \ (BackSlash has ASCII Code = 92) at the end.
        If Right(MyPath, 1) <> Chr(92) Then MyPath = MyPath & Chr(92)
    End Function
End Module
