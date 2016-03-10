'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Class: Login.vb
'Function:
'Purpose: The purpose of this class is to connect and authenticate against the conservative.ca
'        domain and login with admin credentials to elevate the program.
'
'Returns: 
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Imports System.DirectoryServices
Public Class Login


    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Function: Login_KeyPress
    'Purpose: Calling this close the program and threads when the escape key is pressed
    '
    'Returns: 
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Private Sub Login_KeyPress(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode.Equals(Keys.Escape) Then
            Application.Exit()
        End If
    End Sub


    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Function: AuthenticateUser
    'Passing: The LDAP connection string, the username, and password
    'Purpose: Calling this will call the authenicateuser function with passing through the 
    '         values typed by the user
    'Returns: TRUE or FALSE
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Function AuthenticateUser(path As String, user As String, pass As String) As Boolean
        Dim de As New DirectoryEntry(path, user, pass, AuthenticationTypes.Secure)
        Try
            'run a search using those credentials.  
            'If it returns anything, then you're authenticated
            Dim ds As DirectorySearcher = New DirectorySearcher(de)
            ds.FindOne()
            '   ds.SearchScope = SearchScope.Subtree
            '    ds.Filter = "(&(objectClass=user)(sAMAccountName=" + user + ")(memberOf='CN=Domain Admins,OU=Admins,DC=conservative,DC=ca'))"
            '
            '      Dim result As SearchResult = ds.FindOne
            '     If result Is Nothing Then
            '    MsgBox("Denied")
            'Return False
            '   Else

            '   Return True

            '   End If

            If My.User.IsInRole(
                    ApplicationServices.BuiltInRole.Administrator) Then
                ' Insert code to access a resource here.
                Me.Hide()
                '  CIMS.Show()
            Else
                MsgBox("You are not a Local Admin on this machine")
            End If

            Return True

        Catch
            'otherwise, it will crash out so return false
            MsgBox("Denied - Wrong Username or Password")
            Return False
        End Try
    End Function

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Function: btnLogin_Click
    'Purpose: Calling this will call the authenicateuser function with passing through the 
    '         values typed by the user
    'Returns: 
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Private Sub btnLogin_Click(sender As Object, e As EventArgs) Handles btnLogin.Click
        '   AuthenticateUser("LDAP://cpcdc01.conservative.ca", txtUsername.Text, txtPassword.Text)


        Shell("runas /user:" & txtUsername.Text & "\" & txtPassword.Text & " ""D:\CIMS Account Simplifier.exe""", vbNormalFocus)

    End Sub

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Function: txtUsername_Keydown
    'Purpose: Calling this close the program and threads when the escape key is pressed
    '
    'Returns: 
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Private Sub txtUsername_KeyDown(sender As Object, e As KeyEventArgs) Handles txtUsername.KeyDown
        If e.KeyCode.Equals(Keys.Escape) Then
            Application.Exit()
        End If
    End Sub

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Function: txtPassword_Keydown
    'Purpose: Calling this close the program and threads when the escape key is pressed
    '
    'Returns: 
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Private Sub txtPassword_KeyDown(sender As Object, e As KeyEventArgs) Handles txtPassword.KeyDown
        If e.KeyCode.Equals(Keys.Escape) Then
            Application.Exit()
        End If
    End Sub


End Class