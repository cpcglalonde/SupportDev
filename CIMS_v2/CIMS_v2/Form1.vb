'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Class: Form1.vb
'Function:
'Purpose: This class drives the whole CIMS/ AD/ C2G account creation.
'         Parses through the block copied from Zendesk, and splits it to all the 
'         appropriate fields.
'
'Returns: 
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



''''''''''''''''''''''''''''''''''''''''''''''''
'
'Import Statements
' 
''''''''''''''''''''''''''''''''''''''''''''''''
Imports System.DirectoryServices
Imports System.DirectoryServices.ActiveDirectory
Imports System.Text
Imports System.IO
Imports System.Management.Automation
Imports System.Collections.ObjectModel
Imports System.Management.Automation.Runspaces
Imports System.Runtime.InteropServices



Public Class lblWorkingFor

    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto)>
    Private Shared Function SendMessage(ByVal hWnd As IntPtr, ByVal Msg As UInteger, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As IntPtr
    End Function

    <DllImport("User32.dll", SetLastError:=True, CharSet:=CharSet.Auto)>
    Private Shared Function EnumChildWindows(ByVal WindowHandle As IntPtr, ByVal Callback As EnumWindowProcess, ByVal lParam As IntPtr) As Boolean
    End Function


    ''''''''''''''''''''''''''''''''''''''''''''''''
    '
    ' Declaration Statements
    ' 
    ''''''''''''''''''''''''''''''''''''''''''''''''
    Const SW_SHOWNORMAL = 1



    Dim strAlphaNumeric As String
    Dim strChar As String
    Dim CleanedString As String = ""
    Dim myString As String
    Dim myName As String
    Dim num As Integer
    Dim trunc As String
    Dim ridingdash As String
    Dim lower As String
    Dim C2GmyString As String
    Dim C2GmyName As String
    Dim C2GfirstName As String
    Dim C2GLastName As String
    Dim CIMSFirstName As String
    Dim CIMSLastName As String
    Dim myUserName As String
    Dim writeup As String
    Dim writeupFrench As String
    Dim replaceSubmitterName As String
    Dim finalPassword As String
    Dim replaceSubmitterEmail As String
    Dim replaceName As String
    Dim replaceEmail As String
    Dim replacePhone As String
    Dim replaceRiding As String
    Dim replacesomething As String
    Dim replaceworkingfor As String
    Dim replaceworkingwhere As String
    Dim replaceaccesslevel As String
    Dim doubleReplaceRiding As String
    Dim replaceDescRiding As String
    Dim AD_FullName As String
    Dim AD_GivenName As String
    Dim AD_Surname As String
    Dim AD_AccountName As String
    Dim AD_Password As String
    Dim AD_OU As String
    Dim CreateAccCounter As Integer
    Dim AD_VPN As String = "Add-ADGroupMember -Identity ""vpn"" -Members "
    Dim AD_REMOTE As String = "Add-ADGroupMember -Identity ""Remote TS CIMS Access"" -Members "
    Dim AD_UPNSuffix As String = "$Suffix = """
    Dim CompleteUPNSuffix As String = "$CompleteUPN= "
    Dim ADMoveUser As String
    Dim ADLocation As String
    Dim ADEnabled As String
    Dim VPNCheck As String = "IsMember "
    Dim ADEnableAccount As String = "Enable-ADAccount -Identity "
    Dim ADResetPassword As String
    Dim strInt As String
    Dim final As String
    Dim FirstLetter As String
    Dim MiddleLetter As String
    Dim MPPath As String
    Dim EDAPath As String
    Dim ridingnum As String
    Dim Counter As Integer
    Dim LastNumber As Integer
    Dim ADSetPath As String
    Dim sender As Object

    Dim e As System.EventArgs

    Dim AD_DC As String = "DC=conservative,DC=ca"
    Dim AD_PATH As String
    Dim securepass As String
    Dim PowershellScript As String
    '  Dim ChildrenList As New List(Of IntPtr)
    Dim displayname As String

    Dim finalRidingNumber As String
    Dim Powershell_Thread As System.Threading.Thread
    Dim Powershell_Thread_exist As System.Threading.Thread


    ''''''''''''''''''''''''''''''''''''''''''''''''
    '
    ' Declaration Statements -- DATABASE -- 
    ' 
    ''''''''''''''''''''''''''''''''''''''''''''''''
    Dim OwnerOrgID As SByte = 1
    Dim Active As SByte = 1
    Dim Notes As String
    Dim SecurityLevel As Integer = 1
    Dim ScopeLevel As Integer
    Dim CashBoxNo As Integer?
    Dim ext As String = "20"
    Dim RidingNoMPBoundaries As String = "0"
    Dim PreferredLanguageID As Integer?
    Dim ExportToFilePermitted As SByte = 0
    Dim UseMPBoundaries As SByte?
    Dim LastModified As Date
    Dim CreatedDate As Date




    <DllImport("USER32.DLL")>
    Private Shared Function GetShellWindow() As IntPtr
    End Function

    <DllImport("USER32.DLL")>
    Private Shared Function GetWindowText(ByVal hWnd As IntPtr, ByVal lpString As StringBuilder, ByVal nMaxCount As Integer) As Integer
    End Function

    <DllImport("USER32.DLL")>
    Private Shared Function GetWindowTextLength(ByVal hWnd As IntPtr) As Integer
    End Function

    <DllImport("user32.dll", SetLastError:=True)>
    Private Shared Function GetWindowThreadProcessId(ByVal hWnd As IntPtr, <Out()> ByRef lpdwProcessId As UInt32) As UInt32
    End Function

    <DllImport("USER32.DLL")>
    Private Shared Function IsWindowVisible(ByVal hWnd As IntPtr) As Boolean
    End Function

    Private Delegate Function EnumWindowsProc(ByVal hWnd As IntPtr, ByVal lParam As Integer) As Boolean

    <DllImport("USER32.DLL")>
    Private Shared Function EnumWindows(ByVal enumFunc As EnumWindowsProc, ByVal lParam As Integer) As Boolean
    End Function

    Private hShellWindow As IntPtr = GetShellWindow()
    Private dictWindows As New Dictionary(Of IntPtr, String)
    Private currentProcessID As Integer





    '
    '
    '
    Private Sub AddNewUser()

        Dim dc As New DataSQL

        '  dc.ExecuteCommand("INSERT INTO Test (Password) VALUES (" + finalPassword + """)")



        Try
            dc.ExecuteCommand("INSERT INTO Users (OwnerOrgID, UserID, Password, FirstName, LastName, Title, email, Phone, ext, Active, Notes, SecurityLevel, ScopeLevel, CashBoxNo, RidingNo, RidingNoMPBoundaries, PreferredLanguageID, ExportToFilePermitted, UseMPBoundaries, LastModified, Created) values('" & OwnerOrgID & "','" & convertQuotes(myString) & "','" & convertQuotes(finalPassword) & "','" & convertQuotes(CIMSFirstName) & "','" & convertQuotes(CIMSLastName) & "','" & convertQuotes(replaceworkingfor) & "','" & convertQuotes(replaceEmail) & "','" & convertQuotes(replacePhone) & "','" & convertQuotes(ext) & "','" & Active & "','" & convertQuotes(Notes) & "','" & SecurityLevel & "','" & ScopeLevel & "','" & CashBoxNo & "','" & finalRidingNumber & "','" & convertQuotes(RidingNoMPBoundaries) & "','" & PreferredLanguageID & "','" & ExportToFilePermitted & "','" & UseMPBoundaries & "','" & convertQuotes(LastModified) & "','" & convertQuotes(CreatedDate) & "')")

            MsgBox("User " & CIMSFirstName & " Successfully Entered into Database")
        Catch ex As Exception
            MsgBox("Wasn't able to insert into Datebase: Error: " & vbCrLf & vbCrLf & " " & ex.Message)
        End Try

        '  Dim test As String

        '   Test = "INSERT INTO Users (OwnerOrgID, UserID, Password, FirstName, LastName, Title, email, Phone, ext, Active, Notes, SecurityLevel, ScopeLevel, CashBoxNo, RidingNo, RidingNoMPBoundaries, PreferredLanguageID, ExportToFilePermitted, UseMPBoundaries, LastModified, Created) values('" & OwnerOrgID & "','" & convertQuotes(myString) & "','" & convertQuotes(finalPassword) & "','" & convertQuotes(CIMSFirstName) & "','" & convertQuotes(CIMSLastName) & "','" & convertQuotes(replaceworkingfor) & "','" & convertQuotes(replaceEmail) & "','" & convertQuotes(replacePhone) & "','" & convertQuotes(ext) & "','" & Active & "','" & convertQuotes(Notes) & "','" & SecurityLevel & "','" & ScopeLevel & "','" & CashBoxNo & "','" & finalRidingNumber & "','" & convertQuotes(RidingNoMPBoundaries) & "','" & PreferredLanguageID & "','" & ExportToFilePermitted & "','" & UseMPBoundaries & "')"
        '& convertQuotes(LastModified) & "','" & convertQuotes(CreatedDate)
        '    MsgBox(test)

        dc.SubmitChanges()
        'dc.ExecuteCommand("INSERT INTO Test (Password) values('" & convertQuotes(finalPassword) & "')")

        '  Dim addme = From Users In dc.Users


    End Sub

    Public Function convertQuotes(ByVal str As String) As String
        convertQuotes = str.Replace("'", "''")
    End Function

    Sub LoadProcess()
        ComboBox1.Items.Clear()
        '  For Each p As Process In Process.GetProcesses

        ' Try
        'If p.MainWindowTitle.Length <> 0 Then
        'ComboBox1.Items.Add(p.MainWindowTitle & " - " & CStr(p.Id))

        '   ComboBox1.Items.Add(ChildrenList)
        'End If
        'Catch ex As Exception
        '
        'End Try
        'Next
        'If ComboBox1.Items.Count <> 0 Then ComboBox1.SelectedIndex = 0


        For Each P As Process In Process.GetProcesses
            'Get a list of ALL of the open windows associated with the process   
            Dim windows As IDictionary(Of IntPtr, String) = GetOpenWindowsFromPID(P.Id)
            For Each kvp As KeyValuePair(Of IntPtr, String) In windows
                'This small if statement lets us ignore the start menu...
                If kvp.Value.ToLower = "start" = False Then

                    ComboBox1.Items.Add(P.ProcessName & " - " & kvp.Value & " - PID- " & P.Id)
                End If
            Next
        Next

    End Sub


    Sub getprocess()
        ComboBox1.Items.Clear()

        For Each poc In Process.GetProcesses
            If poc.MainWindowTitle.Length > 1 Then

                Dim windows As IDictionary(Of IntPtr, String) = GetOpenWindowsFromPID(poc.Id)


                For Each kvp As KeyValuePair(Of IntPtr, String) In windows
                    poc.EnableRaisingEvents = True

                    Try
                        ComboBox1.Items.Add(" " & kvp.ToString & " - " & poc.Id)
                    Catch ex As Exception
                        '  MsgBox(poc.ProcessName.ToString & " " & ex.Message)
                    End Try
                Next

            End If
            AddHandler poc.Exited, AddressOf OnProcessExit
        Next


    End Sub



    Public Function GetOpenWindowsFromPID(ByVal processID As Integer) As IDictionary(Of IntPtr, String)
        dictWindows.Clear()
        currentProcessID = processID
        EnumWindows(AddressOf enumWindowsInternal, 0)
        Return dictWindows
    End Function

    Private Function enumWindowsInternal(ByVal hWnd As IntPtr, ByVal lParam As Integer) As Boolean
        If (hWnd <> hShellWindow) Then
            Dim windowPid As UInt32
            If Not IsWindowVisible(hWnd) Then
                Return True
            End If
            Dim length As Integer = GetWindowTextLength(hWnd)
            If (length = 0) Then
                Return True
            End If
            GetWindowThreadProcessId(hWnd, windowPid)
            If (windowPid <> currentProcessID) Then
                Return True
            End If
            Dim stringBuilder As New StringBuilder(length)
            GetWindowText(hWnd, stringBuilder, (length + 1))
            dictWindows.Add(hWnd, stringBuilder.ToString)
        End If
        Return True
    End Function
    Private Sub OnProcessExit(ByVal sender As Object, ByVal e As EventArgs)
        ' Avoid late binding and Cast the Object to a Process. 
        Dim poc As Process = CType(sender, Process)
        ' Write the ExitCode 
        'ListBox4.Items.Add(p.ExitCode)
        ' Write the ExitTime 
        MsgBox(poc.ExitTime.ToString)


    End Sub









    ''
    '
    '

    Public Shared Function GetChildWindows(ByVal ParentHandle As IntPtr) As IntPtr()
        Dim ChildrenList As New List(Of IntPtr)
        Dim ListHandle As GCHandle = GCHandle.Alloc(ChildrenList)
        Try
            EnumChildWindows(ParentHandle, AddressOf EnumWindow, GCHandle.ToIntPtr(ListHandle))
        Finally
            If ListHandle.IsAllocated Then ListHandle.Free()
        End Try

        Return ChildrenList.ToArray

    End Function

    Public Delegate Function EnumWindowProcess(ByVal Handle As IntPtr, ByVal Parameter As IntPtr) As Boolean
    Private Shared Function EnumWindow(ByVal Handle As IntPtr, ByVal Parameter As IntPtr) As Boolean
        Dim ChildrenList As List(Of IntPtr) = GCHandle.FromIntPtr(Parameter).Target

        If ChildrenList Is Nothing Then Throw New Exception("GCHandle Target could not be cast as List(Of IntPtr)")
        ChildrenList.Add(Handle)
        Return True
    End Function

    Sub RemovePwd(ByVal input As String)

        Dim p As Process = Process.GetProcessById(input.Substring(input.LastIndexOf("-") + 2))

        For Each I As IntPtr In GetChildWindows(p.MainWindowHandle)
            SendMessage(I, &HCC, 0, 0)
            Invalidate(I)

        Next
    End Sub
    Private Function ExecuteMyPowerShellEnabled(ByVal scriptText As String) As String
        Dim RunspaceSample As Runspace = RunspaceFactory.CreateRunspace()

        RunspaceSample.Open()
        Dim PipelineSample As Pipeline = RunspaceSample.CreatePipeline()

        PipelineSample.Commands.AddScript(scriptText)
        Dim results As Collection(Of PSObject) = PipelineSample.Invoke()





        For Each obj As PSObject In results

            Try
                displayname = obj.Members("Enabled").Value.ToString
            Catch ex As Exception
                displayname = Nothing
            End Try


        Next
        RunspaceSample.Close()

        Return scriptText
    End Function
    Private Function ExecuteMyPowerShellScriptLocation(ByVal scriptText As String) As String
        Dim RunspaceSample As Runspace = RunspaceFactory.CreateRunspace()

        RunspaceSample.Open()
        Dim PipelineSample As Pipeline = RunspaceSample.CreatePipeline()

        PipelineSample.Commands.AddScript(scriptText)
        Dim results As Collection(Of PSObject) = PipelineSample.Invoke()




        For Each obj As PSObject In results

            Try
                displayname = obj.Members("DistinguishedName").Value.ToString
            Catch ex As Exception
                displayname = Nothing
            End Try


        Next
        RunspaceSample.Close()

        Return scriptText

    End Function


    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Function: ExecuteMyPowerShellScript
    'Purpose: When called, this function opens a line to powershell, and uses the string passed 
    '        through to run a script
    'Passes: ScriptText
    'Returns: scriptText
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Private Function ExecuteMyPowerShellScript(ByVal scriptText As String) As String
        Dim RunspaceSample As Runspace = RunspaceFactory.CreateRunspace()

        RunspaceSample.Open()
        Dim PipelineSample As Pipeline = RunspaceSample.CreatePipeline()

        PipelineSample.Commands.AddScript(scriptText)
        Dim results As Collection(Of PSObject) = PipelineSample.Invoke()

        RunspaceSample.Close()

        Return scriptText
    End Function


    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Function: GenerateRandomString
    'Purpose: When called, this function generates a random string between a-z thats 6 characters
    '           Lowercase
    'Passes: Integer, Boolean
    'Returns: IIF
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Public Function GenerateRandomString(ByRef len As Integer, ByRef upper As Boolean) As String
        Dim rand As New Random()
        Dim allowableChars() As Char = "abcdefghijklmnopqrstuvwxyz".ToCharArray()
        Dim final As String = String.Empty
        For i As Integer = 0 To len - 1
            final += allowableChars(rand.Next(allowableChars.Length - 1))
        Next

        Return IIf(upper, final.ToUpper(), final)
    End Function

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Function: GenerateRandomString
    'Purpose: When called, this function generates a random string between A-Z thats 1 character
    '           Uppercase
    'Passes: Integer, Boolean
    'Returns: IIF
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Public Function GenerateRandomStringUpper(ByRef len As Integer, ByRef upper As Boolean) As String
        Dim rand As New Random()
        Dim allowableChars() As Char = "SQUKJHYLEOTCVROXAPIGMNWZBFD".ToCharArray()
        Dim final As String = String.Empty
        For i As Integer = 0 To len - 1
            final += allowableChars(rand.Next(allowableChars.Length - 1))
        Next

        Return IIf(upper, final.ToUpper(), final)
    End Function


    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Function: CleanTheString
    'Purpose: When called, this function will accept the uncleaned string, and strip all characters
    '         that aren't part of the alphabet
    'Passes: theString
    'Returns: CleanedString
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Function CleanTheString(theString)


        'msgbox thestring
        strAlphaNumeric = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ" 'Used to check for numeric characters.
        For i = 1 To Len(theString)
            strChar = Mid(theString, i, 1)
            If InStr(strAlphaNumeric, strChar) Then
                CleanedString += strChar
            End If
        Next
        'msgbox cleanedstring
        CleanTheString = CleanedString
    End Function

    Private Sub txtInfo_Leave(sender As Object, e As EventArgs) Handles txtInfo.Leave

        Dim valid As Integer = 12


        Try


            Dim strText() As String
            strText = Split(txtInfo.Text, vbCrLf)


            Dim submitterName As String = strText(0)
            Dim submitterEmail As String = strText(1)
            Dim Name As String = strText(3)
            Dim Email As String = strText(4)
            Dim Phone As String = strText(5)

            Dim Riding As String = strText(7)
            Dim something As String = strText(8)
            Dim workingfor As String = strText(10)
            Dim workingwhere As String = strText(11)
            Dim accesslevel As String = strText(12)

            If strText(0).StartsWith("Submitter Name: ") Then


                replaceSubmitterName = Replace(submitterName, "Submitter Name: ", "")

                replaceSubmitterEmail = Replace(submitterEmail, "Submitter E-mail Address: ", "")
                replaceName = Replace(Name, "Name: ", "")
                replaceEmail = Replace(Email, "E-mail Address: ", "")
                replacePhone = Replace(Phone, "Phone Number: ", "")
                replaceRiding = Replace(Riding, "Request Type: ", "")
                replacesomething = Replace(something, "Organization Name / ED: ", "")
                replaceworkingwhere = Replace(workingwhere, "Working Where: ", "")


                If (workingfor.Contains("other")) Then
                    replaceworkingfor = Replace(workingfor, "Working For: other: ", "")
                Else

                    replaceworkingfor = Replace(workingfor, "Working For: ", "")
                End If



                replaceaccesslevel = Replace(accesslevel, "Access Level: ", "")

                replaceSubmitterName = LTrim(replaceSubmitterName)
                replaceSubmitterName = RTrim(replaceSubmitterName)

                replaceSubmitterEmail = LTrim(replaceSubmitterEmail)
                replaceSubmitterEmail = RTrim(replaceSubmitterEmail)

                replaceName = LTrim(replaceName)
                replaceName = RTrim(replaceName)

                replaceEmail = LTrim(replaceEmail)
                replaceEmail = RTrim(replaceEmail)

                replacePhone = LTrim(replacePhone)
                replacePhone = RTrim(replacePhone)

                replaceRiding = LTrim(replaceRiding)
                replaceRiding = RTrim(replaceRiding)

                replacesomething = LTrim(replacesomething)
                replacesomething = RTrim(replacesomething)

                replaceworkingfor = LTrim(replaceworkingfor)
                replaceworkingfor = RTrim(replaceworkingfor)

                replaceaccesslevel = LTrim(replaceaccesslevel)
                replaceaccesslevel = RTrim(replaceaccesslevel)

                txtSubmitterName.Text = replaceSubmitterName
                txtSubmitterEmail.Text = replaceSubmitterEmail

                txtName.Text = UnAccent(replaceName)
                txtEmail.Text = replaceEmail

                txtPhone.Text = replacePhone

                txtRequest.Text = replaceRiding
                txtRiding.Text = replacesomething
                txtWorking.Text = replaceworkingfor
                txtAccess.Text = replaceaccesslevel
                txtWorkingWhere.Text = replaceworkingwhere


                doubleReplaceRiding = Replace(replacesomething, "—", "-")
                replaceDescRiding = Replace(doubleReplaceRiding, "-", "")


                txtRiding.Text = doubleReplaceRiding

                myName = replaceName

                If (replacesomething = "--") Then
                    MsgBox("Error: Riding Cannot be NULL -> Submitter Must Re-Submit Application")
                    Exit Sub

                End If
                finalRidingNumber = replacesomething.Substring(0, replacesomething.IndexOf(" "))


                CIMSFirstName = myName.Substring(0, myName.IndexOf(" "))
                CIMSLastName = myName.Substring(myName.IndexOf(" ") + 1)




                myUserName = txtName.Text
                txtName.Text = UnAccent(myUserName)

                myString = CleanTheString(myUserName)
                myUserName = ""
                If myString.Length >= 20 Then
                    num = myString.Length - 20
                    trunc = myString.Substring(0, myString.Length - num)
                    lower = LCase(trunc)
                    myString = lower
                Else
                    lower = LCase(myString)
                    myString = lower
                End If
                '   myString = Replace(Replace(Replace(myUserName, " ", ""), "-", ""), ".", "")



                txtTruncated.Text = UnAccent(myString)
                'My.Computer.Clipboard.SetText(myString)

                FirstLetter = GenerateRandomStringUpper(1, True)
                MiddleLetter = GenerateRandomString(6, False)
                Randomize()
                LastNumber = CInt(Math.Floor((9 - 1 + 1) * Rnd())) + 1
                final = FirstLetter + MiddleLetter & LastNumber
                finalPassword = FirstLetter + MiddleLetter & LastNumber
                txtPassword.Text = FirstLetter + MiddleLetter & LastNumber
                txtCIMSFirst.Text = CIMSFirstName
                txtCIMSLast.Text = CIMSLastName
                txtResult.Text = ("Username: " + UnAccent(myString) & Environment.NewLine & "Password: " + final & Environment.NewLine & Environment.NewLine + "AD & CIMS")


                If (replaceaccesslevel = "Data Entry User") Then
                    ScopeLevel = 5
                ElseIf (replaceaccesslevel = "Riding User")
                    ScopeLevel = 3
                ElseIf (replaceaccesslevel = "Admin User")
                    ScopeLevel = 2
                End If



                '  txtAccess.Text += " (" + SecurityLevel + "-" + ScopeLevel + ")"
                txtInfo.Enabled = False


                btnAD.Enabled = True
                btnCIMS.Enabled = false
                btnNewHire.Enabled = True
                btnData.Enabled = True


                If Riding.Contains("Deactivate Account") Then
                    MsgBox("Are you sure? -- Deactivate the account", MsgBoxStyle.Information)

                End If

            End If

        Catch ex As System.IndexOutOfRangeException
            MsgBox("You Must enter in the data from Zendesk")
        End Try
    End Sub





    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Function: French
    'Purpose: Calling this will generate the french writeup to send to the client in an email
    '
    'Returns: writeupFrench
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Private Function French() As String
        Dim firstName As String = myName.Substring(0, myName.IndexOf(" "))
        Dim EmailName_French As String = "Bonjour  " + firstName + ", "
        writeupFrench = EmailName_French + txtWriteupFrench.Text
        My.Computer.Clipboard.SetText(writeupFrench)
        MsgBox("Successfully copied to clipboard.  (╯°□°）╯︵ ┻━┻")
        txtInfo.Enabled = True
        txtLog.Clear()

        clear()
        Return writeupFrench

    End Function



    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Function: Clear
    'Purpose: This function clears all textboxes that have information in them
    '
    'Returns: 0
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Private Function clear()

        CleanedString = Nothing
        myString = Nothing
        txtName.Text = Nothing
        txtPassword.Text = Nothing
        txtCIMSFirst.Text = Nothing
        txtAccess.Text = Nothing
        txtEmail.Text = Nothing
        txtPhone.Text = Nothing
        txtLog.Text = Nothing
        txtRiding.Text = Nothing
        txtSubmitterEmail.Text = Nothing
        txtSubmitterName.Text = Nothing
        txtRequest.Text = Nothing
        txtWorking.Text = Nothing
        txtTruncated.Text = Nothing
        txtResult.Text = Nothing
        txtCIMSLast.Text = Nothing
        Return 0

    End Function

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Function: English
    'Purpose: Calling this will generate the english writeup to send to the client in an email
    '
    'Returns: Writeup
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Private Function English() As String
        Dim firstName As String = myName.Substring(0, myName.IndexOf(" "))
        Dim EmailName As String = "Hello " + firstName + ", "

        writeup = EmailName + txtWriteup.Text
        My.Computer.Clipboard.SetText(writeup)
        MsgBox("¯\_(ツ)_/¯   Successfully copied to clipboard.  (╯°□°）╯︵ ┻━┻")
        txtInfo.Enabled = True
        txtLog.Clear()


        clear()

        Return writeup

    End Function


    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Function: btnData_Click
    'Purpose: Calling this will determine if the french checkbox is enabled or not
    '         and will call the appropriate function 
    'Returns: 
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Private Sub btnData_Click(sender As Object, e As EventArgs) Handles btnData.Click

        Try

            If CBFrench.Checked = True Then
                French()
            Else
                English()
            End If



        Catch ex As System.NullReferenceException
            MsgBox("Error. No Data to copy to Clipboard")
        End Try
    End Sub


    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Function: RBC2G_CheckedChanged
    'Purpose: Calling this will disable all CIMS buttons, and resizes the form to show the C2G
    '         functionality
    'Returns: 
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Private Sub RBC2G_CheckedChanged(sender As Object, e As EventArgs) Handles RBC2G.CheckedChanged
        Me.Width = 1200
        txtInfo.Enabled = False
        btnData.Enabled = False
        txtC2GInfo.Enabled = True
        btnC2GWriteup.Enabled = True
    End Sub


    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Function: RBCIMS_CheckedChanged
    'Purpose: Calling this will disable all C2G buttons, and resize the form to show all CIMS
    '         functionality
    'Returns: 
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Private Sub RBCIMS_CheckedChanged(sender As Object, e As EventArgs) Handles RBCIMS.CheckedChanged
        Me.Width = 750
        txtInfo.Enabled = True
        btnData.Enabled = True

        txtC2GInfo.Enabled = False
        btnC2GWriteup.Enabled = False
    End Sub



    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Function: CIMS_Load
    'Purpose: When the form is first loaded, it will close the login form, and setup the split
    'Returns: 
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Private Sub CIMS_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        getprocess()
        CheckForIllegalCrossThreadCalls = False
        ' LoadProcess()


        Login.Close()
        lblSplit.TextAlign = ContentAlignment.TopCenter
        lblSplit.Text = String.Join(Environment.NewLine, "||||||||||||||||||||||||||||||||||||||".Select(Function(c) c.ToString()))

        RBCIMS.Checked = True
        txtAccess.ForeColor = Color.DarkRed

        btnAD.Enabled = False
        btnCIMS.Enabled = False
        btnNewHire.Enabled = False
        btnData.Enabled = False


    End Sub


    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Function: UnAccent
    'Purpose: Calling this will convert any french characters to the equivalent enlish char
    'Returns: aString
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Public Function UnAccent(ByVal aString As String) As String
        Dim toReplace() As Char = "àèìòùÀÈÌÒÙ äëïöüÄËÏÖÜ âêîôûÂÊÎÔÛ áéíóúÁÉÍÓÚðÐýÝ ãñõÃÑÕšŠžŽçÇåÅøØ".ToCharArray
        Dim replaceChars() As Char = "aeiouAEIOU aeiouAEIOU aeiouAEIOU aeiouAEIOUdDyY anoANOsSzZcCaAoO".ToCharArray
        For index As Integer = 0 To toReplace.GetUpperBound(0)
            aString = aString.Replace(toReplace(index), replaceChars(index))
        Next
        Return aString
    End Function


    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Function: txtC2GInfo_Leave
    'Purpose: Up leaving the C2G textblock, it will parse through everything, and setup the form
    'Returns: 
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Private Sub txtC2GInfo_Leave(sender As Object, e As EventArgs) Handles txtC2GInfo.Leave
        Try
            Dim strC2GText() As String
            strC2GText = Split(txtC2GInfo.Text, vbCrLf)



            Dim C2GName As String = strC2GText(0)
            Dim C2GEmail As String = strC2GText(1)
            Dim C2GPhone As String = strC2GText(2)
            Dim C2GRiding As String = strC2GText(4)

            If strC2GText(0).StartsWith("Name: ") Then

                Dim ReplaceC2G_Name As String = Replace(C2GName, "Name: ", "")
                Dim ReplaceC2G_Email As String = Replace(C2GEmail, "E-mail Address: ", "")
                Dim ReplaceC2G_Phone As String = Replace(C2GPhone, "Phone Number: ", "")
                Dim ReplaceC2G_Riding As String = Replace(C2GRiding, "Organization Name / ED: ", "")

                ReplaceC2G_Name = LTrim(ReplaceC2G_Name)
                ReplaceC2G_Name = RTrim(ReplaceC2G_Name)

                ReplaceC2G_Email = LTrim(ReplaceC2G_Email)
                ReplaceC2G_Email = RTrim(ReplaceC2G_Email)

                ReplaceC2G_Phone = LTrim(ReplaceC2G_Phone)
                ReplaceC2G_Phone = RTrim(ReplaceC2G_Phone)

                ReplaceC2G_Riding = LTrim(ReplaceC2G_Riding)
                ReplaceC2G_Riding = RTrim(ReplaceC2G_Riding)

                txtC2GName.Text = ReplaceC2G_Name
                txtC2GEmail.Text = ReplaceC2G_Email
                txtC2GNumber.Text = ReplaceC2G_Phone
                txtC2GRiding.Text = ReplaceC2G_Riding


                C2GmyName = ReplaceC2G_Name
                C2GfirstName = C2GmyName.Substring(0, C2GmyName.IndexOf(" "))
                C2GLastName = C2GmyName.Substring(C2GmyName.IndexOf(" ") + 1)

                Dim C2GmyUserName As String

                '    Dim lower As String

                C2GmyUserName = txtC2GName.Text

                txtC2GFirstName.Text = UnAccent(C2GfirstName)
                txtC2GLastName.Text = UnAccent(C2GLastName)
                ''   C2GmyString = Replace(Replace(C2GmyUserName, " ", ""), "-", "")

                '   lower = LCase(C2GmyString)
                '   C2GmyString = lower
                '  txtC2GTruncated.Text = C2GmyString
                'My.Computer.Clipboard.SetText(myString)

                Dim C2Gfinal As String
                Dim C2GFirstLetter As String
                Dim C2GMiddleLetter As String


                Dim C2GLastNumber As Integer
                C2GFirstLetter = GenerateRandomStringUpper(1, True)
                C2GMiddleLetter = GenerateRandomString(6, False)

                C2GLastNumber = CInt(Int((9 * Rnd()) + 1))
                C2Gfinal = C2GFirstLetter + C2GMiddleLetter & C2GLastNumber
                txtC2GPassword.Text = C2GFirstLetter + C2GMiddleLetter & C2GLastNumber


                txtC2GFinal.Text = ("Username: " + ReplaceC2G_Email & Environment.NewLine & "Password: " + C2Gfinal & Environment.NewLine & Environment.NewLine + "C2G Username + Password")
                txtC2GInfo.Enabled = False

            Else
                MsgBox("Error. Please Enter Info From Zendesk.")
            End If

        Catch ex As System.IndexOutOfRangeException
            MsgBox("You Must enter in the data from Zendesk")
        End Try





    End Sub

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Function: btnC2GWriteup_Click
    'Purpose: When pressed, the first name is passed through, and will add it to the string
    '         of the email writeup for C2G
    'Returns:
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Private Sub btnC2GWriteup_Click(sender As Object, e As EventArgs) Handles btnC2GWriteup.Click
        Try

            Dim C2Gwriteup As String

            Dim C2GEmailName As String = "Hello " + C2GfirstName + ", "

            C2Gwriteup = C2GEmailName + txtC2GWriteup.Text
            My.Computer.Clipboard.SetText(C2Gwriteup)
            MsgBox("Successfully copied to clipboard.  (╯°□°）╯︵ ┻━┻", MsgBoxStyle.OkOnly, "COPIED")
            txtC2GInfo.Enabled = True

        Catch ex As System.NullReferenceException
            MsgBox("Error. No Data to copy to Clipboard")
        End Try
    End Sub

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Function: txtCIMSFirst_Enter
    'Purpose: When this textbox is entered, or active, it will copy the contents to the clipboard
    '         As well as does error checking for NULL in the clipboard
    'Returns: 
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Private Sub txtCIMSFirst_Enter(sender As Object, e As EventArgs) Handles txtCIMSFirst.Enter
        Try

            My.Computer.Clipboard.SetText(CIMSFirstName)

        Catch ex As System.ArgumentNullException
            MsgBox("Error. No Data to copy to Clipboard")
        End Try
    End Sub


    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Function: txtCIMSLast_Enter
    'Purpose: When this textbox is entered, or active, it will copy the contents to the clipboard
    '         As well as does error checking for NULL in the clipboard
    'Returns: 
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Private Sub txtCIMSLast_Enter(sender As Object, e As EventArgs) Handles txtCIMSLast.Enter
        Try

            My.Computer.Clipboard.SetText(CIMSLastName)

        Catch ex As System.ArgumentNullException
            MsgBox("Error. No Data to copy to Clipboard")
        End Try
    End Sub


    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Function: txtName_Enter
    'Purpose: When this textbox is entered, or active, it will copy the contents to the clipboard
    '         As well as does error checking for NULL in the clipboard
    'Returns: 
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Private Sub txtName_Enter(sender As Object, e As EventArgs) Handles txtName.Enter
        Try

            My.Computer.Clipboard.SetText(myName)

        Catch ex As System.ArgumentNullException
            MsgBox("Error. No Data to copy to Clipboard")
        End Try
    End Sub


    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Function: txtEmail_Enter
    'Purpose: When this textbox is entered, or active, it will copy the contents to the clipboard
    '         As well as does error checking for NULL in the clipboard
    'Returns: 
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Private Sub txtEmail_Enter(sender As Object, e As EventArgs) Handles txtEmail.Enter
        Try

            My.Computer.Clipboard.SetText(replaceEmail)

        Catch ex As System.ArgumentNullException
            MsgBox("Error. No Data to copy to Clipboard")
        End Try
    End Sub


    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Function: txtPhone_Enter
    'Purpose: When this textbox is entered, or active, it will copy the contents to the clipboard
    '         As well as does error checking for NULL in the clipboard
    'Returns: 
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Private Sub txtPhone_Enter(sender As Object, e As EventArgs) Handles txtPhone.Enter
        Try

            My.Computer.Clipboard.SetText(replacePhone)

        Catch ex As System.ArgumentNullException
            MsgBox("Error. No Data to copy to Clipboard")
        End Try
    End Sub

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Function: txtTruncated_Enter
    'Purpose: When this textbox is entered, or active, it will copy the contents to the clipboard
    '         As well as does error checking for NULL in the clipboard
    'Returns: 
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Private Sub txtTruncated_Enter(sender As Object, e As EventArgs) Handles txtTruncated.Enter
        Try

            My.Computer.Clipboard.SetText(UnAccent(myString))

        Catch ex As System.ArgumentNullException
            MsgBox("Error. No Data to copy to Clipboard")

        Catch ex As System.NullReferenceException
            MsgBox("Error. No Data to copy to Clipboard")
        End Try
    End Sub

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Function: txtPassword_Enter
    'Purpose: When this textbox is entered, or active, it will copy the contents to the clipboard
    'Returns: 
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Private Sub txtPassword_Enter(sender As Object, e As EventArgs) Handles txtPassword.Enter
        Try

            My.Computer.Clipboard.SetText(UnAccent(myString))

        Catch ex As System.ArgumentNullException
            MsgBox("Error. No Data to copy to Clipboard")
        End Try
    End Sub


    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Function: Button1_Click
    'Purpose: The heart of Active Directory integration. This function does it all
    'Returns: 
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles btnAD.Click
        txtLog.Clear()

        AD_FullName = myName
        AD_GivenName = CIMSFirstName
        AD_Surname = CIMSLastName

        ridingnum = replacesomething.Substring(0, 5)

        If txtWorking.Text = "MP" Then
            MPPath = "Get-ADOrganizationalUnit -LDAPFilter '(name=" + ridingnum + " *)' -SearchBase 'OU=Terminal Server Users,OU=CIMS,DC=conservative,DC=ca' -SearchScope Subtree -Properties DistinguishedName"
            ExecuteMyPowerShellScriptLocation(MPPath)
        Else
            EDAPath = "Get-ADOrganizationalUnit -LDAPFilter '(name=" + ridingnum + " *)' -SearchBase 'OU=VPN Users,OU=CIMS,DC=conservative,DC=ca' -SearchScope Subtree -Properties DistinguishedName"
            ExecuteMyPowerShellScriptLocation(EDAPath)
        End If
        Try


            AD_UPNSuffix = "$Suffix = """

            AD_UPNSuffix += myString + """"
            '  securepass = "$SecurePassword=ConvertTo-SecureString '" + finalPassword + "' –asplaintext –force "

            'ExecuteMyPowerShellScriptLocation(EDAPath)
            If CreateAccCounter >= 1 Then
                AD_PATH = "New-ADUser -Name" + " """ + AD_FullName + " "" " + "-GivenName " + AD_GivenName + " -Description '" + replaceDescRiding + "' -OfficePhone '" + txtPhone.Text + "' -Surname " + AD_Surname + " -DisplayName '" + txtName.Text + "' -Server conservative.ca" + " -CannotChangePassword $true -enable $True" + " -EmailAddress " + replaceEmail + " -SamAccountName " + myString + " -UserPrincipalName  ""$($Suffix)@conservative.ca""" + " -AccountPassword (ConvertTo-SecureString -AsPlainText " + finalPassword + " -Force)  -Path '" + displayname + "' -PassThru | Enable-ADAccount "

            Else
                AD_PATH = "New-ADUser -Name" + " """ + AD_FullName + """ " + "-GivenName " + AD_GivenName + " -Description '" + replaceDescRiding + "' -OfficePhone '" + txtPhone.Text + "' -Surname " + AD_Surname + " -DisplayName '" + txtName.Text + "' -Server conservative.ca" + " -CannotChangePassword $true -enable $True" + " -EmailAddress " + replaceEmail + " -SamAccountName " + myString + " -UserPrincipalName  ""$($Suffix)@conservative.ca""" + " -AccountPassword (ConvertTo-SecureString -AsPlainText " + finalPassword + " -Force) -Path '" + displayname + "' -PassThru | Enable-ADAccount "
                CreateAccCounter = 0
            End If
            'securepass & vbCrLf & vbCrLf &

            If txtWorking.Text = "MP" Then
                txtCode.Text = AD_UPNSuffix & vbCrLf & vbCrLf & AD_PATH & vbCrLf & vbCrLf & AD_REMOTE + myString
            Else
                txtCode.Text = AD_UPNSuffix & vbCrLf & vbCrLf & AD_PATH & vbCrLf & vbCrLf & AD_VPN + myString
            End If


            ' My.Computer.Clipboard.SetText(txtCode.Text)
            '  MsgBox("before")
            ExecuteMyPowerShellScript(txtCode.Text)
            txtLog.AppendText(Date.Now + " " + myName + " Object Created in Active Directory")
        Catch ex As CmdletInvocationException
            Powershell_Thread_exist = New System.Threading.Thread(AddressOf Exist)
            Powershell_Thread_exist.Start()
            '   Exist()
        End Try




    End Sub

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Function: Exist
    'Purpose: This function is called if the user exists in AD already
    '        :What to do if that happens
    'Returns: 0
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Private Function Exist() As String
        txtLog.AppendText(vbCrLf & Date.Now + " " + myName + " Object Already Exists in Active Directory")
        ADLocation = "get-aduser " + myString + " -properties DistinguishedName"
        ExecuteMyPowerShellScriptLocation(ADLocation)
        txtLog.AppendText(vbCrLf & vbCrLf & displayname)
        ADEnabled = "get-aduser " + myString + " -properties Enabled"
        ExecuteMyPowerShellEnabled(ADEnabled)
        txtLog.AppendText(vbCrLf & vbCrLf & "Account Status Enabled:" + displayname)
        If displayname.Equals("False") Then
            Counter += 1
            strInt = Counter.ToString
            If MsgBox("What Would you like to do?" & vbCrLf & vbCrLf & "Yes: Add " + strInt + " to the username + ReMake the Account" & vbCrLf & "No: Re-enabled the account and Modify", MsgBoxStyle.YesNoCancel, "Modify User Information") = MsgBoxResult.Yes Then
                myString += strInt
                txtTruncated.Text = myString

                txtResult.Text = ("Username: " + UnAccent(myString) & Environment.NewLine & "Password: " + final & Environment.NewLine & Environment.NewLine + "AD & CIMS")
                txtLog.Clear()
                CreateAccCounter += 1
                Button1_Click(sender, New System.EventArgs())
                ' Button1_Click()
            ElseIf MsgBoxResult.No Then
                SETADUSER()


            ElseIf MsgBoxResult.Cancel Then

            End If
        End If
        Return 0

    End Function


    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Function: Exist
    'Purpose: This function is called if the user exists in AD already
    '        :What to do if that happens
    'Returns: 0
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Private Function SETADUSER() As String
        txtLog.Clear()
        ADEnableAccount += myString
        ExecuteMyPowerShellScript(ADEnableAccount)
        txtLog.AppendText(Date.Now + " " + myName + "'s Account Has Been Enabled" & vbCrLf & vbCrLf)

        If txtWorking.Text = "MP" Then
            ADSetPath = "Get-ADOrganizationalUnit -LDAPFilter '(name=" + ridingnum + " *)' -SearchBase 'OU=Terminal Server Users,OU=CIMS,DC=conservative,DC=ca' -SearchScope Subtree -Properties DistinguishedName"
        Else

            ADSetPath = "Get-ADOrganizationalUnit -LDAPFilter '(name=" + ridingnum + " *)' -SearchBase 'OU=VPN Users,OU=CIMS,DC=conservative,DC=ca' -SearchScope Subtree -Properties DistinguishedName"
        End If
        Try

            AD_UPNSuffix = "$Suffix = """

            ExecuteMyPowerShellScriptLocation(ADSetPath)

            AD_UPNSuffix += myString + """"
            ExecuteMyPowerShellScript(AD_UPNSuffix)




            ADMoveUser = "Get-ADuser " + myString + " | Move-ADObject -TargetPath '" + displayname + "'"
            ExecuteMyPowerShellScript(ADMoveUser)
            txtLog.AppendText(Date.Now + " " + myName + " - Object Moved to the New OU" & vbCrLf & vbCrLf)


            AD_PATH = "Set-ADUser -identity '" + myString + "' -GivenName " + AD_GivenName + " -Description '" + replaceDescRiding + "' -OfficePhone '" + txtPhone.Text + "' -Server conservative.ca" + " -CannotChangePassword $true -enable $True" + " -EmailAddress " + replaceEmail + " -Surname " + AD_Surname + " -DisplayName '" + txtName.Text + "' -SamAccountName '" + myString + "' | Enable-ADAccount "
            ExecuteMyPowerShellScript(AD_PATH)
            txtLog.AppendText(Date.Now + " " + myName + "'s Account information Modified" & vbCrLf & vbCrLf)



            ADResetPassword = "Set-ADAccountPassword " + myString + " -NewPassword (ConvertTo-SecureString -AsPlainText " + finalPassword + " -Force)  -Reset"
            ExecuteMyPowerShellScript(ADResetPassword)
            txtLog.AppendText(Date.Now + " " + myName + " - Password Has Been Changed" & vbCrLf & vbCrLf)

            '    txtCode.Text = ADResetPassword & vbCrLf & vbCrLf & AD_UPNSuffix & vbCrLf & vbCrLf & AD_PATH & vbCrLf & vbCrLf & ADMoveUser & vbCrLf & vbCrLf & AD_VPN + myString
            '      MsgBox("After")
            My.Computer.Clipboard.SetText(txtCode.Text)
            '     ExecuteMyPowerShellScript(txtCode.Text)

            '     txtLog.AppendText(Date.Now + " " + myName + "'s Account information Modified" & vbCrLf & vbCrLf)
            '     txtLog.AppendText(Date.Now + " " + myName + "Object Moved to the New OU")
        Catch ex As System.Management.Automation.CmdletInvocationException
            '    txtLog.AppendText(Date.Now + " " + myName + "'s Account information Modified - VPN" & vbCrLf & vbCrLf)
            '   txtLog.AppendText(Date.Now + " " + myName + "Object Moved to the New OU - VPN")
            Exit Try
        Catch ex As System.Management.Automation.ParseException
            My.Computer.Clipboard.SetText(txtCode.Text)
            txtLog.AppendText(vbCrLf & Date.Now + " " + myName + " Powershell Script Error. Please Manually Create")

            Exit Try

        End Try
        Return 0

    End Function

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Function: Button1_Click_1
    'Purpose: When pressed, launches a messagebox saying that this functionality isn't 
    '         available yet.
    'Returns: 
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Private Sub btnCIMS_click(sender As Object, e As EventArgs) Handles btnCIMS.Click
        '  MsgBox("¯\_(ツ)_/¯" & vbCrLf & vbCrLf & "Function Not available yet", MsgBoxStyle.Exclamation, "Nope Nope Nope")
        If finalRidingNumber.StartsWith("2") Then
            PreferredLanguageID = 2
        End If

        If CBActive.CheckState = CheckState.Unchecked Then
            Active = 0
        End If
        Notes = txtNotes.Text
        CreatedDate = Date.Now
        AddNewUser()

    End Sub


    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Function: txtTruncated_Leave
    'Purpose: Upon leaving the txtbox where the truncated username is kept, it catches any
    '         errors and displays a message box
    'Returns: 
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Private Sub txtTruncated_Leave(sender As Object, e As EventArgs) Handles txtTruncated.Leave
        Try


        Catch ex As System.ArgumentNullException
            MsgBox("Error. No Data to copy to Clipboard")

        Catch ex As System.NullReferenceException
            MsgBox("Error. No Data to copy to Clipboard")
        End Try
    End Sub


    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Function: btnNewHire_Click
    'Purpose: When Pressed, the program will launch the location on sharepoint where all
    '         the newhire documents are located.
    'Returns: 
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Private Sub btnNewHire_Click(sender As Object, e As EventArgs) Handles btnNewHire.Click
        System.Diagnostics.Process.Start("https://cpc-pcc.esp-access.com/staff/resources/Lists/New%20Hires/AllItems.aspx")
    End Sub

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Function: CloseToolStripMenuItem_Click
    'Purpose: When Pressed, this function will close the application along with the thread
    'Returns: 
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Private Sub CloseToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CloseToolStripMenuItem.Click
        Me.Close()

    End Sub


    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Function: AboutToolStripMenuItem_Click
    'Purpose: When Pressed, this function will show the About form
    'Returns: 
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Private Sub AboutToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AboutToolStripMenuItem.Click
        About.Show()
    End Sub


    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Function: DarkToolStripMenuItem_Click_1
    'Purpose: When Pressed, this function will changed all objects on screen to reflect
    '         a dark UI
    'Returns: 
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Private Sub DarkToolStripMenuItem_Click_1(sender As Object, e As EventArgs) Handles DarkToolStripMenuItem.Click
        Dim a As Control
        Dim b As Control
        Dim c As Control
        For Each a In Me.Controls
            If TypeOf a Is TextBox Then
                a.BackColor = Color.Black
                a.ForeColor = Color.White
            End If
        Next

        For Each b In Me.Controls
            If TypeOf b Is Label Then
                b.ForeColor = Color.Red
                b.BackColor = Color.Transparent

            End If
        Next


        For Each c In Me.Controls
            If TypeOf c Is Button Then
                c.BackColor = Color.Black
                c.ForeColor = Color.Red

            End If
        Next

        btnAD.FlatStyle = FlatStyle.Flat
        btnAD.FlatAppearance.BorderColor = Color.Red
        btnData.FlatStyle = FlatStyle.Flat
        btnCIMS.FlatStyle = FlatStyle.Flat
        btnNewHire.FlatStyle = FlatStyle.Flat
        btnC2GWriteup.FlatStyle = FlatStyle.Flat
        MenuStrip1.BackColor = Color.Black
        CBFrench.BackColor = Color.Black
        CBFrench.ForeColor = Color.Red

        RBC2G.BackColor = Color.Black
        RBC2G.ForeColor = Color.Red
        RBCIMS.BackColor = Color.Black
        RBCIMS.ForeColor = Color.Red
        MenuStrip1.BackColor = Color.Black
        HelpToolStripMenuItem.BackColor = Color.Black
        HelpToolStripMenuItem.ForeColor = Color.Red
        MenuStrip1.ForeColor = Color.Red
        Me.BackColor = Color.Black

    End Sub

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Function: LightToolStripMenuItem_Click
    'Purpose: When Pressed, this function will changed all objects on screen to reflect
    '         a light UI
    'Returns: 
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Private Sub LightToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LightToolStripMenuItem.Click
        Dim a As Control
        Dim b As Control
        Dim c As Control
        For Each a In Me.Controls
            If TypeOf a Is TextBox Then
                a.BackColor = SystemColors.Control
                a.ForeColor = SystemColors.ControlText
            End If
        Next

        For Each b In Me.Controls
            If TypeOf b Is Label Then
                b.BackColor = SystemColors.Control
                b.ForeColor = SystemColors.ControlText
            End If
        Next

        For Each c In Me.Controls
            If TypeOf c Is Button Then
                c.BackColor = SystemColors.Control
                c.ForeColor = SystemColors.ControlText



            End If
        Next
        btnAD.FlatStyle = FlatStyle.Standard
        btnData.FlatStyle = FlatStyle.Standard
        btnCIMS.FlatStyle = FlatStyle.Standard
        btnNewHire.FlatStyle = FlatStyle.Standard
        btnC2GWriteup.FlatStyle = FlatStyle.Standard
        Me.BackColor = SystemColors.Control
        RBC2G.BackColor = SystemColors.Control
        RBC2G.ForeColor = SystemColors.ControlText

        txtInfo.BackColor = SystemColors.Window
        txtInfo.ForeColor = SystemColors.ControlText

        RBCIMS.BackColor = SystemColors.Control
        RBCIMS.ForeColor = SystemColors.ControlText

        HelpToolStripMenuItem.BackColor = SystemColors.Control
        HelpToolStripMenuItem.ForeColor = SystemColors.ControlText

        CBFrench.BackColor = SystemColors.Control
        CBFrench.ForeColor = SystemColors.ControlText
        MenuStrip1.ForeColor = SystemColors.ControlText
        MenuStrip1.BackColor = SystemColors.Control
    End Sub

    Private Sub WipePass_Click(sender As Object, e As EventArgs) Handles WipePass.Click
        RemovePwd(ComboBox1.SelectedItem)

    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click

    End Sub

    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Button1.Click
        getprocess()
    End Sub

    Private Sub MenuStrip1_ItemClicked(sender As Object, e As ToolStripItemClickedEventArgs) Handles MenuStrip1.ItemClicked

    End Sub

    Private Sub txtInfo_TextChanged(sender As Object, e As EventArgs) Handles txtInfo.TextChanged

    End Sub
End Class
