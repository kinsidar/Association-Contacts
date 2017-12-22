Option Compare Database
Option Explicit

Private Declare PtrSafe Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal _
   lpBuffer As String, nSize As Long) As Long

Private Const gblnlockProperties As Boolean = False 'For developer settings

Public Function ap_DisableShift()
    ' This function disables the shift at startup. This action causes
    ' the Autoexec macro and Startup properties to always be executed.

    On Error GoTo PROC_ERR

    Dim dbs As DAO.Database
    Dim prop As DAO.Property
    Const conPropNotFound = 3270

    Set dbs = CurrentDb()

    ' This next line disables the shift key on startup.
    dbs.Properties("AllowByPassKey") = False

    If gblnDeveloper Then
        dbs.Properties("AllowByPassKey") = True
        CurrentDb.Properties("StartUpShowDBWindow") = True
        CurrentDb.Properties("AllowSpecialKeys") = True
        CurrentDb.Properties("AllowFullMenus") = True
        DoCmd.ShowToolbar "Ribbon", acToolbarYes
    Else
        dbs.Properties("AllowByPassKey") = False
        ' Ref: http://www.access-programmers.co.uk/forums/showthread.php?t=187697
        ' Disable Navigation Pane
        CurrentDb.Properties("StartUpShowDBWindow") = False
        ' Disable Special Access Keys (F11)
        CurrentDb.Properties("AllowSpecialKeys") = False
        ' Dsiable Full Menus
        CurrentDb.Properties("AllowFullMenus") = False

        '''CurrentDb.Properties("AllowBuiltinToolbars") = False
        '''DoCmd.ShowToolbar "Ribbon", acToolbarNo
    End If

    ' The function is successful.
    Exit Function

PROC_EXIT:
    Exit Function

PROC_ERR:
    ' The first part of this error routine creates the "AllowByPassKey
    ' property if it does not exist.
    If Err = conPropNotFound Then
        Set prop = dbs.CreateProperty("AllowByPassKey", dbBoolean, False)
        dbs.Properties.Append prop
        Resume Next
    Else
        MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ")" & vbCrLf & _
            "Function 'ap_DisableShift' did not complete successfully." & vbCrLf & _
            "Please contact your system administrator.", vbCritical, gstrPROJECT_ACDB_BE
        ErrorCall
        Resume PROC_EXIT
    End If
End Function

Public Function IsDeveloper() As Boolean
    ' Allows only developers to use the shift key during startup

    Dim userName As String
    IsDeveloper = False
    userName = VBA.Environ("Username")

    Select Case userName
        Case "peter"
            IsDeveloper = True
        Case "petere"
            IsDeveloper = True
        Case "peterennis"
            IsDeveloper = True
        Case "kdurst"
            IsDeveloper = True
        Case "kinsa"
            IsDeveloper = True
        Case Else
            IsDeveloper = False
    End Select

    ' NOTE: Additional user test to enter password for developer mode
    Dim strPass As String
    Dim strThePassword As String
    strThePassword = "TestMe"
    If userName = "TheOtherTester" Then
        strPass = InputBox("Enter Password")
        If strPass <> strThePassword Then
            IsDeveloper = False
        Else
            IsDeveloper = True
        End If
    End If

End Function

Public Function IsValidUser() As Boolean
    ' Allows only certain users to edit forms

    Dim userName As String
    IsValidUser = False
    userName = GetUser()
    'Debug.Print UserName

    Select Case userName
        Case "peter"
            IsValidUser = True
        Case "petere"
            IsValidUser = True
        Case "peterennis"
            IsValidUser = True
        Case "kdurst"
            IsValidUser = True
        Case "kinsa"
            IsValidUser = True
        Case "administrator"
            IsValidUser = True
        Case Else
            IsValidUser = False
    End Select
End Function

Public Function IsAdmin() As Boolean
    ' Allows only the admin see the complete feedback datasheet

    Dim userName As String
    IsAdmin = False
    userName = GetUser()
    Debug.Print userName

    Select Case userName
        Case "peter"
            IsAdmin = True
        Case "petere"
            IsAdmin = True
        Case "peterennis"
            IsAdmin = True
        Case "kdurst"
            IsAdmin = True
        Case "kinsa"
            IsAdmin = True
        Case "administrator"
            IsAdmin = True
        Case Else
            IsAdmin = False
    End Select
End Function

Public Function GetUser() As String
    ' Gets the username of the user on the PC

    On Error GoTo PROC_ERR

    Dim strBuffer As String
    Dim lngSize As Long, lngRetVal As Long
   
    lngSize = 199
    strBuffer = String$(200, 0)
   
    lngRetVal = GetUserName(strBuffer, lngSize)
   
    GetUser = Left$(strBuffer, lngSize - 1)

PROC_EXIT:
    Exit Function

PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") Please contact your system administrator."
    ErrorCall
    Resume PROC_EXIT

End Function

Public Function ap_EnableShift()
    ' Ref: https://support.microsoft.com/en-us/kb/826765
    ' This function enables the SHIFT key at startup. This action causes
    ' the Autoexec macro and the Startup properties to be bypassed
    ' if the user holds down the SHIFT key when the user opens the database.

    On Error GoTo PROC_ERR

    Dim db As DAO.Database
    Dim prop As DAO.Property
    Const conPropNotFound = 3270

    Set db = CurrentDb()

    ' This next line of code disables the SHIFT key on startup.
    db.Properties("AllowByPassKey") = True

    ' Function successful
    Exit Function

PROC_ERR:
    ' The first part of this error routine creates the "AllowByPassKey property if it does not exist.
    If Err = conPropNotFound Then
        Set prop = db.CreateProperty("AllowByPassKey", dbBoolean, True)
        db.Properties.Append prop
        Resume Next
    Else
        MsgBox "Function 'ap_EnableShift' did not complete successfully."
        Exit Function
    End If

End Function

Public Function LoadRibbons()
    ' Ref: http://bytes.com/topic/access/answers/952672-how-add-ribbon-name-vba-startup-form

    Dim dbs As DAO.Database
    Dim rst As DAO.Recordset

    Set dbs = Application.CurrentDb
    Set rst = CurrentDb.OpenRecordset("tblRibbons")

    '[RibbonName] - {TEXT}
    '[RibbonXml] - {MEMO}

    Do While Not rst.EOF
        Application.LoadCustomUI rst("RibbonName").Value, rst("RibbonXML").Value
        rst.MoveNext
    Loop

    rst.Close
    Set rst = Nothing
    Set dbs = Nothing

End Function

'==============================================================================
'The following was intended for having a password switch between the user and
'developer settings - it does not work
'==============================================================================

'Private Function DeveloperProperties(gblnlockProperties As Boolean) As Boolean
''
''    Dim gblnlockProperties As Boolean
''    gblnlockProperties = False
'    If Not gblnlockProperties Then
'    'Turns off developer settings
''
''    Application.SetOption "DesignWithData", False 'removes layout view
''    CurrentDb.Properties("StartupForm") = "SwitchBoard"
''    CurrentDb.Properties("ShowDocumentTabs") = True 'allows document tabs
''    CurrentDb.Properties("StartupShowDBWindows") = False 'removes navigation pane
''    CurrentDb.Properties("AllowShortcutMenus") = False
''    CurrentDb.Properties("UseMDIMode") = 0 'turns on tabbed documents
''    CurrentDb.Properties("AllowFullMenus") = False
''    CurrentDb.Properties.Append CurrentDb.CreateProperty("CustomRibbonID", dbText, "RAY")
'        Debug.Print "It worked!"
'    End If
'
'
'End Function
'
'Sub Password()
'    Dim pass As String
'    Dim thepassword As String
'    thepassword = "RDA"
'    pass = InputBox("Enter Password")
'    If pass <> thepassword Then
'        DeveloperProperties (False)
'        'turns off developer setting because password was wrong
'        Debug.Print "Developer settings turned off!"
'        Exit Sub
'    End If
'     ' continue code here
'    DeveloperProperties (True)
'    'turns off developer setting because password was correct
'    Debug.Print "Developer settings turned on!"
'End Sub