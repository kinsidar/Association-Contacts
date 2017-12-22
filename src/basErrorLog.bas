Option Compare Database
Option Explicit

Public Sub ErrorTest()
    On Error GoTo Err_ErrorTest          ' Initialize error handling.
    ' Code to do something here.
    'Dim i As Integer
    'i = 1 / 0
Exit_ErrorTest:                           ' Label to resume after error.
     Exit Sub                 ' Exit before error handler.
Err_ErrorTest:                           ' Label to jump to on error.
    Select Case Err.Number
    Case 9999                        ' Whatever number you anticipate.
        Resume Next                  ' Use this to just ignore the line.
    Case 999
        Resume Exit_ErrorTest         ' Use this to give up on the proc.
    Case Else                        ' Any unexpected error.
        Call LogError(Err.Number, Err.Description, "ErrorTest()")
        Resume Exit_ErrorTest
    End Select             ' Pick up again and quit.
End Sub

Public Sub ErrorCall()
    Call LogError(Err.Number, Err.Description, "ErrorTest()")
End Sub
 
Function LogError(ByVal lngErrNumber As Long, ByVal strErrDescription As String, _
    ByRef strCallingProc As String, Optional ByRef vParameters As String, Optional ByRef bShowUser As Boolean = True) As Boolean

    On Error GoTo Err_LogError

    ' Purpose: Generic error handler.
    ' Logs errors to table "tblLogError".
    ' Arguments: lngErrNumber - value of Err.Number
    ' strErrDescription - value of Err.Description
    ' strCallingProc - name of sub|function that generated the error.
    ' vParameters - optional string: List of parameters to record.
    ' bShowUser - optional boolean: If False, suppresses display.
    ' Author: Allen Browne, allen@allenbrowne.com
    
    FileContentClear 'Clears the output file before writing to it
    
    Dim strMsg As String      ' String for display in MsgBox
    Dim rst As DAO.Recordset  ' The tblLogError table

    Select Case lngErrNumber
    Case 0
        Debug.Print strCallingProc & " called error 0."
    Case 2501                ' Cancelled
        'Do nothing.
    Case 3314, 2101, 2115    ' Can't save.
        If bShowUser Then
            strMsg = "Record cannot be saved at this time." & vbCrLf & _
                "Complete the entry, or press <Esc> to undo."
            MsgBox strMsg, vbExclamation, strCallingProc
            DebugOutput (strMsg) 'Writes the message to a text file
        End If
    Case Else
        If bShowUser Then
            strMsg = "Error " & lngErrNumber & ": " & strErrDescription
            MsgBox strMsg, vbExclamation, strCallingProc
            DebugOutput (strMsg) 'Writes the message to a text file
        End If
        Set rst = CurrentDb.OpenRecordset("tblLogError", , dbAppendOnly) '
        rst.AddNew '
            rst![ErrNumber] = lngErrNumber '===============================================
            rst![ErrDescription] = Left$(strErrDescription, 255) 'sends error info to table
            rst![ErrDate] = Now() '
            rst![CallingProc] = strCallingProc '
            rst![userName] = CurrentUser() '
            rst![ShowUser] = bShowUser '
            If Not IsMissing(vParameters) Then '
                rst![Parameters] = Left(vParameters, 255) '=================================
            End If '
        rst.Update '
        rst.Close '
         
        
        LogError = True
    End Select

Exit_LogError:
    Set rst = Nothing 'end of table info
    DebugOutput (strMsg) 'Writes the message to a text file
    'current version only writes one line to the text file
    'the line is overwritten with consecutive use
    
    Exit Function
    
Err_LogError:
    strMsg = "Please contact your system Administrator." & vbCrLf & _
        "Please write down the following details:" & vbCrLf & vbCrLf & _
        "Calling Proc: " & strCallingProc & vbCrLf & _
        "Error Number " & lngErrNumber & vbCrLf & strErrDescription & vbCrLf & vbCrLf & _
        "Unable to record because Error " & Err.Number & vbCrLf & Err.Description
    MsgBox strMsg, vbCritical, "LogError()"
    DebugOutput (strMsg) 'Writes the message to a text file
    Resume Exit_LogError

End Function