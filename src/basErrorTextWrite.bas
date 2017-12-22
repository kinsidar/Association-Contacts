Option Compare Database
Option Explicit
  
Public Function textWrite(ByRef InputString As String) As Integer
    'Use DebugOutput instead of this function
    On Error GoTo PROC_ERR
    
    Dim objFSO As Object
    Dim objFile As Object
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objFile = objFSO.CreateTextFile("C:\PETER\ITILRDA\DebugOutput.txt") 'set file path here
    objFile.WriteLine InputString 'Writes the input message to a text file in the source folder
    objFile.Close
    Set objFSO = Nothing
    Set objFile = Nothing
PROC_EXIT:
    Exit Function

PROC_ERR:
    
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") Please contact your system administrator."
    ErrorCall
    Resume PROC_EXIT
End Function

Public Function DebugOutput(ByVal Content As String) As Boolean 'changed to byVal 7/14

On Error GoTo PROC_ERR
    If gblnDeveloper Then
        Open "C:\PETER\ITILRDA\DebugOutput.txt" For Append As #1 'set file path here
        Print #1, Content
        Close #1
    Else
        'Does not write to an error log if the user is not a developer
    End If
    'call as: DebugOutput ("yourcontent" & variableHere)
PROC_EXIT:
    Exit Function

PROC_ERR:
    
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") Please contact your system administrator."
    ErrorCall
    Resume PROC_EXIT
End Function

Public Sub FileContentClear(Optional ByRef reblnverbose As Boolean)
    '==============
    'unelegant way of clearing the output file  before writing to it
    On Error GoTo PROC_ERR

    Dim objFSO As Object
    Dim objFile As Object
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objFile = objFSO.CreateTextFile("C:\PETER\ITILRDA\DebugOutput.txt") 'set file path here
    objFile.Write ""
    objFile.Close
    '==============
    'kill file if it exists and flag is off
    If Not reblnverbose Then
        'Kill ("C:\ae\reAccessClass\DebugOutput.txt")
        Dim RemoveFile As String
        RemoveFile = "C:\PETER\ITILRDA\DebugOutput.txt" 'set file path here
        'Check that file exists
        If Len(Dir$(RemoveFile)) > 0 Then
            'First remove readonly attribute, if set
            SetAttr RemoveFile, vbNormal
            'Then delete the file
            Kill RemoveFile
        End If
    End If
PROC_EXIT:
    Exit Sub

PROC_ERR:
    
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") Please contact your system administrator."
    ErrorCall
    Resume PROC_EXIT
End Sub