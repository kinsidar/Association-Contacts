Option Compare Database
Option Explicit

'—————————————————————————-
' Procedure: RefreshTableLinks
' Purpose: Refresh table links to back-ends in the same folder as front end.
' Note: Linked Tables can be in more than one back-end.
' Return: Returns a zero-length string if all tables are relinked.
' Return: Or returns a string listing tables not relinked and errors.
'—————————————————————————-

Public Function RefreshTableLinks() As String

    On Error GoTo ErrHandle
    
    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim strCon As String
    Dim strBackEnd As String
    Dim strMsg As String
    Dim intErrorCount As Integer
    
    Set db = CurrentDb
    
    'MsgBox "The current database is located at " & Application.CurrentProject.Path & "."
    
    'Loop through the TableDefs Collection.
    For Each tdf In db.TableDefs
        'Verify the table is a linked table.
        If Left$(tdf.Connect, 10) = ";DATABASE=" Then
            'Get the existing Connection String.
            strCon = Nz(tdf.Connect, "")
            'Get the name of the back-end database using String Functions.
            strBackEnd = Right$(strCon, (Len(strCon) - (InStrRev(strCon, "\") - 1)))
            'Verify we have a value for the back-end
        If Len(strBackEnd & "") > 0 Then
            'Set a reference to the TableDef Object.
            Set tdf = db.TableDefs(tdf.Name)
            'Build the new Connection Property Value.
            tdf.Connect = ";DATABASE=" & CurrentProject.Path & strBackEnd
            'Refresh the table link.
            tdf.RefreshLink
        Else
            'There was a problem getting the name of the back-end.
            'Add the information to the message to notify the user.
            intErrorCount = intErrorCount + 1
            strMsg = strMsg & "Error getting back-end database name." & vbNewLine
            strMsg = strMsg & "Table Name: " & tdf.Name & vbNewLine
            strMsg = strMsg & "Connect = " & strCon & vbNewLine
        End If
        End If
    Next tdf
    
ExitHere:
    On Error Resume Next
    If intErrorCount > 0 Then
        strMsg = "There were errors refreshing the table links: " _
        & vbNewLine & strMsg & "In Procedure RefreshTableLinks"
        RefreshTableLinks = strMsg
    End If
    Set tdf = Nothing
    Set db = Nothing
    Exit Function
    
ErrHandle:
    intErrorCount = intErrorCount + 1
    strMsg = strMsg & "Error " & Err.Number & " " & Err.Description
    strMsg = strMsg & vbNewLine & "Table Name: " & tdf.Name & vbNewLine
    strMsg = strMsg & "Connect = " & strCon & vbNewLine
    Resume ExitHere

End Function