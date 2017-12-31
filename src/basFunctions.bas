Option Compare Database
Option Explicit

Public Function GetAge(varDOB As Variant, Optional varAsOf As Variant) As String
    ' Source: Allen Browne, http://allenbrowne.com/func-08.html
    ' Purpose: Return the Age in years.
    ' Arguments: varDOB = Date Of Birth
    '            varAsOf = the date to calculate the age at, or today if missing.
    ' Return: Whole number of years.

    Dim dteDOB As Date
    Dim dteAsOf As Date
    Dim dteBDay As Date  ' Birthday in the year of calculation.

    ' Validate parameters
    If IsDate(varDOB) Then
        dteDOB = varDOB

        If Not IsDate(varAsOf) Then    ' Date to calculate age from.
            dteAsOf = Date
        Else
            dteAsOf = varAsOf
        End If

        If dteAsOf >= dteDOB Then      ' Calculate only if it's after person was born.
            dteBDay = DateSerial(Year(dteAsOf), Month(dteDOB), Day(dteDOB))
            GetAge = DateDiff("yyyy", dteDOB, dteAsOf) + (dteBDay > dteAsOf) & " years old"
        End If
    End If
End Function

Public Function IsValidEmail(strAddress As String) As Boolean
    ' Ref: https://www.ozgrid.com/forum/forum/help-forums/excel-general/108987-vba-function-to-confirm-email-address-is-valid
    Dim oRegEx As Object
    Set oRegEx = CreateObject("VBScript.RegExp")
    With oRegEx
        .Pattern = "^[\w-\.]{1,}\@([\da-zA-Z-]{1,}\.){1,}[\da-zA-Z-]{2,3}$"
        IsValidEmail = .Test(strAddress)
    End With
    Set oRegEx = Nothing
End Function

Public Sub RemoveDBOPrefix()
    ' Removes "dbo_" prefixs before imported SQL table names

    Dim obj As AccessObject
    Dim dbs As Object

    Set dbs = Application.CurrentData

    ' Search for open AccessObject objects in AllTables collection.
    For Each obj In dbs.AllTables
        'If found, remove prefix
       If Left(obj.Name, 4) = "dbo_" Then
            DoCmd.Rename Mid(obj.Name, 5), acTable, obj.Name
        End If
    Next obj
End Sub