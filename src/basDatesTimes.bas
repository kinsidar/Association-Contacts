Option Compare Database
Option Explicit

Public Function GetAge(varDOB As Variant, Optional varAsOf As Variant) As String
    'Source: Allen Browne, http://allenbrowne.com/func-08.html
    'Purpose: Return the Age in years.
    'Arguments: varDOB = Date Of Birth
    '           varAsOf = the date to calculate the age at, or today if missing.
    'Return: Whole number of years.

    Dim dteDOB As Date
    Dim dteAsOf As Date
    Dim dteBDay As Date  'Birthday in the year of calculation.

    'Validate parameters
    If IsDate(varDOB) Then
        dteDOB = varDOB

        If Not IsDate(varAsOf) Then  'Date to calculate age from.
            dteAsOf = Date
        Else
            dteAsOf = varAsOf
        End If

        If dteAsOf >= dteDOB Then      'Calculate only if it's after person was born.
            dteBDay = DateSerial(Year(dteAsOf), Month(dteDOB), Day(dteDOB))
            GetAge = DateDiff("yyyy", dteDOB, dteAsOf) + (dteBDay > dteAsOf) & " years old"
        End If
    End If

End Function