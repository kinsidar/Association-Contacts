Option Compare Database
Option Explicit


Function GetAge(varDOB As Variant, Optional varAsOf As Variant) As String
    'Source: Allen Browne
    '        http://allenbrowne.com/func-08.html
    'Purpose:   Return the Age in years.
    'Arguments: varDOB = Date Of Birth
    '           varAsOf = the date to calculate the age at, or today if missing.
    'Return:    Whole number of years.
    Dim dtDOB As Date
    Dim dtAsOf As Date
    Dim dtBDay As Date  'Birthday in the year of calculation.

    'Validate parameters
    If IsDate(varDOB) Then
        dtDOB = varDOB

        If Not IsDate(varAsOf) Then  'Date to calculate age from.
            dtAsOf = Date
        Else
            dtAsOf = varAsOf
        End If

        If dtAsOf >= dtDOB Then      'Calculate only if it's after person was born.
            dtBDay = DateSerial(Year(dtAsOf), Month(dtDOB), Day(dtDOB))
            GetAge = DateDiff("yyyy", dtDOB, dtAsOf) + (dtBDay > dtAsOf) & " years old"
        End If
    End If
End Function