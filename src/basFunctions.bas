Option Compare Database
Option Explicit

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