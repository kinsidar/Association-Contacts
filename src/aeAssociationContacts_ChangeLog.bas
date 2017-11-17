Option Compare Database
Option Explicit

Public gstrClientId As String
Public gstrSvipCrisisId As String
Public gintCurrentUser As Integer
Public gblnHideFormHeader As Boolean
Public gblnDeveloper As Boolean
Public gintUserId As Integer

' Constants for settings of "SVIPDB"
Public Const gblnTEST As Boolean = True
Public Const gstrPROJECT_SVIPDB As String = "AssociationContacts"
Private Const mstrVERSION_SVIPDB As String = "0.0.2"
Private Const mstrDATE_SVIPDB As String = "October 18, 2017"

Public Const SVIPDB_SQL_FRONT_END = False
Public Const SVIPDB_AZSQL_FRONT_END = False
Public Const SVIPDB_STAFF_PERMISSIONS = False
'

Public Function getMyVersion() As String
    On Error GoTo 0
    getMyVersion = mstrVERSION_SVIPDB
End Function

Public Function getMyDate() As String
    On Error GoTo 0
    getMyDate = mstrDATE_SVIPDB
End Function

Public Function getMyProject() As String
    On Error GoTo 0
    getMyProject = gstrPROJECT_SVIPDB
End Function

Public Sub AC_EXPORT(Optional ByVal varDebug As Variant)

    Const THE_FRONT_END_APP = True
    Const THE_SOURCE_FOLDER = ".\src\"
    Const THE_XML_FOLDER = ".\src\xml\"
    Const THE_XML_DATA_FOLDER = ".\src\xmldata\"
    Const THE_BACK_END_SOURCE_FOLDER = "NONE"
    Const THE_BACK_END_XML_FOLDER = "NONE"
    Const THE_BACK_END_DB1 = "NONE"

    On Error GoTo PROC_ERR

    'Debug.Print "THE_BACK_END_DB1 = " & THE_BACK_END_DB1
    If Not IsMissing(varDebug) Then
        aegitClassTest varDebug:="varDebug", varSrcFldr:=THE_SOURCE_FOLDER, varSrcFldrBe:=THE_BACK_END_SOURCE_FOLDER, _
                        varXmlFldr:=THE_XML_FOLDER, varXmlDataFldr:=THE_XML_DATA_FOLDER, _
                        varFrontEndApp:=THE_FRONT_END_APP, _
                        varBackEndDbOne:=THE_BACK_END_DB1
    Else
        aegitClassTest varSrcFldr:=THE_SOURCE_FOLDER, varSrcFldrBe:=THE_BACK_END_SOURCE_FOLDER, _
                        varXmlFldr:=THE_XML_FOLDER, varXmlDataFldr:=THE_XML_DATA_FOLDER, _
                        varFrontEndApp:=THE_FRONT_END_APP, _
                        varBackEndDbOne:=THE_BACK_END_DB1
    End If

PROC_EXIT:
    Exit Sub

PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure SVIPDB_EXPORT"
    Resume Next

End Sub
'
'
'=============================================================================================================================
' Tasks:
' %020 -
' %019 -
' %018 -
' %017 -
' %016 -
' %015 -
' %012 - Add Shift key blocking
' %011 - Add Splash form
' %003 - Relates to GH #9, include version tracking details in the app database change log module
' %002 - Test Helen Fedema add-in for renaming http://www.helenfeddema.com/files/Code10.zip
' %001 - Use ae standards for naming objects - Ref: https://en.wikipedia.org/wiki/Hungarian_notation,
'           https://en.wikipedia.org/wiki/Leszynski_naming_convention
'           RVBA: https://ss64.com/access/syntax-naming.html
'=============================================================================================================================
'
'20171114 - v003 -
    ' FIXED - %014 - Fix error when deleting email and phone number record
    ' FIXED - %010 - Add Users table
    ' FIXED - %013 - Many many bug fixes from refactoring
    ' FIXED - %009 - Add Custom Ribbon
    ' FIXED - %008 - Split database FE/BE
    ' FIXED - %007 - Update forms to use new queries to link to tables
    ' FIXED - %006 - Create queries for the forms to link with tables
'20171017 - v002 -
    ' FIXED - %005 - Fix internal version and date
    ' FIXED - %004 - Add compressed db to the zip folder
'20171009 - v001 - Initial database design based on a sample from:
    ' Ref: https://www.devhut.net/2016/09/01/ms-access-contact-database-template-sample/
    ' Daniel Pineault, Microsoft MVP, 2010-2017
    ' "I am truly pleased to announce that I have been awarded the Title of Microsoft Most Valuable Professional (MVP), years in a row,
    ' for my contributions to the MS Access community. It is a great honor to receive this award directly form the hands of Microsoft
    ' and it is my pleasure to help other developers when I can. This is one of the main reason why I created this website in the first place,
    ' to share knowledge, no strings attached."