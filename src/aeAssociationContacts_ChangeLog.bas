Option Compare Database
Option Explicit

Public gblnHideFormHeader As Boolean
Public gblnDeveloper As Boolean
Public gintUserId As Integer

' Constants for settings of "ACDB"
Public Const gblnTEST As Boolean = True
Public Const gstrPROJECT_ACDB As String = "AssociationContacts"
Private Const mstrVERSION_ACDB As String = "0.1.5"
Private Const mstrDATE_ACDB As String = "December 31, 2017"

Public Const ACDB_SQL_FRONT_END = False
Public Const ACDB_AZSQL_FRONT_END = False
Public Const ACDB_STAFF_PERMISSIONS = False
Public Const ACDB_SHOW_LOGIN_FORM = False
'

Public Function getMyVersion() As String
    On Error GoTo 0
    getMyVersion = mstrVERSION_ACDB
End Function

Public Function getMyDate() As String
    On Error GoTo 0
    getMyDate = mstrDATE_ACDB
End Function

Public Function getMyProject() As String
    On Error GoTo 0
    getMyProject = gstrPROJECT_ACDB
End Function

Public Sub ACDB_EXPORT(Optional ByVal varDebug As Variant)

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
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure ACDB_EXPORT"
    Resume Next

End Sub
'
'
'=============================================================================================================================
' Tasks:
' %050 -
' %049 -
' %048 -
' %047 -
' %046 -
' %031 - Create backend on network SQL Server with renamed front end
' %030 - Create backend on local SQL Server with renamed front end
' %029 - Create backend on SQL Azure with renamed front end
' %028 - Connect to SQL Azure with node.js, https://docs.microsoft.com/en-us/azure/sql-database/sql-database-connect-query-nodejs
' %024 - GH#26, Design a logo, implement in test app, add to DFAQ
' %002 - Test Helen Fedema add-in for renaming http://www.helenfeddema.com/files/Code10.zip
' %001 - Use ae standards for naming objects - Ref: https://en.wikipedia.org/wiki/Hungarian_notation,
'           https://en.wikipedia.org/wiki/Leszynski_naming_convention
'           RVBA: https://ss64.com/access/syntax-naming.html
'=============================================================================================================================
'
'20171231 - v015 -
    ' FIXED - %045 - Open map location for lat lon
'20171230 - v014 -
    ' FIXED - %044 - Fix ribbon button for splash form
    ' FIXED - %022 - GH#30, Filter by Type of Contact
'20171230 - v012 -
    ' FIXED - %043 - Error - cannot run MyAddInInitialize
    ' FIXED - %042 - Fix stray contact_details reference
    ' FIXED - %038 - Implement tblBinary for loading theme artwork internally
    ' FIXED - %027 - GH#34, Use basic validation for email addresses
    ' FIXED - %023 - GH#27, Add ribbon interface
    ' FIXED - %021 - GH#31, LAT and LON not included on the contact form
    ' FIXED - %020 - *BE - GH#32, Unique Id for tables
    ' FIXED - %003 - Relates to GH#9, include version tracking details in the app database change log module
'20171224 - v011 -
    ' FIXED - %041 - Fix forms to match id/field naming standard
    ' FIXED - %040 - Fix queries to match id/field naming standard
'20171221 - v009 -
    ' FIXED - %039 - Display document tabs
    ' FIXED - %037 - Add logo on forms and tabs
    ' FIXED - %036 - Use hover property to set sensible color of command buttons
    ' FIXED - %035 - Update look of contacts form for Access 2016
    ' FIXED - %034 - Add userdev/admin peter
    ' FIXED - %025 - GH#25, Implement frmPersist (FE) and _tblPersist (BE), Ref: https://www.devhut.net/2012/09/29/ms-access-persistent-connection-in-a-split-database/
    ' FIXED - %012 - Add Shift key blocking (basDisableShiftKey etc.)
'20171221 - v008 -
    ' FIXED - %033 - Create theme folder for icons, graphics etc. and use adaept as first test theme
    ' FIXED - %019 - Update aegit to latest (2.0.7)
    ' FIXED - %011 - GH#20, Create splash form
'20171220 - v007 -
    ' FIXED - %032 - Use basFunctions and dte for Date
'20171220 - v006 -
    ' FIXED - %026 - Fix name to ACDB
'20171208 - v005 -
    ' FIXED - %018 - *BE - Lookup tables use tlkp (not tklp)
    ' FIXED - %017 - *BE - GH#22, Create srcbe folder for back end export and then export the back end
'20171127 - v004 -
    ' FIXED - %016 - Rename linked ODBC tables to match database
    ' FIXED - %015 - Link ODBC Driver
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