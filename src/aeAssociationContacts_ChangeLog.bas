Option Compare Database
Option Explicit

' Tools:
' MZ-Tools 8.0 for VBA - Ref: http://www.mztools.com/index.aspx
' TM VBA-Inspector - Ref: http://www.team-moeller.de/en/?Add-Ins:TM_VBA-Inspector
' RibbonX Visual Designer 2010 - Ref: http://www.andypope.info/vba/ribboneditor_2010.htm
' IDBE RibbonCreator 2016 (Office 2016) - Ref: http://www.ribboncreator2016.de/en/?Download
' V-Tools - Ref: http://www.skrol29.com/us/vtools.php
' Bill Mosca - Ref: http://www.thatlldoit.com/Pages/utilsaddins.aspx
' Rubberduck - Ref: https://github.com/rubberduck-vba/Rubberduck
' DataNumen Access Repair - Ref: https://www.datanumen.com/access-repair/
'
'
' Research:
' Ref: http://www.msoutlook.info/question/482 - officeUI-files
' The Ribbon and QAT settings - C:\Users\%username%\AppData\Local\Microsoft\Office
' Ref: http://msdn.microsoft.com/en-us/library/ee704589(v=office.14).aspx
' *** Windows API help - replacing As Any declaration
' Ref: http://allapi.mentalis.org/vbtutor/api1.shtml
' Ref: http://programmersheaven.com/discussion/237489/passing-an-array-as-an-optional-parameter
' *** Example of SQL INSERT / UPDATE using ADODB.Command and ADODB.Parameters objects
' Ref: http://www.access-programmers.co.uk/forums/showthread.php?t=219149
' *** CreateObject("System.Collections.ArrayList")
' Ref: http://www.ozgrid.com/forum/showthread.php?t=167349
' Microsoft Access - Really useful queries - Ref: http://www.sqlquery.com/Microsoft_Access_useful_queries.html
' Ref: http://www.micronetservices.com/manage_remote_backend_access_database.htm
' Microsoft Access Tips and Tricks - Ref: http://www.datagnostics.com/tips.html
'
'
' Guides:
' Office VBA Basic Debugging Techniques
' Ref: http://pubs.logicalexpressions.com/pub0009/LPMArticle.asp?ID=410
' *** Ref: http://www.vb123.com/toolshed/02_accvb/remotequeries.htm - Remote Queries In Microsoft Access
' Ref: http://social.msdn.microsoft.com/Forums/office/en-US/f8a050b9-3e12-465e-9448-36be59827581/vba-code-redirect-results-from-immediate-window-to-an-access-table-or-csv-file?forum=accessdev
' Access Articles- Ref: http://www.databasejournal.com/article.php/1464721
' Long Binary Data - Ref: http://www.ammara.com/support/technologies/long-binary-data.html
'
'
'=============================================================================================================================
' Tasks:
' %005 -
' %004 -
' %003 -
' %002 - Test Helen Fedema add-in for renaming http://www.helenfeddema.com/files/Code10.zip
' %001 - Use ae standards for naming objects - Ref: https://en.wikipedia.org/wiki/Hungarian_notation,
'           https://en.wikipedia.org/wiki/Leszynski_naming_convention
'           RVBA: https://ss64.com/access/syntax-naming.html
'=============================================================================================================================
'
'
'20171009 - v001 - Initial database design based on a sample from:
    ' Ref: https://www.devhut.net/2016/09/01/ms-access-contact-database-template-sample/
    ' Daniel Pineault, Microsoft MVP, 2010-2017
    ' "I am truly pleased to announce that I have been awarded the Title of Microsoft Most Valuable Professional (MVP), years in a row,
    ' for my contributions to the MS Access community. It is a great honor to receive this award directly form the hands of Microsoft
    ' and it is my pleasure to help other developers when I can. This is one of the main reason why I created this website in the first place,
    ' to share knowledge, no strings attached."