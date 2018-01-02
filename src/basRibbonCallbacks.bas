Option Compare Database
Option Explicit

Public Const bolUsePicturesFromTable As Boolean = False       ' True = The images should be loaded from the table "tblBinary"
Public Const strAppPicturePath As String = "pix"              ' The pictures/icons are available below the database directory, %Databasepath%\pix
Public Const bolUseDynamicPicturePath As Boolean = False      ' The images should be loaded from this directory and not from the directory in the Ribbon XML, these values are used in the function "GetImages"
Public bolEnabled As Boolean                                  ' Used in Callback "getEnabled"
Public bolVisible As Boolean                                  ' Used in Callback "getVisible"
Public gobjRibbon As IRibbonUI

' For Sample Callback "GetContent"
Public Type ItemsVal
    id As String
    label As String
    imageMso As String
End Type

Public Sub ribOpenForm(Control As IRibbonControl)
    ' Open the form that is specified in the ribbon tag property
    DoCmd.OpenForm (Control.Tag)
End Sub

Public Sub MyAddinInitialize(ribbon As IRibbonUI)
    ' Callback name for XML "onLoad"
    Set gobjRibbon = ribbon
End Sub

' ae Buttons
Public Function aeNtryPoint(strControl As String, strAction As String)
    ' Callback name for XML "onAction"
    Select Case strControl
        Case "btn0"
            'MsgBox "Button """ & strControl & """ Action", vbInformation, "aeNtryPoint"
            DoCmd.OpenForm "frmSplash"
        Case "btn1"
            MsgBox "Button """ & strControl & """ Action", vbInformation, "aeNtryPoint"
            'DoCmd.OpenForm "frm..."
        Case "btn2"
            MsgBox "Button """ & strControl & """ Action", vbInformation, "aeNtryPoint"
            'DoCmd.OpenForm "frm..."
        Case "btn3"
            MsgBox "Button """ & strControl & """ Action", vbInformation, "aeNtryPoint"
            'DoCmd.OpenForm "frm..."
        Case "btn4"
            MsgBox "Button """ & strControl & """ Action", vbInformation, "aeNtryPoint"
            'DoCmd.OpenForm "frm..."
        Case "btn5"
            MsgBox "Button """ & strControl & """ Action", vbInformation, "aeNtryPoint"
            'DoCmd.OpenForm "frm..."
        Case "btn6"
            MsgBox "Button """ & strControl & """ Action", vbInformation, "aeNtryPoint"
            'DoCmd.OpenForm "frm..."
        Case "btn7"
            MsgBox "Button """ & strControl & """ Action", vbInformation, "aeNtryPoint"
            'DoCmd.OpenForm "frm..."
            'Debug.Print "Button """ & strControl & """ Action"
        Case Else
            MsgBox "Button """ & strControl & """ clicked", vbInformation, "aeNtryPoint"
    End Select
End Function

' Callbacks:

'Public Sub OnRibbonLoad(ribbon As IRibbonUI)
'    ' Callbackname in XML File "onLoad"
'    Set gobjRibbon = ribbon
'End Sub

Public Sub LoadImages(Control, ByRef Image)
    ' Callbackname in XML File "loadImage"
    ' Loads an image with transparency to the ribbon
    ' Modul basGDIPlus is required
    
    Dim strImage        As String
    Dim strPicture      As String
    
    strImage = CStr(Control)
    strPicture = getPic(strImage)
    
    If strImage <> "" Then
        If bolUsePicturesFromTable = True Then
            If strPicture <> "" Then
                Set Image = getIconFromTable(strPicture)
            Else
                Set Image = Nothing
            End If
        Else
            Set Image = LoadPictureGDIP(strImage)
        End If
    Else
        Set Image = Nothing
    End If
End Sub

Public Sub GetImages(Control As IRibbonControl, ByRef Image)
    ' Callbackname in XML File "getImages"
    ' Loads an image with transparency to the ribbon
    ' Modul basGDIPlus is required

    Dim strPicturePath  As String
    Dim strPicture      As String

    strPicture = getTheValue(Control.Tag, "CustomPicture")

    If bolUsePicturesFromTable = True Then
        Set Image = getIconFromTable(strPicture)
    Else
        If bolUseDynamicPicturePath = True Then
            strPicturePath = getAppPath & strAppPicturePath & "\"
        Else
            strPicturePath = getTheValue(Control.Tag, "CustomPicturePath")
        End If
        Set Image = LoadPictureGDIP(strPicturePath & strPicture)
    End If
End Sub

Public Sub GetEnabled(Control As IRibbonControl, ByRef Enabled)
    ' Callbackname in XML File "getEnabled"
    ' To set the property "enabled" to a Ribbon Control
    ' For further information see: http://www.accessribbon.de/en/index.php?Downloads:12

    Select Case Control.id
        Case Else
            Enabled = True
    End Select
End Sub

Public Sub GetVisible(Control As IRibbonControl, ByRef Visible)
    ' Callbackname in XML File "getVisible"
    ' To set the property "Visible" to a Ribbon Control
    ' For further information see: http://www.accessribbon.de/en/index.php?Downloads:12

    Select Case Control.id
        Case Else
            Visible = True
    End Select
End Sub

Public Sub GetLabel(Control As IRibbonControl, ByRef label)
    ' Callbackname in XML File "getLabel"
    ' To set the property "label" to a Ribbon Control

    Select Case Control.id
        ''GetLabel''
        Case Else
            label = "*getLabel*"
    End Select
End Sub

Public Sub GetScreentip(Control As IRibbonControl, ByRef screentip)
    ' Callbackname in XML File "getScreentip"
    ' To set the property "screentip" to a Ribbon Control

    Select Case Control.id
        Case Else
            screentip = "*getScreentip*"
    End Select
End Sub

Public Sub GetSupertip(Control As IRibbonControl, ByRef supertip)
    ' Callbackname in XML File "getSupertip"
    ' To set the property "supertip" to a Ribbon Control

    Select Case Control.id
        Case Else
            supertip = "*getSupertip*"
    End Select
End Sub

Public Sub GetDescription(Control As IRibbonControl, ByRef Description)
    ' Callbackname in XML File "getDescription"
    ' To set the property "Description" to a Ribbon Control

    Select Case Control.id
        Case Else
            Description = "*getDescription*"
    End Select
End Sub

Public Sub GetTitle(Control As IRibbonControl, ByRef title)
    ' Callbackname in XML File "getTitle"
    ' To set the property "title" to a Ribbon MenuSeparator Control

    Select Case Control.id
        Case Else
            title = "*getTitle*"
    End Select
End Sub

' Button

Public Sub OnActionButton(Control As IRibbonControl)
    ' Callbackname in XML File "onAction"
    ' Callback for event button click
    
    Select Case Control.id
        Case Else
            MsgBox "Button """ & Control.id & """ clicked", vbInformation
    End Select
End Sub

' Command Button

Public Sub OnActionButtonHelp(Control As IRibbonControl, ByRef CancelDefault)
    ' Callbackname in XML File Command "onAction"
    ' Callback for command event button click

    MsgBox "Button ""Help"" clicked", vbInformation
    CancelDefault = True
End Sub

' CheckBox

Sub OnActionCheckBox(Control As IRibbonControl, _
                     pressed As Boolean)
    ' Callbackname in XML File "OnActionCheckBox"
    ' Callback for event checkbox click

    Select Case Control.id
        Case Else
            MsgBox "The Value of the Checkbox """ & Control.id & """ is: " & pressed, vbInformation
    End Select
End Sub

Sub GetPressedCheckBox(Control As IRibbonControl, _
                       ByRef bolReturn)
    ' Callbackname in XML File "GetPressedCheckBox"
    ' Callback for checkbox
    ' indicates how the control is displayed

    Select Case Control.id
        Case Else
            If getTheValue(Control.Tag, "DefaultValue") = "1" Then
                bolReturn = True
            Else
                bolReturn = False
            End If
    End Select
End Sub

' ToggleButton

Sub OnActionTglButton(Control As IRibbonControl, _
                      pressed As Boolean)
                              
    ' Callbackname in XML File "onAction"

   Select Case Control.id
        Case Else
            MsgBox "The Value of the Toggle Button """ & Control.id & """ is: " & pressed, vbInformation
    End Select
End Sub

Public Sub GetPressedTglButton(Control As IRibbonControl, _
                        ByRef pressed)
    ' Callbackname in XML File "getPressed"
    ' Callback for an Access ToogleButton Control. Indicates how the control is displayed

    Select Case Control.id
        Case Else
            If getTheValue(Control.Tag, "DefaultValue") = "1" Then
                pressed = True
            Else
                pressed = False
            End If
    End Select
End Sub

'EditBox

Public Sub GetTextEditBox(Control As IRibbonControl, _
                   ByRef strText)
    ' Callbackname in XML File "GetTextEditBox"
    ' Callback for an EditBox Control
    ' Indicates which value is to set to the control

    Select Case Control.id
        Case Else
            strText = getTheValue(Control.Tag, "DefaultValue")
    End Select
End Sub

Public Sub OnChangeEditBox(Control As IRibbonControl, _
                    strText As String)
    ' Callbackname in XML File "OnChangeEditBox"
    ' Callback Editbox: Return value of the Editbox

    Select Case Control.id
        Case Else
            MsgBox "The Value of the EditBox """ & Control.id & """ is: " & strText, vbInformation
    End Select
End Sub

' DropDown

Public Sub OnActionDropDown(Control As IRibbonControl, _
                     selectedId As String, _
                     selectedIndex As Integer)
    ' Callbackname in XML File "OnActionDropDown"
    ' Callback onAction (DropDown)
    
    Select Case Control.id
        Case Else
            Select Case selectedId
                Case Else
                    MsgBox "The selected ItemID of DropDown-Control """ & Control.id & """ is : """ & selectedId & """", vbInformation
            End Select
    End Select
End Sub

Public Sub GetSelectedItemIndexDropDown(Control As IRibbonControl, _
                                 ByRef Index)
    ' Callbackname in XML File "GetSelectedItemIndexDropDown"
    ' Callback getSelectedItemIndex
    
    Dim varIndex As Variant
    varIndex = getTheValue(Control.Tag, "DefaultValue")
    
    If IsNumeric(varIndex) Then
        Select Case Control.id
            Case Else
                Index = getTheValue(Control.Tag, "DefaultValue")
        End Select
    End If
End Sub

' Gallery

Public Sub GetSelectedItemIndexGallery(Control As IRibbonControl, _
                                   ByRef Index)
    ' Callbackname in XML File "GetSelectedItemIndexGallery"
    ' Callback GetSelectedItemIndexGallery
    
    Dim varIndex As Variant
    varIndex = getTheValue(Control.Tag, "DefaultValue")
    
    If IsNumeric(varIndex) Then
        Select Case Control.id
            Case Else
                Index = varIndex
        End Select
    End If
End Sub

Public Sub OnActionGallery(Control As IRibbonControl, _
                     selectedId As String, _
                     selectedIndex As Integer)
    ' Callbackname in XML File "OnActionGallery"
    ' Callback onAction (Gallery)
    
    Select Case Control.id
        Case Else
            Select Case selectedId
                Case Else
                    MsgBox "The selected ItemID of Gallery-Control """ & Control.id & """ is : """ & selectedId & """", vbInformation
            End Select
    End Select
End Sub

' Combobox

Public Sub GetTextComboBox(Control As IRibbonControl, _
                      ByRef strText)
    ' Callbackname im XML File "GetTextComboBox"
    ' Callback getText (Combobox)
                           
    Select Case Control.id
        Case Else
            strText = getTheValue(Control.Tag, "DefaultValue")
    End Select
End Sub

Public Sub OnChangeComboBox(Control As IRibbonControl, _
                               strText As String)
    ' Callbackname im XML File "OnChangeCombobox"
    ' Callback onChange (Combobox)
   
    Select Case Control.id
        Case Else
            MsgBox "The selected Item of Combobox-Control """ & Control.id & """ is : """ & strText & """", vbInformation
    End Select
End Sub

' DynamicMenu

Public Sub GetContent(Control As IRibbonControl, ByRef XMLString)
    ' Sample for a Ribbon XML "getContent" Callback
    ' See also http://www.accessribbon.de/en/index.php?Access_-_Ribbons:Callbacks:dynamicMenu_-_getContent
    '     and: http://www.accessribbon.de/en/index.php?Access_-_Ribbons:Ribbon_XML___Controls:Dynamic_Menu

    Select Case Control.id
        Case Else
            XMLString = getXMLForDynamicMenu()
    End Select
End Sub

' Helper Function

Public Function getXMLForDynamicMenu() As String
    ' Creates a XML String for DynamicMenu CallBack - getContent
   
    Dim lngDummy    As Long
    Dim strDummy    As String
    Dim strContent  As String
    
    Dim Items(4) As ItemsVal
    Items(0).id = "btnDy1"
    Items(0).label = "Item 1"
    Items(0).imageMso = "_1"
    Items(1).id = "btnDy2"
    Items(1).label = "Item 2"
    Items(1).imageMso = "_2"
    Items(2).id = "btnDy3"
    Items(2).label = "Item 3"
    Items(2).imageMso = "_3"
    Items(3).id = "btnDy4"
    Items(3).label = "Item 4"
    Items(3).imageMso = "_4"
    Items(4).id = "btnDy5"
    Items(4).label = "Item 5"
    Items(4).imageMso = "_5"
    
    strDummy = "<menu xmlns=""http://schemas.microsoft.com/office/2009/07/customui"">" & vbCrLf
    
        For lngDummy = LBound(Items) To UBound(Items)
            strContent = strContent & _
                "<button id=""" & Items(lngDummy).id & """" & _
                " label=""" & Items(lngDummy).label & """" & _
                " imageMso=""" & Items(lngDummy).imageMso & """" & _
                " onAction=""OnActionButton""/>" & vbCrLf
        Next

    strDummy = strDummy & strContent & "</menu>"
    getXMLForDynamicMenu = strDummy
End Function

Public Function getTheValue(strTag As String, strValue As String) As String
    ' *************************************************************
    ' Created from     : Avenius
    ' Parameter        : Input String, SuchValue String
    ' Date created     : 05.01.2008
    '
    ' Sample:
    ' getTheValue("DefaultValue:=Test;Enabled:=0;Visible:=1", "DefaultValue")
    ' Return           : "Test"
    ' *************************************************************
      
    On Error Resume Next
      
    Dim workTb()     As String
    Dim Ele()        As String
    Dim myVariabs()  As String
    Dim i            As Integer

    workTb = Split(strTag, ";")
      
    ReDim myVariabs(LBound(workTb) To UBound(workTb), 0 To 1)
    For i = LBound(workTb) To UBound(workTb)
        Ele = Split(workTb(i), ":=")
        myVariabs(i, 0) = Ele(0)
        If UBound(Ele) = 1 Then
            myVariabs(i, 1) = Ele(1)
        End If
    Next

    For i = LBound(myVariabs) To UBound(myVariabs)
        If strValue = myVariabs(i, 0) Then
            getTheValue = myVariabs(i, 1)
        End If
    Next
End Function

Public Function getAppPath() As String
    Dim strDummy As String
    strDummy = CurrentProject.Path
    If Right(strDummy, 1) <> "\" Then strDummy = strDummy & "\"
    getAppPath = strDummy
End Function

Public Function getIconFromTable(strFileName As String) As Picture

    Dim lSize As Long
    Dim arrBin() As Byte
    Dim rs As DAO.Recordset
 
    On Error GoTo PROC_ERR
 
    Set rs = DBEngine(0)(0).OpenRecordset("tblBinary", dbOpenDynaset)
    rs.FindFirst "[FileName]='" & strFileName & "'"
    If rs.NoMatch Then
        Set getIconFromTable = Nothing
    Else
        lSize = rs.Fields("binary").FieldSize
        ReDim arrBin(lSize)
        arrBin = rs.Fields("binary").GetChunk(0, lSize)
        Set getIconFromTable = ArrayToPicture(arrBin)
    End If
    rs.Close
 
PROC_EXIT:
    Reset
    Erase arrBin
    Set rs = Nothing
    Exit Function

PROC_ERR:
    Resume PROC_EXIT

End Function

Public Function getPic(strFullPath As String) As String
    Dim strResult As String
    
    If InStrRev(strFullPath, "\") > 0 Then
        strResult = Mid(strFullPath, InStrRev(strFullPath, "\") + 1)
    Else
        strResult = ""
    End If
   
    getPic = strResult
End Function