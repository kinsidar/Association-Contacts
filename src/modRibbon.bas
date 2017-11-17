Option Compare Database
Option Explicit

' Open the form that is specified in the ribbon tag property
Public Sub ribOpenForm(Control As IRibbonControl)
    DoCmd.OpenForm (Control.Tag)
End Sub