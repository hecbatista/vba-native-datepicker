VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmExample 
   Caption         =   "UserForm1"
   ClientHeight    =   3036
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4584
   OleObjectBlob   =   "frmExample.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmExample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'========================================================================
' frmExample - Minimal example of wiring up the native date picker
'========================================================================
'PURPOSE
' Shows the complete pattern for a form with date fields. Copy this
' pattern for any form that needs a date. Nothing here is required
' by the picker itself - it just demonstrates how a form uses
' modDatePicker.
'
'TO REPRODUCT THIS FORM
' Add these controls to a UserForm named frmExample:
'   - txtStartDate (TextBox) - a date field
'   - cmdPickStart (CommandButton) - a "..." pick button next to it
'   - txtEndDate (TextBox)  - a second date field
'   - cmdPickEnd (CommandButton) - its pick button
'   - cmdOK (CommandButton) - closes the form
'
'DEPENDENCIES
' Requires modDatePicker, frmCalendar, and clsDayButton in the project
'========================================================================
Option Explicit

'------------------------------------------------------------------------
' Default both date boxes to today (formatted to match the picker's
' output)
'------------------------------------------------------------------------
Private Sub UserForm_Initialize()
    txtStartDate = Format(Date, "m/d/yyyy")
    txtEndDate = Format(Date, "m/d/yyyy")
End Sub

'------------------------------------------------------------------------
' Pick buttons: open the calendar seeded from the box and write the
' result back.
'------------------------------------------------------------------------
Private Sub cmdPickStart_Click(): FillDate Me.txtStartDate: End Sub
Private Sub cmdPickEnd_Click(): FillDate Me.txtEndDate: End Sub

'------------------------------------------------------------------------
' AfterUpdate: tidy a typed date to m/d/yyyy when the user leaves the
' field.
' The sub name must match the TextBox name exactly, or it never fires.
'------------------------------------------------------------------------
Private Sub txtStartDate_AfterUpdate(): NormalizeDateBox Me.txtStartDate: End Sub
Private Sub txtEndDate_AfterUpdate(): NormalizeDateBox Me.txtEndDate: End Sub
    
'------------------------------------------------------------------------
' OK: validate both dates before closing (optional but recommended).
'------------------------------------------------------------------------
Private Sub cmdOK_Click()
    If Not IsDate(Me.txtStartDate.Text) Then
        MsgBox "Please enter a valid start date.", vbExclamation
        Me.txtStartDate.SetFocus
        Exit Sub
    End If
    If Not IsDate(Me.txtEndDate.Text) Then
        MsgBox "Please enter a valid end date.", vbExclamation
        Me.txtEndDate.SetFocus
        Exit Sub
    End If
    Me.Hide
End Sub
