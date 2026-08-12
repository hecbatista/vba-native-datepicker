Attribute VB_Name = "modDatePicker"
Option Explicit

'========================================================================
' modDatePicker - Native (ActiveX-free) date picker service
'========================================================================
'PURPOSE
' Provides a pop-up calendar date picker built entirely from native
' MSForms controls. This replaces the old MSCOMCT2 "DTPicker" ActiveX
' control, which required per-machine registration and licensing and does
' NOT exist for 64-bit Office. Nothing here needs to be registered; it
' behaves indentically on 32-bit and 64-bit Office.
'
'THE SYSTEM HAS THREE PARTS THAT WORK TOGETHER
' 1. modDatePicker (this module) - the public entry points forms call
' 2. frmCalendar   (UserForm)    - the pop-up calendar window itself
' 3. clsDayButton  (class)       - wires the Click event for each day cell.
' Keep all three in the project. Removing any one breaks the picker.
'
'HOW A FORM USES IT (see frmChargingInfo for a live example)
'   Add a TextBox (e.g. txtVerdict) and a small "Pick" CommandButton beside it.
'       - In the button's Click event:        FillDate Me.txtVerdict
'       - In the TextBox's AfterUpdate event: NormalizeDateBox Me.txtVerdict
'   No calendar code lives on the form - it only calls these two routines.
'
'DATE FORMAT
'   All dates are written as "m/d/yyyy" (e.g. 8/23/2024 - no leading zeros).
'   IMPORTANT: FillDate and NormalizeDateBox must use the SAME format string
'   so a typed date and a picked date look identical in the box.
'========================================================================

'------------------------------------------------------------------------
' PickDate - shows the calendar and returns the date the user chose.
'
'   StartDate : (optional) the month the calendar opens on, and the day it
'               pre-highlights. Omit or pass 0 to open on today.
'   Returns   : a Date if the user picked a day, or Null if they cancelled
'               (clicked Cancel or the window's X)
'
'   Callers MUST test the result with IsNull before using it, because a
'   cancel returns null, not a date. Example:
'       d = PickDate()
'       If Not IsNull(d) Then ...use d...
'------------------------------------------------------------------------
Public Function PickDate(Optional ByVal StartDate As Date = 0) As Variant
    Dim f As frmCalendar
    Set f = New frmCalendar                 ' a fresh instance each call = no stale state
    
    If StartDate = 0 Then StartDate = Date
    f.SetStart StartDate                    ' tell the calendar which month/day to show
    
    f.Show                                  ' modal - execution pauses here until the
                                            ' the user picks a day or cancels
    
    If f.Cancelled Then
        PickDate = Null
    Else
        PickDate = f.SelectedDate
    End If
    
    Unload f
    Set f = Nothing
End Function

'------------------------------------------------------------------------
' FillDate - the all-in-one helper a "pick date" button calls.
'            Opens the calendar seeded from a TextBox's current value,
'            then writes the chosen date back into the same TextBox.

'   Usage (if the button's Click handler):  FillDate Me.txtVerdict
'
'   - If the box already holds a valid date, the calendar opens on it.
'   - If it's empty or holds junk, the calendar opens on today.
'   - If the user cancels, the box is left unchanged.
'------------------------------------------------------------------------
Public Sub FillDate(tb As MSForms.TextBox)
    Dim seed As Date, d As Variant
    If IsDate(tb.Text) Then seed = CDate(tb.Text) Else seed = Date
    d = PickDate(seed)
    If Not IsNull(d) Then tb.Text = Format(d, "m/d/yyyy")
End Sub

'------------------------------------------------------------------------
' NormalizeDateBox - tidies a date TextBox to m/d/yyyy after the user
'                 finishes typing (so "8/4/24" becomes "8/4/2024", etc.)

'   Usage (in the TextBox's AfterUpdate handler):  NormalizeDateBox Me.txtVerdict
'
'   Empty or non-date text is deliberatly left untouched, so a final validation pass
'   (e.g. on the OK button) can decide how to handle it.
'   AfterUpdate (not Change) is the right event: it fires when the user leaves the
'   field, not on every keystroke.
'------------------------------------------------------------------------
Public Sub NormalizeDateBox(tb As MSForms.TextBox)
    If Len(Trim(tb.Text)) = 0 Then Exit Sub
    If IsDate(tb.Text) Then
        tb.Text = Format(CDate(tb.Text), "m/d/yyyy")
    End If
End Sub


