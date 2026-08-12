VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDayButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'========================================================================
' clsDayButton - Click handler for one day cell in frmCalendar
'========================================================================
'PURPOSE
' frmCalendar builts it's 42 day buttons at run time. Dynamically
' created controls do not raise events on their own, so each day button
' is wrapped in one instance of this class. The WithEvents reference below
' is what lets the button's Click actually fire code.
'
'HOW IT IS USED (see frmCalendar.BuildControls)
'   For each of the 42 cells the form does:
'       Set cls = New clsDayButton
'       Set cls.Btn = <the CommandButton>   ' hooks up the Click event
'       Set cls.Form = Me                   ' so we can call back into the form
'   and keeps the insteance in its mDays colletion. RenderMonth then sets each
'   cell's TheDate / HasDate as it lays out the month.
'
'NOTE
' Blank leading/trailing cells have HasDate = False, so clicking one does nothing
' (the guard in Btn_Click below).
'========================================================================
Public WithEvents Btn As MSForms.CommandButton  'the day button this wraps
Attribute Btn.VB_VarHelpID = -1
Public TheDate As Date                          'the date this cell represents
Public HasDate As Boolean                       'False for blank filler cells
Public Form As Object                           'back-reference to frmCalendar

'------------------------------------------------------------------------
' Btn_Click - fired when the user clicks this day. Ignores blank cells;
'             otherwise tells the calendar which date was chosen.
'------------------------------------------------------------------------
Private Sub Btn_Click()
    If HasDate Then Form.OnDayClick TheDate
End Sub

