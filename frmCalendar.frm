VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCalendar 
   Caption         =   "UserForm1"
   ClientHeight    =   3036
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4584
   OleObjectBlob   =   "frmCalendar.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmCalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'========================================================================
' frmCalendar - Pop-up calendar for the native date picker
'========================================================================
'PURPOSE
' A month-grid calendar window built entirely in code from native MSForms
' controls (no ActiveX, no registration, works on 32- and 64-bit Office).
' It is driven by modDatePicker.PickDate - forms never talk to it directly.
'
'HOW IT IS DRIVEN (the calling sequence matters - see the timing notes below.
' modDatePicker.PickDate does in order:
'   Set f = New frmCalendar -> fires UserForm_Initialize (builds controls)
'   f.SetStart someDate     -> stores which month/day to show
'   f.Show                  -> fires UserForm_Activate (first RenderMonth)
' When the user clicks day, OnDayClick records it and hide the form;
' PickDate then reads SelectedDate / Cancelled.
'
' PUBLIC INTERFACE (what PickDate reads/sets)
'   SetStart(d) - call BEFORE Show to open on a given month and highlight d
'   SelectDate  - the day the user clicked (valid only if Cancelled = False)
'   Cancelled   - True if the user cancelled or closed the window
'
' KEY TIMING NOTES - DO NOT "SIMPLIFY" THESE OR IT WILL BREAK
' * The FIRST RenderMonth is deferred to UserForm_Activate(), NOT initalize.
'   Dynamically-added ComboBoxes cannot accept ListIndex/TopIndex until the
'   form is actually shown
' * Initialize must NOT set mViewYear/mViewMonth. SetStart runs after
'   initalize, so if Initalize set them to today it would overwrite the seed
'   and the calendar would always open on the current month.
' * mNavigating guards against re-entrancy: when RenderMonth sets the month
'   and year combos, that fires their _Change events, which would call
'   RenderMonth again. The guard makes those programmatic changes no-ops so
'   only real USER picks navigate.
'   (This is what fixed the "skips a month bug" when clicking the > arrow.)
'========================================================================

' ---- Public results, read by modDatePicker.PickDate ----
Public SelectedDate As Date
Public Cancelled As Boolean

' ---- Internal State ----
Private mSeedDate As Date
Private mViewYear As Integer
Private mViewMonth As Integer
Private mDays As Collection

' ---- Controls built at run time (see BuildControls) ----
Private WithEvents mBtnPrev As MSForms.CommandButton
Attribute mBtnPrev.VB_VarHelpID = -1
Private WithEvents mBtnNext As MSForms.CommandButton
Attribute mBtnNext.VB_VarHelpID = -1
Private WithEvents mBtnCancel As MSForms.CommandButton
Attribute mBtnCancel.VB_VarHelpID = -1
Private WithEvents mCboMonth As MSForms.ComboBox
Attribute mCboMonth.VB_VarHelpID = -1
Private WithEvents mCboYear As MSForms.ComboBox
Attribute mCboYear.VB_VarHelpID = -1

' ---- Guard flags (see timing notes in header) ----
Private mNavigating As Boolean
Private mShown As Boolean

' ---- Layout constants (points) ----
Private Const MARGIN As Single = 6
Private Const CELLW As Single = 30
Private Const CELLH As Single = 20
Private Const GRIDTOP As Single = 50

'------------------------------------------------------------------------
' Initialize - builds all controls once. Deliberatly does NOT render
'              or set the view month (see timing notes) - that happens
'              in activate.
'------------------------------------------------------------------------
Private Sub UserForm_Initialize()
    Cancelled = True    'default to "cancelled" until a day is picked
    Set mDays = New Collection
    BuildControls
End Sub

'------------------------------------------------------------------------
' Activate - fires as part of .Show, when controls are finally realized.
'            Does the FIRST render here (not in initalize) so the combos
'            can accept ListIndex. Falls back to today if SetStart was
'            never called.
'------------------------------------------------------------------------
Private Sub UserForm_Activate()
    If Not mShown Then
        mShown = True
        If mViewYear = 0 Then
            mViewYear = Year(Date)
            mViewMonth = Month(Date)
        End If
        RenderMonth
    End If
End Sub
    
'------------------------------------------------------------------------
' SetStart - called by PickDate BEFORE the form is shown. Records the
'            month to open on and the day to highlight. Renders
'            immediately only if the form is already visible.
'            (It isn't on the normal first call)
'------------------------------------------------------------------------
Public Sub SetStart(ByVal d As Date)
    mSeedDate = d
    mViewYear = Year(d)
    mViewMonth = Month(d)
    If mShown Then RenderMonth
End Sub

'------------------------------------------------------------------------
' BuildControls - creates every control in code (nav buttons, month/year
'                 dropdowns, weekday headers, the 6x7 grid of day buttons,
'                 and Cancel). Building in code is why the form needs no
'                 design-time controls and stays ActiveX-free.
'------------------------------------------------------------------------
Private Sub BuildControls()
    Dim i As Long, r As Long, c As Long
    Dim lbl As MSForms.Label, Btn As MSForms.CommandButton
    Dim cls As clsDayButton
    Dim days7 As Variant
    
    Me.Caption = "Select a date"
    Me.Width = MARGIN * 2 + CELLW * 7 + 12
    Me.Height = GRIDTOP + CELLH * 6 + 60
    
    ' Previous / next month arrows
    Set mBtnPrev = Me.Controls.Add("Forms.CommandButton.1", "btnPrev", True)
    mBtnPrev.Caption = "<": mBtnPrev.Left = MARGIN: mBtnPrev.Top = 6
    mBtnPrev.Width = 26: mBtnPrev.Height = 18
    
    Set mBtnNext = Me.Controls.Add("Forms.CommandButton.1", "btnNext", True)
    mBtnNext.Caption = ">": mBtnNext.Left = MARGIN + CELLW * 7 - 26: mBtnNext.Top = 6
    mBtnNext.Width = 26: mBtnNext.Height = 18
    
    Dim m As Long, y As Long
    
    ' Month dropdown (pick-only, so users can't type an invalid month)
    Set mCboMonth = Me.Controls.Add("Forms.ComboBox.1", "cboMonth", True)
    mCboMonth.Left = MARGIN + 30: mCboMonth.Top = 6
    mCboMonth.Width = 86
    mCboMonth.Style = fmStyleDropDownList   'pick-only, no typing
    For m = 1 To 12
        mCboMonth.AddItem Format(DateSerial(2000, m, 1), "mmmm")
    Next m
    
    ' Year dropdown. Range is today +/- a span; SyncYearCombo will append
    ' any year navigated to outside this range, so you can never get stuck.
    Set mCboYear = Me.Controls.Add("Forms.ComboBox.1", "cboYear", True)
    mCboYear.Left = MARGIN + 120: mCboYear.Top = 6
    mCboYear.Width = 58
    mCboYear.Style = fmStyleDropDownList
    For y = Year(Date) - 30 To Year(Date) + 15
        mCboYear.AddItem y
    Next y
    
    ' Weekday header row (Sun..Sat)
    days7 = Array("Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat")
    For i = 0 To 6
        Set lbl = Me.Controls.Add("Forms.Label.1", "wd" & i, True)
        lbl.Caption = days7(i)
        lbl.Left = MARGIN + i * CELLW: lbl.Top = 30
        lbl.Width = CELLW: lbl.Height = 14
        lbl.TextAlign = fmTextAlignCenter
    Next i
    
    ' 42 day buttons (6 rows x 7 cols). Each is wrapped in a clsDayButton so
    ' its Click can be handled; the wrappers are kept in mDays (1-based).
    For i = 0 To 41
        r = i \ 7: c = i Mod 7
        Set Btn = Me.Controls.Add("Forms.CommandButton.1", "day" & i, True)
        Btn.Left = MARGIN + c * CELLW: Btn.Top = GRIDTOP + r * CELLH
        Btn.Width = CELLW: Btn.Height = CELLH
        Set cls = New clsDayButton
        Set cls.Btn = Btn
        Set cls.Form = Me
        mDays.Add cls
    Next i
    
    Set mBtnCancel = Me.Controls.Add("Forms.CommandButton.1", "btnCancel", True)
    mBtnCancel.Caption = "Cancel"
    mBtnCancel.Width = 60: mBtnCancel.Height = 20
    mBtnCancel.Top = GRIDTOP + CELLH * 6 + 6
    mBtnCancel.Left = Me.InsideWidth - 60 - MARGIN
End Sub

'------------------------------------------------------------------------
' RenderMonth - fills the grid for mViewYear/mViewMonth: set the
'               dropdowns to match, labels each day cell, disables the
'               blank leading/trailing cells, and colours today (yellow)
'               and the seeded day (blue). Wrapped in mNavigating so the
'               combo syncs don't retrigger it.
'------------------------------------------------------------------------
Private Sub RenderMonth()
    Dim firstDow As Long, daysInMonth As Long, i As Long, dayNum As Long
    Dim firstOfMonth As Date, cls As clsDayButton
     
    mNavigating = True  'suppress combo _Change while we sync them
    
    firstOfMonth = DateSerial(mViewYear, mViewMonth, 1)
    firstDow = Weekday(firstOfMonth, vbSunday) - 1  '0=Sun... 6=Sat
    daysInMonth = Day(DateSerial(mViewYear, mViewMonth + 1, 0)) 'day 0 of next month
    
    mCboMonth.ListIndex = mViewMonth - 1
    SyncYearCombo
    
    For i = 1 To 42
        Set cls = mDays(i)
        dayNum = i - firstDow   ' cell 1 maps to (1 - firstDow); blanks are <=0
        If dayNum >= 1 And dayNum <= daysInMonth Then
            cls.TheDate = DateSerial(mViewYear, mViewMonth, dayNum)
            cls.HasDate = True
            cls.Btn.Caption = CStr(dayNum)
            cls.Btn.Enabled = True
            If cls.TheDate = Int(mSeedDate) Then
                cls.Btn.BackColor = RGB(180, 215, 255)  ' seeded/selected day - blue
            ElseIf cls.TheDate = Date Then
                cls.Btn.BackColor = RGB(255, 244, 200)  ' today - pale yellow
            Else
                cls.Btn.BackColor = &H8000000F  'default button face
            End If
        Else
        ' leading/trailing blank cell - no date, disabled
            cls.HasDate = False
            cls.Btn.Caption = ""
            cls.Btn.Enabled = False
            cls.Btn.BackColor = &H8000000F
        End If
    Next i
    
    mNavigating = False
End Sub

'------------------------------------------------------------------------
' OnDayClick - called by clsDayButton when its day is clicked. Records
'              the chosen date and hides the form, returning control
'              to PickDate.
'------------------------------------------------------------------------
Public Sub OnDayClick(ByVal d As Date)
    SelectedDate = d
    Cancelled = False
    Me.Hide
End Sub

'------------------------------------------------------------------------
' Previous / next month navigation (rolls the year over at the boundaries)
'------------------------------------------------------------------------
Private Sub mBtnPrev_Click()
    mViewMonth = mViewMonth - 1
    If mViewMonth < 1 Then
        mViewMonth = 12
        mViewYear = mViewYear - 1
    End If
    RenderMonth
End Sub

Private Sub mBtnNext_Click()
    mViewMonth = mViewMonth + 1
    If mViewMonth > 12 Then
        mViewMonth = 1
        mViewYear = mViewYear + 1
    End If
    RenderMonth
End Sub

Private Sub mBtnCancel_Click()
    Cancelled = True
    Me.Hide
End Sub

'------------------------------------------------------------------------
' QueryClose - closing via the window X counts as a cancel. We HIDE
'              instead of unloading so PickDate can still read
'              Cancelled/SelectedDate.
'------------------------------------------------------------------------
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ' Closing via the X = cancel. Hide (not unload) so PickDate can read the result.
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        Cancelled = True
        Me.Hide
    End If
End Sub

'------------------------------------------------------------------------
' SyncYearCombo - selects mViewYear in the year dropdown, appending
'                 it first if it falls outside the pre-built range
'                 (so navigation never fails for far-out years).
'------------------------------------------------------------------------
Private Sub SyncYearCombo()
    Dim i As Long
    For i = 0 To mCboYear.ListCount - 1
        If CLng(mCboYear.List(i)) = mViewYear Then
            mCboYear.ListIndex = i
            Exit Sub
        End If
    Next i
    ' year isn't in the list (navigated outside the range) - add it
    mCboYear.AddItem mViewYear
    mCboYear.ListIndex = mCboYear.ListCount - 1
End Sub

'------------------------------------------------------------------------
' Dropdown navigation. Both bail out while mNavigating is True so that the
' programmatic syncs inside RenderMonth don't cause recursive re-rendering;
' only a real user selection changes the month/year.
'------------------------------------------------------------------------
Private Sub mCboMonth_Change()
    If mNavigating Then Exit Sub
    If mCboMonth.ListIndex >= 0 Then
        mViewMonth = mCboMonth.ListIndex + 1
        RenderMonth
    End If
End Sub

Private Sub mCboYear_Change()
    If mNavigating Then Exit Sub
    If mCboYear.ListIndex >= 0 Then
        mViewYear = CLng(mCboYear.List(mCboYear.ListIndex))
        RenderMonth
    End If
End Sub
