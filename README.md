# VBA Native Date Picker

A pop-up calendar date picker for Word/Excel VBA UserForms, built entirely from native MSForms controls. No ActiveX, no registration of MSCOMCT2.OCX, no license keys - works identically on 32-bit and 64-bit Office.

Created as a drop-in replacement for the old MSCOMCT2 `DTPicker` control, which requires per-machine registration and does not exist for 64-bit Office.

## Components (all three are required)

- `modDatePicker.bas` - public entry points your forms call
- `frmCalendar.frm` - the pop-up calendar window
- `clsDayButton.cls` - click handler for each day cell

## Install

1. In the VBE (Alt+F11), File -> Import File for each of the three files.
2. Compile (Debug -> Compile Project) to confirm it's clean.

## Usage

Add a TextBox and a small "pick" button beside it on your form, then:
```
  ' pick button:
  Private Sub cmdPickDate_Click(): FillDate Me.txtDate: End Sub

  ' textbox (auto-formats a typed date on exit):
  Private Sub txtDate_AfterUpdate(): NormalizeDateBox Me.txtDate: End Sub
```
Or call the picker directly:
```
  Dim d As Variant
  d = PickDate()  ' opens on today; returns Null if cancelled
  If Not IsNull(d) Then MsgBox d
```
## Notes

Dates are written as `m/d/yyyy`. Requires Office 2010+ (VBA7).
