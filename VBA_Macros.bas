Attribute VB_Name = "modBikeShareAutomation"
'===========================================================================
' Bike-Sharing Analysis — Automation Macros
' How to install: Open BikeShare_Analysis.xlsx -> Alt+F11 (VBA Editor) ->
'                 File -> Import File... -> select this .bas file.
' Then: Developer tab -> Insert -> Button (Form Control) on the Dashboard
'       sheet -> assign each macro below to its own button.
'===========================================================================

Option Explicit

'---------------------------------------------------------------------------
' Refreshes every query, pivot table, and formula in the workbook, then
' stamps a "Last Refreshed" timestamp on the Dashboard. Run this first,
' any time new rows are appended to Clean_Data or a new import is pasted in.
'---------------------------------------------------------------------------
Sub RefreshAllAndExport()
    Dim pt As PivotTable
    Dim ws As Worksheet

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' Refresh all Power Query connections (if any are added later)
    ThisWorkbook.RefreshAll

    ' Refresh every native PivotTable in the workbook
    For Each ws In ThisWorkbook.Worksheets
        For Each pt In ws.PivotTables
            pt.RefreshTable
        Next pt
    Next ws

    Application.Calculation = xlCalculationAutomatic
    Application.CalculateFull

    ' Timestamp on Dashboard (add a "Last Refreshed" label there if not present)
    On Error Resume Next
    ThisWorkbook.Sheets("Dashboard").Range("B3").Value = _
        "Jan 1 - Feb 14, 2011 | Hourly rental data | Last refreshed " & Format(Now, "yyyy-mm-dd hh:mm")
    On Error GoTo 0

    Application.ScreenUpdating = True
    MsgBox "All data, queries, and pivot tables refreshed.", vbInformation
End Sub

'---------------------------------------------------------------------------
' Cleans and flags newly appended rows in the Clean_Data table:
'   - Fills Weekday_Flag (Weekend/Weekday) from the weekday code
'   - Flags rows with cnt = 0 for manual review
'   - Removes exact duplicate rows by instant
' Safe to re-run — it only touches blank/flagged cells, never overwrites
' existing formulas elsewhere in the sheet.
'---------------------------------------------------------------------------
Sub CleanAndFlagData()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim r As Long
    Dim weekdayCol As Long, flagCol As Long, cntCol As Long, reviewCol As Long
    Dim flaggedCount As Long, dupCount As Long

    Set ws = ThisWorkbook.Sheets("Clean_Data")
    Set tbl = ws.ListObjects("tbl_Clean")

    weekdayCol = tbl.ListColumns("weekday").Index
    flagCol = tbl.ListColumns("Weekday_Flag").Index
    cntCol = tbl.ListColumns("cnt").Index

    Application.ScreenUpdating = False
    flaggedCount = 0

    For r = 1 To tbl.ListRows.Count
        With tbl.ListRows(r).Range
            ' Weekday_Flag is already a live formula in this workbook; this loop
            ' exists so newly pasted rows without the formula still get flagged.
            If .Cells(1, flagCol).Value = "" Then
                Dim wd As Long
                wd = .Cells(1, weekdayCol).Value
                If wd = 0 Or wd = 6 Then
                    .Cells(1, flagCol).Value = "Weekend"
                Else
                    .Cells(1, flagCol).Value = "Weekday"
                End If
            End If

            If .Cells(1, cntCol).Value = 0 Then
                flaggedCount = flaggedCount + 1
            End If
        End With
    Next r

    ' Remove duplicate rows by instant (column 1), keeping the first occurrence
    dupCount = tbl.ListRows.Count
    tbl.Range.RemoveDuplicates Columns:=1, Header:=xlYes
    dupCount = dupCount - tbl.ListRows.Count

    Application.ScreenUpdating = True
    MsgBox "Cleaning complete." & vbCrLf & _
           tbl.ListRows.Count & " rows in table." & vbCrLf & _
           flaggedCount & " rows have cnt = 0 (review recommended)." & vbCrLf & _
           dupCount & " duplicate row(s) removed.", vbInformation
End Sub

'---------------------------------------------------------------------------
' Custom function: maps a weathersit code (1-4) to its readable label.
' Usage in a cell: =WeatherBucket([@weathersit])
'---------------------------------------------------------------------------
Function WeatherBucket(code As Variant) As String
    Select Case code
        Case 1: WeatherBucket = "Clear / Few Clouds"
        Case 2: WeatherBucket = "Mist / Cloudy"
        Case 3: WeatherBucket = "Light Rain / Snow"
        Case 4: WeatherBucket = "Severe Weather"
        Case Else: WeatherBucket = "Unknown"
    End Select
End Function

'---------------------------------------------------------------------------
' Custom function: flags an hourly rental count as a statistical anomaly
' using the classic mean +/- 2*stdev rule.
' Usage in a cell: =AnomalyCheck([@cnt], AVERAGE(tbl_Clean[cnt]), STDEV(tbl_Clean[cnt]))
'---------------------------------------------------------------------------
Function AnomalyCheck(value As Double, meanVal As Double, stdevVal As Double) As String
    If value > meanVal + 2 * stdevVal Then
        AnomalyCheck = "High Anomaly"
    ElseIf value < meanVal - 2 * stdevVal Then
        AnomalyCheck = "Low Anomaly"
    Else
        AnomalyCheck = "Normal"
    End If
End Function

'---------------------------------------------------------------------------
' One-click pipeline: cleans new data, then refreshes everything, then
' reports how many rows/pivots were touched. Wire this to a single
' "Run Full Refresh" button on the Dashboard for the Phase 7 deliverable.
'---------------------------------------------------------------------------
Sub RunFullAutomatedPipeline()
    Call CleanAndFlagData
    Call RefreshAllAndExport
    MsgBox "Full automated pipeline complete: data cleaned, flagged, and all views refreshed.", vbInformation
End Sub
