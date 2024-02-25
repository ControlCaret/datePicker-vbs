Option Explicit

Dim offsetRow As Integer
Dim offsetCol As Integer
Dim outputCell As Range

Sub printCalendar()
    Dim displayDate As Date
    Dim monthStart As Date
    Dim monthEnd As Date
    Dim currentRow As Integer
    Dim currentCol As Integer
    Dim i
    Dim j

    ' Set the offset for the calendar
    offsetRow = 0
    offsetCol = 3

    ' Set the output cell for the date picker
    Set outputCell = Range("A1")

    ' Clear the calendar
    Range("A2:G8").Offset(offsetRow, offsetCol).ClearContents

    ' Print ◀ and ▶
    Range("A1").Offset(offsetRow, offsetCol).Value = "◀"
    Range("G1").Offset(offsetRow, offsetCol).Value = "▶"
    
    ' Get the date to display
    displayDate = DateSerial(Range("C1").Offset(offsetRow, offsetCol).Value, Range("E1").Offset(offsetRow, offsetCol).Value, 1)

    ' Get the start and end of the month
    monthStart = DateSerial(Year(displayDate), Month(displayDate), 1)
    monthEnd = DateSerial(Year(displayDate), Month(displayDate) + 1, 0)

    ' Print the year and month
    'Range("C1").Value = Format(displayDate, "YYYY")
    'Range("E1").Value = Format(displayDate, "M")
    ' Conflict with the change event

    ' Print the days of the week
    For i = 0 To 6
        Cells(2, i + 1).Offset(offsetRow, offsetCol).Value = Format(DateSerial(1, 1, i), "ddd")
    Next i

    ' Print the days of the month
    currentRow = 3
    For i = monthStart To monthEnd
        currentCol = Weekday(i, vbSunday)
        Cells(currentRow + offsetRow, currentCol + offsetCol).Value = Day(i)
        If currentCol = 7 Then
            currentRow = currentRow + 1
        End If
    Next i
End Sub

Private Sub nextMonth()
    Dim y As Integer
    Dim m As Integer
    y = Range("C1").Offset(offsetRow, offsetCol).Value
    m = Range("E1").Offset(offsetRow, offsetCol).Value
    If m = 12 Then
        m = 1
        y = y + 1
    Else
        m = m + 1
    End If
    Range("C1").Offset(offsetRow, offsetCol).Value = y
    Range("E1").Offset(offsetRow, offsetCol).Value = m
    Range("A1").Select
    Call printCalendar()
End Sub

Private Sub previousMonth()
    Dim y As Integer
    Dim m As Integer
    y = Range("C1").Offset(offsetRow, offsetCol).Value
    m = Range("E1").Offset(offsetRow, offsetCol).Value
    If m = 1 Then
        m = 12
        y = y - 1
    Else
        m = m - 1
    End If
    Range("C1").Offset(offsetRow, offsetCol).Value = y
    Range("E1").Offset(offsetRow, offsetCol).Value = m
    Range("A1").Select
    Call printCalendar()
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    ' Change month when the user clicks ◀ or ▶
    If Selection.Count = 1 Then
        If Not Intersect(Target, Range("A1").Offset(offsetRow, offsetCol)) Is Nothing Then
            Call previousMonth()
        Elseif Not Intersect(Target, Range("G1").Offset(offsetRow, offsetCol)) Is Nothing Then
            Call nextMonth()
        Elseif Not Intersect(Target, Range("A3:G8").Offset(offsetRow, offsetCol)) Is Nothing Then
            If Not IsEmpty(Target.Value) Then
                Range(outputCell).Value = DateSerial(Range("C1").Offset(offsetRow, offsetCol).Value, Range("E1").Offset(offsetRow, offsetCol).Value, Target.Value)
            End If
        End If
    End If
End Sub

Private Sub WorkSheet_Change(ByVal Target As Range)
    If Not Intersect(Target, Range("C1:E1").Offset(offsetRow, offsetCol)) Is Nothing Then
        Call printCalendar()
    End If
End Sub
