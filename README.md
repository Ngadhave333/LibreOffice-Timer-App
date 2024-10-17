' Global Arrays to hold the state for 14 people, each with 12 planets
Dim timers(13, 11) As Date ' 14 persons, 12 planets (0-13 for people, 0-11 for planets)
Dim isRunning(13, 11) As Boolean
Dim timerValues(13, 11) As Long

Sub StartTimer(personIndex As Integer, planetIndex As Integer)
    Dim oSheet As Object
    oSheet = ThisComponent.Sheets(0)

    If isRunning(personIndex, planetIndex) Then
        MsgBox "The timer for Person " & (personIndex + 1) & ", Planet " & (planetIndex + 1) & " is already running."
        Exit Sub
    End If

    ' Initialize timer settings
    isRunning(personIndex, planetIndex) = True
    timers(personIndex, planetIndex) = Now
    timerValues(personIndex, planetIndex) = 14400 ' 4 hours in seconds

    ' Start the timer update routine
    TimerUpdateRoutine
End Sub

Sub TimerUpdateRoutine()
    Dim oSheet As Object
    oSheet = ThisComponent.Sheets(0)
    Dim i, j As Integer

    For i = 0 To 13 ' 14 persons
        For j = 0 To 11 ' 12 planets per person
            If isRunning(i, j) Then
                ' Calculate remaining time
                Dim elapsedTime As Long
                elapsedTime = DateDiff("s", timers(i, j), Now)
                timerValues(i, j) = 14400 - elapsedTime

                If timerValues(i, j) <= 0 Then
                    oSheet.getCellByPosition(1 + j, 2 + (i * 3)).String = "Attack Ready!" ' Adjust row/col accordingly
                    isRunning(i, j) = False
                Else
                    ' Update the cell with remaining time in HH:MM:SS format
                    Dim hoursLeft As Long
                    Dim minutesLeft As Long
                    Dim secondsLeft As Long

                    hoursLeft = Int(timerValues(i, j) / 3600)
                    minutesLeft = Int((timerValues(i, j) Mod 3600) / 60)
                    secondsLeft = timerValues(i, j) Mod 60

                    oSheet.getCellByPosition(1 + j, 2 + (i * 3)).String = _
                        Format(hoursLeft, "00") & ":" & Format(minutesLeft, "00") & ":" & Format(secondsLeft, "00")
                End If
            End If
        Next j
    Next i

    ' Wait for a second and update again
    Wait(1000)
    TimerUpdateRoutine
End Sub

' Person 1 (xyz1)
Sub StartTimerPerson1Planet1()
    StartTimer(0, 0) ' Person 1, Planet 1 (Main)
End Sub
Sub StartTimerPerson1Planet2()
    StartTimer(0, 1) ' Person 1, Planet 2 (c1)
End Sub
Sub StartTimerPerson1Planet3()
    StartTimer(0, 2) ' Person 1, Planet 3 (c2)
End Sub
Sub StartTimerPerson1Planet4()
    StartTimer(0, 3) ' Person 1, Planet 4 (c3)
End Sub
Sub StartTimerPerson1Planet5()
    StartTimer(0, 4) ' Person 1, Planet 5 (c4)
End Sub
Sub StartTimerPerson1Planet6()
    StartTimer(0, 5) ' Person 1, Planet 6 (c5)
End Sub
Sub StartTimerPerson1Planet7()
    StartTimer(0, 6) ' Person 1, Planet 7 (c6)
End Sub
Sub StartTimerPerson1Planet8()
    StartTimer(0, 7) ' Person 1, Planet 8 (c7)
End Sub
Sub StartTimerPerson1Planet9()
    StartTimer(0, 8) ' Person 1, Planet 9 (c8)
End Sub
Sub StartTimerPerson1Planet10()
    StartTimer(0, 9) ' Person 1, Planet 10 (c9)
End Sub
Sub StartTimerPerson1Planet11()
    StartTimer(0, 10) ' Person 1, Planet 11 (c10)
End Sub
Sub StartTimerPerson1Planet12()
    StartTimer(0, 11) ' Person 1, Planet 12 (c11)
End Sub

' Person 2 (xyz2)
Sub StartTimerPerson2Planet1()
    StartTimer(1, 0) ' Person 2, Planet 1 (Main)
End Sub
Sub StartTimerPerson2Planet2()
    StartTimer(1, 1) ' Person 2, Planet 2 (c1)
End Sub
Sub StartTimerPerson2Planet3()
    StartTimer(1, 2) ' Person 2, Planet 3 (c2)
End Sub
Sub StartTimerPerson2Planet4()
    StartTimer(1, 3) ' Person 2, Planet 4 (c3)
End Sub
Sub StartTimerPerson2Planet5()
    StartTimer(1, 4) ' Person 2, Planet 5 (c4)
End Sub
Sub StartTimerPerson2Planet6()
    StartTimer(1, 5) ' Person 2, Planet 6 (c5)
End Sub
Sub StartTimerPerson2Planet7()
    StartTimer(1, 6) ' Person 2, Planet 7 (c6)
End Sub
Sub StartTimerPerson2Planet8()
    StartTimer(1, 7) ' Person 2, Planet 8 (c7)
End Sub
Sub StartTimerPerson2Planet9()
    StartTimer(1, 8) ' Person 2, Planet 9 (c8)
End Sub
Sub StartTimerPerson2Planet10()
    StartTimer(1, 9) ' Person 2, Planet 10 (c9)
End Sub
Sub StartTimerPerson2Planet11()
    StartTimer(1, 10) ' Person 2, Planet 11 (c10)
End Sub
Sub StartTimerPerson2Planet12()
    StartTimer(1, 11) ' Person 2, Planet 12 (c11)
End Sub

' Person 3 (xyz3)
Sub StartTimerPerson3Planet1()
    StartTimer(2, 0) ' Person 3, Planet 1 (Main)
End Sub
Sub StartTimerPerson3Planet2()
    StartTimer(2, 1) ' Person 3, Planet 2 (c1)
End Sub
Sub StartTimerPerson3Planet3()
    StartTimer(2, 2) ' Person 3, Planet 3 (c2)
End Sub
Sub StartTimerPerson3Planet4()
    StartTimer(2, 3) ' Person 3, Planet 4 (c3)
End Sub
Sub StartTimerPerson3Planet5()
    StartTimer(2, 4) ' Person 3, Planet 5 (c4)
End Sub
Sub StartTimerPerson3Planet6()
    StartTimer(2, 5) ' Person 3, Planet 6 (c5)
End Sub
Sub StartTimerPerson3Planet7()
    StartTimer(2, 6) ' Person 3, Planet 7 (c6)
End Sub
Sub StartTimerPerson3Planet8()
    StartTimer(2, 7) ' Person 3, Planet 8 (c7)
End Sub
Sub StartTimerPerson3Planet9()
    StartTimer(2, 8) ' Person 3, Planet 9 (c8)
End Sub
Sub StartTimerPerson3Planet10()
    StartTimer(2, 9) ' Person 3, Planet 10 (c9)
End Sub
Sub StartTimerPerson3Planet11()
    StartTimer(2, 10) ' Person 3, Planet 11 (c10)
End Sub
Sub StartTimerPerson3Planet12()
    StartTimer(2, 11) ' Person 3, Planet 12 (c11)
End Sub

' Person 4 (xyz4)
Sub StartTimerPerson4Planet1()
    StartTimer(3, 0) ' Person 4, Planet 1 (Main)
End Sub
Sub StartTimerPerson4Planet2()
    StartTimer(3, 1) ' Person 4, Planet 2 (c1)
End Sub
Sub StartTimerPerson4Planet3()
    StartTimer(3, 2) ' Person 4, Planet 3 (c2)
End Sub
Sub StartTimerPerson4Planet4()
    StartTimer(3, 3) ' Person 4, Planet 4 (c3)
End Sub
Sub StartTimerPerson4Planet5()
    StartTimer(3, 4) ' Person 4, Planet 5 (c4)
End Sub
Sub StartTimerPerson4Planet6()
    StartTimer(3, 5) ' Person 4, Planet 6 (c5)
End Sub
Sub StartTimerPerson4Planet7()
    StartTimer(3, 6) ' Person 4, Planet 7 (c6)
End Sub
Sub StartTimerPerson4Planet8()
    StartTimer(3, 7) ' Person 4, Planet 8 (c7)
End Sub
Sub StartTimerPerson4Planet9()
    StartTimer(3, 8) ' Person 4, Planet 9 (c8)
End Sub
Sub StartTimerPerson4Planet10()
    StartTimer(3, 9) ' Person 4, Planet 10 (c9)
End Sub
Sub StartTimerPerson4Planet11()
    StartTimer(3, 10) ' Person 4, Planet 11 (c10)
End Sub
Sub StartTimerPerson4Planet12()
    StartTimer(3, 11) ' Person 4, Planet 12 (c11)
End Sub

' Person 5 (xyz5)
Sub StartTimerPerson5Planet1()
    StartTimer(4, 0) ' Person 5, Planet 1 (Main)
End Sub
Sub StartTimerPerson5Planet2()
    StartTimer(4, 1) ' Person 5, Planet 2 (c1)
End Sub
Sub StartTimerPerson5Planet3()
    StartTimer(4, 2) ' Person 5, Planet 3 (c2)
End Sub
Sub StartTimerPerson5Planet4()
    StartTimer(4, 3) ' Person 5, Planet 4 (c3)
End Sub
Sub StartTimerPerson5Planet5()
    StartTimer(4, 4) ' Person 5, Planet 5 (c4)
End Sub
Sub StartTimerPerson5Planet6()
    StartTimer(4, 5) ' Person 5, Planet 6 (c5)
End Sub
Sub StartTimerPerson5Planet7()
    StartTimer(4, 6) ' Person 5, Planet 7 (c6)
End Sub
Sub StartTimerPerson5Planet8()
    StartTimer(4, 7) ' Person 5, Planet 8 (c7)
End Sub
Sub StartTimerPerson5Planet9()
    StartTimer(4, 8) ' Person 5, Planet 9 (c8)
End Sub
Sub StartTimerPerson5Planet10()
    StartTimer(4, 9) ' Person 5, Planet 10 (c9)
End Sub
Sub StartTimerPerson5Planet11()
    StartTimer(4, 10) ' Person 5, Planet 11 (c10)
End Sub
Sub StartTimerPerson5Planet12()
    StartTimer(4, 11) ' Person 5, Planet 12 (c11)
End Sub

' Helper Wait Function
Sub Wait(ByVal milliseconds As Long)
    Dim startTime As Long
    startTime = GetCurrentTime
    Do While GetCurrentTime < startTime + milliseconds
        DoEvents ' Allow other processes to run
    Loop
End Sub

' Helper Function to Get Current Time
Function GetCurrentTime() As Long
    GetCurrentTime = Timer * 1000 ' Return current time in milliseconds
End Function
