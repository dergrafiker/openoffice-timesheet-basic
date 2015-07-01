REM  *****  BASIC  *****

Public Function minBreakTime(ByVal workTime As Double, ByVal breakTime As Double) as Double
    Dim minBreak As Double
    workTime = workTime * 24
    breakTime = breakTime * 24

    If workTime <= 6 Then
    	minBreak = 0
    ElseIf workTime <= 9 Then
        minBreak = 30# / 60#
    Else
        minBreak = 45# / 60#
    End If

    If workTime = 0 Then
        minBreakTime = 0
    ElseIf breakTime < minBreak Then
        minBreakTime = minBreak / 24
    Else
        minBreakTime = breakTime / 24
    End If
End Function

Public Function timeDiff(dateA As Date, dateB As Date) As String
    Dim timeDifference As String
    timeDifference = Format(Abs(dateA - dateB), "[hh]:mm")

    If (dateA = 0) Then
        timeDiff = Format(0, "[hh]:mm")
    ElseIf (timeDifference = "0:00") Then
        timeDiff = timeDifference
    ElseIf ((dateB * 24) > (dateA * 24)) Then
        timeDiff = "-" + timeDifference
    Else
        timeDiff = timeDifference
    End If
End Function

Public Function roundEven(X As Double, Optional Anzahl_Stellen As Long)
    If( IsMissing(Anzahl_Stellen) ) Then
        Anzahl_Stellen = 0
    End If
    roundEven = Round(X, Anzahl_Stellen)
End Function

Public Function vacationCount(text As String) As Double
    vacationCount = 0#
    If (text Like "UT*") Then
        vacationCount = 1#
    ElseIf (text Like "HUT*") Then
        vacationCount = 0.5
    End If
End Function

Public Function sickDay(text As String) As Integer
    sickDay = 0
    If (text Like "KRANK*") Then
        sickDay = 1
    End If
End Function

Public Function isHoliday(text As String) As Boolean
     isHoliday = text Like "FREI*"
End Function

Public Function expectedWorktime(actualWorkTime As Double, comment As String, defaultWorkTime As Double) As Double
    expectedWorktime = 0
    
    If (actualWorkTime > 0) Then
        If (vacationCount(comment) = 0 And sickDay(comment) = 0 And Not isHoliday(comment)) Then
            expectedWorktime = defaultWorkTime
        End If
    End If
End Function

Public Function calcWorkTime(maxWorkTime as Double,worktime as Double,manualTime as Double) as Double
	Dim breakTime As Double
	IF (manualTime > 0) Then
		calcWorkTime = manualTime
	Else
		If(worktime > 0) Then
    		breakTime = maxWorkTime - workTime
	    End If
    	breakTime = CDate(minBreakTime(maxWorkTime, breakTime))
    
   		'MsgBox "maxWorkTime "&Format(maxWorkTime, "[h]:mm") & chr(13) & "worktime "&Format(worktime, "[h]:mm") & chr(13) & "breakTime "&Format(breakTime, "[h]:mm")
	    calcWorkTime = maxWorkTime - breakTime
    End If
End Function
 
Public Function calc6(t1 as Double,t2 as Double,t3 as Double,t4 as Double,t5 as Double,t6 as Double, manualTime as Double) as Double
	Dim workTime, maxWorkTime As Double
	If (manualTime = 0) Then
		workTime = ABS(t1-t2) + ABS(t3-t4) + ABS(t5-t6)
    	maxWorkTime = ABS(t1-t6)
	End If
    calc6 = calcWorkTime(maxWorkTime,workTime,manualTime)
End Function

Public Function calc4(t1 as Double,t2 as Double,t3 as Double,t4 as Double, manualTime as Double) as Double
	Dim workTime, maxWorkTime As Double
	
	If (manualTime = 0) Then
		workTime = ABS(t1-t2) + ABS(t3-t4)
    	maxWorkTime = ABS(t1-t4)
    End If
   	calc4 = calcWorkTime(maxWorkTime,workTime,manualTime)
End Function

