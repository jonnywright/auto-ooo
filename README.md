# auto-ooo
Out of Office Assistant

Includes Sub-routine (although not currently used) to auto-enable out of office based on day of week and time of day Outlook is closed;


Sub CheckDayOfWeekAndTime()

    'If it is a Thursday...
    If Weekday(Now(), vbThursday) = 1 Then

        'And after 13:00:00...
        If Time > TimeValue("13:00:00") Then

        'Turn on out-of-office
        OutOfOffice True

        Else
        End If
    Else
    End If

    End Sub
