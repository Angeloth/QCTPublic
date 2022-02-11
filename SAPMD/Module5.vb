Module Module5
    Private alarmTime As Date
    Private WithEvents timer As Timer = New Timer()

    Public Sub Start(ByVal time As Date)
        alarmTime = time
        timer.Start()
    End Sub

    Public Sub TerminaTime()
        timer.Stop()
    End Sub
    Public Sub AgregaTiempo()
        alarmTime = Date.Now.AddMinutes(5)
    End Sub
    Public Sub TimerTick(ByVal sender As Object, ByVal e As EventArgs) Handles timer.Tick

        If alarmTime < Date.Now Then
            timer.Stop()
            RaiseEvent TimerComplete("Time's up.")
        Else
            Dim remainingTime As TimeSpan = alarmTime.Subtract(Date.Now)
            RaiseEvent Elapsed(remainingTime)
        End If

    End Sub

    'This will fire when the duration of the timer has elapsed
    Public Event TimerComplete As Action(Of String)

    'This will fire each time the timer ticks unless it's finished
    Public Event Elapsed As Action(Of TimeSpan)

End Module
