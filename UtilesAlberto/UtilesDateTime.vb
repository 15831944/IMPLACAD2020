Imports System.Threading

Partial Public Class Utiles
    Private Shared mTime As Stopwatch
    '
    Public Shared Sub Retardo(segundos As Integer)
        Threading.Thread.Sleep(segundos * 1000)
    End Sub

    Public Shared Sub mTime_Inicio()
        mTime = New Stopwatch
        mTime.Start()
    End Sub

    Public Shared Function mTime_Duration(Optional message As Boolean = False) As String
        If mTime Is Nothing Then
            Return ""
            Exit Function
        End If
        Dim resultado As String = ""
        mTime.Stop()
        ' Get the elapsed time as a TimeSpan value.
        Dim ts As TimeSpan = mTime.Elapsed

        ' Format and display the TimeSpan value.
        resultado = String.Format("{0:00}:{1:00}:{2:00}.{3:00}", ts.Hours, ts.Minutes, ts.Seconds, ts.Milliseconds / 10)
        If message = True Then
            MsgBox("Duration : " + resultado)
        End If
        'Console.WriteLine("Duración : " + resultado)
        'Debug.Print("Duración : " + resultado)
        mTime = Nothing
        Return resultado
    End Function
    Public Shared Sub EjemploTiempo()
        Dim stopWatch As New Stopwatch()
        stopWatch.Start()
        Thread.Sleep(10000)
        stopWatch.Stop()
        ' Get the elapsed time as a TimeSpan value.
        Dim ts As TimeSpan = stopWatch.Elapsed

        ' Format and display the TimeSpan value.
        Dim elapsedTime As String = String.Format("{0:00}:{1:00}:{2:00}.{3:00}", ts.Hours, ts.Minutes, ts.Seconds, ts.Milliseconds / 10)
        'Console.WriteLine("RunTime " + elapsedTime)
        MsgBox(elapsedTime)
    End Sub 'Main
End Class