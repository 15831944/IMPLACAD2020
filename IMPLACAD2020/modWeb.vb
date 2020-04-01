Imports System.Net

Module modWeb

    ''
    Public Sub DescargaFicheroWEB(origenURL As String, destinoHD As String, Optional quePb As System.Windows.Forms.ProgressBar = Nothing)
        ' Put your command code here
        'If MsgBox("Desea comprobar si existen actualizaciones?", MsgBoxStyle.YesNo) = MsgBoxResult.No Then
        'Exit Sub
        'End If
        ''
        ''
        Dim webRequest As HttpWebRequest = Net.WebRequest.Create(origenURL)
        webRequest.Method = "GET"
        webRequest.ContentType = "application/x-www-form-urlencoded"
        'webRequest.CookieContainer = cookies

        Dim bytesProcessed As Integer = 0
        Dim remoteStream As Stream
        Dim localStream As Stream
        Dim response As WebResponse

        response = webRequest.GetResponse()
        If Not response Is Nothing Then
            If quePb IsNot Nothing Then
                quePb.Minimum = 0
                quePb.Maximum = CInt(DameTamañoFicheroWEB(origenURL))
            End If
            remoteStream = response.GetResponseStream()
            localStream = File.Create(destinoHD)

            'Declare buffer as byte array
            Dim myBuffer As Byte()
            'Byte array initialization
            ReDim myBuffer(1024)

            Dim bytesRead As Integer
            bytesRead = remoteStream.Read(myBuffer, 0, 1024)
            Do While (bytesRead > 0)
                localStream.Write(myBuffer, 0, bytesRead)
                bytesProcessed += bytesRead
                If quePb IsNot Nothing Then
                    If bytesProcessed > quePb.Maximum Then
                        quePb.Value = quePb.Maximum
                    Else
                        quePb.Value = bytesProcessed
                    End If
                    quePb.Invalidate()
                End If
                bytesRead = remoteStream.Read(myBuffer, 0, 1024)
                System.Windows.Forms.Application.DoEvents()
            Loop
            remoteStream.Close()
            localStream.Close()
        End If
    End Sub

    ''
    Public Function DameTamañoFicheroWEB(origenURL As String) As Long
        Dim resultado As Long = 0
        ''
        If My.Computer.Network.IsAvailable = False Then
            MsgBox("No dispone de conexión a internet. No se puede actualizar.")
            Return resultado
            Exit Function
        End If
        ''
        Dim webRequest As HttpWebRequest = Net.WebRequest.Create(origenURL)
        webRequest.Method = "GET"
        webRequest.ContentType = "application/x-www-form-urlencoded"
        'webRequest.CookieContainer = cookies

        Dim response As WebResponse

        response = webRequest.GetResponse()
        If Not response Is Nothing Then
            resultado = response.ContentLength
            response.Close()
        End If
        Return resultado
        ''
    End Function

    ''
    Public Function DameFechaFicheroWEB(origenURL As String, Optional verdatos As Boolean = False) As Date
        Dim resultado As Date = New System.DateTime(0)
        ''
        If My.Computer.Network.IsAvailable = False Then
            MsgBox("No dispone de conexión a internet. No se puede actualizar.")
            Return resultado
            Exit Function
        End If
        ''
        Dim webRequest As HttpWebRequest = Nothing
        Try
            webRequest = Net.WebRequest.Create(origenURL)
        Catch ex As Exception
            '' No existe el fichero o ha dado error. salimos de la función
            Return resultado
            Exit Function
        End Try
        ''
        webRequest.Method = "GET"
        webRequest.ContentType = "application/x-www-form-urlencoded"
        'webRequest.CookieContainer = cookies

        Dim response As WebResponse

        response = webRequest.GetResponse()
        'resultado = WebRequestMethods.File.
        Dim mensaje As String = ""
        If Not response Is Nothing Then
            'resultado = WebRequestMethods.Ftp.GetDateTimestamp
            mensaje = resultado
            For x As Integer = 0 To response.Headers.Count - 1
                mensaje &= response.Headers.Keys(x) & " = " & response.Headers(x).ToString & vbCrLf
                If response.Headers.Keys(x).StartsWith("Last-") Then
                    If IsDate(response.Headers(x).ToString) Then
                        resultado = CDate(response.Headers(x).ToString).ToLocalTime
                    Else
                        resultado = Date.Now.ToLocalTime
                        'Exit For
                    End If
                End If
            Next
            response.Close()
        End If
        If verdatos = True Then
            MsgBox(mensaje)
            Debug.Print(mensaje)
        End If
        Return resultado
        ''
    End Function

    ''
    Public Function DameDatosFicheroWEB(origenURL As String) As String
        Dim resultado As String = ""
        ''
        If My.Computer.Network.IsAvailable = False Then
            MsgBox("No dispone de conexión a internet. No se puede actualizar.")
            Return resultado
            Exit Function
        End If
        ''
        Dim webRequest As HttpWebRequest = Net.WebRequest.Create(origenURL)
        webRequest.Method = "GET"
        webRequest.ContentType = "application/x-www-form-urlencoded"
        'webRequest.CookieContainer = cookies

        Dim response As WebResponse

        response = webRequest.GetResponse()
        If Not response Is Nothing Then
            Try
                For x As Integer = 0 To response.Headers.Count - 1
                    resultado &= response.Headers.Keys(x) & " = " & response.Headers(x) & vbCrLf
                Next
            Catch ex As Exception
                '' No hacemos nada y continuamos.
            End Try
            'resultado = response.GetResponseStream
            response.Close()
        End If
        Return resultado
        ''
    End Function

    ''
End Module