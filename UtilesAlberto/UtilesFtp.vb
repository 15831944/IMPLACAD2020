Imports System.Net.FtpWebRequest
Imports System.Net
Imports System.IO
Imports System.Windows.Forms

Partial Public Class Utiles
    Public Const sep As String = "||"
    Public Shared pasive As Boolean = True
    '
    Public Shared Function FTP_ExisteDir(queDir As String, user As String, pass As String) As String
        Dim resultado As String = ""
        Dim ftpRequ As FtpWebRequest = Nothing
        Dim ftpResp As FtpWebResponse = Nothing
        Try
            ftpRequ = DirectCast(FtpWebRequest.Create(queDir), FtpWebRequest)
            ftpRequ.Credentials = New NetworkCredential(user, pass)
            ftpRequ.AuthenticationLevel = Security.AuthenticationLevel.MutualAuthRequested
            ftpRequ.Method = WebRequestMethods.Ftp.ListDirectory
            'ftpRequ.Proxy = Nothing
            ftpRequ.UsePassive = pasive
            '
            ftpResp = DirectCast(ftpRequ.GetResponse, FtpWebResponse)
            resultado = ftpResp.StatusDescription
        Catch ex As WebException
            resultado = "ERROR: " & FTP_ResponseString(ex.Status.ToString & " | " & ex.Message & " | " & ex.Response.ResponseUri.ToString)
        Catch ex As Exception
            resultado = "ERROR: " & FTP_ResponseString(ex.ToString)
        Finally
            ftpRequ = Nothing
            If ftpResp IsNot Nothing Then
                ftpResp.Close()
            End If
            ftpResp = Nothing
        End Try
        '
        Return resultado
    End Function
    Public Shared Function FTP_EliminarFichero(ByVal fichero As String, user As String, pass As String) As String
        Dim resultado As String = ""
        Dim ftpRequ As FtpWebRequest = Nothing
        Dim ftpResp As FtpWebResponse = Nothing

        ' Creamos una petición FTP con la dirección del fichero a eliminar
        ftpRequ = DirectCast(WebRequest.Create(New Uri(fichero)), FtpWebRequest)
        ftpRequ.Credentials = New NetworkCredential(user, pass)
        ftpRequ.AuthenticationLevel = Security.AuthenticationLevel.MutualAuthRequested
        ftpRequ.Method = WebRequestMethods.Ftp.DeleteFile
        ftpRequ.UsePassive = pasive
        Try
            ftpResp = DirectCast(ftpRequ.GetResponse(), FtpWebResponse)
            resultado = FTP_ResponseString(ftpResp.ToString)
        Catch ex As WebException
            resultado = "ERROR: " & FTP_ResponseString(ex.Status.ToString & " | " & ex.Message & " | " & ex.Response.ResponseUri.ToString)
        Catch ex As Exception
            ' Si se produce algún fallo, se devolverá el mensaje del error
            Return "ERROR " & ex.ToString
            resultado = "ERROR: " & FTP_ResponseString(ex.ToString)
        Finally
            ftpRequ = Nothing
            If ftpResp IsNot Nothing Then ftpResp.Close()
            ftpResp = Nothing
        End Try
        '
        Return resultado
    End Function

    Public Shared Function FTP_ExisteObjeto(ByVal dir As String, user As String, pass As String) As Boolean
        Dim peticionFTP As FtpWebRequest

        ' Creamos una peticion FTP con la dirección del objeto que queremos saber si existe
        peticionFTP = CType(WebRequest.Create(New Uri(dir)), FtpWebRequest)

        ' Fijamos el usuario y la contraseña de la petición
        peticionFTP.Credentials = New NetworkCredential(user, pass)

        ' Para saber si el objeto existe, solicitamos la fecha de creación del mismo
        peticionFTP.Method = WebRequestMethods.Ftp.GetDateTimestamp

        peticionFTP.UsePassive = False

        Try
            ' Si el objeto existe, se devolverá True
            Dim respuestaFTP As FtpWebResponse
            respuestaFTP = CType(peticionFTP.GetResponse(), FtpWebResponse)
            Return True
        Catch ex As WebException
            Return False
        Catch ex As Exception
            ' Si el objeto no existe, se producirá un error y al entrar por el Catch
            ' se devolverá falso
            Return False
        End Try
    End Function

    Public Shared Function FTP_CrearDirectorio(ByVal dir As String, user As String, pass As String) As String
        Dim resultado As String = ""
        Dim ftpRequ As FtpWebRequest = Nothing
        Dim ftpResp As FtpWebResponse = Nothing

        ' Creamos una petición FTP con la dirección del fichero a eliminar
        ftpRequ = DirectCast(WebRequest.Create(New Uri(dir)), FtpWebRequest)
        ftpRequ.Credentials = New NetworkCredential(user, pass)
        ftpRequ.AuthenticationLevel = Security.AuthenticationLevel.MutualAuthRequested
        ftpRequ.Method = WebRequestMethods.Ftp.MakeDirectory
        ftpRequ.UsePassive = pasive
        Try
            ftpResp = DirectCast(ftpRequ.GetResponse(), FtpWebResponse)
            resultado = FTP_ResponseString(ftpResp.ToString)
        Catch ex As WebException
            resultado = "ERROR: " & FTP_ResponseString(ex.Status.ToString & " | " & ex.Message & " | " & ex.Response.ResponseUri.ToString)
        Catch ex As Exception
            resultado = "ERROR: " & FTP_ResponseString(ex.ToString)
        Finally
            ftpRequ = Nothing
            If ftpResp IsNot Nothing Then ftpResp.Close()
            ftpResp = Nothing
        End Try
        '
        Return resultado
    End Function

    Public Shared Function FTP_Upload(ByVal ficheroLocal As String, ByVal ficheroFtp As String, user As String, pass As String) As String
        Dim resultado As String = ""
        Dim ftpRequ As FtpWebRequest = Nothing
        Dim ftpResp As FtpWebResponse = Nothing

        ' Creamos una petición FTP con la dirección del fichero a eliminar
        ftpRequ = DirectCast(WebRequest.Create(New Uri(ficheroFtp)), FtpWebRequest)
        ftpRequ.Credentials = New NetworkCredential(user, pass)
        ftpRequ.AuthenticationLevel = Security.AuthenticationLevel.MutualAuthRequested
        ftpRequ.Method = WebRequestMethods.Ftp.UploadFile
        ftpRequ.UsePassive = pasive
        '
        ' Fichero local a subir, en Byte()
        Dim btfile() As Byte = IO.File.ReadAllBytes(ficheroLocal)
        Dim strFile As IO.Stream = Nothing
        Try
            strFile = ftpRequ.GetRequestStream()
            'Upload Each Byte'
            strFile.Write(btfile, 0, btfile.Length)
            resultado = FTP_ResponseString(ftpResp.ToString)
        Catch ex As WebException
            resultado = "ERROR: " & FTP_ResponseString(ex.Status.ToString & " | " & ex.Message & " | " & ex.Response.ResponseUri.ToString)
        Catch ex As Exception
            resultado = "ERROR: " & FTP_ResponseString(ex.ToString)
        Finally
            If strFile IsNot Nothing Then
                strFile.Flush()
                strFile.Close()
            End If
            strFile = Nothing
            ftpRequ = Nothing
            If ftpResp IsNot Nothing Then ftpResp.Close()
            ftpResp = Nothing
        End Try
        '
        Return resultado
    End Function
    '
    Public Shared Function FTP_Borra(ficheroFtp As String, user As String, pass As String) As String
        Dim resultado As String = ""
        Dim ftpRequ As FtpWebRequest = Nothing
        Dim ftpResp As FtpWebResponse = Nothing

        ' Creamos una petición FTP con la dirección del fichero a eliminar
        ftpRequ = DirectCast(WebRequest.Create(New Uri(ficheroFtp)), FtpWebRequest)
        ftpRequ.Credentials = New NetworkCredential(user, pass)
        ftpRequ.AuthenticationLevel = Security.AuthenticationLevel.MutualAuthRequested
        ftpRequ.Method = WebRequestMethods.Ftp.DeleteFile
        ftpRequ.UsePassive = pasive
        '
        Try

            ftpResp = DirectCast(ftpRequ.GetResponse(), FtpWebResponse)
            resultado = FTP_ResponseString(ftpResp.ToString)
        Catch ex As WebException
            resultado = "ERROR: " & FTP_ResponseString(ex.Status.ToString & " | " & ex.Message & " | " & ex.Response.ResponseUri.ToString)
        Catch ex As Exception
            resultado = "ERROR: " & FTP_ResponseString(ex.ToString)
        Finally
            ftpRequ = Nothing
            If ftpResp IsNot Nothing Then ftpResp.Close()
            ftpResp = Nothing
        End Try
        '
        Return resultado
    End Function

    Public Shared Function FTP_ListaDir(ByVal queDir As String, user As String, pass As String) As String
        Dim resultado As String = ""
        Dim ftpRequ As FtpWebRequest = Nothing
        Dim ftpResp As FtpWebResponse = Nothing
        Try
            ftpRequ = DirectCast(FtpWebRequest.Create(queDir), FtpWebRequest)
            ' Si no se necesitara autentificación con user/password, comentar esta linea
            ftpRequ.Credentials = New NetworkCredential(user, pass)
            ftpRequ.AuthenticationLevel = Security.AuthenticationLevel.MutualAuthRequested
            ftpRequ.Method = WebRequestMethods.Ftp.ListDirectoryDetails
            'ftpRequ.Proxy = Nothing
            ftpRequ.UsePassive = pasive

            ftpResp = DirectCast(ftpRequ.GetResponse, FtpWebResponse)
            Dim sreader As New IO.StreamReader(ftpResp.GetResponseStream)
            While Not sreader.Peek = -1
                Dim ftpList As String() = sreader.ReadLine.Split(" "c)
                'Dim ftpBites As String = FormatNumber(CDbl(ftpList(ftpList.GetUpperBound(0) - 4)) / 1000000, 2, TriState.True) & " Mb"
                Dim ftpBites As String = ftpList(ftpList.GetUpperBound(0) - 4)
                Dim ftpFecha As String = ftpList(ftpList.GetUpperBound(0) - 2) & "/" & ftpList(ftpList.GetUpperBound(0) - 3) & " " & ftpList(ftpList.GetUpperBound(0) - 1)
                Dim ftpfile As String = ftpList(ftpList.GetUpperBound(0))
                'Console.WriteLine(ftpfile)
                'If ftpfile.Contains(".bsp") And Not ftpfile.Contains(".ztmp") Then
                'lwMaps.Items.Add(ftpfile, 6)
                resultado &= ftpfile & vbTab & ftpBites & vbTab & ftpFecha & vbCrLf
                'End If
            End While
        Catch ex As WebException
            resultado = "ERROR: " & FTP_ResponseString(ex.Status.ToString & " | " & ex.Message & " | " & ex.Response.ResponseUri.ToString)
        Catch ex As Exception
            resultado = "ERROR: " & FTP_ResponseString(ex.ToString)
        Finally
            ftpRequ = Nothing
            If ftpResp IsNot Nothing Then
                ftpResp.Close()
            End If
            ftpResp = Nothing
        End Try

        Return resultado
        '-		ftpList	{Length=17}	String()
        '(0)	"-rw-r--r--"	String
        '(1)	"1"	String
        '(2)	"ftp"	String
        '(3)	"ftp"	String
        '(4)	""	String
        '(5)	""	String
        '(6)	""	String
        '(7)	""	String
        '(8)	""	String
        '(9)	""	String
        '(10)	""	String
        '(11)	""	String
        '(12)	"234001"	String
        '(13)	"Dec"	String
        '(14)	"05"	String
        '(15)	"16:04"	String
        '(16)	"CommandNames.txt"	String
    End Function

    'Public Shared Function existeObjeto1(ByVal dir As String, user As String, pass As String) As Boolean
    '    Dim peticionFTP As FtpWebRequest

    '    ' Creamos una peticion FTP con la dirección del objeto que queremos saber si existe
    '    peticionFTP = CType(WebRequest.Create(New Uri(dir)), FtpWebRequest)

    '    ' Fijamos el usuario y la contraseña de la petición
    '    peticionFTP.Credentials = New NetworkCredential(user, pass)

    '    ' Para saber si el objeto existe, solicitamos la fecha de creación del mismo
    '    peticionFTP.Method = WebRequestMethods.Ftp.GetDateTimestamp

    '    peticionFTP.UsePassive = False

    '    Try
    '        ' Si el objeto existe, se devolverá True
    '        Dim respuestaFTP As FtpWebResponse
    '        respuestaFTP = CType(peticionFTP.GetResponse(), FtpWebResponse)
    '        Return True
    '    Catch ex As Exception
    '        ' Si el objeto no existe, se producirá un error y al entrar por el Catch
    '        ' se devolverá falso
    '        Return False
    '    End Try
    'End Function

    Public Enum queDatoFichero
        tamaño = 0
        fecha = 1
        '-		ftpList	{Length=17}	String()
        '(0)	"-rw-r--r--"	String
        '(1)	"1"	String
        '(2)	"ftp"	String
        '(3)	"ftp"	String
        '(4)	""	String
        '(5)	""	String
        '(6)	""	String
        '(7)	""	String
        '(8)	""	String
        '(9)	""	String
        '(10)	""	String
        '(11)	""	String
        '(12)	"234001"	String
        '(13)	"Dec"	String
        '(14)	"05"	String
        '(15)	"16:04"	String
        '(16)	"CommandNames.txt"	String
    End Enum

    Public Shared Function FTP_DameDatosFichero(ByVal queFichero As String, ByVal queDato As queDatoFichero, user As String, pass As String) As String
        Dim resultado As String = ""
        Dim ftpRequ As FtpWebRequest = Nothing
        Dim ftpResp As WebResponse = Nothing
        ' Creamos una peticion FTP con la dirección del objeto que queremos saber si existe
        ftpRequ = CType(WebRequest.Create(New Uri(queFichero)), FtpWebRequest)
        ftpRequ.Credentials = New NetworkCredential(user, pass)
        ftpRequ.Method = Nothing

        '' PETICIÓN QUE ESTAMOS SOLICITANDO AL SERVIDOR
        If queDato = queDatoFichero.fecha Then
            ftpRequ.Method = WebRequestMethods.Ftp.GetDateTimestamp
        ElseIf queDato = queDatoFichero.tamaño Then
            ftpRequ.Method = WebRequestMethods.Ftp.GetFileSize
        End If
        ftpRequ.AuthenticationLevel = Security.AuthenticationLevel.MutualAuthRequested
        ftpRequ.UsePassive = pasive

        Try
            ftpResp = CType(ftpRequ.GetResponse(), FtpWebResponse)
            resultado = ftpResp.ToString
        Catch ex As WebException
            resultado = "ERROR: " & FTP_ResponseString(ex.Status.ToString & " | " & ex.Message & " | " & ex.Response.ResponseUri.ToString)
        Catch ex As Exception
            resultado = "ERROR: " & FTP_ResponseString(ex.ToString)
        End Try
        '
        Return resultado
    End Function
    '
    Private Shared Function FTP_ResponseString(queResu As String) As String
        Return queResu.Replace(":", "|").Replace(";", "|").Replace(vbCrLf, "|")
    End Function
End Class
