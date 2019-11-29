Imports System.Net.FtpWebRequest
Imports System.Net
Imports System.IO

Public Class clsFTP
    Public Const sep As String = "||"
    Public host, user, pass As String

    Public Sub New(ByVal host As String, ByVal user As String, ByVal pass As String)
        Me.host = host
        Me.user = user
        Me.pass = pass
    End Sub

    Public Function eliminarFichero(ByVal fichero As String) As String
        Dim peticionFTP As FtpWebRequest

        ' Creamos una petición FTP con la dirección del fichero a eliminar
        peticionFTP = CType(WebRequest.Create(New Uri(fichero)), FtpWebRequest)

        ' Fijamos el usuario y la contraseña de la petición
        peticionFTP.Credentials = New NetworkCredential(user, pass)

        ' Seleccionamos el comando que vamos a utilizar: Eliminar un fichero
        peticionFTP.Method = WebRequestMethods.Ftp.DeleteFile
        peticionFTP.UsePassive = False

        Try
            Dim respuestaFTP As FtpWebResponse
            respuestaFTP = CType(peticionFTP.GetResponse(), FtpWebResponse)
            respuestaFTP.Close()
            ' Si todo ha ido bien, devolvemos String.Empty
            Return String.Empty
        Catch ex As Exception
            ' Si se produce algún fallo, se devolverá el mensaje del error
            Return ex.Message
        End Try
    End Function

    Public Function existeObjeto(ByVal dir As String) As Boolean
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
        Catch ex As Exception
            ' Si el objeto no existe, se producirá un error y al entrar por el Catch
            ' se devolverá falso
            Return False
        End Try
    End Function

    Public Function creaDirectorio(ByVal dir As String) As String
        Dim peticionFTP As FtpWebRequest

        ' Creamos una peticion FTP con la dirección del directorio que queremos crear
        peticionFTP = CType(WebRequest.Create(New Uri(dir)), FtpWebRequest)

        ' Fijamos el usuario y la contraseña de la petición
        peticionFTP.Credentials = New NetworkCredential(user, pass)

        ' Seleccionamos el comando que vamos a utilizar: Crear un directorio
        peticionFTP.Method = WebRequestMethods.Ftp.MakeDirectory

        Try
            Dim respuesta As FtpWebResponse
            respuesta = CType(peticionFTP.GetResponse(), FtpWebResponse)
            respuesta.Close()
            ' Si todo ha ido bien, se devolverá String.Empty
            Return String.Empty
        Catch ex As Exception
            ' Si se produce algún fallo, se devolverá el mensaje del error
            Return ex.Message
        End Try
    End Function

    Public Function subirFichero(ByVal fichero As String, ByVal destino As String, _
    ByVal dir As String) As String
        Dim infoFichero As New FileInfo(fichero)

        Dim uri As String
        uri = destino

        ' Si no existe el directorio, lo creamos
        If Not existeObjeto(dir) Then
            creaDirectorio(dir)
        End If

        Dim peticionFTP As FtpWebRequest

        ' Creamos una peticion FTP con la dirección del fichero que vamos a subir
        peticionFTP = CType(FtpWebRequest.Create(New Uri(destino)), FtpWebRequest)

        ' Fijamos el usuario y la contraseña de la petición
        peticionFTP.Credentials = New NetworkCredential(user, pass)

        peticionFTP.KeepAlive = False
        peticionFTP.UsePassive = False

        ' Seleccionamos el comando que vamos a utilizar: Subir un fichero
        peticionFTP.Method = WebRequestMethods.Ftp.UploadFile

        ' Especificamos el tipo de transferencia de datos
        peticionFTP.UseBinary = True

        ' Informamos al servidor sobre el tamaño del fichero que vamos a subir
        peticionFTP.ContentLength = infoFichero.Length

        ' Fijamos un buffer de 2KB
        Dim longitudBuffer As Integer
        longitudBuffer = 2048
        Dim lector As Byte() = New Byte(2048) {}

        Dim num As Integer

        ' Abrimos el fichero para subirlo
        Dim fs As FileStream
        fs = infoFichero.OpenRead()

        Try
            Dim escritor As Stream
            escritor = peticionFTP.GetRequestStream()

            ' Leemos 2 KB del fichero en cada iteración
            num = fs.Read(lector, 0, longitudBuffer)

            While (num <> 0)
                ' Escribimos el contenido del flujo de lectura en el
                ' flujo de escritura del comando FTP
                escritor.Write(lector, 0, num)
                num = fs.Read(lector, 0, longitudBuffer)
            End While

            escritor.Close()
            fs.Close()
            ' Si todo ha ido bien, se devolverá String.Empty
            Return String.Empty
        Catch ex As Exception
            ' Si se produce algún fallo, se devolverá el mensaje del error
            Return ex.Message
        End Try
    End Function

    Public Sub DescargaFicheroSimple(ByVal oriFTP As String, ByVal desPC As String)
        Try
            'My.Computer.Network.DownloadFile(origen, destino, usuario, clave, True, 100, True)
            Dim pp As New Microsoft.VisualBasic.Devices.Network
            If pp.IsAvailable Then _
                pp.DownloadFile(oriFTP, desPC, Me.user, Me.pass, True, 100000, True)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub


    Public Function DescargarFicheroFTP(ByVal DirFtp As String, ByVal user As String, _
            ByVal pass As String, ByVal dirLocal As String, ByVal Fichero As String, _
            Optional ByRef quePb As System.Windows.Forms.ProgressBar = Nothing) As String

        Dim localFile As String = Fichero
        Dim remoteFile As String = Fichero
        Dim host As String = DirFtp 'nombre de la carpeta en nuestro server FTP donde estan los archivos que deseamos descargar
        ' colocamos el nombre de usuario y password respectivo para acceder al server, si este no poseyera, dejar solo las comillas, osea ""
        Dim username As String = user
        Dim password As String = pass
        Dim URI As String = host & remoteFile ' nombre completo de la ruta del archivo  & remoteFile
        Dim ftp As System.Net.FtpWebRequest = CType(Net.FtpWebRequest.Create(URI), Net.FtpWebRequest)
        ftp.Credentials = New System.Net.NetworkCredential(username, password)
        ftp.KeepAlive = False
        ftp.UseBinary = True
        ftp.Method = System.Net.WebRequestMethods.Ftp.DownloadFile
        Try
            Using response As System.Net.FtpWebResponse = CType(ftp.GetResponse, System.Net.FtpWebResponse)
                Using responseStream As IO.Stream = response.GetResponseStream
                    Using fs As New IO.FileStream(dirLocal & "\" & Fichero, IO.FileMode.Create)
                        Dim buffer(2047) As Byte
                        Dim read As Integer = 0
                        Do
                            read = responseStream.Read(buffer, 0, buffer.Length)
                            fs.Write(buffer, 0, read)
                            ''
                            If quePb IsNot Nothing AndAlso quePb.Value + read <= quePb.Maximum Then
                                quePb.Value += read
                                quePb.Refresh()
                            End If
                        Loop Until read = 0
                        responseStream.Close()
                        fs.Flush()
                        fs.Close()
                        fs.Dispose()
                    End Using
                    responseStream.Close()
                    responseStream.Dispose()
                End Using
                response.Close()
            End Using
        Catch ex As Exception
            Return ex.Message
            Exit Function
        End Try
        Return "Salgo de Bajar Fichero Ftp"
    End Function

    Public Sub CargarFicheroSimple(ByVal oriPC As String, ByVal desFTP As String)
        'Dim oriPC As String = "C:\Temp\CommandNames.txt"
        'Dim desFTP As String = "ftp://80.34.88.43/2aCAD/2aCADAdvancedTools/Inventor2011/CommandNames.txt"
        Try
            'My.Computer.Network.DownloadFile(origen, destino, usuario, clave, True, 100, True)
            Dim pp As New Microsoft.VisualBasic.Devices.Network
            If pp.IsAvailable Then _
                pp.UploadFile(oriPC, desFTP, Me.user, Me.pass, True, 100000, FileIO.UICancelOption.DoNothing)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Public Function ListarFicherosDir(ByVal queDir As String) As String()
        Dim mensaje(-1) As String
        Dim ftp As FtpWebRequest = DirectCast(FtpWebRequest.Create(Me.host & queDir), FtpWebRequest)
        '/// ************************

        '///NOTE if you need to authenticate with username / password you would add the following 2 lines ...
        Dim cred As New NetworkCredential(Me.user, Me.pass)
        ftp.Credentials = cred
        '/// ************************

        ftp.AuthenticationLevel = Security.AuthenticationLevel.MutualAuthRequested
        ftp.Method = WebRequestMethods.Ftp.ListDirectoryDetails
        ftp.Proxy = Nothing

        Dim ftpresp As FtpWebResponse = DirectCast(ftp.GetResponse, FtpWebResponse)

        '/// **** these 3 lines are NOT needed, just shows you a bit of the server logon messages you would see ...
        Console.WriteLine(ftp.Headers.ToString)
        Console.WriteLine(ftpresp.BannerMessage)
        Console.WriteLine(ftpresp.Headers.ToString)
        '/// ****

        Dim sreader As New IO.StreamReader(ftpresp.GetResponseStream)

        While Not sreader.Peek = -1
            Dim ftpList As String() = sreader.ReadLine.Split(" ")
            'Dim ftpBites As String = FormatNumber(CDbl(ftpList(ftpList.GetUpperBound(0) - 4)) / 1000000, 2, TriState.True) & " Mb"
            Dim ftpBites As String = ftpList(ftpList.GetUpperBound(0) - 4)
            Dim ftpFecha As String = ftpList(ftpList.GetUpperBound(0) - 2) & "/" & ftpList(ftpList.GetUpperBound(0) - 3) & " " & ftpList(ftpList.GetUpperBound(0) - 1)
            Dim ftpfile As String = ftpList(ftpList.GetUpperBound(0))
            Console.WriteLine(ftpfile)
            'If ftpfile.Contains(".bsp") And Not ftpfile.Contains(".ztmp") Then
            'lwMaps.Items.Add(ftpfile, 6)
            ReDim Preserve mensaje(mensaje.GetUpperBound(0) + 1)
            mensaje(mensaje.GetUpperBound(0)) = ftpfile '& vbTab & ftpBites & vbTab & ftpFecha
            'mensaje &= ftpfile & vbTab & ftpBites & vbTab & ftpFecha & vbCrLf
            'End If
        End While

        ftpresp.Close()
        Return mensaje
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

    Public Function existeObjeto1(ByVal dir As String) As Boolean
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
        Catch ex As Exception
            ' Si el objeto no existe, se producirá un error y al entrar por el Catch
            ' se devolverá falso
            Return False
        End Try
    End Function

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

    Public Function DameDatosFichero(ByVal queFichero As String, ByVal queDato As queDatoFichero) As String
        Dim resultado As String = ""
        Dim peticionFTP As FtpWebRequest
        Try
            ' Creamos una peticion FTP con la dirección del objeto que queremos saber si existe
            peticionFTP = CType(WebRequest.Create(New Uri(queFichero)), FtpWebRequest)

            ' Fijamos el usuario y la contraseña de la petición
            peticionFTP.Credentials = New NetworkCredential(user, pass)

            ' Para saber si el objeto existe, solicitamos la fecha de creación del mismo
            'peticionFTP.Method = WebRequestMethods.Ftp.GetDateTimestamp

            '' PETICIÓN QUE ESTAMOS SOLICITANDO AL SERVIDOR
            If queDato = queDatoFichero.fecha Then
                peticionFTP.Method = WebRequestMethods.Ftp.GetDateTimestamp
            ElseIf queDato = queDatoFichero.tamaño Then
                peticionFTP.Method = WebRequestMethods.Ftp.GetFileSize
            End If

            peticionFTP.UsePassive = False

            Try
                ' Si el objeto existe, se devolverá True
                Dim respuestaFTP As FtpWebResponse
                respuestaFTP = CType(peticionFTP.GetResponse(), FtpWebResponse)
                If queDato = queDatoFichero.fecha Then
                    'Return respuestaFTP.LastModified
                    'resultado = respuestaFTP.LastModified.ToLongDateString
                    resultado = respuestaFTP.LastModified
                ElseIf queDato = queDatoFichero.tamaño Then
                    'Return respuestaFTP.ContentLength
                    resultado = respuestaFTP.ContentLength.ToString
                End If
            Catch ex As Exception
                ' Si el objeto no existe, se producirá un error y al entrar por el Catch
                ' se devolverá falso
            End Try
        Catch ex As Exception

        End Try
        Return resultado
    End Function
End Class
