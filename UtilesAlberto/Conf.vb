Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Configuration
Imports System.Collections
Imports System.Collections.Specialized
Imports System.Reflection

''' <summary>
''' Siempre instanciar esta clase (Public cfg as UtilesAlberto.Conf) en la aplicación que desarrollemos
''' cfg = New UtilesAlberto.Conf (O cfg= New us.Conf, si Imports ua = UtilesAlberto)
''' </summary>
Public Class Conf
    ' Constantes
    Public Const nFijo As String = "2aCAD"
    Private Const keyPre As String = "aiiao2K19"
    Private Const keySuf As String = "2aCAD"
    Private ReadOnly activado As Boolean = False

    ' Variables aplicacion iniciales
    Public _appFullPath As String
    Public _appname As String
    Public _appnameext As String
    Public _appfolder As String
    Public _appfolderandname As String
    Public _appconfig As String
    Public _applog As String
    Public _appini As String
    Public _appOUT As String
    Public _appRESOURCES As String
    Public _appCONFIGURATORS As String
    Public _appPROJECTS As String           ' Para Inventor
    Public _appTEMPLATES As String          ' Para Inventor
    Public _appversion As String
    Public _appnameversion As String
    Public _appdate As String
    '
    Public _OAssembly As Assembly
    Public _Appiwin32 As IntPtr

    '
    Public _User As String = System.Environment.UserName
    Public _Machine As String = System.Environment.MachineName
    Public _Domain As String = System.Environment.UserDomainName

    ' Variables de configuración (XXXXXX.dll.config)
    Public _Log As Boolean = False       ' Generar fichero log (True/False)
    Public _Messages As Boolean = False  ' Mostrar mensajes en la aplicación (True/False)
    Public _Interval As Integer = 5      ' Para usos varios con Timers (Segundos)
    Public _connectionString As String = ""
    '
    ' ***** Llenar Conf primero *****
    ' Conf.PonFijos(ass:=Reflection.Assembly.GetExecutingAssembly, bolLog:=Log, bolMensajes:=Log, douIntervalo:=5)

    Public Sub New(key As String, Optional ass As Assembly = Nothing,
                               Optional bLog As Boolean = False,
                               Optional bMensajes As Boolean = False,
                               Optional dInterval As Double = 20)
        If key <> keyPre & "_" & keySuf Then
            activado = False
            Throw New System.Exception("Not allowed")
            Return
        Else
            activado = True
        End If

        If ass Is Nothing Then
            Me._OAssembly = System.Reflection.Assembly.GetExecutingAssembly
        Else
            Me._OAssembly = ass
        End If
        '
        Me._appFullPath = Me._OAssembly.Location
        Me._appname = IO.Path.GetFileNameWithoutExtension(_appFullPath)
        Me._appnameext = IO.Path.GetFileName(_appFullPath)
        Me._appfolder = IO.Path.GetDirectoryName(_appFullPath)
        Me._appfolderandname = IO.Path.Combine(_appfolder, _appname)
        Me._appconfig = _appFullPath & ".config"
        Me._applog = IO.Path.ChangeExtension(_appFullPath, ".log")
        Me._appini = IO.Path.ChangeExtension(_appFullPath, ".ini")
        Me._appOUT = IO.Path.Combine(_appfolder, "OUT")
        Me._appRESOURCES = IO.Path.Combine(_appfolder, "RESOURCES")
        Me._appCONFIGURATORS = IO.Path.Combine(_appfolder, "CONFIGURATORS")
        Me._appPROJECTS = IO.Path.Combine(_appfolder, "PROJECTS")
        Me._appTEMPLATES = IO.Path.Combine(_appfolder, "TEMPLATES")
        Me._appversion = _OAssembly.GetName.Version.ToString
        Me._appnameversion = _appname & " - " & _appversion
        Me._appdate = System.IO.File.GetLastWriteTime(Me._appFullPath).ToShortDateString
        '
        Try
            Me._Log = bLog
            Me._Messages = bMensajes
            Me._Interval = dInterval
        Catch ex As Exception
            Debug.Print(ex.ToString)
        End Try
    End Sub

    Public Sub PonLog(ByVal quetexto As String, Optional ByVal borrar As Boolean = False)
        If Me._OAssembly Is Nothing Then
            'OAssembly = System.Reflection.Assembly.GetExecutingAssembly
            Exit Sub
        End If
        '
        If borrar = True And IO.File.Exists(Me._applog) Then IO.File.Delete(Me._applog)
        If quetexto.EndsWith(vbCrLf) = False Then quetexto &= vbCrLf
        IO.File.AppendAllText(Me._applog, Date.Now & vbTab & quetexto)
    End Sub
    Public Overrides Function ToString() As String
        Dim resultado As New StringBuilder()
        resultado.AppendLine("_appfull = " & Me._appFullPath)
        resultado.AppendLine("_appname = " & Me._appname)
        resultado.AppendLine("_appnameext = " & Me._appnameext)
        resultado.AppendLine("_appfolder = " & Me._appfolder)
        resultado.AppendLine("_appfolderandname = " & Me._appfolderandname)
        resultado.AppendLine("_appversion = " & Me._appversion)

        resultado.AppendLine("_appconfig = " & Me._appconfig)
        resultado.AppendLine("_applog = " & Me._applog)
        resultado.AppendLine("_appini = " & Me._appini)

        resultado.AppendLine("_user = " & Me._User)
        resultado.AppendLine("_machine = " & Me._Machine)
        resultado.AppendLine("_domain = " & Me._Domain)

        resultado.AppendLine("_Log = " & Me._Log.ToString)
        resultado.AppendLine("_Messages = " & Me._Messages.ToString)
        resultado.AppendLine("_Interval = " & Me._Interval.ToString)

        resultado.AppendLine("_appOUT = " & Me._appOUT)
        resultado.AppendLine("_appRESOURCES = " & _appRESOURCES)
        resultado.AppendLine("_appCONFIGURATORS = " & Me._appCONFIGURATORS)
        '
        Return resultado.ToString
    End Function
End Class
