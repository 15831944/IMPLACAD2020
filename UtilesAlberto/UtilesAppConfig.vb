Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Configuration
Imports System.Collections
Imports System.Collections.Specialized

Partial Public Class Utiles
    Private Const pathapplicationSettings = "/configuration/applicationSettings"
    Private Const pathconfigSections As String = "/configuration/configSections"
    Private Const pathconnectionstring As String = "/configuration/connectionStrings"
    Private Shared pathapplicationSettingsFull As String = ""
    Private Shared pathuserSettingsFull As String = ""

    ' ***** EJEMPLOS PARA LEER CONFIGURACIÓN desde Nombre.dll.config
    ' Public Function ReadConfig() As String
    '    Dim result As String = ""
    '    If clsI Is Nothing Then clsI = New Inventor2acad.Inventor2acad(oApp, log_path)
    '    LIBRARY = ConfigLee("LIBRARY").ToString  'My.Settings.BIBLIOTECA
    '    If LIBRARY.StartsWith(".\") Then LIBRARY = LIBRARY.Replace(".\", app_folder & "\")
    '    log = ConfigLee("log")  ' My.Settings.log
    '    Return result
    'End Function
    Private Shared Sub PonPathApplicationSettings(FullPathConfig As String)
        If IO.File.Exists(FullPathConfig) = False Then Exit Sub
        If pathapplicationSettingsFull <> "" Then Exit Sub
        '   <configuration>
        '      <configSections>
        '          <sectionGroup name = "applicationSettings" type="System.Configuration.ApplicationSettingsGroup, System, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089">
        '              <section name = "MKINVWEB.Settings" type="System.Configuration.ClientSettingsSection, System, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089" requirePermission="false"/>
        '           </sectionGroup>
        '   </configSections>
        '
        Dim doc As New Xml.XmlDocument
        Dim configSections As Xml.XmlNodeList
        ''
        ''
        doc.Load(FullPathConfig)
        ''
        configSections = doc.SelectNodes(pathconfigSections)
        '
        For Each oNo As Xml.XmlNode In configSections
            ' Recorrer <configSections> para sacar los grupos (applicationSettings y userSettings)
            For Each oNod As Xml.XmlNode In oNo.ChildNodes
                Try
                    Dim name As String = oNod.Attributes.GetNamedItem("name").Value
                    If name = "applicationSettings" Then
                        pathapplicationSettingsFull = pathapplicationSettings & "/" & oNod.FirstChild.Attributes.GetNamedItem("name").Value
                    ElseIf name = "userSettings" Then
                        pathuserSettingsFull = pathapplicationSettings & "/" & oNod.FirstChild.Attributes.GetNamedItem("name").Value
                    End If
                Catch ex As Exception
                    Continue For
                End Try
            Next
        Next
    End Sub
    Public Shared Function ConfigLee(cfg As Conf, nombrePro As String, Optional FullPathConfig As String = "") As Object
        If FullPathConfig = "" Or IO.File.Exists(FullPathConfig) = False Then
            FullPathConfig = cfg._appconfig
        End If
        ' Coger del fichero .config la rama de la configuración
        PonPathApplicationSettings(FullPathConfig)
        '
        Dim resultado As String = ""
        'Dim solonombre As String = IO.Path.GetFileNameWithoutExtension(IO.Path.GetFileNameWithoutExtension(FullPathConfig))
        ''
        Dim user As Boolean = False
        '
        Dim doc As New Xml.XmlDocument
        Dim settings As Xml.XmlNodeList
        ''
        If IO.File.Exists(FullPathConfig) = False Then
            Return "ERROR"
            Exit Function
        End If
        ''
        doc.Load(FullPathConfig)
        ''
        settings = doc.SelectNodes(pathapplicationSettingsFull)
        ''
REPITE:
        For Each oNo As Xml.XmlNode In settings
            For Each oNod As Xml.XmlNode In oNo.ChildNodes
                If oNod.Attributes.GetNamedItem("name").Value.ToUpper(Globalization.CultureInfo.CurrentCulture) = nombrePro.ToUpper(Globalization.CultureInfo.CurrentCulture) Then
                    resultado = oNod.FirstChild.InnerText
                    Exit For
                End If
            Next
        Next
        ''
        If resultado = "" And user = False Then  '' Por si la configuración esta en userSettings
            settings = doc.SelectNodes(pathuserSettingsFull)
            user = True
            GoTo REPITE
        ElseIf resultado = "True" Or resultado = "False" Then
            Return CType(resultado, Boolean)
        ElseIf resultado <> "" AndAlso IsNumeric(resultado) Then
            Return CDbl(resultado)
        Else
            Return resultado.ToString
        End If
    End Function
    ''
    Public Shared Sub ConfigEscribe(cfg As Conf, nombrePro As String, valorPro As Object, Optional FullPathConfig As String = "")
        If FullPathConfig = "" OrElse IO.File.Exists(FullPathConfig) = False Then
            FullPathConfig = cfg._appconfig
        End If
        '
        'Dim solonombre As String = IO.Path.GetFileNameWithoutExtension(IO.Path.GetFileNameWithoutExtension(FullPathConfig))
        Dim doc As New Xml.XmlDocument
        Dim settings As Xml.XmlNodeList
        ''
        Dim user As Boolean = False
        '
        If IO.File.Exists(FullPathConfig) = False Then
            Exit Sub
        End If
        ''
        doc.Load(FullPathConfig)
        ''
        'settings = doc.SelectNodes("/configuration/applicationSettings/" & solonombre & ".My.MySettings/setting")
        settings = doc.SelectNodes(pathapplicationSettingsFull)
        ''
REPITE:
        Dim cambiado As Boolean = False
        For Each oNo As Xml.XmlNode In settings
            For Each oNod As Xml.XmlNode In oNo.ChildNodes
                If oNod.Attributes.GetNamedItem("name").Value.ToUpper(Globalization.CultureInfo.CurrentCulture) = nombrePro.ToUpper(Globalization.CultureInfo.CurrentCulture) Then
                    oNod.FirstChild.InnerText = valorPro
                    cambiado = True
                    Exit For
                End If
            Next
        Next
        '
        If cambiado = False And user = False Then    ' Por si la configuración está en "userSettings"
            'settings = doc.SelectNodes("/configuration/userSettings/" & solonombre & ".My.MySettings/setting")
            settings = doc.SelectNodes(pathuserSettingsFull)
            user = True
            GoTo REPITE
        ElseIf cambiado = True Then
            doc.Save(FullPathConfig)
        End If
    End Sub
    '
    ' ConexionString
    Public Shared Function ConfigLee_connectionStrings(cfg As Conf, nombrePro As String, Optional FullPathConfig As String = "") As Object
        If FullPathConfig = "" Or IO.File.Exists(FullPathConfig) = False Then
            FullPathConfig = cfg._appconfig
        End If
        '
        Dim resultado As String = ""
        '
        Dim doc As New Xml.XmlDocument
        Dim settings As Xml.XmlNodeList
        ''
        If IO.File.Exists(FullPathConfig) = False Then
            Return "ERROR"
            Exit Function
        End If
        ''
        doc.Load(FullPathConfig)
        ''
        settings = doc.SelectNodes(pathconnectionstring)
        ''
        For Each oNo As Xml.XmlNode In settings
            For Each oNod As Xml.XmlNode In oNo.ChildNodes
                If oNod.Attributes.GetNamedItem("name").Value.ToUpper(Globalization.CultureInfo.CurrentCulture) = nombrePro.ToUpper(Globalization.CultureInfo.CurrentCulture) Then
                    resultado = oNod.Attributes.GetNamedItem("connectionString").Value
                    Exit For
                End If
            Next
        Next
        Return resultado.ToString
    End Function
    ''
    Public Shared Sub ConfigEscribe_connectionStrings(cfg As Conf, nombrePro As String, valorPro As Object, Optional FullPathConfig As String = "")
        If FullPathConfig = "" OrElse IO.File.Exists(FullPathConfig) = False Then
            FullPathConfig = cfg._appconfig
        End If
        '
        Dim doc As New Xml.XmlDocument
        Dim settings As Xml.XmlNodeList
        '
        If IO.File.Exists(FullPathConfig) = False Then
            Exit Sub
        End If
        ''
        doc.Load(FullPathConfig)
        ''
        settings = doc.SelectNodes(pathconnectionstring)
        '
        Dim cambiado As Boolean = False
        For Each oNo As Xml.XmlNode In settings
            For Each oNod As Xml.XmlNode In oNo.ChildNodes
                If oNod.Attributes.GetNamedItem("name").Value.ToUpper(Globalization.CultureInfo.CurrentCulture) = nombrePro.ToUpper(Globalization.CultureInfo.CurrentCulture) Then
                    oNod.Attributes.GetNamedItem("connectionString").Value = valorPro
                    cambiado = True
                    Exit For
                End If
            Next
        Next
        '
        If cambiado = True Then
            doc.Save(FullPathConfig)
        End If
    End Sub
End Class
