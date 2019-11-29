Module modAppConfig
    ' ***** EJEMPLOS PARA LEER CONFIGURACIÓN desde Nombre.dll.config
    ' Public Function ReadConfig() As String

    '    Dim result As String = ""
    '    If clsI Is Nothing Then clsI = New Inventor2acad.Inventor2acad(oApp, log_path)
    '    ''
    '    LIBRARY = ConfigLee("LIBRARY").ToString  'My.Settings.BIBLIOTECA
    '    If LIBRARY.StartsWith(".\") Then LIBRARY = LIBRARY.Replace(".\", app_folder & "\")
    '    ''
    '    TEMPLATES = ConfigLee("TEMPLATES").ToString    ' My.Settings.PLANTILLAS
    '    If TEMPLATES.StartsWith(".\") Then TEMPLATES = TEMPLATES.Replace(".\", app_folder & "\")
    '    ''
    '    log = ConfigLee("log")  ' My.Settings.log
    '    ''
    '    lastfolder = ConfigLee("lastfolder")  'My.Settings.ultimoDir
    '    ''
    '    zMin = CDbl(ConfigLee("zMin"))
    '    ''
    '    tolTowers = CInt(ConfigLee("tolTowers"))
    '    ''
    '    tolMark = CInt(ConfigLee("tolMark"))
    '    ''
    '    minMark = CInt(ConfigLee("minMark"))
    '    ''
    '    usuario = System.Environment.UserName
    '    maquina = System.Environment.MachineName
    '    year = Date.Now.Year.ToString
    '    ''
    '    If log Then clsI.PonLog("ReadConfig...")
    '    If log Then clsI.PonLog(vbTab & "log = " & log.ToString)
    '    If log Then clsI.PonLog(vbTab & "Nombre PC = " & maquina)
    '    If log Then clsI.PonLog(vbTab & "LIBRARY = " & LIBRARY)
    '    If log Then clsI.PonLog(vbTab & "TEMPLATES = " & TEMPLATES)
    '    If log Then clsI.PonLog(vbTab & "lastfolder = " & lastfolder)
    '    If log Then clsI.PonLog(vbTab & "preGAUGE = " & preGAUGE)
    '    If log Then clsI.PonLog(vbTab & "preGAUGET = " & preGAUGET)
    '    If log Then clsI.PonLog(vbTab & "tolTowers = " & tolTowers & " mm")
    '    If log Then clsI.PonLog(vbTab & "tolMark = " & tolMark & " mm")
    '    If log Then clsI.PonLog(vbTab & "minMark = " & minMark & " mm")
    '    ''
    '    '' Create and Load images in imageLists
    '    'ImagelistLoadFull()
    '    ''
    '    '' If error write in result
    '    If My.Computer.FileSystem.DirectoryExists(LIBRARY) = False Then
    '        result &= "The folder does not exist LIBRARY: " & vbCrLf & TEMPLATES & vbCrLf
    '    End If
    '    ''
    '    If My.Computer.FileSystem.DirectoryExists(TEMPLATES) = False Then
    '        result &= "The folder does not exist TEMPLATES: " & vbCrLf & TEMPLATES & vbCrLf
    '    End If
    '    ''
    '    For Each queExt As String In extensiones
    '        Dim fiPlan As String = preGAUGE & queExt
    '        Dim fiFin As String = IO.Path.Combine(TEMPLATES, fiPlan)
    '        ''
    '        If My.Computer.FileSystem.FileExists(fiFin) = False Then
    '            result &= "Does not exist --> " & fiPlan & vbCrLf & "In folder TEMPLATES -->  " & TEMPLATES & vbCrLf & vbCrLf
    '        Else
    '            Select Case queExt
    '                Case ".iam" : pgIAM = fiFin
    '                Case ".ipn" : pgIPN = fiFin
    '                Case ".idw" : pgIDW = fiFin
    '                Case ".ipt" : pgIPT = fiFin
    '                Case "T.ipt" : pgIPTT = fiFin
    '            End Select
    '        End If
    '    Next
    '    ''
    '    Return result
    'End Function
    Public Function ConfigLee(FullPathConfig As String, nombrePro As String) As Object
        Dim resultado As String = ""
        Dim solonombre As String = IO.Path.GetFileNameWithoutExtension(FullPathConfig).Replace(".dll", "")
        ''
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
        settings = doc.SelectNodes("/configuration/applicationSettings/" & solonombre & ".My.MySettings/setting")
        ''
        For Each oNod As Xml.XmlNode In settings
            If oNod.Attributes.GetNamedItem("name").Value.ToUpper = nombrePro.ToUpper Then
                resultado = oNod.FirstChild.InnerText
                Exit For
            End If
        Next
        ''
        If resultado = "True" Or resultado = "False" Then
            Return CType(resultado, Boolean)
        Else
            Return resultado.ToString
        End If
    End Function
    ''
    Public Sub ConfigEscribe(FullPathConfig As String, nombrePro As String, valorPro As Object)
        Dim solonombre As String = IO.Path.GetFileNameWithoutExtension(FullPathConfig)
        Dim doc As New Xml.XmlDocument
        Dim settings As Xml.XmlNodeList
        ''
        If IO.File.Exists(FullPathConfig) = False Then
            Exit Sub
        End If
        ''
        doc.Load(FullPathConfig)
        ''
        settings = doc.SelectNodes("/configuration/applicationSettings/" & solonombre & ".My.MySettings/setting")
        ''
        Dim cambiado As Boolean = False
        For Each oNod As Xml.XmlNode In settings
            If oNod.Attributes.GetNamedItem("name").Value.ToUpper = nombrePro.ToUpper Then
                oNod.FirstChild.InnerText = valorPro
                cambiado = True
                Exit For
            End If
        Next
        If cambiado = True Then doc.Save(FullPathConfig)
    End Sub
End Module
