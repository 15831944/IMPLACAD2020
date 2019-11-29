Imports System
Imports System.IO
Imports System.Xml
Imports System.ComponentModel
Imports System.Data
Imports System.Drawing
Imports System.Windows.Forms
Imports System.Data.SqlClient
''
Imports Autodesk.AutoCAD.EditorInput
'Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.Interop
Imports Autodesk.AutoCAD.Interop.Common

Module modVar
    ''
    '' ***** FORMULARIOS
    Public frmE As frmEtiquetas
    '' OBJETOS AUTOCAD
    Public WithEvents oApp As Autodesk.AutoCAD.Interop.AcadApplication = Nothing
    Public WithEvents oDoc As Autodesk.AutoCAD.Interop.AcadDocument = Nothing
    'Public oDoc As Autodesk.AutoCAD.Interop.AcadDocument = Nothing
    'Public oDocImprimir As Autodesk.AutoCAD.Interop.AcadDocument = Nothing
    Public oSel As Autodesk.AutoCAD.Interop.AcadSelectionSet = Nothing       '"SERICAD"
    Public oSelTemp As Autodesk.AutoCAD.Interop.AcadSelectionSet = Nothing   '"TEMPORAL"
    Public bloqueEditar As Autodesk.AutoCAD.Interop.Common.AcadBlockReference = Nothing
    Public TablaDatos As Autodesk.AutoCAD.Interop.Common.AcadTable = Nothing
    Public TablaDatosParcial As Autodesk.AutoCAD.Interop.Common.AcadTable = Nothing
    Public TablaEscaleras As Autodesk.AutoCAD.Interop.Common.AcadTable = Nothing
    Public textoEstilo As Autodesk.AutoCAD.Interop.Common.AcadTextStyle = Nothing
    ''
    '' CLASSES
    Public cIni As clsINI
    ''
    '' OTROS OBJETOS
    'Public WithEvents oTimer As Timer
    ''
    '' COLECCIONES
    Public arrBloquesIdParcial As ArrayList
    Public arrBalizaSuelo As ArrayList
    Public arrBalizaPared As ArrayList
    Public colBloquesCantidadParcial As Hashtable  '' Colección de key=nombre del bloque, value=cantidad
    Public colBloquesCantidad As Hashtable  '' Colección de key=nombre del bloque, value=cantidad
    Public arrBloquesId As ArrayList        '' Array de los ID de los bloques de etiqueta.
    Public colBotones As Hashtable          '' Hashtable de botones (Key=Nombre, Value=estado (true/false)
    Public arrBotones As ArrayList          '' ArrayList con el nombre de todos los botones.
    ''
    ''
    '' CONSTANTES
    Public Const regAPP As String = "IMPLACAD"
    Public Const estilotexto As String = "RRC_arial"
    Public Const codigoactivacion As String = "VIP2796"
    Public Const balizas As String = "PLANOS_Y_BALIZAMIENTOS"
    ''
    ''
    '' VARIABLES
    'Public bloqueEditar As String = ""      ' Nombre del bloque a cambiar.
    'Public bloqueID As Long          ' Object ID del bloque a cambiar.
    '' escala=1 (si todo está en mm) / escala=0.1 (si todo está en cm) / escala=0.01 (Si todo esta en m)
    Public escalaTotal As Double = 0.02     '' Con esta variable escalaremos bloques y texto
    'Public dirApp As String
    Public dirBase As String            '' Directorio Base de bloques C:\ProgramData\IMPLACAD  (Ponemos la barra al final)
    Public actualizardatos As String    '' enlace para descargar actualización de datos
    Public actualizarbd As String       '' enlace para descargar la actualización de bd
    Public webActualiza As String       '' Direccion Web del directorio de descarga.
    Public vermensajes As Boolean = True
    Public nombreviejo As String = ""
    Public Tipo As String = "*"
    Public Tipo1 As String = "*"
    Public Tipo2 As String = "*"
    Public Tipo3 As String = "*"
    ''
    Public Function INICargar(Optional ByVal nombreINI As String = "") As String
        If nombreINI = "" Then nombreINI = nIni
        If cIni Is Nothing Then cIni = New clsINI
        Call Utilidades.PermisosTodo()
        Dim mensaje As String = ""
        '; Este es un archivo de configuracion de ejemplo
        '; Los comentarios comienzan con ';'
        '[OPCIONES]
        'dirBase=C:\ProgramData\IMPLACAD
        'log = 1
        ''
        ''
        '' RELLENAR VARIABLES SIMPLES
        dirBase = cIni.IniGet(nombreINI, "OPCIONES", "dirBase")        ' Directorio base por defecto
        If dirBase.EndsWith("\") = False Then dirBase &= "\"
        ''
        actualizardatos = cIni.IniGet(nombreINI, "OPCIONES", "actualizardatos")        ' Enlace para actualziar datos
        actualizarbd = cIni.IniGet(nombreINI, "OPCIONES", "actualizarbd")        ' Enlace para actualizar BD
        webActualiza = cIni.IniGet(nombreINI, "OPCIONES", "webActualiza")        ' Enlace para actualizar BD
        If webActualiza.EndsWith("/") = False Then webActualiza &= "/"
        log = IIf(cIni.IniGet(nombreINI, "OPCIONES", "log") = "1", True, False)         '' Fichero log para control errores (Si o No)
        Tipo = cIni.IniGet(nombreINI, "OPCIONES", "Tipo")        ' Tipo
        Tipo1 = cIni.IniGet(nombreINI, "OPCIONES", "Tipo1")        ' Tipo
        Tipo2 = cIni.IniGet(nombreINI, "OPCIONES", "Tipo2")        ' Tipo
        Tipo3 = cIni.IniGet(nombreINI, "OPCIONES", "Tipo3")        ' Tipo
        ''
        ''
        ''
        My.Computer.FileSystem.CurrentDirectory = dirApp
        ''

        Dim btnTodos As String() = New String() { _
            "IMPLACADMENU", "INSERTAREDITAR", "ADECUA", "BALIZARSUELO", "BALIZARPARED", _
            "BALIZARESCALERA", "RUTAEVACUACION", "TABLADATOS", "TABLAESCALERAS", _
            "CAPASCOBERTURA", "GROSORLINEAS", "ESCALAM", "IMPRIMIRINS", "IMPRIMIREVA", _
            "EXPLOTAEVA", "TABLAPARCIAL", "_ETRANSMIT...", "ACTUALIZARIMPLACAD"}
        arrBotones = New ArrayList(btnTodos)
        colBotones = New Hashtable
        For Each nBoton As String In btnTodos
            colBotones.Add(nBoton, True)
        Next
        INICargar = mensaje
    End Function
    ''
    Public Function Image2Bytes(ByVal img As Image) As Byte()
        Dim sTemp As String = Path.GetTempFileName()
        Dim fs As New FileStream(sTemp, FileMode.OpenOrCreate, FileAccess.ReadWrite)
        img.Save(fs, System.Drawing.Imaging.ImageFormat.Png)
        fs.Position = 0
        '
        Dim imgLength As Integer = CInt(fs.Length)
        Dim bytes(0 To imgLength - 1) As Byte
        fs.Read(bytes, 0, imgLength)
        fs.Close()
        'Return bytes
        Image2Bytes = bytes
        Exit Function
    End Function

    Public Function Bytes2Image(ByVal bytes() As Byte) As Image
        If bytes Is Nothing Then Return Nothing
        '
        Dim ms As New MemoryStream(bytes)
        Dim bm As Bitmap = Nothing
        Try
            bm = New Bitmap(ms)
        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine(ex.Message)
        End Try
        'Return bm
        Bytes2Image = bm
        Exit Function
    End Function

    Public Sub Bytes2EscribeFichero(ByVal bytes() As Byte, queFi As String)
        Dim fs As FileStream = New FileStream(queFi, FileMode.Create)
        ''Y escribimos en disco el array de bytes que conforman
        ''el fichero que sea (DWG, Word, Excel, etc.)
        fs.Write(bytes, 0, Convert.ToInt32(bytes.Length))
        fs.Close()
    End Sub
    ''
    Public Function ImplacadEstado() As Estado
        ''ActiNo = 0
        ''ActiSi = 1
        ''ActuNo = 2
        ''ActuSi = 3
        ''EscalaMSi = 4
        ''EscalaMNo = 5
        If oApp Is Nothing Then _
        oApp = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
        ''
        Dim resultado As Integer = 0
        ''
        '' Si no está activado. no dejamos hacer nada. Desactivar todos los botones menos actualizar.
        If IO.File.Exists(dirApp & nApp & ".imp") = False Then
            resultado = Estado.ActiNo
            Exit Function
        Else
            Dim quecodigo As String = IO.File.ReadAllText(dirApp & nApp & ".imp")
            If quecodigo.ToUpper <> codigoactivacion.ToUpper Then
                resultado = Estado.ActiNo
                Exit Function
            End If
        End If
        ''
        '' Si está activo. Dejamos actualizar sólo. Activar botones, a la espera de EscalaMSi.
        If IO.Directory.Exists("C:\ProgramData\IMPLACAD") = False Then
            resultado = Estado.ActiSi + Estado.ActuNo
        Else
            If PropiedadLee("ESCALAM") <> "" Then
                Try
                    'Call oApp.Application.MenuGroups.Item("EXPRESS")
                    resultado = Estado.ActiSi + Estado.ActuSi + Estado.EscalaMSi + Estado.ExpressToolsSi
                Catch ex As Exception
                    'resultado = Estado.ActiSi + Estado.ActuSi + Estado.EscalaMSi + Estado.ExpressToolsNo
                End Try
            Else
                resultado = Estado.ActiSi + Estado.ActuSi + Estado.EscalaMNo
            End If
        End If
        ''
        Return resultado
    End Function
    ''
    Public Function ImplacadActivado(Optional conMensajes As Boolean = True) As Boolean
        If oApp Is Nothing Then _
        oApp = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
        ''
        If oApp.Documents.Count = 0 Then Exit Function
        Dim resultado As Boolean = False
        ''
        '' Si no está activado. no dejamos hacer nada. Desactivar todos los botones menos actualizar.
        If IO.File.Exists(dirApp & nApp & ".imp") = False Then
            If conMensajes Then
                MsgBox("¡¡ IMPLACAD no está activado !!" & vbCrLf & vbCrLf & _
                       "Introduzca el código de activación desde ACTUALIZAIMPLACAD.", MsgBoxStyle.Critical)
            End If
            resultado = False
        Else
            Dim quecodigo As String = IO.File.ReadAllText(dirApp & nApp & ".imp")
            If quecodigo.ToUpper = codigoactivacion.ToUpper Then
                resultado = True
            Else
                If conMensajes Then
                    MsgBox("¡¡ IMPLACAD no está activado !!" & vbCrLf & vbCrLf & _
                           "Introduzca el código de activación desde ACTUALIZAIMPLACAD.", MsgBoxStyle.Critical)
                End If
                resultado = False
            End If
        End If
        ''
        Return resultado
    End Function
    ''
    'Private Sub oApp_AppActivate() Handles oApp.AppActivate
    '    If oApp.Documents.Count = 0 Then Exit Sub
    '    ''
    '    Dim queEstado As Integer = ImplacadEstado()
    '    Try
    '        If queEstado = Estado.ActiNo Or queEstado = Estado.ActiSi + Estado.ActuNo Then
    '            '' Solo botón actualizar. Pendiente de activar. O activo pero no actualizado
    '            BotonesImplacad(ActivarBotones.NoActivado)
    '        ElseIf queEstado = Estado.ActiSi + Estado.ActuSi + Estado.EscalaMSi + Estado.ExpressToolsSi Then
    '            '' Todo activado
    '            BotonesImplacad(ActivarBotones.SiActivadoConEscala)
    '        ElseIf queEstado = Estado.ActiSi + Estado.ActuSi + Estado.EscalaMSi + Estado.ExpressToolsNo Then
    '            '' Todo activado y quitamos después "ADECUA"
    '            BotonesImplacad(ActivarBotones.SiActivadoConEscala)
    '            'BotonesImplacad(ActivarBotones.Cualquiera, "ADECUA", False)
    '        ElseIf queEstado = Estado.ActiSi + Estado.ActuSi + Estado.EscalaMNo Then
    '            '' Solo botón actualizar. Pendiente de activar. O activo pero no actualizado
    '            BotonesImplacad(ActivarBotones.SiActivadoSinEscala)
    '        End If
    '        'If oApp.ActiveDocument.Name.Contains("·EVA·A4H.dwg") = False Or _
    '        '    oApp.ActiveDocument.Name.Contains("·EVA·A4H.dwg") = False Or _
    '        '    oApp.ActiveDocument.Name.Contains("·EVA·A4H.dwg") = False Or _
    '        '    oApp.ActiveDocument.Name.Contains("·EVA·A4H.dwg") = False Then
    '        '    BotonesImplacad(ActivarBotones.Cualquiera, "EXPLOTAEVA", False)
    '        'End If

    '    Catch ex As Exception
    '        'MsgBox(ex.Message)
    '    End Try
    'End Sub

    'Private Sub oApp_EndCommand(CommandName As String) Handles oApp.EndCommand
    '    'Debug.Print(CommandName)
    '    ''
    '    If oApp.Documents.Count = 0 Then Exit Sub
    '    ''
    '    If CommandName <> "NEW" And CommandName <> "OPEN" Then
    '        Exit Sub
    '    End If
    '    ''
    '    oDoc = oApp.ActiveDocument
    '    ''
    '    Dim queEstado As Integer = ImplacadEstado()
    '    Try
    '        If queEstado = Estado.ActiNo Or queEstado = Estado.ActiSi + Estado.ActuNo Then
    '            '' Solo botón actualizar. Pendiente de activar. O activo pero no actualizado
    '            BotonesImplacad(ActivarBotones.NoActivado)
    '        ElseIf queEstado = Estado.ActiSi + Estado.ActuSi + Estado.EscalaMSi + Estado.ExpressToolsSi Then
    '            '' Todo activado
    '            BotonesImplacad(ActivarBotones.SiActivadoConEscala)
    '        ElseIf queEstado = Estado.ActiSi + Estado.ActuSi + Estado.EscalaMSi + Estado.ExpressToolsNo Then
    '            '' Todo activado y quitamos después "ADECUA"
    '            BotonesImplacad(ActivarBotones.SiActivadoConEscala)
    '            'BotonesImplacad(ActivarBotones.Cualquiera, "ADECUA", False)
    '        ElseIf queEstado = Estado.ActiSi + Estado.ActuSi + Estado.EscalaMNo Then
    '            '' Solo botón actualizar. Pendiente de activar. O activo pero no actualizado
    '            BotonesImplacad(ActivarBotones.SiActivadoSinEscala)
    '        End If
    '    Catch ex As Exception
    '        MsgBox(ex.Message)
    '    End Try
    'End Sub
    ' ''
    'Public Sub oApp_EndOpen(FileName As String) Handles oApp.EndOpen
    '    ''
    '    oDoc = oApp.ActiveDocument
    '    Dim queEstado As Integer = ImplacadEstado()
    '    Try
    '        If queEstado = Estado.ActiNo Or queEstado = Estado.ActiSi + Estado.ActuNo Then
    '            '' Solo botón actualizar. Pendiente de activar. O activo pero no actualizado
    '            BotonesImplacad(ActivarBotones.NoActivado)
    '        ElseIf queEstado = Estado.ActiSi + Estado.ActuSi + Estado.EscalaMSi + Estado.ExpressToolsSi Then
    '            '' Todo activado
    '            BotonesImplacad(ActivarBotones.SiActivadoConEscala)
    '        ElseIf queEstado = Estado.ActiSi + Estado.ActuSi + Estado.EscalaMSi + Estado.ExpressToolsNo Then
    '            '' Todo activado y quitamos después "ADECUA"
    '            BotonesImplacad(ActivarBotones.SiActivadoConEscala)
    '            'BotonesImplacad(ActivarBotones.Cualquiera, "ADECUA", False)
    '        ElseIf queEstado = Estado.ActiSi + Estado.ActuSi + Estado.EscalaMNo Then
    '            '' Solo botón actualizar. Pendiente de activar. O activo pero no actualizado
    '            BotonesImplacad(ActivarBotones.SiActivadoSinEscala)
    '        End If
    '        'If oApp.ActiveDocument.Name.Contains("·EVA·A4H.dwg") = False Or _
    '        '    oApp.ActiveDocument.Name.Contains("·EVA·A4H.dwg") = False Or _
    '        '    oApp.ActiveDocument.Name.Contains("·EVA·A4H.dwg") = False Or _
    '        '    oApp.ActiveDocument.Name.Contains("·EVA·A4H.dwg") = False Then
    '        '    BotonesImplacad(ActivarBotones.Cualquiera, "EXPLOTAEVA", False)
    '        'End If

    '    Catch ex As Exception
    '        'MsgBox(ex.Message)
    '    End Try
    'End Sub
    ''
    'Private Sub oDoc_Activate() Handles oDoc.Activate
    '    ''
    '    oDoc = oApp.ActiveDocument
    '    ''
    '    Dim queEstado As Integer = ImplacadEstado()
    '    Try
    '        If queEstado = Estado.ActiNo Or queEstado = Estado.ActiSi + Estado.ActuNo Then
    '            '' Solo botón actualizar. Pendiente de activar. O activo pero no actualizado
    '            BotonesImplacad(ActivarBotones.NoActivado)
    '        ElseIf queEstado = Estado.ActiSi + Estado.ActuSi + Estado.EscalaMSi + Estado.ExpressToolsSi Then
    '            '' Todo activado
    '            BotonesImplacad(ActivarBotones.SiActivadoConEscala)
    '        ElseIf queEstado = Estado.ActiSi + Estado.ActuSi + Estado.EscalaMSi + Estado.ExpressToolsNo Then
    '            '' Todo activado y quitamos después "ADECUA"
    '            BotonesImplacad(ActivarBotones.SiActivadoConEscala)
    '            'BotonesImplacad(ActivarBotones.Cualquiera, "ADECUA", False)
    '        ElseIf queEstado = Estado.ActiSi + Estado.ActuSi + Estado.EscalaMNo Then
    '            '' Solo botón actualizar. Pendiente de activar. O activo pero no actualizado
    '            BotonesImplacad(ActivarBotones.SiActivadoSinEscala)
    '        End If
    '    Catch ex As Exception
    '        MsgBox(ex.Message)
    '    End Try
    'End Sub
End Module
