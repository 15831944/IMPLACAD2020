Imports System.Drawing

'
'Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.Interop.Common

Module modVar

    ''
    '' ***** FORMULARIOS
    Public FrmE As frmEtiquetas

    Public FrmZ As frmZonas

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
    Public TablaZonas As Autodesk.AutoCAD.Interop.Common.AcadTable = Nothing
    Public TablaZona As Autodesk.AutoCAD.Interop.Common.AcadTable = Nothing
    Public textoEstilo As Autodesk.AutoCAD.Interop.Common.AcadTextStyle = Nothing

    ''
    '' CLASSES
    Public cIni As clsINI = Nothing

    Public cfg As UtilesAlberto.Conf = Nothing
    Public ua As UtilesAlberto.Utiles = Nothing
    Public oEtis As Etiquetas = Nothing

    ''
    '' OTROS OBJETOS
    'Public WithEvents oTimer As Timer
    ''
    '' COLECCIONES
    Public arrBloquesIdParcial As ArrayList

    Public arrBalizaSuelo As ArrayList
    Public arrBalizaPared As ArrayList
    Public colBloquesCantidadParcial As SortedList  ' Hashtable  '' Colección de key=nombre del bloque, value=cantidad
    Public colBloquesCantidad As SortedList ' Hashtable  '' Colección de key=nombre del bloque, value=cantidad

    'Public arrBloquesId As ArrayList        '' Array de los ID de los bloques de etiqueta.
    Public colBotones As Hashtable          '' Hashtable de botones (Key=Nombre, Value=estado (true/false)

    Public arrBotones As ArrayList          '' ArrayList con el nombre de todos los botones.
    Public arrpreEti As ArrayList           '' ArrayList de los prefijos de etiquetas IMPLACAD (Solo estos se tendrán en cuenta) El resto se borrará XData

    'Public colConjuntos As Hashtable         '' Hashtable de conjuntos de etiquetas (Key=REFERENCIA, Value=Array de Referencias)
    'Public colSustituciones As Hashtable    '' Hashtable de sustituciones para plano EVA (Key=Referencia de bloque, Value=nombre de la imagen que la sustituye [Sin .png])
    'Public LConjuntos As List(Of String)    ' List de Conjuntos (TIPO3=CONJUNTO o Largo/Ancho=0)
    Public LLwPolyline As List(Of AcadLWPolyline) = Nothing

    Public LZona As List(Of Zona) = Nothing

    '
    ' CONSTANTES
    Public IMPLACAD_DATA As String = "C:\ProgramData\IMPLACAD\"

    Public Const IMPLACAD_BUNDLE = "C:\ProgramData\Autodesk\ApplicationPlugins\IMPLACAD.bundle\"
    Public arrPaths As List(Of String)
    Public Const regAPP As String = "IMPLACAD"
    Public Const estilotexto As String = "RRC_arial"
    Public Const codigoactivacion As String = "VIP2796"
    Public Const balizas As String = "PLANOS_Y_BALIZAMIENTOS"
    Public sep As String = IO.Path.DirectorySeparatorChar
    Public nSel As String = "Temporal"          ' Nombre fijo de el SelectionSet.
    Public PreEtiquetas As String() = {"AC", "AD", "EV", "EX", "OB", "PR", "SIA"}
    ' CAMINOS
    Public appPath As String = System.Reflection.Assembly.GetExecutingAssembly.Location

    Public appDir As String = IO.Path.GetDirectoryName(appPath)
    Public appFile As String = IO.Path.GetFileName(appPath)
    Public appName As String = IO.Path.GetFileNameWithoutExtension(appPath)
    Public appXLSX As String = IO.Path.ChangeExtension(appPath, ".xlsx")

    'Public appSDF As String = IO.Path.ChangeExtension(appPath, ".sdf")
    '
    ' VARIABLES
    'Public bloqueEditar As String = ""      ' Nombre del bloque a cambiar.
    'Public bloqueID As Long          ' Object ID del bloque a cambiar.
    '' escala=1 (si todo está en mm) / escala=0.1 (si todo está en cm) / escala=0.01 (Si todo esta en m)
    Public escalaTotal As Double = 0.02     '' Con esta variable escalaremos bloques y texto

    'Public dirApp As String
    'Public dirBase As String            '' Directorio Base de bloques C:\ProgramData\IMPLACAD  (Ponemos la barra al final)
    Public webActualiza As String       '' Direccion Web del directorio de descarga.

    Public vermensajes As Boolean = True
    Public nombreviejo As String = ""
    Public Tipo As String = "*"
    Public Tipo1 As String = "*"
    Public Tipo2 As String = "*"
    Public Tipo3 As String = "*"
    Public app_procesointerno As Boolean = False 'afleta

    ''
    Public Function INICargar(Optional ByVal nombreINI As String = "") As String
        If nombreINI = "" Then nombreINI = nIni
        Call Utilidades.PermisosTodo()
        Dim mensaje As String = ""
        '; Este es un archivo de configuracion de ejemplo
        '; Los comentarios comienzan con ';'
        '[OPCIONES]
        ''dirBase=C:\ProgramData\IMPLACAD
        ''webActualiza=http://www.impla.es/IMPLACAD/ACTUALIZACIONES/
        ''log = 1
        ''Tipo=*
        ''Tipo1=*
        ''Tipo2=*
        ''preEti=EV,EX,SIA,KIT
        ''
        '' RELLENAR VARIABLES SIMPLES
        IMPLACAD_DATA = cIni.IniGet(nombreINI, "OPCIONES", "IMPLACAD_DATA")        ' Directorio base por defecto
        'If dirBase.EndsWith("\") = False Then dirBase &= "\"
        Try
            If IO.Directory.Exists(IMPLACAD_DATA) = False Then
                IO.Directory.CreateDirectory(IMPLACAD_DATA)
            End If
            Call PermisosTodoCarpeta(IMPLACAD_DATA)
        Catch ex As System.Exception
            ''
        End Try
        ''
        arrPaths = New List(Of String)
        arrPaths.Add(IMPLACAD_DATA) : arrPaths.Add(IMPLACAD_BUNDLE)
        webActualiza = cIni.IniGet(nombreINI, "OPCIONES", "webActualiza")        ' Enlace para actualizar BD
        If webActualiza.EndsWith("/") = False Then webActualiza &= "/"
        log = IIf(cIni.IniGet(nombreINI, "OPCIONES", "log") = "1", True, False)         '' Fichero log para control errores (Si o No)
        Tipo = cIni.IniGet(nombreINI, "OPCIONES", "Tipo")        ' Tipo
        Tipo1 = cIni.IniGet(nombreINI, "OPCIONES", "Tipo1")        ' Tipo
        Tipo2 = cIni.IniGet(nombreINI, "OPCIONES", "Tipo2")        ' Tipo
        Tipo3 = cIni.IniGet(nombreINI, "OPCIONES", "Tipo3")        ' Tipo
        ''
        arrpreEti = New ArrayList
        For Each preEti As String In cIni.IniGet(nombreINI, "OPCIONES", "preEti").Split(",")
            If arrpreEti.Contains(preEti) = False Then arrpreEti.Add(preEti)
        Next
        '
        My.Computer.FileSystem.CurrentDirectory = dirApp
        '
        Dim btnTodos As String() = New String() {
            "IMPLACADMENU", "INSERTAREDITAR", "ADECUA", "BALIZARSUELO", "BALIZARPARED",
            "BALIZARESCALERA", "RUTAEVACUACION", "TABLADATOS", "TABLAESCALERAS",
            "CAPASCOBERTURA", "GROSORLINEAS", "ESCALAM", "IMPRIMIRINS", "IMPRIMIREVA",
            "EXPLOTAEVA", "TABLAPARCIAL", "_ETRANSMIT...", "ACTUALIZARIMPLACAD", "RESOLVERREFX"}
        arrBotones = New ArrayList(btnTodos)
        colBotones = New Hashtable
        For Each nBoton As String In btnTodos
            colBotones.Add(nBoton, True)
        Next
        ' Por si no existe IMPLACAD.xlsx
        If IO.File.Exists(appXLSX) = False Then
            Dim ressourceName = My.Application.Info.AssemblyName & "." & IO.Path.GetFileName(appXLSX)   ' "Dwf3D2aCAD.AdskAssetViewer.dll"
            Using stream = System.Reflection.Assembly.GetExecutingAssembly().GetManifestResourceStream(ressourceName)
                Dim assemblyData(CInt(stream.Length) - 1) As Byte
                stream.Read(assemblyData, 0, assemblyData.Length)
                'If IO.File.Exists(appXLSX) = False Then
                IO.File.WriteAllBytes(appXLSX, assemblyData)
                'End If
                ' Si es una DLL, hay que cargarla.
                'Return System.Reflection.Assembly.LoadFrom(appXLSX)
            End Using
        End If
        INICargar = mensaje
    End Function

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
        If IO.File.Exists(nImp) = False Then
            resultado = Estado.ActiNo
            Exit Function
        Else
            Dim quecodigo As String = IO.File.ReadAllText(nImp)
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
    Public Function ImplacadActualizado(Optional conMensajes As Boolean = True) As Boolean
        If oApp Is Nothing Then _
        oApp = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
        ''
        If oApp.Documents.Count = 0 Then Exit Function
        ''
        '' Comprobar si se ha actualizado la aplicación y se tienen los recursos principales.
        Dim ficheropri As String = "C:\ProgramData\IMPLACAD\IMPLACADDATOS.zip"
        If IO.File.Exists(ficheropri) = False Then
            MsgBox("No existe fichero '" & ficheropri & "'" & vbCrLf & vbCrLf &
                   "Actualice la aplicación desde ACTUALIZARIMPLACAD", MsgBoxStyle.Critical)
            Return False
            Exit Function
        End If
        ''
        Dim nCarpetasAct As Integer = IO.Directory.GetDirectories(IO.Path.GetDirectoryName(ficheropri), "*.*", SearchOption.AllDirectories).Length
        Dim nFicherosActDwg As Integer = IO.Directory.GetFiles(IO.Path.GetDirectoryName(ficheropri), "*.dwg", SearchOption.AllDirectories).Length
        Dim nFicherosActPng As Integer = IO.Directory.GetFiles(IO.Path.GetDirectoryName(ficheropri), "*.png", SearchOption.AllDirectories).Length
        ''
        Dim nCarpetasMin As Integer = FicheroZIP_DameCantidadFicheros(ficheropri, QueTipo.Carpetas)
        Dim nFicherosMinDwg As Integer = FicheroZIP_DameCantidadFicheros(ficheropri, QueTipo.FicherosDWG)
        Dim nFicherosMinPng As Integer = FicheroZIP_DameCantidadFicheros(ficheropri, QueTipo.FicherosPNG)

        ''
        If (nFicherosActDwg < nFicherosMinDwg) Or (nFicherosActPng < nFicherosMinPng) Or (nCarpetasAct < nCarpetasMin) Then
            MsgBox("Faltan recurso en sub-carpetas de '" & IO.Path.GetDirectoryName(ficheropri) & "'" & vbCrLf & vbCrLf &
                   "Actualice la aplicación desde ACTUALIZARIMPLACAD para tenerlos", MsgBoxStyle.Critical)
            Return False
            Exit Function
        End If
        ''
        Return True
    End Function

    ''
    Public Function ImplacadEscalaM(Optional conMensajes As Boolean = True) As Boolean
        If oApp Is Nothing Then _
        oApp = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
        ''
        If oApp.Documents.Count = 0 Then Exit Function
        Dim resultado As Boolean = False
        ''
        '' Si no se ejecuta ESCALAM (Crea y rellena la propiead del dibujo ESCALAM) No continuamos.
        If PropiedadLee("ESCALAM") <> "" Then
            resultado = True
        Else
            MsgBox("¡¡ ESCALAM no se ha ejecutado aún !!" & vbCrLf & vbCrLf &
                   "Use primero utilidad ESCALAM para poner la escala en metros y guarde el dibujo.", MsgBoxStyle.Critical)
            resultado = False
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
        If IO.File.Exists(nImp) = False Then
            If conMensajes Then
                MsgBox("¡¡ IMPLACAD no está activado !!" & vbCrLf & vbCrLf &
                       "Introduzca el código de activación desde ACTUALIZAIMPLACAD.", MsgBoxStyle.Critical)
            End If
            resultado = False
        Else
            Dim quecodigo As String = IO.File.ReadAllText(nImp)
            If quecodigo.ToUpper = codigoactivacion.ToUpper Then
                resultado = True
            Else
                If conMensajes Then
                    MsgBox("¡¡ IMPLACAD no está activado !!" & vbCrLf & vbCrLf &
                           "Introduzca el código de activación desde ACTUALIZAIMPLACAD.", MsgBoxStyle.Critical)
                End If
                resultado = False
            End If
        End If
        ''
        Return resultado
    End Function

    ''
    Public Function ImplacadGuardado() As Boolean
        Dim resultado As Boolean = False
        If oApp.ActiveDocument.FullName = "" Then
            MsgBox("Este documento aún no se ha guardado por primera vez" & vbCrLf & vbCrLf &
                   "Guarde antes el documento y vuelva a probar...")
            resultado = False
        ElseIf (oApp.ActiveDocument.FullName <> "" And oApp.ActiveDocument.Saved = False) Then
            oApp.ActiveDocument.Save()
            resultado = True
        ElseIf (oApp.ActiveDocument.FullName <> "" And oApp.ActiveDocument.Saved = True) Then
            resultado = True
        End If
        ''
        Return resultado
    End Function

    ''
    Public Sub ImplacadDibujoCierra(queFull As String)
        Dim oDoc As Autodesk.AutoCAD.Interop.AcadDocument = Nothing
        ''
        For Each oDoc In oApp.Documents
            If oDoc.FullName = queFull Then
                On Error Resume Next
                oDoc.Close(False)
                Exit For
            End If
        Next
    End Sub

    ''
    Public Sub ImplacadActivaDocumento(queFull As String,
                                       Optional activar As Boolean = True,
                                       Optional guardar As Boolean = False)
        Dim oDoc As Autodesk.AutoCAD.Interop.AcadDocument = Nothing
        ''
        For Each oDoc In oApp.Documents
            If oDoc.FullName = queFull Then
                oDoc.Activate()
                If guardar = True Then oDoc.Save()
                Exit For
            End If
        Next
    End Sub

    ''
    Public Sub ImplacadGuardaDocumentos()
        Dim oDoc As Autodesk.AutoCAD.Interop.AcadDocument = Nothing
        ''
        For Each oDoc In oApp.Documents
            oDoc.Save()
        Next
    End Sub

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