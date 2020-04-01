Imports System.Linq
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Interop
Imports Autodesk.AutoCAD.Interop.Common

Module modCapas
    ' NOMBRE CAPAS
    Public Const NBalizamientoSuelo As String = "BALIZAMIENTO SUELO"
    Public Const NBalizamientoPared As String = "BALIZAMIENTO PARED"
    Public Const NBalizamientoEscalera As String = "BALIZAMIENTO ESCALERA"
    Public Const NRutaEvacuacion1 As String = "Ruta evacuación primaria"
    Public Const NCoverturaEvacuacion As String = "Zona Cobertura Evacuación"
    Public Const NCoverturaExtincion As String = "Zona Cobertura Extincion"
    Public Const NZonas As String = "Zonas"
    Public Const NZonasTexto As String = "Zonas_Texto"
    Public Const NTablas As String = "Tablas"
    '
    Public Sub CapaCeroActiva()
        '' Pondremos la capa 0 como activa.
        If oApp Is Nothing Then _
        oApp = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
        '' Activar Capa 0 primero.
        oApp.ActiveDocument.ActiveLayer = oApp.ActiveDocument.Layers.Item("0")
        ''
    End Sub
    Public Sub CapaCreaActivaZonas(Optional activar As Boolean = True)
        If oApp Is Nothing Then _
            oApp = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
        '' Activar Capa 0 primero.
        'oApp.ActiveDocument.ActiveLayer = oApp.ActiveDocument.Layers.Item("0")
        ''
        '' Coger la capa BALIZAMIENTO SUELO o crearla
        Dim oLayer As AcadLayer = Nothing
        Try
            oLayer = oApp.ActiveDocument.Layers.Item(NZonas)
        Catch ex As System.Exception
            oLayer = oApp.ActiveDocument.Layers.Add(NZonas)
        End Try
        ''
        '' Poner la capa como visible (LayerOn=True) y Reutilizada (Freeze=False)
        oLayer.LayerOn = True
        oLayer.Freeze = False
        ''
        '' Poner alguna de sus características
        Dim oColor As New AcadAcCmColor
        oColor.ColorIndex = 7
        oLayer.TrueColor = oColor
        oLayer.Lineweight = ACAD_LWEIGHT.acLnWt025
        ''
        If activar Then oApp.ActiveDocument.ActiveLayer = oLayer    ' oApp.ActiveDocument.Layers.Item("BALIZAMIENTO SUELO")
        oLayer = Nothing
        oColor = Nothing
        '
        'oApp.ActiveDocument.Regen(AcRegenType.acAllViewports)
    End Sub
    Public Sub CapaCreaActivaBalizamientoSuelo(Optional activar As Boolean = True)
        ''
        '' Crear una capa y poner sus características.
        ''
        If oApp Is Nothing Then _
            oApp = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
        '' Activar Capa 0 primero.
        oApp.ActiveDocument.ActiveLayer = oApp.ActiveDocument.Layers.Item("0")
        ''
        '' Coger la capa BALIZAMIENTO SUELO o crearla
        Dim oLayer As AcadLayer = Nothing
        Try
            oLayer = oApp.ActiveDocument.Layers.Item(NBalizamientoSuelo)
        Catch ex As System.Exception
            oLayer = oApp.ActiveDocument.Layers.Add(NBalizamientoSuelo)
        End Try
        ''
        '' Poner la capa como visible (LayerOn=True) y Reutilizada (Freeze=False)
        oLayer.LayerOn = True
        oLayer.Freeze = False
        ''
        '' Poner alguna de sus características
        Dim oColor As New AcadAcCmColor
        oColor.ColorIndex = 95
        oLayer.TrueColor = oColor
        oLayer.Lineweight = ACAD_LWEIGHT.acLnWt140
        ''
        If activar Then oApp.ActiveDocument.ActiveLayer = oLayer    ' oApp.ActiveDocument.Layers.Item("BALIZAMIENTO SUELO")
        oLayer = Nothing
        oColor = Nothing
        ''
        oApp.ActiveDocument.Regen(AcRegenType.acAllViewports)
    End Sub
    Public Sub CapaCreaActivaRutaEvacuacion(queTipo As TipoEvacuacion)
        ''
        '' Crear una capa y poner sus características.
        ''
        If oApp Is Nothing Then _
            oApp = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
        '' Activar Capa 0 primero.
        'oApp.ActiveDocument.ActiveLayer = oApp.ActiveDocument.Layers.Item("0")
        ''
        '' Coger la capa BALIZAMIENTO SUELO o crearla
        Dim oLayer As AcadLayer = Nothing
        Dim tipolinea As String = ""
        Dim ficherolinea As String = dirApp & "rutasevacuacion.lin"
        If queTipo = TipoEvacuacion.Primaria Then
            Try
                oLayer = oApp.ActiveDocument.Layers.Item(NRutaEvacuacion1)
            Catch ex As System.Exception
                oLayer = oApp.ActiveDocument.Layers.Add(NRutaEvacuacion1)
            End Try
            'Dim oBl As AcadBlockReference = oApp.ActiveDocument.ModelSpace.InsertBlock( _
            ' CType(oPt, Point3d).ToArray, _
            'dirApp & "..\..\Resources\etiquetaescalera.dwg", _
            '1, 1, 1, 0)
            tipolinea = "EVACUACIONPRI"
        ElseIf queTipo = TipoEvacuacion.Accesibilidad Then
            Try
                oLayer = oApp.ActiveDocument.Layers.Item("Ruta evacuación accesibilidad")
            Catch ex As System.Exception
                oLayer = oApp.ActiveDocument.Layers.Add("Ruta evacuación accesibilidad")
            End Try
            tipolinea = "EVACUACIONPRI"
        End If
        ''
        '' Poner la capa como visible (LayerOn=True) y Reutilizada (Freeze=False)
        oLayer.LayerOn = True
        oLayer.Freeze = False
        ''
        '' Poner alguna de sus características
        Dim oColor As New AcadAcCmColor
        If queTipo = TipoEvacuacion.Primaria Then
            oColor.ColorIndex = 104
        ElseIf queTipo = TipoEvacuacion.Accesibilidad Then
            oColor.SetRGB(0, 51, 171)
        End If
        'oColor.ColorIndex = RGB(0, 51, 171)
        oLayer.TrueColor = oColor
        oLayer.Lineweight = ACAD_LWEIGHT.acLnWt030
        '' Cargar el tipo de linea de la capa
        Try
            '' Cargamos el fichero de tipos de lineas. Si ya existe error.
            oApp.ActiveDocument.Linetypes.Load(tipolinea, ficherolinea)
        Catch ex As System.Exception
            '' Ya existe el tipo de linea cargado.
            'MsgBox(ex.Message)
        End Try
        '' Poner el tipo de linea correcto.
        oLayer.Linetype = tipolinea
        ''
        oApp.ActiveDocument.ActiveLayer = oLayer    ' oApp.ActiveDocument.Layers.Item("BALIZAMIENTO SUELO")
        oLayer = Nothing
        oColor = Nothing
        ''
        'oApp.Preferences.User.ADCInsertUnitsDefaultSource = AcInsertUnits.acInsertUnitsMillimeters
        'oApp.Preferences.User.ADCInsertUnitsDefaultTarget = AcInsertUnits.acInsertUnitsMeters
        oApp.ActiveDocument.Regen(AcRegenType.acAllViewports)
    End Sub
    Public Sub CapaCreaActivaTablaDatos(Optional activar As Boolean = True)
        '(SETQ capacuadro "CUADRO")
        '(IF   (NULL (TBLOBJNAME "LAYER" capacuadro))
        '(ENTMAKE (LIST '(0 . "LAYER")'(100 . "AcDbSymbolTableRecord")'(100 . "AcDbLayerTableRecord")
        '                     (CONS 2 capacuadro)'(70 . 0)(CONS 62 7) (CONS 370 25)))
        ' )
        '' 70 = reutilizar 0
        '' 62 = Color 7
        '' 370 = ACAD_LWEIGHT.acLnWt25 'grosor 2 mm
        ''
        '' Crear una capa y poner sus características.
        ''
        If oApp Is Nothing Then _
            oApp = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
        '' Activar Capa 0 primero.
        'oApp.ActiveDocument.ActiveLayer = oApp.ActiveDocument.Layers.Item("0")
        ''
        '' Coger la capa BALIZAMIENTO SUELO o crearla
        Dim oLayer As AcadLayer = Nothing
        Try
            oLayer = oApp.ActiveDocument.Layers.Item(NTablas)
        Catch ex As System.Exception
            oLayer = oApp.ActiveDocument.Layers.Add(NTablas)
        End Try
        ''
        '' Poner la capa como visible (LayerOn=True) y Reutilizada (Freeze=False)
        oLayer.LayerOn = True
        oLayer.Freeze = False
        ''
        '' Poner alguna de sus características
        Dim oColor As New AcadAcCmColor
        oColor.ColorIndex = 7
        oLayer.TrueColor = oColor
        oLayer.Lineweight = ACAD_LWEIGHT.acLnWt025
        ''
        If activar Then oApp.ActiveDocument.ActiveLayer = oLayer    ' oApp.ActiveDocument.Layers.Item("BALIZAMIENTO SUELO")
        oLayer = Nothing
        oColor = Nothing
        ''
        oApp.ActiveDocument.Regen(AcRegenType.acAllViewports)
    End Sub
    Public Sub CapaCreaActivaTablaZonas(Optional activar As Boolean = True)
        '(SETQ capacuadro "CUADRO")
        '(IF   (NULL (TBLOBJNAME "LAYER" capacuadro))
        '(ENTMAKE (LIST '(0 . "LAYER")'(100 . "AcDbSymbolTableRecord")'(100 . "AcDbLayerTableRecord")
        '                     (CONS 2 capacuadro)'(70 . 0)(CONS 62 7) (CONS 370 25)))
        ' )
        '' 70 = reutilizar 0
        '' 62 = Color 7
        '' 370 = ACAD_LWEIGHT.acLnWt25 'grosor 2 mm
        ''
        '' Crear una capa y poner sus características.
        ''
        If oApp Is Nothing Then _
            oApp = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
        '' Activar Capa 0 primero.
        'oApp.ActiveDocument.ActiveLayer = oApp.ActiveDocument.Layers.Item("0")
        ''
        '' Coger la capa BALIZAMIENTO SUELO o crearla
        Dim oLayer As AcadLayer = Nothing
        Try
            oLayer = oApp.ActiveDocument.Layers.Item(NTablas)
        Catch ex As System.Exception
            oLayer = oApp.ActiveDocument.Layers.Add(NTablas)
        End Try
        ''
        '' Poner la capa como visible (LayerOn=True) y Reutilizada (Freeze=False)
        oLayer.LayerOn = True
        oLayer.Freeze = False
        ''
        '' Poner alguna de sus características
        Dim oColor As New AcadAcCmColor
        oColor.ColorIndex = 7
        oLayer.TrueColor = oColor
        oLayer.Lineweight = ACAD_LWEIGHT.acLnWt025
        ''
        If activar Then oApp.ActiveDocument.ActiveLayer = oLayer
        oLayer = Nothing
        oColor = Nothing
        ''
        oApp.ActiveDocument.Regen(AcRegenType.acAllViewports)
    End Sub
    Public Sub CapaCreaActivaBalizamientoPared(Optional activar As Boolean = True)
        ' (DEFUN c:BalizarPared ( / )
        '  (SETQ capabalizar1 "BALIZAMIENTO PARED")
        '   (IF   (NULL (TBLOBJNAME "LAYER" capabalizar1))
        '   (ENTMAKE (LIST '(0 . "LAYER")'(100 . "AcDbSymbolTableRecord")'(100 . "AcDbLayerTableRecord")
        '                        (CONS 2 capabalizar1)'(70 . 0)(CONS 62 52) (CONS 370 100)))
        '     );If
        '(SETVAR "clayer" capabalizar1)
        '  (COMMAND "_line")
        '  (PRINC)
        ')
        '' 70 = reutilizar 0
        '' 62 = Color 52
        '' 370 = ACAD_LWEIGHT.acLnWt100 'grosor 2 mm
        ''
        '' Crear una capa y poner sus características.
        ''
        If oApp Is Nothing Then _
        oApp = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
        '' Activar Capa 0 primero.
        'oApp.ActiveDocument.ActiveLayer = oApp.ActiveDocument.Layers.Item("0")
        ''
        '' Coger la capa BALIZAMIENTO SUELO o crearla
        Dim oLayer As AcadLayer = Nothing
        Try
            oLayer = oApp.ActiveDocument.Layers.Item(NBalizamientoPared)
        Catch ex As System.Exception
            oLayer = oApp.ActiveDocument.Layers.Add(NBalizamientoPared)
        End Try
        ''
        '' Poner la capa como visible (LayerOn=True) y Reutilizada (Freeze=False)
        oLayer.LayerOn = True
        oLayer.Freeze = False
        ''
        '' Poner alguna de sus características
        Dim oColor As New AcadAcCmColor
        oColor.ColorIndex = 52
        oLayer.TrueColor = oColor
        oLayer.Lineweight = ACAD_LWEIGHT.acLnWt100
        ''
        If activar Then oApp.ActiveDocument.ActiveLayer = oLayer    ' oApp.ActiveDocument.Layers.Item("BALIZAMIENTO PARED")
        oLayer = Nothing
        oColor = Nothing
        ''oApp.ActiveDocument.SendCommand("_line ")
        oApp.ActiveDocument.Regen(AcRegenType.acAllViewports)
    End Sub
    Public Sub CapaCreaActivaBalizamientoEscalera(Optional activar As Boolean = True)
        ' (SETQ capabalizarescalera "BALIZAMIENTO ESCALERA")
        '   (IF   (NULL (TBLOBJNAME "LAYER" capabalizarescalera))
        '   (ENTMAKE (LIST '(0 . "LAYER")'(100 . "AcDbSymbolTableRecord")'(100 . "AcDbLayerTableRecord")
        '                        (CONS 2 capabalizarescalera)'(70 . 0)(CONS 62 68) (CONS 370 100)))
        '     );If
        '(SETQ capacuadro "CUADRO")
        '   (IF   (NULL (TBLOBJNAME "LAYER" capacuadro))
        '   (ENTMAKE (LIST '(0 . "LAYER")'(100 . "AcDbSymbolTableRecord")'(100 . "AcDbLayerTableRecord")
        '                        (CONS 2 capacuadro)'(70 . 0)(CONS 62 7) (CONS 370 25)))
        '    )
        '' 70 = reutilizar 0
        '' 62 = Color 68
        '' 370 = ACAD_LWEIGHT.acLnWt100
        ''
        '' Crear una capa y poner sus características.
        ''
        If oApp Is Nothing Then _
        oApp = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
        '' Activar Capa 0 primero.
        'CapaCeroActiva()
        ''
        '' Coger la capa BALIZAMIENTO SUELO o crearla
        Dim oLayer As AcadLayer = Nothing
        Try
            oLayer = oApp.ActiveDocument.Layers.Item(NBalizamientoEscalera)
        Catch ex As System.Exception
            oLayer = oApp.ActiveDocument.Layers.Add(NBalizamientoEscalera)
        End Try
        ''
        '' Poner la capa como visible (LayerOn=True) y Reutilizada (Freeze=False)
        oLayer.LayerOn = True
        oLayer.Freeze = False
        ''
        '' Poner alguna de sus características
        Dim oColor As New AcadAcCmColor
        oColor.ColorIndex = 7       ' Antes era 68, pero no se veían las lineas que dibujamos para anchura.
        oLayer.TrueColor = oColor
        oLayer.Lineweight = ACAD_LWEIGHT.acLnWt100
        ''
        If activar Then oApp.ActiveDocument.ActiveLayer = oLayer    ' oApp.ActiveDocument.Layers.Item("BALIZAMIENTO PARED")
        oLayer = Nothing
        oColor = Nothing
        ''oApp.ActiveDocument.SendCommand("_line ")
        oApp.ActiveDocument.Regen(AcRegenType.acAllViewports)
    End Sub
    Public Sub CapaZonaCoberturaACTDES(Optional estado As CapaEstado = CapaEstado.Inverso)
        If oApp Is Nothing Then _
        oApp = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
        Dim oColor As New AcadAcCmColor
        oColor.ColorMethod = AcColorMethod.acColorMethodByACI
        ''
        '' Configurar la capa de cobertura (0=desactivada / 1=activada / 2=Not(estado actual)
        Dim oLayer As AcadLayer = Nothing
        ''
        '' ** Capa "Zona Cobertura Evacuación"
        Try
            oLayer = oApp.ActiveDocument.Layers.Item(NCoverturaEvacuacion)
        Catch ex As System.Exception
            '' Si no existia, la creamos.
            oLayer = oApp.ActiveDocument.Layers.Add(NCoverturaEvacuacion)
            oColor.ColorIndex = 104
            oLayer.TrueColor = oColor
        End Try
        ''
        If estado = CapaEstado.Desactivar Then
            oLayer.LayerOn = False ' True=Activar  / False=DesActivar
            oLayer.Freeze = True ' True=Inutilizar  / False=Reutilizar
        ElseIf estado = CapaEstado.Activar Then
            oLayer.LayerOn = True
            oLayer.Freeze = False
        ElseIf estado = CapaEstado.Inverso Then
            oLayer.LayerOn = Not (oLayer.LayerOn)
            oLayer.Freeze = Not (oLayer.Freeze)
        End If
        ''
        '' ** Capa "Zona Cobertura Extincion"
        Try
            oLayer = oApp.ActiveDocument.Layers.Item(NCoverturaExtincion)
        Catch ex As System.Exception
            '' Si no existia, la creamos.
            oLayer = oApp.ActiveDocument.Layers.Add(NCoverturaExtincion)
            oColor.ColorIndex = 1
            oLayer.TrueColor = oColor
        End Try
        ''
        If estado = CapaEstado.Desactivar Then
            oLayer.LayerOn = False ' True=Activar  / False=DesActivar
            oLayer.Freeze = True ' True=Inutilizar  / False=Reutilizar
        ElseIf estado = CapaEstado.Activar Then
            oLayer.LayerOn = True
            oLayer.Freeze = False
        ElseIf estado = CapaEstado.Inverso Then
            oLayer.LayerOn = Not (oLayer.LayerOn)
            oLayer.Freeze = Not (oLayer.Freeze)
        End If
        ''
        oLayer = Nothing
        oApp.ActiveDocument.Regen(AcRegenType.acAllViewports)
    End Sub
End Module
