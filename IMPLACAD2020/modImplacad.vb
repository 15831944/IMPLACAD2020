Imports System
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.Interop
Imports Autodesk.AutoCAD.Interop.Common
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Ribbon
Imports System.Linq
Imports System.Collections.Generic


Module modImplacad
    ''
    Public WithEvents oTimer As System.Timers.Timer
    ''
    '' Configurar Dibujo Actual
    Public Sub ConfiguraDibujo()
        If oApp Is Nothing Then _
        oApp = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
        'oDoc = oApp.ActiveDocument
        ''
        Dim oPrefApp As AcadPreferences = oApp.Preferences
        ''
        '' Poner las unidades de contenido origen y de dibujo destino (Opciones-Preferencias de usuario)
        oPrefApp.User.ADCInsertUnitsDefaultSource = AcInsertUnits.acInsertUnitsUnitless
        'oPrefApp.User.ADCInsertUnitsDefaultSource = AcInsertUnits.acInsertUnitsMillimeters
        'oPrefApp.User.ADCInsertUnitsDefaultSource = AcInsertUnits.acInsertUnitsMeters
        oPrefApp.User.ADCInsertUnitsDefaultTarget = AcInsertUnits.acInsertUnitsMeters
        oPrefApp = Nothing
        ''
        'Dim oPrefDoc As AcadDatabasePreferences = oDoc.Preferences
        'oPrefDoc = Nothing
        '' Configurar unidades del dibujo y escala de inserción Sin unidad
        '' 0= Sin unidad, 4 = milimetros, 5 = centimetros, 6 = metros
        'Application.DocumentManager.MdiActiveDocument.Database.Insunits = UnitsValue.Meters
        Application.DocumentManager.MdiActiveDocument.Database.Insunits = UnitsValue.Undefined
        'oDoc.SendCommand("INSUNITS" & vbCr & "4" & vbCr)
        ''
        '' Crear estilo RRC_ARIAL
        'Dim estilotexto As String = "RRC_arial"
        ''
        'If textoEstilo Is Nothing Then
        Try
            textoEstilo = oApp.ActiveDocument.TextStyles.Add(estilotexto)
        Catch ex As System.Exception
            textoEstilo = oApp.ActiveDocument.TextStyles.Item(estilotexto)
        End Try
        ''
        Try
            textoEstilo.SetFont("ARIAL", False, False, 0, 34)
        Catch ex As System.Exception
            MsgBox(ex.Message)
        End Try
        'End If
        ''
        '' Para que no se muestren los marcos de las imágenes.
        Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("IMAGEFRAME", 1)
        oApp.ActiveDocument.Regen(AcRegenType.acActiveViewport)
        Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("IMAGEFRAME", 0)
        '' Poner la escala de anotación 1:1
        Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("CANNOSCALE", "1:1")
        '' Poner la escala de los tipos de linea a 1
        Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("LTSCALE", 1)
        oApp.ActiveDocument.Regen(AcRegenType.acActiveViewport)
        'CANNOSCALE
        'Call DameDependencias()
        'XrefComprueba()
        'XrefImagenDame()
    End Sub
    ''
    Public Function GetAcadPath() As String
        Dim resultado As String = ""
        'Dim prodId As String = Autodesk.AutoCAD.Runtime.SystemObjects.DynamicLinker.ProductLcid.ToString
        'Dim strPath As String = My.Computer.Registry.LocalMachine.OpenSubKey(prodId).GetValue("AcadLocation")
        'Return strPath
        'Return Autodesk.AutoCAD.Runtime.Utilities.
        '' Buscamos en diferentes ubicaciones, según la versión de AutoCAD 2020 instalada.
        ''
        '' AutoCAD 2020 o AutoCAD Mechanical 2020 o AutoCAD Architecture 2020
        Dim rk As Microsoft.Win32.RegistryKey
        Dim arrRks As String() = New String() {
            "SOFTWARE\Autodesk\AutoCAD\R23.1\ACAD-3001:40A",
        "SOFTWARE\Autodesk\AutoCAD\R23.1\ACAD-3004:40A",
        "SOFTWARE\Autodesk\AutoCAD\R23.1\ACAD-3005:40A"}
        '"HKEY_LOCAL_MACHINE\SOFTWARE\Autodesk\AutoCAD\R20.1\ACAD-F005:40A"
        Try
            rk = My.Computer.Registry.LocalMachine.OpenSubKey("SOFTWARE\Autodesk\AutoCAD\R23.1\ACAD-3001", False)
            resultado = rk.GetValue("GlobUPILocation")
        Catch ex As System.Exception
            For Each queRk As String In arrRks
                Try
                    rk = My.Computer.Registry.LocalMachine.OpenSubKey(queRk)
                    resultado = rk.GetValue("AcadLocation")
                    Exit For
                Catch ex1 As System.Exception
                    resultado = ""
                    Continue For
                End Try
            Next
        End Try
        '' 
        Return resultado
    End Function
    ''
    Public Sub PonTRUSTEDPATHS()
        Dim str_TR As String = Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable("TRUSTEDPATHS")
        str_TR.Replace(";;", ";")
        str_TR.Replace(";;", ";")
        If str_TR = "." Then str_TR = ""
        ''
        Dim C_Paths As String = LCase(str_TR)

        Dim Old_Path_Ary As List(Of String) = New List(Of String)
        Old_Path_Ary.AddRange(C_Paths.Split(";"))

        Dim New_paths As List(Of String) = New List(Of String)

        New_paths.Add(My.Application.Info.DirectoryPath)
        New_paths.Add(IMPLACAD_DATA)

        For Each Str As String In New_paths
            If Not Old_Path_Ary.Contains(LCase(Str)) Then
                Old_Path_Ary.Add(Str)
            End If
        Next
        '' Quitar los que están vacios en 
        For x As Integer = Old_Path_Ary.Count - 1 To 0 Step -1
            If Old_Path_Ary(x) = "" Then
                Old_Path_Ary.RemoveAt(x)
            End If
        Next
        ''
        Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("TRUSTEDPATHS", String.Join(";", Old_Path_Ary.ToArray()))
    End Sub
    ''
    Public Function EsParaTrabajar() As Boolean
        Dim queValor As Object = Nothing
        Try
            '' Leemos el valor de estado.
            oApp.ActiveDocument.SummaryInfo.GetCustomByKey("ESTADO", queValor)
        Catch ex As System.Exception
            Return True
            Exit Function
        End Try
        ''
        If IsNothing(queValor) Then
            Return True
        Else
            If queValor.ToString = "PARAIMPRIMIR" Then
                Return False
            Else
                Return True
            End If
        End If
        '' Ponemos la propiedad personalizada ESTADO=PARAIMPRIMIR
        'Try
        '    oApp.ActiveDocument.SummaryInfo.AddCustomInfo("ESTADO", "PARAIMPRIMIR")
        'Catch ex As System.Exception
        '    oApp.ActiveDocument.SummaryInfo.SetCustomByKey("ESTADO", "PARAIMPRIMIR")
        'End Try
        ''
    End Function
    ''
    Public Function PropiedadLee(quePro As String) As String
        Dim queValor As String = ""
        Try
            '' Leemos el valor de ESCALAM
            oApp.ActiveDocument.SummaryInfo.GetCustomByKey(quePro, queValor)
        Catch ex As System.Exception
            queValor = ""
        End Try
        ''
        Return queValor
    End Function
    ''
    Public Sub PropiedadEscribe(quePro As String, queVal As String)
        Dim valoractual As String = ""
        Try
            '' Si no existía, la creamos
            oApp.ActiveDocument.SummaryInfo.AddCustomInfo(quePro, queVal)
        Catch ex As System.Exception
            '' Leemos el valor de estado.
            oApp.ActiveDocument.SummaryInfo.GetCustomByKey(quePro, valoractual)
            If valoractual <> queVal Then
                oApp.ActiveDocument.SummaryInfo.SetCustomByKey(quePro, queVal)
            End If
        End Try
        ''
    End Sub
    ''
#Region "TABLADATOS"
    ''
    '' Insertar tabla para datos de etiquetas o la actualiza si ya existe.
    Public Sub TablaDatosInserta()
        ConfiguraDibujo()
        ''
        '' Comprobar si ya existe la tabla de datos
        Dim TablaDatosTemp As AcadTable = DameTablaDatos()
        ''
        '' Si no existe, la creamos. Y si existe la actualizamos.
        If TablaDatosTemp Is Nothing And vermensajes = True Then
            If oApp Is Nothing Then _
                    oApp = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
            'oDoc = oApp.ActiveDocument
            oApp.ActiveDocument.Activate()
            Dim oPunto As Object = Nothing
            Try
                'oPunto = PuntoDame_NET("Punto de inserción de la tabla:")
                oPunto = oApp.ActiveDocument.Utility.GetPoint(, "Punto de inserción de la tabla:")
            Catch ex As System.Exception
                'MsgBox(ex.Message)
            End Try
            ''
            If oPunto Is Nothing Then
                Exit Sub
            End If
            ''
            CapaCreaActivaTablaDatos()
            ''
            '' Insertamos la tabla.
            'Dim oPunto As Object = oDoc.Utility.GetPoint(, "Punto de inserción de la Tabla")
            'TablaDatos = oDoc.ModelSpace.AddTable(oPunto, 6, 2, 50 * escala, 400 * escala)
            TablaDatos = oApp.ActiveDocument.ModelSpace.AddTable(oPunto, 6, 2, 1, 1)
            'Temporalmente desactivamos el recalculo de la tabla
            TablaDatos.RegenerateTableSuppressed = True
            ''
            '' Ponemos los textos en las cabeceras
            TablaDatos.SetText(0, 0, "IMPLACAD ES UN PRODUCTO IMPLASER")        '' Titulo tabla
            TablaDatos.SetText(1, 0, "** BALIZAMIENTO **")                  '' Cabecera
            TablaDatos.SetText(1, 1, "** METROS **")                '' Cabecera
            TablaDatos.SetText(2, 0, "SUELO")        '' 
            TablaDatos.SetText(3, 0, "PARED")        '' 
            TablaDatos.SetText(4, 0, "TOTAL SEÑALES --->")        '' Titulo tabla
            'TablaDatos.SetText(4, 1, DameTotalBloquesImplacad)        '' Titulo tabla
            TablaDatos.SetText(5, 0, "** REFERENCIA **")        '' Titulo tabla
            TablaDatos.SetText(5, 1, "** CANTIDAD **")        '' Titulo tabla
            '' Regenerar la tabla
            'TablaDatos.RegenerateTableSuppressed = False
            'TablaDatos.RecomputeTableBlock(True)
            ';(initget 6)
            ';; REGENERAMOS LA TABLA
            '(vla-put-RegenerateTableSuppressed tb :vlax-false)
            TablaDatosActualiza()
            '' Regenerar la tabla
            TablaDatos.RegenerateTableSuppressed = False
            TablaDatos.RecomputeTableBlock(True)
        Else
            TablaDatos = CType(oApp.ActiveDocument.ObjectIdToObject(TablaDatosTemp.ObjectID), AcadTable) 'TablaDatosTemp ' oDoc.ObjectIdToObject(TablaDatosTemp.ObjectID)
            'Temporalmente desactivamos el recalculo de la tabla
            TablaDatos.RegenerateTableSuppressed = True
            TablaDatosActualiza()
            '' Regenerar la tabla
            TablaDatos.RegenerateTableSuppressed = False
            TablaDatos.RecomputeTableBlock(True)
            If vermensajes = True Then MsgBox("Tabla DATOS actualizada porque ya existía...")
        End If
        TablaDatos = Nothing
        'vermensajes = True
    End Sub
    ''
    Public Sub TablaDatosActualiza()
        If colBloquesCantidad IsNot Nothing Then colBloquesCantidad.Clear()
        colBloquesCantidad = Nothing
        ''
        TablaDatos.SetText(2, 1, DameTotalBaliza(queCapa.BALIZAMIENTOSUELO))        '' Total Baliza Suelo
        TablaDatos.SetText(3, 1, DameTotalBaliza(queCapa.BALIZAMIENTOPARED))        '' Total Baliza Pared
        TablaDatos.SetText(4, 1, DameTotalBloquesImplacad)        '' Total etiquetas

        ''
        '' Ahora actualizaremos el listado de bloques y cantidades.
        '' Borrar todas las filas superiores a la 5
        If TablaDatos.Rows > 6 Then
            TablaDatos.DeleteRows(6, TablaDatos.Rows - 5)
        End If
        ''
        '' Ahora creamos las filas que sean necesarias.
        If colBloquesCantidad IsNot Nothing AndAlso colBloquesCantidad.Count > 0 Then
            Dim colIni As Integer = 6
            TablaDatos.InsertRows(colIni, 0.3, colBloquesCantidad.Count)
            ''
            For Each nBloque As String In colBloquesCantidad.Keys
                TablaDatos.SetText(colIni, 0, nBloque)        '' Nombre del bloque
                TablaDatos.SetText(colIni, 1, colBloquesCantidad(nBloque))  '' Cantidad de bloques.
                colIni += 1
            Next
            colBloquesCantidad.Clear()
            colBloquesCantidad = Nothing
        End If
        '' Formateamos la tabla.
        TablaDatos.TitleSuppressed = False              '' Activar Titulo
        TablaDatos.HeaderSuppressed = False             '' Activar Cabecera
        TablaDatos.RegenerateTableSuppressed = True     '' Regenerar/Suprimirla
        '' Margen de las celdas
        TablaDatos.VertCellMargin = 0.1                '' Margen vertical celda
        TablaDatos.HorzCellMargin = 0.25               '' Margen horizontal celda
        '' Altura de las filas de cabecera
        TablaDatos.SetRowHeight(0, 0.48)                '' Altura Fila 0
        TablaDatos.SetRowHeight(1, 0.4)                '' Altura Fila 1
        TablaDatos.SetRowHeight(2, 0.4)                '' Altura Fila 2
        TablaDatos.SetRowHeight(3, 0.4)                '' Altura Fila 3
        TablaDatos.SetRowHeight(4, 0.4)                '' Altura Fila 4
        TablaDatos.SetRowHeight(5, 0.4)                '' Altura Fila 5
        '' Anchura de las 2 columnas
        TablaDatos.SetColumnWidth(0, 3.5)               '' Ancho columna 1
        TablaDatos.SetColumnWidth(1, 3.5)               '' Ancho columna 2
        '' Estilo de texto
        TablaDatos.SetTextStyle(RowType.TitleRow, textoEstilo.Name)  '' Estilo de texto de titulo
        TablaDatos.SetTextStyle(RowType.HeaderRow, textoEstilo.Name) '' Estilo de texto de cabecera
        TablaDatos.SetTextStyle(RowType.DataRow, textoEstilo.Name)   '' Estilo de texto de datos
        '' Altura de los textos
        TablaDatos.SetTextHeight(RowType.TitleRow, 0.25)    ' 22 * escalaTotal)    '' Altura texto titulo=0.25
        TablaDatos.SetTextHeight(RowType.HeaderRow, 0.2)    ' 2 * escalaTotal)    '' Altura texto cabecera=0.2
        TablaDatos.SetTextHeight(RowType.DataRow, 0.2)      ' 2 * escalaTotal)    '' Altura texto celdas=0.2
        ''
        TablaDatos.SetAlignment(RowType.TitleRow, AcCellAlignment.acMiddleCenter)   '' Alineacion titulo
        TablaDatos.SetAlignment(RowType.HeaderRow, AcCellAlignment.acMiddleCenter)   '' Alineacion cabecera
        TablaDatos.SetAlignment(RowType.DataRow, AcCellAlignment.acMiddleCenter)   '' Alineacion datos
        ''
        '' Altura del resto de las filas con datos a 0.3
        For y As Integer = 6 To TablaDatos.Rows - 1
            TablaDatos.SetRowHeight(y, 0.3)
        Next
        ''
        '' Regenerar la tabla
        'TablaDatos.RegenerateTableSuppressed = False
        'TablaDatos.RecomputeTableBlock(True)
        XData.XNuevo(TablaDatos, "Clase=tabla")
        'TablaDatos = Nothing
    End Sub
    ''
    Public Function DameTablaDatos() As AcadTable
        Dim resultado As AcadTable = Nothing
        'Dim cSeleccion As AcadSelectionSets
        Dim F1(1) As Short
        Dim F2(1) As Object
        Dim vF1 As Object = Nothing
        Dim vF2 As Object = Nothing
        '' 
        '' Filtros para seleccionar la Tabla
        'F1(0) = 100 : F2(0) = "AcDbSymbolTableRecord"
        'F1(1) = 100 : F2(1) = "AcDbLayerTableRecord"
        F1(0) = 0 : F2(0) = "ACAD_TABLE"
        F1(1) = 1001 : F2(1) = regAPP
        ' F1(1) = 1000 : F2(1) = "Clase=tabla"
        ''
        vF1 = F1
        vF2 = F2
        ''
        If oApp Is Nothing Then _
        oApp = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
        ''
        ''
        ''
        Try
            oSelTemp = oApp.ActiveDocument.SelectionSets.Add("TEMPORAL")
        Catch ex As System.Exception
            oSelTemp = oApp.ActiveDocument.SelectionSets.Item("TEMPORAL")
        End Try
        ''
        ''
        oSelTemp.Clear()
        Try
            oSelTemp.Select(AcSelect.acSelectionSetAll, , , vF1, vF2)
        Catch ex As System.Exception
            Debug.Print(ex.Message)
        End Try
        If oSelTemp IsNot Nothing And oSelTemp.Count > 0 Then
            Try
                For Each oEnt As AcadEntity In oSelTemp
                    Dim texto As String = XData.XLeeDato(oEnt, xT.TEXTOS)
                    If TypeOf oEnt Is AcadTable And texto = "Clase=tabla" Then
                        resultado = CType(oApp.ActiveDocument.ObjectIdToObject(oEnt.ObjectID), AcadTable)
                        Exit For
                    End If
                Next
            Catch ex As System.Exception
                MsgBox(ex.Message)
            End Try
            oSelTemp.Clear()
        End If
        ''
        oSelTemp = Nothing
        Return resultado
        Exit Function
    End Function
#End Region
    ''
#Region "TABLAPARCIAL"

    ''
    '' Insertar tabla para datos de etiquetas o la actualiza si ya existe.
    Public Sub TablaDatosParcialInserta()
        ConfiguraDibujo()
        ''
        '' Siempre crearemos una nueva tabla con los elementos seleccionados.
        If oApp Is Nothing Then _
                    oApp = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
        'oDoc = oApp.ActiveDocument
        oApp.ActiveDocument.Activate()
        Dim oPunto As Object = Nothing
        Try
            'oPunto = PuntoDame_NET("Punto de inserción de la tabla:")
            oPunto = oApp.ActiveDocument.Utility.GetPoint(, "Punto de inserción de la tabla:")
        Catch ex As System.Exception
            'MsgBox(ex.Message)
        End Try
        ''
        If oPunto Is Nothing Then
            Exit Sub
        End If
        ''
        CapaCreaActivaTablaDatos()
        ''
        '' Insertamos la tabla.
        'Dim oPunto As Object = oDoc.Utility.GetPoint(, "Punto de inserción de la Tabla")
        'TablaDatos = oDoc.ModelSpace.AddTable(oPunto, 6, 2, 50 * escala, 400 * escala)
        TablaDatosParcial = oApp.ActiveDocument.ModelSpace.AddTable(oPunto, 6, 2, 1, 1)
        'Temporalmente desactivamos el recalculo de la tabla
        TablaDatosParcial.RegenerateTableSuppressed = True
        ''
        '' Ponemos los textos en las cabeceras
        TablaDatosParcial.SetText(0, 0, "IMPLACAD ES UN PRODUCTO IMPLASER (TABLA PARCIAL)")        '' Titulo tabla
        TablaDatosParcial.SetText(1, 0, "** BALIZAMIENTO **")                  '' Cabecera
        TablaDatosParcial.SetText(1, 1, "** METROS **")                '' Cabecera
        TablaDatosParcial.SetText(2, 0, "SUELO")        '' 
        TablaDatosParcial.SetText(3, 0, "PARED")        '' 
        TablaDatosParcial.SetText(4, 0, "TOTAL SEÑALES --->")        '' Titulo tabla
        'TablaDatos.SetText(4, 1, DameTotalBloquesImplacad)        '' Titulo tabla
        TablaDatosParcial.SetText(5, 0, "** REFERENCIA **")        '' Titulo tabla
        TablaDatosParcial.SetText(5, 1, "** CANTIDAD **")        '' Titulo tabla
        ''
        TablaDatosParcialActualiza()
        '' Regenerar la tabla
        TablaDatosParcial.RegenerateTableSuppressed = False
        TablaDatosParcial.RecomputeTableBlock(True)
        TablaDatosParcial = Nothing
        'vermensajes = True
    End Sub
    ''
    Public Sub TablaDatosParcialActualiza()
        ''
        '' Borrar Hashtable colBloquesCantidadParcial antes de recontar
        If colBloquesCantidadParcial IsNot Nothing Then colBloquesCantidadParcial.Clear()
        colBloquesCantidadParcial = Nothing
        ''
        '' Rellenamos los 3 arrays con los objetos seleccionados: Balizas Suelo, Balizas Pared y Bloques Implacad.
        TablaDatosParcialRellenaColecciones(oApp.ActiveDocument.ActiveSelectionSet)
        ''
        TablaDatosParcial.SetText(2, 1, DameTotalBaliza(queCapa.BALIZAMIENTOSUELO, arrBalizaSuelo))        '' Total Baliza Suelo en selección
        TablaDatosParcial.SetText(3, 1, DameTotalBaliza(queCapa.BALIZAMIENTOPARED, arrBalizaPared))        '' Total Baliza Pared en selección
        TablaDatosParcial.SetText(4, 1, DameTotalBloquesImplacad(arrBloquesIdParcial))       '' Total etiquetas en selección

        ''
        '' Ahora actualizaremos el listado de bloques y cantidades.
        '' Borrar todas las filas superiores a la 5
        If TablaDatosParcial.Rows > 6 Then
            TablaDatosParcial.DeleteRows(6, TablaDatosParcial.Rows - 5)
        End If
        ''
        '' Ahora creamos las filas que sean necesarias.
        If colBloquesCantidadParcial IsNot Nothing AndAlso colBloquesCantidadParcial.Count > 0 Then
            Dim colIni As Integer = 6
            TablaDatosParcial.InsertRows(colIni, 0.3, colBloquesCantidadParcial.Count)
            ''
            For Each nBloque As String In colBloquesCantidadParcial.Keys
                TablaDatosParcial.SetText(colIni, 0, nBloque)        '' Nombre del bloque
                TablaDatosParcial.SetText(colIni, 1, colBloquesCantidadParcial(nBloque))  '' Cantidad de bloques.
                colIni += 1
            Next
        End If
        '' Formateamos la tabla.
        TablaDatosParcial.TitleSuppressed = False              '' Activar Titulo
        TablaDatosParcial.HeaderSuppressed = False             '' Activar Cabecera
        TablaDatosParcial.RegenerateTableSuppressed = True     '' Regenerar/Suprimirla
        '' Margen de las celdas
        TablaDatosParcial.VertCellMargin = 0.1                '' Margen vertical celda
        TablaDatosParcial.HorzCellMargin = 0.25               '' Margen horizontal celda
        '' Altura de las filas de cabecera
        TablaDatosParcial.SetRowHeight(0, 0.48)                '' Altura Fila 0
        TablaDatosParcial.SetRowHeight(1, 0.4)                '' Altura Fila 1
        TablaDatosParcial.SetRowHeight(2, 0.4)                '' Altura Fila 2
        TablaDatosParcial.SetRowHeight(3, 0.4)                '' Altura Fila 3
        TablaDatosParcial.SetRowHeight(4, 0.4)                '' Altura Fila 4
        TablaDatosParcial.SetRowHeight(5, 0.4)                '' Altura Fila 5
        '' Anchura de las 2 columnas
        TablaDatosParcial.SetColumnWidth(0, 3.5)               '' Ancho columna 1
        TablaDatosParcial.SetColumnWidth(1, 3.5)               '' Ancho columna 2
        '' Estilo de texto
        TablaDatosParcial.SetTextStyle(RowType.TitleRow, textoEstilo.Name)  '' Estilo de texto de titulo
        TablaDatosParcial.SetTextStyle(RowType.HeaderRow, textoEstilo.Name) '' Estilo de texto de cabecera
        TablaDatosParcial.SetTextStyle(RowType.DataRow, textoEstilo.Name)   '' Estilo de texto de datos
        '' Altura de los textos
        TablaDatosParcial.SetTextHeight(RowType.TitleRow, 0.25)    ' 22 * escalaTotal)    '' Altura texto titulo=0.25
        TablaDatosParcial.SetTextHeight(RowType.HeaderRow, 0.2)    ' 2 * escalaTotal)    '' Altura texto cabecera=0.2
        TablaDatosParcial.SetTextHeight(RowType.DataRow, 0.2)      ' 2 * escalaTotal)    '' Altura texto celdas=0.2
        ''
        TablaDatosParcial.SetAlignment(RowType.TitleRow, AcCellAlignment.acMiddleCenter)   '' Alineacion titulo
        TablaDatosParcial.SetAlignment(RowType.HeaderRow, AcCellAlignment.acMiddleCenter)   '' Alineacion cabecera
        TablaDatosParcial.SetAlignment(RowType.DataRow, AcCellAlignment.acMiddleCenter)   '' Alineacion datos
        ''
        '' Altura del resto de las filas con datos a 0.3
        For y As Integer = 6 To TablaDatosParcial.Rows - 1
            TablaDatosParcial.SetRowHeight(y, 0.3)
        Next
        ''
        '' Regenerar la tabla
        'TablaDatos.RegenerateTableSuppressed = False
        'TablaDatos.RecomputeTableBlock(True)
        XData.XNuevo(TablaDatosParcial, "Clase=tablaparcial")
        'TablaDatos = Nothing
        If arrBloquesIdParcial IsNot Nothing Then arrBloquesIdParcial.Clear()
        arrBloquesIdParcial = Nothing
    End Sub
#End Region
    ''
#Region "TABLAESCALERAS"
    '' Insertar tabla para ESCALERAS o la actualiza si ya existe.
    Public Sub TablaEscalerasInserta()
        If DameBalizasEscaleras(queTipoBE.Todas).Count = 0 Then
            If vermensajes = True Then MsgBox("No existen balizas escaleras en el dibujo...")
            Exit Sub
        End If
        ConfiguraDibujo()
        ''
        '' Comprobar si ya existe la tabla de datos
        Dim TablaEscalerasTemp As AcadTable = DameTablaEscaleras()
        ''
        '' Si no existe, la creamos. Y si existe la actualizamos.
        If TablaEscalerasTemp Is Nothing And vermensajes = True Then
            If oApp Is Nothing Then _
                    oApp = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
            'oDoc = oApp.ActiveDocument
            oApp.ActiveDocument.Activate()
            Dim oPunto As Object = Nothing
            Try
                'oPunto = PuntoDame_NET("Punto de inserción de la tabla:")
                oPunto = oApp.ActiveDocument.Utility.GetPoint(, vbLf & "Punto de inserción de la tabla :")
            Catch ex As System.Exception
                'MsgBox(ex.Message)
            End Try
            ''
            If oPunto Is Nothing Then
                Exit Sub
            End If
            ''
            '' Activar capa "capacuadro"
            CapaCreaActivaTablaDatos()
            ''
            '' Insertamos la tabla.
            'Dim oPunto As Object = oDoc.Utility.GetPoint(, "Punto de inserción de la Tabla")
            'TablaDatos = oDoc.ModelSpace.AddTable(oPunto, 6, 2, 50 * escala, 400 * escala)
            TablaEscaleras = oApp.ActiveDocument.ModelSpace.AddTable(oPunto, 4, 2, 1, 1)
            'Temporalmente desactivamos el recalculo de la tabla
            TablaEscaleras.RegenerateTableSuppressed = True
            ''
            '' Ponemos los textos en las cabeceras
            TablaEscaleras.SetText(0, 0, "BALIZAMIENTO EN ESCALERAS")        '' Fila 0 = Titulo tabla
            TablaEscaleras.SetText(1, 0, "TOTAL ESCALERAS")             '' Fila 1.0 = Texto total
            TablaEscaleras.SetText(1, 1, DameBalizasEscaleras(queTipoBE.Todas).Count)     '' Fila 1.1 = Total escaleras
            TablaEscaleras.SetText(2, 0, "  - Asc. BA 0307L (m)")             '' Fila 2.0 = Texto total
            'TablaEscaleras.SetText(2, 1, DameBalizasEscaleras(queTipoBE.Ascendente).Count)     '' Fila 1.1 = Total escaleras
            TablaEscaleras.SetText(2, 1, "0")     '' Fila 1.1 = Total escaleras
            TablaEscaleras.SetText(3, 0, "  - Des. BA 0401L (m)")             '' Fila 3.0 = Texto total
            'TablaEscaleras.SetText(3, 1, DameBalizasEscaleras(queTipoBE.Descendente).Count)     '' Fila 1.1 = Total escaleras
            TablaEscaleras.SetText(3, 1, "0")     '' Fila 1.1 = Total escaleras
            '' Actualizamos la tabla ESCALERAS
            TablaEscalerasActualiza()
            '' Regenerar la tabla
            TablaEscaleras.RegenerateTableSuppressed = False
            TablaEscaleras.RecomputeTableBlock(True)
        Else
            TablaEscaleras = CType(oApp.ActiveDocument.ObjectIdToObject(TablaEscalerasTemp.ObjectID), AcadTable) 'TablaEscalerasTemp ' oDoc.ObjectIdToObject(TablaDatosTemp.ObjectID)
            'Temporalmente desactivamos el recalculo de la tabla
            TablaEscaleras.RegenerateTableSuppressed = True
            TablaEscalerasActualiza()
            '' Regenerar la tabla
            TablaEscaleras.RegenerateTableSuppressed = False
            TablaEscaleras.RecomputeTableBlock(True)
            If vermensajes = True Then MsgBox("Tabla ESCALERAS actualizada porque ya existía...")
        End If
        TablaEscaleras = Nothing
        'vermensajes = True
    End Sub
    ''
    ''
    Public Sub TablaEscalerasActualiza()
        ''
        '' Ahora actualizaremos el total de balizas escalera y sus datos.
        TablaEscaleras.SetText(1, 1, DameBalizasEscaleras(queTipoBE.Todas).Count)     '' Fila 1.1 = Total escaleras
        '' Borrar todas las filas superiores a la 2
        If TablaEscaleras.Rows > 4 Then
            TablaEscaleras.DeleteRows(4, TablaEscaleras.Rows - 4)
        End If
        ''
        '' Ahora creamos las filas que sean necesarias.
        Dim arrBEscaleras As ArrayList = DameBalizasEscaleras()
        ''arrBEscaleras.Sort()  'Da error si ordenamos
        If arrBEscaleras IsNot Nothing AndAlso arrBEscaleras.Count > 0 Then
            Dim colIni As Integer = 4   '' indice de la primera fila a insertar.
            '' Creamos 6 lineas por cada bloque (1 linea de asteriscos + 4 lineas de datos)
            '' Las insertamos a partir de la linea 4.
            TablaEscaleras.InsertRows(colIni, 0.3, (arrBEscaleras.Count * 5))
            ''
            '' Arrays para almacenar los totales
            Dim arrAsc As Array = New Double() {0, 0}
            Dim arrDes As Array = New Double() {0, 0}
            ''
            For x As Integer = 0 To arrBEscaleras.Count - 1 ' Each oBe As AcadBlockReference In arrBEscaleras
                Dim oBe As AcadBlockReference = arrBEscaleras(x)
                '' Cogemos los datos de los atributos
                Dim nombreescalera As String = oBe.GetAttributes(0).TextString
                Dim numeroescalones As String = oBe.GetAttributes(1).TextString
                Dim clase As String = oBe.GetAttributes(2).TextString
                Dim ancho As String = oBe.GetAttributes(3).TextString
                'Dim metrosescalon As Double = Format(CDbl(ancho / 1000), "##0.00")
                Dim metrosescalon As Double = Format(CDbl(ancho), "##0.00")
                Dim metrostotales As Double = Format(Math.Truncate((CDbl(numeroescalones) * metrosescalon) + 1), "##0.#")
                ''
                '(4)
                TablaEscaleras.SetText(colIni, 0, StrDup(28, "*"))             '' Fila 2.0 = separador 
                TablaEscaleras.SetText(colIni, 1, StrDup(28, "*"))             '' Fila 2.1 = separador
                ''
                colIni += 1 '(5)
                TablaEscaleras.SetText(colIni, 0, "Nombre escalera")
                TablaEscaleras.SetText(colIni, 1, nombreescalera)  ' NOMBREESCALERA)
                ''
                colIni += 1 '(6)
                If clase = "Ascendente" Or clase.ToUpper.StartsWith("A") Then
                    TablaEscaleras.SetText(colIni, 0, "BA 0307L (m)")
                ElseIf clase = "Descendente" Or clase.ToUpper.StartsWith("D") Then
                    TablaEscaleras.SetText(colIni, 0, "BA 0401L (m)")
                End If
                TablaEscaleras.SetText(colIni, 1, metrostotales)
                ''
                colIni += 1 '(7)
                TablaEscaleras.SetText(colIni, 0, "Nº Escalones")
                TablaEscaleras.SetText(colIni, 1, numeroescalones)  ' NUMEROESCALONES
                ''
                colIni += 1 '(8)
                TablaEscaleras.SetText(colIni, 0, "Metros escalon")
                TablaEscaleras.SetText(colIni, 1, metrosescalon)    ' ANCHO
                colIni += 1 '(9)
                ''
                '' Acumular totales Asc o Des
                If clase = "Ascendente" Or clase.ToUpper.StartsWith("A") Then
                    arrAsc(0) += 1 : arrAsc(1) += Format(metrostotales, "##0.#")
                ElseIf clase = "Descendente" Or clase.ToUpper.StartsWith("D") Then
                    arrDes(0) += 1 : arrDes(1) += Format(metrostotales, "##0.#")
                End If
            Next
            ''
            '' Ponemos los totales que hemos calculado
            TablaEscaleras.SetText(2, 1, "(" & arrAsc(0) & ") " & Format(arrAsc(1), "##0.#") & " m")             '' Fila 3.0 = Texto total
            TablaEscaleras.SetText(3, 1, "(" & arrDes(0) & ") " & Format(arrDes(1), "##0.#") & " m")     '' Fila 1.1 = Total escaleras
        End If
        '' Formateamos la tabla.
        TablaEscaleras.TitleSuppressed = False              '' Activar Titulo
        TablaEscaleras.HeaderSuppressed = False             '' Activar Cabecera
        TablaEscaleras.RegenerateTableSuppressed = True     '' Regenerar/Suprimirla
        TablaEscaleras.VertCellMargin = 0.1                 '' Margen vertical celda
        TablaEscaleras.HorzCellMargin = 0.25                '' Margen horizontal celda
        '' Altura filas fijas
        TablaEscaleras.SetRowHeight(0, 0.48)                '' Altura Fila 0
        TablaEscaleras.SetRowHeight(1, 0.4)                '' Altura Fila 1
        TablaEscaleras.SetRowHeight(2, 0.4)                '' Altura Fila 2
        TablaEscaleras.SetRowHeight(3, 0.4)                '' Altura Fila 3
        '' Anchura de cada una de las 2 columnas
        TablaEscaleras.SetColumnWidth(0, 3.5)               '' Ancho columna 0
        TablaEscaleras.SetColumnWidth(1, 3.5)               '' Ancho columna 1
        '' Estilo de texto la tabla
        TablaEscaleras.SetTextStyle(RowType.TitleRow, textoEstilo.Name)  '' Estilo de texto de titulo
        TablaEscaleras.SetTextStyle(RowType.HeaderRow, textoEstilo.Name) '' Estilo de texto de cabecera
        TablaEscaleras.SetTextStyle(RowType.DataRow, textoEstilo.Name)   '' Estilo de texto de datos
        '' Altura de texto de los datos
        TablaEscaleras.SetTextHeight(RowType.TitleRow, 0.25)    '22 * escalaTotal)    '' Altura texto titulo=0.25
        TablaEscaleras.SetTextHeight(RowType.HeaderRow, 0.2)    ' 2 * escalaTotal)    '' Altura texto cabecera=0.2
        TablaEscaleras.SetTextHeight(RowType.DataRow, 0.2)     '2 * escalaTotal)    '' Altura texto celdas=0.2
        ''
        TablaEscaleras.SetAlignment(RowType.TitleRow, AcCellAlignment.acMiddleCenter)   '' Alineacion titulo
        TablaEscaleras.SetAlignment(RowType.HeaderRow, AcCellAlignment.acMiddleCenter)   '' Alineacion cabecera
        TablaEscaleras.SetAlignment(RowType.DataRow, AcCellAlignment.acMiddleCenter)   '' Alineacion datos
        ''
        '' Altura del resto de las filas con datos a 0.3
        For y As Integer = 4 To TablaEscaleras.Rows - 1
            TablaEscaleras.SetRowHeight(y, 0.3)
        Next
        ''
        XData.XNuevo(TablaEscaleras, "Clase=tablaescaleras")
        'TablaEscaleras = Nothing
    End Sub
    ''
    Public Function DameTablaEscaleras() As AcadTable
        Dim resultado As AcadTable = Nothing
        'Dim cSeleccion As AcadSelectionSets
        Dim F1(1) As Short
        Dim F2(1) As Object
        Dim vF1 As Object = Nothing
        Dim vF2 As Object = Nothing
        '' 
        '' Filtros para seleccionar la Tabla
        'F1(0) = 100 : F2(0) = "AcDbSymbolTableRecord"
        'F1(1) = 100 : F2(1) = "AcDbLayerTableRecord"
        F1(0) = 0 : F2(0) = "ACAD_TABLE"
        F1(1) = 1001 : F2(1) = regAPP
        'F1(1) = 1000 : F2(1) = "Clase=tablaescaleras"
        ''
        vF1 = F1
        vF2 = F2
        ''
        If oApp Is Nothing Then _
        oApp = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
        ''
        ''
        ''
        Try
            oSelTemp = oApp.ActiveDocument.SelectionSets.Add("TEMPORAL")
        Catch ex As System.Exception
            oSelTemp = oApp.ActiveDocument.SelectionSets.Item("TEMPORAL")
        End Try
        ''
        ''
        oSelTemp.Clear()
        Try
            oSelTemp.Select(AcSelect.acSelectionSetAll, , , vF1, vF2)
        Catch ex As System.Exception
            Debug.Print(ex.Message)
        End Try
        If oSelTemp IsNot Nothing And oSelTemp.Count > 0 Then
            Try
                For Each oEnt As AcadEntity In oSelTemp
                    Dim texto As String = XData.XLeeDato(oEnt, xT.TEXTOS)
                    If TypeOf oEnt Is AcadTable And texto = "Clase=tablaescaleras" Then
                        resultado = CType(oApp.ActiveDocument.ObjectIdToObject(oEnt.ObjectID), AcadTable)
                        Exit For
                    End If
                Next
            Catch ex As System.Exception
                MsgBox(ex.Message)
            End Try
            oSelTemp.Clear()
        End If
        ''
        oSelTemp = Nothing
        Return resultado
        Exit Function
    End Function
    ''
    Public Function DameBalizasEscaleras(Optional queTipo As queTipoBE = queTipoBE.Todas) As ArrayList
        Dim resultado As New ArrayList
        'Dim cSeleccion As AcadSelectionSets
        Dim F1(1) As Short
        Dim F2(1) As Object
        Dim vF1 As Object = Nothing
        Dim vF2 As Object = Nothing
        '' 
        '' Filtros para seleccionar la Tabla
        'F1(0) = 100 : F2(0) = "AcDbSymbolTableRecord"
        'F1(0) = 0 : F2(0) = "Table"
        F1(0) = 1001 : F2(0) = regAPP
        F1(1) = 100 : F2(1) = "AcDbBlockReference"
        'F1(1) = 1000 : F2(1) = "Clase=balizaescalera"
        'F1(2) = 0 : F2(2) = "INSERT"
        ''
        vF1 = F1
        vF2 = F2
        ''
        If oApp Is Nothing Then _
        oApp = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
        ''
        ''
        ''
        Try
            oSelTemp = oApp.ActiveDocument.SelectionSets.Add("TEMPORAL")
        Catch ex As System.Exception
            oSelTemp = oApp.ActiveDocument.SelectionSets.Item("TEMPORAL")
        End Try
        ''
        ''
        oSelTemp.Clear()
        Try
            oSelTemp.Select(AcSelect.acSelectionSetAll, , , vF1, vF2)
        Catch ex As System.Exception
            Debug.Print(ex.Message)
        End Try
        If oSelTemp IsNot Nothing And oSelTemp.Count > 0 Then
            For Each oEnt As AcadEntity In oSelTemp
                Dim texto As String = XData.XLeeDato(oEnt, xT.TEXTOS)
                If TypeOf oEnt Is AcadBlockReference And texto = "Clase=balizaescalera" Then
                    Dim oBlo As AcadBlockReference = oEnt
                    Dim noEs As String = oBlo.GetAttributes(0).TextString    ' NOMBREESCALERA
                    Dim nuEs As String = oBlo.GetAttributes(1).TextString    ' NUMEROESCALONES
                    Dim clEs As String = oBlo.GetAttributes(2).TextString    ' CLASE
                    Dim anEs As String = oBlo.GetAttributes(3).TextString    ' ANCHO
                    ''
                    If queTipo = queTipoBE.Todas Then
                        resultado.Add(oBlo)
                    ElseIf queTipo = queTipoBE.Ascendente And clEs = queTipoBE.Ascendente.ToString Then
                        resultado.Add(oBlo)
                    ElseIf queTipo = queTipoBE.Descendente And clEs = queTipoBE.Descendente.ToString Then
                        resultado.Add(oBlo)
                    End If
                End If
            Next
            oSelTemp.Clear()
            'resultado.Sort()
        End If
        ''
        oSelTemp = Nothing
        Return resultado
        Exit Function
    End Function
#End Region
    ''
#Region "UTILIDADES"
    ''
    '' Permite escalar el dibujo para ponerlo en metros.
    '' Pide 2 puntos para calcular medida y nos pide cuando debe medir esto en metros.
    Public Function EscalaDibujoMetros() As Double
        Dim resultado As Double = 0
        ConfiguraDibujo()
        ''
        If oApp Is Nothing Then _
                oApp = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
        'oDoc = oApp.ActiveDocument
        oApp.ActiveDocument.Activate()
        Dim oPt1 As Object = Nothing
        'Dim oPt2 As Object = Nothing
        Dim medidaorigen As Double = 0
        Dim medidadestino As Double = 0
        Dim queescala As Double = 0
        Try
            'oPunto = PuntoDame_NET("Punto de inserción de la tabla:")
            oPt1 = oApp.ActiveDocument.Utility.GetPoint(, vbLf & "Primer punto para MEDIR:")
            'oPt2 = oApp.ActiveDocument.Utility.GetPoint(oPt1, vbLf & "Segundo punto para MEDIR:")
            medidaorigen = oApp.ActiveDocument.Utility.GetDistance(oPt1, vbLf & "Segundo punto para MEDIR:")
            medidadestino = oApp.ActiveDocument.Utility.GetDistance(, "Medida final en METROS")
        Catch ex As System.Exception
            'MsgBox(ex.Message)
            Return 0
            Exit Function
        End Try
        ''
        If oPt1 Is Nothing Or medidaorigen = 0 Or medidadestino = 0 Then
            resultado = 0
            Exit Function
        Else
            queescala = medidadestino / medidaorigen
            resultado = queescala
        End If
        ''
        Try
            Dim cadena As String = "._scale t  0,0 " & queescala & " "
            Application.DocumentManager.MdiActiveDocument.SendStringToExecute(cadena, True, False, False)
        Catch ex As System.Exception
            MsgBox(ex)
            resultado = 0
        End Try
        ''
        Return resultado
    End Function
#End Region
    ''
#Region "DAMETOTALES"
    ''
    Public Function DameTodoImplacad() As ArrayList
        Dim arrEnt As ArrayList = Nothing
        'Dim cSeleccion As AcadSelectionSets
        Dim F1(1) As Short
        Dim F2(1) As Object
        'Dim cSeleccion As AcadSelectionSets
        'Dim F1(0) As Short
        'Dim F2(0) As Object
        Dim vF1 As Object = Nothing
        Dim vF2 As Object = Nothing

        '' Las 2 maneras valen igual. AcDbBlckReference es mejor (Solo coge bloques) INSERT coge sombreados también.
        F1(0) = 100 : F2(0) = "AcDbBlockReference"
        'F1(0) = 0 : F2(0) = "INSERT"
        F1(1) = 1001 : F2(1) = regAPP
        'F1(1) = 1000 : F2(1) = "Clase=etiqueta"
        ''
        vF1 = F1
        vF2 = F2
        ''
        If oApp Is Nothing Then _
        oApp = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
        ''
        ''
        ''
        Try
            oSelTemp = oApp.ActiveDocument.SelectionSets.Add("TEMPORAL")
        Catch ex As System.Exception
            oSelTemp = oApp.ActiveDocument.SelectionSets.Item("TEMPORAL")
        End Try
        ''
        ''
        oSelTemp.Clear()
        Try
            oSelTemp.Select(AcSelect.acSelectionSetAll, , , vF1, vF2)
        Catch ex As System.Exception
            Debug.Print(ex.Message)
        End Try
        ''
        If oSelTemp IsNot Nothing And oSelTemp.Count > 0 Then
            arrEnt = New ArrayList
            For Each oEnt As AcadEntity In oSelTemp
                '' Si no es un bloque, continuamos
                If Not (TypeOf oEnt Is AcadBlockReference) Then Continue For
                Dim oBl As AcadBlockReference = oApp.ActiveDocument.ObjectIdToObject(oEnt.ObjectID)
                '' Si el nombre no empieza con EX, EV, SIA o KIT (arrpreEti) Borramos Xdata de regApp y Continuamos
                Dim borraXdata As Boolean = True
                For Each preEti As String In arrpreEti
                    If oBl.Name.ToUpper.StartsWith(preEti) Then
                        borraXdata = False
                        Exit For
                    End If
                Next
                ''
                If borraXdata = True Then
                    XData.XBorrar(oBl)
                    Continue For
                Else
                    arrEnt.Add(oEnt)
                End If
            Next
            oSelTemp.Clear()
        End If
        ''
        oSelTemp = Nothing
        Return arrEnt
        ''
    End Function
    ''
    Public Sub TablaDatosParcialRellenaColecciones(queSel As AcadSelectionSet)
        '' Rellenar los 3 arrays con: Balizas Pared, Balizas Suelo, BloquesImplacad.
        '' Sacados de la selección actual.
        If queSel IsNot Nothing And queSel.Count > 0 Then
            'resultado = oSelTemp.Count
            '' Para asegurarnos que no cuenta las tablas.
            arrBalizaPared = New ArrayList
            arrBalizaSuelo = New ArrayList
            arrBloquesIdParcial = New ArrayList
            ''
            For Each oEnt As AcadEntity In queSel
                Dim texto As String = XData.XLeeDato(oEnt, xT.TEXTOS)
                If TypeOf oEnt Is AcadBlockReference And texto = "Clase=etiqueta" Then
                    ''
                    If oEnt.Layer <> "0" Then Continue For
                    ''
                    Dim oBloque As AcadBlockReference = oApp.ActiveDocument.ObjectIdToObject(oEnt.ObjectID)
                    If arrBloquesIdParcial.Contains(oEnt.ObjectID) = False Then
                        arrBloquesIdParcial.Add(oEnt.ObjectID)
                    End If
                    oBloque = Nothing
                ElseIf TypeOf oEnt Is AcadLine Then
                    Dim oLine As AcadLine = oApp.ActiveDocument.ObjectIdToObject(oEnt.ObjectID)
                    If oLine.Layer = "BALIZAMIENTO SUELO" And arrBalizaSuelo.Contains(oEnt.ObjectID) = False Then
                        arrBalizaSuelo.Add(oEnt.ObjectID)
                    ElseIf oLine.Layer = "BALIZAMIENTO PARED" And arrBalizaPared.Contains(oEnt.ObjectID) = False Then
                        arrBalizaPared.Add(oEnt.ObjectID)
                    End If
                    oLine = Nothing
                End If
                texto = ""
            Next
        End If
    End Sub
    ''
    Public Function DameTotalBloquesImplacad(Optional queArr As ArrayList = Nothing) As Integer
        Dim resultado As Integer = 0
        'Dim cSeleccion As AcadSelectionSets
        Dim F1(1) As Short
        Dim F2(1) As Object
        'Dim cSeleccion As AcadSelectionSets
        'Dim F1(0) As Short
        'Dim F2(0) As Object
        Dim vF1 As Object = Nothing
        Dim vF2 As Object = Nothing

        '' Las 2 maneras valen igual. AcDbBlckReference es mejor (Solo coge bloques) INSERT coge sombreados también.
        F1(0) = 100 : F2(0) = "AcDbBlockReference"
        'F1(0) = 0 : F2(0) = "INSERT"
        F1(1) = 1001 : F2(1) = regAPP
        'F1(1) = 1000 : F2(1) = "Clase=etiqueta"
        ''
        vF1 = F1
        vF2 = F2
        ''
        If oApp Is Nothing Then _
        oApp = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
        ''
        ''
        ''
        Try
            oSelTemp = oApp.ActiveDocument.SelectionSets.Add("TEMPORAL")
        Catch ex As System.Exception
            oSelTemp = oApp.ActiveDocument.SelectionSets.Item("TEMPORAL")
        End Try
        ''
        ''
        oSelTemp.Clear()
        Try
            oSelTemp.Select(AcSelect.acSelectionSetAll, , , vF1, vF2)
        Catch ex As System.Exception
            Debug.Print(ex.Message)
        End Try
        If oSelTemp IsNot Nothing And oSelTemp.Count > 0 Then
            'resultado = oSelTemp.Count
            '' Para asegurarnos que no cuenta las tablas.
            colBloquesCantidad = New SortedList ' Hashtable
            colBloquesCantidadParcial = New SortedList  'Hashtable
            'arrBloquesId = New ArrayList
            For Each oEnt As AcadEntity In oSelTemp
                ''
                '' Si no estan en estas capas no los contamos (Solo etiquetas)
                If oEnt.Layer <> "BALIZAMIENTO SUELO" And
                    oEnt.Layer <> "BALIZAMIENTO PARED" And
                    oEnt.Layer <> "0" Then
                    Continue For
                End If
                ''
                Dim texto As String = XData.XLeeDato(oEnt, xT.TEXTOS)
                If TypeOf oEnt Is AcadBlockReference And texto = "Clase=etiqueta" Then
                    Dim oBloque As AcadBlockReference = oApp.ActiveDocument.ObjectIdToObject(oEnt.ObjectID)
                    '' Si el nombre no empieza con EX, EV, SIA o KIT (arrpreEti) Borramos Xdata de regApp y Continuamos
                    Dim borraXdata As Boolean = True
                    For Each preEti As String In arrpreEti
                        If oBloque.Name.ToUpper.StartsWith(preEti) Then
                            borraXdata = False
                            Exit For
                        End If
                    Next

                    '' Si el nombre no empieza con EX, EV, SIA o KIT (arrpreEti) Continuamos
                    If borraXdata = True Then
                        XData.XBorrar(oBloque)
                        Continue For
                    End If
                    ''
                    If queArr Is Nothing Then
                        ''
                        If oEtis.DConjuntos.ContainsKey(oBloque.EffectiveName) = True Then
                            'SumaConjuntos(oBloque.EffectiveName, colBloquesCantidad, resultado)
                            Dim oSubs As Object = oBloque.Explode
                            For Each oSubEnt As AcadEntity In oSubs ' oBloque.Explode
                                Dim t As String = ""
                                If TypeOf oSubEnt Is AcadMText Then
                                    t = CType(oSubEnt, AcadMText).TextString
                                ElseIf TypeOf oSubEnt Is AcadText Then
                                    t = CType(oSubEnt, AcadText).TextString
                                End If
                                If t.Contains(";") Then
                                    t = t.Split(";")(1)
                                End If
                                t = t.Replace("}", "")
                                Dim textovalido As Boolean = False
                                For Each pre As String In arrpreEti
                                    If t.StartsWith(pre) = True Then
                                        textovalido = True
                                        Exit For
                                    End If
                                Next
                                If textovalido = False Then Continue For
                                Dim etiquetas As String() = t.Split(" ")
                                For Each eti As String In etiquetas
                                    If colBloquesCantidad.ContainsKey(eti) Then
                                        colBloquesCantidad(eti) += 1
                                    Else
                                        colBloquesCantidad.Add(key:=eti, value:=1)
                                    End If
                                    resultado += 1
                                Next
                                Exit For
                            Next
                            For Each oE As AcadEntity In oSubs
                                oE.Delete()
                            Next
                        Else
                            If colBloquesCantidad.ContainsKey(oBloque.EffectiveName) Then
                                colBloquesCantidad(oBloque.EffectiveName) += 1
                            Else
                                colBloquesCantidad.Add(key:=oBloque.EffectiveName, value:=1)
                            End If
                            resultado += 1
                        End If
                        'arrBloquesId.Add(oBloque.ObjectID)
                    ElseIf queArr IsNot Nothing AndAlso queArr.Contains(oEnt.ObjectID) Then
                        ''
                        If oEtis.DConjuntos.ContainsKey(oBloque.EffectiveName) = True Then
                            'SumaConjuntos(oBloque.EffectiveName, colBloquesCantidadParcial, resultado)
                            Dim oSubs As Object = oBloque.Explode
                            For Each oSubEnt As AcadEntity In oSubs ' oBloque.Explode
                                Dim t As String = ""
                                If TypeOf oSubEnt Is AcadMText Then
                                    t = CType(oSubEnt, AcadMText).TextString
                                ElseIf TypeOf oSubEnt Is AcadText Then
                                    t = CType(oSubEnt, AcadText).TextString
                                End If
                                If t.Contains(";") Then
                                    t = t.Split(";")(1)
                                End If
                                t = t.Replace("}", "")
                                Dim textovalido As Boolean = False
                                For Each pre As String In arrpreEti
                                    If t.StartsWith(pre) = True Then
                                        textovalido = True
                                        Exit For
                                    End If
                                Next
                                If textovalido = False Then Continue For
                                Dim etiquetas As String() = t.Split(" ")
                                For Each eti As String In etiquetas
                                    If colBloquesCantidadParcial.ContainsKey(eti) Then
                                        colBloquesCantidadParcial(eti) += 1
                                    Else
                                        colBloquesCantidadParcial.Add(key:=eti, value:=1)
                                    End If
                                    resultado += 1
                                Next
                                Exit For
                            Next
                            For Each oE As AcadEntity In oSubs
                                oE.Delete()
                            Next
                        Else
                            If colBloquesCantidadParcial.ContainsKey(oBloque.EffectiveName) Then
                                colBloquesCantidadParcial(oBloque.EffectiveName) += 1
                            Else
                                colBloquesCantidadParcial.Add(key:=oBloque.EffectiveName, value:=1)
                            End If
                            resultado += 1
                            'arrBloquesIdParcial.Add(oBloque.ObjectID)
                        End If
                    Else
                        Continue For
                    End If
                    oBloque = Nothing
                End If
                texto = ""
            Next
            oSelTemp.Clear()
        End If
        ''
        oSelTemp = Nothing
        Return resultado
        Exit Function
    End Function
    ''
    'Public Sub SumaConjuntos(nomBlo As String, ByRef colBl As SortedList, ByRef cont As Integer)
    '    ''
    '    '' Recorremos el conjunto de etiquetas para añadir las que contenga
    '    If oEtis.DConjuntos.ContainsKey(nomBlo) = False Then Exit Sub
    '    If oEtis.DConjuntos(nomBlo).COMPONENTES Is Nothing OrElse oEtis.DConjuntos(nomBlo).COMPONENTES = "" Then Exit Sub
    '    ''
    '    'Dim queEtis() As String = oEtis.DConjuntos(nomBlo).COMPONENTES
    '    For Each queEti As String In oEtis.DConjuntos(nomBlo).COMPONENTES.Split(";"c)
    '        If colBl.ContainsKey(queEti) Then
    '            colBl(queEti) += 1
    '        Else
    '            colBl.Add(key:=queEti, value:=1)
    '        End If
    '        cont += 1
    '    Next
    'End Sub
    Public Function DameTotalBaliza(queCapa As queCapa, Optional queArr As ArrayList = Nothing) As Double
        'AcDbLayerTableRecord
        Dim resultado As Double = 0
        'Dim cSeleccion As AcadSelectionSets
        Dim F1(1) As Short
        Dim F2(1) As Object
        Dim vF1 As Object = Nothing
        Dim vF2 As Object = Nothing
        Dim capa As String = ""
        ''  
        '' Filtro de nombre entidad
        F1(0) = 0 : F2(0) = "LINE"
        'F1(0) = 100 : F2(0) = "AcDbLine"
        'DatosX Siempre tiene que estar despues de entidad. Si no no funciona
        'F1(0) = 1001 : F2(0) = regAPP
        'Filtro de bloque, nombre y capa
        If queCapa = queCapa.BALIZAMIENTOSUELO Then
            capa = "BALIZAMIENTO SUELO"
        ElseIf queCapa = Global.IMPLACAD.queCapa.BALIZAMIENTOPARED Then
            capa = "BALIZAMIENTO PARED"
        End If
        F1(1) = 8 : F2(1) = capa
        ''
        vF1 = F1
        vF2 = F2
        ''
        If oApp Is Nothing Then _
        oApp = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
        ''
        ''
        ''
        Try
            oSelTemp = oApp.ActiveDocument.SelectionSets.Add("TEMPORAL")
        Catch ex As System.Exception
            oSelTemp = oApp.ActiveDocument.SelectionSets.Item("TEMPORAL")
        End Try
        ''
        ''
        oSelTemp.Clear()
        Try
            oSelTemp.Select(AcSelect.acSelectionSetAll, , , vF1, vF2)
        Catch ex As System.Exception
            Debug.Print(ex.Message)
        End Try
        If oSelTemp IsNot Nothing And oSelTemp.Count > 0 Then
            For Each oEnt As AcadEntity In oSelTemp
                If TypeOf oEnt Is AcadLine Then
                    Dim oLine As AcadLine = oApp.ActiveDocument.ObjectIdToObject(oEnt.ObjectID)
                    If queArr Is Nothing Then
                        resultado += oLine.Length + 1
                    ElseIf queArr IsNot Nothing AndAlso queArr.Contains(oLine.ObjectID) Then
                        resultado += oLine.Length + 1
                    Else
                        Continue For
                    End If
                    oLine = Nothing
                End If
            Next
            oSelTemp.Clear()
        End If
        ''
        oSelTemp = Nothing
        '' Devolvemos los mm pasados a metros y con formato de 2 decimales.
        'Return Format((resultado / 1000), "0.000")
        '' Devolvemos los m con formato de 2 decimales.
        If resultado > 0 Then
            Return Format(Math.Truncate(resultado + 1), "##0.#")
        Else
            Return Format(0, "##0.#")
        End If
        Exit Function
    End Function
#End Region
    ''
    Public Sub SeleccionaBloquesImplacad(Optional ByVal nombre As Object = "", Optional ByVal capa As Object = "", Optional ByVal DatosX As Boolean = True, Optional ByVal SelTemp As Boolean = False)
        'Dim cSeleccion As AcadSelectionSets
        Dim F1(0) As Short
        Dim F2(0) As Object
        Dim vF1 As Object = Nothing
        Dim vF2 As Object = Nothing

        '' Las 2 maneras valen igual. AcDbBlckReference es mejor (Solo coge bloques) INSERT coge sombreados también.
        F1(0) = 100 : F2(0) = "AcDbBlockReference"
        'F1(0) = 0 : F2(0) = "INSERT"
        'DatosX Siempre tiene que estar despues de entidad. Si no no funciona
        If DatosX = True Then
            ReDim Preserve F1(F1.Length)
            ReDim Preserve F2(F2.Length)
            'Filtro de bloque, nombre y capa
            F1(F1.Length - 1) = 1001 : F2(F2.Length - 1) = regAPP
        End If
        If nombre <> "" Then
            ReDim Preserve F1(F1.Length)
            ReDim Preserve F2(F2.Length)
            'Filtro de bloque y nombre
            F1(F1.Length - 1) = 2 : F2(F2.Length - 1) = nombre
        End If
        If capa <> "" Then
            ReDim Preserve F1(F1.Length)
            ReDim Preserve F2(F2.Length)
            'Filtro de bloque, nombre y capa
            F1(F1.Length - 1) = 8 : F2(F1.Length - 1) = capa
        End If

        vF1 = F1
        vF2 = F2
        ''
        oApp = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
        ''
        ''
        Try
            oSel = oApp.ActiveDocument.SelectionSets.Add(regAPP)
        Catch ex As System.Exception
            oSel = oApp.ActiveDocument.SelectionSets.Item(regAPP)
        End Try
        ''
        Try
            oSelTemp = oApp.ActiveDocument.SelectionSets.Add("TEMPORAL")
        Catch ex As System.Exception
            oSelTemp = oApp.ActiveDocument.SelectionSets.Item("TEMPORAL")
        End Try
        ''
        ''
        If SelTemp = False Then
            oSel.Clear()
            Try
                oSel.Select(AcSelect.acSelectionSetAll, , , vF1, vF2)
            Catch ex As System.Exception
                Debug.Print(ex.Message)
            End Try
            MsgBox(oSel.Count)
        Else
            oSelTemp.Clear()
            Try
                oSelTemp.Select(AcSelect.acSelectionSetAll, , , vF1, vF2)
            Catch ex As System.Exception
                Debug.Print(ex.Message)
            End Try
            MsgBox(oSel.Count)
        End If
    End Sub

    Public Function SeleccionaDameBloque(Optional ByVal nombre As Object = "", Optional ByVal capa As Object = "", Optional ByVal DatosX As Boolean = True, Optional ByVal SelTemp As Boolean = False) As AcadBlockReference
        Dim resultado As AcadBlockReference = Nothing
        'Dim cSeleccion As AcadSelectionSets
        Dim F1(0) As Short
        Dim F2(0) As Object
        Dim vF1 As Object = Nothing
        Dim vF2 As Object = Nothing

        '' Las 2 maneras valen igual. AcDbBlckReference es mejor (Solo coge bloques) INSERT coge sombreados también.
        F1(0) = 100 : F2(0) = "AcDbBlockReference"
        'F1(0) = 0 : F2(0) = "INSERT"
        'DatosX Siempre tiene que estar despues de entidad. Si no no funciona
        If DatosX = True Then
            ReDim Preserve F1(F1.Length)
            ReDim Preserve F2(F2.Length)
            'Filtro de bloque, nombre y capa
            F1(F1.Length - 1) = 1001 : F2(F2.Length - 1) = regAPP
        End If
        If nombre <> "" Then
            ReDim Preserve F1(F1.Length)
            ReDim Preserve F2(F2.Length)
            'Filtro de bloque y nombre
            F1(F1.Length - 1) = 2 : F2(F2.Length - 1) = nombre
        End If
        If capa <> "" Then
            ReDim Preserve F1(F1.Length)
            ReDim Preserve F2(F2.Length)
            'Filtro de bloque, nombre y capa
            F1(F1.Length - 1) = 8 : F2(F1.Length - 1) = capa
        End If

        vF1 = F1
        vF2 = F2
        ''
        If oApp Is Nothing Then _
        oApp = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
        ''
        ''
        Try
            oSel = oApp.ActiveDocument.SelectionSets.Add(regAPP)
        Catch ex As System.Exception
            oSel = oApp.ActiveDocument.SelectionSets.Item(regAPP)
        End Try
        ''
        Try
            oSelTemp = oApp.ActiveDocument.SelectionSets.Add("TEMPORAL")
        Catch ex As System.Exception
            oSelTemp = oApp.ActiveDocument.SelectionSets.Item("TEMPORAL")
        End Try
        ''
        ''
        oApp.ActiveDocument.SetVariable("pickadd", 0)   '' Solo una selección. Se quita lo que hubiera
        If SelTemp = False Then
            oSel.Clear()
            Try
                oSel.SelectOnScreen(vF1, vF2)
            Catch ex As System.Exception
                'Debug.Print(ex.Message)
            End Try
            'MsgBox(oSel.Count)
            If oSel.Count > 0 Then
                If TypeOf oSel.Item(oSel.Count - 1) Is AcadBlockReference Then
                    resultado = oSel.Item(oSel.Count - 1)
                Else
                    resultado = Nothing
                End If
            Else
                resultado = Nothing
            End If
        Else
            oSelTemp.Clear()
            Try
                oSelTemp.Select(vF1, vF2)
            Catch ex As System.Exception
                'Debug.Print(ex.Message)
            End Try
            'MsgBox(oSel.Count)
            If oSelTemp.Count > 0 Then
                'Dim texto As String = XData.XLeeDato(oEnt, xT.TEXTOS)
                'If TypeOf oEnt Is AcadTable Or texto = "Clase=tabla" Then
                If TypeOf oSelTemp.Item(oSelTemp.Count - 1) Is AcadTable Then
                    resultado = Nothing
                Else
                    resultado = oSelTemp.Item(oSelTemp.Count - 1)
                End If
            Else
                resultado = Nothing
            End If
        End If
        ''
        oApp.ActiveDocument.SetVariable("pickadd", 2)   '' La seleccion actual se suma a la que hubiera.
        Return resultado
        Exit Function
    End Function

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
            oLayer = oApp.ActiveDocument.Layers.Item("Zona Cobertura Evacuación")
        Catch ex As System.Exception
            '' Si no existia, la creamos.
            oLayer = oApp.ActiveDocument.Layers.Add("Zona Cobertura Evacuación")
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
            oLayer = oApp.ActiveDocument.Layers.Item("Zona Cobertura Extincion")
        Catch ex As System.Exception
            '' Si no existia, la creamos.
            oLayer = oApp.ActiveDocument.Layers.Add("Zona Cobertura Extincion")
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

    Public Sub GrosorLineasACTDES(Optional estado As CapaEstado = CapaEstado.Inverso)
        If oApp Is Nothing Then _
        oApp = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
        ''
        '' Configurar grosores de linea (0=desactivados  /  1=activados  / 2=Not(estado actual)
        Dim oLayer As AcadLayer = Nothing
        Try
            If estado = CapaEstado.Desactivar Then
                oApp.ActiveDocument.Preferences.LineWeightDisplay = False
            ElseIf estado = CapaEstado.Activar Then
                oApp.ActiveDocument.Preferences.LineWeightDisplay = True
            ElseIf estado = CapaEstado.Inverso Then
                oApp.ActiveDocument.Preferences.LineWeightDisplay = Not (oApp.ActiveDocument.Preferences.LineWeightDisplay)
            End If
            oApp.ActiveDocument.Regen(AcRegenType.acAllViewports)
        Catch ex As System.Exception
            ''
        End Try
    End Sub
    ''
    Public Sub CambiaBloquePorImagen(ByRef oDoc As AcadDocument)
        ''
        If oApp Is Nothing Then _
            oApp = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
        '' Activar Capa 0 primero.
        oDoc.ActiveLayer = oDoc.Layers.Item("0")
        ''
        '' Variables de imagenes
        Dim imgPath As String = ""
        Dim imgPunto As Object = Nothing
        Dim imgEscala As Double = 1
        Dim imgRotacion As Double = oDoc.Utility.AngleToReal("0", AcAngleUnits.acRadians)
        Dim rot0 As Double = oDoc.Utility.AngleToReal("0", AcAngleUnits.acRadians)
        Dim rot90 As Double = oDoc.Utility.AngleToReal("90", AcAngleUnits.acRadians)
        Dim rot180 As Double = oDoc.Utility.AngleToReal("180", AcAngleUnits.acRadians)
        Dim rot270 As Double = oDoc.Utility.AngleToReal("270", AcAngleUnits.acRadians)

        ''
        '' 1.- Cogemos todos los bloques de IMPLACAD insertados
        Dim arrBlo As ArrayList = DameTodoImplacad()
        ''
        '' 2.- Cogemos todas las imagenes insertadas de la plantilla
        'Dim colImg As New Hashtable
        'For Each oEnt As AcadEntity In oDoc.ModelSpace
        '    If TypeOf oEnt Is AcadRasterImage Then
        '        Dim oImg As AcadRasterImage = oDoc.ObjectIdToObject(oEnt.ObjectID)
        '        If oImg.ImageFile.ToUpper.Contains("PLANOS_Y_BALIZAMIENTOS") Then
        '            If colImg.ContainsKey(oImg.Name) = False Then colImg.Add(oImg.Name, oImg.ObjectID)
        '        End If
        '    End If
        'Next
        ' ''
        'If colImg.Count = 0 Then
        '    Exit Sub
        'End If
        ''
        '' Recorrer cada bloque y cambiarlo por la imagen que le corresponda.
        For Each oEnt As AcadEntity In arrBlo
            If Not (TypeOf oEnt Is AcadBlockReference) Then Continue For
            ''
            Dim oBl As AcadBlockReference = oApp.ActiveDocument.ObjectIdToObject(oEnt.ObjectID)
            Dim oImgFin As AcadRasterImage = Nothing
            ''
            Dim bln As String = oBl.EffectiveName.ToUpper
            ''
            '' Si lo encontramos en colSustituciones, definimos la imagen.
            '' Si no lo encontramos, recorremos colSustituciones para ver si empieza por StartsWith
            '' para definir la imagen.
            Dim encontrado As Boolean = False
            If oEtis.DSustituciones.ContainsKey(bln) Then
                '' Lo encontramos en oEtis.DSustituciones. Definimos su imagen
                imgPath = IO.Path.Combine(IMPLACAD_DATA, balizas, oEtis.DSustituciones(bln) & ".png")
                encontrado = True
            Else
                '' Buscaremos en colSustituciones para ver si nombre bloque empieza por (menos de 10)
                For Each queRef As String In oEtis.DSustituciones.Keys
                    '' Si es >= 10 caracteres, saltamos al siguiente
                    If queRef.Length >= 10 Then Continue For
                    ''
                    If bln.ToUpper.StartsWith(queRef.ToUpper) Then
                        imgPath = IO.Path.Combine(IMPLACAD_DATA, balizas, oEtis.DSustituciones(queRef) & ".png")
                        encontrado = True
                        Exit For
                    End If
                Next
            End If
            ''
            If encontrado = False Then
                oBl.Delete()
                Continue For
            End If
            imgPunto = oBl.InsertionPoint
            ''
            '' Insertamos la imagen que corresponda (desplazada XX unidades)
            If imgPath <> "" AndAlso IO.File.Exists(imgPath) Then
                Try
                    oImgFin = oApp.ActiveDocument.ModelSpace.AddRaster(imgPath, imgPunto, imgEscala, imgRotacion)
                Catch ex As System.Exception
                    MsgBox(ex.Message &
                           "Comprueba el fichero : " & vbCrLf & vbCrLf &
                           imgPath)
                End Try
            End If
            ''
            If oImgFin IsNot Nothing Then
                If oImgFin.ImageHeight = oImgFin.ImageWidth Then
                    oImgFin.ScaleEntity(imgPunto, 5)     ' 5 = 8.74  /  2.86 = 5
                ElseIf oImgFin.ImageHeight < oImgFin.ImageWidth Then
                    Dim factor As Double = 5 / oImgFin.ImageHeight
                    oImgFin.ImageHeight *= factor
                    oImgFin.ImageWidth *= factor
                ElseIf oImgFin.ImageWidth < oImgFin.ImageHeight Then
                    Dim factor As Double = 5 / oImgFin.ImageWidth
                    oImgFin.ImageWidth *= factor
                    oImgFin.ImageHeight *= factor
                End If
                oImgFin.Update()
                ''
                '' La volvemos a posicionar en el punto de inserción del bloque
                oImgFin.Move(oImgFin.Origin, oBl.InsertionPoint)
                oImgFin.Update()
                '' Ahora movemos la imagen para que quede alineada con el bloque al que sustutuye
                '' Punto al que move la imagen
                Dim ptFin(2) As Double
                ptFin(0) = oImgFin.Origin(0)
                ptFin(1) = oImgFin.Origin(1)
                ptFin(2) = oImgFin.Origin(2)
                ''
                Dim medX As Double = oImgFin.ImageWidth
                Dim medY As Double = oImgFin.ImageHeight
                '' Rotacion del bloque en grados (180 grados = PI radianes)
                '' grados = (180 x radianes) / PI
                '' 10 = (5 x 4) / 2
                '' radianes = (grados / PI) / 180
                '' 4 = (10 x 2) / 5
                'Dim oBlr As Double = Math.Round((180 * oBl.Rotation) / Math.PI, 0)
                Dim oBlr As Double = Math.Round(RadGra(oBl.Rotation), 0)
                ''
                If oBlr = 0 Then
                    ptFin(0) = ptFin(0) - (medX / 2.0#)      'IIf(ptFin(0) > 0, medX, -medX)
                    ptFin(1) = ptFin(1) - medY      'IIf(ptFin(1) > 0, medY, -medY)
                ElseIf oBlr > 0 And oBlr < 90 Then
                    'ptFin(0) = ptFin(0) - (medX / 2.0#)     ' IIf(ptFin(0) > 0, medX, -medX)
                    ptFin(1) = ptFin(1) - medY      'IIf(ptFin(1) > 0, medY, -medY)
                ElseIf oBlr = 90 Then
                    'ptFin(0) = ptFin(0) - (medX / 2.0#)      'IIf(ptFin(0) > 0, medX, -medX)
                    ptFin(1) = ptFin(1) - (medY * 0.5#)      'IIf(ptFin(1) > 0, medY, -medY)
                ElseIf oBlr > 90 And oBlr < 180 Then
                    'ptFin(0) = ptFin(0) - (medX / 2.0#)      'IIf(ptFin(0) > 0, medX, -medX)
                    'ptFin(1) = ptFin(1) - medY      'IIf(ptFin(1) > 0, medY, -medY)
                ElseIf oBlr = 180 Then
                    ptFin(0) = ptFin(0) - (medX * 0.5#)      'IIf(ptFin(0) > 0, medX, -medX)
                    'ptFin(1) = ptFin(1) - medY      'IIf(ptFin(1) > 0, medY, -medY)
                ElseIf oBlr > 180 And oBlr < 270 Then
                    ptFin(0) = ptFin(0) - medX      'IIf(ptFin(0) > 0, (medX * 2), -(medX * 2))
                    'ptFin(1) = ptFin(1) - medY      'IIf(ptFin(1) > 0, medY, -medY)
                ElseIf oBlr = 270 Then
                    ptFin(0) = ptFin(0) - medX      'IIf(ptFin(0) > 0, (medX * 2), -(medX * 2))
                    ptFin(1) = ptFin(1) - (medY * 0.5#)      'IIf(ptFin(1) > 0, (medY / 2), -(medY / 2))
                ElseIf oBlr > 270 Then
                    ptFin(0) = ptFin(0) - medX      'IIf(ptFin(0) > 0, (medX * 2), -(medX * 2))
                    ptFin(1) = ptFin(1) - medY      'IIf(ptFin(1) > 0, medY, -medY)
                End If
                ''
                '' Mover la imagen
                If ptFin(0) <> oImgFin.Origin(0) Or ptFin(1) <> oImgFin.Origin(1) Then
                    oImgFin.Move(oImgFin.Origin, ptFin)
                    oImgFin.Update()
                End If
            End If
            ''borramos el bloque.
            oBl.Delete()
            oBl = Nothing
            oImgFin = Nothing
        Next
        ''
        oApp.ActiveDocument.Regen(AcRegenType.acAllViewports)
        oDoc.Save()
    End Sub
    ''
    Public Sub CapaCreaActivaBalizamientoSuelo()
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
            oLayer = oApp.ActiveDocument.Layers.Item("BALIZAMIENTO SUELO")
        Catch ex As System.Exception
            oLayer = oApp.ActiveDocument.Layers.Add("BALIZAMIENTO SUELO")
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
        oApp.ActiveDocument.ActiveLayer = oLayer    ' oApp.ActiveDocument.Layers.Item("BALIZAMIENTO SUELO")
        oLayer = Nothing
        oColor = Nothing
        ''
        oApp.ActiveDocument.Regen(AcRegenType.acAllViewports)
    End Sub
    ''
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
                oLayer = oApp.ActiveDocument.Layers.Item("Ruta evacuación primaria")
            Catch ex As System.Exception
                oLayer = oApp.ActiveDocument.Layers.Add("Ruta evacuación primaria")
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
    ''
    ''
    Public Sub CapaCreaActivaTablaDatos()
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
        oApp.ActiveDocument.ActiveLayer = oApp.ActiveDocument.Layers.Item("0")
        ''
        '' Coger la capa BALIZAMIENTO SUELO o crearla
        Dim oLayer As AcadLayer = Nothing
        Try
            oLayer = oApp.ActiveDocument.Layers.Item("capacuadro")
        Catch ex As System.Exception
            oLayer = oApp.ActiveDocument.Layers.Add("capacuadro")
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
        oApp.ActiveDocument.ActiveLayer = oLayer    ' oApp.ActiveDocument.Layers.Item("BALIZAMIENTO SUELO")
        oLayer = Nothing
        oColor = Nothing
        ''
        oApp.ActiveDocument.Regen(AcRegenType.acAllViewports)
    End Sub
    ''
    ''
    Public Sub CapaCreaActivaBalizamientoPared()
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
        oApp.ActiveDocument.ActiveLayer = oApp.ActiveDocument.Layers.Item("0")
        ''
        '' Coger la capa BALIZAMIENTO SUELO o crearla
        Dim oLayer As AcadLayer = Nothing
        Try
            oLayer = oApp.ActiveDocument.Layers.Item("BALIZAMIENTO PARED")
        Catch ex As System.Exception
            oLayer = oApp.ActiveDocument.Layers.Add("BALIZAMIENTO PARED")
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
        oApp.ActiveDocument.ActiveLayer = oLayer    ' oApp.ActiveDocument.Layers.Item("BALIZAMIENTO PARED")
        oLayer = Nothing
        oColor = Nothing
        ''oApp.ActiveDocument.SendCommand("_line ")
        oApp.ActiveDocument.Regen(AcRegenType.acAllViewports)
    End Sub
    ''
    ''
    Public Sub CapaCreaActivaBalizamientoEscalera()
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
        CapaCeroActiva()
        ''
        '' Coger la capa BALIZAMIENTO SUELO o crearla
        Dim oLayer As AcadLayer = Nothing
        Try
            oLayer = oApp.ActiveDocument.Layers.Item("BALIZAMIENTO ESCALERA")
        Catch ex As System.Exception
            oLayer = oApp.ActiveDocument.Layers.Add("BALIZAMIENTO ESCALERA")
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
        oApp.ActiveDocument.ActiveLayer = oLayer    ' oApp.ActiveDocument.Layers.Item("BALIZAMIENTO PARED")
        oLayer = Nothing
        oColor = Nothing
        ''oApp.ActiveDocument.SendCommand("_line ")
        oApp.ActiveDocument.Regen(AcRegenType.acAllViewports)
    End Sub

    Public Sub CapaCeroActiva()
        '' Pondremos la capa 0 como activa.
        If oApp Is Nothing Then _
        oApp = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
        '' Activar Capa 0 primero.
        oApp.ActiveDocument.ActiveLayer = oApp.ActiveDocument.Layers.Item("0")
        ''
    End Sub

    Private Sub oTimer_Elapsed(sender As Object, e As Timers.ElapsedEventArgs) Handles oTimer.Elapsed
        '' Pondremos la capa 0 como activa.
        If oApp Is Nothing Then _
        oApp = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
        '' Activar Capa 0 primero.
        oApp.ActiveDocument.ActiveLayer = oApp.ActiveDocument.Layers.Item("0")
        ''
    End Sub
    ''
    Public Function GetPuntoDame_NET(mensaje As String) As Object
        Dim resultado As Object = Nothing
        '' Get the current database and start the Transaction Manager
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = acDoc.Editor
        Dim acCurDb As Database = acDoc.Database
        Dim pPtRes As PromptPointResult
        Dim pPtOpts As PromptPointOptions = New PromptPointOptions("")
        '' Prompt for the start point
        pPtOpts.Message = vbLf & mensaje
        pPtOpts.AllowNone = True    ' Permite terminar con boton derecho.
        Using oUi As EditorUserInteraction = ed.StartUserInteraction(Application.MainWindow.Handle)
            ' Using (acDoc.LockDocument) '' Necesario para dialogos modales.
            '' Poner el foco en la zona de dibujo
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
            pPtRes = acDoc.Editor.GetPoint(pPtOpts)
            Dim ptStart As Autodesk.AutoCAD.Geometry.Point3d = pPtRes.Value

            '' Exit if the user presses ESC or cancels the command
            If pPtRes.Status = PromptStatus.Cancel Then
                resultado = Nothing
            ElseIf pPtRes.Status = PromptStatus.OK Then
                resultado = ptStart
            End If
            ''
            oUi.End()
        End Using
        acDoc = Nothing
        acCurDb = Nothing
        Return resultado
    End Function
    ''
    Public Function GetOpcionDame_NET() As String
        Dim resultado As String = ""
        ''
        If oApp Is Nothing Then _
        oApp = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
        'oDoc = oApp.ActiveDocument
        oApp.ActiveDocument.Activate()
        ''

        ' Define the list of valid keywords
        Dim listaopciones As String
        listaopciones = "Ascendente Descendente"
        oApp.ActiveDocument.Utility.InitializeUserInput(1, listaopciones)

        ' Prompt para que el usuario introduzca una palabra. Return "Ascendente", "Descendente" en
        ' resultado variable o puede introducir "A", "D".
        'Dim returnString As String
        resultado = oApp.ActiveDocument.Utility.GetKeyword(vbLf & "Tipo de Escalera [Ascendente/Descendente] : ")
        'MsgBox("You entered " & resultado, , "GetKeyword Example")
        Return resultado
    End Function
    ''
    Public Function GetOpcionPrimariaSecundariaDame() As String
        Dim resultado As String = ""
        ''
        If oApp Is Nothing Then _
        oApp = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
        'oDoc = oApp.ActiveDocument
        oApp.ActiveDocument.Activate()
        ''

        ' Define the list of valid keywords
        Dim listaopciones As String
        listaopciones = "Primaria Accesibilidad"
        oApp.ActiveDocument.Utility.InitializeUserInput(0, listaopciones)   '' 0 para poder contestar con return.

        ' Prompt para que el usuario introduzca una palabra. Return "Primaria", "Secundaria" en
        ' resultado variable o puede introducir "P", "S".
        'Dim returnString As String
        resultado = oApp.ActiveDocument.Utility.GetKeyword(vbLf & "Tipo de Ruta evacuación [Primaria/Accesibilidad] <Primaria>: ")
        'MsgBox("You entered " & resultado, , "GetKeyword Example")
        If resultado = "" Then resultado = "Primaria"
        Return resultado
    End Function
    ''
    Public Function DameOpcionTexto(mensaje As String, queopciones As String()) As String
        Dim resultado As String = ""
        ''
        If oApp Is Nothing Then _
        oApp = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
        'oDoc = oApp.ActiveDocument
        oApp.ActiveDocument.Activate()
        ''

        ' Define the list of valid keywords
        Dim listaopciones As String = String.Join(" ", queopciones)
        'listaopciones = "Primaria Secundaria"
        oApp.ActiveDocument.Utility.InitializeUserInput(0, listaopciones)   '' 0 para poder contestar con return.

        ' Prompt para que el usuario introduzca una palabra del array. queopciones(0) sera el valor por defecto.
        Dim cadenaopciones As String = String.Join("/", queopciones)
        Dim cadenadefecto As String = queopciones(0)
        resultado = oApp.ActiveDocument.Utility.GetKeyword(vbLf & mensaje & " [" & cadenaopciones & "] <" & cadenadefecto & ">: ")
        'MsgBox("You entered " & resultado, , "GetKeyword Example")
        If resultado = "" Then
            resultado = cadenadefecto
        ElseIf InStr(listaopciones, resultado) = 0 Then
            resultado = ""
        End If
        ''
        Return resultado
    End Function
    ''
    ''
    Public Function GetDINDame_NET() As String
        Dim resultado As String = ""
        ''
        If oApp Is Nothing Then _
        oApp = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
        'oDoc = oApp.ActiveDocument
        oApp.ActiveDocument.Activate()
        '' Define the list of valid keywords
        Dim listaopciones As String
        listaopciones = "A4 A3 A2 A1 A0"
        'oDoc.Utility.InitializeUserInput(1, listaopciones)
        oApp.ActiveDocument.Utility.InitializeUserInput(7, listaopciones)

        ' Prompt para que el usuario introduzca una palabra. Return "Ascendente", "Descendente" en
        ' resultado variable o puede introducir "A", "D".
        'Dim returnString As String
        resultado = oApp.ActiveDocument.Utility.GetKeyword(vbLf & "Formato DIN [A4/A3/A2/A1/A0] : ")
        'MsgBox("You entered " & resultado, , "GetKeyword Example")
        Return resultado
    End Function
    ''
    '' Le daremos array de ancho, largo (A4, A3, A2, A1 o A0), el largo actual (X) y el ancho actual (Y)
    '' Nos devolverá qué escala tenemos que aplicar para encajarlo en el papel indicado.
    Public Function DameEscala(queDin As Array, queL As Double, queA As Double) As Double
        Dim resultado As Double = 1
        Dim escalaX As Double
        Dim escalaY As Double
        Dim largoIni As Double
        Dim anchoIni As Double
        Dim largoFin As Double
        Dim anchoFin As Double
        ''
        If queL > queA Then     '' En Horizontal
            largoIni = queDin(0)
            anchoIni = queDin(1)
        ElseIf queL < queA Then '' En Vertical
            largoIni = queDin(1)
            anchoIni = queDin(0)
        End If
        ''
        escalaX = Math.Abs(largoIni / queL)
        escalaY = Math.Abs(anchoIni / queA)
        largoFin = queL * escalaX
        anchoFin = queA * escalaY
        ''
        If queL > largoIni Then
            '' el valor menor. Porque hay que reducir
            resultado = Math.Max(escalaX, escalaY)
        Else
            '' el valor mayor. Porque hay que ampliar
            resultado = Math.Min(escalaX, escalaY)
        End If
        ''
        Return resultado
    End Function
    ''
    Public Sub CierraDibujo(queFullname As String)
        If oApp Is Nothing Then _
    oApp = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
        '' Activar Capa 0 primero.
        For Each queDoc As AcadDocument In oApp.Documents
            If queDoc.FullName = queFullname Then
                queDoc.Close()
                Exit For
            End If
        Next
    End Sub

    Public Sub CierraDibujoTodos()
        If oApp Is Nothing Then _
oApp = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
        'oApp.Documents.Close()
        '' Activar Capa 0 primero.
        For Each queDoc As AcadDocument In oApp.Documents
            queDoc.Activate()
            queDoc.Save()
            'DameDependencias()
        Next
        oApp.Documents.Close()
    End Sub
    ''
    Public Function XrefComprueba() As Hashtable
        Dim resultado As Hashtable = New Hashtable
        If oApp Is Nothing Then _
oApp = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
        ''
        Dim mensaje As String = ""
        For Each oEref As AcadBlock In oApp.ActiveDocument.Blocks
            If TypeOf oEref Is RasterImage Then
                Debug.Print(oEref.Name)
            End If
            If oEref.IsXRef AndAlso IO.File.Exists(oEref.Path) = False Then
                If resultado.Contains(oEref.EffectiveName) = False Then
                    resultado.Add(oEref.EffectiveName, oEref.Path)
                    mensaje &= oEref.EffectiveName & " / " & oEref.Path & vbCrLf
                End If
            End If
        Next
        ''
        If mensaje <> "" Then MsgBox(mensaje)
        Return resultado
    End Function

    ''
    Public Function XrefImagenDame() As Hashtable
        Dim resultado As Hashtable = New Hashtable
        If oApp Is Nothing Then _
oApp = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
        ''
        Dim mensaje As String = ""
        Dim oDict As AcadDictionary
        oDict = oApp.ActiveDocument.Dictionaries.Item("ACAD_IMAGE_DICT")
        Dim oImage As AcadRasterImage = Nothing
        Dim oImageDef As Object
        MsgBox(oDict.Count)
        For Each oImageDef In oDict
            'If TypeOf oImageDef Is AcadRasterImage Then
            MsgBox(oImageDef.GetType.ToString)
            'oImage = oImageDef
            'If resultado.Contains(oImage.Name) = False Then
            '    resultado.Add(oImage.Name, oImage)
            '    mensaje &= oImage.Name & " / " & oImage.ImageFile & vbCrLf
            'End If
            'End If
        Next
        ''
        If mensaje <> "" Then MsgBox(mensaje)
        Return resultado
    End Function
    ''
    Public Function UltimoObjeto() As AcadEntity
        If oApp Is Nothing Then _
        oApp = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
        ''
        Return oApp.ActiveDocument.ModelSpace.Item(oApp.ActiveDocument.ModelSpace.Count - 1)
    End Function
    ''
    Public Function DameBloque(quenombre As String) As AcadBlockReference
        Dim oBl As AcadBlockReference = Nothing
        If oApp Is Nothing Then _
        oApp = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
        ''
        For Each oEnt As AcadEntity In oApp.ActiveDocument.ModelSpace
            If TypeOf oEnt Is AcadBlockReference Then
                oBl = oApp.ActiveDocument.ObjectIdToObject(oEnt.ObjectID)
                If oBl.Name = quenombre Or oBl.EffectiveName = quenombre Then
                    Exit For
                Else
                    oBl = Nothing
                End If
            End If
        Next
        Return oBl
    End Function

    Public Sub XRef_DWGListar(Optional ConMensaje As Boolean = False)
        Dim doc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor
        Dim mensaje As String = "Listado de Xrefs DWG:" & vbCrLf & vbCrLf
        Dim cOid As New ObjectIdCollection
        Using lock As DocumentLock = doc.LockDocument
            Using tx As Transaction = db.TransactionManager.StartTransaction
                db.ResolveXrefs(True, False)
                Dim bTable As BlockTable = tx.GetObject(db.BlockTableId, OpenMode.ForRead)
                For Each oId As ObjectId In bTable
                    Dim bTr As BlockTableRecord = tx.GetObject(oId, OpenMode.ForRead)
                    If bTr IsNot Nothing AndAlso bTr.IsFromExternalReference Then
                        mensaje &= vbTab & bTr.Name & " = " & bTr.PathName & vbCrLf
                        ' Solo procesaremos los que contengan IMPLACAD
                        If IO.File.Exists(bTr.PathName) = True Then Continue For
                        '
                        Dim pathOld As String = IO.Path.GetFullPath(bTr.PathName)
                        Dim partes() As String = pathOld.Split("\")
                        partes(2) = "IMPLACAD"
                        Dim pathNew As String = String.Join("\", partes)
                        If IO.File.Exists(pathNew) = False Then
                            Dim lista As String() = IO.Directory.GetFiles(IMPLACAD_DATA, IO.Path.GetFileName(pathOld), SearchOption.AllDirectories)
                            If lista IsNot Nothing AndAlso lista.Length > 0 Then
                                pathNew = lista(0)
                            End If
                        End If
                        If IO.File.Exists(pathNew) Then
                            bTr.UpgradeOpen()
                            bTr.PathName = pathNew  '.Replace("\", "/")
                            mensaje &= vbTab & IO.Path.GetFileNameWithoutExtension(pathNew) & " = NEW(" & pathNew & ")" & vbCrLf
                            cOid.Add(bTr.ObjectId)
                        Else
                            Dim nombre As String = IO.Path.GetFileName(pathNew)
                            ' Si no existe el fichero PNG. Lo buscamos
                            Dim fi As String() = IO.Directory.GetFiles(IMPLACAD_DATA, nombre, SearchOption.AllDirectories)
                            If fi IsNot Nothing AndAlso fi.Count > 0 Then
                                pathNew = fi.First
                            End If
                        End If
                    End If
                Next
                Try
                    If cOid.Count > 0 Then db.ReloadXrefs(cOid)
                Catch ex As System.Exception

                End Try
                tx.Commit()
            End Using
        End Using
        If ConMensaje Then MsgBox(mensaje)
    End Sub
    Public Sub XRef_IMGListar(Optional ConMensaje As Boolean = False)
        Dim doc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor
        Dim imgDefs As New Dictionary(Of ObjectId, RasterImageDef)
        Dim cOid As New ObjectIdCollection
        Dim mensaje As String = "Listado de Xrefs IMG:" & vbCrLf & vbCrLf
        Using lock As DocumentLock = doc.LockDocument
            Using tx As Transaction = db.TransactionManager.StartTransaction
                Dim dNames As DBDictionary = tx.GetObject(db.NamedObjectsDictionaryId, OpenMode.ForRead)
                Dim PDF_DEFINITION As String = "ACAD_PDFDEFINITIONS"
                Dim DWF_DEFINITION As String = "ACAD_DWFDEFINITIONS"
                Dim ACAD_FIELDLIST As String = "ACAD_FIELDLIST"
                Dim idImgDict As ObjectId = RasterImageDef.GetImageDictionary(db)
                'Dim imgKeyName As String = Database.UnderlayDefinition.GetDictionaryKey(TypeOf  AcDb.PdfDefinition));
                Dim imgDic As DBDictionary = tx.GetObject(idImgDict, OpenMode.ForWrite)
                For Each entry As DictionaryEntry In imgDic
                    Dim id As ObjectId = entry.Value
                    Dim imgDef As RasterImageDef = tx.GetObject(id, OpenMode.ForRead)
                    mensaje &= vbTab & imgDef.SourceFileName & vbCrLf
                    If IO.File.Exists(imgDef.SourceFileName) = True Then
                        Continue For
                    End If
                    '
                    Dim pathOld As String = IO.Path.GetFullPath(imgDef.SourceFileName)
                    Dim partes() As String = pathOld.Split("\")
                    partes(2) = "IMPLACAD"
                    Dim pathNew As String = String.Join("\", partes)
                    If IO.File.Exists(pathNew) = False Then
                        Dim lista As String() = IO.Directory.GetFiles(IMPLACAD_DATA, IO.Path.GetFileName(pathOld), SearchOption.AllDirectories)
                        If lista IsNot Nothing AndAlso lista.Length > 0 Then
                            pathNew = lista(0)
                        End If
                    End If
                    If IO.File.Exists(pathNew) Then
                        imgDef.UpgradeOpen()
                        imgDef.SourceFileName = pathNew
                        imgDef.Load()
                        mensaje &= vbTab & IO.Path.GetFileNameWithoutExtension(pathNew) & " = NEW(" & pathNew & ")" & vbCrLf
                    Else
                        Dim nombre As String = IO.Path.GetFileName(pathNew)
                        ' Si no existe el fichero PNG. Lo buscamos
                        Dim fi As String() = IO.Directory.GetFiles(IMPLACAD_DATA, nombre, SearchOption.AllDirectories)
                        If fi IsNot Nothing AndAlso fi.Count > 0 Then
                            pathNew = fi.First
                        End If
                    End If
                Next
                tx.Commit()
            End Using
        End Using
        If ConMensaje Then MsgBox(mensaje)
    End Sub
    'Public Sub XRef_Listar()
    '    Dim doc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
    '    Dim db As Database = doc.Database
    '    Dim ed As Editor = doc.Editor
    '    Dim mensaje As String = "Listado de Xrefs:" & vbCrLf & vbCrLf
    '    Using tx As Transaction = db.TransactionManager.StartTransaction
    '        db.ResolveXrefs(True, False)
    '        Dim xg As XrefGraph = db.GetHostDwgXrefGraph(True)
    '        If xg Is Nothing Then Exit Sub
    '        'mensaje &= "XRef dibujo actual (Graph)" & vbCrLf
    '        Dim root As GraphNode = xg.RootNode
    '        If root Is Nothing Then Exit Sub
    '        ' Recursivamente
    '        XRef_Listar_Recursivo(mensaje, root, vbTab, ed, tx)
    '        MsgBox(mensaje)
    '    End Using
    'End Sub

    'Public Sub XRef_Listar_Recursivo(ByRef mensaje As String, root As GraphNode, separador As String, ed As Editor, ByRef tx As Transaction)
    '    For x As Integer = 0 To root.NumOut - 1
    '        Dim child As XrefGraphNode = root.Out(x)
    '        If child.XrefStatus = XrefStatus.Resolved Then
    '            Dim bl As BlockTableRecord = tx.GetObject(child.BlockTableRecordId, OpenMode.ForRead)
    '            mensaje &= separador & child.Database.Filename
    '            If bl.IsFromExternalReference Then
    '                mensaje &= " (" & bl.PathName & ")" & vbCrLf
    '            End If
    '            XRef_Listar_Recursivo(mensaje, child, separador & "| ", ed, tx)
    '        ElseIf child.XrefStatus = XrefStatus.FileNotFound Then

    '        End If
    '    Next
    'End Sub

    Public Sub PATHS_Set(New_paths As List(Of String))
        Dim acad_pref As AcadPreferences = Autodesk.AutoCAD.ApplicationServices.Application.Preferences
        Dim C_Paths As String = LCase(acad_pref.Files.SupportPath)

        Dim Old_Path_Ary As List(Of String) = New List(Of String)
        Old_Path_Ary = C_Paths.Split(";").ToList

        'Dim New_paths As List(Of String) = New List(Of String)

        'New_paths.Add("C:\Program files\Folder1")
        'New_paths.Add("C:\Program files\Folder2")
        'New_paths.Add("C:\Program files\Folder3")
        'New_paths.Add("C:\Program files\Folder4")

        For Each Str As String In New_paths
            If Not Old_Path_Ary.Contains(Str) Then
                Old_Path_Ary.Add(Str)
            End If
        Next
        acad_pref.Files.SupportPath = String.Join(";", Old_Path_Ary.ToArray())
    End Sub
End Module

Public Enum queCapa As Integer
    BALIZAMIENTOSUELO = 0
    BALIZAMIENTOPARED = 1
End Enum
''
Public Enum queTipoBE As Integer
    Todas = 0
    Ascendente = 1
    Descendente = 2
End Enum
''
Public Enum CapaEstado As Integer
    Desactivar = 0
    Activar = 1
    Inverso = 2
End Enum
''
Public Enum TipoEvacuacion As Integer
    Primaria = 0
    Accesibilidad = 1
End Enum
''
Public Enum ActivarBotones As Integer
    NoActivado = 1
    SiActivadoSinEscala = 2
    SiActivadoConEscala = 3
    Cualquiera = 4
End Enum
''
Public Enum Estado As Integer
    ActiNo = 0
    ActiSi = 1
    ActuNo = 2
    ActuSi = 3
    EscalaMSi = 4
    EscalaMNo = 5
    ExpressToolsNo = 6
    ExpressToolsSi = 7
End Enum