Imports System.Linq
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Interop
Imports Autodesk.AutoCAD.Interop.Common

Module modTablas

#Region "TABLADATOS"
    '' Insertar tabla para datos de etiquetas o la actualiza si ya existe.
    Public Sub TablaDatos_Inserta()
        ConfiguraDibujo()
        ''
        '' Comprobar si ya existe la tabla de datos
        Dim TablaDatosTemp As AcadTable = TablaDatos_Dame()
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
            TablaDatos_Actualiza()
            '' Regenerar la tabla
            TablaDatos.RegenerateTableSuppressed = False
            TablaDatos.RecomputeTableBlock(True)
        Else
            TablaDatos = CType(oApp.ActiveDocument.ObjectIdToObject(TablaDatosTemp.ObjectID), AcadTable) 'TablaDatosTemp ' oDoc.ObjectIdToObject(TablaDatosTemp.ObjectID)
            'Temporalmente desactivamos el recalculo de la tabla
            TablaDatos.RegenerateTableSuppressed = True
            TablaDatos_Actualiza()
            '' Regenerar la tabla
            TablaDatos.RegenerateTableSuppressed = False
            TablaDatos.RecomputeTableBlock(True)
            If vermensajes = True Then MsgBox("Tabla DATOS actualizada porque ya existía...")
        End If
        TablaDatos = Nothing
        'vermensajes = True
    End Sub

    Public Sub TablaDatos_Actualiza()
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

    Public Function TablaDatos_Dame() As AcadTable
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
            oSelTemp = oApp.ActiveDocument.SelectionSets.Add(nSel)
        Catch ex As System.Exception
            oSelTemp = oApp.ActiveDocument.SelectionSets.Item(nSel)
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
                    Dim texto As String = XData.XLeeDato(CType(oEnt, AcadObject), "Clase")
                    If TypeOf oEnt Is AcadTable And texto = "tabladatos" Then  '"Clase=tabla"
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
#Region "TABLADATOSPARCIAL"
    '' Insertar tabla para datos de etiquetas o la actualiza si ya existe.
    Public Sub TablaDatosParcial_Inserta()
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
        TablaDatosParcial_Actualiza()
        '' Regenerar la tabla
        TablaDatosParcial.RegenerateTableSuppressed = False
        TablaDatosParcial.RecomputeTableBlock(True)
        TablaDatosParcial = Nothing
        'vermensajes = True
    End Sub
    Public Sub TablaDatosParcial_Actualiza()
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
    Public Sub TablaEscaleras_Inserta()
        If BalizasEscaleras_Dame(queTipoBE.Todas).Count = 0 Then
            If vermensajes = True Then MsgBox("No existen balizas escaleras en el dibujo...")
            Exit Sub
        End If
        ConfiguraDibujo()
        ''
        '' Comprobar si ya existe la tabla de datos
        Dim TablaEscalerasTemp As AcadTable = TablaEscaleras_Dame()
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
            TablaEscaleras.SetText(1, 1, BalizasEscaleras_Dame(queTipoBE.Todas).Count)     '' Fila 1.1 = Total escaleras
            TablaEscaleras.SetText(2, 0, "  - Asc. BA 0307L (m)")             '' Fila 2.0 = Texto total
            'TablaEscaleras.SetText(2, 1, DameBalizasEscaleras(queTipoBE.Ascendente).Count)     '' Fila 1.1 = Total escaleras
            TablaEscaleras.SetText(2, 1, "0")     '' Fila 1.1 = Total escaleras
            TablaEscaleras.SetText(3, 0, "  - Des. BA 0401L (m)")             '' Fila 3.0 = Texto total
            'TablaEscaleras.SetText(3, 1, DameBalizasEscaleras(queTipoBE.Descendente).Count)     '' Fila 1.1 = Total escaleras
            TablaEscaleras.SetText(3, 1, "0")     '' Fila 1.1 = Total escaleras
            '' Actualizamos la tabla ESCALERAS
            TablaEscaleras_Actualiza()
            '' Regenerar la tabla
            TablaEscaleras.RegenerateTableSuppressed = False
            TablaEscaleras.RecomputeTableBlock(True)
        Else
            TablaEscaleras = CType(oApp.ActiveDocument.ObjectIdToObject(TablaEscalerasTemp.ObjectID), AcadTable) 'TablaEscalerasTemp ' oDoc.ObjectIdToObject(TablaDatosTemp.ObjectID)
            'Temporalmente desactivamos el recalculo de la tabla
            TablaEscaleras.RegenerateTableSuppressed = True
            TablaEscaleras_Actualiza()
            '' Regenerar la tabla
            TablaEscaleras.RegenerateTableSuppressed = False
            TablaEscaleras.RecomputeTableBlock(True)
            If vermensajes = True Then MsgBox("Tabla ESCALERAS actualizada porque ya existía...")
        End If
        TablaEscaleras = Nothing
        'vermensajes = True
    End Sub

    Public Sub TablaEscaleras_Actualiza()
        ''
        '' Ahora actualizaremos el total de balizas escalera y sus datos.
        TablaEscaleras.SetText(1, 1, BalizasEscaleras_Dame(queTipoBE.Todas).Count)     '' Fila 1.1 = Total escaleras
        '' Borrar todas las filas superiores a la 2
        If TablaEscaleras.Rows > 4 Then
            TablaEscaleras.DeleteRows(4, TablaEscaleras.Rows - 4)
        End If
        ''
        '' Ahora creamos las filas que sean necesarias.
        Dim arrBEscaleras As ArrayList = BalizasEscaleras_Dame()
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

    Public Function TablaEscaleras_Dame() As AcadTable
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
            oSelTemp = oApp.ActiveDocument.SelectionSets.Add(nSel)
        Catch ex As System.Exception
            oSelTemp = oApp.ActiveDocument.SelectionSets.Item(nSel)
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
                    Dim texto As String = XData.XLeeDato(CType(oEnt, AcadObject), "Clase")
                    If TypeOf oEnt Is AcadTable And texto = "tablaescaleras" Then     ' "Clase=tablaescaleras"
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

    Public Function BalizasEscaleras_Dame(Optional queTipo As queTipoBE = queTipoBE.Todas) As ArrayList
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
            oSelTemp = oApp.ActiveDocument.SelectionSets.Add(nSel)
        Catch ex As System.Exception
            oSelTemp = oApp.ActiveDocument.SelectionSets.Item(nSel)
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
                Dim texto As String = XData.XLeeDato(oEnt, "Clase")
                If TypeOf oEnt Is AcadBlockReference And texto = "balizaescalera" Then        '"Clase=balizaescalera"
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


#Region "TABLAZONAS"
    '' Insertar tabla para datos de etiquetas o la actualiza si ya existe.
    Public Sub TablaZonas_Inserta()
        ConfiguraDibujo()
        ''
        '' Comprobar si ya existe la tabla de datos
        Dim TablaZonasTemp As AcadTable = TablaZonas_Dame()
        ''
        '' Si no existe, la creamos. Y si existe la actualizamos.
        If TablaZonasTemp Is Nothing Then   ' And vermensajes = True Then
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
            CapaCreaActivaTablaZonas()
            ''
            '' Insertamos la tabla.
            'Dim oPunto As Object = oDoc.Utility.GetPoint(, "Punto de inserción de la Tabla")
            'TablaZonas = oDoc.ModelSpace.AddTable(oPunto, 6, 2, 50 * escala, 400 * escala)
            TablaZonas = oApp.ActiveDocument.ModelSpace.AddTable(oPunto, 6, 2, 1, 1)
            'Temporalmente desactivamos el recalculo de la tabla
            TablaZonas.RegenerateTableSuppressed = True
            ''
            '' Ponemos los textos en las cabeceras
            TablaZonas.SetText(0, 0, "IMPLACAD ES UN PRODUCTO IMPLASER")        '' Titulo tabla
            TablaZonas.SetText(1, 0, "** BALIZAMIENTO **")                  '' Cabecera
            TablaZonas.SetText(1, 1, "** METROS **")                '' Cabecera
            TablaZonas.SetText(2, 0, "SUELO")        ''
            TablaZonas.SetText(3, 0, "PARED")        ''
            TablaZonas.SetText(4, 0, "TOTAL SEÑALES --->")        '' Titulo tabla
            'TablaZonas.SetText(4, 1, DameTotalBloquesImplacad)        '' Titulo tabla
            TablaZonas.SetText(5, 0, "** REFERENCIA **")        '' Titulo tabla
            TablaZonas.SetText(5, 1, "** CANTIDAD **")        '' Titulo tabla
            '' Regenerar la tabla
            TablaZonas_Actualiza()
            '' Regenerar la tabla
            TablaZonas.RegenerateTableSuppressed = False
            TablaZonas.RecomputeTableBlock(True)
        Else
            TablaZonas = CType(oApp.ActiveDocument.ObjectIdToObject(TablaZonasTemp.ObjectID), AcadTable) 'TablaDatosTemp ' oDoc.ObjectIdToObject(TablaDatosTemp.ObjectID)
            'Temporalmente desactivamos el recalculo de la tabla
            TablaZonas.RegenerateTableSuppressed = True
            TablaZonas_Actualiza()
            '' Regenerar la tabla
            TablaZonas.RegenerateTableSuppressed = False
            TablaZonas.RecomputeTableBlock(True)
            If vermensajes = True Then MsgBox("Tabla DATOS actualizada porque ya existía...")
        End If
        TablaZonas = Nothing
        'vermensajes = True
    End Sub

    Public Sub TablaZonas_Actualiza()
        If colBloquesCantidad IsNot Nothing Then colBloquesCantidad.Clear()
        colBloquesCantidad = Nothing
        ''
        TablaZonas.SetText(2, 1, DameTotalBaliza(queCapa.BALIZAMIENTOSUELO))        '' Total Baliza Suelo
        TablaZonas.SetText(3, 1, DameTotalBaliza(queCapa.BALIZAMIENTOPARED))        '' Total Baliza Pared
        TablaZonas.SetText(4, 1, DameTotalBloquesImplacad)        '' Total etiquetas

        ''
        '' Ahora actualizaremos el listado de bloques y cantidades.
        '' Borrar todas las filas superiores a la 5
        If TablaZonas.Rows > 6 Then
            TablaZonas.DeleteRows(6, TablaDatos.Rows - 5)
        End If
        ''
        '' Ahora creamos las filas que sean necesarias.
        If colBloquesCantidad IsNot Nothing AndAlso colBloquesCantidad.Count > 0 Then
            Dim colIni As Integer = 6
            TablaZonas.InsertRows(colIni, 0.3, colBloquesCantidad.Count)
            ''
            For Each nBloque As String In colBloquesCantidad.Keys
                TablaZonas.SetText(colIni, 0, nBloque)        '' Nombre del bloque
                TablaZonas.SetText(colIni, 1, colBloquesCantidad(nBloque))  '' Cantidad de bloques.
                colIni += 1
            Next
            colBloquesCantidad.Clear()
            colBloquesCantidad = Nothing
        End If
        '' Formateamos la tabla.
        TablaZonas.TitleSuppressed = False              '' Activar Titulo
        TablaZonas.HeaderSuppressed = False             '' Activar Cabecera
        TablaZonas.RegenerateTableSuppressed = True     '' Regenerar/Suprimirla
        '' Margen de las celdas
        TablaZonas.VertCellMargin = 0.1                '' Margen vertical celda
        TablaZonas.HorzCellMargin = 0.25               '' Margen horizontal celda
        '' Altura de las filas de cabecera
        TablaZonas.SetRowHeight(0, 0.48)                '' Altura Fila 0
        TablaZonas.SetRowHeight(1, 0.4)                '' Altura Fila 1
        TablaZonas.SetRowHeight(2, 0.4)                '' Altura Fila 2
        TablaZonas.SetRowHeight(3, 0.4)                '' Altura Fila 3
        TablaZonas.SetRowHeight(4, 0.4)                '' Altura Fila 4
        TablaZonas.SetRowHeight(5, 0.4)                '' Altura Fila 5
        '' Anchura de las 2 columnas
        TablaZonas.SetColumnWidth(0, 3.5)               '' Ancho columna 1
        TablaZonas.SetColumnWidth(1, 3.5)               '' Ancho columna 2
        '' Estilo de texto
        TablaZonas.SetTextStyle(RowType.TitleRow, textoEstilo.Name)  '' Estilo de texto de titulo
        TablaZonas.SetTextStyle(RowType.HeaderRow, textoEstilo.Name) '' Estilo de texto de cabecera
        TablaZonas.SetTextStyle(RowType.DataRow, textoEstilo.Name)   '' Estilo de texto de datos
        '' Altura de los textos
        TablaZonas.SetTextHeight(RowType.TitleRow, 0.25)    ' 22 * escalaTotal)    '' Altura texto titulo=0.25
        TablaZonas.SetTextHeight(RowType.HeaderRow, 0.2)    ' 2 * escalaTotal)    '' Altura texto cabecera=0.2
        TablaZonas.SetTextHeight(RowType.DataRow, 0.2)      ' 2 * escalaTotal)    '' Altura texto celdas=0.2
        ''
        TablaZonas.SetAlignment(RowType.TitleRow, AcCellAlignment.acMiddleCenter)   '' Alineacion titulo
        TablaZonas.SetAlignment(RowType.HeaderRow, AcCellAlignment.acMiddleCenter)   '' Alineacion cabecera
        TablaZonas.SetAlignment(RowType.DataRow, AcCellAlignment.acMiddleCenter)   '' Alineacion datos
        ''
        '' Altura del resto de las filas con datos a 0.3
        For y As Integer = 6 To TablaDatos.Rows - 1
            TablaZonas.SetRowHeight(y, 0.3)
        Next
        ''
        '' Regenerar la tabla
        'TablaDatos.RegenerateTableSuppressed = False
        'TablaDatos.RecomputeTableBlock(True)
        XData.XNuevo(TablaZonas, "Clase=tablazonas")
        'TablaZonas = Nothing
    End Sub

    Public Function TablaZonas_Dame() As AcadTable
        Dim resultado As AcadTable = Nothing

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
        ' F1(1) = 1000 : F2(1) = "Clase=tablazonas"
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
            oSelTemp = oApp.ActiveDocument.SelectionSets.Add(nSel)
        Catch ex As System.Exception
            oSelTemp = oApp.ActiveDocument.SelectionSets.Item(nSel)
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
                    If TypeOf oEnt Is AcadTable = False Then Continue For
                    Dim texto As String = XData.XLeeDato(oEnt, "Clase")
                    If texto = "tablazonas" Then  '"Clase=tablazonas"
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
End Module
