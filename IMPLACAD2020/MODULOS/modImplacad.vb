﻿Imports System.Linq
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Interop
Imports Autodesk.AutoCAD.Interop.Common

Module modImplacad
    Public WithEvents oTimer As System.Timers.Timer

    Public Sub AutoCAD_PonFoco()
        ' Foco en AutoCAD
        Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
    End Sub

    Public Sub AutoCAD_PonFoco1()
        ' Foco en AutoCAD
        Call Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Window.Focus()
        'Dim doc As AcadDocument = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.GetAcadDocument
    End Sub

    Public Sub AutoCAD_PonFoco2()
        ' Foco en AutoCAD
        Call Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf)
        'Dim doc As AcadDocument = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.GetAcadDocument
    End Sub

    ''
    '' Configurar Dibujo Actual
    Public Sub ConfiguraDibujo()
        If oApp Is Nothing Then _
        oApp = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
        Using lock As DocumentLock = Eventos.AXDoc.LockDocument
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
            Try
                If Application.DocumentManager.MdiActiveDocument.Database.Insunits <> UnitsValue.Undefined Then
                    Application.DocumentManager.MdiActiveDocument.Database.Insunits = UnitsValue.Undefined
                End If
            Catch ex As System.Exception
            End Try

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
            'oApp.ActiveDocument.Regen(AcRegenType.acActiveViewport)
            Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("IMAGEFRAME", 0)
            '' Poner la escala de anotación 1:1
            Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("CANNOSCALE", "1:1")
            '' Poner la escala de los tipos de linea a 1
            Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("LTSCALE", 1)
            'oApp.ActiveDocument.Regen(AcRegenType.acActiveViewport)
            'CANNOSCALE
            'Call DameDependencias()
            'XrefComprueba()
            'XrefImagenDame()
        End Using
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

    '
    ' Nos dará todas las Polilineas en capa "Zonas"
    Public Function DamePolilineasZonasImplacad(Optional capa As String = "Zonas") As List(Of AcadLWPolyline)
        Dim resultado As New List(Of AcadLWPolyline)
        'Dim cSeleccion As AcadSelectionSets
        Dim F1(1) As Short
        Dim F2(1) As Object
        Dim vF1 As Object = Nothing
        Dim vF2 As Object = Nothing
        '
        'F1(0) = 1001 : F2(0) = regAPP
        F1(0) = 100 : F2(0) = "AcDbPolyline"   ',AcDbPolyline" '"*Polyline"
        F1(1) = 8 : F2(1) = capa
        'F1(3) = 1000 : F2(3) = "Clase=etiqueta"
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
            oSelTemp = oApp.ActiveDocument.SelectionSets.Item(nSel)
        Catch ex As System.Exception
            oSelTemp = oApp.ActiveDocument.SelectionSets.Add(nSel)
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
            For Each oEnt As AcadEntity In oSelTemp
                '' Si no es un bloque, continuamos
                If oEnt.EntityName <> "AcDbPolyline" Then Continue For
                'If Not (TypeOf oEnt Is AcadLWPolyline) Then Continue For
                Dim oPol As AcadLWPolyline = DirectCast(oApp.ActiveDocument.ObjectIdToObject(oEnt.ObjectID), AcadLWPolyline)
                If oPol Is Nothing Then Continue For
                resultado.Add(oPol)
                '' Si el nombre no empieza con EX, EV, SIA o KIT (arrpreEti) Borramos Xdata de regApp y Continuamos
                'Dim borraXdata As Boolean = True
                'For Each preEti As String In arrpreEti
                '    If oBl.Name.ToUpper.StartsWith(preEti) Then
                '        borraXdata = False
                '        Exit For
                '    End If
                'Next
                '
                'If borraXdata = True Then
                '    XData.XBorrar(oBl)
                '    Continue For
                'Else
                '    arrEnt.Add(oEnt)
                'End If
            Next
            oSelTemp.Clear()
        End If
        ''
        oSelTemp = Nothing
        Return resultado
    End Function

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
                Dim texto As String = XData.XLeeDato(oEnt, "Clase")
                If TypeOf oEnt Is AcadBlockReference And texto = "etiqueta" Then      '"Clase=etiqueta"
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
                Dim texto As String = XData.XLeeDato(oEnt, "Clase")
                If TypeOf oEnt Is AcadBlockReference And texto = "etiqueta" Then      '"Clase=etiqueta"
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
            'Return Format(Math.Truncate(resultado + 1), "##0.#")
            Return Format(Math.Truncate(resultado), "##0.#")
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
            oSelTemp = oApp.ActiveDocument.SelectionSets.Add(nSel)
        Catch ex As System.Exception
            oSelTemp = oApp.ActiveDocument.SelectionSets.Item(nSel)
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
            oSelTemp = oApp.ActiveDocument.SelectionSets.Add(nSel)
        Catch ex As System.Exception
            oSelTemp = oApp.ActiveDocument.SelectionSets.Item(nSel)
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
                Dim ImgNewName As String = ""
                Dim SUSTITUCION As String = oEtis.DSustituciones(bln).Trim.Replace(" ", "")
                If SUSTITUCION.Contains(";") Then
                    ImgNewName = SUSTITUCION.Split(";")(1) & ".png"
                Else
                    ImgNewName = SUSTITUCION & ".png"
                End If
                ' encontramos en oEtis.DSustituciones. Definimos su imagen
                imgPath = IO.Path.Combine(IMPLACAD_DATA, balizas, ImgNewName)
                encontrado = True
                'Else
                '    '' Buscaremos en colSustituciones para ver si nombre bloque empieza por (menos de 10)
                '    For Each queRef As String In oEtis.DSustituciones.Keys
                '        '' Si es >= 10 caracteres, saltamos al siguiente
                '        If queRef.Length >= 10 Then Continue For
                '        ''
                '        If bln.ToUpper.StartsWith(queRef.ToUpper) Then
                '            imgPath = IO.Path.Combine(IMPLACAD_DATA, balizas, oEtis.DSustituciones(queRef) & ".png")
                '            encontrado = True
                '            Exit For
                '        End If
                '    Next
            End If
            ' No es un bloque para sustituir. Lo borramos
            If encontrado = False Then
                oBl.Delete()
                Continue For
            End If
            ' No existe la imagen en IMPLACAD_DATA\balizas
            If imgPath = "" OrElse IO.File.Exists(imgPath) = False Then
                Continue For
            End If
            '
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
    Public Sub CambiaBloqueViejoPorNuevo(ByRef oDoc As AcadDocument)
        ''
        If oApp Is Nothing Then _
            oApp = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
        '' Activar Capa 0 primero.
        'oDoc.ActiveLayer = oDoc.Layers.Item("0")
        CapaCeroActiva()

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
        '
        ' 1.- Cogemos todos los bloques de IMPLACAD insertados
        Dim arrBlo As ArrayList = DameTodoImplacad()
        '
        ' 2.- Sacar el total de bloques viejos a sustituir y preguntar.
        Dim LBlo As New Dictionary(Of String, String)     ' Bloque viejo, path nuevo
        For Each oEnt As AcadEntity In arrBlo
            If Not (TypeOf oEnt Is AcadBlockReference) Then Continue For
            ''
            Dim oBl As AcadBlockReference = oApp.ActiveDocument.ObjectIdToObject(oEnt.ObjectID)
            Dim oBlName As String = oBl.EffectiveName.ToUpper
            Dim oBlNew As AcadBlockReference = Nothing
            ''
            ''
            '' Si lo encontramos en colSustituciones, cambiarlo.
            If oEtis.DSustituciones.ContainsKey(oBlName) Then
                Dim oBlNewName As String = ""
                Dim SUSTITUCION As String = oEtis.DSustituciones(oBlName).Trim.Replace(" ", "")
                If SUSTITUCION.Contains(";") Then
                    oBlNewName = SUSTITUCION.Split(";")(0) & ".dwg"
                Else
                    oBlNewName = SUSTITUCION & ".dwg"
                End If
                Dim busco As String() = IO.Directory.GetFiles(IMPLACAD_DATA, oBlNewName, SearchOption.AllDirectories)
                ' Si no lo hemos encontrado, continuar
                If busco Is Nothing OrElse busco.Count = 0 Then Continue For
                ' Encontrado
                Dim oBlNewPath As String = busco.First
                '
                If LBlo.ContainsKey(oBl.Handle) = False Then LBlo.Add(oBl.Handle, oBlNewPath)
            End If
        Next
        '
        If LBlo.Count > 0 Then
            If MsgBox(
                "Se han encontrado (" & LBlo.Count & ") etiquetas antiguas que se pueden sustituir por nuevas." & vbCrLf &
                "¿Desea sustituirlas?",
                MsgBoxStyle.Information Or MsgBoxStyle.YesNo,
                "SUSTITUIR ETIQUETAS ANTIGUAS") = MsgBoxResult.No Then
                Exit Sub
            End If
        Else
            MsgBox(
                "No se han encontrado referencias que se pueden sustituir.",
                MsgBoxStyle.Information Or MsgBoxStyle.OkOnly,
                "SUSTITUIR ETIQUETAS ANTIGUAS")
            Exit Sub
        End If
        '
        ' Recorrer cada etiqueta de bloque y cambiarla por la nueva etiqueta.
        For Each oEnt As String In LBlo.Keys
            Dim oBl As AcadBlockReference = oApp.ActiveDocument.HandleToObject(oEnt)
            Dim oBlNew As AcadBlockReference = Nothing
            ' Lo insertamos y ponemos las misma propiedades que oBl
            oBlNew = oDoc.ModelSpace.InsertBlock(oBl.InsertionPoint, LBlo(oEnt), oBl.XScaleFactor, oBl.YScaleFactor, oBl.ZScaleFactor, oBl.Rotation)
            If oBlNew IsNot Nothing Then
                'poner datos al bloque nuevo
                XData.XNuevo(oBlNew, "Clase=etiqueta")
                oBlNew = Nothing
                'borramos el bloque viejo.
                oBl.Delete()
                oBl = Nothing
            End If
        Next
        ''
        oApp.ActiveDocument.Regen(AcRegenType.acAllViewports)
        oDoc.Save()
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

    Public Sub HazZoomObjeto(ByVal obj As AcadObject, Optional reduce As Double = 1, Optional selecciona As Boolean = True)
        Dim pt1 As Object = Nothing
        Dim pt2 As Object = Nothing
        obj.GetBoundingBox(pt1, pt2)
        Dim distX As Double = Math.Abs(pt2(0) - pt1(0))
        Dim distY As Double = Math.Abs(pt2(1) - pt1(1))
        '
        pt1(0) -= (distX) * reduce : pt1(1) -= (distY) * reduce
        pt2(0) += (distX) * reduce : pt2(1) += (distY) * reduce
        oApp.ZoomWindow(pt1, pt2)
        If selecciona = True Then
            'Dim oIPrt As New IntPtr(obj.ObjectID)
            'Dim oId As New ObjectId(oIPrt)
            'Dim arrIds() As ObjectId = {oId}
            'Autodesk.AutoCAD.Internal.Utils.SelectObjects(arrIds)
            Selecciona_AcadObject(obj)
        End If
    End Sub

    Public Sub SeleccionaPorHandle(ThisDrawing As AcadDocument, objEnt As AcadEntity, Optional comando As String = "")
        Dim queHandle As String = objEnt.Handle
        Dim lisp As String = "(handent " & Chr(34) & queHandle & Chr(34) & ") "
        ThisDrawing.SendCommand("_SELECT ")
        ThisDrawing.SendCommand(lisp)
        ' Si hay un comando se ejecutará sobre la selección
        If comando = " " Or comando = "  " Or comando = "" Then
            ThisDrawing.SendCommand(comando)
        ElseIf comando <> "" Then
            If comando.Contains("[handle]") Then comando = comando.Replace("[handle]", lisp)
            If comando.Contains("[Handle]") Then comando = comando.Replace("[Hhandle]", lisp)
            If comando.Contains("[HANDLE]") Then comando = comando.Replace("[HANDLE]", lisp)
            ThisDrawing.SendCommand(comando & " ")
        End If
        ' Volver a seleccionar el objecto. Dara error si se ha borrado
        Try
            ThisDrawing.SendCommand("_SELECT ")
            ThisDrawing.SendCommand(lisp & vbCrLf)
            '(ssget "x" '((5 . "157")))
            'ThisDrawing.SendCommand("(ssget " & Chr(34) & "X" & Chr(34) & " '((5 . " & Chr(34) & objEnt.Handle & Chr(34) & "))) ")
            ' ThisDrawing.SendCommand("_SELECT (handent " & Chr(34) & queHandle & Chr(34) & ")  ")
            'ThisDrawing.SendCommand("(setq sel1 (ssget '((0 . " & Chr(34) & "INSERT" & Chr(34) & ")(5 . " & Chr(34) & objEnt.Handle & Chr(34) & "))))")
            'ThisDrawing.SendCommand("(ssget '((5 . " & Chr(34) & objEnt.Handle & Chr(34) & ")))")
            '(setq sel1 (ssget '((0 . "CIRCLE"))))
        Catch ex As System.Exception
            '
        End Try
    End Sub

    Public Sub Selecciona_AcadObject(objEnt As AcadObject)
        oApp.ActiveDocument.SetVariable("pickadd", 0)   ' Quitar la seleccion que hubiera.
        Dim queHandle As String = objEnt.Handle
        Dim obj As AcadObject = oApp.ActiveDocument.HandleToObject(queHandle)
        ' Volver a seleccionar el objecto. Dara error si se ha borrado
        Try
            Dim oIPrt As New IntPtr(obj.ObjectID)
            Dim oId As New ObjectId(oIPrt)
            Dim arrIds() As ObjectId = {oId}
            Autodesk.AutoCAD.Internal.Utils.SelectObjects(arrIds)
        Catch ex As System.Exception
            '
        End Try
        oApp.ActiveDocument.SetVariable("pickadd", 2)   ' Quitar la seleccion que hubiera.
    End Sub

    Public Sub Selecciona_AcadID(IdEnt As Long)
        oApp.ActiveDocument.SetVariable("pickadd", 0)   ' Quitar la seleccion que hubiera.
        Try
            Dim oIPrt As New IntPtr(IdEnt)
            Dim oId As New ObjectId(oIPrt)
            Dim arrIds() As ObjectId = {oId}
            Autodesk.AutoCAD.Internal.Utils.SelectObjects(arrIds)
        Catch ex As System.Exception
            '
        End Try
        oApp.ActiveDocument.SetVariable("pickadd", 2)   '' La seleccion actual se suma a la que hubiera.
    End Sub

    Public Sub Selecciona_AcadID(IdEnt As Long())
        oApp.ActiveDocument.SetVariable("pickadd", 0)   ' Quitar la seleccion que hubiera.
        Dim colIds As New List(Of ObjectId)
        For Each LongId As Long In IdEnt
            Dim oId As New ObjectId(New IntPtr(LongId))
            colIds.Add(oId)
        Next
        Try
            Autodesk.AutoCAD.Internal.Utils.SelectObjects(colIds.ToArray)
        Catch ex As System.Exception
            '
        End Try
        oApp.ActiveDocument.SetVariable("pickadd", 2)   '' La seleccion actual se suma a la que hubiera.
    End Sub

    Public Sub Selecciona_AcadID(IdEnt As ObjectId())
        oApp.ActiveDocument.SetVariable("pickadd", 0)   ' Quitar la seleccion que hubiera.
        Try
            Autodesk.AutoCAD.Internal.Utils.SelectObjects(IdEnt)
        Catch ex As System.Exception
            '
        End Try
        oApp.ActiveDocument.SetVariable("pickadd", 2)   '' La seleccion actual se suma a la que hubiera.
    End Sub

    ''' <summary>
    ''' Seleccionamos objetos indicando su tipo, capa, Xdata y
    ''' acadselectionset a utilizar "oSel" u "oSelTemp" (true o false)
    ''' </summary>
    ''' <param name="tipo">Tipo de opbjeto Autocad: POLYLINE, LWPOLILINE, AEC_WALL, INSERT</param>
    ''' <param name="capa">Nombre de la capa o modvariables.precapa + ultZona</param>
    ''' <param name="DatosX">True tendrá en cuenta los XData o False no los tendrá en cuenta</param>
    ''' <remarks></remarks>
    Public Function SeleccionaTodosObjetos_IDs(Optional ByVal tipo As Object = Nothing, Optional ByVal capa As Object = "", Optional ByVal DatosX As Boolean = False) As List(Of Long)
        Dim resultado As New List(Of Long)
        'Dim cSeleccion As AcadSelectionSets
        Dim F1(-1) As Short   'Dim F1(0 To 5) As Integer
        Dim F2(-1) As Object    'Dim F2(0 To 5) As Object
        Dim vF1 As Object
        Dim vF2 As Object
        ' 0 para tipo / 2 para nombre / 8 para capa
        ' tipo objeto o TODOS si no ponemos nadaDatosX.
        ' Siempre tiene que estar despues de entidad. Si no no funciona.
        ' "AEC_WALL" "LWPOLYLINE" "POLYLINE" "INSERT"
        If Not (tipo Is Nothing) Then
            ReDim Preserve F1(F1.Length)
            ReDim Preserve F2(F2.Length)
            F1(F1.Length - 1) = 0 : F2(F2.Length - 1) = tipo
        End If
        'F1(0) = 0 : F2(0) = tipo

        'DatosX Siempre tiene que estar despues de entidad. Si no no funciona
        If DatosX = True Then
            ReDim Preserve F1(F1.Length)
            ReDim Preserve F2(F2.Length)
            F1(F1.Length - 1) = 1001 : F2(F2.Length - 1) = regAPP   ' CType(regAPP, Object)
        End If

        If capa <> "" Then
            ReDim Preserve F1(F1.Length)
            ReDim Preserve F2(F2.Length)
            F1(F1.Length - 1) = 8 : F2(F2.Length - 1) = capa
        End If
        vF1 = F1
        vF2 = F2
        '
        Try
            oSel = oApp.ActiveDocument.SelectionSets.Add(nSel)
        Catch ex As System.Exception
            oSel = oApp.ActiveDocument.SelectionSets.Item(nSel)
        End Try
        '
        oSel.Clear()
        oSel.Select(AcSelect.acSelectionSetAll, , , vF1, vF2)
        '
        If oSel IsNot Nothing AndAlso oSel.Count > 0 Then
            For Each oEnt As AcadEntity In oSel
                If resultado.Contains(oEnt.ObjectID) = False Then resultado.Add(oEnt.ObjectID)
            Next
        End If
        '
        oSel.Clear()
        oSel.Delete()
        oSel = Nothing
        '
        Return resultado
    End Function

    ''' <summary>
    ''' Seleccionamos objetos indicando su tipo, capa, Xdata y
    ''' acadselectionset a utilizar "oSel" u "oSelTemp" (true o false)
    ''' </summary>
    ''' <param name="tipo">Tipo de opbjeto Autocad: POLYLINE, LWPOLILINE, AEC_WALL, INSERT</param>
    ''' <param name="capa">Nombre de la capa o modvariables.precapa + ultZona</param>
    ''' <param name="DatosX">True tendrá en cuenta los XData o False no los tendrá en cuenta</param>
    ''' <remarks></remarks>
    Public Function SeleccionaTodosObjetos_Handle(Optional ByVal tipo As Object = Nothing, Optional ByVal capa As Object = "", Optional ByVal DatosX As Boolean = False) As List(Of String)
        Dim resultado As New List(Of String)
        'Dim cSeleccion As AcadSelectionSets
        Dim F1(-1) As Short   'Dim F1(0 To 5) As Integer
        Dim F2(-1) As Object    'Dim F2(0 To 5) As Object
        Dim vF1 As Object
        Dim vF2 As Object
        ' 0 para tipo / 2 para nombre / 8 para capa
        ' tipo objeto o TODOS si no ponemos nadaDatosX.
        ' Siempre tiene que estar despues de entidad. Si no no funciona.
        ' "AEC_WALL" "LWPOLYLINE" "POLYLINE" "INSERT"
        If Not (tipo Is Nothing) Then
            ReDim Preserve F1(F1.Length)
            ReDim Preserve F2(F2.Length)
            F1(F1.Length - 1) = 0 : F2(F2.Length - 1) = tipo
        End If
        'F1(0) = 0 : F2(0) = tipo

        'DatosX Siempre tiene que estar despues de entidad. Si no no funciona
        If DatosX = True Then
            ReDim Preserve F1(F1.Length)
            ReDim Preserve F2(F2.Length)
            F1(F1.Length - 1) = 1001 : F2(F2.Length - 1) = regAPP   ' CType(regAPP, Object)
        End If

        If capa <> "" Then
            ReDim Preserve F1(F1.Length)
            ReDim Preserve F2(F2.Length)
            F1(F1.Length - 1) = 8 : F2(F2.Length - 1) = capa
        End If
        vF1 = F1
        vF2 = F2
        '
        Try
            oSel = oApp.ActiveDocument.SelectionSets.Add(nSel)
        Catch ex As System.Exception
            oSel = oApp.ActiveDocument.SelectionSets.Item(nSel)
        End Try
        '
        oSel.Clear()
        oSel.Select(AcSelect.acSelectionSetAll, , , vF1, vF2)
        '
        If oSel IsNot Nothing AndAlso oSel.Count > 0 Then
            For Each oEnt As AcadEntity In oSel
                If resultado.Contains(oEnt.Handle) = False Then resultado.Add(oEnt.Handle)
            Next
        End If
        '
        oSel.Clear()
        oSel.Delete()
        oSel = Nothing
        '
        Return resultado
    End Function

    Public Function SeleccionaTodosObjetosXData(nombreXData As String, valueXData As String, Optional igual As Boolean = False) As List(Of Long)
        Dim resultado As New List(Of Long)
        'Dim cSeleccion As AcadSelectionSets
        Dim F1(1) As Short   'Dim F1(0 To 5) As Integer
        Dim F2(1) As Object    'Dim F2(0 To 5) As Object
        Dim vF1 As Object
        Dim vF2 As Object
        ' Todos los tipos de objetos
        F1(0) = 0 : F2(0) = "*"
        ' DatosX Siempre tiene que estar despues de entidad. Si no no funciona
        F1(1) = 1001 : F2(1) = regAPP   ' CType(regAPP, Object)
        vF1 = F1
        vF2 = F2
        '
        Try
            oSel = oApp.ActiveDocument.SelectionSets.Add(nSel)
        Catch ex As System.Exception
            oSel = oApp.ActiveDocument.SelectionSets.Item(nSel)
        End Try
        '
        oSel.Clear()
        oSel.Select(AcSelect.acSelectionSetAll, , , vF1, vF2)
        '
        If oSel IsNot Nothing AndAlso oSel.Count > 0 Then
            For Each oEnt As AcadEntity In oSel
                Dim queGrupo As String = XLeeDato(oEnt.Handle, nombreXData)
                If queGrupo = "" Then Continue For
                If igual = True AndAlso queGrupo.ToUpper <> valueXData.ToUpper Then
                    Continue For
                ElseIf igual = False AndAlso queGrupo.ToUpper.Contains(valueXData.ToUpper) = False Then
                    Continue For
                End If
                '
                If resultado.Contains(oEnt.ObjectID) = False Then resultado.Add(oEnt.ObjectID)
            Next
        End If
        '
        oSel.Clear()
        oSel.Delete()
        oSel = Nothing
        '
        Return resultado
    End Function

    Public Function SeleccionaDameEntitiesEnCapa(queCapa As String) As ArrayList
        Dim resultado As New ArrayList
        ''
        'Add bit about counting text on layer
        Dim myEd As Editor = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor
        Dim myTVs(0) As TypedValue
        myTVs.SetValue(New TypedValue(DxfCode.LayerName, queCapa), 0)
        Dim myFilter As New SelectionFilter(myTVs)
        Dim myPSR As PromptSelectionResult = myEd.SelectAll(myFilter)
        Dim oSel As SelectionSet = Nothing
        If myPSR.Status = PromptStatus.OK Then
            oSel = myPSR.Value
        End If
        '    Dim myTVs(3) As TypedValue
        'myTVs.SetValue(New TypedValue(DxfCode.Operator, "<AND"), 0)
        'myTVs.SetValue(New TypedValue(DxfCode.Start, "TEXT"), 1)
        'myTVs.SetValue(New TypedValue(DxfCode.LayerName, "0"), 2)
        'myTVs.SetValue(New TypedValue(DxfCode.Operator, "AND>"), 3)
        ''myTVs(0) = New TypedValue(DxfCode.l .Start, "MTEXT")
        'Dim myFilter As New SelectionFilter(myTVs)
        '    Dim myPSR As PromptSelectionResult = myEd.SelectAll(myFilter)
        '    If myPSR.Status = PromptStatus.OK Then
        '    Dim mySS As SelectionSet = myPSR.Value
        'myForm.Label2.Text = mySS.Count
        'End If
        ''
        If oApp Is Nothing Then _
    oApp = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
        '
        oApp.ActiveDocument.SetVariable("pickadd", 0)   '' Solo una selección. Se quita lo que hubiera
        oApp.ActiveDocument.ActiveSelectionSet.Clear()
        '
        If oSel IsNot Nothing AndAlso oSel.Count > 0 Then
            For x As Integer = 0 To oSel.Count - 1
                resultado.Add(oApp.ActiveDocument.ObjectIdToObject(oSel.Item(x).ObjectId.OldIdPtr))
            Next
        Else
            resultado = Nothing
        End If
        ''
        oApp.ActiveDocument.SetVariable("pickadd", 2)   '' La seleccion actual se suma a la que hubiera.
        '
        oSel = Nothing
        Return resultado
        Exit Function
    End Function

    '
    ''' <summary>
    ''' Devuelve arrayList con todas las polilineas que cumplan el criterio
    ''' nombreApp = '*' por defecto. Le podemos indicar un nombre de APP registrada (1001=nombreApp)
    ''' nombrecapa = '*' por defecto. Le podemos indicar un nombre de capa
    ''' ** Le podemos indicar carácterés comodin (Ej.: nombrecapa=planta*) Utilizar * o ?
    ''' </summary>
    ''' <param name="nombreApp">Nombre de la app que registro el XData. O filtro con carácteres comodín</param>
    ''' <param name="nombrecapa">Nombre de la capa o filtro con carácteres comodin</param>
    ''' <returns></returns>
    Public Function SeleccionaDamePolilineasTODAS(Optional nombreApp As String = "*", Optional nombrecapa As String = "*") As ArrayList
        Dim resultado As New ArrayList
        'Dim cSeleccion As AcadSelectionSets
        Dim F1(3) As Short
        Dim F2(3) As Object
        Dim vF1 As Object = Nothing
        Dim vF2 As Object = Nothing

        '' Las 2 maneras valen igual. AcDbBlckReference es mejor (Solo coge bloques) INSERT coge sombreados también.
        F1(0) = 100 : F2(0) = "AcDbBlockReference"
        F1(1) = 0 : F2(1) = "LWPOLYLINE"
        F1(2) = 1001 : F2(2) = nombreApp
        F1(4) = 8 : F2(4) = nombrecapa
        ''
        vF1 = F1
        vF2 = F2
        '
        Try
            oSel = oApp.ActiveDocument.SelectionSets.Add(regAPP)
        Catch ex As System.Exception
            oSel = oApp.ActiveDocument.SelectionSets.Item(regAPP)
        End Try
        ''
        oSel.Clear()
        Try
            oSel.Select(AcSelect.acSelectionSetAll, , , vF1, vF2)
        Catch ex As System.Exception
            Debug.Print(ex.Message)
        End Try
        ''
        If oSel.Count > 0 Then
            For Each oEnt As AcadEntity In oSel
                If Not (TypeOf oEnt Is AcadBlockReference) Then Continue For
                resultado.Add(oApp.ActiveDocument.ObjectIdToObject(oEnt.ObjectID))
            Next
            oSel.Clear()
            oSel.Delete()
            oSel = Nothing
        End If
        ''
        Return resultado
    End Function

    ''
    '' Seleccionamos una polilinea cerrada o no y la usamos
    '' para hacer una seleccion PV (Poligono Ventana) y obtener todos
    '' los objetos AutoCAD que hay dentro.
    Public Function SeleccionaDameEntitiesDentroPolilinea(Optional conMensaje As Boolean = False) As ArrayList
        Dim resultado As New ArrayList
        ''
repetir:
        Dim objCadEnt As AcadEntity = Nothing
        Dim vrRetPnt As Object = Nothing
        Try
            oApp.ActiveDocument.Utility.GetEntity(objCadEnt, vrRetPnt, "Seleccione Polilinea")
            '' Si no seleccionamos nada, salimos
            If objCadEnt Is Nothing Then
                '' Volver a solicitar entidad
                GoTo repetir
                'Return resultado
                'Exit Function
            ElseIf Not (TypeOf objCadEnt Is AcadLWPolyline) Then
                '' Volver a solicitar entidad
                GoTo repetir
            End If
        Catch ex As System.Exception
            Return resultado
            Exit Function
        End Try
        '' Si el objeto seleccionado es una polilinea
        If objCadEnt.ObjectName = "AcDbPolyline" Then   '|-- Checking for 2D Polylines --|
            Dim objLWPline As AcadLWPolyline
            Dim objSSet As AcadSelectionSet
            Dim dblCurCords() As Double
            Dim dblNewCords() As Double
            Dim iMaxCurArr, iMaxNewArr As Integer
            Dim iCurArrIdx, iNewArrIdx, iCnt As Integer
            objLWPline = objCadEnt
            dblCurCords = objLWPline.Coordinates    '|-- The returned coordinates are 2D only --|
            iMaxCurArr = UBound(dblCurCords)
            If iMaxCurArr = 3 Then
                oApp.ActiveDocument.Utility.Prompt("La polilinea debe tener un minimo de 2 segmentos...")
                Return resultado
                Exit Function
            Else
                '|-- The 2D Coordinates are insufficient to use in SelectByPolygon method --|
                '|-- So convert those into 3D coordinates --|
                iMaxNewArr = ((iMaxCurArr + 1) * 1.5) - 1   '|-- New array dimension
                ReDim dblNewCords(iMaxNewArr)
                iCurArrIdx = 0 : iCnt = 1
                For iNewArrIdx = 0 To iMaxNewArr
                    If iCnt = 3 Then    '|-- The z coordinate is set to 0 --|
                        dblNewCords(iNewArrIdx) = 0
                        iCnt = 1
                    Else
                        dblNewCords(iNewArrIdx) = dblCurCords(iCurArrIdx)
                        iCurArrIdx = iCurArrIdx + 1
                        iCnt = iCnt + 1
                    End If
                Next
                ''
                '' Creamos el selectionsets para poner ahí la nueva selección
                Try
                    objSSet = oApp.ActiveDocument.SelectionSets.Item(nSel)
                Catch ex As System.Exception
                    objSSet = oApp.ActiveDocument.SelectionSets.Add(nSel)
                End Try
                '' Quitamos los objetos que hubiera seleccionados.
                objSSet.Clear()
                ''
                '' Para filtrar entidades
                ''Dim gpCode(0) As Integer
                'Dim dataValue(0) As Variant
                'gpCode(0) = 0
                'dataValue(0) = "Circle"
                'Dim groupCode As Variant, dataCode As Variant
                'groupCode = gpCode
                'dataCode = dataValue
                oApp.ActiveDocument.Activate()
                objSSet.SelectByPolygon(AcSelect.acSelectionSetWindowPolygon, dblNewCords)
                objSSet.Highlight(True)
                'objSSet.Delete
                '' Mostrar o no mensaje.
                Dim mensaje As String
                mensaje = "Nº de Objetos = " & objSSet.Count & vbCrLf & vbCrLf
                Dim nBlo, nPol, nSom, nLin As Integer
                For x As Integer = 0 To objSSet.Count - 1
                    If TypeOf objSSet.Item(x) Is AcadPolyline Then
                        nPol += 1
                    ElseIf TypeOf objSSet.Item(x) Is AcadLWPolyline Then
                        nPol += 1
                    ElseIf TypeOf objSSet.Item(x) Is AcadBlockReference Then
                        nBlo += 1
                    ElseIf TypeOf objSSet.Item(x) Is AcadHatch Then
                        nSom += 1
                    ElseIf TypeOf objSSet.Item(x) Is AcadLine Then
                        nLin += 1
                    End If
                    resultado.Add(objSSet.Item(x).ObjectID)
                Next
                ''
                mensaje = mensaje &
        "Bloques = " & nBlo & vbCrLf &
        "Polilineas = " & nPol & vbCrLf &
        "Sombreados = " & nSom & vbCrLf &
        "Lineas = " & nLin
                'MsgBox ("Objetos seleccionados = " & objSSet.Count)
                If conMensaje Then
                    MsgBox(mensaje)
                End If
                objSSet.Highlight(False)
                objSSet.Delete()
                objSSet = Nothing
            End If
            ''
        End If
        ''
        Return resultado
    End Function

    ''
    '' Obtener todos los objetos AutoCAD que hay dentro.
    Public Function DameEntitiesDentroPolilinea(objLWPline As AcadLWPolyline, Optional conMensaje As Boolean = False, Optional seleccionar As Boolean = True) As List(Of Long)
        Dim resultado As New List(Of Long)
        ''

        '' Si el objeto seleccionado es una polilinea
        Dim objSSet As AcadSelectionSet
        Dim dblCurCords() As Double
        Dim dblNewCords() As Double
        Dim iMaxCurArr, iMaxNewArr As Integer
        Dim iCurArrIdx, iNewArrIdx, iCnt As Integer
        dblCurCords = objLWPline.Coordinates    '|-- The returned coordinates are 2D only --|
        iMaxCurArr = UBound(dblCurCords)
        If iMaxCurArr = 3 Then
            oApp.ActiveDocument.Utility.Prompt("La polilinea debe tener un minimo de 2 segmentos...")
            Return resultado
            Exit Function
        Else
            '|-- The 2D Coordinates are insufficient to use in SelectByPolygon method --|
            '|-- So convert those into 3D coordinates --|
            iMaxNewArr = ((iMaxCurArr + 1) * 1.5) - 1   '|-- New array dimension
            ReDim dblNewCords(iMaxNewArr)
            iCurArrIdx = 0 : iCnt = 1
            For iNewArrIdx = 0 To iMaxNewArr
                If iCnt = 3 Then    '|-- The z coordinate is set to 0 --|
                    dblNewCords(iNewArrIdx) = 0
                    iCnt = 1
                Else
                    dblNewCords(iNewArrIdx) = dblCurCords(iCurArrIdx)
                    iCurArrIdx = iCurArrIdx + 1
                    iCnt = iCnt + 1
                End If
            Next
            ''
            '' Creamos el selectionsets para poner ahí la nueva selección
            Try
                objSSet = oApp.ActiveDocument.SelectionSets.Item(nSel)
            Catch ex As System.Exception
                objSSet = oApp.ActiveDocument.SelectionSets.Add(nSel)
            End Try
            '' Quitamos los objetos que hubiera seleccionados.
            objSSet.Clear()
            ''
            '' Para filtrar entidades
            ''Dim gpCode(0) As Integer
            'Dim dataValue(0) As Variant
            'gpCode(0) = 0
            'dataValue(0) = "Circle"
            'Dim groupCode As Variant, dataCode As Variant
            'groupCode = gpCode
            'dataCode = dataValue
            oApp.ActiveDocument.Activate()
            objSSet.SelectByPolygon(AcSelect.acSelectionSetWindowPolygon, dblNewCords)
            If seleccionar Then objSSet.Highlight(True)
            'objSSet.Delete
            '' Mostrar o no mensaje.
            Dim mensaje As String
            mensaje = "Nº de Objetos = " & objSSet.Count & vbCrLf & vbCrLf
            Dim nBlo, nPol, nSom, nLin As Integer
            For x As Integer = 0 To objSSet.Count - 1
                If TypeOf objSSet.Item(x) Is AcadPolyline Then
                    nPol += 1
                ElseIf TypeOf objSSet.Item(x) Is AcadLWPolyline Then
                    nPol += 1
                ElseIf TypeOf objSSet.Item(x) Is AcadBlockReference Then
                    nBlo += 1
                ElseIf TypeOf objSSet.Item(x) Is AcadHatch Then
                    nSom += 1
                ElseIf TypeOf objSSet.Item(x) Is AcadLine Then
                    nLin += 1
                End If
                resultado.Add(objSSet.Item(x).ObjectID)
            Next
            ''
            mensaje = mensaje &
        "Bloques = " & nBlo & vbCrLf &
        "Polilineas = " & nPol & vbCrLf &
        "Sombreados = " & nSom & vbCrLf &
        "Lineas = " & nLin
            'MsgBox ("Objetos seleccionados = " & objSSet.Count)
            If conMensaje Then
                MsgBox(mensaje)
            End If
            objSSet.Highlight(False)
            objSSet.Delete()
            objSSet = Nothing
        End If
        ''
        ''
        Return resultado
    End Function

    Public Sub VaciaMemoria()
        Try
            GC.Collect()
            GC.WaitForPendingFinalizers()
            GC.Collect()
        Catch ex As System.Exception
            '
        End Try
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