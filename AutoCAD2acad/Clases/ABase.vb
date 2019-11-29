Imports System.Diagnostics
Imports System.Collections
Imports System.Windows.Forms
Imports System.Drawing


Imports Autodesk.AutoCAD.Interop
Imports Autodesk.AutoCAD.Interop.Common
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.ApplicationServices
Imports oAppS = Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.EditorInput

Namespace A2acad
    Partial Public Class A2acad
        ''' <summary>
        ''' Crear nueva instancia de clsAutoCAD2acad (oApp as AcadApplicationk, fullPath DLL que lo llama, appReg [CLIENTE]2acad)
        ''' </summary>
        ''' <param name="queApp">AcadApplication que cargamos al inicio del desarrollo</param>
        ''' <param name="_appFullPathPadreDll">fullPath de la DLL que instancia esta clase (app_fullpath)</param>
        ''' <param name="queAppReg">regApp que creamos en variables [CLIENTE]2acad</param>
        Public Sub New(queApp As Autodesk.AutoCAD.Interop.AcadApplication, _appFullPathPadreDll As String, queAppReg As String, key As String)
            If key <> keyPre & "_" & keySuf Then
                activado = False
                Throw New System.Exception("Not allowed")
                Return
            Else
                activado = True
            End If
            Control.CheckForIllegalCrossThreadCalls = False
            oAppA = queApp
            Me.regAPPA = queAppReg
            If IO.File.Exists(_appFullPathPadreDll) Then DatosPadrePon(_appFullPathPadreDll)
            ConectaAcad()
            ConectaDibujo()
            ventanas = New System.Collections.Specialized.StringDictionary
        End Sub
        Public Sub DatosPadrePon(_appPathPadre)
            _appPath = _appPathPadre
            _appNombre = IO.Path.GetFileNameWithoutExtension(_appPath)
            _appDir = IO.Path.GetDirectoryName(_appPath)
            _appLog = IO.Path.Combine(_appDir, _appNombre & ".log")
        End Sub

        Public Sub ConectaAcad()
            'oAppAT = oAppA.Caption
            Me.oAppAintP = clsAPI.DameIntPtr(oAppA.Caption)
            Dim PRs() As System.Diagnostics.Process = System.Diagnostics.Process.GetProcessesByName("acad", My.Computer.Name)
            prApp = PRs(0)       ' El IntPtr del proceso. Para controlar SendMessage
            '
            If (OC Is Nothing) Then
                Try
                    If oAppA.Version.StartsWith("21") Then  '.Contains("2017") Or oAppA.Version.Contains("R21") Then
                        OC = CType(oAppA.GetInterfaceObject("AutoCAD.AcCmColor.21"), AcadAcCmColor)
                    ElseIf oAppA.Version.StartsWith("22") Then  '.Contains("2020") Or oAppA.Version.Contains("R22") Then
                        OC = CType(oAppA.GetInterfaceObject("AutoCAD.AcCmColor.22"), AcadAcCmColor)
                    ElseIf oAppA.Version.StartsWith("23") Then  '.Contains("2019") Or oAppA.Version.Contains("R23") Then
                        OC = CType(oAppA.GetInterfaceObject("AutoCAD.AcCmColor.23"), AcadAcCmColor)
                    End If
                Catch ex As Exception
                    Try
                        OC = CType(oAppA.GetInterfaceObject("AutoCAD.AcCmColor"), AcadAcCmColor)
                    Catch ex1 As Exception
                        ''
                    End Try
                End Try
            End If

            If (oLsm Is Nothing) Then
                Try
                    If oAppA.Version.StartsWith("21") Then 'oAppA.Version.Contains("2017") Or oAppA.Version.Contains("R21") Then
                        oLsm = oAppA.GetInterfaceObject("AutoCAD.AcadLayerStateManager.21")
                    ElseIf oAppA.Version.StartsWith("22") Then  'oAppA.Version.Contains("2020") Or oAppA.Version.Contains("R22") Then
                        oLsm = oAppA.GetInterfaceObject("AutoCAD.AcadLayerStateManager.22")
                    ElseIf oAppA.Version.StartsWith("23") Then  'oAppA.Version.Contains("2019") Or oAppA.Version.Contains("R23") Then
                        oLsm = oAppA.GetInterfaceObject("AutoCAD.AcadLayerStateManager.23")
                    End If
                Catch ex As Exception
                    Try
                        oLsm = oAppA.GetInterfaceObject("AutoCAD.AcadLayerStateManager")
                    Catch ex1 As Exception
                        ''
                    End Try
                End Try
            End If
            '
            'Me.LiberaApp()          'Liberamos AutoCAD si está ocupado en algo.
        End Sub

        Public Sub ConectaDibujo()
            Try
                'Exit Sub
                If Not (oAppA Is Nothing) Then
                    If oAppA.Documents.Count = 0 Then Exit Sub
                    Try
                        Try
                            oDoc.RegisteredApplications.Add(regAPPA)
                        Catch ex As Exception
                            ' La aplicacion ya esta registrada.
                        End Try
                        '
                        Me.oDocFull = oDoc.FullName
                        'Dim oDocNCorto As String = IO.Path.GetFileName(oDoc.FullName)
                        oBd = oDoc.Database
                        Me.oDochw = oDoc.HWND
                        Me.oDocintP = clsAPI.FindWindow(Nothing, oAppA.Caption)
                    Catch ex As Exception
                        MessageBox.Show("Tiene que haber un dibujo Abierto y Guardado..." & vbCrLf & vbCrLf & vbCrLf & ex.Message)
                        'frmInicio.Form1_Load(Nothing, Nothing)
                        Exit Sub
                    End Try
                Else
                    MessageBox.Show("AutoCAD no está abierto...")
                End If
                'Me.ActivaApp(, True)
                'Me.LiberaRecursos() ' bloque SERICAD
                'Me.InsertaRecursos(, modVariables.dirApp & regAPP & ".dwg")
            Catch ex As Exception

            End Try
        End Sub

        Public Sub AutoCAD_PonFoCo()
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
            Autodesk.AutoCAD.ApplicationServices.Application.MainWindow.Focus()
        End Sub
        Public Function DameArrPol(ByVal quePol As AcadLWPolyline) As Double()
            Dim pt() As Double = quePol.Coordinates
            Dim puntos(-1) As Double
            For x = 0 To pt.GetLength(0) - 1 Step 2
                Dim pt1(2) As Double
                pt1(0) = pt(x) : pt1(1) = pt(x + 1) : pt1(2) = 0

                ReDim Preserve puntos(puntos.GetUpperBound(0) + 1)
                puntos(puntos.GetUpperBound(0)) = pt1(0)

                ReDim Preserve puntos(puntos.GetUpperBound(0) + 1)
                puntos(puntos.GetUpperBound(0)) = pt1(1)

                ReDim Preserve puntos(puntos.GetUpperBound(0) + 1)
                puntos(puntos.GetUpperBound(0)) = pt1(2)

            Next
            '
            Return puntos
        End Function
        ' Da error
        'Public Sub HazZoomObjeto(ByVal obj As AcadObject, Optional ByVal poco As Boolean = False)
        '    Using oLock As DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
        '        Using acTrans As Transaction = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Database.TransactionManager.StartOpenCloseTransaction()
        '            'poco = true hace que el zoom se mueve 5 unidades de dibujo hacia fuera de zoom
        '            'Si no ponemos este valor el zoom hará que veamos el objeto a mital de su tamaño
        '            Dim pt1() As Double = Nothing
        '            Dim pt2() As Double = Nothing
        '            obj.GetBoundingBox(pt1, pt2)
        '            Dim distX As Double = Math.Abs(pt2(0) - pt1(0))
        '            Dim distY As Double = Math.Abs(pt2(1) - pt1(1))
        '            If poco = True Then
        '                distX = 1 : distY = 1
        '            End If
        '            'If oAppA.ActiveDocument.UserCoordinateSystems.
        '            'Dim margen As Integer
        '            If oDoc.GetVariable("INSUNITSDEFTARGET") = 4 Then
        '                'margen = multi * 1000
        '            ElseIf oDoc.GetVariable("INSUNITSDEFTARGET") = 5 Then
        '                'margen = multi * 100
        '            ElseIf oDoc.GetVariable("INSUNITSDEFTARGET") = 6 Then
        '                'margen = multi * 10
        '            End If

        '            pt1(0) -= (distX) : pt1(1) -= (distY)
        '            pt2(0) += (distX) : pt2(1) += (distY)
        '            oAppA.ZoomWindow(pt1, pt2)
        '            acTrans.Commit()
        '        End Using
        '    End Using
        'End Sub

        Public Sub HazZoomObjeto(ByVal obj As AcadObject, Optional reduce As Double = 1, Optional selecciona As Boolean = True)
            Dim pt1 As Object = Nothing
            Dim pt2 As Object = Nothing
            obj.GetBoundingBox(pt1, pt2)
            Dim distX As Double = Math.Abs(pt2(0) - pt1(0))
            Dim distY As Double = Math.Abs(pt2(1) - pt1(1))
            '
            pt1(0) -= (distX) * reduce : pt1(1) -= (distY) * reduce
            pt2(0) += (distX) * reduce : pt2(1) += (distY) * reduce
            oAppA.ZoomWindow(pt1, pt2)
            If selecciona = True Then
                'Dim oIPrt As New IntPtr(obj.ObjectID)
                'Dim oId As New ObjectId(oIPrt)
                'Dim arrIds() As ObjectId = {oId}
                'Autodesk.AutoCAD.Internal.Utils.SelectObjects(arrIds)
                Selecciona_AcadObject(obj)
            End If
        End Sub
        'Public Sub HazZoomSeleccion(ByVal oSel As AcadSelectionSet, Optional reduce As Double = 1, Optional selecciona As Boolean = True)
        '    Dim pt1 As Object = Nothing
        '    Dim pt2 As Object = Nothing
        '    obj.GetBoundingBox(pt1, pt2)
        '    Dim distX As Double = Math.Abs(pt2(0) - pt1(0))
        '    Dim distY As Double = Math.Abs(pt2(1) - pt1(1))
        '    '
        '    pt1(0) -= (distX) * reduce : pt1(1) -= (distY) * reduce
        '    pt2(0) += (distX) * reduce : pt2(1) += (distY) * reduce
        '    oAppA.zo.ZoomWindow(pt1, pt2)
        '    If selecciona = True Then
        '        'Dim oIPrt As New IntPtr(obj.ObjectID)
        '        'Dim oId As New ObjectId(oIPrt)
        '        'Dim arrIds() As ObjectId = {oId}
        '        'Autodesk.AutoCAD.Internal.Utils.SelectObjects(arrIds)
        '        Selecciona_AcadObject(obj)
        '    End If
        'End Sub



        ' Seleccionar con los puntos de la polilinea (se puede agregar capa, dotosX)
        Public Sub SeleccionaConPuntos(ByVal opol As AcadLWPolyline,
                                   Optional ByVal tipoObj As String = "INSERT",
                                   Optional ByVal puntos() As Double = Nothing,
                                   Optional ByVal capa As String = "",
                                   Optional ByVal DatosX As Boolean = False,
                                   Optional ByVal SelTemp As Boolean = False,
                                   Optional ByVal iluminarSel As Boolean = False)
            Dim F1(0) As Short
            Dim F2(0) As Object
            Dim vF1 As Object = Nothing
            Dim vF2 As Object = Nothing

            F1(0) = 0 : F2(0) = tipoObj ' "INSERT"
            'DatosX Siempre tiene que estar despues de entidad. Si no no funciona
            If DatosX = True Then
                ReDim Preserve F1(F1.Length)
                ReDim Preserve F2(F2.Length)
                'Filtro de bloque, nombre y capa
                F1(F1.Length - 1) = 1001 : F2(F2.Length - 1) = regAPPA
            End If
            If capa <> "" Then
                ReDim Preserve F1(F1.Length)
                ReDim Preserve F2(F2.Length)
                'Filtro de bloque, nombre y capa
                F1(F1.Length - 1) = 8 : F2(F1.Length - 1) = capa
            End If

            vF1 = F1
            vF2 = F2
            Me.oSel.Clear()
            'oSel.Select(AcSelect.acSelectionSetWindowPolygon, puntos, puntoFicticio vF1, vF2)  'Punto ficticio es un 2ª punto que habría que poner
            If puntos Is Nothing Then
                'oSel.SelectByPolygon(AcSelect.acSelectionSetWindowPolygon, Me.DamePuntos3DPolilinea(opol), vF1, vF2)
                oSel.SelectByPolygon(AcSelect.acSelectionSetCrossingPolygon, Me.DamePuntos3DPolilinea(opol), vF1, vF2)
            Else
                'oSel.SelectByPolygon(AcSelect.acSelectionSetWindowPolygon, puntos, vF1, vF2)
                oSel.SelectByPolygon(AcSelect.acSelectionSetCrossingPolygon, puntos, vF1, vF2)
            End If

            If oSel.Count > 0 AndAlso iluminarSel = True Then
                oSel.Highlight(True)
                oSel.Update()
            End If
            'oSel.Select(AcSelect.acSelectionSetWindowPolygon, puntos, vF1, vF2)
        End Sub

        Public Function ConvTo3dPoints(ByVal objCoors() As Double) As Double()
            Dim i As Long, j As Long
            Dim convPts() As Double = Nothing

            j = 0
            For i = LBound(objCoors) To UBound(objCoors) Step 2
                ReDim Preserve convPts(0 To j)
                convPts(j) = objCoors(i)
                ReDim Preserve convPts(0 To j + 1)
                convPts(j + 1) = objCoors(i + 1)
                ReDim Preserve convPts(0 To j + 2)
                convPts(j + 2) = 0
                j = j + 3

            Next
            ConvTo3dPoints = convPts
        End Function

        Public Function ArreglaCoordenadas(ByVal puntos() As Double) As Double()
            Dim puntosNuevos(-1) As Double
            Dim contador As Integer = 0
            For Each p As Double In puntos
                ReDim Preserve puntosNuevos(puntosNuevos.GetUpperBound(0) + 1)
                puntosNuevos(puntosNuevos.GetUpperBound(0)) = p
                contador += 1
                If contador > 0 Then
                    If Math.IEEERemainder(CDbl(contador), 2.0#) = 0 Then
                        ReDim Preserve puntosNuevos(puntosNuevos.GetUpperBound(0) + 1)
                        puntosNuevos(puntosNuevos.GetUpperBound(0)) = 0
                    End If
                End If
            Next
            ArreglaCoordenadas = puntosNuevos
        End Function

        Public Function DamePuntos3DPolilinea(ByVal opol As AcadLWPolyline) As Double()
            Dim puntosNuevos(-1) As Double
            Dim contador As Integer = 0
            For Each coor As Double In opol.Coordinates
                ReDim Preserve puntosNuevos(puntosNuevos.GetUpperBound(0) + 1)
                puntosNuevos(puntosNuevos.GetUpperBound(0)) = coor
                contador += 1
                If contador > 0 Then
                    If Math.IEEERemainder(CDbl(contador), 2.0#) = 0 Then
                        ReDim Preserve puntosNuevos(puntosNuevos.GetUpperBound(0) + 1)
                        puntosNuevos(puntosNuevos.GetUpperBound(0)) = 0
                    End If
                End If
            Next
            DamePuntos3DPolilinea = puntosNuevos
        End Function
        ' Quitamos esta funcion, hasta saber por qué falla.
        Public Sub LiberaApp(Optional ByVal MoverC As Boolean = True)
            Exit Sub
            Try
                If MoverC = True Then Cursor.Position = New Point(400, Screen.PrimaryScreen.WorkingArea.Height / 2)
                AppActivate(oAppA.Caption)
                If Not (Me.ventanas Is Nothing) Then Me.ventanas.Clear()
                Me.ventanas = clsVentanas.DameVentanasHijas(Me.oAppAintP)
                ''clsVentanas.EnumerarVentanas() 'Mensaje con todas las ventanas en Windows
                ''If Me.hanDoc > 0 Then clsVentanas.enumerarVentanasHijas(Me.hwoApp) 'Mensaje con ventanas AutoCAD.
                If oAppA.GetAcadState.IsQuiescent = False Then
                    Call modTeclas.SendMessage(clsVentanas.lc, teclas.WM_IME_KEYUP, teclas.VK_ESCAPE, 0)
                    Call modTeclas.SendMessage(clsVentanas.lc, teclas.WM_IME_KEYDOWN, teclas.VK_ESCAPE, 0)
                    Call modTeclas.SendMessage(clsVentanas.lc, teclas.WM_IME_KEYUP, teclas.VK_ESCAPE, 0)
                    Call modTeclas.SendMessage(clsVentanas.lc, teclas.WM_IME_KEYDOWN, teclas.VK_ESCAPE, 0)
                End If
                'Este es el bueno (MountTam)   'HeadLand
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        End Sub

        Public Sub ActivaApp(Optional ByVal MoverC As Boolean = False, Optional wNormal As Boolean = False)
            AppActivate(oAppA.Caption)
            If MoverC = True Then Cursor.Position = New Point(400, Screen.PrimaryScreen.WorkingArea.Height / 2)
            If wNormal = True And Me.oAppA.WindowState <> Autodesk.AutoCAD.Interop.Common.AcWindowState.acNorm Then
                Me.oAppA.WindowState = Autodesk.AutoCAD.Interop.Common.AcWindowState.acNorm
            End If
        End Sub

        Public Sub ActivaAppAPI()
            clsAPI.SetForegroundWindow(Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.CurrentDocument.Window.Handle)
            clsAPI.SetFocus(Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.CurrentDocument.Window.Handle.ToInt64)
        End Sub

        'Public Sub ActivaAppViejo(Optional ByVal queApp As String = "", Optional ByVal MoverC As Boolean = True)
        '    'LiberaApp()
        '    If queApp = "" Then queApp = oAppA.Caption
        '    Dim intP As IntPtr = clsAPI.FindWindow(Nothing, queApp)
        '    If intP <> IntPtr.Zero Then
        '        clsAPI.SetForegroundWindow(intP)
        '        AppActivate(queApp)
        '        If oAppA.WindowState <> Autodesk.AutoCAD.Interop.Common.AcWindowState.acNorm Then
        '            oAppA.WindowState = Autodesk.AutoCAD.Interop.Common.AcWindowState.acNorm
        '        End If
        '        If MoverC = True Then Cursor.Position = New Point(400, Screen.PrimaryScreen.WorkingArea.Height / 2)
        '        'If (Me.oDoc Is Nothing) Then Me.oDoc = oAppA.ActiveDocument
        '        'Me.oAppA.ActiveDocument.Activate()

        '        'If oAppA.GetAcadState.IsQuiescent = True Then oAppA.ActiveDocument.SendCommand(Chr(27))
        '        'Me.oDoc.Utility.Prompt(vbCrLf & "")
        '        'Me.oAppA.ActiveDocument.ActivePViewport.GridOn = True
        '        'Me.oDoc.SendCommand(Chr(27))
        '        'Me.oAppA.ActiveDocument.ActivePViewport.GridOn = False
        '    End If
        'End Sub

        '************************ PARA BLOQUES

        ''' <summary>
        ''' Nos devuelve un ArrayList con los atributos del bloque
        ''' </summary>
        ''' <param name="bl">Objeto bloque que pasaremos</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function DameAtributos(ByVal bl As AcadBlockReference) As ArrayList
            Dim att As AcadAttributeReference
            Dim arrA As New ArrayList
            Dim atts() As Object
            atts = bl.GetAttributes
            'If att.Length > 0 Then
            For Each atrr As Object In atts
                att = CType(atrr, AcadAttributeReference)
                arrA.Add(att)
            Next

            DameAtributos = arrA
            Exit Function
        End Function

        Public Function DameAtributosC(ByVal bl As AcadBlockReference) As ArrayList
            Dim att As AcadAttribute
            Dim arrA As New ArrayList
            Dim atts() As Object
            atts = bl.GetConstantAttributes
            'If att.Length > 0 Then
            For Each atrr As Object In atts
                att = CType(atrr, AcadAttribute)
                arrA.Add(att)
            Next

            DameAtributosC = arrA
            Exit Function
        End Function


        Public Function DameZonaC(ByVal bl As AcadBlockReference) As String
            Dim att As AcadAttribute
            Dim resultado As String = ""
            Dim atts() As Object
            atts = bl.GetConstantAttributes
            'If att.Length > 0 Then
            For Each atrr As Object In atts
                att = CType(atrr, AcadAttribute)
                If att.TagString.ToString.ToUpper = "ZONCLIENTE" Then
                    resultado = att.TextString
                    Exit For
                End If
            Next
            DameZonaC = resultado
            Exit Function
        End Function

        Public Function DameAtributosColl(ByVal bl As AcadBlockReference) As Collection
            Dim att As AcadAttribute
            Dim arrA As New Collection
            Dim atts() As Object
            atts = bl.GetConstantAttributes
            'If att.Length > 0 Then
            For Each atrr As Object In atts
                att = CType(atrr, AcadAttribute)
                arrA.Add(att, att.TagString)
            Next

            DameAtributosColl = arrA
            Exit Function
        End Function

        Public Function MensajeAtributos(ByVal bl As AcadBlockReference) As String
            Dim resultado As String = ""
            Dim att As AcadAttribute
            Dim arrA As New ArrayList
            Dim atts() As Object
            atts = bl.GetConstantAttributes

            If atts.Length > 0 Then
                For Each atrr As Object In atts
                    att = CType(atrr, AcadAttribute)
                    resultado &= (att.TagString & " :" & vbTab & att.TextString) & vbCrLf
                Next
            End If
            MensajeAtributos = resultado
            Exit Function
        End Function

        Public Sub PonAtributosC(ByVal bl As AcadBlockReference)
            Dim att As AcadAttribute
            Dim arrA As New ArrayList
            Dim atts() As Object
            atts = bl.GetConstantAttributes

            'Dim nCapa As String = bl.Layer.Replace(modVariables.preCapa, "")
            If atts.Length > 0 Then
                For Each atrr As Object In atts
                    att = CType(atrr, AcadAttribute)
                    If att.TagString = "NOMBLOQUE" Then
                        att.TextString = bl.Name.ToUpper   'Saldremos del bucle porque no cambiamos más datos.
                        Exit For
                    ElseIf att.TagString = "ZONCLIENTE" Then
                        'att.TextString = Me.LeeXDataZona(bl)
                    ElseIf att.TagString = "CONTROL" Then
                        'No hacemos nada
                    Else
                        'att.TextString = att.TagString.ToLower
                    End If
                    'ct.Text += (att.TagString & " / " & att.TextString) & vbCrLf
                Next
            End If
        End Sub
    End Class
End Namespace