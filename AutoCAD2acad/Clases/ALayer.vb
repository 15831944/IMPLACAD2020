Imports System.Diagnostics
Imports System.Collections
Imports System.Windows.Forms
Imports System.Drawing


Imports Autodesk.AutoCAD.Interop
Imports Autodesk.AutoCAD.Interop.Common
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.EditorInput

Namespace A2acad
    Partial Public Class A2acad
        Public Function CapaDame(nombreCapa As String) As AcadLayer
            Dim resultado As AcadLayer = Nothing
            For Each oL As AcadLayer In oAppA.ActiveDocument.Layers
                If oL.Name.ToUpper.Equals(nombreCapa.ToUpper) Then
                    resultado = oL
                    Exit For
                End If
            Next
            Return resultado
        End Function
        Public Sub CapaCeroActiva()
            If oAppA Is Nothing Then _
        oAppA = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
            Try
                '' Activar Capa 0.
                oAppA.ActiveDocument.ActiveLayer = oAppA.ActiveDocument.Layers.Item("0")
            Catch ex As Exception
                '
            End Try
            ''
        End Sub
        Public Sub CapaActiva(nombreCapa As String)
            For Each oL As AcadLayer In oAppA.ActiveDocument.Layers
                If oL.Name.ToUpper.Equals(nombreCapa.ToUpper) Then
                    oAppA.ActiveDocument.ActiveLayer = oL
                    Exit For
                End If
            Next
        End Sub
        Public Sub CapaCrea(nombreCapa As String)
            Try
                oAppA.ActiveDocument.Layers.Add(nombreCapa)
            Catch ex As Exception
                'Console.Write(ex.Message)
            End Try
        End Sub
        Public Sub CapaCrea(nombresCapas As Collection)
            For Each nombrecapa As String In nombresCapas
                Try
                    oAppA.ActiveDocument.Layers.Add(nombrecapa)
                Catch ex As Exception
                    Console.Write(ex.Message)
                End Try
            Next
        End Sub
        Public Sub CapaCreaActiva(nombreCapa As String,
                              Optional queColor As Integer = 0,
                              Optional visible As Boolean = True,
                              Optional reactivada As Boolean = True,
                              Optional ponercomoactiva As Boolean = False,
                              Optional grosor As ACAD_LWEIGHT = ACAD_LWEIGHT.acLnWt025)
            ''
            '' Crear una capa y poner sus características.
            ''
            If oAppA Is Nothing Then _
        oAppA = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
            '' Activar Capa 0 primero.
            'CapaCeroActiva()
            ''
            '' Coger la capa nombreCapa o crearla
            Dim existia As Boolean = False
            Dim oLayer As AcadLayer = Nothing
            Dim oColor As AcadAcCmColor = Nothing
            Dim doc As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Using acLckDoc As Autodesk.AutoCAD.ApplicationServices.DocumentLock = doc.LockDocument()
                Try
                    oLayer = oAppA.ActiveDocument.Layers.Item(nombreCapa)
                    existia = True
                    ''
                    '' Ya existía, salimos sin hacer nada mas.
                    'Exit Sub
                Catch ex As System.Exception
                    oLayer = oAppA.ActiveDocument.Layers.Add(nombreCapa)
                    existia = False
                End Try
                ''
                '' Poner la capa como visible (LayerOn=True) y Reutilizada (Freeze=False)
                If visible = True And oLayer.LayerOn = False Then
                    oLayer.LayerOn = True
                ElseIf visible = False And oLayer.LayerOn = True Then
                    oLayer.LayerOn = False
                End If
                ''
                If reactivada = True And oLayer.Freeze = True Then
                    oLayer.Freeze = False
                ElseIf reactivada = False And oLayer.Freeze = False Then
                    oLayer.Freeze = True
                End If
                '
                ' Si no existía, poner color y grosor.
                If existia = False Then
                    '' Color
                    oColor = New AcadAcCmColor
                    Try
                        If queColor >= 0 And queColor <= 256 Then
                            oColor.ColorIndex = queColor
                            If oLayer.TrueColor IsNot oColor Then oLayer.TrueColor = oColor
                        Else
                            oColor.ColorIndex = 7
                            If oLayer.TrueColor IsNot oColor Then oLayer.TrueColor = oColor
                        End If
                    Catch ex As Exception

                    End Try
                    ' Grosor de 0.25
                    If oLayer.Lineweight <> grosor Then oLayer.Lineweight = grosor
                End If
                '
                Try
                    If ponercomoactiva = True And oAppA.ActiveDocument.ActiveLayer.Name <> oLayer.Name Then
                        oAppA.ActiveDocument.ActiveLayer = oLayer
                    End If
                Catch ex As Exception

                End Try
            End Using
            oLayer = Nothing
            oColor = Nothing
            ''oAppA.ActiveDocument.SendCommand("_line ")
            oAppA.ActiveDocument.Regen(AcRegenType.acAllViewports)
        End Sub
        Public Function CapaDameDatos(nombreCapa As String) As String
            Dim resultado As String = ""
            Dim oLay As AcadLayer = CapaDame(nombreCapa)
            If oLay Is Nothing Then
                Return resultado
                Exit Function
            End If
            '
            resultado &= "Nombre : " & oLay.Name & vbCrLf
            resultado &= "Descripción : " & oLay.Description & vbCrLf
            resultado &= "Inutilizada : " & oLay.Freeze.ToString & vbCrLf   ' Ni se ve ni se regenera
            resultado &= "Activa : " & oLay.LayerOn.ToString & vbCrLf       ' No se ve pero si se regenera
            resultado &= "Bloqueada : " & oLay.Lock.ToString & vbCrLf       ' No se puede seleccionar nada
            resultado &= "Handle : " & oLay.Handle & vbCrLf                 ' Identificador unico de la capa
            resultado &= "HasExtensionDictionary : " & oLay.HasExtensionDictionary & vbCrLf
            resultado &= "Tipo de linea : " & oLay.Linetype & vbCrLf
            resultado &= "Grosor de linea : " & oLay.Lineweight.ToString & vbCrLf
            resultado &= "Material : " & oLay.Material & vbCrLf
            resultado &= "ObjectId : " & oLay.ObjectID.ToString & vbCrLf    ' Identifador unico (Long)
            resultado &= "ObjectName : " & oLay.ObjectName & vbCrLf         ' AutoCAD Class name
            resultado &= "OwnerID : " & oLay.HasExtensionDictionary & vbCrLf    ' Padre. Identificador único
            resultado &= "PlotStyleName : " & oLay.PlotStyleName & vbCrLf   ' Estilo de impresión actual
            resultado &= "Plottable : " & oLay.Plottable.ToString & vbCrLf   ' Si se imprime o no
            resultado &= "TrueColor : (Rojo: " & oLay.TrueColor.Red & ", Verde: " & oLay.TrueColor.Green & ", Azul: " & oLay.TrueColor.Blue & ")" & vbCrLf   ' Color RGB de la capa
            resultado &= "EnUso : " & oLay.Used.ToString & vbCrLf           ' Si está en uso (Si tiene algo dentro)
            resultado &= "ViewportDefault : " & oLay.ViewportDefault.ToString           ' Si está desactivada en nuevas ventanas
            '
            Return resultado
        End Function
        Public Function CapaDameNombresVacias() As ArrayList
            Dim resultado As New ArrayList
            ''
            'Add bit about counting text on layer
            oAppA.ActiveDocument.SetVariable("pickadd", 0)   '' Se quitan las selecciones anteriores
            For Each oLa As AcadLayer In oAppA.ActiveDocument.Layers
                Dim queCapa As String = oLa.Name
                Dim myTVs(0) As TypedValue
                myTVs.SetValue(New TypedValue(DxfCode.LayerName, queCapa), 0)
                Dim myFilter As New SelectionFilter(myTVs)
                Dim myEd As Editor = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor
                Dim myPSR As PromptSelectionResult = myEd.SelectAll(myFilter)
                Dim oSel As SelectionSet = Nothing
                If myPSR.Value.Count = 0 Then
                    resultado.Add(queCapa)
                End If
            Next
            ''
            oAppA.ActiveDocument.SetVariable("pickadd", 2)   '' La seleccion actual se suma a la que hubiera.
            Return resultado
        End Function

        Public Sub CapaCambiaBloque(ByVal acb As AcadBlockReference, ByVal oDoc As AcadDocument, ByVal queCapa As String)
            If queCapa = "" Then
                queCapa = "0"
            End If
            '
            Dim ly As AcadLayer = Nothing
            Try
                ly = oDoc.Layers.Item(queCapa)
            Catch ex As Exception
                ly = oDoc.Layers.Add(queCapa)
            End Try
            '
            acb.Layer = queCapa
            acb.Update()
        End Sub
        '
        Public Sub CapaCambiaEntitiy(ByVal oEnt As AcadEntity, ByVal oDoc As AcadDocument, ByVal queCapa As String)
            If queCapa = "" Then
                queCapa = "0"
            End If
            '
            Dim ly As AcadLayer = Nothing
            Try
                ly = oDoc.Layers.Item(queCapa)
            Catch ex As Exception
                ly = oDoc.Layers.Add(queCapa)
            End Try
            '
            Try
                oEnt.Layer = queCapa
                oEnt.Update()
            Catch ex As Exception
                Console.Write(ex.Message.ToString)
            End Try
        End Sub
        Public Function CapaBorraConPrefijo(arrPreCapas As ArrayList) As Integer
            Dim resultado As Integer = 0
            '' Podemos indicar un ArrayList con los prefijos de capas
            '' a tener en cuenta para borrar.
            '' Pondremos la capa 0 como activa.
            CapaCeroActiva()
            'CapaActiva("0")
            For Each oLay As AcadLayer In oAppA.ActiveDocument.Layers
                If oLay.Name = "0" Then Continue For
                Try
                    For Each preC As String In arrPreCapas
                        If oLay.Name.ToUpper.StartsWith(preC.ToUpper) Then  ' AndAlso oLay.Used = False Then
                            Dim arrEnt As ArrayList = SeleccionaDameEntitiesEnCapa(oLay.Name)
                            If arrEnt Is Nothing OrElse arrEnt.Count = 0 Then
                                oLay.Delete()
                                resultado += 1
                                Exit For
                            End If
                        End If
                    Next
                Catch ex As System.Exception
                    '' No hacemos nada, la capa no estaba vacia.
                End Try
            Next
            ''
            Return resultado
        End Function
        Public Function CapaBorraVaciasTodo() As Integer
            Dim resultado As Integer = 0
            '' Tiene en cuenta todas las capas que estén vacías.
            '' Pondremos la capa 0 como activa.
            CapaCeroActiva()
            'CapaActiva("0")
            'Dim arrCapasBorrar As New ArrayList     '' Arraylist de las capas a borrar
            Dim esCapaAnemed As Boolean = False
            For Each oLay As AcadLayer In oAppA.ActiveDocument.Layers
                Try
                    Dim arrEnt As ArrayList = SeleccionaDameEntitiesEnCapa(oLay.Name)
                    If arrEnt Is Nothing OrElse arrEnt.Count = 0 Then
                        'arrCapasBorrar.Add(oLay.Name)
                        'oAppA.ActiveDocument.Layers.Item(oLay.Name).Delete()
                        oLay.Delete()
                        'Exit For
                        resultado += 1
                    End If

                Catch ex As System.Exception
                    '' No hacemos nada, la capa no estaba vacia.
                End Try
            Next
            ''
            Return resultado
        End Function

        Public Sub CapasFreezePViewport(layers As List(Of String))
            '' Agregue nombres de capa para congelar/descongelar
            '' en las ventanas de espacio papel separados por comas
            ''Dim layers As List(Of String) = {"Wall", "ANNO-TEXT"}.ToList() ''<-- layers for test
            Dim doc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim ed As Editor = doc.Editor
            Dim db As Database = doc.Database
            Dim idLay As ObjectId
            Dim idLayTblRcd As ObjectId
            Dim lt As LayerTableRecord
            Dim layOut As Layout
            Dim tm As Autodesk.AutoCAD.ApplicationServices.TransactionManager = db.TransactionManager
            Dim ta As Transaction = tm.StartTransaction()
            Try
                Dim acLayoutMgr As LayoutManager
                acLayoutMgr = LayoutManager.Current
                Dim layDict As DBDictionary = DirectCast(ta.GetObject(db.LayoutDictionaryId, OpenMode.ForRead, False), DBDictionary)
                For Each itmdict As DBDictionaryEntry In layDict
                    layOut = DirectCast(ta.GetObject(itmdict.Value, OpenMode.ForRead, False), Layout)
                    ed.WriteMessage(vbLf + "Layout: {0}" + vbLf, layOut.LayoutName)
                    If layOut.LayoutName <> "Model" Then
                        acLayoutMgr.CurrentLayout = layOut.LayoutName
                        Dim ltt As LayerTable = DirectCast(ta.GetObject(db.LayerTableId, OpenMode.ForRead, False), LayerTable)
                        For Each lname As String In layers
                            idLay = ltt(lname)
                            lt = ta.GetObject(idLay, OpenMode.ForRead)
                            If ltt.Has(lname) Then
                                idLayTblRcd = ltt.Item(lname)
                            Else
                                ed.WriteMessage("Layer: """ + lname + """ not available")
                                Return
                            End If
                            Dim idCol As ObjectIdCollection = New ObjectIdCollection
                            idCol.Add(idLayTblRcd)
                            ' Check that we are in paper space 
                            Dim vpid As ObjectId = ed.CurrentViewportObjectId
                            If vpid.IsNull() Then
                                ed.WriteMessage("No Viewport current.")
                                Return
                            End If
                            'VP need to be open for write 
                            Dim oViewport As Viewport = DirectCast(tm.GetObject(vpid, OpenMode.ForWrite, False), Viewport)
                            If Not oViewport.IsLayerFrozenInViewport(idLayTblRcd) Then
                                oViewport.FreezeLayersInViewport(idCol.GetEnumerator())
                            Else
                                oViewport.ThawLayersInViewport(idCol.GetEnumerator())
                            End If
                        Next
                    End If
                Next
                ta.Commit()
            Finally
                ta.Dispose()
            End Try
        End Sub
        Public Sub CapaFreezePViewport(VPlayer As String)

            '****************************************
            '*** Code from VisibleVisual.com ********
            '****************************************
            ' freeze the layer in the CURRENT viewport
            Dim objEntity As AcadObject = Nothing
            Dim objPViewport As AcadObject = Nothing
            Dim objPViewport2 As AcadObject = Nothing
            Dim XdataType As Object = Nothing
            Dim XdataValue As Object = Nothing
            Dim I As Integer
            Dim Counter As Integer
            Dim PT1 As Object = Nothing
            ' Get the active ViewPort
            objPViewport = oAppA.ActiveDocument.ActivePViewport
            ' Get the Xdata from the Viewport
            objPViewport.GetXData("ACAD", XdataType, XdataValue)
            For I = LBound(XdataType) To UBound(XdataType)
                ' Look for frozen Layers in this viewport
                If CInt(XdataType(I)) = 1003 Then
                    ' Set the counter AFTER the position of the Layer frozen layer(s)
                    Counter = I + 1
                    ' If the layer is already in the frozen layers xdata of this viewport the
                    ' exit this sub program
                    If XdataValue(I).ToString = VPlayer Then Exit Sub
                End If
            Next
            ' If no frozen layers exist in this viewport then
            ' find the Xdata location 1002 and place the frozen layer infront of the "}"
            ' found at Xdata location 1002
            If Counter = 0 Then
                For I = LBound(XdataType) To UBound(XdataType)
                    If CInt(XdataType(I)) = 1002 Then Counter = I - 1
                Next
            End If
            ' set the Xdata for the layer that is beeing frozen
            XdataType(Counter) = 1003
            XdataValue(Counter) = VPlayer
            ReDim Preserve XdataType(Counter + 1)
            ReDim Preserve XdataValue(Counter + 1)
            ' put the first "}" back into the xdata array
            XdataType(Counter + 1) = 1002
            XdataValue(Counter + 1) = "}"
            ' Keep the xdata Array and add one more to the array
            ReDim Preserve XdataType(Counter + 2)
            ReDim Preserve XdataValue(Counter + 2)
            ' put the second "}" back into the xdata array
            XdataType(Counter + 2) = 1002
            XdataValue(Counter + 2) = "}"
            ' Reset the Xdata on to the viewport
            objPViewport.SetXData(XdataType, XdataValue)
            'If no change is visible run VPupdate after this code.
            ViewportUpdate()
        End Sub
        ''
        Public Sub ViewportUpdate()
            ' Update the viewport...
            Dim objPViewport As AcadObject = oAppA.ActiveDocument.ActiveViewport
            oDoc.MSpace = False
            'objPViewport.Display (False)
            objPViewport.Display(True)
            oDoc.MSpace = True
            oDoc.Utility.Prompt("Viewport Updated!" & vbCr)
        End Sub
        ''
        ' Freezes the selected layers in all other existing viewport layouts
        Public Sub CapaFreezeOtherLayouts(ByVal pageNumber As Integer, ByVal layersToFreezeLayerIds As ObjectIdCollection)
            Dim doc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim db As Database = doc.Database
            Dim ed As Editor = doc.Editor
            Dim vp As Viewport = Nothing
            Dim viewPortFound As Boolean
            Dim freezeVPtrans As Transaction = Nothing

            Try
                freezeVPtrans = db.TransactionManager.StartTransaction()
                Dim myBT As BlockTable = db.BlockTableId.GetObject(OpenMode.ForRead)
                For Each btrID As ObjectId In myBT
                    Dim myBTR As BlockTableRecord = btrID.GetObject(OpenMode.ForRead)
                    ' If the block table record is a layout
                    If myBTR.IsLayout Then
                        viewPortFound = False
                        If Not myBTR.Name = "*Model_Space" Then
                            Dim layOut As Layout = myBTR.LayoutId.GetObject(OpenMode.ForRead)
                            'Dim initId As ObjectId = Nothing
                            'initId = layOut.Initialize()
                            ' If the layout is not the new layout
                            If layOut.TabOrder <> pageNumber Then
                                LayoutActiva(layOut.LayoutName)
                                For Each id As ObjectId In myBTR
                                    Dim obj As DBObject = id.GetObject(OpenMode.ForWrite)
                                    ' If the object is a viewport (there is a model viewport which is found first, we want the second one)
                                    If TypeOf obj Is Viewport And viewPortFound = True Then
                                        Dim vpref As Viewport = DirectCast(obj, Viewport)
                                        ' Selected Viewport for write.
                                        vp = freezeVPtrans.GetObject(vpref.ObjectId, OpenMode.ForWrite)
                                        Dim layerTable As LayerTable = freezeVPtrans.GetObject(db.LayerTableId, OpenMode.ForRead)
                                        ' Freeze the selected layers in the viewports
                                        vp.FreezeLayersInViewport(layersToFreezeLayerIds.GetEnumerator())
                                        'vp.UpdateDisplay()
                                        ed.Regen()
                                        layerTable.Dispose()
                                    End If
                                    If TypeOf obj Is Viewport Then
                                        viewPortFound = True
                                    End If
                                Next
                            End If
                        End If
                    End If
                Next
            Catch
                ed.WriteMessage("Error!")
            Finally
                freezeVPtrans.Commit()
                ed.Regen()
                freezeVPtrans.Dispose()
            End Try
        End Sub

        Public Sub LayoutActiva(ByVal LayoutName As String)
            Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim acCurDb As Database = acDoc.Database
            Using doclock As DocumentLock = acDoc.LockDocument
                Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
                    Try
                        Dim acLayoutMgr As LayoutManager
                        acLayoutMgr = LayoutManager.Current
                        acLayoutMgr.CurrentLayout = LayoutName
                        acTrans.Commit()
                    Catch ex As Autodesk.AutoCAD.Runtime.Exception
                        Autodesk.AutoCAD.ApplicationServices.Application.ShowAlertDialog(ex.Message & vbCr & ex.StackTrace)
                        acTrans.Abort()
                    End Try
                End Using
            End Using
        End Sub
        Private Function LayoutDameLayerState(Optional queLayout As String = "") As String
            Dim resultado As String = ""
            Dim Layout As AcadLayout = Nothing
            Try
                If queLayout = "" Then
                    Layout = oAppA.ActiveDocument.ActiveLayout
                Else
                    Layout = oAppA.ActiveDocument.Layouts.Item(queLayout)
                End If
            Catch ex As Exception
                Return resultado
                Exit Function
            End Try
            '
            If Not (Layout.HasExtensionDictionary) Then LayoutPonLayerState()

            Dim XRec As AcadXRecord = Layout.GetExtensionDictionary("TabHasLState")

            Dim dxfCodes As Object = Nothing
            Dim dxfData As Object = Nothing
            XRec.GetXRecordData(dxfCodes, dxfData)

            Dim i As Integer
            For i = LBound(dxfCodes) To UBound(dxfCodes)
                If CInt(dxfCodes(i)) = 1 Then
                    resultado = dxfData(i)
                    Exit For
                End If
            Next i
            XRec = Nothing
            Layout = Nothing
            '
            Return resultado
        End Function
        Public Sub LayoutPonLayerState()
            Dim Name As String = estadocapas
            'Name = InputBox("Specify the LayerState to assign to the current layout.", "Assign LayerState")
            Dim XRec As AcadXRecord = oAppA.ActiveDocument.ActiveLayout.GetExtensionDictionary.AddXRecord("TabHasLState")
            Dim dxfCodes(0) As Integer
            Dim dxfData(0) As Object
            dxfCodes(0) = 1 : dxfData(0) = Name
            Try
                XRec.SetXRecordData(dxfCodes, dxfData)
            Catch ex As Exception
                ' Ya existe el estado de capas del Layout
            End Try
            XRec = Nothing
        End Sub
        Public Sub LayerStateSave(Optional Name As String = "")
            If Name = "" Then Name = estadocapas
            If oLsm Is Nothing Then Exit Sub
            oLsm.SetDatabase(oAppA.ActiveDocument.Database)
            Try
                If LayerStateExists(oLsm, Name) = False Then
                    ' oLsm.Restore(Name)
                    'Else
                    oLsm.Save(Name, AcLayerStateMask.acLsAll)
                End If
            Catch ex As Exception
            End Try
        End Sub
        Public Sub LayerStateRestore(Name As String)
            If oLsm Is Nothing Then Exit Sub
            oLsm.SetDatabase(oAppA.ActiveDocument.Database)
            If LayerStateExists(oLsm, Name) Then
                LayerStateRestore(Name)
                oAppA.ActiveDocument.Regen(AcRegenType.acActiveViewport)
            End If
        End Sub
        Public Function LayerStateExists(LSMan As AcadLayerStateManager, Name As String) As Boolean
            On Error Resume Next
            LSMan.Restore(Name)
            LayerStateExists = (Err.Number = 0)
            'If Not (LStateExists) Then
            '    MsgBox("The LayerState """ & Name & """ is assigned to the current layout, but cannot be found." &
            '        vbCrLf & "Please use the SetTabLState macro to reassign the LayerState, or recreate the Layer State.", vbInformation, "Missing assigned LayerState")
            'End If
        End Function
    End Class
End Namespace