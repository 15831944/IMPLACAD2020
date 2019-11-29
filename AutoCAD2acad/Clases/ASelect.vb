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
        Public WithEvents oTresalte As System.Timers.Timer
        Public Sub SelectionSet_Delete(Optional queSet As String = "")
            ' Borra el SelectionSet (No los objetos que tuviera)
            If queSet = "" Then queSet = regAPPA
            Try
                oSel = oAppA.ActiveDocument.SelectionSets.Item(queSet)
                For x As Integer = 0 To oSel.Count
                    oSel.Item(x).Highlight(False)
                Next
                oSel = Nothing
                oAppA.ActiveDocument.SelectionSets.Item(queSet).Delete()
            Catch ex As System.Exception
                ' No existe
            End Try
        End Sub
        Public Sub Seleccion_Quitar()
            ' No usamos objecto. Por si los usaramos
            Dim F1(1) As Short
            Dim F2(1) As Object
            Dim vF1 As Object = Nothing
            Dim vF2 As Object = Nothing

            '' Las 2 maneras valen igual. AcDbBlckReference es mejor (Solo coge bloques) INSERT coge sombreados también.
            F1(0) = 100 : F2(0) = "*"
            F1(1) = 0 : F2(1) = "*"
            vF1 = F1
            vF2 = F2
            ''
            If oAppA Is Nothing Then _
    oAppA = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
            ''
            ''
            Try
                oSel = oAppA.ActiveDocument.SelectionSets.Add(regAPPA)
            Catch ex As System.Exception
                oSel = oAppA.ActiveDocument.SelectionSets.Item(regAPPA)
            End Try
            ''
            oSel.Clear()
            ''
            oAppA.ActiveDocument.SetVariable("pickadd", 0)   '' Solo una selección. Se quita lo que hubiera
            '
            Dim point() As Double = {10000, 10000, 100000}
            oSel.SelectAtPoint(point)
            '
            oAppA.ActiveDocument.SetVariable("pickadd", 2)   '' La seleccion actual se suma a la que hubiera.
        End Sub

        Public Sub Seleccion_CreaResalta(queEntidades As ArrayList, conZoom As Boolean)
            ''
            Try
                oSel = oAppA.ActiveDocument.SelectionSets.Add(regAPPA)
            Catch ex As System.Exception
                oSel = oAppA.ActiveDocument.SelectionSets.Item(regAPPA)
            End Try
            ''
            ''
            oAppA.ActiveDocument.SetVariable("pickadd", 0)   '' Solo una selección. Se quita lo que hubiera
            oSel.Clear()
            ''
            '' Recorremos el diccionario para añadir las entidades a oSelTemp y resaltarlas o quitar resalte.
            'For Each queEnt As AcadEntity In queEntidades
            'Dim objeto(0) As Object : objeto(0) = oApp.ActiveDocument.ObjectIdToObject(queEnt.ObjectID)
            'oSelTemp.AddItems(objeto)
            'Next
            Dim arrObj(queEntidades.Count - 1) As AcadEntity
            Dim arrIds(queEntidades.Count - 1) As ObjectId
            For x As Integer = 0 To queEntidades.Count - 1
                arrObj(x) = queEntidades(x)
                arrIds(x) = New ObjectId(CType(queEntidades(x), AcadEntity).ObjectID)
            Next
            oSel.AddItems(arrObj)
            ''
            'oSelTemp.Highlight(True)
            oAppA.ActiveDocument.Regen(AcRegenType.acActiveViewport)
            Autodesk.AutoCAD.Internal.Utils.SelectObjects(arrIds)
            If conZoom Then
                Autodesk.AutoCAD.Internal.Utils.ZoomObjects(False)
            End If
            ''
            oAppA.ActiveDocument.SetVariable("pickadd", 2)   '' La seleccion actual se suma a la que hubiera.
        End Sub
        Public Sub SeleccionCreaResalta(queEntidades As ArrayList, Optional tiempo As Integer = 5000, Optional conZoom As Boolean = True)
            ''
            Try
                oSel = oAppA.ActiveDocument.SelectionSets.Add(regAPPA)
            Catch ex As System.Exception
                oSel = oAppA.ActiveDocument.SelectionSets.Item(regAPPA)
            End Try
            ''
            ''
            oAppA.ActiveDocument.SetVariable("pickadd", 0)   '' Solo una selección. Se quita lo que hubiera
            oSel.Clear()
            ''
            '' Recorremos el diccionario para añadir las entidades a oSelTemp y resaltarlas o quitar resalte.
            'For Each queEnt As AcadEntity In queEntidades
            'Dim objeto(0) As Object : objeto(0) = oApp.ActiveDocument.ObjectIdToObject(queEnt.ObjectID)
            'oSelTemp.AddItems(objeto)
            'Next
            Dim arrObj(queEntidades.Count - 1) As AcadEntity
            Dim arrIds(queEntidades.Count - 1) As ObjectId
            For x As Integer = 0 To queEntidades.Count - 1
                arrObj(x) = queEntidades(x)
                arrIds(x) = New ObjectId(CType(queEntidades(x), AcadEntity).ObjectID)
            Next
            oSel.AddItems(arrObj)
            ''
            'oSelTemp.Highlight(True)
            oAppA.ActiveDocument.Regen(AcRegenType.acActiveViewport)
            Autodesk.AutoCAD.Internal.Utils.SelectObjects(arrIds)
            If conZoom Then
                Autodesk.AutoCAD.Internal.Utils.ZoomObjects(False)
            End If
            ''
            oAppA.ActiveDocument.SetVariable("pickadd", 2)   '' La seleccion actual se suma a la que hubiera.
            If IsNumeric(tiempo) And tiempo > 0 Then
                oTresalte = New System.Timers.Timer
                oTresalte.Interval = tiempo
                oTresalte.Start()
            End If
        End Sub
        ''
        Private Sub oTresalte_Elapsed(sender As Object, e As Timers.ElapsedEventArgs) Handles oTresalte.Elapsed
            If oSel IsNot Nothing Then ''
                Try
                    If Autodesk.AutoCAD.ApplicationServices.Application.IsQuiescent = False Then
                        oSel.Highlight(False)
                        oAppA.ActiveDocument.Regen(AcRegenType.acActiveViewport)
                        '' Quitar la selección actual
                        Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.SendStringToExecute(Chr(27), False, False, False)
                        Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.SetImpliedSelection(New ObjectId() {})
                        '
                        Try
                            oSel.Clear()
                            oSel.Delete()
                        Catch ex As Exception
                            '
                        End Try
                    End If
                Catch ex As Exception

                End Try
                oSel = Nothing
            End If
            oTresalte.Stop()
        End Sub
        Public Function SeleccionaBloquePorHandleXdata(ByVal queHandle As String) As AcadBlockReference
            Dim resultado As AcadBlockReference = Nothing
            'Dim cSeleccion As AcadSelectionSets
            Dim F1(2) As Short
            Dim F2(2) As Object
            Dim vF1 As Object = Nothing
            Dim vF2 As Object = Nothing

            '' Las 2 maneras valen igual. AcDbBlckReference es mejor (Solo coge bloques) INSERT coge sombreados también.
            F1(0) = 100 : F2(0) = "AcDbBlockReference"
            F1(1) = 0 : F2(1) = "INSERT"
            F1(2) = 5 : F2(2) = queHandle
            '
            vF1 = F1
            vF2 = F2
            ''
            If oAppA Is Nothing Then _
    oAppA = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
            ''
            ''
            Try
                oSel = oAppA.ActiveDocument.SelectionSets.Add(nSel)
            Catch ex As System.Exception
                oSel = oAppA.ActiveDocument.SelectionSets.Item(nSel)
            End Try
            ''
            oAppA.ActiveDocument.SetVariable("pickadd", 0)   '' Solo una selección. Se quita lo que hubiera
            oSel.Clear()
            Try
                oSel.Select(vF1, vF2)
            Catch ex As System.Exception
                'Debug.Print(ex.Message)
            End Try
            'MsgBox(oSel.Count)
            If oSel.Count > 0 Then
                'Dim texto As String = XData.XLeeDato(oEnt, xT.TEXTOS)
                'If TypeOf oEnt Is AcadTable Or texto = "Clase=tabla" Then
                If TypeOf oSel.Item(oSel.Count - 1) Is AcadTable Then
                    resultado = Nothing
                Else
                    resultado = oSel.Item(oSel.Count - 1)
                End If
            Else
                resultado = Nothing
            End If
            ''
            oSel.Clear()
            oSel.Delete()
            oSel = Nothing
            oAppA.ActiveDocument.SetVariable("pickadd", 2)   '' La seleccion actual se suma a la que hubiera.
            Return resultado
        End Function
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
            Catch ex As Exception
                '
            End Try
        End Sub

        Public Sub Selecciona_AcadObject(objEnt As AcadObject)
            oAppA.ActiveDocument.SetVariable("pickadd", 0)   ' Quitar la seleccion que hubiera.
            Dim queHandle As String = objEnt.Handle
            Dim obj As AcadObject = oAppA.ActiveDocument.HandleToObject(queHandle)
            ' Volver a seleccionar el objecto. Dara error si se ha borrado
            Try
                Dim oIPrt As New IntPtr(obj.ObjectID)
                Dim oId As New ObjectId(oIPrt)
                Dim arrIds() As ObjectId = {oId}
                Autodesk.AutoCAD.Internal.Utils.SelectObjects(arrIds)
            Catch ex As Exception
                '
            End Try
            oAppA.ActiveDocument.SetVariable("pickadd", 2)   ' Quitar la seleccion que hubiera.
        End Sub
        Public Sub Selecciona_AcadID(IdEnt As Long)
            oAppA.ActiveDocument.SetVariable("pickadd", 0)   ' Quitar la seleccion que hubiera.
            Try
                Dim oIPrt As New IntPtr(IdEnt)
                Dim oId As New ObjectId(oIPrt)
                Dim arrIds() As ObjectId = {oId}
                Autodesk.AutoCAD.Internal.Utils.SelectObjects(arrIds)
            Catch ex As Exception
                '
            End Try
            oAppA.ActiveDocument.SetVariable("pickadd", 2)   '' La seleccion actual se suma a la que hubiera.
        End Sub
        Public Sub Selecciona_AcadID(IdEnt As Long())
            oAppA.ActiveDocument.SetVariable("pickadd", 0)   ' Quitar la seleccion que hubiera.
            Dim colIds As New List(Of ObjectId)
            For Each LongId As Long In IdEnt
                Dim oId As New ObjectId(New IntPtr(LongId))
                colIds.Add(oId)
            Next
            Try
                Autodesk.AutoCAD.Internal.Utils.SelectObjects(colIds.ToArray)
            Catch ex As Exception
                '
            End Try
            oAppA.ActiveDocument.SetVariable("pickadd", 2)   '' La seleccion actual se suma a la que hubiera.
        End Sub
        Public Sub Selecciona_AcadID(IdEnt As ObjectId())
            oAppA.ActiveDocument.SetVariable("pickadd", 0)   ' Quitar la seleccion que hubiera.
            Try
                Autodesk.AutoCAD.Internal.Utils.SelectObjects(IdEnt)
            Catch ex As Exception
                '
            End Try
            oAppA.ActiveDocument.SetVariable("pickadd", 2)   '' La seleccion actual se suma a la que hubiera.
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
                F1(F1.Length - 1) = 1001 : F2(F2.Length - 1) = regAPPA   ' CType(regAPP, Object)
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
                oSel = oAppA.ActiveDocument.SelectionSets.Add(nSel)
            Catch ex As System.Exception
                oSel = oAppA.ActiveDocument.SelectionSets.Item(nSel)
            End Try
            '
            Me.oSel.Clear()
            Me.oSel.Select(AcSelect.acSelectionSetAll, , , vF1, vF2)
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
                F1(F1.Length - 1) = 1001 : F2(F2.Length - 1) = regAPPA   ' CType(regAPP, Object)
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
                oSel = oAppA.ActiveDocument.SelectionSets.Add(nSel)
            Catch ex As System.Exception
                oSel = oAppA.ActiveDocument.SelectionSets.Item(nSel)
            End Try
            '
            Me.oSel.Clear()
            Me.oSel.Select(AcSelect.acSelectionSetAll, , , vF1, vF2)
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
            F1(1) = 1001 : F2(1) = regAPPA   ' CType(regAPP, Object)
            vF1 = F1
            vF2 = F2
            '
            Try
                oSel = oAppA.ActiveDocument.SelectionSets.Add(nSel)
            Catch ex As System.Exception
                oSel = oAppA.ActiveDocument.SelectionSets.Item(nSel)
            End Try
            '
            Me.oSel.Clear()
            Me.oSel.Select(AcSelect.acSelectionSetAll, , , vF1, vF2)
            '
            If oSel IsNot Nothing AndAlso oSel.Count > 0 Then
                For Each oEnt As AcadEntity In oSel
                    Dim queGrupo As String = Me.XLeeDato(oEnt.Handle, nombreXData)
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
            If oAppA Is Nothing Then _
        oAppA = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
            '
            oAppA.ActiveDocument.SetVariable("pickadd", 0)   '' Solo una selección. Se quita lo que hubiera
            oAppA.ActiveDocument.ActiveSelectionSet.Clear()
            '
            If oSel IsNot Nothing AndAlso oSel.Count > 0 Then
                For x As Integer = 0 To oSel.Count - 1
                    resultado.Add(oAppA.ActiveDocument.ObjectIdToObject(oSel.Item(x).ObjectId.OldIdPtr))
                Next
            Else
                resultado = Nothing
            End If
            ''
            oAppA.ActiveDocument.SetVariable("pickadd", 2)   '' La seleccion actual se suma a la que hubiera.
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
                oSel = oAppA.ActiveDocument.SelectionSets.Add(regAPPA)
            Catch ex As System.Exception
                oSel = oAppA.ActiveDocument.SelectionSets.Item(regAPPA)
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
                    resultado.Add(oAppA.ActiveDocument.ObjectIdToObject(oEnt.ObjectID))
                Next
                oSel.Clear()
                oSel.Delete()
                oSel = Nothing
            End If
            ''
            Return resultado
        End Function
        '
        Public Function SeleccionaDamePolilineasONSCREEN() As ArrayList
            Dim resultado As New ArrayList
            'Dim cSeleccion As AcadSelectionSets
            Dim F1(1) As Short
            Dim F2(1) As Object
            Dim vF1 As Object = Nothing
            Dim vF2 As Object = Nothing

            '' Las 2 maneras valen igual. AcDbBlckReference es mejor (Solo coge bloques) INSERT coge sombreados también.
            F1(0) = 100 : F2(0) = "AcDbBlockReference"
            F1(1) = 0 : F2(1) = "LWPOLYLINE"
            vF1 = F1
            vF2 = F2
            ''
            If oAppA Is Nothing Then _
    oAppA = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
            ''
            ''
            Try
                oSel = oAppA.ActiveDocument.SelectionSets.Add(nSel)
            Catch ex As System.Exception
                oSel = oAppA.ActiveDocument.SelectionSets.Item(nSel)
            End Try
            ''
            oSel.Clear()
            ''
            oAppA.ActiveDocument.SetVariable("pickadd", 2)   '' Solo una selección. Se quita lo que hubiera
            ''
            Try
                oSel.SelectOnScreen(vF1, vF2)
            Catch ex As System.Exception
                'Debug.Print(ex.Message)
            End Try
            'MsgBox(oSel.Count)
            If oSel.Count > 0 Then
                For Each oEnt As AcadEntity In oSel
                    If TypeOf oEnt Is AcadBlockReference Then
                        resultado.Add(CType(oEnt, AcadBlockReference))
                    End If
                Next
                '
                oSel.Clear()
                oSel.Delete()
                oSel = Nothing
            Else
                resultado = Nothing
            End If
            ''
            oAppA.ActiveDocument.SetVariable("pickadd", 2)   '' La seleccion actual se suma a la que hubiera.
            Return resultado
        End Function

        Public Sub SeleccionaBloquesTodos(Optional ByVal nombre As Object = "", Optional ByVal capa As Object = "", Optional ByVal DatosX As Boolean = False)
            'Dim cSeleccion As AcadSelectionSets
            Dim F1(0) As Short
            Dim F2(0) As Object
            Dim vF1 As Object = Nothing
            Dim vF2 As Object = Nothing

            F1(0) = 0 : F2(0) = "INSERT"
            'DatosX Siempre tiene que estar despues de entidad. Si no no funciona
            If DatosX = True Then
                ReDim Preserve F1(F1.Length)
                ReDim Preserve F2(F2.Length)
                'Filtro de bloque, nombre y capa
                F1(F1.Length - 1) = 1001 : F2(F2.Length - 1) = regAPPA
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
            '
            oAppA = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
            ''
            Try
                oSel = oAppA.ActiveDocument.SelectionSets.Add(nSel)
            Catch ex As System.Exception
                oSel = oAppA.ActiveDocument.SelectionSets.Item(nSel)
            End Try
            oSel.Clear()

            Try
                Me.oSel.Select(AcSelect.acSelectionSetAll, , , vF1, vF2)
            Catch ex As Exception
                Debug.Print(ex.Message)
            End Try
        End Sub
        ''
        Public Function SeleccionaDameBloqueEnPuntoXDATA(quePunto As Object, Optional prenombre As String = "*") As AcadBlockReference
            Dim resultado As AcadBlockReference = Nothing
            'Dim cSeleccion As AcadSelectionSets
            Dim F1(2) As Short
            Dim F2(2) As Object
            Dim vF1 As Object = Nothing
            Dim vF2 As Object = Nothing

            '' Las 2 maneras valen igual. AcDbBlckReference es mejor (Solo coge bloques) INSERT coge sombreados también.
            'F1(0) = -4 : F2(0) = "<and"
            F1(0) = 100 : F2(0) = "AcDbBlockReference"
            F1(1) = 0 : F2(1) = "INSERT"
            F1(2) = 2 : F2(2) = prenombre
            'F1(3) = -4 : F2(3) = ">=,>=,*"
            'F1(4) = 10 : F2(4) = pt1
            'F1(5) = -4 : F2(5) = "<=,<=,*"
            'F1(6) = 10 : F2(6) = pt2
            'F1(7) = -4 : F2(7) = "and>"
            'Dim op1 As New TypedValue(DxfCode.Operator, "<and")
            ''
            vF1 = F1
            vF2 = F2
            ''
            oAppA = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
            ''
            Try
                oSel = oAppA.ActiveDocument.SelectionSets.Add(nSel)
            Catch ex As System.Exception
                oSel = oAppA.ActiveDocument.SelectionSets.Item(nSel)
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
                    ''
                    Dim oBl As AcadBlockReference = oAppA.ActiveDocument.ObjectIdToObject(oEnt.ObjectID)
                    Dim pt1 As Object() = {0.0, 0.0}
                    Dim pt2 As Object() = {0.0, 0.0}
                    Dim ptMin As Object = Nothing
                    Dim ptMax As Object = Nothing
                    Dim ptIns As New Autodesk.AutoCAD.Geometry.Point3d
                    oBl.GetBoundingBox(ptMin, ptMax)
                    Dim medida As Double = 0.8
                    'oBl.IntersectWith()
                    pt1(0) = oBl.InsertionPoint(0) - medida
                    pt1(1) = oBl.InsertionPoint(1) - medida
                    pt2(0) = oBl.InsertionPoint(0) + medida
                    pt2(1) = oBl.InsertionPoint(1) + medida
                    '
                    Dim distancia As Double = Math.Sqrt(
                (oBl.InsertionPoint(0) - quePunto(0)) * (oBl.InsertionPoint(0) - quePunto(0)) _
                +
                (oBl.InsertionPoint(1) - quePunto(1)) * (oBl.InsertionPoint(1) - quePunto(1))
                )

                    If (CDbl(oBl.InsertionPoint(0)) = CDbl(quePunto(0)) And CDbl(oBl.InsertionPoint(1)) = CDbl(quePunto(1))) Then
                        resultado = oBl
                        Exit For
                    ElseIf distancia < medida Then
                        resultado = oBl
                        Exit For
                    End If
                Next
                '
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
                oAppA.ActiveDocument.Utility.GetEntity(objCadEnt, vrRetPnt, "Seleccione Polilinea")
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
                    oAppA.ActiveDocument.Utility.Prompt("La polilinea debe tener un minimo de 2 segmentos...")
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
                        objSSet = oAppA.ActiveDocument.SelectionSets.Add("SEL_ENT")
                    Catch ex As System.Exception
                        objSSet = oAppA.ActiveDocument.SelectionSets.Item("SEL_ENT")
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
                    oAppA.ActiveDocument.Activate()
                    objSSet.SelectByPolygon(AcSelect.acSelectionSetWindowPolygon, dblNewCords)
                    objSSet.Highlight(True)
                    'objSSet.Delete
                    '' Mostrar o no mensaje.
                    Dim mensaje As String
                    mensaje = "Nº de Objetos = " & objSSet.Count & vbCrLf & vbCrLf
                    Dim nBlo, nPol, nSom, nLin As Integer
                    For x = 0 To objSSet.Count - 1
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
        '' Seleccionamos una polilinea cerrada o no y la usamos
        '' para hacer una seleccion Borde (Todo lo que este tocando) y obtener todos
        '' los objetos AutoCAD que hay dentro.
        Public Function SeleccionaDameEntitiesTocandoPolilinea(Optional conMensaje As Boolean = False, Optional solobloques As Boolean = True) As ArrayList
            Dim resultado As New ArrayList
            ''
repetir:
            Dim objCadEnt As AcadEntity = Nothing
            Dim vrRetPnt As Object = Nothing
            Try
                oAppA.ActiveDocument.Utility.GetEntity(objCadEnt, vrRetPnt, "Seleccione Polilinea")
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
                '
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
                    objSSet = oAppA.ActiveDocument.SelectionSets.Add("SEL_ENT")
                Catch ex As System.Exception
                    objSSet = oAppA.ActiveDocument.SelectionSets.Item("SEL_ENT")
                End Try
                '' Quitamos los objetos que hubiera seleccionados.
                oAppA.ActiveDocument.SetVariable("pickadd", 0)   '' Solo una selección. Se quita lo que hubiera
                objSSet.Clear()
                ''
                '' Para filtrar entidades
                Dim gpCode(0) As Short
                Dim dataValue(0) As Object
                gpCode(0) = 100 : dataValue(0) = "AcDbBlockReference"
                'gpCode(0) = 0 : dataValue(0) = "Insert"
                Dim groupCode As Object, dataCode As Object
                groupCode = gpCode
                dataCode = dataValue
                oAppA.ActiveDocument.Activate()
                If solobloques Then
                    '' Solo Bloques
                    objSSet.SelectByPolygon(AcSelect.acSelectionSetFence, dblNewCords, groupCode, dataCode)
                Else
                    '' Todos
                    objSSet.SelectByPolygon(AcSelect.acSelectionSetFence, dblNewCords)
                End If
                ''
                objSSet.Highlight(True)
                'objSSet.Delete
                '' Mostrar o no mensaje.
                Dim mensaje As String
                mensaje = "Nº de Objetos = " & objSSet.Count & vbCrLf & vbCrLf
                Dim nBlo, nPol, nSom, nLin As Integer
                For x = 0 To objSSet.Count - 1
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
            'End If
            ''
            oAppA.ActiveDocument.SetVariable("pickadd", 2)   '' Restaurar varias selecciones.
            Return resultado
        End Function
        Public Function SeleccionaDameBloqueUno(Optional ByVal nombre As Object = "", Optional ByVal capa As Object = "") As AcadBlockReference
            Dim resultado As AcadBlockReference = Nothing
            'Dim cSeleccion As AcadSelectionSets
            Dim F1(1) As Short
            Dim F2(1) As Object
            Dim vF1 As Object = Nothing
            Dim vF2 As Object = Nothing

            '' Las 2 maneras valen igual. AcDbBlckReference es mejor (Solo coge bloques) INSERT coge sombreados también.
            F1(0) = 100 : F2(0) = "AcDbBlockReference"
            F1(1) = 0 : F2(1) = "INSERT"
            'DatosX Siempre tiene que estar despues de entidad. Si no no funciona
            'If DatosX = True Then
            '    ReDim Preserve F1(F1.Length)
            '    ReDim Preserve F2(F2.Length)
            '    'Filtro de bloque, nombre y capa
            '    F1(F1.Length - 1) = 1001 : F2(F2.Length - 1) = regAPP
            'End If
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
            If oAppA Is Nothing Then _
    oAppA = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
            ''
            ''
            Try
                oSel = oAppA.ActiveDocument.SelectionSets.Add(nSel)
            Catch ex As System.Exception
                oSel = oAppA.ActiveDocument.SelectionSets.Item(nSel)
            End Try
            ''
            oAppA.ActiveDocument.SetVariable("pickadd", 0)   '' Solo una selección. Se quita lo que hubiera
            oSel.Clear()
            Try
                oSel.Select(vF1, vF2)
            Catch ex As System.Exception
                'Debug.Print(ex.Message)
            End Try
            'MsgBox(oSel.Count)
            If oSel.Count > 0 Then
                'Dim texto As String = XData.XLeeDato(oEnt, xT.TEXTOS)
                'If TypeOf oEnt Is AcadTable Or texto = "Clase=tabla" Then
                If TypeOf oSel.Item(oSel.Count - 1) Is AcadTable Then
                    resultado = Nothing
                Else
                    resultado = oSel.Item(oSel.Count - 1)
                End If
                '
                oSel.Clear()
                oSel.Delete()
                oSel = Nothing
            Else
                resultado = Nothing
            End If
            ''
            oAppA.ActiveDocument.SetVariable("pickadd", 2)   '' La seleccion actual se suma a la que hubiera.
            Return resultado
        End Function

        Public Function SeleccionaDameBloquesTODOS_ListAcadBlockReference(Optional nombreApp As String = "*", Optional nombrebloque As String = "*", Optional nombrecapa As String = "*") As List(Of AcadBlockReference)
            Dim resultado As New List(Of AcadBlockReference)
            Dim temp As ArrayList = SeleccionaDameBloquesTODOS(nombreApp, nombrebloque, nombrecapa)
            If temp Is Nothing OrElse temp.Count = 0 Then
                resultado = Nothing
            Else
                For Each obj As AcadObject In temp
                    resultado.Add(CType(obj, AcadBlockReference))
                Next
            End If
            Return resultado
        End Function
        '
        ''' <summary>
        ''' Devuelve arrayList con todos los bloques que cumplan el criterio
        ''' nombreApp = '*' por defecto. Le podemos indicar un nombre de APP registrada (1001=nombreApp)
        ''' nombrebloque = '*' por defecto. Le podemos indicar un nombre de bloque
        ''' nombrecapa = '*' por defecto. Le podemos indicar un nombre de capa
        ''' ** Le podemos indicar carácterés comodin (Ej.: nombrecapa=planta*) Utilizar * o ?
        ''' </summary>
        ''' <param name="nombreApp">Nombre de la app que registro el XData. O filtro con carácteres comodín</param>
        ''' <param name="nombrebloque">Nombre del bloque o filtro con carácteres comodin</param>
        ''' <param name="nombrecapa">Nombre de la capa o filtro con carácteres comodin</param>
        ''' <returns></returns>
        Public Function SeleccionaDameBloquesTODOS(Optional nombreApp As String = "*", Optional nombrebloque As String = "*", Optional nombrecapa As String = "*") As ArrayList
            Dim resultado As New ArrayList
            'Dim cSeleccion As AcadSelectionSets
            Dim F1(4) As Short
            Dim F2(4) As Object
            Dim vF1 As Object = Nothing
            Dim vF2 As Object = Nothing

            '' Las 2 maneras valen igual. AcDbBlckReference es mejor (Solo coge bloques) INSERT coge sombreados también.
            F1(0) = 100 : F2(0) = "AcDbBlockReference"
            F1(1) = 0 : F2(1) = "INSERT"
            F1(2) = 1001 : F2(2) = nombreApp
            F1(3) = 2 : F2(3) = IIf(nombrebloque = "*", "*", IIf(nombrebloque.EndsWith("*"), nombrebloque & "*", nombrebloque)) '& ",`*U*"      ' ,`*U* Para añadir también los bloques dinámicos.
            F1(4) = 8 : F2(4) = nombrecapa
            ''
            vF1 = F1
            vF2 = F2
            '
            Try
                oSel = oAppA.ActiveDocument.SelectionSets.Add(regAPPA)
            Catch ex As System.Exception
                oSel = oAppA.ActiveDocument.SelectionSets.Item(regAPPA)
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
                    resultado.Add(oAppA.ActiveDocument.ObjectIdToObject(oEnt.ObjectID))
                Next
                '
                oSel.Clear()
                oSel.Delete()
                oSel = Nothing
            End If
            ''
            Return resultado
        End Function
        '
        Public Function SeleccionaDameBloquesTODOStexto(Optional nombreApp As String = "*", Optional nombrebloque As String = "*", Optional nombrecapa As String = "*", Optional textoXdata As String = "*") As ArrayList
            Dim resultado As New ArrayList
            'Dim cSeleccion As AcadSelectionSets
            Dim F1(5) As Short
            Dim F2(5) As Object
            Dim vF1 As Object = Nothing
            Dim vF2 As Object = Nothing

            '' Las 2 maneras valen igual. AcDbBlckReference es mejor (Solo coge bloques) INSERT coge sombreados también.
            F1(0) = 100 : F2(0) = "AcDbBlockReference"
            F1(1) = 0 : F2(1) = "INSERT"
            F1(2) = 1001 : F2(2) = nombreApp
            F1(3) = 2 : F2(3) = nombrebloque '& ",`*U*"      ' ,`*U* Para añadir también los bloques dinámicos.
            F1(4) = 8 : F2(4) = nombrecapa
            F1(5) = 1000 : F2(5) = textoXdata
            ''
            vF1 = F1
            vF2 = F2
            '
            Try
                oSel = oAppA.ActiveDocument.SelectionSets.Add(regAPPA)
            Catch ex As System.Exception
                oSel = oAppA.ActiveDocument.SelectionSets.Item(regAPPA)
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
                    resultado.Add(oAppA.ActiveDocument.ObjectIdToObject(oEnt.ObjectID))
                Next
                '
                oSel.Clear()
                oSel.Delete()
                oSel = Nothing
            End If
            ''
            Return resultado
        End Function
        '
        Public Function SeleccionaDameBloquesPorNombreID(Optional nombrebloque As String = "*") As List(Of Long)
            Dim resultado As New List(Of Long)
            'Dim cSeleccion As AcadSelectionSets
            Dim F1(4) As Short
            Dim F2(4) As Object
            Dim vF1 As Object = Nothing
            Dim vF2 As Object = Nothing
            ' F1(x) = 100 : F2(x) = "AcDbBlockReference"
            ' F1(y) = 0 : F2(y) = "INSERT"
            '' Las 2 maneras valen igual. AcDbBlckReference es mejor (Solo coge bloques) INSERT coge sombreados también.
            F1(0) = -4 : F2(0) = "<AND"
            F1(1) = 100 : F2(1) = "AcDbBlockReference"  ' Solo bloques
            F1(2) = 0 : F2(2) = "INSERT"                ' Bloques y sombreados
            F1(3) = 2 : F2(3) = nombrebloque            ' Solo con este nombre
            F1(4) = -4 : F2(4) = "AND>"
            ''
            vF1 = F1
            vF2 = F2
            '
            Try
                oSel = oAppA.ActiveDocument.SelectionSets.Add(regAPPA)
            Catch ex As System.Exception
                oSel = oAppA.ActiveDocument.SelectionSets.Item(regAPPA)
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
                    resultado.Add(oEnt.ObjectID)
                Next
                '
                oSel.Clear()
                oSel.Delete()
                oSel = Nothing
            End If
            ''
            Return resultado
        End Function

        Public Function SeleccionaDameHandle_PorNombreBloque(Optional nombrebloque As String = "*") As List(Of String)
            Dim resultado As New List(Of String)
            'Dim cSeleccion As AcadSelectionSets
            Dim F1(4) As Short
            Dim F2(4) As Object
            Dim vF1 As Object = Nothing
            Dim vF2 As Object = Nothing
            ' F1(x) = 100 : F2(x) = "AcDbBlockReference"
            ' F1(y) = 0 : F2(y) = "INSERT"
            '' Las 2 maneras valen igual. AcDbBlckReference es mejor (Solo coge bloques) INSERT coge sombreados también.
            F1(0) = -4 : F2(0) = "<AND"
            F1(1) = 100 : F2(1) = "AcDbBlockReference"  ' Solo bloques
            F1(2) = 0 : F2(2) = "INSERT"                ' Bloques y sombreados
            F1(3) = 2 : F2(3) = nombrebloque            ' Solo con este nombre
            F1(4) = -4 : F2(4) = "AND>"
            ''
            vF1 = F1
            vF2 = F2
            '
            Try
                oSel = oAppA.ActiveDocument.SelectionSets.Add(regAPPA)
            Catch ex As System.Exception
                oSel = oAppA.ActiveDocument.SelectionSets.Item(regAPPA)
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
                    resultado.Add(oEnt.Handle)
                Next
                '
                oSel.Clear()
                oSel.Delete()
                oSel = Nothing
            End If
            ''
            Return resultado
        End Function
        '
        Public Function Selecciona_TodoEnPunto(quePunto As Double, Optional acTipo As String = "*", Optional Tipo As String = "*") As List(Of Long)
            Dim resultado As New List(Of Long)
            'Dim cSeleccion As AcadSelectionSets
            Dim F1(1) As Short
            Dim F2(1) As Object
            Dim vF1 As Object = Nothing
            Dim vF2 As Object = Nothing

            '' Las 2 maneras valen igual. AcDbBlckReference es mejor (Solo coge bloques) INSERT coge sombreados también.
            F1(0) = 100 : F2(0) = acTipo
            F1(1) = 0 : F2(1) = Tipo
            vF1 = F1
            vF2 = F2
            ''
            If oAppA Is Nothing Then _
    oAppA = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
            ''
            ''
            Try
                oSel = oAppA.ActiveDocument.SelectionSets.Add(regAPPA)
            Catch ex As System.Exception
                oSel = oAppA.ActiveDocument.SelectionSets.Item(regAPPA)
            End Try
            ''
            oSel.Clear()
            ''
            oAppA.ActiveDocument.SetVariable("pickadd", 0)   '' Solo una selección. Se quita lo que hubiera
            ''
            Try
                'Dim point() As Double = {10000, 10000, 100000}
                oSel.SelectAtPoint(quePunto, vF1, vF2)
            Catch ex As System.Exception
                'Debug.Print(ex.Message)
            End Try
            'MsgBox(oSel.Count)
            If oSel.Count > 0 Then
                For Each oEnt As AcadEntity In oSel
                    If TypeOf oEnt Is AcadBlockReference Then
                        resultado.Add(oEnt.ObjectID)
                    End If
                Next
                '
                oSel.Clear()
                oSel.Delete()
                oSel = Nothing
            Else
                resultado = Nothing
            End If
            ''
            oAppA.ActiveDocument.SetVariable("pickadd", 2)   '' La seleccion actual se suma a la que hubiera.
            Return resultado
        End Function
        Public Function SeleccionaDameBloquesONSCREEN() As ArrayList
            Dim resultado As New ArrayList
            'Dim cSeleccion As AcadSelectionSets
            Dim F1(1) As Short
            Dim F2(1) As Object
            Dim vF1 As Object = Nothing
            Dim vF2 As Object = Nothing

            '' Las 2 maneras valen igual. AcDbBlckReference es mejor (Solo coge bloques) INSERT coge sombreados también.
            F1(0) = 100 : F2(0) = "AcDbBlockReference"
            F1(1) = 0 : F2(1) = "INSERT"
            vF1 = F1
            vF2 = F2
            ''
            If oAppA Is Nothing Then _
    oAppA = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
            ''
            ''
            Try
                oSel = oAppA.ActiveDocument.SelectionSets.Add(nSel)
            Catch ex As System.Exception
                oSel = oAppA.ActiveDocument.SelectionSets.Item(nSel)
            End Try
            ''
            oSel.Clear()
            ''
            oAppA.ActiveDocument.SetVariable("pickadd", 0)   '' Solo una selección. Se quita lo que hubiera
            ''
            Try
                oSel.SelectOnScreen(vF1, vF2)
            Catch ex As System.Exception
                'Debug.Print(ex.Message)
            End Try
            'MsgBox(oSel.Count)
            If oSel.Count > 0 Then
                For Each oEnt As AcadEntity In oSel
                    If TypeOf oEnt Is AcadBlockReference Then
                        resultado.Add(CType(oEnt, AcadBlockReference))
                    End If
                Next
                '
                oSel.Clear()
                oSel.Delete()
                oSel = Nothing
            Else
                resultado = Nothing
            End If
            ''
            oAppA.ActiveDocument.SetVariable("pickadd", 2)   '' La seleccion actual se suma a la que hubiera.
            Return resultado
        End Function
        Public Function SeleccionaDameEntitiesONSCREEN(Optional solouna As Boolean = True) As ArrayList
            ''
            AutoCAD_PonFoCo()
            Dim resultado As New ArrayList
            'Dim cSeleccion As AcadSelectionSets
            Dim F1(1) As Short
            Dim F2(1) As Object
            Dim vF1 As Object = Nothing
            Dim vF2 As Object = Nothing

            '' Las 2 maneras valen igual. AcDbBlckReference es mejor (Solo coge bloques) INSERT coge sombreados también.
            F1(0) = 100 : F2(0) = "*"
            F1(1) = 0 : F2(1) = "*"
            vF1 = F1
            vF2 = F2
            ''
            If oAppA Is Nothing Then _
    oAppA = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
            ''
            ''
            Try
                oSel = oAppA.ActiveDocument.SelectionSets.Add(nSel)
            Catch ex As System.Exception
                oSel = oAppA.ActiveDocument.SelectionSets.Item(nSel)
            End Try
            ''
            oSel.Clear()
            ''
            If solouna = True Then
                oAppA.ActiveDocument.SetVariable("pickadd", 0)   '' Solo una selección. Se quita lo que hubiera
            Else
                oAppA.ActiveDocument.SetVariable("pickadd", 2)   '' La seleccion actual se suma a la que hubiera.
            End If

            Try
                oSel.SelectOnScreen(vF1, vF2)
            Catch ex As System.Exception
                'Debug.Print(ex.Message)
            End Try
            'MsgBox(oSel.Count)
            If oSel.Count > 0 Then
                For Each oEnt As AcadEntity In oSel
                    resultado.Add(oEnt)
                Next
                '
                oSel.Clear()
                oSel.Delete()
                oSel = Nothing
            Else
                resultado = Nothing
            End If
            ''
            oAppA.ActiveDocument.SetVariable("pickadd", 2)   '' La seleccion actual se suma a la que hubiera.
            Return resultado
        End Function

        'Añadido AFLETA
        Public Function ObjectId_SelectOne(ByVal pModo As OpenMode, ByVal pArrayType() As Type) As ObjectId

            Dim doc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim db As Database = doc.Database
            Dim ed As Editor = doc.Editor
            Dim result As ObjectId
            Dim strMensaje = ""
            Try

                Using acTrans As Transaction = db.TransactionManager.StartOpenCloseTransaction()
                    Dim acBlkTbl As BlockTable = CType(acTrans.GetObject(db.BlockTableId, OpenMode.ForRead), BlockTable)
                    Dim opts As PromptEntityOptions = New PromptEntityOptions("")

                    opts.SetRejectMessage("Only one Entity.")
                    For Each i As Type In pArrayType

                        opts.AddAllowedClass(i, True)
                        strMensaje = strMensaje & i.Name & " "
                    Next

                    opts.Message = vbLf & "Select " & strMensaje & ": "


                    Dim selRes As PromptEntityResult = ed.GetEntity(opts)

                    If selRes.Status <> PromptStatus.OK Then
                        ed.WriteMessage(vbLf & "Selected " & strMensaje & " FAILED.")
                        Return Nothing
                    End If

                    result = selRes.ObjectId
                End Using

                ed.WriteMessage(vbLf & "Selected " & strMensaje & " CORRECT.")

                Return result

            Catch ex As Autodesk.AutoCAD.Runtime.Exception
                'ed.WriteMessage(ex.Message.ToString())
                Return Nothing
            End Try
        End Function
        '
        Public Sub Selecctionar_Zoom_Objetos(arrIds() As ObjectId, Optional hazzoom As Boolean = True)
            Call Autodesk.AutoCAD.Internal.Utils.SelectObjects(arrIds)
            If hazzoom Then
                Autodesk.AutoCAD.Internal.Utils.ZoomObjects(False)
                oAppA.ZoomScaled(3, AcZoomScaleType.acZoomScaledRelative)
            End If
        End Sub
    End Class
End Namespace