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
#Region "BLOQUES"
        ''' <summary>
        ''' Inserta un bloque en el punto indicado. Le pasamos punto de insercion, nombre completo del bloque,
        ''' escala x, y, z y rotación
        ''' </summary>
        ''' <param name="pt">Punto de inserción (Array(2) de coordenadas)</param>
        ''' <param name="nombre">Nombre solo (Tienen que estar ya cargado) o fullPath (Si no está cargado)</param>
        ''' <param name="sX">Escala X</param>
        ''' <param name="sY">Escala Y</param>
        ''' <param name="sZ">Escala Z</param>
        ''' <param name="rotacion">Rotación (En radianes)</param>
        Public Sub Bloque_Inserta(Optional ByVal pt() As Double = Nothing, Optional ByVal nombre As String = "",
                                 Optional ByVal sX As Double = 1.0#, Optional ByVal sY As Double = 1.0#,
                                 Optional ByVal sZ As Double = 1.0#, Optional ByVal rotacion As Double = 0)

            'Me.LiberaApp(True)
            Dim oDoc As Autodesk.AutoCAD.Interop.AcadDocument = oAppA.ActiveDocument
            'Me.oDoc.SetVariable("NOMUTT", 1)
            'AppActivate(clsA.oAppT)
            Dim oBl As AcadBlock = Nothing
            Dim existe As Boolean = True
            If IO.File.Exists(nombre) = False Then
                existe = False
                Try
                    ' Para saber si ya está cargado (Solo el nombre)
                    oBl = oDoc.Blocks.Item(nombre)
                    existe = True
                Catch ex As Exception
                    ' No existe este bloque cargado
                    existe = False
                End Try
            End If
            '
            If existe = False Then
                MsgBox("No existe el bloque " & nombre, MsgBoxStyle.Critical)
                oBl = Nothing
                VaciaMemoria()
                Exit Sub
            End If
            '
            Try
                Dim oBlr As AcadBlockReference
                '********************************************************
                'oDoc.SendCommand(Chr(27))
                If pt Is Nothing Then pt = oAppA.ActiveDocument.Utility.GetPoint(, "Punto Inserción Bloque")
                If pt Is Nothing Then Exit Sub
                Try
                    oBlr = oDoc.ModelSpace.InsertBlock(pt, nombre, sX, sY, sZ, rotacion)
                    VaciaMemoria()
                Catch ex As Exception
                    Exit Sub
                End Try
                '
                oBlult = oBlr
                'Me.oDoc.Blocks.Item(acb.Name).Explodable = False
                'acb.Layer = preCapa & "TEMP"
                'XPonDato(acb, xT.CAPA, preCapa & "TEMP")     'Ponemos el XDato "SERICAD" como aplicación registrado (codigo 1001)
            Catch ex As Exception
                'No hacemos nada. Ya que el usuario a podido cancelar la inserción.
            Finally
                VaciaMemoria()
                'Me.oDoc.SetVariable("NOMUTT", 0)
            End Try
        End Sub

        ''' <summary>
        ''' Inserta un bloque en el punto indicado. Le pasamos punto de insercion, nombre completo del bloque,
        ''' escala x, y, z y rotación
        ''' </summary>
        ''' <param name="pt">Punto de inserción (Array(2) de coordenadas)</param>
        ''' <param name="nombre">Nombre solo (Tienen que estar ya cargado) o fullPath (Si no está cargado)</param>
        ''' <param name="sX">Escala X</param>
        ''' <param name="sY">Escala Y</param>
        ''' <param name="sZ">Escala Z</param>
        ''' <param name="rotacion">Rotación (En radianes)</param>
        Public Function Bloque_InsertaMultiple(Optional ByVal pt() As Double = Nothing, Optional ByVal nombre As String = "",
                                 Optional ByVal sX As Double = 1.0#, Optional ByVal sY As Double = 1.0#,
                                 Optional ByVal sZ As Double = 1.0#, Optional ByVal rotacion As Double = 0) As AcadBlockReference

            'Me.LiberaApp(True)
            Dim oDoc As Autodesk.AutoCAD.Interop.AcadDocument = oAppA.ActiveDocument
            'Me.oDoc.SetVariable("NOMUTT", 1)
            'AppActivate(clsA.oAppT)
            Dim oBl As AcadBlock = Nothing
            Dim oBlr As AcadBlockReference = Nothing
            Dim existe As Boolean = True
            If IO.File.Exists(nombre) = False Then
                existe = False
                Try
                    ' Para saber si ya está cargado (Solo el nombre)
                    oBl = oDoc.Blocks.Item(nombre)
                    existe = True
                Catch ex As Exception
                    ' No existe este bloque cargado
                    existe = False
                End Try
            End If
            '
            If existe = False Then
                MsgBox("No existe el bloque " & nombre, MsgBoxStyle.Critical)
                oBl = Nothing
                VaciaMemoria()
                Return Nothing
            End If
            '
            Try
                '********************************************************
                'oDoc.SendCommand(Chr(27))
                If pt Is Nothing Then pt = oAppA.ActiveDocument.Utility.GetPoint(, "Punto Inserción Bloque")
                If pt Is Nothing Then
                    Return Nothing
                End If
                Try
                    oBlr = oDoc.ModelSpace.InsertBlock(pt, nombre, sX, sY, sZ, rotacion)
                    VaciaMemoria()
                Catch ex As Exception
                    Return Nothing
                End Try
                '
                oBlult = oBlr
                'Me.oDoc.Blocks.Item(acb.Name).Explodable = False
                'acb.Layer = preCapa & "TEMP"
                'XPonDato(acb, xT.CAPA, preCapa & "TEMP")     'Ponemos el XDato "SERICAD" como aplicación registrado (codigo 1001)
            Catch ex As Exception
                'No hacemos nada. Ya que el usuario a podido cancelar la inserción.
            Finally
                VaciaMemoria()
                'Me.oDoc.SetVariable("NOMUTT", 0)
            End Try
            Return oBlr
        End Function
        ''
        Public Function Bloque_InsertaDerecha(queFiBlo As String, queBloEsta As AcadBlockReference, queEscala As Double, queDistDcha As Double) As AcadBlockReference
            Dim resultado As AcadBlockReference = Nothing
            If oAppA Is Nothing Then _
            oAppA = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
            ''
            ''
            '' Buscamos este bloque en el dibujo actual. Y, si no está,
            '' lo cargamos del disco duro (queFiBlo = camino completo al fichero DWG)
            Dim nombre As String = IO.Path.GetFileNameWithoutExtension(queFiBlo)
            For Each oBl As AcadBlock In oAppA.ActiveDocument.Blocks
                If oBl.Name = nombre Then
                    queFiBlo = nombre
                    VaciaMemoria()
                    Exit For
                End If
            Next
            ''
            Dim oApp As Autodesk.AutoCAD.Interop.AcadApplication =
                CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication,
    Autodesk.AutoCAD.Interop.AcadApplication)
            '
            Dim puntoInserta(0 To 2) As Double
            puntoInserta(0) = queBloEsta.InsertionPoint(0) + queDistDcha
            puntoInserta(1) = queBloEsta.InsertionPoint(1)
            puntoInserta(2) = 0
            '
            resultado = oAppA.ActiveDocument.ActiveLayout.Block.InsertBlock(puntoInserta, queFiBlo, queEscala, queEscala, queEscala, 0)
            oAppA.ActiveDocument.Regen(AcRegenType.acActiveViewport)
            '
            VaciaMemoria()
            Return resultado
        End Function
        Public Sub Bloque_BorraLimpia(ByVal nombre As String)
            SeleccionaBloquesTodos(nombre)
            If Me.oSel.Count > 0 Then
                For Each ao As AcadObject In Me.oSel
                    ao.Delete()
                Next
                '
                oSel.Clear()
                oSel.Delete()
                oSel = Nothing
            End If
            Try
                Me.oDoc.Blocks.Item(nombre).Delete()
                VaciaMemoria()
            Catch ex As Exception

            End Try
        End Sub
        Public Function Bloque_SeleccionaDame(Optional conconfirmacion As Boolean = False) As AcadBlockReference
            ' Begin the selection
            Dim bloque As AcadBlockReference = Nothing
            Dim obj As AcadObject = Nothing
            Dim basePnt As Object = Nothing

            ' Dim doc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            'Using acLckDoc As DocumentLock = doc.LockDocument()
            ' Foco en AutoCAD
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()

            ' The following example waits for a selection from the user
            On Error Resume Next
RETRY:
            oAppA.ActiveDocument.Utility.GetEntity(obj, basePnt, vbCrLf & "Seleccione Bloque : ")
            'MessageBox.Show(obj.ObjectName)

            If Err.Number <> 0 Then
                Err.Clear()
                'MsgBox("Program ended.", , "GetEntity Example")
                VaciaMemoria()
                Return bloque
                Exit Function
            End If
            '
            If obj.ObjectName = "AcDbBlockReference" Then
                If conconfirmacion = True Then
                    Dim resultado As String = ""
                    resultado = oAppA.ActiveDocument.Utility.GetString(False, vbLf & "[Intro] o 'S' Acepta Selección / 'N' anula selección --> : ")
                    If resultado.ToUpper = "N" Then
                        obj = Nothing
                        GoTo RETRY
                    End If
                End If
                bloque = CType(oAppA.ActiveDocument.HandleToObject(obj.Handle), AcadBlockReference)
                bloque.Update()
            Else
                obj = Nothing
                basePnt = Nothing
                GoTo RETRY
            End If
            'End Using
            'doc = Nothing
            VaciaMemoria()
            Return bloque
        End Function
        Public Function Bloques_DameNombreContiene(quenombre As String, Optional exacto As Boolean = False) As List(Of AcadBlockReference)
            Dim resultado As New List(Of AcadBlockReference)
            If oAppA Is Nothing Then _
            oAppA = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
            ''
            For Each oEnt As AcadEntity In oAppA.ActiveDocument.ModelSpace
                If TypeOf oEnt Is AcadBlockReference Then
                    Dim oBl As AcadBlockReference = oAppA.ActiveDocument.ObjectIdToObject(oEnt.ObjectID)
                    If exacto = True Then
                        If oBl.EffectiveName.ToUpper = quenombre.ToUpper OrElse oBl.Name.ToUpper = quenombre.ToUpper Then
                            resultado.Add(oBl)
                        End If
                    Else
                        If oBl.EffectiveName.ToUpper.Contains(quenombre.ToUpper) OrElse oBl.Name.ToUpper.Contains(quenombre.ToUpper) Then
                            resultado.Add(oBl)
                        End If
                    End If
                    oBl = Nothing
                End If
            Next
            VaciaMemoria()
            Return resultado
        End Function

        Public Function Bloques_DameNombreContiene(quenombres() As String) As List(Of AcadBlockReference)
            Dim resultado As New List(Of AcadBlockReference)
            If oAppA Is Nothing Then _
            oAppA = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
            '
            For Each oEnt As AcadEntity In oAppA.ActiveDocument.ModelSpace
                If TypeOf oEnt Is AcadBlockReference Then
                    Dim oBl As AcadBlockReference = oAppA.ActiveDocument.ObjectIdToObject(oEnt.ObjectID)
                    For Each nombre In quenombres
                        If oBl.EffectiveName.ToUpper.Contains(nombre.ToUpper) OrElse oBl.Name.ToUpper.Contains(nombre.ToUpper) Then
                            resultado.Add(oBl)
                        End If
                        oBl = Nothing
                    Next
                End If
                System.Windows.Forms.Application.DoEvents()
            Next
            VaciaMemoria()
            Return resultado
        End Function
        Public Function Bloque_Cambia(ByRef blq As AcadBlockReference, fullNuevo As String, ByVal todo As Boolean) As AcadBlockReference
            Dim resultado As AcadBlockReference = Nothing
            Control.CheckForIllegalCrossThreadCalls = False
            'Me.ultBlv = blq        'El ultimo bloque a cambiar (borrar) 'Me.ultBl              'El ultimo bloque insertado en el dibujo
            Dim nombreViejo As String = blq.Name
            If todo = True Then
                Dim contador As Integer = 0
                Dim nBloques As Integer = 0

                Me.SeleccionaBloquesTodos(nombreViejo)
                For Each obj As AcadObject In Me.oSel
                    If TypeOf obj Is AcadBlockReference Then ' obj.ObjectName = "AcDbBlockReference" Then
                        Dim objB As AcadBlockReference = oAppA.ActiveDocument.HandleToObject(obj.Handle)    ' Me.oDoc.ObjectIdToObject(obj.ObjectID)
                        If objB.Name = nombreViejo Then
                            Try
                                Me.Bloque_Inserta(objB.InsertionPoint, fullNuevo, objB.XScaleFactor, objB.YScaleFactor, objB.ZScaleFactor, objB.Rotation)
                            Catch ex As Exception
                                Me.Bloque_Inserta(objB.InsertionPoint, fullNuevo, objB.XScaleFactor, objB.YScaleFactor, objB.ZScaleFactor, objB.Rotation)
                            End Try
                            Me.oBlult.Layer = objB.Layer
                            Me.oBlult.Update()
                            objB.Delete()
                            nBloques += 1
                        End If
                    End If
                    contador += 1
                    System.Windows.Forms.Application.DoEvents()
                Next
                '
                oSel.Clear()
                oSel.Delete()
                oSel = Nothing
                oAppA.ActiveDocument.Utility.Prompt(vbCrLf & "( " & nBloques & " ) Cambiados" & vbCrLf)
                'oDoc.SendCommand(Chr(27))
            Else
                Try
                    Me.Bloque_Inserta(blq.InsertionPoint, fullNuevo, blq.XScaleFactor, blq.YScaleFactor, blq.ZScaleFactor, blq.Rotation)
                Catch ex As Exception
                    Me.Bloque_Inserta(blq.InsertionPoint, fullNuevo, blq.XScaleFactor, blq.YScaleFactor, blq.ZScaleFactor, blq.Rotation)
                End Try
                Me.oBlult.Layer = blq.Layer
                Me.oBlult.Update()
                blq.Delete()
                'ActivaApp(False)
            End If
            oAppA.ActiveDocument.Regen(AcRegenType.acActiveViewport)
            VaciaMemoria()
            Return oBlult
        End Function

        Public Function Bloque_AtributoDame(queId As Long, nombreAtri As String) As String
            Dim resultado As String = ""
            ''
            If oAppA Is Nothing Then _
            oAppA = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
            ''
            Dim oBl As Autodesk.AutoCAD.Interop.Common.AcadBlockReference = oAppA.ActiveDocument.ObjectIdToObject(queId)

            ' Cargamos los atributos
            Dim varAttributes As Object
            varAttributes = oBl.GetAttributes
            Dim consAttributes As Object
            consAttributes = oBl.GetConstantAttributes

            ' Localizar los atributos y añadir valores a la colección (ORDEN y NUMERO). O todos
            '' Atributos editables.
            Dim oAtri As Object ' Autodesk.AutoCAD.Interop.Common.AcadAttribute
            For I = LBound(varAttributes) To UBound(varAttributes)
                oAtri = varAttributes(I)
                If oAtri.TagString.ToString.ToUpper = nombreAtri.ToUpper Then
                    resultado = oAtri.TextString
                    Exit For
                End If
            Next
            '' Atributos constantes.
            For I = LBound(consAttributes) To UBound(consAttributes)
                oAtri = consAttributes(I)
                If oAtri.TagString.ToString.ToUpper = nombreAtri.ToUpper Then
                    resultado = oAtri.TextString
                    Exit For
                End If
            Next
            ''
            VaciaMemoria()
            oAtri = Nothing
            oBl = Nothing
            ''
            Return resultado
        End Function
        '

        Public Function Bloque_AtributoDame(arrAtributos As Object, nombreAtri As String, Optional arrAtributosConstantes As Object = Nothing) As String
            Dim resultado As String = ""
            '
            ' Localizar los atributos y añadir valores a la colección (ORDEN y NUMERO). O todos
            '' Atributos editables.
            Dim oAtri As Object ' Autodesk.AutoCAD.Interop.Common.AcadAttribute
            For I = LBound(arrAtributos) To UBound(arrAtributos)
                oAtri = arrAtributos(I)
                If oAtri.TagString.ToString.ToUpper = nombreAtri.ToUpper Then
                    resultado = oAtri.TextString
                    Exit For
                End If
            Next
            If resultado = "" AndAlso arrAtributosConstantes IsNot Nothing Then
                ' Atributos constantes.
                For I = LBound(arrAtributosConstantes) To UBound(arrAtributosConstantes)
                    oAtri = arrAtributosConstantes(I)
                    If oAtri.TagString.ToString.ToUpper = nombreAtri.ToUpper Then
                        resultado = oAtri.TextString
                        Exit For
                    End If
                Next
            End If
            ''
            VaciaMemoria()
            oAtri = Nothing
            ''
            Return resultado
        End Function
        '
        'Public Function Bloque_DameDato_AttPropX(h As String, arrAtt As Object, arrPro As Object, nameprobusco As String, Optional arrAttCons As Object = Nothing) As String
        '    Dim resultado As String = ""
        '    ' Buscar en XData
        '    resultado = XLeeDato(h, nameprobusco)
        '    ' Buscar en Atributos
        '    If resultado = "" AndAlso arrAtt IsNot Nothing Then
        '        resultado = Bloque_AtributoDame(arrAtt, nameprobusco, arrAttCons)
        '    End If
        '    ' Buscar en Parametros
        '    If resultado = "" AndAlso arrPro IsNot Nothing Then
        '        resultado = BloqueDinamico_ParametroDame(arrPro, nameprobusco)
        '    End If

        '    Return resultado
        'End Function
        Public Function Bloque_DameDato_AttPropX(oBlD As Bloque_Datos, nameprobusco As String) As String
            Dim resultado As String = ""
            ' Buscar en XData
            resultado = XLeeDato(oBlD.handle, nameprobusco)
            ' Buscar en Atributos
            If resultado = "" AndAlso oBlD.atributos IsNot Nothing Then
                resultado = Bloque_AtributoDame(oBlD.atributos, nameprobusco, oBlD.atributosConst)
            End If
            ' Buscar en Parametros
            If resultado = "" AndAlso oBlD.propiedades IsNot Nothing Then
                resultado = BloqueDinamico_ParametroDame(oBlD.propiedades, nameprobusco)
            End If

            Return resultado
        End Function
        '' Devuelto todos los atributos (GetAttributes + GetConstantAttributes)
        Public Function Bloque_AtributosDameTodos(queId As Long) As Hashtable
            Dim resultado As New Hashtable
            ''
            If oAppA Is Nothing Then _
            oAppA = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
            ''
            Dim oBl As Autodesk.AutoCAD.Interop.Common.AcadBlockReference = oAppA.ActiveDocument.ObjectIdToObject(queId)
            Dim oAtri As Object ' Autodesk.AutoCAD.Interop.Common.AcadAttribute
            ''
            ' Cargamos los atributos editables
            Dim varAttributes As Object
            varAttributes = oBl.GetAttributes
            ' Localizar los atributos y añadir valores a la colección (ORDEN, NUMERO, etc.). O todos
            For I = LBound(varAttributes) To UBound(varAttributes)
                oAtri = varAttributes(I)
                If resultado.ContainsKey(oAtri.TagString.ToString.ToUpper) = False Then _
                    resultado.Add(oAtri.TagString.ToUpper, oAtri.TextString)
            Next
            ''
            ' Cargamos los atributos constantes
            Dim consAttributes As Object
            consAttributes = oBl.GetConstantAttributes
            ' Localizar los atributos y añadir valores a la colección
            For I = LBound(consAttributes) To UBound(consAttributes)
                oAtri = consAttributes(I)
                If resultado.ContainsKey(oAtri.TagString.toupper) = False Then _
                    resultado.Add(oAtri.TagString.ToUpper, oAtri.TextString)
            Next
            ''
            oAtri = Nothing
            oBl = Nothing
            ''
            VaciaMemoria()
            Return resultado
        End Function
        '
        ''
        '' Devuelto todos los AcadAttributeReference (GetAttributes)
        Public Function Bloque_AtributosDameTodos_AttRef(queId As Long) As Dictionary(Of String, AcadAttributeReference)
            Dim resultado As New Dictionary(Of String, AcadAttributeReference)
            ''
            If oAppA Is Nothing Then _
            oAppA = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
            ''
            Dim oBl As Autodesk.AutoCAD.Interop.Common.AcadBlockReference = oAppA.ActiveDocument.ObjectIdToObject(queId)
            Dim oAtri As Object ' Autodesk.AutoCAD.Interop.Common.AcadAttribute
            ''
            ' Cargamos los atributos editables
            Dim varAttributes As Object
            varAttributes = oBl.GetAttributes
            ' Localizar los atributos y añadir valores a la colección (ORDEN, NUMERO, etc.). O todos
            For I = LBound(varAttributes) To UBound(varAttributes)
                oAtri = varAttributes(I)
                If resultado.ContainsKey(oAtri.TagString.ToString.ToUpper) = False Then _
                    resultado.Add(oAtri.TagString.ToString.ToUpper, CType(oAtri, AcadAttributeReference))
            Next
            ''
            oAtri = Nothing
            oBl = Nothing
            ''
            VaciaMemoria()
            Return resultado
        End Function
        '
        ''
        '' Devuelto todos los AcadAttribute (GetConstantAttributes)
        Public Function Bloque_AtributosDameTodos_AttConst(queId As Long) As Dictionary(Of String, AcadAttribute)
            Dim resultado As New Dictionary(Of String, AcadAttribute)
            ''
            If oAppA Is Nothing Then _
            oAppA = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
            ''
            Dim oBl As Autodesk.AutoCAD.Interop.Common.AcadBlockReference = oAppA.ActiveDocument.ObjectIdToObject(queId)
            Dim oAtri As Object ' Autodesk.AutoCAD.Interop.Common.AcadAttribute
            '
            ' Cargamos los atributos constantes
            Dim consAttributes As Object
            consAttributes = oBl.GetConstantAttributes
            ' Localizar los atributos y añadir valores a la colección
            For I = LBound(consAttributes) To UBound(consAttributes)
                oAtri = consAttributes(I)
                If resultado.ContainsKey(oAtri.TagString.ToString.ToUpper) = False Then _
                    resultado.Add(oAtri.TagString.ToString.ToUpper, CType(oAtri, AcadAttribute))
            Next
            ''
            oAtri = Nothing
            oBl = Nothing
            ''
            VaciaMemoria()
            Return resultado
        End Function
        '
        ' Copia todos los valores de atributos editables de un bloque (idOri) a otro (idDes)
        ' Solo los que se llamen igual (Deberían ser todos si idDes es un Clon de idOri)
        Public Sub Bloque_AtributosCopiaOriDes(idOri As Long, idDes As Long)
            Dim resultado As New Hashtable
            ''
            If oAppA Is Nothing Then _
            oAppA = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
            ''
            Dim oBlOri As Autodesk.AutoCAD.Interop.Common.AcadBlockReference = oAppA.ActiveDocument.ObjectIdToObject(idOri)
            Dim oBlDes As Autodesk.AutoCAD.Interop.Common.AcadBlockReference = oAppA.ActiveDocument.ObjectIdToObject(idDes)
            Dim oAtriO As Object ' Autodesk.AutoCAD.Interop.Common.AcadAttribute     'Object
            Dim oAtriD As Object    ' Autodesk.AutoCAD.Interop.Common.AcadAttribute     'Object
            '
            ' Cargamos los atributos editables
            Dim oAtriOri As Object
            Dim oAtriDes As Object
            oAtriOri = oBlOri.GetAttributes
            oAtriDes = oBlDes.GetAttributes
            ' Localizar los atributos y copiar el valor de oBlOri a oBlDes
            For I = LBound(oAtriOri) To UBound(oAtriOri)
                oAtriO = oAtriOri(I)
                'oAtri.TagString.ToUpper, oAtri.TextString
                For J = LBound(oAtriDes) To UBound(oAtriDes)
                    oAtriD = oAtriDes(J)
                    If oAtriD.TagString.ToString.ToUpper = oAtriO.TagString.ToString.ToUpper Then
                        oAtriD.TextString = oAtriO.TextString
                        Exit For
                    End If
                Next
            Next
            '
            oAtriO = Nothing
            oAtriD = Nothing
            oBlOri = Nothing
            oBlDes = Nothing
            oAtriOri = Nothing
            oAtriDes = Nothing
            ''
            VaciaMemoria()
        End Sub
        Public Sub Bloque_AtributoEscribe(queId As Long, nombreAtri As String, valorAtri As String)
            ''
            If oAppA Is Nothing Then _
            oAppA = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
            ''
            Dim oBl As Autodesk.AutoCAD.Interop.Common.AcadBlockReference = oAppA.ActiveDocument.ObjectIdToObject(queId)

            ' Cargamos los atributos
            Dim varAttributes As Object
            varAttributes = oBl.GetAttributes
            ''
            '' Atributos editables.
            Dim oAtri As Object ' Object ' Autodesk.AutoCAD.Interop.Common.AcadAttribute
            For I = LBound(varAttributes) To UBound(varAttributes)
                oAtri = varAttributes(I)
                If oAtri.TagString.ToString.ToUpper = nombreAtri.ToUpper Then
                    oAtri.TextString = valorAtri
                    Exit For
                End If
            Next
            oBl.Update()
            ''
            oBl = Nothing
            varAttributes = Nothing
            oAtri = Nothing
            ''
            VaciaMemoria()
        End Sub
        Public Sub Bloque_AtributoEscribe(queId As Long, dicNombreValor As Dictionary(Of String, String))
            ''
            If oAppA Is Nothing Then _
            oAppA = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
            ''
            Dim oBl As Autodesk.AutoCAD.Interop.Common.AcadBlockReference = oAppA.ActiveDocument.ObjectIdToObject(queId)

            ' Cargamos los atributos
            Dim varAttributes As Object
            varAttributes = oBl.GetAttributes
            ''
            '' Atributos editables.
            Dim oAtri As Object ' Object ' Autodesk.AutoCAD.Interop.Common.AcadAttribute
            For I = LBound(varAttributes) To UBound(varAttributes)
                oAtri = varAttributes(I)
                If dicNombreValor.ContainsKey(oAtri.TagString.ToString.ToUpper) Then
                    If oAtri.TextString.ToString <> dicNombreValor(oAtri.TagString.ToString.ToUpper) Then
                        oAtri.TextString = dicNombreValor(oAtri.TagString.ToString.ToUpper)
                    End If
                End If
            Next
            oBl.Update()
            ''
            oBl = Nothing
            varAttributes = Nothing
            oAtri = Nothing
            ''
            VaciaMemoria()
        End Sub
        '

        Public Function Bloque_LargoEnX(queId As Long) As Boolean
            Dim resultado As Boolean = False
            '
            If oAppA Is Nothing Then _
            oAppA = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
            '
            Try
                Dim oBl As Autodesk.AutoCAD.Interop.Common.AcadBlockReference = oAppA.ActiveDocument.ObjectIdToObject(queId)

                ' Cargamos los atributos
                Dim minExt As Object = Nothing
                Dim maxExt As Object = Nothing
                Dim largo As Double = 0   ' En X
                Dim ancho As Double = 0   ' En Y
                ' Rellenar coordenadas de minExt(2) y maxExt(2)
                oBl.GetBoundingBox(minExt, maxExt)
                ' Rellenar largo en X y ancho en Y
                largo = maxExt(0) - minExt(0) ' 0 es la coordenada en X
                ancho = maxExt(1) - minExt(1) ' 1 es la coordeanda en Y
                '
                If largo >= ancho Then
                    resultado = True
                Else
                    resultado = False
                End If
                '
                oBl = Nothing
            Catch ex As Exception
                resultado = False
            End Try
            '
            Return resultado
        End Function
        '
        Public Sub Bloque_NUMERO()
            ''
            If oAppA Is Nothing Then _
            oAppA = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
            ''
            Dim contador As Integer = 0
            ''
            For Each oEnt As AcadEntity In oAppA.ActiveDocument.ModelSpace
                If Not (TypeOf oEnt Is AcadBlockReference) Then Continue For
                ''
                Dim oBl As AcadBlockReference = oEnt
                ' Cargamos los atributos
                Dim varAttributes As Object
                varAttributes = oBl.GetAttributes
                ''
                '' Atributos editables.
                Dim oAtri As Object ' Object ' Autodesk.AutoCAD.Interop.Common.AcadAttribute
                For I = LBound(varAttributes) To UBound(varAttributes)
                    oAtri = varAttributes(I)
                    Dim tagS As String = oAtri.TagString
                    If tagS.ToUpper = atributoNUMERO.ToUpper Then
                        contador += 1
                        oAtri.TextString = contador.ToString
                        oBl.Update()
                        Exit For
                    End If
                Next
                oBl = Nothing
                varAttributes = Nothing
                oAtri = Nothing
            Next
            '' Guardar, si tiene cambios.
            If oAppA.ActiveDocument.Saved = False Then oAppA.ActiveDocument.Save()
            ''
            MsgBox("Bloques totales numerados = " & contador)
            VaciaMemoria()
        End Sub
        ''
        Public Sub Bloque_NUMERONuevosSolo()
            ''
            If oAppA Is Nothing Then _
            oAppA = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
            ''
            Dim contador As Integer = 0
            Dim arrBlNumera As New ArrayList
            Dim arrNumeros As New ArrayList
            ''
            For Each oEnt As AcadEntity In oAppA.ActiveDocument.ModelSpace
                If Not (TypeOf oEnt Is AcadBlockReference) Then Continue For
                ''
                Dim oBl As AcadBlockReference = oEnt
                ' Cargamos los atributos
                Dim varAttributes As Object
                varAttributes = oBl.GetAttributes
                ''
                '' Buscamos NUMERO para localizar el más superior.
                '' 
                Dim oAtri As Object ' Object ' Autodesk.AutoCAD.Interop.Common.AcadAttribute
                For I = LBound(varAttributes) To UBound(varAttributes)
                    oAtri = varAttributes(I)
                    Dim tagS As String = oAtri.TagString
                    If tagS.ToUpper = atributoNUMERO.ToUpper Then
                        Dim textS As String = oAtri.TextString
                        If textS = "" Or IsNumeric(oAtri.TextString) = False Or arrNumeros.Contains(oAtri.TextString) Then
                            arrBlNumera.Add(oBl.ObjectID)
                        ElseIf IsNumeric(oAtri.TextString) Then
                            arrNumeros.Add(oAtri.TextString)
                            If CInt(oAtri.TextString) > contador Then contador = CInt(oAtri.TextString)
                        End If
                        ''
                        Exit For
                    End If
                Next
                oBl = Nothing
                varAttributes = Nothing
                oAtri = Nothing
            Next
            ''
            '' Ahora renumeramos sólo los bloques nuevos, empezando en contador +1 (Si es mayor que 0)
            If contador = 0 Then Exit Sub
            contador += 1
            ''
            For Each queId As Long In arrBlNumera
                Bloque_AtributoEscribe(queId, atributoNUMERO, contador)
                contador += 1
            Next
            '' Guardar, si tiene cambios.
            If oAppA.ActiveDocument.Saved = False Then oAppA.ActiveDocument.Save()
            ''
            MsgBox("Bloques nuevos numerados = " & arrBlNumera.Count & vbCrLf &
                   "Ultimo número asignado = " & contador)
        End Sub
        ''
        ''
        '' Escala todos los bloques del dibujo. Solo los del desarrollo actual.
        Public Sub Bloques_EscalaTodos(escala As Double, arrPreBloques As List(Of String))
            If oAppA Is Nothing Then _
            oAppA = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
            ''
            For Each oEnt As AcadEntity In oAppA.ActiveDocument.ModelSpace
                If Not (TypeOf oEnt Is AcadBlockReference) Then Continue For
                ''
                Dim oBl As AcadBlockReference = oEnt
                For Each quePre As String In arrPreBloques
                    If oBl.EffectiveName.ToUpper.StartsWith(quePre.ToUpper) Then
                        'oBl.ScaleEntity(oBl.InsertionPoint, escala)
                        If oBl.XScaleFactor <> escala Then oBl.XScaleFactor = escala
                        If oBl.YScaleFactor <> escala Then oBl.YScaleFactor = escala
                        If oBl.ZScaleFactor <> escala Then oBl.ZScaleFactor = escala
                    End If
                Next
            Next
            oAppA.ActiveDocument.Regen(AcRegenType.acAllViewports)
        End Sub
        ''
        Public Function Bloque_Inserta(queFiBlo As String, queEscala As Double, Optional repetir As Boolean = False) As AcadBlockReference
            Dim oBlRef As AcadBlockReference = Nothing
            If oAppA Is Nothing Then _
            oAppA = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
            ''
            ''
            '' Buscamos este bloque en el dibujo actual. Y, si no está,
            '' lo cargamos del disco duro (queFiBlo = camino completo al fichero DWG)
            Dim nombre As String = IO.Path.GetFileNameWithoutExtension(queFiBlo)
            For Each oBl As AcadBlock In oAppA.ActiveDocument.Blocks
                If oBl.Name = nombre Then
                    queFiBlo = nombre
                    Exit For
                End If
            Next
            ''
repetir:
            Dim resultado As PromptPointResult = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.CurrentDocument.Editor.GetPoint("Punto de insercion :")
            If resultado.Status = PromptStatus.OK Then
                Dim oPoint As Autodesk.AutoCAD.Geometry.Point3d = resultado.Value
                Dim puntoInserta(0 To 2) As Double
                puntoInserta(0) = oPoint.X
                puntoInserta(1) = oPoint.Y
                puntoInserta(2) = 0 'oPoint.Z
                Dim inserta As AcadBlockReference = Nothing
                oBlRef = oAppA.ActiveDocument.ModelSpace.InsertBlock(puntoInserta, queFiBlo, queEscala, queEscala, queEscala, 0)
                oAppA.ActiveDocument.Regen(AcRegenType.acActiveViewport)
            Else
                repetir = False
                oBlRef = Nothing
            End If
            ''
            If repetir = True Then GoTo repetir
            ''
            Return oBlRef
        End Function
        ''
        Public Sub Bloque_Cambia(queFiBlo As String, queArrBlo As ArrayList, queEscala As Double)
            Dim oBlRef As AcadBlockReference = Nothing
            If oAppA Is Nothing Then _
            oAppA = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
            ''
            ''
            '' Buscamos este bloque en el dibujo actual. Y, si no está,
            '' lo cargamos del disco duro (queFiBlo = camino completo al fichero DWG)
            Dim nombre As String = IO.Path.GetFileNameWithoutExtension(queFiBlo)
            For Each oBl As AcadBlock In oAppA.ActiveDocument.Blocks
                If oBl.Name = nombre Then
                    queFiBlo = nombre
                    Exit For
                End If
            Next
            ''
            Dim oApp As Autodesk.AutoCAD.Interop.AcadApplication =
                CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication,
    Autodesk.AutoCAD.Interop.AcadApplication)
            'Me.Visible = False
            '' Iteramos con el ArrayList de bloques insertados
            For Each oBlr As AcadBlockReference In queArrBlo
                Dim puntoInserta(0 To 2) As Double
                puntoInserta(0) = oBlr.InsertionPoint(0)
                puntoInserta(1) = oBlr.InsertionPoint(1)
                puntoInserta(2) = 0
                Dim inserta As AcadBlockReference = Nothing
                inserta = oAppA.ActiveDocument.ActiveLayout.Block.InsertBlock(puntoInserta, queFiBlo, queEscala, queEscala, queEscala, 0)
                oBlr.Delete()
            Next
            oAppA.ActiveDocument.Regen(AcRegenType.acActiveViewport)
        End Sub
        ''
        Public Sub Bloque_CambiaTodosSubDir(calidadFin As String, colCalidades As List(Of String), datoPath As String, queArrBlo As ArrayList, Optional nombreparametro As String = "CALIDAD")
            Dim oBlRef As AcadBlockReference = Nothing
            If oAppA Is Nothing Then _
            oAppA = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
            ''
            If queArrBlo Is Nothing Then Exit Sub
            '' Iteramos con el ArrayList de bloques insertados
            For Each oBlr As AcadBlockReference In queArrBlo
                'Cogemos el FullPath del bloque insertado
                Dim fullPath As String = XLeeDato(oBlr.Handle, datoPath)
                If fullPath = "" Or IO.File.Exists(fullPath) = False Then Exit Sub
                'Cogemos el nombre
                Dim nombre As String = IO.Path.GetFileName(fullPath)
                'Cogemos el fullpath del directorio
                Dim queDir As String = IO.Path.GetDirectoryName(fullPath)
                If queDir = "" Or IO.Directory.Exists(queDir) = False Then Exit Sub
                'Cogemos sólo el nombre del directorio
                Dim nDir As String = IO.Path.GetFileName(queDir)
                Dim subDir As String = IO.Path.GetDirectoryName(queDir)
                '
                ' Si no es un bloque que este en directorio de colCalidades,
                ' evaluar estado de visibilidad y si tampoco tiene CALIDAD, continuar.
                If colCalidades.Contains(nDir) = False Then
                    ' Evaluar estado de visibilidad y si tampoco tiene CALIDAD, continuar.
                    Call BloqueDinamico_ParametroEscribe(oBlr.ObjectID, nombreparametro, calidadFin)
                    Continue For
                End If
                'Si ya está en el directorio de la calidadFin elegida, continuar. Si no, cambiar.
                If nDir = calidadFin Then Continue For
                '
                ' Nuevo directorio desde el que coger el elemento a cambiar.
                Dim nuevoFullPath = IO.Path.Combine(subDir, calidadFin, nombre)
                '
                Dim puntoInserta(0 To 2) As Double
                puntoInserta(0) = oBlr.InsertionPoint(0)
                puntoInserta(1) = oBlr.InsertionPoint(1)
                puntoInserta(2) = 0
                Dim escX As Double = oBlr.XScaleFactor
                Dim escY As Double = oBlr.YScaleFactor
                Dim escZ As Double = oBlr.ZScaleFactor
                Dim rotacion As Double = oBlr.Rotation
                Dim inserta As AcadBlockReference = Nothing
                inserta = oAppA.ActiveDocument.ActiveLayout.Block.InsertBlock(puntoInserta, nuevoFullPath, escX, escY, escZ, rotacion)
                oBlr.Delete()
                If inserta IsNot Nothing Then
                    XPonDato(inserta.Handle, datoPath, nuevoFullPath)
                End If
            Next
            oAppA.ActiveDocument.Regen(AcRegenType.acActiveViewport)
        End Sub
        Public Function BloqueDinamico_ParametroDame(queid As Long, quePar As String) As String
            Dim resultado As String = ""
            ''
            If oAppA Is Nothing Then _
            oAppA = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
            ''
            Dim oBlo As Autodesk.AutoCAD.Interop.Common.AcadBlockReference = oAppA.ActiveDocument.ObjectIdToObject(queid)
            ''
            '' Si el bloquedinamico
            If oBlo.IsDynamicBlock Then
                'Debug.Print oBlo.Name & "( " & oBlo.EffectiveName & " )"
                Dim oPs As Object = oBlo.GetDynamicBlockProperties
                Dim oDp As AcadDynamicBlockReferenceProperty
                For x As Integer = 0 To UBound(oPs)
                    oDp = oPs(x)
                    If oDp.PropertyName.ToUpper = quePar.ToUpper Then
                        resultado = oDp.Value
                        Exit For
                    End If
                Next
                oDp = Nothing
                oPs = Nothing
            End If
            oBlo = Nothing
            '
            Return resultado
        End Function
        '

        Public Function BloqueDinamico_ParametroDame(arrParameters As Object, quePar As String) As String
            Dim resultado As String = ""
            Dim oDp As AcadDynamicBlockReferenceProperty
            For x As Integer = 0 To UBound(arrParameters)
                oDp = arrParameters(x)
                If oDp.PropertyName.ToUpper = quePar.ToUpper Then
                    resultado = oDp.Value.ToString
                    Exit For
                End If
            Next
            oDp = Nothing
            '
            Return resultado
        End Function
        Public Function BloqueDinamico_ParametroEscribe(queid As Long, quePar As String, queVal As Object) As Boolean
            Dim resultado As Boolean = False
            ''
            If oAppA Is Nothing Then _
            oAppA = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
            ''
            Dim oBlo As Autodesk.AutoCAD.Interop.Common.AcadBlockReference = oAppA.ActiveDocument.ObjectIdToObject(queid)
            ''
            '' Si el bloquedinamico
            If oBlo.IsDynamicBlock Then
                'Debug.Print oBlo.Name & "( " & oBlo.EffectiveName & " )"
                Dim oPs As Object = oBlo.GetDynamicBlockProperties
                Dim oDp As AcadDynamicBlockReferenceProperty
                For x As Integer = 0 To UBound(oPs)
                    oDp = oPs(x)
                    If oDp.PropertyName.ToUpper = quePar.ToUpper Then
                        oDp.Value = queVal
                        oBlo.Update()
                        resultado = True
                        Exit For
                    End If
                Next
                oDp = Nothing
                oPs = Nothing
            Else
                resultado = False
            End If
            oBlo = Nothing
            '
            Return resultado
        End Function
        Public Function BloqueDinamicoParametrosDameTodos(queid As Long) As Hashtable
            Dim resultado As New Hashtable
            ''
            If oAppA Is Nothing Then _
            oAppA = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
            ''
            Dim oBlo As Autodesk.AutoCAD.Interop.Common.AcadBlockReference = oAppA.ActiveDocument.ObjectIdToObject(queid)
            ''
            '' Si es bloquedinamico
            If oBlo.IsDynamicBlock Then
                'Debug.Print oBlo.Name & "( " & oBlo.EffectiveName & " )"
                Dim oPs As Object = oBlo.GetDynamicBlockProperties
                Dim oDp As AcadDynamicBlockReferenceProperty
                For x As Integer = 0 To UBound(oPs)
                    oDp = oPs(x)
                    'Debug.Print vbTab & (oDp.PropertyName)
                    If resultado.ContainsKey(oDp.PropertyName) = False And
                    resultado.ContainsKey(oDp.PropertyName.ToUpper) = False And
                    resultado.ContainsKey(oDp.PropertyName.ToLower) = False Then
                        Try
                            resultado.Add(oDp.PropertyName.ToUpper, oDp.Value.ToString.ToUpper)
                        Catch ex As Exception
                            Debug.Print(ex.ToString)
                        End Try
                    End If
                Next
                oDp = Nothing
                oPs = Nothing
            End If
            oBlo = Nothing
            ''
            Return resultado
        End Function
        Public Function BloqueDinamicoParametrosDameTodos_Dictionary(queid As Long) As Dictionary(Of String, AcadDynamicBlockReferenceProperty)
            Dim resultado As New Dictionary(Of String, AcadDynamicBlockReferenceProperty)
            ''
            If oAppA Is Nothing Then _
            oAppA = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
            ''
            Dim oBlo As Autodesk.AutoCAD.Interop.Common.AcadBlockReference = oAppA.ActiveDocument.ObjectIdToObject(queid)
            ''
            '' Si es bloquedinamico
            If oBlo.IsDynamicBlock Then
                Dim oPs As Object = oBlo.GetDynamicBlockProperties
                Dim oDp As AcadDynamicBlockReferenceProperty
                For x As Integer = 0 To UBound(oPs)
                    oDp = oPs(x)
                    'Debug.Print vbTab & (oDp.PropertyName)
                    If resultado.ContainsKey(oDp.PropertyName) = False And
                    resultado.ContainsKey(oDp.PropertyName.ToUpper) = False And
                    resultado.ContainsKey(oDp.PropertyName.ToLower) = False Then
                        Try
                            resultado.Add(oDp.PropertyName.ToUpper, oDp)
                        Catch ex As Exception
                            Debug.Print(ex.ToString)
                        End Try
                    End If
                Next
                oDp = Nothing
                oPs = Nothing
            End If
            oBlo = Nothing
            ''
            Return resultado
        End Function
        Public Function Bloque_EsregApp(ByVal oBl As AcadBlockReference) As Boolean
            Dim arr() As Object = oBl.GetConstantAttributes
            Dim resultado As Boolean = False
            Try
                ' Buscar en atributos CONTROL=regApp
                If arr.Length = 0 Then
                    resultado = False
                Else
                    Dim acAt As AcadAttribute
                    For Each obj As Object In arr
                        acAt = CType(obj, AcadAttribute)
                        If acAt.TagString.ToString = "CONTROL" And acAt.TextString.ToString = regAPPA Then
                            resultado = True
                            Exit For
                        End If
                    Next
                End If
            Catch ex As Exception
                resultado = False
            End Try
            '
            ' Leer XData para ver si está registrada regAPP
            If resultado = False Then
                Try
                    If XLeeDato(oBl.Handle, regAPPA) = regAPPA Then resultado = True
                Catch
                End Try
            End If
            '
            Return resultado
        End Function
        ''' <summary>
        ''' Procedimiento que detecta si un bloque está dentro de una polilinea (Region)
        ''' </summary>
        ''' <param name="oBlRef">BlockReference que vamos a evaluar (Su punto de insercion)</param>
        ''' <param name="oPol">Polyline delimitadora (Tiene que estar cerrada)</param>
        ''' <returns></returns>
        Public Function Bloque_EstaDentroPolilinea(ByRef oBlRef As BlockReference, ByRef oPol As Polyline) As Boolean
            Dim OkRetour As Boolean = False
            '===========================================================================
            'Gestión de errores. Salir si se cumplen. (Objetos Nothing o polilinea abierta)
            If IsNothing(oBlRef) = True Or IsNothing(oPol) = True Then Return OkRetour
            'Si polilinea no esta cerrada
            If oPol.Closed = False Then Return False
            'Punto de inserción del bloque
            Dim pt As New Autodesk.AutoCAD.Geometry.Point2d(oBlRef.Position.X, oBlRef.Position.Y)
            'Si el punto de inserción está dentro de la polilinea
            Dim vn As Integer = oPol.NumberOfVertices
            Dim colpt() As Autodesk.AutoCAD.Geometry.Point2d = Nothing
            ReDim colpt(vn)
            For i As Integer = 0 To vn - 1
                'Convertir los puntos 3D a puntos 2D
                Dim pts As Autodesk.AutoCAD.Geometry.Point2d = oPol.GetPoint2dAt(i)
                colpt(i) = New Autodesk.AutoCAD.Geometry.Point2d(pts.X, pts.Y)
                'Comprobar si el punto de inserción está encima de la polilinea.
                Dim seg As Autodesk.AutoCAD.Geometry.Curve2d = Nothing
                Dim segType As SegmentType = oPol.GetSegmentType(i)
                If (segType = SegmentType.Arc) Then
                    seg = oPol.GetArcSegment2dAt(i)
                ElseIf (segType = SegmentType.Line) Then
                    seg = oPol.GetLineSegment2dAt(i)
                End If
                If IsNothing(seg) = False Then
                    OkRetour = seg.IsOn(pt)
                    If OkRetour = True Then
                        Return True
                    End If
                End If
            Next
            'Hacer coincidir punto final e inicial de la polilinea
            colpt(vn) = oPol.GetPoint2dAt(0)
            '===========================================================================
            Dim RetFonction As Double
            RetFonction = wn_PnPoly(pt, colpt, vn)
            If RetFonction = 0 Then
                OkRetour = False
            Else
                OkRetour = True
            End If
            Return OkRetour
        End Function

        'AFLETA
        Public Function Bloques_DameTodosPorAtributo(queatributo As String, quevalor As String, Optional exacto As Boolean = False) As ArrayList
            Dim resultado As New ArrayList
            If oAppA Is Nothing Then _
            oAppA = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
            ''
            For Each oEnt As AcadEntity In oAppA.ActiveDocument.ModelSpace
                If TypeOf oEnt Is AcadBlockReference Then
                    Dim oBlk As AcadBlockReference = oAppA.ActiveDocument.ObjectIdToObject(oEnt.ObjectID)

                    Dim strValor As String = Bloque_AtributoDame(oBlk.ObjectID, queatributo)

                    If exacto = True Then
                        If Bloque_AtributoDame(oBlk.ObjectID, queatributo).ToUpper = quevalor.ToUpper Then
                            resultado.Add(oBlk)
                        End If
                    Else
                        If Bloque_AtributoDame(oBlk.ObjectID, queatributo).ToUpper.Contains(quevalor.ToUpper) Then
                            resultado.Add(oBlk)
                        End If
                    End If
                    oBlk = Nothing
                End If
            Next
            VaciaMemoria()
            Return resultado
        End Function
#End Region
        '
#Region "ATRIBUTOS"
        Public Function AtributosConstant_DameTodos(ByVal bl As AcadBlockReference) As String
            Dim att As AcadAttribute
            Dim resultado As String = ""
            Dim atts() As Object
            atts = bl.GetConstantAttributes

            If atts.Length > 0 Then
                For Each atrr As Object In atts
                    att = CType(atrr, AcadAttribute)
                    resultado &= (att.TagString & ": " & att.TextString) & vbCrLf
                Next
            End If
            '
            AtributosConstant_DameTodos = resultado
        End Function
#End Region
    End Class

    Public Class Bloque_Datos
        Public handle As String = ""
        Public id As Long = 0
        Public name As String = ""
        Public eName As String = ""
        Public atributos As Object = Nothing
        Public atributosConst As Object = Nothing
        Public propiedades As Object = Nothing

        Public Sub New(oBl As AcadBlockReference)
            handle = oBl.Handle
            id = oBl.ObjectID
            name = oBl.Name
            eName = oBl.EffectiveName
            atributos = oBl.GetAttributes
            atributosConst = oBl.GetConstantAttributes
            If oBl.IsDynamicBlock Then
                propiedades = oBl.GetDynamicBlockProperties
            End If
        End Sub
    End Class
End Namespace
