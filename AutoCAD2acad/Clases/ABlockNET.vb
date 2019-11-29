Imports System.Diagnostics
Imports System.Collections
Imports System.Windows.Forms
Imports System.Drawing


Imports Autodesk.AutoCAD.Interop
Imports Autodesk.AutoCAD.Interop.Common
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.DatabaseServices.Filters
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.EditorInput

Namespace A2acad
    Partial Public Class A2acad

#Region "BLOQUESACTIVEX"
        Private Shared acBlkRefCached As BlockReference = Nothing
        Private Shared acIdBlkRefCached As ObjectId = ObjectId.Null
        Private Shared acAttRefCached As AttributeReference = Nothing
        Private Shared acIdAttRefCached As ObjectId = ObjectId.Null

        Public Function BlockReference_SelectOne(ByVal pModo As OpenMode) As BlockReference
            Dim doc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim db As Database = doc.Database
            Dim ed As Editor = doc.Editor
            Dim acBlkRef As BlockReference
            Try

                Using acTrans As Transaction = db.TransactionManager.StartOpenCloseTransaction()
                    Dim acBlkTbl As BlockTable = CType(acTrans.GetObject(db.BlockTableId, OpenMode.ForRead), BlockTable)
                    Dim opts As PromptEntityOptions = New PromptEntityOptions(vbLf & "Select BlockReference: ")
                    opts.SetRejectMessage("Only one BlockReference.")
                    opts.AddAllowedClass(GetType(BlockReference), True)
                    Dim selRes As PromptEntityResult = ed.GetEntity(opts)
                    If selRes.Status <> PromptStatus.OK Then
                        ed.WriteMessage(vbLf & "Selected BlockReference FAILED.")
                        Return Nothing
                    End If

                    acBlkRef = CType(acTrans.GetObject(selRes.ObjectId, pModo), BlockReference)
                End Using

                ed.WriteMessage(vbLf & "Selected BlockReference CORRECT.")
                '
                VaciaMemoria()
                Return acBlkRef
            Catch ex As Autodesk.AutoCAD.Runtime.Exception
                ed.WriteMessage(ex.Message.ToString())
                Return Nothing
            End Try
        End Function
        Public Function BlockReference_SelectMultiples(ByVal pListFilter As TypedValue(), ByVal pModo As OpenMode) As List(Of BlockReference)
            Dim doc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim db As Database = doc.Database
            Dim ed As Editor = doc.Editor
            Dim listResult As List(Of BlockReference) = New List(Of BlockReference)()
            Try
                Using acTrans As Transaction = db.TransactionManager.StartOpenCloseTransaction()
                    Dim acBlkTbl As BlockTable = CType(acTrans.GetObject(db.BlockTableId, OpenMode.ForRead), BlockTable)
                    Dim filter As SelectionFilter = New SelectionFilter(pListFilter)
                    Dim opts As PromptSelectionOptions = New PromptSelectionOptions()
                    opts.MessageForAdding = "Select BlockReference: "
                    Dim selRes As PromptSelectionResult = ed.GetSelection(opts, filter)
                    If selRes.Status <> PromptStatus.OK Then
                        ed.WriteMessage(vbLf & "Selected BlockReferences FAILED.")
                        Return listResult
                    End If

                    If selRes.Value.Count <> 0 Then
                        Dim [set] As SelectionSet = selRes.Value
                        For Each id As ObjectId In [set].GetObjectIds()
                            Dim acBlkRef As BlockReference = CType(acTrans.GetObject(id, pModo), BlockReference)
                            listResult.Add(acBlkRef)
                        Next

                        ed.WriteMessage(vbLf & "Selected BlockReferences CORRECT.")
                    End If
                End Using
                '
                VaciaMemoria()
                Return listResult
            Catch ex As Autodesk.AutoCAD.Runtime.Exception
                ed.WriteMessage(ex.Message.ToString())
                Return listResult
            End Try
        End Function

        'Public Function FindBloques(ByVal pListFilterSelection As TypedValue(), ByVal pPropiedades As Dictionary(Of String, Object)) As List(Of BlockReference)
        '    Dim doc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        '    Dim db As Database = doc.Database
        '    Dim ed As Editor = doc.Editor
        '    Dim listResult As List(Of BlockReference) = New List(Of BlockReference)()
        '    Try
        '        Using acTrans As Transaction = db.TransactionManager.StartOpenCloseTransaction()
        '            Dim acBlkTbl As BlockTable = CType(acTrans.GetObject(db.BlockTableId, OpenMode.ForRead), BlockTable)
        '            Dim filter As SelectionFilter = New SelectionFilter(pListFilterSelection)
        '            Dim selRes As PromptSelectionResult = ed.SelectAll(filter)
        '            If selRes.Status <> PromptStatus.OK Then
        '                ed.WriteMessage(vbLf & "No ha seleccionado un bloque correcto.")
        '                Return listResult
        '            End If

        '            If selRes.Value.Count <> 0 Then
        '                Dim [set] As SelectionSet = selRes.Value
        '                listResult = [set].GetObjectIds().[Select](Function(i) OpenBlockReference(acTrans, i)).ToList(Of BlockReference).Where(Function(b) b.AttributeCollection.Cast(Of ObjectId)().[Select](Function(IdA) OpenAttributeReference(acTrans, IdA)).ToList(Of AttributeReference)().Where(Function(a)
        '                                                                                                                                                                                                                                                                                                    If pPropiedades.ContainsKey(a.Tag) AndAlso pPropiedades(a.Tag).ToString() = a.TextString Then
        '                                                                                                                                                                                                                                                                                                        Return True
        '                                                                                                                                                                                                                                                                                                    Else
        '                                                                                                                                                                                                                                                                                                        Return False
        '                                                                                                                                                                                                                                                                                                    End If
        '                                                                                                                                                                                                                                                                                                End Function).Count() = pPropiedades.Count).ToList(Of BlockReference)()
        '                ed.WriteMessage(vbLf & "Selección completada correctamente.")
        '            End If
        '        End Using
        '        'Vacia()
        '        Return listResult
        '    Catch ex As Autodesk.AutoCAD.Runtime.Exception
        '        ed.WriteMessage(ex.Message.ToString())
        '        Return listResult
        '    End Try
        'End Function

        'Public Function Count_Bloques(ByVal pListFilterSelection As TypedValue(), ByVal pPropiedades As Dictionary(Of String, Object)) As Integer
        '    Dim doc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        '    Dim db As Database = doc.Database
        '    Dim ed As Editor = doc.Editor
        '    Dim intTotal As Integer = 0
        '    Try
        '        Using acTrans As Transaction = db.TransactionManager.StartOpenCloseTransaction()
        '            Dim acBlkTbl As BlockTable = CType(acTrans.GetObject(db.BlockTableId, OpenMode.ForRead), BlockTable)
        '            Dim filter As SelectionFilter = New SelectionFilter(pListFilterSelection)
        '            Dim selRes As PromptSelectionResult = ed.SelectAll(filter)
        '            If selRes.Status <> PromptStatus.OK Then
        '                ed.WriteMessage(vbLf & "No ha seleccionado un bloque correcto.")
        '                Return 0
        '            End If

        '            If selRes.Value.Count <> 0 Then
        '                Dim [set] As SelectionSet = selRes.Value
        '                intTotal = [set].GetObjectIds().[Select](Function(i) OpenBlockReference(acTrans, i)).ToList(Of BlockReference)().Where(Function(b) b.AttributeCollection.Cast(Of ObjectId)().[Select](Function(IdA) OpenAttributeReference(acTrans, IdA)).ToList(Of AttributeReference)().Where(Function(a)
        '                                                                                                                                                                                                                                                                                                    If pPropiedades.ContainsKey(a.Tag) AndAlso pPropiedades(a.Tag).ToString() = a.TextString Then
        '                                                                                                                                                                                                                                                                                                        Return True
        '                                                                                                                                                                                                                                                                                                    Else
        '                                                                                                                                                                                                                                                                                                        Return False
        '                                                                                                                                                                                                                                                                                                    End If
        '                                                                                                                                                                                                                                                                                                End Function).Count() = pPropiedades.Count).Count()
        '                ed.WriteMessage(vbLf & "Selección completada correctamente.")
        '            End If
        '        End Using
        '        'Vacia()
        '        Return intTotal
        '    Catch ex As Autodesk.AutoCAD.Runtime.Exception
        '        ed.WriteMessage(ex.Message.ToString())
        '        Return intTotal
        '    End Try
        'End Function
        Private Function BlockReference_Open(ByVal pAcTrans As Transaction, ByVal pObjId As ObjectId) As BlockReference
            If acIdBlkRefCached <> pObjId OrElse acBlkRefCached Is Nothing Then
                acBlkRefCached = CType(pAcTrans.GetObject(pObjId, OpenMode.ForRead), BlockReference)
                acIdBlkRefCached = acBlkRefCached.Id
            End If

            VaciaMemoria()
            Return acBlkRefCached
        End Function
        Public Function BlockTableRecord_Clone(ByVal idBlTR As ObjectId, ByVal newName As String, ByVal BlTROrigin As Point3d) As BlockTableRecord
            Dim doc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim db As Database = doc.Database
            Dim acBlkTblRDestino As BlockTableRecord
            Try
                Using acLckDoc As DocumentLock = doc.LockDocument()
                    Using acTrans As Transaction = db.TransactionManager.StartOpenCloseTransaction()
                        Dim copyId As ObjectId = ObjectId.Null
                        Using temDB As Database = db.Wblock(idBlTR)
                            copyId = db.Insert(newName, temDB, True)
                        End Using

                        If copyId <> ObjectId.Null Then
                            acBlkTblRDestino = TryCast(acTrans.GetObject(copyId, OpenMode.ForWrite), BlockTableRecord)
                            'acBlkTblRDestino.Origin = BlTROrigin
                        Else
                            acBlkTblRDestino = Nothing
                        End If

                        acTrans.Commit()
                    End Using
                End Using
                '
                VaciaMemoria()
                Return acBlkTblRDestino
            Catch ex As Autodesk.AutoCAD.Runtime.Exception
                MessageBox.Show(ex.Message)
                Return Nothing
            End Try
        End Function

        Public Function BlockTableRecord_Clone(ByVal idTableRecord As ObjectId, newName As String, ByVal newPoint3D As Point3d?) As ObjectId
            Dim doc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim db As Database = doc.Database
            Dim copyId As ObjectId = ObjectId.Null
            Try
                Using acLckDoc As DocumentLock = doc.LockDocument()
                    Using acTrans As Transaction = db.TransactionManager.StartOpenCloseTransaction()
                        Dim oBlTRorigen As BlockTableRecord = CType(acTrans.GetObject(idTableRecord, OpenMode.ForRead), BlockTableRecord)
                        Dim strNombreBloqueDestino As String = BlockTableRecord_NewName(idTableRecord)
                        Dim temDB As Database = db.Wblock(idTableRecord)
                        copyId = db.Insert(strNombreBloqueDestino, temDB, True)
                        If newPoint3D IsNot Nothing Then
                            Dim acBlkTblRDestino As BlockTableRecord = TryCast(acTrans.GetObject(copyId, OpenMode.ForWrite), BlockTableRecord)
                            acBlkTblRDestino.Origin = newPoint3D.Value
                        End If

                        acTrans.Commit()
                    End Using
                End Using
                VaciaMemoria()
            Catch ex As Autodesk.AutoCAD.Runtime.Exception
                MessageBox.Show(ex.Message)
            End Try

            Return copyId
        End Function

        Public Function BlockTableRecord_Clone(ByVal pNombreOriginal As String, ByVal pNombreDestino As String, ByVal pPoint3D As Point3d) As BlockTableRecord
            Dim doc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim db As Database = doc.Database
            Dim acBlkTblRDestino As BlockTableRecord
            Try
                Using acLckDoc As DocumentLock = doc.LockDocument()
                    Using acTrans As Transaction = db.TransactionManager.StartOpenCloseTransaction()
                        Dim acBlkTbl As BlockTable = TryCast(acTrans.GetObject(db.BlockTableId, OpenMode.ForRead), BlockTable)
                        Dim acBlkTblROriginal As BlockTableRecord = TryCast(acTrans.GetObject(acBlkTbl(pNombreOriginal), OpenMode.ForRead), BlockTableRecord)
                        Dim copyId As ObjectId = ObjectId.Null
                        Using temDB As Database = db.Wblock(acBlkTblROriginal.Id)
                            copyId = db.Insert(pNombreDestino, temDB, True)
                        End Using

                        If copyId <> ObjectId.Null Then
                            acBlkTblRDestino = TryCast(acTrans.GetObject(copyId, OpenMode.ForWrite), BlockTableRecord)
                            acBlkTblRDestino.Origin = pPoint3D
                        Else
                            acBlkTblRDestino = Nothing
                        End If

                        acTrans.Commit()
                    End Using
                End Using
                '
                VaciaMemoria()
                Return acBlkTblRDestino
            Catch ex As Autodesk.AutoCAD.Runtime.Exception
                MessageBox.Show(ex.Message)
                Return Nothing
            End Try
        End Function

        Public Function BlockTableRecord_CambiaPuntoInsercion(ByVal idBlTR As ObjectId, idBl As ObjectId, oldPoint3D As Point3d, ByVal newPoint3D As Point3d) As BlockTableRecord
            Dim doc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim ed As Editor = doc.Editor
            Dim db As Database = doc.Database
            Dim ucs As Matrix3d = ed.CurrentUserCoordinateSystem        '' Actual UCS
            Dim insPt As Point3d = newPoint3D.TransformBy(ucs)          '' Punto en coordenadas UCS
            Dim newPt As Point3d = Point3d.Origin                       '' Nuevo punto en coordenadas 0,0,0
            Dim hasAtts As Boolean = False                              '' Tiene atributos el bloque
            Dim btr As BlockTableRecord = Nothing

            Using acLckDoc As DocumentLock = doc.LockDocument()
                Using tr As Transaction = db.TransactionManager.StartTransaction()
                    Dim br As BlockReference = TryCast(tr.GetObject(idBl, OpenMode.ForWrite), BlockReference)
                    btr = TryCast(tr.GetObject(idBlTR, OpenMode.ForWrite), BlockTableRecord)
                    ' Transformaciones del bloque insertado
                    Dim blockmatrix As Matrix3d = br.BlockTransform
                    ' Punto seleccionado en coordenadas del bloque
                    insPt = insPt.TransformBy(blockmatrix.Inverse())
                    ' Coordenadas para mover los objetos que contiene el BlockTableRecord
                    Dim move As Matrix3d = Matrix3d.Displacement(btr.Origin - insPt)
                    For Each id As ObjectId In btr
                        Dim ent As Entity = TryCast(tr.GetObject(id, OpenMode.ForWrite), Entity)
                        ent.TransformBy(move)
                        ' Si tienen atributos
                        If ent.[GetType]() = GetType(AttributeDefinition) Then hasAtts = True
                    Next
                    ' Coleccion de todos los BlockReferente que tiene el BlockTableRecord
                    Dim ids As ObjectIdCollection = btr.GetBlockReferenceIds(False, True)
                    ' Nueva Matrix3D para mover cada BlockReference
                    Dim mBref As Matrix3d = New Matrix3d()
                    ' Mover cada Blockreference a su nueva ubicación.
                    For Each brefId As ObjectId In ids
                        Dim btrans As Boolean = False
                        Dim b As BlockReference = TryCast(tr.GetObject(brefId, OpenMode.ForWrite), BlockReference)
                        ' Posición actual del Blockreference
                        Dim bposition As Point3d = b.Position
                        If b.BlockTableRecord = idBlTR Then
                            ' Mover los Blockrefence que cuelga directamente del BlockTableRecord
                            newPt = insPt.TransformBy(b.BlockTransform)
                            mBref = Matrix3d.Displacement(newPt - b.Position)
                            b.TransformBy(mBref)
                            btrans = True
                        Else
                            ' Mover los BlockReference que cuelga de algún hijo (Recorrer sus BlockReference hijos)
                            Dim nestedBref As BlockTableRecord = TryCast(tr.GetObject(b.BlockTableRecord, OpenMode.ForWrite), BlockTableRecord)
                            For Each id As ObjectId In nestedBref
                                If id = brefId Then
                                    newPt = insPt.TransformBy(b.BlockTransform)
                                    b.TransformBy(mBref)
                                    btrans = True
                                    Exit For
                                End If
                            Next
                        End If
                        ' Si hay que moverlo btrans = True
                        If btrans Then
                            Dim id As ObjectId = b.ExtensionDictionary
                            If id = ObjectId.Null Then Continue For
                            '
                            Dim dic As DBDictionary = TryCast(tr.GetObject(id, OpenMode.ForWrite), DBDictionary)
                            Const spatialName As String = "SPATIAL"
                            If Not dic.Contains("ACAD_FILTER") Then Continue For
                            '
                            Dim FilterDic As DBDictionary = TryCast(tr.GetObject(dic.GetAt("ACAD_FILTER"), OpenMode.ForWrite), DBDictionary)
                            If Not FilterDic.Contains(spatialName) Then Continue For
                            '
                            Dim sf As Autodesk.AutoCAD.DatabaseServices.Filters.SpatialFilter = TryCast(tr.GetObject(FilterDic.GetAt(spatialName), OpenMode.ForWrite), SpatialFilter)
                            Dim def As SpatialFilterDefinition = sf.Definition
                            Dim pts As Point2dCollection = New Point2dCollection()
                            Dim m As Matrix3d = sf.OriginalInverseBlockTransform
                            Dim cs As CoordinateSystem3d = m.CoordinateSystem3d
                            Dim plane As Plane = New Plane(cs.Origin, cs.Zaxis)
                            Dim disp As Matrix2d = Matrix2d.Displacement(newPt.Convert2d(plane) - b.Position.Convert2d(plane))
                            For Each pt As Point2d In def.GetPoints()
                                Dim l As Point3d = New Point3d(plane, pt)
                                l = l.TransformBy(move)
                                l = l.TransformBy(m)
                                pts.Add(l.Convert2d(plane))
                            Next

                            pts.TrimToSize()
                            Dim NewDef As SpatialFilterDefinition = New SpatialFilterDefinition(pts, Vector3d.ZAxis, 0, def.FrontClip, def.BackClip, True)
                            Dim NewSf As SpatialFilter = New SpatialFilter()
                            NewSf.Definition = NewDef
                            FilterDic.Remove(spatialName)
                            FilterDic.SetAt(spatialName, NewSf)
                            tr.AddNewlyCreatedDBObject(NewSf, True)
                        End If
                        ' Si tiene atributos, los actualizamos.
                        If hasAtts Then
                            AttributeReference_Update(idBl, db, tr)
                        End If
                    Next
                    tr.Commit()
                End Using
            End Using
            '
            VaciaMemoria()
            ed.Regen()
            Return btr
        End Function

        Public Function BlockTableRecord_CloneTransform(ByVal idTableRecord As ObjectId, ByVal newName As String, oldPoint3D As Point3d, ByVal newPoint3D As Point3d) As BlockTableRecord
            Dim doc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim db As Database = doc.Database
            Dim acBlkTblRDestino As BlockTableRecord
            Try
                Using acLckDoc As DocumentLock = doc.LockDocument()
                    Using acTrans As Transaction = db.TransactionManager.StartOpenCloseTransaction()
                        Dim copyId As ObjectId = ObjectId.Null
                        Using temDB As Database = db.Wblock(idTableRecord)
                            copyId = db.Insert(newName, temDB, True)
                        End Using

                        If copyId <> ObjectId.Null Then
                            acBlkTblRDestino = TryCast(acTrans.GetObject(copyId, OpenMode.ForWrite), BlockTableRecord)
                            'Dim ptini As Point3d = acBlkTblRDestino.Origin
                            'Dim ptfin As Point3d = newPoint3D
                            acBlkTblRDestino.Origin = newPoint3D
                            ' Mover todos los objetos
                            '' Create a matrix and move from ptini to ptfin
                            'Dim acVec3d As Vector3d = oldPoint3D.GetVectorTo(newPoint3D)
                            'For Each oId As ObjectId In acBlkTblRDestino
                            '    Dim oEnt As Entity = TryCast(acTrans.GetObject(oId, OpenMode.ForWrite), Entity)
                            '    oEnt.TransformBy(Matrix3d.Displacement(acVec3d))
                            '    acTrans.AddNewlyCreatedDBObject(oEnt, True)
                            'Next
                            '' Create a matrix and move the circle using a vector from (0,0,0) to (2,0,0)
                            ''Dim acPt3d As Point3d = New Point3d(0, 0, 0)
                            ''Dim acVec3d As Vector3d = acPt3d.GetVectorTo(New Point3d(2, 0, 0))
                            ''acCirc.TransformBy(Matrix3d.Displacement(acVec3d))
                        Else
                            acBlkTblRDestino = Nothing
                        End If

                        acTrans.Commit()
                    End Using
                End Using
                '
                VaciaMemoria()
                Return acBlkTblRDestino
            Catch ex As Autodesk.AutoCAD.Runtime.Exception
                MessageBox.Show(ex.Message)
                Return Nothing
            End Try
        End Function

        Public Function BlockTableRecord_CloneTransform(ByVal idTableRecord As ObjectId, idOld As ObjectId, ByVal newPoint3D As Point3d?) As ObjectId
            Dim doc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim db As Database = doc.Database
            Dim copyId As ObjectId = ObjectId.Null
            Try
                Using acLckDoc As DocumentLock = doc.LockDocument()
                    Using acTrans As Transaction = db.TransactionManager.StartOpenCloseTransaction()
                        Dim oBlTRorigen As BlockTableRecord = CType(acTrans.GetObject(idTableRecord, OpenMode.ForRead), BlockTableRecord)
                        Dim strNombreBloqueDestino As String = BlockTableRecord_NewName(idOld)
                        Dim temDB As Database = db.Wblock(idTableRecord)
                        copyId = db.Insert(strNombreBloqueDestino, temDB, True)
                        If newPoint3D IsNot Nothing Then
                            Dim acBlkTblRDestino As BlockTableRecord = TryCast(acTrans.GetObject(copyId, OpenMode.ForWrite), BlockTableRecord)
                            acBlkTblRDestino.Origin = newPoint3D.Value
                        End If

                        acTrans.Commit()
                    End Using
                End Using
                VaciaMemoria()
            Catch ex As Autodesk.AutoCAD.Runtime.Exception
                MessageBox.Show(ex.Message)
            End Try

            Return copyId
        End Function

        Public Function BlockTableRecord_CloneTransform(ByVal pNombreOriginal As String, ByVal pNombreDestino As String, ByVal pPoint3D As Point3d) As BlockTableRecord
            Dim doc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim db As Database = doc.Database
            Dim acBlkTblRDestino As BlockTableRecord
            Try
                Using acLckDoc As DocumentLock = doc.LockDocument()
                    Using acTrans As Transaction = db.TransactionManager.StartOpenCloseTransaction()
                        Dim acBlkTbl As BlockTable = TryCast(acTrans.GetObject(db.BlockTableId, OpenMode.ForRead), BlockTable)
                        Dim acBlkTblROriginal As BlockTableRecord = TryCast(acTrans.GetObject(acBlkTbl(pNombreOriginal), OpenMode.ForRead), BlockTableRecord)
                        Dim copyId As ObjectId = ObjectId.Null
                        Using temDB As Database = db.Wblock(acBlkTblROriginal.Id)
                            copyId = db.Insert(pNombreDestino, temDB, True)
                        End Using

                        If copyId <> ObjectId.Null Then
                            acBlkTblRDestino = TryCast(acTrans.GetObject(copyId, OpenMode.ForWrite), BlockTableRecord)
                            acBlkTblRDestino.Origin = pPoint3D
                        Else
                            acBlkTblRDestino = Nothing
                        End If

                        acTrans.Commit()
                    End Using
                End Using
                '
                VaciaMemoria()
                Return acBlkTblRDestino
            Catch ex As Autodesk.AutoCAD.Runtime.Exception
                MessageBox.Show(ex.Message)
                Return Nothing
            End Try
        End Function
        Public Function BlockTableRecord_Get(ByVal oBlRef As BlockReference) As BlockTableRecord
            Dim doc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim db As Database = doc.Database
            Dim oBlRefTR As BlockTableRecord
            Try
                Using acTrans As Transaction = db.TransactionManager.StartOpenCloseTransaction()
                    oBlRefTR = TryCast(acTrans.GetObject(oBlRef.BlockTableRecord, OpenMode.ForRead), BlockTableRecord)
                    acTrans.Commit()
                End Using

                VaciaMemoria()
                Return oBlRefTR
            Catch ex As Autodesk.AutoCAD.Runtime.Exception
                MessageBox.Show(ex.Message)
                Return Nothing
            End Try
        End Function
        Public Function BlockTableRecord_Get(ByVal pIdBlkTblR As ObjectId) As BlockTableRecord
            Dim doc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim db As Database = doc.Database
            Dim acBlkTblR As BlockTableRecord
            Try
                Using acTrans As Transaction = db.TransactionManager.StartOpenCloseTransaction()
                    acBlkTblR = TryCast(acTrans.GetObject(pIdBlkTblR, OpenMode.ForRead), BlockTableRecord)
                    acTrans.Commit()
                End Using
                '
                VaciaMemoria()
                Return acBlkTblR
            Catch ex As Autodesk.AutoCAD.Runtime.Exception
                MessageBox.Show(ex.Message)
                Return Nothing
            End Try
        End Function

        Public Function BlockTableRecord_Get(ByVal pStrName As String) As BlockTableRecord
            Dim doc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim db As Database = doc.Database
            Dim acBlkTblR As BlockTableRecord = Nothing
            Try
                Using acTrans As Transaction = db.TransactionManager.StartOpenCloseTransaction()
                    Dim acBlkTbl As BlockTable = TryCast(acTrans.GetObject(db.BlockTableId, OpenMode.ForRead), BlockTable)
                    acBlkTblR = TryCast(acTrans.GetObject(acBlkTbl(pStrName), OpenMode.ForRead), BlockTableRecord)
                    acTrans.Commit()
                End Using
                '
                VaciaMemoria()
            Catch ex As Autodesk.AutoCAD.Runtime.Exception
                MessageBox.Show(ex.Message)
            End Try

            Return acBlkTblR
        End Function
        Public Function BlockTableRecord_GetCountChilds(oBlTR As BlockTableRecord) As Integer
            Dim resultado As Integer = 0
            For Each oAcId As ObjectId In oBlTR
                resultado += 1
                System.Windows.Forms.Application.DoEvents()
            Next
            VaciaMemoria()
            Return resultado
        End Function

        Public Function Entity_Get(ByVal pObjectId As ObjectId) As Entity
            Dim doc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim db As Database = doc.Database
            Dim acEntity As Entity
            Try
                Using acTrans As Transaction = db.TransactionManager.StartOpenCloseTransaction()
                    acEntity = TryCast(acTrans.GetObject(pObjectId, OpenMode.ForRead), Entity)
                    acTrans.Commit()
                End Using
                '
                VaciaMemoria()
                Return acEntity
            Catch ex As Autodesk.AutoCAD.Runtime.Exception
                MessageBox.Show(ex.Message)
                Return Nothing
            End Try
        End Function
        Public Function Entity_Get(ByVal pObjectId As Long) As Entity
            Dim oId As ObjectId = New ObjectId(New IntPtr(pObjectId))
            Return Entity_Get(oId)
        End Function
        Public Function DBObject_Get(ByVal pObjectId As ObjectId) As DBObject
            Dim doc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim db As Database = doc.Database
            Dim acDBObject As DBObject = Nothing
            Dim acEntity As Entity = Nothing
            Try
                Using acTrans As Transaction = db.TransactionManager.StartOpenCloseTransaction()
                    acDBObject = TryCast(acTrans.GetObject(pObjectId, OpenMode.ForRead), DBObject)
                    'acEntity = TryCast(acTrans.GetObject(pObjectId, OpenMode.ForRead), Entity)
                    acTrans.Commit()
                End Using
                '
                VaciaMemoria()
                Return acDBObject
            Catch ex As Autodesk.AutoCAD.Runtime.Exception
                'MessageBox.Show(ex.Message)
                Return Nothing
            End Try
        End Function
        Public Function DBObject_Get(ByVal pObjectId As Long) As DBObject
            Dim oId As ObjectId = New ObjectId(New IntPtr(pObjectId))
            Return DBObject_Get(oId)
        End Function

        Public Function BlockReference_Get(ByVal pIdBlkRef As ObjectId) As BlockReference
            Dim doc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim db As Database = doc.Database
            Dim acBlkTblRef As BlockReference
            Try
                Using acTrans As Transaction = db.TransactionManager.StartOpenCloseTransaction()
                    acBlkTblRef = TryCast(acTrans.GetObject(pIdBlkRef, OpenMode.ForRead), BlockReference)
                End Using
                '
                VaciaMemoria()
                Return acBlkTblRef
            Catch ex As System.Exception
                Return Nothing
            End Try
        End Function
        Public Function BlockReference_Insert(ByVal idBlTR As ObjectId, idBlRefOld As ObjectId, Optional attributes As Dictionary(Of String, Object) = Nothing) As ObjectId
            Dim doc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim db As Database = doc.Database
            Try
                Dim idBlRefNew As ObjectId
                Using acLckDoc As DocumentLock = doc.LockDocument()
                    Using acTrans As Transaction = db.TransactionManager.StartOpenCloseTransaction()
                        Dim oBlTRCurrentSpace As BlockTableRecord = TryCast(acTrans.GetObject(db.CurrentSpaceId, OpenMode.ForWrite), BlockTableRecord)
                        Dim oBlTR As BlockTableRecord = TryCast(acTrans.GetObject(idBlTR, OpenMode.ForRead), BlockTableRecord)
                        Dim oBlkRefOld As BlockReference = TryCast(acTrans.GetObject(idBlRefOld, OpenMode.ForRead), BlockReference)
                        Dim oBlkRefNew As BlockReference = New BlockReference(oBlkRefOld.Position, idBlTR)
                        oBlkRefNew.Rotation = oBlkRefOld.Rotation
                        oBlkRefNew.ScaleFactors = oBlkRefOld.ScaleFactors
                        oBlkRefNew.Layer = oBlkRefOld.Layer
                        '
                        idBlRefNew = oBlTRCurrentSpace.AppendEntity(oBlkRefNew)
                        acTrans.AddNewlyCreatedDBObject(oBlkRefNew, True)
                        '
                        If attributes IsNot Nothing Then
                            For Each item As KeyValuePair(Of String, Object) In attributes
                                Dim attDef As AttributeDefinition = AttributeDefinition_AddBlockFromTableRecord(oBlTR, New DictionaryEntry(item.Key, item.Value), acTrans)
                                Using attRef As AttributeReference = New AttributeReference()
                                    attRef.SetAttributeFromBlock(attDef, oBlkRefNew.BlockTransform)
                                    attRef.TextString = attributes(attDef.Tag).ToString()
                                    oBlkRefNew.AttributeCollection.AppendAttribute(attRef)
                                End Using
                            Next
                        End If
                        acTrans.Commit()
                    End Using
                End Using
                'Vacia()
                VaciaMemoria()
                Return idBlRefNew
            Catch ex As Autodesk.AutoCAD.Runtime.Exception
                MessageBox.Show(ex.Message)
                Return ObjectId.Null
            End Try
        End Function

        Public Function BlockReference_Insert(ByVal pStrName As String, idBlRefOld As ObjectId, Optional attributes As Dictionary(Of String, Object) = Nothing) As ObjectId
            Dim doc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim db As Database = doc.Database
            Try
                Dim blkRefId As ObjectId = ObjectId.Null
                Using acLckDoc As DocumentLock = doc.LockDocument()
                    Using acTrans As Transaction = db.TransactionManager.StartOpenCloseTransaction()
                        Dim oBlT As BlockTable = TryCast(acTrans.GetObject(db.BlockTableId, OpenMode.ForRead), BlockTable)
                        Dim oBlTRCurrentSpace As BlockTableRecord = TryCast(acTrans.GetObject(db.CurrentSpaceId, OpenMode.ForWrite), BlockTableRecord)
                        If oBlT.Has(pStrName) Then
                            Dim oBlTR As BlockTableRecord = TryCast(acTrans.GetObject(oBlT(pStrName), OpenMode.ForRead), BlockTableRecord)
                            Dim oBLRefOld As BlockReference = TryCast(acTrans.GetObject(idBlRefOld, OpenMode.ForRead), BlockReference)
                            Dim oBlRefNew As BlockReference = New BlockReference(oBLRefOld.Position, oBlTR.Id)
                            oBlRefNew.Rotation = oBLRefOld.Rotation
                            oBlRefNew.ScaleFactors = oBLRefOld.ScaleFactors
                            oBlRefNew.Layer = oBLRefOld.Layer
                            blkRefId = oBlTRCurrentSpace.AppendEntity(oBlRefNew)
                            acTrans.AddNewlyCreatedDBObject(oBlRefNew, True)
                            If attributes IsNot Nothing Then
                                For Each item As KeyValuePair(Of String, Object) In attributes
                                    Dim attDef As AttributeDefinition = AttributeDefinition_AddBlockFromTableRecord(oBlTR, New DictionaryEntry(item.Key, item.Value), acTrans)
                                    Using attRef As AttributeReference = New AttributeReference()
                                        attRef.SetAttributeFromBlock(attDef, oBlRefNew.BlockTransform)
                                        attRef.TextString = attributes(attDef.Tag).ToString()
                                        oBlRefNew.AttributeCollection.AppendAttribute(attRef)
                                    End Using
                                Next
                            End If
                        End If
                        acTrans.Commit()
                    End Using
                End Using

                VaciaMemoria()
                Return blkRefId
            Catch ex As Autodesk.AutoCAD.Runtime.Exception
                MessageBox.Show(ex.Message)
                Return ObjectId.Null
            End Try
        End Function

        Public Function BlockReference_Insert(ByVal idBlTR As ObjectId, newPt As Point3d, newRotation As Double, newScale As Double, newLayer As String, Optional attributes As Dictionary(Of String, Object) = Nothing) As ObjectId
            Dim doc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim db As Database = doc.Database
            Try
                Dim idBlRefNew As ObjectId
                Using acLckDoc As DocumentLock = doc.LockDocument()
                    Using acTrans As Transaction = db.TransactionManager.StartOpenCloseTransaction()
                        Dim oBlTRCurrentSpace As BlockTableRecord = TryCast(acTrans.GetObject(db.CurrentSpaceId, OpenMode.ForWrite), BlockTableRecord)
                        Dim oBlTR As BlockTableRecord = TryCast(acTrans.GetObject(idBlTR, OpenMode.ForRead), BlockTableRecord)
                        Dim oBlkRefNew As BlockReference = New BlockReference(newPt, idBlTR)
                        oBlkRefNew.Rotation = newRotation
                        oBlkRefNew.ScaleFactors = New Scale3d(newScale)
                        oBlkRefNew.Layer = newLayer
                        '
                        idBlRefNew = oBlTRCurrentSpace.AppendEntity(oBlkRefNew)
                        acTrans.AddNewlyCreatedDBObject(oBlkRefNew, True)
                        '
                        If attributes IsNot Nothing Then
                            For Each item As KeyValuePair(Of String, Object) In attributes
                                Dim attDef As AttributeDefinition = AttributeDefinition_AddBlockFromTableRecord(oBlTR, New DictionaryEntry(item.Key, item.Value), acTrans)
                                Using attRef As AttributeReference = New AttributeReference()
                                    attRef.SetAttributeFromBlock(attDef, oBlkRefNew.BlockTransform)
                                    attRef.TextString = attributes(attDef.Tag).ToString()
                                    oBlkRefNew.AttributeCollection.AppendAttribute(attRef)
                                End Using
                            Next
                        End If
                        acTrans.Commit()
                    End Using
                End Using
                '
                VaciaMemoria()
                Return idBlRefNew
            Catch ex As Autodesk.AutoCAD.Runtime.Exception
                MessageBox.Show(ex.Message)
                Return ObjectId.Null
            End Try
        End Function

        Public Function BlockReference_Insert(ByVal pStrName As String, newPt As Point3d, newRotation As Double, newScale As Double, newLayer As String, Optional attributes As Dictionary(Of String, Object) = Nothing) As ObjectId
            Dim doc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim db As Database = doc.Database
            Try
                Dim blkRefId As ObjectId = ObjectId.Null
                Using acLckDoc As DocumentLock = doc.LockDocument()
                    Using acTrans As Transaction = db.TransactionManager.StartOpenCloseTransaction()
                        Dim oBlT As BlockTable = TryCast(acTrans.GetObject(db.BlockTableId, OpenMode.ForRead), BlockTable)
                        Dim oBlTRCurrentSpace As BlockTableRecord = TryCast(acTrans.GetObject(db.CurrentSpaceId, OpenMode.ForWrite), BlockTableRecord)
                        If oBlT.Has(pStrName) Then
                            Dim oBlTR As BlockTableRecord = TryCast(acTrans.GetObject(oBlT(pStrName), OpenMode.ForRead), BlockTableRecord)
                            Dim oBlRefNew As BlockReference = New BlockReference(newPt, oBlTR.Id)
                            oBlRefNew.Rotation = newRotation
                            oBlRefNew.ScaleFactors = New Scale3d(newScale)
                            oBlRefNew.Layer = newLayer
                            blkRefId = oBlTRCurrentSpace.AppendEntity(oBlRefNew)
                            acTrans.AddNewlyCreatedDBObject(oBlRefNew, True)
                            If attributes IsNot Nothing Then
                                For Each item As KeyValuePair(Of String, Object) In attributes
                                    Dim attDef As AttributeDefinition = AttributeDefinition_AddBlockFromTableRecord(oBlTR, New DictionaryEntry(item.Key, item.Value), acTrans)
                                    Using attRef As AttributeReference = New AttributeReference()
                                        attRef.SetAttributeFromBlock(attDef, oBlRefNew.BlockTransform)
                                        attRef.TextString = attributes(attDef.Tag).ToString()
                                        oBlRefNew.AttributeCollection.AppendAttribute(attRef)
                                    End Using
                                Next
                            End If
                        End If
                        acTrans.Commit()
                    End Using
                End Using

                VaciaMemoria()
                Return blkRefId
            Catch ex As Autodesk.AutoCAD.Runtime.Exception
                MessageBox.Show(ex.Message)
                Return ObjectId.Null
            End Try
        End Function

        'AFLETA
        Public Function Entity_Erased(ByVal idBlkRef As ObjectId) As Entity
            Dim doc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim db As Database = doc.Database
            Dim ent As Entity = Nothing
            Try

                Using acTrans As Transaction = db.TransactionManager.StartOpenCloseTransaction()
                    If (idBlkRef.IsErased) Then
                        ent = acTrans.GetObject(idBlkRef, OpenMode.ForRead, True)

                    End If

                End Using

                VaciaMemoria()
                Return ent
            Catch ex As Autodesk.AutoCAD.Runtime.Exception

                Return ent
            End Try
        End Function

        Public Function BlockTableRecord_DameVector3D(ByRef oBlTR As BlockTableRecord, ByRef acTrans As Transaction) As Vector3d
            Dim resultado As Vector3d = New Vector3d(0, 0, 0)
            ' Calcular el desplazamiento necesario para las entidades.
            ' Cojer la primera que podamos interpretar.
            Dim ptOri As Point3d = oBlTR.Origin
            '
            For Each oId As ObjectId In oBlTR
                Dim oEnt As Entity = TryCast(acTrans.GetObject(oId, OpenMode.ForWrite), Entity)
                Select Case oEnt.GetType
                    Case GetType(BlockReference)
                        resultado = CType(oEnt, BlockReference).Position.GetVectorTo(ptOri)
                End Select
                If resultado.X <> 0 Then Exit For
            Next
            '
            VaciaMemoria()
            Return resultado
        End Function
        Public Sub BlockTableRecord_Change(ByVal pAcBlkRef As BlockReference, ByVal pNmBlkTblRHijo As String)
            Dim doc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim db As Database = doc.Database
            Try
                Using acLckDoc As DocumentLock = doc.LockDocument()
                    Using acTrans As Transaction = db.TransactionManager.StartOpenCloseTransaction()
                        Dim acBlkTbl As BlockTable = TryCast(acTrans.GetObject(db.BlockTableId, OpenMode.ForWrite), BlockTable)
                        If acBlkTbl.Has(pNmBlkTblRHijo) Then
                            Dim acBlkTblR As BlockTableRecord = TryCast(acTrans.GetObject(acBlkTbl(pNmBlkTblRHijo), OpenMode.ForWrite), BlockTableRecord)
                            pAcBlkRef.BlockTableRecord = acBlkTblR.Id
                        End If

                        acTrans.Commit()
                    End Using
                End Using
                '
                VaciaMemoria()
            Catch ex As Autodesk.AutoCAD.Runtime.Exception
                MessageBox.Show(ex.Message)
            End Try
        End Sub
        Public Function BlockReference_Delete(ByVal pIdAcBlkRef As ObjectId) As Boolean
            Dim doc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim db As Database = doc.Database
            Try
                Using acLckDoc As DocumentLock = doc.LockDocument()
                    Using acTrans As Transaction = db.TransactionManager.StartOpenCloseTransaction()
                        Dim blkRef As BlockReference = CType(acTrans.GetObject(pIdAcBlkRef, OpenMode.ForWrite), BlockReference)
                        If blkRef.IsErased = False Then
                            blkRef.[Erase]()
                        End If

                        acTrans.Commit()
                        db.TransactionManager.QueueForGraphicsFlush()
                    End Using
                End Using
                '
                VaciaMemoria()
                Return True
            Catch ex As Autodesk.AutoCAD.Runtime.Exception
                MessageBox.Show(ex.Message)
                Return False
            End Try
        End Function

        Public Function BlockReference_Delete(ByVal pIdBlkTblR As ObjectId, ByVal pNmBlockDelete As String) As Boolean
            Dim doc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim db As Database = doc.Database
            Try
                Using acLckDoc As DocumentLock = doc.LockDocument()
                    Using acTrans As Transaction = db.TransactionManager.StartOpenCloseTransaction()
                        Dim AcBlkTblR As BlockTableRecord = TryCast(acTrans.GetObject(pIdBlkTblR, OpenMode.ForRead), BlockTableRecord)
                        For Each _ObjectIdEntity As ObjectId In AcBlkTblR
                            Dim acEntity As Entity = TryCast(acTrans.GetObject(_ObjectIdEntity, OpenMode.ForRead), Entity)
                            If acEntity.[GetType]() = GetType(BlockReference) Then
                                Dim acBlkRef As BlockReference = CType(acEntity, BlockReference)
                                If acBlkRef.Name = pNmBlockDelete Then
                                    acEntity.UpgradeOpen()
                                    acEntity.[Erase]()
                                End If
                            End If
                        Next

                        acTrans.Commit()
                    End Using
                End Using

                VaciaMemoria()
                Return True
            Catch ex As Autodesk.AutoCAD.Runtime.Exception
                MessageBox.Show(ex.Message)
                Return False
            End Try
        End Function

        Public Function BlockTableRecord_ExistName(ByVal pNombre As String) As Boolean
            Dim resultado As Boolean = False
            Dim doc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim db As Database = doc.Database
            Try
                Using acTrans As Transaction = db.TransactionManager.StartOpenCloseTransaction()
                    Dim acBlkTbl As BlockTable = TryCast(acTrans.GetObject(db.BlockTableId, OpenMode.ForRead), BlockTable)
                    If acBlkTbl.Has(pNombre) Then
                        resultado = True
                    Else
                        resultado = False
                    End If
                End Using
                '
                VaciaMemoria()
                Return resultado
            Catch ex As Autodesk.AutoCAD.Runtime.Exception
                MessageBox.Show(ex.Message)
                Return Nothing
            End Try
        End Function
        Public Function BlockTableRecord_NewNameAutonumeric(ByVal pNombre As String) As String
            Dim doc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim db As Database = doc.Database
            Try
                Using acTrans As Transaction = db.TransactionManager.StartOpenCloseTransaction()
                    Dim acBlkTbl As BlockTable = TryCast(acTrans.GetObject(db.BlockTableId, OpenMode.ForRead), BlockTable)
                    Dim intPos As Integer = 0
                    While acBlkTbl.Has(pNombre)
                        Dim arrayNombre As String() = pNombre.Split({"-pi"}, StringSplitOptions.None)
                        If arrayNombre.Length = 2 Then
                            Integer.TryParse(arrayNombre(1), intPos)
                            intPos = intPos + 1
                            pNombre = arrayNombre(0) & "-pi" + intPos.ToString.PadLeft(2, intPos.ToString)
                        Else
                            pNombre = pNombre & "-pi01"
                        End If
                    End While
                End Using
                VaciaMemoria()
                Return pNombre
            Catch ex As Autodesk.AutoCAD.Runtime.Exception
                MessageBox.Show(ex.Message)
                Return Nothing
            End Try
        End Function
        Public Function BlockTableRecord_NewName(ByVal oId As ObjectId) As String
            Dim doc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim db As Database = doc.Database
            Dim strNombre As String = ""
            Try
                Using acTrans As Transaction = db.TransactionManager.StartOpenCloseTransaction()
                    Dim acBlkTblR As BlockTableRecord = TryCast(acTrans.GetObject(oId, OpenMode.ForRead), BlockTableRecord)
                    strNombre = acBlkTblR.Name.Replace("*", "")
                End Using
                '
                VaciaMemoria()
                Return strNombre
            Catch ex As Autodesk.AutoCAD.Runtime.Exception
                MessageBox.Show(ex.Message)
                Return Nothing
            End Try
        End Function
        '
        Public Function BlockTableRecord_GetBlocksReferences(ByVal pNmBloque As String) As ObjectIdCollection
            Dim doc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim db As Database = doc.Database
            Dim ListObjectId As ObjectIdCollection = New ObjectIdCollection()
            Try
                Using acTrans As Transaction = db.TransactionManager.StartOpenCloseTransaction()
                    Dim acBlkTbl As BlockTable = TryCast(acTrans.GetObject(db.BlockTableId, OpenMode.ForRead), BlockTable)
                    If acBlkTbl.Has(pNmBloque) Then
                        Dim acBlkTblR As BlockTableRecord = TryCast(acTrans.GetObject(acBlkTbl(pNmBloque), OpenMode.ForRead), BlockTableRecord)
                        ListObjectId = acBlkTblR.GetBlockReferenceIds(False, True)
                    End If
                End Using
                VaciaMemoria()
            Catch ex As Autodesk.AutoCAD.Runtime.Exception
                MessageBox.Show(ex.Message)
            End Try
            '
            Return ListObjectId
        End Function
        '
        Public Function BlockTable_GetBlockTableRecordsNames() As List(Of String)
            Dim doc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim db As Database = doc.Database
            Dim ListNames As New List(Of String)
            Try
                Using acTrans As Transaction = db.TransactionManager.StartOpenCloseTransaction()
                    Dim oBlT As BlockTable = TryCast(acTrans.GetObject(db.BlockTableId, OpenMode.ForRead), BlockTable)
                    For Each oId In oBlT
                        Dim oBlTR As BlockTableRecord = TryCast(acTrans.GetObject(oId, OpenMode.ForRead), BlockTableRecord)
                        If ListNames.Contains(oBlTR.Name) = False Then ListNames.Add(oBlTR.Name)
                    Next
                End Using
                VaciaMemoria()
            Catch ex As Autodesk.AutoCAD.Runtime.Exception
                MessageBox.Show(ex.Message)
            End Try
            '
            Return ListNames
        End Function
#End Region
#Region "ATRIBUTOSACTIVEX"


        'Public Function AttributeReference_Find(ByVal pListFilterSelection As TypedValue(), ByVal pStrNmAtributo As String) As List(Of AttributeReference)
        '    Dim doc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        '    Dim db As Database = doc.Database
        '    Dim ed As Editor = doc.Editor
        '    Dim listResult As List(Of AttributeReference) = New List(Of AttributeReference)()
        '    Try
        '        Using acTrans As Transaction = db.TransactionManager.StartOpenCloseTransaction()
        '            Dim acBlkTbl As BlockTable = CType(acTrans.GetObject(db.BlockTableId, OpenMode.ForRead), BlockTable)
        '            Dim filter As SelectionFilter = New SelectionFilter(pListFilterSelection)
        '            Dim selRes As PromptSelectionResult = ed.SelectAll(filter)
        '            If selRes.Status <> PromptStatus.OK Then
        '                ed.WriteMessage(vbLf & "No ha seleccionado un bloque correcto.")
        '                Return listResult
        '            End If

        '            If selRes.Value.Count <> 0 Then
        '                Dim selSet As SelectionSet = selRes.Value
        '                listResult = selSet.GetObjectIds().Select(Function(i)
        '                                                              Return BlockReference_Open(acTrans, i)
        '                                                          End Function).ToList().SelectMany(Function(b)
        '                                                                                                Return b.AttributeCollection.Cast(Of ObjectId)().Select(Function(idAtt)
        '                                                                                                                                                            Return AttributeReference_Open(acTrans, idAtt)
        '                                                                                                                                                        End Function).Where(Function(att)
        '                                                                                                                                                                                Return att.Tag = pStrNmAtributo
        '                                                                                                                                                                            End Function).ToList()
        '                                                                                            End Function).ToList()
        '                ed.WriteMessage(vbLf & "Selección completada correctamente.")
        '            End If
        '        End Using

        '        Return listResult
        '    Catch ex As Autodesk.AutoCAD.Runtime.Exception
        '        ed.WriteMessage(ex.Message.ToString())
        '        Return listResult
        '    End Try
        'End Function

        'Public Function FindAttributeReference(ByVal pListFilterSelection As TypedValue(), ByVal pStrNmAtributo As String, ByVal pStrValue As String) As List(Of AttributeReference)
        '    Dim doc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        '    Dim db As Database = doc.Database
        '    Dim ed As Editor = doc.Editor
        '    Dim listResult As List(Of AttributeReference) = New List(Of AttributeReference)()
        '    Try
        '        Using acTrans As Transaction = db.TransactionManager.StartOpenCloseTransaction()
        '            Dim acBlkTbl As BlockTable = CType(acTrans.GetObject(db.BlockTableId, OpenMode.ForRead), BlockTable)
        '            Dim filter As SelectionFilter = New SelectionFilter(pListFilterSelection)
        '            Dim selRes As PromptSelectionResult = ed.SelectAll(filter)
        '            If selRes.Status <> PromptStatus.OK Then
        '                ed.WriteMessage(vbLf & "No ha seleccionado un bloque correcto.")
        '                Return listResult
        '            End If

        '            If selRes.Value.Count <> 0 Then
        '                Dim [set] As SelectionSet = selRes.Value
        '                listResult = [set].GetObjectIds().[Select](Function(i) OpenBlockReference(acTrans, i)).ToList(Of BlockReference)().SelectMany(Function(b) b.AttributeCollection.Cast(Of ObjectId)().[Select](Function(idAtt) OpenAttributeReference(acTrans, idAtt)).Where(Function(att) att.Tag = pStrNmAtributo AndAlso att.TextString = pStrValue).ToList(Of AttributeReference)()).ToList()
        '                ed.WriteMessage(vbLf & "Selección completada correctamente.")
        '            End If
        '        End Using
        '        'Vacia()
        '        Return listResult
        '    Catch ex As Autodesk.AutoCAD.Runtime.Exception
        '        ed.WriteMessage(ex.Message.ToString())
        '        Return listResult
        '    End Try
        'End Function

        Private Function AttributeReference_Open(ByVal pAcTrans As Transaction, ByVal pObjId As ObjectId) As AttributeReference
            If acIdAttRefCached <> pObjId OrElse acAttRefCached Is Nothing Then
                acAttRefCached = CType(pAcTrans.GetObject(pObjId, OpenMode.ForRead), AttributeReference)
                acIdAttRefCached = acAttRefCached.Id
            End If

            VaciaMemoria()
            Return acAttRefCached
        End Function
        Public Function AttributeReference_Get(ByVal pIdAttRef As ObjectId, ByVal pModo As OpenMode) As AttributeReference
            Dim doc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim db As Database = doc.Database
            Dim acAttRef As AttributeReference = Nothing
            Try
                Using acTrans As Transaction = db.TransactionManager.StartOpenCloseTransaction()
                    acAttRef = CType(acTrans.GetObject(pIdAttRef, pModo), AttributeReference)
                End Using
                VaciaMemoria()
            Catch ex As Autodesk.AutoCAD.Runtime.Exception
            End Try

            Return acAttRef
        End Function

        Public Function AttributeReference_Get(ByVal pAcBlkRef As BlockReference, ByVal pStrNmAtributo As String, ByVal pModo As OpenMode) As AttributeReference
            Dim doc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim db As Database = doc.Database
            Dim acAttRef As AttributeReference = Nothing
            Try
                Using acTrans As Transaction = db.TransactionManager.StartOpenCloseTransaction()
                    For Each idAtt As ObjectId In pAcBlkRef.AttributeCollection
                        Dim acAttRefTemp As AttributeReference = CType(acTrans.GetObject(idAtt, pModo), AttributeReference)
                        If acAttRefTemp.Tag.ToUpper() = pStrNmAtributo.ToUpper() Then
                            acAttRef = acAttRefTemp
                            Exit For
                        End If
                    Next

                    If acAttRef IsNot Nothing AndAlso acAttRef.IsObjectIdsInFlux = True Then
                        acTrans.Commit()
                    End If
                End Using
                VaciaMemoria()
            Catch ex As Autodesk.AutoCAD.Runtime.Exception
                MessageBox.Show(ex.Message)
            End Try

            Return acAttRef
        End Function

        Public Function AttributeReference_Get(ByVal pIdAcBlkRef As ObjectId, ByVal pStrNmAtributo As String, ByVal pModo As OpenMode) As AttributeReference
            Dim doc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim db As Database = doc.Database
            Dim acAttRef As AttributeReference = Nothing
            Try
                Using acTrans As Transaction = db.TransactionManager.StartOpenCloseTransaction()
                    Dim acBlkRef As BlockReference = CType(acTrans.GetObject(pIdAcBlkRef, OpenMode.ForRead), BlockReference)
                    For Each idAtt As ObjectId In acBlkRef.AttributeCollection
                        Dim acAttRefTemp As AttributeReference = CType(acTrans.GetObject(idAtt, pModo), AttributeReference)
                        If acAttRefTemp.Tag.ToUpper() = pStrNmAtributo.ToUpper() Then
                            acAttRef = acAttRefTemp
                            Exit For
                        End If
                    Next

                    If acAttRef IsNot Nothing AndAlso acAttRef.IsObjectIdsInFlux = True Then
                        acTrans.Commit()
                    End If
                End Using
            Catch ex As Autodesk.AutoCAD.Runtime.Exception
                MessageBox.Show(ex.Message)
            End Try

            Return acAttRef
        End Function

        'AFLETA
        Public Function AttributeReference_Get(ByVal pIdAcBlkRef As ObjectId, ByVal pStrNmAtributo As String, ByVal pModo As OpenMode, ByVal BlkRefIsErased As Boolean) As AttributeReference
            Dim doc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim db As Database = doc.Database
            Dim acAttRef As AttributeReference = Nothing
            Try
                If BlkRefIsErased = False Then
                    AttributeReference_Get(pIdAcBlkRef, pStrNmAtributo, pModo)
                Else
                    Using acTrans As Transaction = db.TransactionManager.StartOpenCloseTransaction()
                        Dim acBlkRef As BlockReference = CType(acTrans.GetObject(pIdAcBlkRef, OpenMode.ForRead, True), BlockReference)
                        For Each idAtt As ObjectId In acBlkRef.AttributeCollection
                            Dim acAttRefTemp As AttributeReference = CType(acTrans.GetObject(idAtt, pModo), AttributeReference)
                            If acAttRefTemp.Tag.ToUpper() = pStrNmAtributo.ToUpper() Then
                                acAttRef = acAttRefTemp
                                Exit For
                            End If
                        Next

                        If acAttRef IsNot Nothing AndAlso acAttRef.IsObjectIdsInFlux = True Then
                            acTrans.Commit()
                        End If
                    End Using
                End If


            Catch ex As Autodesk.AutoCAD.Runtime.Exception
                MessageBox.Show(ex.Message)
            End Try

            Return acAttRef
        End Function

        Public Sub AttributeReference_Set(ByVal pAttRef As AttributeReference, ByVal pStrValue As String)
            Dim doc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim db As Database = doc.Database
            If pAttRef Is Nothing Then Exit Sub
            Try
                Using acLckDoc As DocumentLock = doc.LockDocument()
                    Using acTrans As Transaction = db.TransactionManager.StartOpenCloseTransaction()
                        acTrans.Commit()
                        Try
                            If Not pAttRef.IsWriteEnabled Then
                                If pAttRef.IsReadEnabled Then pAttRef.UpgradeOpen() Else pAttRef = CType(acTrans.GetObject(pAttRef.Id, OpenMode.ForWrite), AttributeReference)
                            End If
                            pAttRef.TextString = pStrValue
                            acTrans.Commit()
                        Catch ex As Exception
                            acTrans.Abort()
                        End Try
                    End Using
                End Using
                VaciaMemoria()
            Catch ex As Autodesk.AutoCAD.Runtime.Exception
                MessageBox.Show(ex.Message)
            End Try
        End Sub
        ' Copiar AttributeDefinición y crear AttributeReference de BlockReferenceClonados.
        Public Function AttributeDefinition_AddALLClonados(ByVal idBlRefOrigen As ObjectId, idBlRefDestino As ObjectId) As AttributeCollection
            Dim doc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim db As Database = doc.Database
            Dim oBlRefDestino As BlockReference = Nothing
            Using acLckDoc As DocumentLock = doc.LockDocument()
                Using acTrans As Transaction = db.TransactionManager.StartOpenCloseTransaction()
                    Try
                        Dim acAttDef As AttributeDefinition = New AttributeDefinition()
                        Dim oBlRefOrigen As BlockReference = CType(acTrans.GetObject(idBlRefOrigen, OpenMode.ForRead), BlockReference)
                        Dim oBlTROrigen As BlockTableRecord = CType(acTrans.GetObject(oBlRefOrigen.BlockTableRecord, OpenMode.ForRead), BlockTableRecord)
                        oBlRefDestino = CType(acTrans.GetObject(idBlRefDestino, OpenMode.ForWrite), BlockReference)
                        If oBlTROrigen.HasAttributeDefinitions And oBlRefOrigen.AttributeCollection.Count > 0 Then
                            For Each objID As ObjectId In oBlTROrigen
                                Dim dbObj As DBObject = TryCast(acTrans.GetObject(objID, OpenMode.ForRead), DBObject)
                                If TypeOf dbObj Is AttributeDefinition Then
                                    Dim oAttDef As AttributeDefinition = TryCast(dbObj, AttributeDefinition)
                                    Dim oAttRef As New AttributeReference
                                    Try
                                        oAttRef.SetAttributeFromBlock(oAttDef, oBlRefDestino.BlockTransform)
                                        oAttRef.Position = oAttDef.Position.TransformBy(oBlRefDestino.BlockTransform)
                                        oAttRef.TextString = AttributeReference_Get(oBlRefOrigen, oAttDef.Tag.ToString, OpenMode.ForRead).TextString
                                        oBlRefDestino.AttributeCollection.AppendAttribute(oAttRef)
                                        acTrans.AddNewlyCreatedDBObject(oAttRef, True)
                                    Catch ex As Exception
                                        Continue For
                                    End Try
                                End If
                            Next
                            acTrans.Commit()
                        End If
                    Catch ex As Exception
                        acTrans.Abort()
                    End Try
                End Using
            End Using
            '
            VaciaMemoria()
            Return oBlRefDestino.AttributeCollection
        End Function
        ' Copiar AttributeDefinición y crear AttributeReference de BlockReferenceClonados.
        Public Function AttributeDefinition_AddALLBlockBlockReferencesClonados(ByVal idBlRefOrigen As ObjectId, idBlRefDestino As ObjectId, ByVal oTrans As Transaction) As AttributeCollection
            Dim acAttDef As AttributeDefinition = New AttributeDefinition()
            Dim oBlRefOrigen As BlockReference = CType(oTrans.GetObject(idBlRefOrigen, OpenMode.ForRead), BlockReference)
            Dim oBlTROrigen As BlockTableRecord = CType(oTrans.GetObject(oBlRefOrigen.BlockTableRecord, OpenMode.ForRead), BlockTableRecord)
            Dim oBlRefDestino As BlockReference = CType(oTrans.GetObject(idBlRefDestino, OpenMode.ForWrite), BlockReference)
            If oBlTROrigen.HasAttributeDefinitions Then
                For Each objID As ObjectId In oBlTROrigen
                    Dim dbObj As DBObject = TryCast(oTrans.GetObject(objID, OpenMode.ForRead), DBObject)
                    If TypeOf dbObj Is AttributeDefinition Then
                        Dim oAttDef As AttributeDefinition = TryCast(dbObj, AttributeDefinition)
                        Dim oAttRef As New AttributeReference
                        oAttRef.SetAttributeFromBlock(oAttDef, oBlRefDestino.BlockTransform)
                        oAttRef.Position = oAttDef.Position.TransformBy(oBlRefDestino.BlockTransform)
                        oAttRef.TextString = oAttRef.TextString
                        oBlRefDestino.AttributeCollection.AppendAttribute(oAttRef)
                        oTrans.AddNewlyCreatedDBObject(oAttRef, True)
                    End If
                Next
            End If
            '
            VaciaMemoria()
            Return oBlRefDestino.AttributeCollection
        End Function
        Public Function AttributeDefinition_AddBlockFromTableRecord(ByVal oBlTR As BlockTableRecord, ByVal pItem As System.Collections.DictionaryEntry, ByVal oTrans As Transaction) As AttributeDefinition
            Dim bolExistePropiedad As Boolean = False
            Dim acAttDef As AttributeDefinition = New AttributeDefinition()
            If oBlTR.HasAttributeDefinitions Then
                For Each objID As ObjectId In oBlTR
                    Dim dbObj As DBObject = TryCast(oTrans.GetObject(objID, OpenMode.ForRead), DBObject)
                    If TypeOf dbObj Is AttributeDefinition Then
                        Dim acAtt As AttributeDefinition = TryCast(dbObj, AttributeDefinition)
                        If pItem.Key.ToString().ToUpper() = acAtt.Tag.ToUpper() Then
                            bolExistePropiedad = True
                            acAttDef = acAtt
                            Exit For
                        End If
                    End If
                Next
            End If
            '
            If Not bolExistePropiedad Then
                acAttDef.Position = New Point3d(0, 0, 0)
                acAttDef.Verifiable = False
                acAttDef.Prompt = pItem.Key.ToString()
                acAttDef.Tag = pItem.Key.ToString()
                acAttDef.TextString = pItem.Value.ToString()
                acAttDef.Height = 1
                acAttDef.Justify = AttachmentPoint.MiddleCenter
                acAttDef.Constant = False
                acAttDef.Visible = True
                acAttDef.Invisible = False
                oBlTR.AppendEntity(acAttDef)
            End If

            VaciaMemoria()
            Return acAttDef
        End Function

        Public Function AttributeDefinition_AddBlockFromTableRecord(ByVal oBlTR As BlockTableRecord, nameAtt As String, valueAtt As String, ByVal oTrans As Transaction) As AttributeDefinition
            Dim bolExistePropiedad As Boolean = False
            Dim acAttDef As AttributeDefinition = New AttributeDefinition()
            If oBlTR.HasAttributeDefinitions Then
                For Each objID As ObjectId In oBlTR
                    Dim dbObj As DBObject = TryCast(oTrans.GetObject(objID, OpenMode.ForRead), DBObject)
                    If TypeOf dbObj Is AttributeDefinition Then
                        Dim acAtt As AttributeDefinition = TryCast(dbObj, AttributeDefinition)
                        If nameAtt.ToUpper() = acAtt.Tag.ToUpper() Then
                            bolExistePropiedad = True
                            acAttDef = acAtt
                            Exit For
                        End If
                    End If
                Next
            End If

            If Not bolExistePropiedad Then
                acAttDef.Position = New Point3d(0, 0, 0)
                acAttDef.Verifiable = False
                acAttDef.Prompt = nameAtt
                acAttDef.Tag = nameAtt
                acAttDef.TextString = valueAtt
                acAttDef.Height = 1
                acAttDef.Justify = AttachmentPoint.MiddleCenter
                acAttDef.Constant = False
                acAttDef.Visible = True
                acAttDef.Invisible = False
                oBlTR.AppendEntity(acAttDef)
            End If

            VaciaMemoria()
            Return acAttDef
        End Function
        Public Function AttributeReference_AddFromBlockReference(ByVal oAttDefLleno As AttributeDefinition, ByVal oBlRefSinAtt As BlockReference, ByVal tr As Transaction) As AttributeReference
            Dim acAttRef As AttributeReference = New AttributeReference()
            acAttRef.SetAttributeFromBlock(oAttDefLleno, oBlRefSinAtt.BlockTransform)
            acAttRef.Position = oAttDefLleno.Position.TransformBy(oBlRefSinAtt.BlockTransform)
            acAttRef.TextString = oAttDefLleno.TextString
            oBlRefSinAtt.AttributeCollection.AppendAttribute(acAttRef)
            tr.AddNewlyCreatedDBObject(acAttRef, True)
            VaciaMemoria()
            Return acAttRef
        End Function
        '' Los BlockReference tienen que ser iguales. Se acaban de clonar y tienen los mismos atributos.
        'Public Function AttributeReferences_CopyFromCloneBlockReference(ByVal oBlRefOri As BlockReference, oBlRefDes As BlockReference) As AttributeCollection
        '    Dim doc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        '    Dim db As Database = doc.Database
        '    Dim oAttRef As AttributeReference = Nothing
        '    Try
        '        Using acLckDoc As DocumentLock = doc.LockDocument()
        '            Using acTrans As Transaction = db.TransactionManager.StartOpenCloseTransaction()
        '                For Each idAtt As ObjectId In oBlRefOri.AttributeCollection
        '                    oAttRef = CType(acTrans.GetObject(idAtt, OpenMode.ForRead), AttributeReference)
        '                    Try
        '                        Dim oAttRefNew As AttributeReference = AttributeReference_Get(oBlRefDes, oAttRef.Tag, OpenMode.ForRead)
        '                        AttributeReference_Set(oAttRefNew, oAttRef.TextString)
        '                    Catch ex As Exception
        '                        Continue For
        '                    End Try
        '                Next
        '                acTrans.Commit()
        '            End Using
        '        End Using
        '        Vacia()
        '        '
        '        Return oBlRefDes.AttributeCollection
        '    Catch ex As Autodesk.AutoCAD.Runtime.Exception
        '        MessageBox.Show(ex.Message)
        '        Return Nothing
        '    End Try
        'End Function
        ' Actualizar atributos. Ya hay iniciada una transacción.
        Public Sub AttributeReference_Update(idBlRef As ObjectId, db As Database, tr As Transaction)
            Dim doc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Try
                Dim br As BlockReference = CType(tr.GetObject(idBlRef, OpenMode.ForRead), BlockReference)
                Using btr As BlockTableRecord = CType(tr.GetObject(br.BlockTableRecord, OpenMode.ForRead), BlockTableRecord)
                    '
                    If btr.HasAttributeDefinitions Then
                        For Each id As ObjectId In btr
                            Dim obj As DBObject = tr.GetObject(id, OpenMode.ForRead)
                            If TypeOf obj Is AttributeDefinition Then
                                Dim ad As AttributeDefinition = CType(tr.GetObject(id, OpenMode.ForWrite), AttributeDefinition)
                                '
                            End If
                        Next

                    End If
                End Using
                VaciaMemoria()
            Catch ex As Autodesk.AutoCAD.Runtime.Exception
                MessageBox.Show(ex.Message)
            End Try
        End Sub
#End Region
    End Class
End Namespace