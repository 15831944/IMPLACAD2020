Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry

Namespace A2acad
    Partial Public Class A2acad
        'AFLETA
        Public Sub MLeader_Insert(oPointInsert As Point3d, pID As String)
            Dim doc As Document = Application.DocumentManager.MdiActiveDocument
            Dim db As Database = doc.Database
            Dim ed As Editor = doc.Editor

            Try
                Using acLckDoc As DocumentLock = doc.LockDocument()


                    Using acTrans As Transaction = db.TransactionManager.StartOpenCloseTransaction()
                        Dim acBlkTbl As BlockTable = CType(acTrans.GetObject(db.BlockTableId, OpenMode.ForRead), BlockTable)
                        Dim model As BlockTableRecord = acTrans.GetObject(db.CurrentSpaceId, OpenMode.ForWrite)

                        If (acBlkTbl.Has("BloqueProxy") = False) Then
                            MsgBox("There is not BloqueProxy.")
                        End If

                        Dim leader As MLeader = New MLeader
                        leader.SetDatabaseDefaults()
                        leader.ContentType = ContentType.BlockContent
                        leader.BlockContentId = acBlkTbl("BloqueProxy")
                        leader.BlockPosition = New Point3d(oPointInsert.X + 50, oPointInsert.Y + 50, 0)

                        Dim idx As Integer = leader.AddLeaderLine(leader.BlockPosition)

                        leader.AddFirstVertex(idx, New Point3d(oPointInsert.X, oPointInsert.Y, 0))
                        'Handle Block Attributes

                        Dim blkLeader As BlockTableRecord = CType(acTrans.GetObject(leader.BlockContentId, OpenMode.ForRead), BlockTableRecord)
                        'Doesn't take in consideration oLeader.BlockRotation
                        Dim Transfo As Matrix3d = Matrix3d.Displacement(leader.BlockPosition.GetAsVector)

                        If blkLeader.HasAttributeDefinitions Then
                            For Each blkEntId As ObjectId In blkLeader

                                Dim dbObj As DBObject = CType(acTrans.GetObject(blkEntId, OpenMode.ForRead), DBObject)
                                If TypeOf dbObj Is AttributeDefinition Then

                                    Dim oAttDef As AttributeDefinition = TryCast(dbObj, AttributeDefinition)

                                    If (Not (oAttDef) Is Nothing) Then
                                        Dim AttributeRef As AttributeReference = New AttributeReference

                                        AttributeRef.SetAttributeFromBlock(oAttDef, Transfo)
                                        AttributeRef.Position = oAttDef.Position.TransformBy(Transfo)

                                        If (oAttDef.Tag.ToUpper() = "ELEMENTO") Then
                                            AttributeRef.TextString = pID

                                        Else
                                            AttributeRef.TextString = ""
                                        End If

                                        leader.SetBlockAttribute(blkEntId, AttributeRef)


                                    End If
                                End If

                            Next
                        End If

                        model.AppendEntity(leader)
                        acTrans.AddNewlyCreatedDBObject(leader, True)
                        acTrans.Commit()

                    End Using
                End Using


            Catch ex As Autodesk.AutoCAD.Runtime.Exception
                ed.WriteMessage(ex.Message.ToString())
            End Try

        End Sub

        'AFLETA
        Public Function AttributeReference_Get_FromMLeader(ByVal pIdAcBlkRef As ObjectId, ByVal pStrNmAtributo As String, ByVal pModo As OpenMode, ByVal BlkRefIsErased As Boolean) As AttributeReference
            Dim doc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim db As Database = doc.Database
            Dim acAttRef As AttributeReference = Nothing
            Try
                'If BlkRefIsErased = False Then
                '    AttributeReference_Get(pIdAcBlkRef, pStrNmAtributo, pModo)
                'Else
                Using acTrans As Transaction = db.TransactionManager.StartOpenCloseTransaction()
                    Dim acMleader As MLeader = CType(acTrans.GetObject(pIdAcBlkRef, OpenMode.ForRead, BlkRefIsErased), MLeader)
                    Dim blkLeader As BlockTableRecord = CType(acTrans.GetObject(acMleader.BlockContentId, OpenMode.ForRead), BlockTableRecord)
                    If blkLeader.HasAttributeDefinitions Then
                        For Each idAtt As ObjectId In blkLeader

                            Dim dbObj As DBObject = CType(acTrans.GetObject(idAtt, OpenMode.ForRead), DBObject)
                            If TypeOf dbObj Is AttributeDefinition Then
                                Dim acAttRefTemp As AttributeReference = acMleader.GetBlockAttribute(dbObj.Id)
                                If acAttRefTemp.Tag.ToUpper() = pStrNmAtributo.ToUpper() Then
                                    acAttRef = acAttRefTemp
                                    Exit For
                                End If
                            End If

                        Next

                        If acAttRef IsNot Nothing AndAlso acAttRef.IsObjectIdsInFlux = True Then
                            acTrans.Commit()
                        End If
                    End If

                End Using
                'End If
                Return acAttRef

            Catch ex As Autodesk.AutoCAD.Runtime.Exception
                Return Nothing
            End Try

            Return acAttRef
        End Function

    End Class
End Namespace