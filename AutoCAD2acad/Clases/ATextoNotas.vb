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
        Public Sub CoberturaOnOff(Optional activar As Boolean = False)  '_WIPEOUT
            If activar = False Then
                oAppA.ActiveDocument.SendCommand("_FRAME 0 ")
                oAppA.ActiveDocument.SendCommand("_WIPEOUTFRAME 0 ")
            Else
                oAppA.ActiveDocument.SendCommand("_FRAME 2 ")
                oAppA.ActiveDocument.SendCommand("_WIPEOUTFRAME 2 ")
            End If
            ' Solo funciona en Español
            'If activar = True Then
            '    oAppA.ActiveDocument.SendCommand("_WIPEOUT M ACT ")
            'Else
            '    oAppA.ActiveDocument.SendCommand("_WIPEOUT M DES ")
            'End If
            ' Solo funciona en Ingles
            'If activar = True Then
            '    oAppA.ActiveDocument.SendCommand("_WIPEOUT F ON ")
            'Else
            '    oAppA.ActiveDocument.SendCommand("_WIPEOUT F OFF ")
            'End If
        End Sub

        Public Sub WipeOut_CreatePolilyne()
            Dim doc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim db As Database = doc.Database
            Dim tr As Transaction = db.TransactionManager.StartTransaction()

            Using tr
                Dim bt As BlockTable = CType(tr.GetObject(db.BlockTableId, OpenMode.ForRead, False), BlockTable)
                Dim btr As BlockTableRecord = CType(tr.GetObject(bt(BlockTableRecord.ModelSpace), OpenMode.ForWrite, False), BlockTableRecord)
                Dim pts As Point2dCollection = New Point2dCollection(5)
                pts.Add(New Point2d(0.0, 0.0))
                pts.Add(New Point2d(100.0, 0.0))
                pts.Add(New Point2d(100.0, 100.0))
                pts.Add(New Point2d(0.0, 100.0))
                pts.Add(New Point2d(0.0, 0.0))
                Dim wo As Wipeout = New Wipeout()
                wo.SetDatabaseDefaults(db)
                wo.SetFrom(pts, New Vector3d(0.0, 0.0, 0.1))
                btr.AppendEntity(wo)
                tr.AddNewlyCreatedDBObject(wo, True)
                tr.Commit()
            End Using
        End Sub


        Public Function MleaderDameTodos_PorNombreBloque(quenombre As String, Optional exacto As Boolean = False) As ArrayList
            Dim resultado As New ArrayList
            If oAppA Is Nothing Then _
        oAppA = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
            ''
            For Each oEnt As AcadEntity In oAppA.ActiveDocument.ModelSpace
                If TypeOf oEnt Is AcadMLeader Then
                    Dim oMl As AcadMLeader = oAppA.ActiveDocument.ObjectIdToObject(oEnt.ObjectID)
                    If exacto = True Then
                        If oMl.ContentBlockName.ToUpper = quenombre.ToUpper Then
                            resultado.Add(oMl)
                        End If
                    Else
                        If oMl.ContentBlockName.ToUpper.Contains(quenombre.ToUpper) Then
                            resultado.Add(oMl)
                        End If
                    End If
                    oMl = Nothing
                End If
            Next
            VaciaMemoria()
            Return resultado
        End Function

        Public Function MLeaderBlock_DameValorAtributo(oMl As AcadMLeader, nombreAtt As String) As String
            Dim resultado As String = ""
            ' This example creates an MLeader object and gets and sets values for
            ' the block attribute type.
            'Dim points(0 To 5) As Double
            'points(0) = 0 : points(1) = 4 : points(2) = 0
            'points(3) = 1.5 : points(4) = 5 : points(5) = 0
            'Dim i As Long
            'Dim oML As AcadMLeader
            'oML = ThisDrawing.ModelSpace.AddMLeader(points, i)
            'oML.ContentType = acBlockContent
            'oML.ContentBlockType = acBlockBox

            Dim sBlock As String
            sBlock = oMl.ContentBlockName

            Dim o As AcadEntity
            For Each o In oAppA.ActiveDocument.Blocks.Item(sBlock)
                If o.ObjectName = "AcDbAttributeDefinition" Then
                    'oMl.SetBlockAttributeValue o.ObjectID, "123"
                    'Dim oAttDef As AttributeDefinition = HostApplicationServices.WorkingDatabase.get oAppA.ActiveDocument.Database.ObjectIdToObject(o.ObjectID)
                    If o.tagstring.ToString.ToUpper = nombreAtt.ToUpper Then
                        resultado = oMl.GetBlockAttributeValue(o.ObjectID)
                        Exit For
                    End If
                End If
            Next o

            'Update
            'ZoomExtents
            Return resultado
        End Function
        Public Function MLeaderBlock_PonValorAtributo(oMl As AcadMLeader, nombreAtt As String, valor As String) As Boolean
            Dim resultado As Boolean = False
            Dim sBlock As String
            sBlock = oMl.ContentBlockName
            Dim o As AcadEntity
            For Each o In oAppA.ActiveDocument.Blocks.Item(sBlock)
                If o.ObjectName = "AcDbAttributeDefinition" Then
                    'oMl.SetBlockAttributeValue o.ObjectID, "123"
                    'Dim oAttDef As AttributeDefinition = HostApplicationServices.WorkingDatabase.get oAppA.ActiveDocument.Database.ObjectIdToObject(o.ObjectID)
                    If o.Tagstring.ToString = nombreAtt Then
                        Try
                            oMl.SetBlockAttributeValue(o.ObjectID, valor)
                            oMl.Update()
                            resultado = True
                        Catch ex As Exception
                            resultado = False
                        End Try
                    End If
                End If
            Next o
            '
            Return resultado
        End Function
        Public Function MLeader_InsertaCommand() As AcadMLeader
            ''Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.SendStringToExecute("\x03\x03", False, True, False)
            Dim resultado As AcadMLeader = Nothing
            Dim ptIns(2) As Double
            'Dim oCir As AcadCircle = Nothing
            Try
                ' Creamos un circulo para que este sea el último objecto del EspacioModelo
                'oCir = oAppA.ActiveDocument.ModelSpace.AddCircle(ptIns, 10)
                oAppA.ActiveDocument.SendCommand("_MLEADER ")
                '
                Dim oEnt As AcadEntity = oAppA.ActiveDocument.ModelSpace.Item(oAppA.ActiveDocument.ModelSpace.Count - 1)
                If TypeOf oEnt Is AcadMLeader Then
                    resultado = DirectCast(oEnt, AcadMLeader)
                End If
            Catch ex As Exception
            End Try
            ''
            'If oCir IsNot Nothing Then oCir.Delete()
            Return resultado
        End Function
        Public Sub MLeader_Inserta()
            Dim objUtil As AcadUtility = oAppA.ActiveDocument.Utility
            Dim strPrompt1 As String = vbCrLf & "First Leader Point: "
            Dim strPrompt2 As String = vbCrLf & "Second Leader Point: "
            Dim dblPnts(5) As Double
            '
            With objUtil
                Dim pt1 As Object = Nothing
                Dim pt2 As Object = Nothing
                'Dim pointUCS As Object = Nothing
                '
                pt1 = .GetPoint(, strPrompt1)
                If pt1 Is Nothing Then Exit Sub
                'pointUCS = .TranslateCoordinates(varPnt, AcCoordinateSystem.acWorld, AcCoordinateSystem.acUCS, False)
                dblPnts(0) = pt1(0)  ' pointUCS(0)
                dblPnts(1) = pt1(1)  'pointUCS(1)
                dblPnts(2) = pt1(2)  'pointUCS(2)
                '
                pt2 = .GetPoint(pt1, strPrompt2)
                If pt2 Is Nothing Then Exit Sub
                'pointUCS = .TranslateCoordinates(varPnt, AcCoordinateSystem.acWorld, AcCoordinateSystem.acUCS, False)
                dblPnts(3) = pt2(0)  'pointUCS(0)
                dblPnts(4) = pt2(1)  'pointUCS(1)
                dblPnts(5) = pt2(2)  'pointUCS(2)

                'varPnt = .GetPoint(pointUCS, Prompt:="Insertion point for annotation:")
                'pointUCS = .TranslateCoordinates(varPnt, AcCoordinateSystem.acWorld, AcCoordinateSystem.acUCS, False)
                'dblPnts(6) = dblPnts(3) + 10
                'dblPnts(7) = dblPnts(4)
                'dblPnts(8) = dblPnts(5)
            End With
            '
            Dim oMl As AcadMLeader = Nothing
            '
            Dim index As Integer = 0
            If oAppA.ActiveDocument.ActiveSpace = AcActiveSpace.acModelSpace Then
                oMl = oAppA.ActiveDocument.ModelSpace.AddMLeader(dblPnts, index)

            Else
                oMl = oAppA.ActiveDocument.PaperSpace.AddMLeader(dblPnts, index)

            End If
            oMl.ContentType = AcMLeaderContentType.acBlockContent
            oMl.ContentBlockType = AcPredefBlockType.acBlockUserDefined
            'oMl.SetBlockAttributeValue(attdefId:=, Value:=)
            '
            'Dim r As Integer = oMl.AddLeader
            'dblPnts(4) += 8
            'Call oMl.AddLeaderLine(r, dblPnts)
        End Sub
        'Sub Example_MLeaderLine()
        '    Dim oML As AcadMLeader
        '    Dim points(0 To 5) As Double
        '    points(0) = 1 : points(1) = 1 : points(2) = 0
        '    points(3) = 4 : points(4) = 4 : points(5) = 0
        '    Dim i As Long
        'Set oML = ThisDrawing.ModelSpace.AddMLeader(points, i)

        'Dim r As Long
        '    r = oML.AddLeader()

        '    points(4) = 10
        '    Call oML.AddLeaderLine(r, points)

        '    MsgBox "LeaderCount = " & oML.LeaderCount
        'ZoomExtents
        'End Sub

        'AFLETA
        Public Function MleaderDameTodos_PorAtributo(queatributo As String, quevalor As String, Optional exacto As Boolean = False) As ArrayList
            Dim resultado As New ArrayList
            If oAppA Is Nothing Then _
        oAppA = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
            ''
            For Each oEnt As AcadEntity In oAppA.ActiveDocument.ModelSpace
                If TypeOf oEnt Is AcadMLeader Then
                    Dim oMl As AcadMLeader = oAppA.ActiveDocument.ObjectIdToObject(oEnt.ObjectID)

                    If exacto = True Then
                        If MLeaderBlock_DameValorAtributo(oMl, queatributo).ToUpper = quevalor.ToUpper Then
                            resultado.Add(oMl)
                        End If
                    Else
                        If MLeaderBlock_DameValorAtributo(oMl, queatributo).ToUpper.Contains(quevalor.ToUpper) Then
                            resultado.Add(oMl)
                        End If
                    End If
                    oMl = Nothing
                End If
            Next
            VaciaMemoria()
            Return resultado
        End Function
    End Class
End Namespace