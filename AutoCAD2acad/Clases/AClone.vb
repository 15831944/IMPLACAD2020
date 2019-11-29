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
        Public Sub Clona_TodoDesdeDWGInsertando(quePlantilla As String, Optional borrabloque As Boolean = True, Optional cargarsiempre As Boolean = False)
            If System.IO.File.Exists(quePlantilla) = False Then Exit Sub
            '
            ' Poner el estilo Standard antes de insertar recursos
            Try
                Dim oDic As AcadDictionary = oAppA.ActiveDocument.Dictionaries.Item("ACAD_MLEADERSTYLE")
                Dim oMlS As AcadMLeaderStyle = CType(oAppA.ActiveDocument.ObjectIdToObject(oDic.Item(0).ObjectID), AcadMLeaderStyle)
                Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("CMLEADERSTYLE", oMlS.Name)
                oAppA.ActiveDocument.Regen(AcRegenType.acActiveViewport)
            Catch ex As Exception
                Debug.Print(ex.ToString)
            End Try
            '
            Dim existe As Boolean = False
            For Each oBl1 As AcadBlock In oAppA.ActiveDocument.Blocks
                If oBl1.Name = IO.Path.GetFileNameWithoutExtension(quePlantilla) Then
                    existe = True
                    Exit For
                End If
            Next
            '
            If cargarsiempre = True Or existe = False Then
                Dim ptIns(2) As Double
                ptIns(0) = 0 : ptIns(1) = 0 : ptIns(2) = 0
                Dim oBl As AcadBlockReference = oAppA.ActiveDocument.ModelSpace.InsertBlock(ptIns, quePlantilla, 1, 1, 1, 0)
                If borrabloque And oBl IsNot Nothing Then
                    Try
                        oBl.Delete()
                        'oAppA.ActiveDocument.ModelSpace.Item(oAppA.ActiveDocument.ObjectIdToObject(oBl.ObjectID)).Delete()
                    Catch ex As Exception
                        Debug.Print(ex.ToString)
                    End Try
                End If
            End If
            ' Poner el estilo ULMA
            Try
                Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("CMLEADERSTYLE", "ULMA")
                oAppA.ActiveDocument.Regen(AcRegenType.acActiveViewport)
            Catch ex As Exception
                Debug.Print(ex.ToString)
            End Try
        End Sub
        '
        Public Sub Clona_TodoDesdeDWGDatabase(quePlantilla As String)
            If System.IO.File.Exists(quePlantilla) = False Then Exit Sub
            '
            Dim doc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim db As Database = doc.Database
            Using tr As Transaction = db.TransactionManager.StartTransaction()
                Using sourceDb As Database = New Database(False, True)
                    sourceDb.ReadDwgFile(quePlantilla, System.IO.FileShare.Read, True, "")
                    db.Insert(quePlantilla, sourceDb, False)
                End Using                   '
                tr.Commit()
            End Using
            VaciaMemoria()
        End Sub
        Public Sub Clona_UnBloqueDesdeDWG(quePlantilla As String, queBloque As String)
            If System.IO.File.Exists(quePlantilla) = False Then Exit Sub
            '
            Dim doc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Using OpenDb As New Database(False, True)
                OpenDb.ReadDwgFile(quePlantilla, System.IO.FileShare.ReadWrite, True, "")
                Dim ids As New ObjectIdCollection()
                Using tr As Transaction = OpenDb.TransactionManager.StartTransaction()
                    'For example, Get the block by name "TEST"
                    Dim bt As BlockTable
                    bt = DirectCast(tr.GetObject(OpenDb.BlockTableId, OpenMode.ForRead), BlockTable)

                    If bt.Has(queBloque) Then
                        ids.Add(bt(queBloque))
                    End If

                    'if found, add the block
                    If ids.Count <> 0 Then
                        'get the current drawing database
                        Dim destdb As Database = doc.Database

                        Dim iMap As New IdMapping()
                        destdb.WblockCloneObjects(ids, destdb.BlockTableId, iMap, DuplicateRecordCloning.Replace, False)
                        destdb.Dispose()
                    End If
                    '
                    tr.Commit()
                End Using
            End Using
            VaciaMemoria()
        End Sub
        Public Sub Clona_TodosBloquesDesdeDWG(quePlantilla As String)
            If System.IO.File.Exists(quePlantilla) = False Then Exit Sub
            '
            Dim doc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Using OpenDb As New Database(False, True)
                OpenDb.ReadDwgFile(quePlantilla, System.IO.FileShare.ReadWrite, True, "")
                Dim ids As New ObjectIdCollection()
                Using tr As Transaction = OpenDb.TransactionManager.StartTransaction()
                    'For example, Get the block by name "TEST"
                    Dim bt As BlockTable
                    bt = DirectCast(tr.GetObject(OpenDb.BlockTableId, OpenMode.ForRead), BlockTable)
                    '
                    For Each queId As ObjectId In bt
                        ids.Add(queId)
                    Next
                    'if found, add the block
                    If ids.Count <> 0 Then
                        'get the current drawing database
                        Dim destdb As Database = doc.Database

                        Dim iMap As New IdMapping()
                        destdb.WblockCloneObjects(ids, destdb.BlockTableId, iMap, DuplicateRecordCloning.Replace, False)
                        destdb.Dispose()
                    End If
                    tr.Commit()
                End Using
            End Using
            VaciaMemoria()
        End Sub
        '
        Public Sub Clona_TodosEstilosDesdeDWG(quePlantilla As String)
            If System.IO.File.Exists(quePlantilla) = False Then Exit Sub
            '
            Dim doc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Using OpenDb As New Database(False, True)
                OpenDb.ReadDwgFile(quePlantilla, System.IO.FileShare.ReadWrite, True, "")
                Dim ids As New ObjectIdCollection()
                Using tr As Transaction = OpenDb.TransactionManager.StartTransaction()
                    'For example, Get the block by name "TEST"
                    Dim bt As TextStyleTable
                    bt = DirectCast(tr.GetObject(OpenDb.TextStyleTableId, OpenMode.ForRead), TextStyleTable)
                    '
                    For Each queId As ObjectId In bt
                        ids.Add(queId)
                    Next
                    'if found, add the block
                    If ids.Count <> 0 Then
                        'get the current drawing database
                        Dim destdb As Database = doc.Database

                        Dim iMap As New IdMapping()
                        destdb.WblockCloneObjects(ids, destdb.BlockTableId, iMap, DuplicateRecordCloning.Replace, False)
                        destdb.Dispose()
                    End If
                    tr.Commit()
                End Using
            End Using
            VaciaMemoria()
        End Sub
        Public Sub Clona_TodosEstilosMLeaderDesdeDWG(quePlantilla As String)
            If System.IO.File.Exists(quePlantilla) = False Then Exit Sub
            '
            Dim doc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Using OpenDb As New Database(False, True)
                OpenDb.ReadDwgFile(quePlantilla, System.IO.FileShare.ReadWrite, True, "")
                Dim ids As New ObjectIdCollection()
                Using tr As Transaction = OpenDb.TransactionManager.StartTransaction()
                    'For example, Get the block by name "TEST"
                    Dim bt As DBDictionary
                    bt = DirectCast(tr.GetObject(OpenDb.MLeaderStyleDictionaryId, OpenMode.ForRead), DBDictionary)
                    '
                    'Dim custEnum As IEnumerator = bt.GetEnumerator
                    For Each dEnt As DictionaryEntry In bt
                        ids.Add(bt.GetAt(dEnt.Key))
                    Next
                    'if found, add the block
                    If ids.Count <> 0 Then
                        'get the current drawing database
                        Dim destdb As Database = doc.Database
                        Dim iMap As New IdMapping()
                        destdb.WblockCloneObjects(ids, destdb.BlockTableId, iMap, DuplicateRecordCloning.Replace, False)
                        destdb.Dispose()
                    End If
                    tr.Commit()
                End Using
            End Using
            VaciaMemoria()
        End Sub
        Public Sub Clona_UnEstiloDesdeDWG(quePlantilla As String, mleaderStyle As String)
            Dim dm As DocumentCollection = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager
            Dim doc As Document = dm.MdiActiveDocument
            Dim db As Database = doc.Database
            'Dim ed As Editor = doc.Editor
            'Dim openFileDia As New Autodesk.AutoCAD.Windows.OpenFileDialog("Select Template", "", "dwt", "Select DWG 3D Template", OpenFileDialog.OpenFileDialogFlags.SearchPath)
            'Dim diagResult As System.Windows.Forms.DialogResult = openFileDia.ShowDialog()
            'exit the sub if the dialog result returns a response other than ok
            'If Not diagResult = Windows.Forms.DialogResult.OK Then Exit Sub
            'Dim quePlantilla As String = "C:\Civil 3D Project Templates\testing.dwt" 'openFileDia.Filename
            '
            If System.IO.File.Exists(quePlantilla) = False Then Exit Sub
            Dim dwt As Document = dm.Open(quePlantilla, False, "")
            Dim dwtdb As Database = dwt.Database
            '
            Dim styleId As ObjectId
            Dim styleColl As New ObjectIdCollection()
            Dim styleList As New List(Of ObjectId)
            Using targetLock As DocumentLock = doc.LockDocument()
                Using trans As Transaction = dwtdb.TransactionManager.StartOpenCloseTransaction
                    Try
                        Dim bt As BlockTable = trans.GetObject(dwtdb.BlockTableId, OpenMode.ForRead)
                        Dim dbDic As DBDictionary = trans.GetObject(dwtdb.MLeaderStyleDictionaryId, OpenMode.ForRead)
                        '
                        If dbDic.Contains(mleaderStyle) Then
                            styleId = dbDic.GetAt(mleaderStyle)
                            styleColl.Add(styleId)
                            Dim idMapping As New IdMapping
                            dwt.Database.WblockCloneObjects(styleColl, doc.Database.BlockTableId, idMapping, DuplicateRecordCloning.Replace, False)
                            idMapping.Dispose()
                        End If
                        dwt.CloseAndDiscard()
                        'targetLock.Dispose()
                        '
                    Catch ex As Exception
                        doc.Editor.WriteMessage(vbCrLf + "Error: " + ex.Message)
                    End Try
                End Using
            End Using
            VaciaMemoria()
        End Sub
        '
        'Public Sub testupdate()
        '    Dim dm As DocumentCollection = Application.DocumentManager
        '    Dim doc As Document = dm.MdiActiveDocument
        '    Try
        '        Dim civilDoc As Autodesk.Civil.ApplicationServices.CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument
        '        'Dim openFileDia As New Autodesk.AutoCAD.Windows.OpenFileDialog("Select Template", "", "dwt", "Select Civil 3D Template", OpenFileDialog.OpenFileDialogFlags.SearchPath)
        '        'Dim diagResult As System.Windows.Forms.DialogResult = openFileDia.ShowDialog()
        '        'exit the sub if the dialog result returns a response other than ok
        '        'If Not diagResult = Windows.Forms.DialogResult.OK Then Exit Sub
        '        Dim dwtName As String = "C:\Civil 3D Project Templates\testing.dwt" 'openFileDia.Filename
        '        If System.IO.File.Exists(dwtName) = False Then Exit Sub
        '        Dim dwt As Document = dm.Open(dwtName, False, "")
        '        Dim dwtCivilDoc As Autodesk.Civil.ApplicationServices.CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument
        '        Dim targetLock As DocumentLock = doc.LockDocument()
        '        Dim style As Autodesk.Civil.DatabaseServices.Styles.StyleBase
        '        Dim styleColl As New ObjectIdCollection()
        '        Dim styleList As New List(Of ObjectId)
        '        Dim tm As Autodesk.AutoCAD.DatabaseServices.TransactionManager = dwt.Database.TransactionManager
        '        Using trans As Transaction = tm.StartTransaction
        '            Dim bt As BlockTable = trans.GetObject(dwt.Database.BlockTableId, OpenMode.ForRead)
        '            For i As Integer = 0 To dwtCivilDoc.Styles.PointStyles.Count - 1
        '                style = trans.GetObject(dwtCivilDoc.Styles.PointStyles.Item(i), OpenMode.ForRead)
        '                styleColl.Add(style.ObjectId)
        '            Next
        '        End Using
        '        Dim idMapping As New IdMapping
        '        dwt.Database.WblockCloneObjects(styleColl, doc.Database.BlockTableId, idMapping, DuplicateRecordCloning.MangleName, False)
        '        idMapping.Dispose()
        '        dwt.CloseAndDiscard()
        '        targetLock.Dispose()
        '    Catch ex As Exception
        '        doc.Editor.WriteMessage(vbCrLf + "Error: " + ex.Message)
        '    End Try

        'End Sub
    End Class
End Namespace