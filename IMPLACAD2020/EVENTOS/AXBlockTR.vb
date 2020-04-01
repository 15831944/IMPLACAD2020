Imports Autodesk.AutoCAD.DatabaseServices

Partial Public Class Eventos

    Public Shared Sub Subscribe_AXBlockTR()
        AddHandler BlockTableRecord.BlockInsertionPoints, AddressOf AXBlockTR_BlockInsertionPoints
    End Sub

    Public Shared Sub Unsubscribe_AXBlockTR()
        Try
            'RemoveHandler BlockTableRecord.BlockInsertionPoints, AddressOf AXBlockTR_BlockInsertionPoints
        Catch ex As System.Exception
            Debug.Print(ex.ToString)
        End Try
    End Sub

    Public Shared Sub AXBlockTR_BlockInsertionPoints(sender As Object, e As BlockInsertionPointsEventArgs)
        'AXDoc.Editor.WriteMessage("AXBlockTR_BlockInsertionPoints")
        If logeventos Then PonLogEv("AXBlockTR_BlockInsertionPoints")
        If coneventos = False Then Exit Sub  ' Para que no haga nada después de un comando.
        'AXEditor.WriteMessage("BlockTableRecord_BlockInsertionPoints - Block Name: {0}\n", e.BlockTableRecord.Name)
    End Sub

End Class