Imports System
Imports System.Text
Imports System.Linq
Imports System.Xml
Imports System.Reflection
Imports System.ComponentModel
Imports System.Collections
Imports System.Collections.Generic
Imports System.Windows
Imports System.Windows.Media.Imaging
Imports System.Windows.Forms
Imports System.IO

Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Interop
Imports Autodesk.AutoCAD.Interop.Common
Imports Autodesk.AutoCAD.Geometry
Imports AXApp = Autodesk.AutoCAD.ApplicationServices.Application
Imports AXDoc = Autodesk.AutoCAD.ApplicationServices.Document
Imports AXWin = Autodesk.AutoCAD.Windows
Imports a2 = AutoCAD2acad.A2acad

Partial Public Class Eventos
    Public Shared Sub Subscribe_AXBlockTR()
        AddHandler BlockTableRecord.BlockInsertionPoints, AddressOf AXBlockTR_BlockInsertionPoints
    End Sub
    Public Shared Sub Unsubscribe_AXBlockTR()
        Try
            'RemoveHandler BlockTableRecord.BlockInsertionPoints, AddressOf AXBlockTR_BlockInsertionPoints
        Catch ex As system.Exception
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
