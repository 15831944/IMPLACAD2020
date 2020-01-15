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
'
Partial Public Class Eventos
    Public Shared Sub Subscribe_AXEventM()
        AddHandler Autodesk.AutoCAD.Internal.Reactors.ApplicationEventManager.Instance.ApplicationDockLayoutChanged, AddressOf AXEventM_ApplicationDockLayoutChanged
        AddHandler Autodesk.AutoCAD.Internal.Reactors.ApplicationEventManager.Instance.ApplicationDocumentFrameChanged, AddressOf AXEventM_ApplicationDocumentFrameChanged
        AddHandler Autodesk.AutoCAD.Internal.Reactors.ApplicationEventManager.Instance.ApplicationMainWindowMoved, AddressOf AXEventM_ApplicationMainWindowMoved
        AddHandler Autodesk.AutoCAD.Internal.Reactors.ApplicationEventManager.Instance.ApplicationMainWindowSized, AddressOf AXEventM_ApplicationMainWindowSized
        AddHandler Autodesk.AutoCAD.Internal.Reactors.ApplicationEventManager.Instance.ApplicationMainWindowVisibleChanged, AddressOf AXEventM_ApplicationMainWindowVisibleChanged
    End Sub

    Public Shared Sub Unsubscribe_AXEventM()
        Try
            'RemoveHandler Autodesk.AutoCAD.Internal.Reactors.ApplicationEventManager.Instance.ApplicationDockLayoutChanged, AddressOf AXEventM_ApplicationDockLayoutChanged
            'RemoveHandler Autodesk.AutoCAD.Internal.Reactors.ApplicationEventManager.Instance.ApplicationDocumentFrameChanged, AddressOf AXEventM_ApplicationDocumentFrameChanged
            'RemoveHandler Autodesk.AutoCAD.Internal.Reactors.ApplicationEventManager.Instance.ApplicationMainWindowMoved, AddressOf AXEventM_ApplicationMainWindowMoved
            'RemoveHandler Autodesk.AutoCAD.Internal.Reactors.ApplicationEventManager.Instance.ApplicationMainWindowSized, AddressOf AXEventM_ApplicationMainWindowSized
            'RemoveHandler Autodesk.AutoCAD.Internal.Reactors.ApplicationEventManager.Instance.ApplicationMainWindowVisibleChanged, AddressOf AXEventM_ApplicationMainWindowVisibleChanged
        Catch ex As system.Exception
            Debug.Print(ex.ToString)
        End Try
    End Sub

    Public Shared Sub AXEventM_ApplicationDockLayoutChanged(sender As Object, e As EventArgs)
        'AXDoc.Editor.WriteMessage("AXEventM_ApplicationDockLayoutChanged")
        If logeventos Then PonLogEv("AXEventM_ApplicationDockLayoutChanged")
    End Sub

    Public Shared Sub AXEventM_ApplicationDocumentFrameChanged(sender As Object, e As EventArgs)
        'AXDoc.Editor.WriteMessage("AXEventM_ApplicationDocumentFrameChanged")
        If logeventos Then PonLogEv("AXEventM_ApplicationDocumentFrameChanged")
    End Sub

    Public Shared Sub AXEventM_ApplicationMainWindowMoved(sender As Object, e As EventArgs)
        'AXDoc.Editor.WriteMessage("AXEventM_ApplicationMainWindowMoved")
        If logeventos Then PonLogEv("AXEventM_ApplicationMainWindowMoved")
    End Sub

    Public Shared Sub AXEventM_ApplicationMainWindowSized(sender As Object, e As EventArgs)
        'AXDoc.Editor.WriteMessage("AXEventM_ApplicationMainWindowSized")
        If logeventos Then PonLogEv("AXEventM_ApplicationMainWindowSized")
    End Sub

    Public Shared Sub AXEventM_ApplicationMainWindowVisibleChanged(sender As Object, e As EventArgs)
        'AXDoc.Editor.WriteMessage("AXEventM_ApplicationMainWindowVisibleChanged")
        If logeventos Then PonLogEv("AXEventM_ApplicationMainWindowVisibleChanged")
    End Sub
End Class
