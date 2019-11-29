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
Imports uau = UtilesAlberto.Utiles
Imports a2 = AutoCAD2acad.A2acad
'
Partial Public Class Eventos
    Public Shared Sub Subscribe_AXObj(ByRef AXObj As Autodesk.AutoCAD.DatabaseServices.DBObject)
        If AXObj Is Nothing Then Exit Sub
        '
        'AddHandler AXObj.Cancelled, AddressOf AXObj_Cancelled
        'AddHandler AXObj.Copied, AddressOf AXObj_Copied
        'AddHandler AXObj.Erased, AddressOf AXObj_Erased
        'AddHandler AXObj.Goodbye, AddressOf AXObj_Goodbye
        'AddHandler AXObj.Modified, AddressOf AXObj_Modified
        'AddHandler AXObj.ModifiedXData, AddressOf AXObj_ModifiedXData
        'AddHandler AXObj.ModifyUndone, AddressOf AXObj_ModifyUndone
        'AddHandler AXObj.ObjectClosed, AddressOf AXObj_ObjectClosed
        'AddHandler AXObj.OpenedForModify, AddressOf AXObj_OpenedForModify
        'AddHandler AXObj.Reappended, AddressOf AXObj_Reappended
        'AddHandler AXObj.SubObjectModified, AddressOf AXObj_SubObjectModified
        'AddHandler AXObj.Unappended, AddressOf AXObj_Unappended
    End Sub
    Public Shared Sub Unsubscribe_AXObj(ByRef AXObj As DBObject)
        Try
            If AXObj Is Nothing OrElse AXObj.IsDisposed = True Then Exit Sub
            '
            'RemoveHandler AXObj.Cancelled, AddressOf AXObj_Cancelled
            'RemoveHandler AXObj.Copied, AddressOf AXObj_Copied
            'RemoveHandler AXObj.Erased, AddressOf AXObj_Erased
            'RemoveHandler AXObj.Goodbye, AddressOf AXObj_Goodbye
            'RemoveHandler AXObj.Modified, AddressOf AXObj_Modified
            'RemoveHandler AXObj.ModifiedXData, AddressOf AXObj_ModifiedXData
            'RemoveHandler AXObj.ModifyUndone, AddressOf AXObj_ModifyUndone
            'RemoveHandler AXObj.ObjectClosed, AddressOf AXObj_ObjectClosed
            'RemoveHandler AXObj.OpenedForModify, AddressOf AXObj_OpenedForModify
            'RemoveHandler AXObj.Reappended, AddressOf AXObj_Reappended
            'RemoveHandler AXObj.SubObjectModified, AddressOf AXObj_SubObjectModified
            'RemoveHandler AXObj.Unappended, AddressOf AXObj_Unappended
        Catch ex As system.Exception
            Debug.Print(ex.ToString)
        End Try
    End Sub
    Public Shared Sub AXObj_Cancelled(sender As Object, e As EventArgs)
        'AXDoc.Editor.WriteMessage("AXObj_Cancelled")
        If logeventos Then PonLogEv("AXObj_Cancelled")
    End Sub

    Public Shared Sub AXObj_Copied(sender As Object, e As ObjectEventArgs)
        'AXDoc.Editor.WriteMessage("AXObj_Copied")
        If logeventos Then PonLogEv("AXObj_Copied")
    End Sub

    Public Shared Sub AXObj_Erased(sender As Object, e As ObjectErasedEventArgs)
        'AXDoc.Editor.WriteMessage("AXObj_Erased")
        If logeventos Then PonLogEv("AXObj_Erased")
        'If e.DBObject IsNot Nothing AndAlso e.Erased = False Then
        '    'Debug.Print(sender.Name)
        '    Unsubscribe_EvObjDB(sender)
        '    Unsubscribre_EvObjCOM(CType(sender, DBObject).AcadObject)
        'End If
    End Sub

    Public Shared Sub AXObj_Goodbye(sender As Object, e As EventArgs)
        'AXDoc.Editor.WriteMessage("AXObj_Goodbye")
        If logeventos Then PonLogEv("AXObj_Goodbye")
        'Try
        '    Dim colPurge As New ObjectIdCollection
        '    Dim oDbO As DBObject = CType(sender, DBObject)
        '    If oDbO.IsErased Then
        '        colPurge.Add(oDbO.ObjectId)
        '        EvDocM.CurrentDocument.Database.Purge(colPurge)
        '    End If
        '    colPurge = Nothing
        '    oDbO = Nothing

        'Catch ex As System.Exception

        'End Try
    End Sub

    Public Shared Sub AXObj_Modified(sender As Object, e As EventArgs)
        'AXDoc.Editor.WriteMessage("AXObj_Modified")
        If logeventos Then PonLogEv("AXObj_Modified")
        'Dim dbobj As DBObject = CType(sender, DBObject)
        'If TypeOf dbobj Is Autodesk.AutoCAD.DatabaseServices.BlockReference Then
        '    '        Using Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
        '    '            'modTavil.AcadBlockReference_Modified(CType(queObj, Autodesk.AutoCAD.Interop.Common.AcadBlockReference))
        '    '            Try
        '    '                Dim queTipo As String = clsA.XLeeDato(CType(queObj, AcadObject), "tipo")
        '    '                'MsgBox(oBlr.EffectiveName & " Modificado")
        '    '                If colIds Is Nothing Then colIds = New List(Of Long)
        '    '                If colHan Is Nothing Then colHan = New List(Of String)
        '    '                If colIds.Contains(CType(queObj, AcadBlockReference).ObjectID) = False And queTipo = "cinta" Then
        '    '                    colIds.Add(CType(queObj, AcadBlockReference).ObjectID)
        '    '                    'colHan.Add(CType(queObj, AcadBlockReference).Handle)
        '    '                    ' Si vamos a modificar algo, poner app_procesointer = true (Para que no active eventos)
        '    '                    app_procesointerno = True
        '    '                    'clsA.SeleccionaPorHandle(Ev.EvApp.ActiveDocument, queObj, "_UPDATEFIELD")
        '    '                End If
        '    '            Catch ex As Exception
        '    '                Debug.Print(ex.ToString)
        '    '            End Try
        '    '        End Using
        'ElseIf TypeOf dbobj Is Circle Then
        '    EvDocM.CurrentDocument.Editor.WriteMessage(vbLf & "ActiveX Radio: " & CType(dbobj, Circle).Radius)
        'End If
    End Sub

    Public Shared Sub AXObj_ModifiedXData(sender As Object, e As EventArgs)
        'AXDoc.Editor.WriteMessage("AXObj_ModifiedXData")
        If logeventos Then PonLogEv("AXObj_ModifiedXData")
    End Sub

    Public Shared Sub AXObj_ModifyUndone(sender As Object, e As EventArgs)
        'AXDoc.Editor.WriteMessage("AXObj_ModifyUndone")
        If logeventos Then PonLogEv("AXObj_ModifyUndone")
    End Sub

    Public Shared Sub AXObj_ObjectClosed(sender As Object, e As ObjectClosedEventArgs)
        'AXDoc.Editor.WriteMessage("AXObj_ObjectClosed")
        If logeventos Then PonLogEv("AXObj_ObjectClosed")
        If sender Is Nothing Then Exit Sub
        'Debug.Print("Hola")
    End Sub

    Public Shared Sub AXObj_OpenedForModify(sender As Object, e As EventArgs)
        'AXDoc.Editor.WriteMessage("AXObj_OpenedForModify")
        If logeventos Then PonLogEv("AXObj_OpenedForModify")
    End Sub

    Public Shared Sub AXObj_Reappended(sender As Object, e As EventArgs)
        'AXDoc.Editor.WriteMessage("AXObj_Reappended")
        If logeventos Then PonLogEv("AXObj_Reappended")
    End Sub

    Public Shared Sub AXObj_SubObjectModified(sender As Object, e As ObjectEventArgs)
        'AXDoc.Editor.WriteMessage("AXObj_SubObjectModified")
        If logeventos Then PonLogEv("AXObj_SubObjectModified")
    End Sub

    Public Shared Sub AXObj_Unappended(sender As Object, e As EventArgs)
        'AXDoc.Editor.WriteMessage("AXObj_Unappended")
        If logeventos Then PonLogEv("AXObj_Unappended")
        'If sender IsNot Nothing Then
        '    Unsubscribe_EvObjDB(sender)
        '    Unsubscribre_EvObjCOM(CType(sender, DBObject).AcadObject)
        'End If
    End Sub
End Class
'
'Cancelled          Se activa cuando la apertura del objeto es texto cancelado.
'Copied             Activado después de clonar el objeto.
'Erased             Se activa cuando el objeto se marca para borrar o se borra.
'Goodbye            Se activa cuando el objeto está a punto de eliminarse de la memoria porque su base de datos asociada se está destruyendo.
'Modified           Se activa cuando se modifica el objeto.
'ModifiedXData      Se activa cuando se modifica el XData adjunto al objeto.
'ModifyUndone       Se activa cuando se deshacen los cambios anteriores al objeto.
'ObjectClosed       Se activa cuando el objeto está cerrado.
'OpenedForModify    Se activa antes de que se modifique el objeto.
'Reappended         Se activa cuando el objeto se elimina de la base de datos después de una operación Deshacer Y luego se vuelve a agregar con una operación Rehacer..
'SubObjectModified  Se activa cuando se modifica un subobjeto del objeto.
'Unappended         Se activa cuando el objeto se elimina de la base de datos después de una operación Deshacer.

'Los siguientes son algunos de los eventos utilizados para responder a los cambios de objetos a nivel de base de datos
'ObjectAppended         Se activa cuando se agrega un objeto a una base de datos.
'ObjectErased           Se activa cuando un objeto se borra o borra de una base de datos.
'ObjectModified         Se activa cuando un objeto ha sido modificado.
'ObjectOpenedForModify  Se activa antes de que se modifique un objeto.
'ObjectReappended       Se activa cuando un objeto se elimina de una base de datos después de una operación Deshacer Y luego se vuelve a agregar con una operación Rehacer.
'ObjectUnappended       Se activa cuando un objeto se elimina de una base de datos después de una operación Deshacer.


