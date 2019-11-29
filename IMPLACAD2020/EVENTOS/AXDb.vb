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

Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Interop
Imports Autodesk.AutoCAD.Interop.Common
Imports Autodesk.AutoCAD.Geometry
Imports AcadApplication = Autodesk.AutoCAD.ApplicationServices.Application
Imports AcadDocument = Autodesk.AutoCAD.ApplicationServices.Document
Imports AcadWindows = Autodesk.AutoCAD.Windows
Imports uau = UtilesAlberto.Utiles
Imports a2 = AutoCAD2acad.A2acad
'
Partial Public Class Eventos
    Public Shared Sub Subscribe_AXDB() 'docDb As Autodesk.AutoCAD.DatabaseServices.Database)
        If AXDb() Is Nothing Then Exit Sub
        AddHandler AXDb.AbortDxfOut, AddressOf AXDB_AbortDxfOut
        AddHandler AXDb.AbortSave, AddressOf AXDB_AbortSave
        AddHandler AXDb.BeginDeepClone, AddressOf AXDB_BeginDeepClone
        AddHandler AXDb.BeginDeepCloneTranslation, AddressOf AXDB_BeginDeepCloneTranslation
        AddHandler AXDb.BeginDxfIn, AddressOf AXDB_BeginDxfIn
        AddHandler AXDb.BeginDxfOut, AddressOf AXDB_BeginDxfOut
        AddHandler AXDb.BeginInsert, AddressOf AXDB_BeginInsert
        AddHandler AXDb.BeginSave, AddressOf AXDB_BeginSave
        AddHandler AXDb.BeginWblockBlock, AddressOf AXDB_BeginWblockBlock
        AddHandler AXDb.BeginWblockEntireDatabase, AddressOf AXDB_BeginWblockEntireDatabase
        AddHandler AXDb.BeginWblockObjects, AddressOf AXDB_BeginWblockObjects
        AddHandler AXDb.BeginWblockSelectedObjects, AddressOf AXDB_BeginWblockSelectedObjects
        AddHandler Autodesk.AutoCAD.DatabaseServices.Database.DatabaseConstructed, AddressOf AXDB_DatabaseConstructed
        AddHandler AXDb.DatabaseToBeDestroyed, AddressOf AXDB_DatabaseToBeDestroyed
        AddHandler AXDb.DeepCloneAborted, AddressOf AXDB_DeepCloneAborted
        AddHandler AXDb.DeepCloneEnded, AddressOf AXDB_DeepCloneEnded
        AddHandler AXDb.Disposed, AddressOf AXDB_Disposed
        AddHandler AXDb.DwgFileOpened, AddressOf AXDB_DwgFileOpened
        AddHandler AXDb.DxfInComplete, AddressOf AXDB_DxfInComplete
        AddHandler AXDb.DxfOutComplete, AddressOf AXDB_DxfOutComplete
        AddHandler AXDb.InitialDwgFileOpenComplete, AddressOf AXDB_InitialDwgFileOpenComplete
        AddHandler AXDb.InsertAborted, AddressOf AXDB_InsertAborted
        AddHandler AXDb.InsertEnded, AddressOf AXDB_InsertEnded
        AddHandler AXDb.InsertMappingAvailable, AddressOf AXDB_InsertMappingAvailable
        AddHandler AXDb.ObjectAppended, AddressOf AXDB_ObjectAppended
        AddHandler AXDb.ObjectErased, AddressOf AXDB_ObjectErased
        AddHandler AXDb.ObjectModified, AddressOf AXDB_ObjectModified
        AddHandler AXDb.ObjectOpenedForModify, AddressOf AXDB_ObjectOpenedForModify
        AddHandler AXDb.ObjectReappended, AddressOf AXDB_ObjectReappended
        AddHandler AXDb.ObjectUnappended, AddressOf AXDB_ObjectUnappended
        AddHandler AXDb.PartialOpenNotice, AddressOf AXDB_PartialOpenNotice
        AddHandler AXDb.ProxyResurrectionCompleted, AddressOf AXDB_ProxyResurrectionCompleted
        AddHandler AXDb.SaveComplete, AddressOf AXDB_SaveComplete
        AddHandler AXDb.SystemVariableChanged, AddressOf AXDB_SystemVariableChanged
        AddHandler AXDb.SystemVariableWillChange, AddressOf AXDB_SystemVariableWillChange
        AddHandler AXDb.WblockAborted, AddressOf AXDB_WblockAborted
        AddHandler AXDb.WblockEnded, AddressOf AXDB_WblockEnded
        AddHandler AXDb.WblockMappingAvailable, AddressOf AXDB_WblockMappingAvailable
        AddHandler AXDb.WblockNotice, AddressOf AXDB_WblockNotice
        AddHandler Autodesk.AutoCAD.DatabaseServices.Database.XrefAttachAborted, AddressOf AXDB_XrefAttachAborted
        AddHandler AXDb.XrefAttachEnded, AddressOf AXDB_XrefAttachEnded
        AddHandler AXDb.XrefBeginAttached, AddressOf AXDB_XrefBeginAttached
        AddHandler AXDb.XrefBeginOtherAttached, AddressOf AXDB_XrefBeginOtherAttached
        AddHandler AXDb.XrefBeginRestore, AddressOf AXDB_XrefBeginRestore
        AddHandler AXDb.XrefComandeered, AddressOf AXDB_XrefComandeered
        AddHandler AXDb.XrefPreXrefLockFile, AddressOf AXDB_XrefPreXrefLockFile
        AddHandler AXDb.XrefRedirected, AddressOf AXDB_XrefRedirected
        AddHandler AXDb.XrefRestoreAborted, AddressOf AXDB_XrefRestoreAborted
        AddHandler AXDb.XrefRestoreEnded, AddressOf AXDB_XrefRestoreEnded
        AddHandler AXDb.XrefSubCommandAborted, AddressOf AXDB_XrefSubCommandAborted
        AddHandler AXDb.XrefSubCommandEnd, AddressOf AXDB_XrefSubCommandEnd
        AddHandler AXDb.XrefSubCommandStart, AddressOf AXDB_XrefSubCommandStart
    End Sub

    Public Shared Sub Unsubscribe_AXDB()
        Try
            If AXDb() Is Nothing Then Exit Sub
            'RemoveHandler AXDb.AbortDxfIn, AddressOf AXDB_AbortDxfIn
            'RemoveHandler AXDb.AbortDxfOut, AddressOf AXDB_AbortDxfOut
            'RemoveHandler AXDb.AbortSave, AddressOf AXDB_AbortSave
            'RemoveHandler AXDb.BeginDeepClone, AddressOf AXDB_BeginDeepClone
            'RemoveHandler AXDb.BeginDeepCloneTranslation, AddressOf AXDB_BeginDeepCloneTranslation
            'RemoveHandler AXDb.BeginDxfIn, AddressOf AXDB_BeginDxfIn
            'RemoveHandler AXDb.BeginDxfOut, AddressOf AXDB_BeginDxfOut
            'RemoveHandler AXDb.BeginInsert, AddressOf AXDB_BeginInsert
            'RemoveHandler AXDb.BeginSave, AddressOf AXDB_BeginSave
            'RemoveHandler AXDb.BeginWblockBlock, AddressOf AXDB_BeginWblockBlock
            'RemoveHandler AXDb.BeginWblockEntireDatabase, AddressOf AXDB_BeginWblockEntireDatabase
            'RemoveHandler AXDb.BeginWblockObjects, AddressOf AXDB_BeginWblockObjects
            'RemoveHandler AXDb.BeginWblockSelectedObjects, AddressOf AXDB_BeginWblockSelectedObjects
            'RemoveHandler Autodesk.AutoCAD.DatabaseServices.Database.DatabaseConstructed, AddressOf AXDB_DatabaseConstructed
            'RemoveHandler AXDb.DatabaseToBeDestroyed, AddressOf AXDB_DatabaseToBeDestroyed
            'RemoveHandler AXDb.DeepCloneAborted, AddressOf AXDB_DeepCloneAborted
            'RemoveHandler AXDb.DeepCloneEnded, AddressOf AXDB_DeepCloneEnded
            'RemoveHandler AXDb.Disposed, AddressOf AXDB_Disposed
            'RemoveHandler AXDb.DwgFileOpened, AddressOf AXDB_DwgFileOpened
            'RemoveHandler AXDb.DxfInComplete, AddressOf AXDB_DxfInComplete
            'RemoveHandler AXDb.DxfOutComplete, AddressOf AXDB_DxfOutComplete
            'RemoveHandler AXDb.InitialDwgFileOpenComplete, AddressOf AXDB_InitialDwgFileOpenComplete
            'RemoveHandler AXDb.InsertAborted, AddressOf AXDB_InsertAborted
            'RemoveHandler AXDb.InsertEnded, AddressOf AXDB_InsertEnded
            'RemoveHandler AXDb.InsertMappingAvailable, AddressOf AXDB_InsertMappingAvailable
            'RemoveHandler AXDb.ObjectAppended, AddressOf AXDB_ObjectAppended
            'RemoveHandler AXDb.ObjectErased, AddressOf AXDB_ObjectErased
            'RemoveHandler AXDb.ObjectModified, AddressOf AXDB_ObjectModified
            'RemoveHandler AXDb.ObjectOpenedForModify, AddressOf AXDB_ObjectOpenedForModify
            'RemoveHandler AXDb.ObjectReappended, AddressOf AXDB_ObjectReappended
            'RemoveHandler AXDb.ObjectUnappended, AddressOf AXDB_ObjectUnappended
            'RemoveHandler AXDb.PartialOpenNotice, AddressOf AXDB_PartialOpenNotice
            'RemoveHandler AXDb.ProxyResurrectionCompleted, AddressOf AXDB_ProxyResurrectionCompleted
            'RemoveHandler AXDb.SaveComplete, AddressOf AXDB_SaveComplete
            'RemoveHandler AXDb.SystemVariableChanged, AddressOf AXDB_SystemVariableChanged
            'RemoveHandler AXDb.SystemVariableWillChange, AddressOf AXDB_SystemVariableWillChange
            'RemoveHandler AXDb.WblockAborted, AddressOf AXDB_WblockAborted
            'RemoveHandler AXDb.WblockEnded, AddressOf AXDB_WblockEnded
            'RemoveHandler AXDb.WblockMappingAvailable, AddressOf AXDB_WblockMappingAvailable
            'RemoveHandler AXDb.WblockNotice, AddressOf AXDB_WblockNotice
            'RemoveHandler Autodesk.AutoCAD.DatabaseServices.Database.XrefAttachAborted, AddressOf AXDB_XrefAttachAborted
            'RemoveHandler AXDb.XrefAttachEnded, AddressOf AXDB_XrefAttachEnded
            'RemoveHandler AXDb.XrefBeginAttached, AddressOf AXDB_XrefBeginAttached
            'RemoveHandler AXDb.XrefBeginOtherAttached, AddressOf AXDB_XrefBeginOtherAttached
            'RemoveHandler AXDb.XrefBeginRestore, AddressOf AXDB_XrefBeginRestore
            'RemoveHandler AXDb.XrefComandeered, AddressOf AXDB_XrefComandeered
            'RemoveHandler AXDb.XrefPreXrefLockFile, AddressOf AXDB_XrefPreXrefLockFile
            'RemoveHandler AXDb.XrefRedirected, AddressOf AXDB_XrefRedirected
            'RemoveHandler AXDb.XrefRestoreAborted, AddressOf AXDB_XrefRestoreAborted
            'RemoveHandler AXDb.XrefRestoreEnded, AddressOf AXDB_XrefRestoreEnded
            'RemoveHandler AXDb.XrefSubCommandAborted, AddressOf AXDB_XrefSubCommandAborted
            'RemoveHandler AXDb.XrefSubCommandEnd, AddressOf AXDB_XrefSubCommandEnd
            'RemoveHandler AXDb.XrefSubCommandStart, AddressOf AXDB_XrefSubCommandStart
        Catch ex As system.Exception
            Debug.Print(ex.ToString)
        End Try
    End Sub
    Public Shared Sub AXDB_AbortDxfIn(sender As Object, e As EventArgs)
        'AXDoc.Editor.WriteMessage("AXDB_AbortDxfIn")
        If logeventos Then PonLogEv("AXDB_AbortDxfIn")
    End Sub

    Public Shared Sub AXDB_AbortDxfOut(sender As Object, e As EventArgs)
        'AXDoc.Editor.WriteMessage("AXDB_AbortDxfOut")
        If logeventos Then PonLogEv("AXDB_AbortDxfOut")
    End Sub

    Public Shared Sub AXDB_AbortSave(sender As Object, e As EventArgs)
        'AXDoc.Editor.WriteMessage("AXDB_AbortSave")
        If logeventos Then PonLogEv("AXDB_AbortSave")
    End Sub

    Public Shared Sub AXDB_BeginDeepClone(sender As Object, e As IdMappingEventArgs)
        'AXDoc.Editor.WriteMessage("AXDB_BeginDeepClone")
        If logeventos Then PonLogEv("AXDB_BeginDeepClone")
    End Sub

    Public Shared Sub AXDB_BeginDeepCloneTranslation(sender As Object, e As IdMappingEventArgs)
        'AXDoc.Editor.WriteMessage("AXDB_BeginDeepCloneTranslation")
        If logeventos Then PonLogEv("AXDB_BeginDeepCloneTranslation")
    End Sub

    Public Shared Sub AXDB_BeginDxfIn(sender As Object, e As EventArgs)
        'AXDoc.Editor.WriteMessage("AXDB_BeginDxfIn")
        If logeventos Then PonLogEv("AXDB_BeginDxfIn")
    End Sub

    Public Shared Sub AXDB_BeginDxfOut(sender As Object, e As EventArgs)
        'AXDoc.Editor.WriteMessage("AXDB_BeginDxfOut")
        If logeventos Then PonLogEv("AXDB_BeginDxfOut")
    End Sub

    Public Shared Sub AXDB_BeginInsert(sender As Object, e As BeginInsertEventArgs)
        'AXDoc.Editor.WriteMessage("AXDB_BeginInsert")
        If logeventos Then PonLogEv("AXDB_BeginInsert")
    End Sub

    Public Shared Sub AXDB_BeginSave(sender As Object, e As DatabaseIOEventArgs)
        'AXDoc.Editor.WriteMessage("AXDB_BeginSave")
        If logeventos Then PonLogEv("AXDB_BeginSave")
    End Sub

    Public Shared Sub AXDB_BeginWblockBlock(sender As Object, e As BeginWblockBlockEventArgs)
        'AXDoc.Editor.WriteMessage("AXDB_BeginWblockBlock")
        If logeventos Then PonLogEv("AXDB_BeginWblockBlock")
    End Sub

    Public Shared Sub AXDB_BeginWblockEntireDatabase(sender As Object, e As BeginWblockEntireDatabaseEventArgs)
        'AXDoc.Editor.WriteMessage("AXDB_BeginWblockEntireDatabase")
        If logeventos Then PonLogEv("AXDB_BeginWblockEntireDatabase")
    End Sub

    Public Shared Sub AXDB_BeginWblockObjects(sender As Object, e As BeginWblockObjectsEventArgs)
        'AXDoc.Editor.WriteMessage("AXDB_BeginWblockObjects")
        If logeventos Then PonLogEv("AXDB_BeginWblockObjects")
    End Sub

    Public Shared Sub AXDB_BeginWblockSelectedObjects(sender As Object, e As BeginWblockSelectedObjectsEventArgs)
        'AXDoc.Editor.WriteMessage("AXDB_BeginWblockSelectedObjects")
        If logeventos Then PonLogEv("AXDB_BeginWblockSelectedObjects")
    End Sub

    Public Shared Sub AXDB_DatabaseConstructed(sender As Object, e As EventArgs)
        'AXDoc.Editor.WriteMessage("AXDB_DatabaseConstructed")
        If logeventos Then PonLogEv("AXDB_DatabaseConstructed")
    End Sub

    Public Shared Sub AXDB_DatabaseToBeDestroyed(sender As Object, e As EventArgs)
        'AXDoc.Editor.WriteMessage("AXDB_DatabaseToBeDestroyed")
        If logeventos Then PonLogEv("AXDB_DatabaseToBeDestroyed")
    End Sub

    Public Shared Sub AXDB_DeepCloneAborted(sender As Object, e As EventArgs)
        'AXDoc.Editor.WriteMessage("AXDB_DeepCloneAborted")
        If logeventos Then PonLogEv("AXDB_DeepCloneAborted")
    End Sub

    Public Shared Sub AXDB_DeepCloneEnded(sender As Object, e As EventArgs)
        'AXDoc.Editor.WriteMessage("AXDB_DeepCloneEnded")
        If logeventos Then PonLogEv("AXDB_DeepCloneEnded")
    End Sub

    Public Shared Sub AXDB_Disposed(sender As Object, e As EventArgs)
        'AXDoc.Editor.WriteMessage("AXDB_Disposed")
        If logeventos Then PonLogEv("AXDB_Disposed")
    End Sub

    Public Shared Sub AXDB_DwgFileOpened(sender As Object, e As DatabaseIOEventArgs)
        'AXDoc.Editor.WriteMessage("AXDB_DwgFileOpened")
        If logeventos Then PonLogEv("AXDB_DwgFileOpened")
    End Sub

    Public Shared Sub AXDB_DxfInComplete(sender As Object, e As EventArgs)
        'AXDoc.Editor.WriteMessage("AXDB_DxfInComplete")
        If logeventos Then PonLogEv("AXDB_DxfInComplete")
    End Sub

    Public Shared Sub AXDB_DxfOutComplete(sender As Object, e As EventArgs)
        'AXDoc.Editor.WriteMessage("AXDB_DxfOutComplete")
        If logeventos Then PonLogEv("AXDB_DxfOutComplete")
    End Sub

    Public Shared Sub AXDB_InitialDwgFileOpenComplete(sender As Object, e As EventArgs)
        'AXDoc.Editor.WriteMessage("AXDB_InitialDwgFileOpenComplete")
        If logeventos Then PonLogEv("AXDB_InitialDwgFileOpenComplete")
    End Sub

    Public Shared Sub AXDB_InsertAborted(sender As Object, e As EventArgs)
        'AXDoc.Editor.WriteMessage("AXDB_InsertAborted")
        If logeventos Then PonLogEv("AXDB_InsertAborted")
    End Sub

    Public Shared Sub AXDB_InsertEnded(sender As Object, e As EventArgs)
        'AXDoc.Editor.WriteMessage("AXDB_InsertEnded")
        If logeventos Then PonLogEv("AXDB_InsertEnded")
        If coneventos = False Then Exit Sub  ' Para que no haga nada después de un comando.
    End Sub

    Public Shared Sub AXDB_InsertMappingAvailable(sender As Object, e As IdMappingEventArgs)
        'AXDoc.Editor.WriteMessage("AXDB_InsertMappingAvailable")
        If logeventos Then PonLogEv("AXDB_InsertMappingAvailable")
    End Sub

    Public Shared Sub AXDB_ObjectAppended(sender As Object, e As ObjectEventArgs)
        'AXDoc.Editor.WriteMessage("AXDB_ObjectAppended")
        If logeventos Then PonLogEv("AXDB_ObjectAppended")
        If (app_procesointerno = True) Then Exit Sub
        If coneventos = False Then Exit Sub  ' Para que no haga nada después de un comando.
    End Sub

    Public Shared Sub AXDB_ObjectErased(sender As Object, e As ObjectErasedEventArgs)
        'AXDoc.Editor.WriteMessage("AXDB_ObjectErased")
        If logeventos Then PonLogEv("AXDB_ObjectErased")
        If (app_procesointerno = True) Then Exit Sub
        If coneventos = False Then Exit Sub  ' Para que no haga nada después de un comando.

        app_procesointerno = False
    End Sub

    Public Shared Sub AXDB_ObjectModified(sender As Object, e As ObjectEventArgs)
        'AXDoc.Editor.WriteMessage("AXDB_ObjectModified")
        If logeventos Then PonLogEv("AXDB_ObjectModified")
        If (app_procesointerno = True) Then Exit Sub
        If coneventos = False Then Exit Sub  ' Para que no haga nada después de un comando.


        app_procesointerno = False
    End Sub

    Public Shared Sub AXDB_ObjectOpenedForModify(sender As Object, e As ObjectEventArgs)
        'AXDoc.Editor.WriteMessage("AXDB_ObjectOpenedForModify")
        If logeventos Then PonLogEv("AXDB_ObjectOpenedForModify")
    End Sub

    Public Shared Sub AXDB_ObjectReappended(sender As Object, e As ObjectEventArgs)
        'AXDoc.Editor.WriteMessage("AXDB_ObjectReappended")
        If logeventos Then PonLogEv("AXDB_ObjectReappended")
    End Sub

    Public Shared Sub AXDB_ObjectUnappended(sender As Object, e As ObjectEventArgs)
        'AXDoc.Editor.WriteMessage("AXDB_ObjectUnappended")
        If logeventos Then PonLogEv("AXDB_ObjectUnappended")

    End Sub

    Public Shared Sub AXDB_PartialOpenNotice(sender As Object, e As EventArgs)
        'AXDoc.Editor.WriteMessage("AXDB_PartialOpenNotice")
        If logeventos Then PonLogEv("AXDB_PartialOpenNotice")
    End Sub

    Public Shared Sub AXDB_ProxyResurrectionCompleted(sender As Object, e As ProxyResurrectionCompletedEventArgs)
        'AXDoc.Editor.WriteMessage("AXDB_ProxyResurrectionCompleted")
        If logeventos Then PonLogEv("AXDB_ProxyResurrectionCompleted")
    End Sub

    Public Shared Sub AXDB_SaveComplete(sender As Object, e As DatabaseIOEventArgs)
        'AXDoc.Editor.WriteMessage("AXDB_SaveComplete")
        If logeventos Then PonLogEv("AXDB_SaveComplete")
    End Sub

    Public Shared Sub AXDB_SystemVariableChanged(sender As Object, e As Autodesk.AutoCAD.DatabaseServices.SystemVariableChangedEventArgs)
        'AXDoc.Editor.WriteMessage("AXDB_SystemVariableChanged")
        If logeventos Then PonLogEv("AXDB_SystemVariableChanged")
    End Sub

    Public Shared Sub AXDB_SystemVariableWillChange(sender As Object, e As Autodesk.AutoCAD.DatabaseServices.SystemVariableChangingEventArgs)
        'AXDoc.Editor.WriteMessage("AXDB_SystemVariableWillChange")
        If logeventos Then PonLogEv("AXDB_SystemVariableWillChange")
    End Sub

    Public Shared Sub AXDB_WblockAborted(sender As Object, e As EventArgs)
        'AXDoc.Editor.WriteMessage("AXDB_WblockAborted")
        If logeventos Then PonLogEv("AXDB_WblockAborted")
    End Sub

    Public Shared Sub AXDB_WblockEnded(sender As Object, e As EventArgs)
        'AXDoc.Editor.WriteMessage("AXDB_WblockEnded")
        If logeventos Then PonLogEv("AXDB_WblockEnded")
    End Sub

    Public Shared Sub AXDB_WblockMappingAvailable(sender As Object, e As IdMappingEventArgs)
        'AXDoc.Editor.WriteMessage("AXDB_WblockMappingAvailable")
        If logeventos Then PonLogEv("AXDB_WblockMappingAvailable")
    End Sub

    Public Shared Sub AXDB_WblockNotice(sender As Object, e As WblockNoticeEventArgs)
        'AXDoc.Editor.WriteMessage("AXDB_WblockNotice")
        If logeventos Then PonLogEv("AXDB_WblockNotice")
    End Sub

    Public Shared Sub AXDB_XrefAttachAborted(sender As Object, e As EventArgs)
        'AXDoc.Editor.WriteMessage("AXDB_XrefAttachAborted")
        If logeventos Then PonLogEv("AXDB_XrefAttachAborted")
    End Sub

    Public Shared Sub AXDB_XrefAttachEnded(sender As Object, e As EventArgs)
        'AXDoc.Editor.WriteMessage("AXDB_XrefAttachEnded")
        If logeventos Then PonLogEv("AXDB_XrefAttachEnded")
    End Sub

    Public Shared Sub AXDB_XrefBeginAttached(sender As Object, e As XrefBeginOperationEventArgs)
        'AXDoc.Editor.WriteMessage("AXDB_XrefBeginAttached")
        If logeventos Then PonLogEv("AXDB_XrefBeginAttached")
    End Sub

    Public Shared Sub AXDB_XrefBeginOtherAttached(sender As Object, e As XrefBeginOperationEventArgs)
        'AXDoc.Editor.WriteMessage("AXDB_XrefBeginOtherAttached")
        If logeventos Then PonLogEv("AXDB_XrefBeginOtherAttached")
    End Sub

    Public Shared Sub AXDB_XrefBeginRestore(sender As Object, e As XrefBeginOperationEventArgs)
        'AXDoc.Editor.WriteMessage("AXDB_XrefBeginRestore")
        If logeventos Then PonLogEv("AXDB_XrefBeginRestore")
    End Sub

    Public Shared Sub AXDB_XrefComandeered(sender As Object, e As XrefComandeeredEventArgs)
        'AXDoc.Editor.WriteMessage("AXDB_XrefComandeered")
        If logeventos Then PonLogEv("AXDB_XrefComandeered")
    End Sub

    Public Shared Sub AXDB_XrefPreXrefLockFile(sender As Object, e As XrefPreXrefLockFileEventArgs)
        'AXDoc.Editor.WriteMessage("AXDB_XrefPreXrefLockFile")
        If logeventos Then PonLogEv("AXDB_XrefPreXrefLockFile")
    End Sub

    Public Shared Sub AXDB_XrefRedirected(sender As Object, e As XrefRedirectedEventArgs)
        'AXDoc.Editor.WriteMessage("AXDB_XrefRedirected")
        If logeventos Then PonLogEv("AXDB_XrefRedirected")
    End Sub

    Public Shared Sub AXDB_XrefRestoreAborted(sender As Object, e As EventArgs)
        'AXDoc.Editor.WriteMessage("AXDB_XrefRestoreAborted")
        If logeventos Then PonLogEv("AXDB_XrefRestoreAborted")
    End Sub

    Public Shared Sub AXDB_XrefRestoreEnded(sender As Object, e As EventArgs)
        'AXDoc.Editor.WriteMessage("AXDB_XrefRestoreEnded")
        If logeventos Then PonLogEv("AXDB_XrefRestoreEnded")
    End Sub

    Public Shared Sub AXDB_XrefSubCommandAborted(sender As Object, e As XrefSubCommandEventArgs)
        'AXDoc.Editor.WriteMessage("AXDB_XrefSubCommandAborted")
        If logeventos Then PonLogEv("AXDB_XrefSubCommandAborted")
    End Sub

    Public Shared Sub AXDB_XrefSubCommandEnd(sender As Object, e As XrefSubCommandEventArgs)
        'AXDoc.Editor.WriteMessage("AXDB_XrefSubCommandEnd")
        If logeventos Then PonLogEv("AXDB_XrefSubCommandEnd")
    End Sub

    Public Shared Sub AXDB_XrefSubCommandStart(sender As Object, e As XrefVetoableSubCommandEventArgs)
        'AXDoc.Editor.WriteMessage("AXDB_XrefSubCommandStart")
        If logeventos Then PonLogEv("AXDB_XrefSubCommandStart")
    End Sub
End Class
'
'Los siguientes son algunos de los eventos utilizados para responder a los cambios de objetos a nivel de base de datos
'ObjectAppended         Se activa cuando se agrega un objeto a una base de datos.
'ObjectErased           Se activa cuando un objeto se borra o borra de una base de datos.
'ObjectModified         Se activa cuando un objeto ha sido modificado.
'ObjectOpenedForModify  Se activa antes de que se modifique un objeto.
'ObjectReappended       Se activa cuando un objeto se elimina de una base de datos después de una operación Deshacer Y luego se vuelve a agregar con una operación Rehacer.
'ObjectUnappended       Se activa cuando un objeto se elimina de una base de datos después de una operación Deshacer.