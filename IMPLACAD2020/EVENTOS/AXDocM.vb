Imports Autodesk.AutoCAD.ApplicationServices

Partial Public Class Eventos

    Public Shared Sub Subscribe_AXDocM()
        AddHandler AXDocM.DocumentActivated, AddressOf AXDocM_DocumentActivated
        AddHandler AXDocM.DocumentActivationChanged, AddressOf AXDocM_DocumentActivationChanged
        AddHandler AXDocM.DocumentBecameCurrent, AddressOf AXDocM_DocumentBecameCurrent
        AddHandler AXDocM.DocumentCreated, AddressOf AXDocM_DocumentCreated
        AddHandler AXDocM.DocumentCreateStarted, AddressOf AXDocM_DocumentCreateStarted
        AddHandler AXDocM.DocumentCreationCanceled, AddressOf AXDocM_DocumentCreationCanceled
        AddHandler AXDocM.DocumentDestroyed, AddressOf AXDocM_DocumentDestroyed
        AddHandler AXDocM.DocumentLockModeChanged, AddressOf AXDocM_DocumentLockModeChanged
        AddHandler AXDocM.DocumentLockModeChangeVetoed, AddressOf AXDocM_DocumentLockModeChangeVetoed
        AddHandler AXDocM.DocumentLockModeWillChange, AddressOf AXDocM_DocumentLockModeWillChange
        AddHandler AXDocM.DocumentToBeActivated, AddressOf AXDocM_DocumentToBeActivated
        AddHandler AXDocM.DocumentToBeDeactivated, AddressOf AXDocM_DocumentToBeDeactivated
        AddHandler AXDocM.DocumentToBeDestroyed, AddressOf AXDocM_DocumentToBeDestroyed
    End Sub

    Public Shared Sub Unsubscribe_AXDocM()
        Try
            'RemoveHandler AXDocM.DocumentActivated, AddressOf AXDocM_DocumentActivated
            'RemoveHandler AXDocM.DocumentActivationChanged, AddressOf AXDocM_DocumentActivationChanged
            'RemoveHandler AXDocM.DocumentBecameCurrent, AddressOf AXDocM_DocumentBecameCurrent
            'RemoveHandler AXDocM.DocumentCreated, AddressOf AXDocM_DocumentCreated
            'RemoveHandler AXDocM.DocumentCreateStarted, AddressOf AXDocM_DocumentCreateStarted
            'RemoveHandler AXDocM.DocumentCreationCanceled, AddressOf AXDocM_DocumentCreationCanceled
            'RemoveHandler AXDocM.DocumentDestroyed, AddressOf AXDocM_DocumentDestroyed
            'RemoveHandler AXDocM.DocumentLockModeChanged, AddressOf AXDocM_DocumentLockModeChanged
            'RemoveHandler AXDocM.DocumentLockModeChangeVetoed, AddressOf AXDocM_DocumentLockModeChangeVetoed
            'RemoveHandler AXDocM.DocumentLockModeWillChange, AddressOf AXDocM_DocumentLockModeWillChange
            'RemoveHandler AXDocM.DocumentToBeActivated, AddressOf AXDocM_DocumentToBeActivated
            'RemoveHandler AXDocM.DocumentToBeDeactivated, AddressOf AXDocM_DocumentToBeDeactivated
            'RemoveHandler AXDocM.DocumentToBeDestroyed, AddressOf AXDocM_DocumentToBeDestroyed
        Catch ex As System.Exception
            Debug.Print(ex.ToString)
        End Try
    End Sub

    Public Shared Sub AXDocM_DocumentActivated(sender As Object, e As DocumentCollectionEventArgs)
        'AXDoc.Editor.WriteMessage("AXDocM_DocumentActivated")
        If logeventos Then PonLogEv("AXDocM_DocumentActivated;" & e.Document.Name)
    End Sub

    Public Shared Sub AXDocM_DocumentActivationChanged(sender As Object, e As DocumentActivationChangedEventArgs)
        'AXDoc.Editor.WriteMessage("AXDocM_DocumentActivationChanged")
        If logeventos Then PonLogEv("AXDocM_DocumentActivationChanged")
    End Sub

    Public Shared Sub AXDocM_DocumentBecameCurrent(sender As Object, e As DocumentCollectionEventArgs)
        'AXDoc.Editor.WriteMessage("AXDocM_DocumentBecameCurrent")
        If logeventos Then PonLogEv("AXDocM_DocumentBecameCurrent")
    End Sub

    Public Shared Sub AXDocM_DocumentCreated(sender As Object, e As DocumentCollectionEventArgs)
        'AXDoc.Editor.WriteMessage("AXDocM_DocumentCreated")
        If logeventos Then PonLogEv("AXDocM_DocumentCreated;" & e.Document.Name)
        Subscribe_AXDoc()
    End Sub

    Public Shared Sub AXDocM_DocumentCreateStarted(sender As Object, e As DocumentCollectionEventArgs)
        'AXDoc.Editor.WriteMessage("AXDocM_DocumentCreateStarted")
        If logeventos Then PonLogEv("AXDocM_DocumentCreateStarted")
    End Sub

    Public Shared Sub AXDocM_DocumentCreationCanceled(sender As Object, e As DocumentCollectionEventArgs)
        'AXDoc.Editor.WriteMessage("AXDocM_DocumentCreationCanceled")
        If logeventos Then PonLogEv("AXDocM_DocumentCreationCanceled")
    End Sub

    Public Shared Sub AXDocM_DocumentDestroyed(sender As Object, e As DocumentDestroyedEventArgs)
        'AXDoc.Editor.WriteMessage("AXDocM_DocumentDestroyed")
        If logeventos Then PonLogEv("AXDocM_DocumentDestroyed")
    End Sub

    Public Shared Sub AXDocM_DocumentLockModeChanged(sender As Object, e As DocumentLockModeChangedEventArgs)
        'AXDoc.Editor.WriteMessage("AXDocM_DocumentLockModeChanged")
        'If logeventos Then PonLogEv("AXDocM_DocumentLockModeChanged")
    End Sub

    Public Shared Sub AXDocM_DocumentLockModeChangeVetoed(sender As Object, e As DocumentLockModeChangeVetoedEventArgs)
        'AXDoc.Editor.WriteMessage("AXDocM_DocumentLockModeChangeVetoed")
        'If logeventos Then PonLogEv("AXDocM_DocumentLockModeChangeVetoed")
    End Sub

    Public Shared Sub AXDocM_DocumentLockModeWillChange(sender As Object, e As DocumentLockModeWillChangeEventArgs)
        'AXDoc.Editor.WriteMessage("AXDocM_DocumentLockModeWillChange")
        'If logeventos Then PonLogEv("AXDocM_DocumentLockModeWillChange")
    End Sub

    Public Shared Sub AXDocM_DocumentToBeActivated(sender As Object, e As DocumentCollectionEventArgs)
        'AXDoc.Editor.WriteMessage("AXDocM_DocumentToBeActivated")
        If logeventos Then PonLogEv("AXDocM_DocumentToBeActivated")
    End Sub

    Public Shared Sub AXDocM_DocumentToBeDeactivated(sender As Object, e As DocumentCollectionEventArgs)
        'AXDoc.Editor.WriteMessage("AXDocM_DocumentToBeDeactivated")
        If logeventos Then PonLogEv("AXDocM_DocumentToBeDeactivated")
    End Sub

    Public Shared Sub AXDocM_DocumentToBeDestroyed(sender As Object, e As DocumentCollectionEventArgs)
        'AXDoc.Editor.WriteMessage("AXDocM_DocumentToBeDestroyed")
        If logeventos Then PonLogEv("AXDocM_DocumentToBeDestroyed;" & e.Document.Name)
        Unsubscribe_AXDoc()
    End Sub

End Class

'
'DocumentActivated              Se activa cuando se activa una ventana de documento.
'DocumentActivationChanged      Se activa después de que la ventana del documento activo se desactiva o destruye.
'DocumentBecameCurrent          Se activa cuando una ventana de documento se configura como actual y es diferente de la ventana de documento activa anterior.
'DocumentCreated                Se activa después de crear una ventana de documento. Se produce después de que se crea un nuevo dibujo o se abre un dibujo existente.
'DocumentCreateStarted          Se activa justo antes de que se cree una ventana de documento. Ocurre antes de que se cree un nuevo dibujo o se abra un dibujo existente.
'DocumentCreationCanceled       Se activa cuando se cancela una solicitud para crear un nuevo dibujo o para abrir un dibujo existente.
'DocumentDestroyed              Se activa antes de que se destruya una ventana de documento y se elimine su objeto de base de datos asociado.
'DocumentLockModeChanged        Se activa después de que el modo de bloqueo de un documento ha cambiado.
'DocumentLockModeChangeVetoed   Se activa después de que se veta el cambio del modo de bloqueo.
'DocumentLockModeWillChange     Se activa antes de que se cambie el modo de bloqueo de un documento.
'DocumentToBeActivated          Se activa cuando un documento está a punto de activarse.
'DocumentToBeDeactivated        Se activa cuando un documento está a punto de desactivarse.
'DocumentToBeDestroyed          Se activa cuando un documento está a punto de ser destruido.