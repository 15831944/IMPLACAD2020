Imports Autodesk.AutoCAD.Interop.Common

'
Partial Public Class Eventos

    Public Shared Sub Subscribe_COMApp()
        AddHandler COMApp.AppActivate, AddressOf COMApp_AppActivate
        AddHandler COMApp.AppDeactivate, AddressOf COMApp_AppDeactivate
        AddHandler COMApp.ARXLoaded, AddressOf COMApp_ARXLoaded
        AddHandler COMApp.ARXUnloaded, AddressOf COMApp_ARXUnloaded
        AddHandler COMApp.BeginCommand, AddressOf COMApp_BeginCommand
        AddHandler COMApp.BeginFileDrop, AddressOf COMApp_BeginFileDrop
        AddHandler COMApp.BeginLisp, AddressOf COMApp_BeginLisp
        AddHandler COMApp.BeginModal, AddressOf COMApp_BeginModal
        AddHandler COMApp.BeginOpen, AddressOf COMApp_BeginOpen
        AddHandler COMApp.BeginPlot, AddressOf COMApp_BeginPlot
        AddHandler COMApp.BeginQuit, AddressOf COMApp_BeginQuit
        AddHandler COMApp.BeginSave, AddressOf COMApp_BeginSave
        AddHandler COMApp.EndCommand, AddressOf COMApp_EndCommand
        AddHandler COMApp.EndLisp, AddressOf COMApp_EndLisp
        AddHandler COMApp.EndModal, AddressOf COMApp_EndModal
        AddHandler COMApp.EndOpen, AddressOf COMApp_EndOpen
        AddHandler COMApp.EndPlot, AddressOf COMApp_EndPlot
        AddHandler COMApp.EndSave, AddressOf COMApp_EndSave
        AddHandler COMApp.LispCancelled, AddressOf COMApp_LispCancelled
        AddHandler COMApp.NewDrawing, AddressOf COMApp_NewDrawing
        AddHandler COMApp.SysVarChanged, AddressOf COMApp_SysVarChanged
        AddHandler COMApp.WindowChanged, AddressOf COMApp_WindowChanged
        AddHandler COMApp.WindowMovedOrResized, AddressOf COMApp_WindowMovedOrResized
    End Sub

    Public Shared Sub Unsubscribe_COMApp()
        Try
            'RemoveHandler COMApp.AppActivate, AddressOf COMApp_AppActivate
            'RemoveHandler COMApp.AppDeactivate, AddressOf COMApp_AppDeactivate
            'RemoveHandler COMApp.ARXLoaded, AddressOf COMApp_ARXLoaded
            'RemoveHandler COMApp.ARXUnloaded, AddressOf COMApp_ARXUnloaded
            'RemoveHandler COMApp.BeginCommand, AddressOf COMApp_BeginCommand
            'RemoveHandler COMApp.BeginFileDrop, AddressOf COMApp_BeginFileDrop
            'RemoveHandler COMApp.BeginLisp, AddressOf COMApp_BeginLisp
            'RemoveHandler COMApp.BeginModal, AddressOf COMApp_BeginModal
            'RemoveHandler COMApp.BeginOpen, AddressOf COMApp_BeginOpen
            'RemoveHandler COMApp.BeginPlot, AddressOf COMApp_BeginPlot
            'RemoveHandler COMApp.BeginQuit, AddressOf COMApp_BeginQuit
            'RemoveHandler COMApp.BeginSave, AddressOf COMApp_BeginSave
            'RemoveHandler COMApp.EndCommand, AddressOf COMApp_EndCommand
            'RemoveHandler COMApp.EndLisp, AddressOf COMApp_EndLisp
            'RemoveHandler COMApp.EndModal, AddressOf COMApp_EndModal
            'RemoveHandler COMApp.EndOpen, AddressOf COMApp_EndOpen
            'RemoveHandler COMApp.EndPlot, AddressOf COMApp_EndPlot
            'RemoveHandler COMApp.EndSave, AddressOf COMApp_EndSave
            'RemoveHandler COMApp.LispCancelled, AddressOf COMApp_LispCancelled
            'RemoveHandler COMApp.NewDrawing, AddressOf COMApp_NewDrawing
            'RemoveHandler COMApp.SysVarChanged, AddressOf COMApp_SysVarChanged
            'RemoveHandler COMApp.WindowChanged, AddressOf COMApp_WindowChanged
            'RemoveHandler COMApp.WindowMovedOrResized, AddressOf COMApp_WindowMovedOrResized
        Catch ex As System.Exception
            Debug.Print(ex.ToString)
        End Try
    End Sub

    Public Shared Sub COMApp_AppActivate()
        'AXDoc.Editor.WriteMessage("COMApp_AppActivate")
        If logeventos Then PonLogEv("COMApp_AppActivate")
    End Sub

    Public Shared Sub COMApp_AppDeactivate()
        'AXDoc.Editor.WriteMessage("COMApp_AppDeactivate")
        If logeventos Then PonLogEv("COMApp_AppDeactivate")
    End Sub

    Public Shared Sub COMApp_ARXLoaded(AppName As String)
        'AXDoc.Editor.WriteMessage("COMApp_ARXLoaded")
        If logeventos Then PonLogEv("COMApp_ARXLoaded")
    End Sub

    Public Shared Sub COMApp_ARXUnloaded(AppName As String)
        'AXDoc.Editor.WriteMessage("COMApp_ARXUnloaded")
        If logeventos Then PonLogEv("COMApp_ARXUnloaded")
    End Sub

    Public Shared Sub COMApp_BeginCommand(CommandName As String)
        'AXDoc.Editor.WriteMessage("COMApp_BeginCommand")
        If logeventos Then PonLogEv("COMApp_BeginCommand")
    End Sub

    Public Shared Sub COMApp_BeginFileDrop(FileName As String, ByRef Cancel As Boolean)
        'AXDoc.Editor.WriteMessage("COMApp_BeginFileDrop")
        If logeventos Then PonLogEv("COMApp_BeginFileDrop")
    End Sub

    Public Shared Sub COMApp_BeginLisp(FirstLine As String)
        'AXDoc.Editor.WriteMessage("COMApp_BeginLisp")
        If logeventos Then PonLogEv("COMApp_BeginLisp")
    End Sub

    Public Shared Sub COMApp_BeginModal()
        'AXDoc.Editor.WriteMessage("COMApp_BeginModal")
        If logeventos Then PonLogEv("COMApp_BeginModal")
    End Sub

    Public Shared Sub COMApp_BeginOpen(ByRef FileName As String)
        'AXDoc.Editor.WriteMessage("COMApp_BeginOpen")
        If logeventos Then PonLogEv("COMApp_BeginOpen")
    End Sub

    Public Shared Sub COMApp_BeginPlot(DrawingName As String)
        'AXDoc.Editor.WriteMessage("COMApp_BeginPlot")
        If logeventos Then PonLogEv("COMApp_BeginPlot")
    End Sub

    Public Shared Sub COMApp_BeginQuit(ByRef Cancel As Boolean)
        'AXDoc.Editor.WriteMessage("COMApp_BeginQuit")
        If logeventos Then PonLogEv("COMApp_BeginQuit")
    End Sub

    Public Shared Sub COMApp_BeginSave(FileName As String)
        'AXDoc.Editor.WriteMessage("COMApp_BeginSave")
        If logeventos Then PonLogEv("COMApp_BeginSave")
    End Sub

    Public Shared Sub COMApp_EndCommand(CommandName As String)
        'AXDoc.Editor.WriteMessage("COMApp_EndCommand")
        If logeventos Then PonLogEv("COMApp_EndCommand")
    End Sub

    Public Shared Sub COMApp_EndLisp()
        'AXDoc.Editor.WriteMessage("COMApp_EndLisp")
        If logeventos Then PonLogEv("COMApp_EndLisp")
    End Sub

    Public Shared Sub COMApp_EndModal()
        'AXDoc.Editor.WriteMessage("COMApp_EndModal")
        If logeventos Then PonLogEv("COMApp_EndModal")
    End Sub

    Public Shared Sub COMApp_EndOpen(FileName As String)
        'AXDoc.Editor.WriteMessage("COMApp_EndOpen")
        If logeventos Then PonLogEv("COMApp_EndOpen")
        'Subscribre_EvDocS(EvDocM.CurrentDocument)
    End Sub

    Public Shared Sub COMApp_EndPlot(DrawingName As String)
        'AXDoc.Editor.WriteMessage("COMApp_EndPlot")
        If logeventos Then PonLogEv("COMApp_EndPlot")
    End Sub

    Public Shared Sub COMApp_EndSave(FileName As String)
        'AXDoc.Editor.WriteMessage("COMApp_EndSave")
        If logeventos Then PonLogEv("COMApp_EndSave")
    End Sub

    Public Shared Sub COMApp_LispCancelled()
        'AXDoc.Editor.WriteMessage("COMApp_LispCancelled")
        If logeventos Then PonLogEv("COMApp_LispCancelled")
    End Sub

    Public Shared Sub COMApp_NewDrawing()
        'AXDoc.Editor.WriteMessage("COMApp_NewDrawing")
        If logeventos Then PonLogEv("COMApp_NewDrawing")
    End Sub

    Public Shared Sub COMApp_SysVarChanged(SysvarName As String, newVal As Object)
        'AXDoc.Editor.WriteMessage("COMApp_SysVarChanged")
        If logeventos Then PonLogEv("COMApp_SysVarChanged")
    End Sub

    Public Shared Sub COMApp_WindowChanged(WindowState As AcWindowState)
        'AXDoc.Editor.WriteMessage("COMApp_WindowChanged")
        If logeventos Then PonLogEv("COMApp_WindowChanged")
    End Sub

    Public Shared Sub COMApp_WindowMovedOrResized(HWNDFrame As Integer, bMoved As Boolean)
        'AXDoc.Editor.WriteMessage("COMApp_WindowMovedOrResized")
        If logeventos Then PonLogEv("COMApp_WindowMovedOrResized")
    End Sub

End Class

'
'AppActivate:           Se activa justo antes de que se active la ventana principal de la aplicación.
'AppDeactivate:         Se activa justo antes de que se desactive la ventana principal de la aplicación.
'ARXLoaded:             Se activa cuando se carga una aplicación ObjectARX.
'ARXUnloaded:           Se activa cuando una aplicación ObjectARX se ha descargado.
'BeginCommand:          Se activa inmediatamente después de que se emite un comando, pero antes de que se complete.
'BeginFileDrop:         Se activa cuando se suelta un archivo en la ventana principal de la aplicación.
'BeginLISP:             Se activa inmediatamente después de que AutoCAD recibe una solicitud para evaluar una expresión LISP.
'BeginModal:            Se activa justo antes de que se muestre un cuadro de diálogo modal.
'BeginOpen:             Se activa inmediatamente después de que AutoCAD recibe una solicitud para abrir un dibujo existente.
'BeginPlot:             Se activa inmediatamente después de que AutoCAD recibe una solicitud para imprimir un dibujo.
'BeginQuit:             Se activa justo antes de que finalice una sesión de AutoCA.
'BeginSave:             Se activa inmediatamente después de que AutoCAD recibe una solicitud para guardar el dibujo.
'EndCommand:            Se activa inmediatamente después de que se completa un comandos.
'EndLISP:               Se activa al finalizar la evaluación de una expresión LISP.
'EndModal:              Se activa justo después de que se descarta un cuadro de diálogo modal.
'EndOpen:               Se activa inmediatamente después de que AutoCAD termina de abrir un dibujo existente.
'EndPlot:               Se activa después de que se haya enviado un documento a la impresora.
'EndSave:               Se activa cuando AutoCAD ha terminado de guardar el dibujog.
'LISPCancelled:         Se activa cuando se cancela la evaluación de una expresión LISP.
'NewDrawing:            Se activa justo antes de que se cree un nuevo dibujo.
'SysVarChanged:         Se activa cuando se cambia el valor de una variable del sistema.
'WindowChanged:         Se activa cuando hay un cambio en la ventana de la aplicación.
'WindowMovedOrResized:  Se activa justo después de que la ventana de la Aplicación se haya movido o cambiado de tamaño.