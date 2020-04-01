Imports System.Runtime.InteropServices
Imports Autodesk.AutoCAD.Interop
Imports Autodesk.AutoCAD.Interop.Common

'
Partial Public Class Eventos

    Public Shared Sub Subscribe_COMDoc()
        If COMDoc() Is Nothing Then Exit Sub
        AddHandler COMDoc.Activate, AddressOf COMDoc_Activate
        AddHandler COMDoc.BeginClose, AddressOf COMDoc_BeginClose
        AddHandler COMDoc.BeginDocClose, AddressOf COMDoc_BeginDocClose
        AddHandler COMDoc.BeginCommand, AddressOf COMDoc_BeginCommand
        AddHandler COMDoc.BeginDoubleClick, AddressOf COMDoc_BeginDoubleClick
        AddHandler COMDoc.BeginCommand, AddressOf COMDoc_BeginCommand
        AddHandler COMDoc.BeginLisp, AddressOf COMDoc_BeginLisp
        AddHandler COMDoc.BeginPlot, AddressOf COMDoc_BeginPlot
        AddHandler COMDoc.BeginRightClick, AddressOf COMDoc_BeginRightClick
        AddHandler COMDoc.BeginSave, AddressOf COMDoc_BeginSave
        AddHandler COMDoc.BeginShortcutMenuCommand, AddressOf COMDoc_BeginShortcutMenuCommand
        AddHandler COMDoc.BeginShortcutMenuDefault, AddressOf COMDoc_BeginShortcutMenuDefault
        AddHandler COMDoc.BeginShortcutMenuEdit, AddressOf COMDoc_BeginShortcutMenuEdit
        AddHandler COMDoc.BeginShortcutMenuGrip, AddressOf COMDoc_BeginShortcutMenuGrip
        AddHandler COMDoc.BeginShortcutMenuOsnap, AddressOf COMDoc_BeginShortcutMenuOsnap
        AddHandler COMDoc.Deactivate, AddressOf COMDoc_Deactivate
        AddHandler COMDoc.EndCommand, AddressOf COMDoc_EndCommand
        AddHandler COMDoc.EndLisp, AddressOf COMDoc_EndLisp
        AddHandler COMDoc.EndPlot, AddressOf COMDoc_EndPlot
        AddHandler COMDoc.EndSave, AddressOf COMDoc_EndSave
        AddHandler COMDoc.EndShortcutMenu, AddressOf COMDoc_EndShortcutMenu
        AddHandler COMDoc.LayoutSwitched, AddressOf COMDoc_LayoutSwitched
        AddHandler COMDoc.LispCancelled, AddressOf COMDoc_LispCancelled
        AddHandler COMDoc.ObjectAdded, AddressOf COMDoc_ObjectAdded
        AddHandler COMDoc.ObjectErased, AddressOf COMDoc_ObjectErased
        AddHandler COMDoc.ObjectModified, AddressOf COMDoc_ObjectModified
        AddHandler COMDoc.SelectionChanged, AddressOf COMDoc_SelectionChanged
        AddHandler COMDoc.WindowChanged, AddressOf COMDoc_WindowChanged
        AddHandler COMDoc.WindowMovedOrResized, AddressOf COMDoc_WindowMovedOrResized
    End Sub

    Public Shared Sub Unsubscribe_COMDoc()
        Try
            If COMDoc() Is Nothing Then Exit Sub
            'RemoveHandler COMDoc.Activate, AddressOf COMDoc_Activate
            'RemoveHandler COMDoc.BeginClose, AddressOf COMDoc_BeginClose
            'RemoveHandler COMDoc.BeginDocClose, AddressOf COMDoc_BeginDocClose
            'RemoveHandler COMDoc.BeginCommand, AddressOf COMDoc_BeginCommand
            'RemoveHandler COMDoc.BeginDoubleClick, AddressOf COMDoc_BeginDoubleClick
            'RemoveHandler COMDoc.BeginCommand, AddressOf COMDoc_BeginCommand
            'RemoveHandler COMDoc.BeginLisp, AddressOf COMDoc_BeginLisp
            'RemoveHandler COMDoc.BeginPlot, AddressOf COMDoc_BeginPlot
            'RemoveHandler COMDoc.BeginRightClick, AddressOf COMDoc_BeginRightClick
            'RemoveHandler COMDoc.BeginSave, AddressOf COMDoc_BeginSave
            'RemoveHandler COMDoc.BeginShortcutMenuCommand, AddressOf COMDoc_BeginShortcutMenuCommand
            'RemoveHandler COMDoc.BeginShortcutMenuDefault, AddressOf COMDoc_BeginShortcutMenuDefault
            'RemoveHandler COMDoc.BeginShortcutMenuEdit, AddressOf COMDoc_BeginShortcutMenuEdit
            'RemoveHandler COMDoc.BeginShortcutMenuGrip, AddressOf COMDoc_BeginShortcutMenuGrip
            'RemoveHandler COMDoc.BeginShortcutMenuOsnap, AddressOf COMDoc_BeginShortcutMenuOsnap
            'RemoveHandler COMDoc.Deactivate, AddressOf COMDoc_Deactivate
            'RemoveHandler COMDoc.EndCommand, AddressOf COMDoc_EndCommand
            'RemoveHandler COMDoc.EndLisp, AddressOf COMDoc_EndLisp
            'RemoveHandler COMDoc.EndPlot, AddressOf COMDoc_EndPlot
            'RemoveHandler COMDoc.EndSave, AddressOf COMDoc_EndSave
            'RemoveHandler COMDoc.EndShortcutMenu, AddressOf COMDoc_EndShortcutMenu
            'RemoveHandler COMDoc.LayoutSwitched, AddressOf COMDoc_LayoutSwitched
            'RemoveHandler COMDoc.LispCancelled, AddressOf COMDoc_LispCancelled
            'RemoveHandler COMDoc.ObjectAdded, AddressOf COMDoc_ObjectAdded
            'RemoveHandler COMDoc.ObjectErased, AddressOf COMDoc_ObjectErased
            'RemoveHandler COMDoc.ObjectModified, AddressOf COMDoc_ObjectModified
            'RemoveHandler COMDoc.SelectionChanged, AddressOf COMDoc_SelectionChanged
            'RemoveHandler COMDoc.WindowChanged, AddressOf COMDoc_WindowChanged
            'RemoveHandler COMDoc.WindowMovedOrResized, AddressOf COMDoc_WindowMovedOrResized
        Catch ex As System.Exception
            Debug.Print(ex.ToString)
        End Try
    End Sub

    Public Shared Sub COMDoc_Activate()
        'AXDoc.Editor.WriteMessage("COMDoc_Activate")
        If logeventos Then PonLogEv("COMDoc_Activate")
    End Sub

    Public Shared Sub COMDoc_BeginClose()
        'AXDoc.Editor.WriteMessage("COMDoc_BeginClose")
        If logeventos Then PonLogEv("COMDoc_BeginClose")
    End Sub

    Public Shared Sub COMDoc_BeginCommand(CommandName As String)
        'AXDoc.Editor.WriteMessage("COMDoc_BeginCommand")
        If logeventos Then PonLogEv("COMDoc_BeginCommand;" & CommandName)
        If coneventos = False Then Exit Sub
        'If CommandName = "SAVE" Or CommandName = "QSAVE" Then
        'If clsA Is Nothing Then clsA = New AutoCAD2acad.clsAutoCAD2acad(Ev.EvApp, cfg._appFullPath, regAPPCliente)
        'clsA.SeleccionaTodosObjetos("INSERT",, True)
        'Ev.EvApp.ActiveDocument.SendCommand("_UPDATEFIELD _ALL  ")

        '
        'Dim arrIdBloques As ArrayList     '' Arraylist con los IDs de los bloques que empiezan por
        'arrIdBloques = clsA.SeleccionaTodosObjetos("INSERT",, True)
        '
        ' Procesamos todos los bloques (arralist de Ids)
        'Dim arrHighlight As New ArrayList
        'For Each queId As Long In arrIdBloques
        '    Dim oBl As AcadBlockReference = Ev.EvApp.ActiveDocument.ObjectIdToObject(queId)
        '    ''
        '    Dim capa As String = oBl.Layer
        '    Dim partes() As String = capa.Split("·"c)
        '    Dim nivel As String = modAutoCAD.BloqueAtributoDame(queId, "NIVEL")
        '    If partes.Length > 1 AndAlso (partes(1) = ultimoNivel Or nivel = ultimoNivel) Then
        '        arrHighlight.Add(CType(oBl, AcadEntity))
        '    End If
        'Next
        ''
        'If arrHighlight IsNot Nothing AndAlso arrHighlight.Count > 0 Then
        '    modAutoCAD.SeleccionCreaResalta(queEntidades:=arrHighlight, tiempo:=5000, conZoom:=True)
        '    'Autodesk.AutoCAD.Internal.Utils.SelectObjects(ids:=Autodesk.AutoCAD.DatabaseServices.ObjectId)
        '    'Autodesk.AutoCAD.Internal.Utils.ZoomObjects(False)
        '    Ev.EvApp.ZoomScaled(0.75, AcZoomScaleType.acZoomScaledRelative)
        'End If
        'End If
    End Sub

    Public Shared Sub COMDoc_BeginDocClose(ByRef Cancel As Boolean)
        'AXDoc.Editor.WriteMessage("COMDoc_BeginDocClose")
        If logeventos Then PonLogEv("COMDoc_BeginDocClose")
    End Sub

    Public Shared Sub COMDoc_BeginDoubleClick(PickPoint As Object)
        'AXDoc.Editor.WriteMessage("COMDoc_BeginDoubleClick")
        If logeventos Then PonLogEv("COMDoc_BeginDoubleClick;" & PickPoint.ToString)
    End Sub

    Public Shared Sub COMDoc_BeginLisp(FirstLine As String)
        'AXDoc.Editor.WriteMessage("COMDoc_BeginLisp")
        If logeventos Then PonLogEv("COMDoc_BeginLisp;" & FirstLine)
    End Sub

    Public Shared Sub COMDoc_BeginPlot(DrawingName As String)
        'AXDoc.Editor.WriteMessage("COMDoc_BeginPlot")
        If logeventos Then PonLogEv("COMDoc_BeginPlot;" & DrawingName)
    End Sub

    Public Shared Sub COMDoc_BeginRightClick(PickPoint As Object)
        'AXDoc.Editor.WriteMessage("COMDoc_BeginRightClick")
        If logeventos Then PonLogEv("COMDoc_BeginRightClick;" & PickPoint.ToString)
    End Sub

    Public Shared Sub COMDoc_BeginSave(FileName As String)
        'AXDoc.Editor.WriteMessage("COMDoc_BeginSave")
        If logeventos Then PonLogEv("COMDoc_BeginSave;" & FileName)
    End Sub

    Public Shared Sub COMDoc_BeginShortcutMenuCommand(ByRef ShortcutMenu As AcadPopupMenu, Command As String)
        'AXDoc.Editor.WriteMessage("COMDoc_BeginShortcutMenuCommand")
        If logeventos Then PonLogEv("COMDoc_BeginShortcutMenuCommand;" & ShortcutMenu.Name & "|" & Command)
        'AcadPopupMenuItem_PonerQuitar(ShortcutMenu, "TAVILACERCADE", "Tavil. Acerca de...", poner:=True)
    End Sub

    Public Shared Sub COMDoc_BeginShortcutMenuDefault(ByRef ShortcutMenu As AcadPopupMenu)
        'AXDoc.Editor.WriteMessage("COMDoc_BeginShortcutMenuDefault")
        If logeventos Then PonLogEv("COMDoc_BeginShortcutMenuDefault;" & ShortcutMenu.Name)
        'AcadPopupMenuItem_PonerQuitar(ShortcutMenu, "TAVILACERCADE", "Tavil. Acerca de...", poner:=True)
    End Sub

    Public Shared Sub COMDoc_BeginShortcutMenuEdit(ByRef ShortcutMenu As AcadPopupMenu, ByRef SelectionSet As AcadSelectionSet)
        'AXDoc.Editor.WriteMessage("COMDoc_BeginShortcutMenuEdit")
        If logeventos Then PonLogEv("COMDoc_BeginShortcutMenuEdit;" & ShortcutMenu.Name & SelectionSet.Count)
        'AcadPopupMenuItem_PonerQuitar(ShortcutMenu, "TAVILACERCADE", "Tavil. Acerca de...", poner:=True)
    End Sub

    Public Shared Sub COMDoc_BeginShortcutMenuGrip(ByRef ShortcutMenu As AcadPopupMenu)
        'AXDoc.Editor.WriteMessage("COMDoc_BeginShortcutMenuGrip")
        If logeventos Then PonLogEv("COMDoc_BeginShortcutMenuGrip;" & ShortcutMenu.Name)
        'AcadPopupMenuItem_PonerQuitar(ShortcutMenu, "TAVILACERCADE", "Tavil. Acerca de...", poner:=True)
    End Sub

    Public Shared Sub COMDoc_BeginShortcutMenuOsnap(ByRef ShortcutMenu As AcadPopupMenu)
        'AXDoc.Editor.WriteMessage("COMDoc_BeginShortcutMenuOsnap")
        If logeventos Then PonLogEv("COMDoc_BeginShortcutMenuOsnap;" & ShortcutMenu.Name)
        'AcadPopupMenuItem_PonerQuitar(ShortcutMenu, "TAVILACERCADE", "Tavil. Acerca de...", poner:=True)
    End Sub

    Public Shared Sub COMDoc_Deactivate()
        'AXDoc.Editor.WriteMessage("COMDoc_Deactivate")
        If logeventos Then PonLogEv("COMDoc_Deactivate")
    End Sub

    Public Shared Sub COMDoc_EndCommand(CommandName As String)
        'AXDoc.Editor.WriteMessage("COMDoc_EndCommand")
        If logeventos Then PonLogEv("COMDoc_EndCommand;" & CommandName)
        If coneventos = False Then Exit Sub
        'If CommandName = "INSERT" Then

        '    'finComando = True
        'End If
    End Sub

    Public Shared Sub COMDoc_EndLisp()
        'AXDoc.Editor.WriteMessage("COMDoc_EndLisp")
        If logeventos Then PonLogEv("COMDoc_EndLisp")
    End Sub

    Public Shared Sub COMDoc_EndPlot(DrawingName As String)
        'AXDoc.Editor.WriteMessage("COMDoc_EndPlot")
        If logeventos Then PonLogEv("COMDoc_EndPlotCommandName;" & DrawingName)
    End Sub

    Public Shared Sub COMDoc_EndSave(FileName As String)
        'AXDoc.Editor.WriteMessage("COMDoc_EndSave")
        If logeventos Then PonLogEv("COMDoc_EndSave;" & FileName)
    End Sub

    Public Shared Sub COMDoc_EndShortcutMenu(ByRef ShortcutMenu As AcadPopupMenu)
        'AXDoc.Editor.WriteMessage("COMDoc_EndShortcutMenu")
        If logeventos Then PonLogEv("COMDoc_EndShortcutMenu;" & ShortcutMenu.Name)
        'AcadPopupMenuItem_PonerQuitar(ShortcutMenu, "TAVILACERCADE", "Tavil. Acerca de...", poner:=False)
    End Sub

    Public Shared Sub COMDoc_LayoutSwitched(LayoutName As String)
        'AXDoc.Editor.WriteMessage("COMDoc_LayoutSwitched")
        If logeventos Then PonLogEv("COMDoc_LayoutSwitched;" & LayoutName)
    End Sub

    Public Shared Sub COMDoc_LispCancelled()
        'AXDoc.Editor.WriteMessage("COMDoc_LispCancelled")
        If logeventos Then PonLogEv("COMDoc_LispCancelled")
    End Sub

    Public Shared Sub COMDoc_ObjectAdded([Object] As Object)
        'AXDoc.Editor.WriteMessage("COMDoc_ObjectAdded")
        If logeventos Then PonLogEv("COMDoc_ObjectAdded;" & [Object].GetType.ToString)
        'If clsA Is Nothing Then clsA = New AutoCAD2acad.A2acad.A2acad(Eventos.COMApp, cfg._appFullPath, regAPPCliente)
        ''Debug.Print([Object].EntityName)
        'Dim acadO As AcadObject = EvCOMApp.ActiveDocument.ObjectIdToObject([Object].ObjectId)
        'Dim oId As ObjectId = New ObjectId(acadO.ObjectID)
        'Dim acadDb As DBObject = clsA.DBObject_Get(oId)
        'Debug.Print(acadO.ObjectName)
        ''
        'If lTypesAXObj.Contains(acadDb.GetType.Name) AndAlso
        '    lHasCode.Contains(acadDb.AcadObject.OBjectId.ToString) = False Then
        '    Subscribre_EvCOMObj(acadO)
        '    Subscribe_EvAXObj(acadDb)
        'End If
    End Sub

    Public Shared Sub COMDoc_ObjectErased(<ComAliasName("AXDBLib.LONG_PTR")> ObjectId As Long)
        'AXDoc.Editor.WriteMessage("COMDoc_ObjectErased")
        If logeventos Then PonLogEv("COMDoc_ObjectErased;" & ObjectId.ToString)
        'Try
        '    Unsubscribe_EvAXObj(clsA.DBObject_Get(ObjectId))
        '    Dim colIds As New ObjectIdCollection
        '    colIds.Add(New Autodesk.AutoCAD.DatabaseServices.ObjectId(New IntPtr(ObjectId)))
        '    EvAXDocM.CurrentDocument.Database.Purge(colIds)
        'Catch ex As System.Exception

        'End Try
    End Sub

    Public Shared Sub COMDoc_ObjectModified(queObj As Object)
        'AXDoc.Editor.WriteMessage("COMDoc_ObjectModified")
        If logeventos Then PonLogEv("COMDoc_ObjectModified;" & queObj.GetType.ToString)
        'If TypeOf queObj Is Autodesk.AutoCAD.Interop.Common.AcadBlockReference Then
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
        '    '    End If
        '    '    ' Volver a false para activar los eventos.
        '    '    app_procesointerno = False
        '    '    'Ev.EvApp.ActiveDocument.Regen(AcRegenType.acActiveViewport)
        'ElseIf TypeOf queObj Is Autodesk.AutoCAD.Interop.Common.AcadCircle Then
        '    Debug.Print(CType(queObj, Autodesk.AutoCAD.Interop.Common.AcadCircle).Diameter)
        'End If
    End Sub

    Public Shared Sub COMDoc_SelectionChanged()
        'AXDoc.Editor.WriteMessage("COMDoc_SelectionChanged")
        If logeventos Then PonLogEv("COMDoc_SelectionChanged")
    End Sub

    Public Shared Sub COMDoc_WindowChanged(WindowState As AcWindowState)
        'AXDoc.Editor.WriteMessage("COMDoc_WindowChanged")
        If logeventos Then PonLogEv("COMDoc_WindowChanged;" & WindowState.ToString)
    End Sub

    Public Shared Sub COMDoc_WindowMovedOrResized(<ComAliasName("AXDBLib.LONG_PTR")> HWNDFrame As Long, bMoved As Boolean)
        'AXDoc.Editor.WriteMessage("COMDoc_WindowMovedOrResized")
        If logeventos Then PonLogEv("COMDoc_WindowMovedOrResized;" & HWNDFrame & "|" & bMoved.ToString)
    End Sub

End Class

'
'Activate:                  Se activa cuando se activa una ventana de documento.
'BeginClose:                Se activa inmediatamente después de que AutoCAD recibe una solicitud para cerrar un dibujo.
'BeginDocClose:             Se activa justo después de recibir una solicitud para cerrar un dibujo.
'BeginCommand:              Se activa inmediatamente después de que se emite un comando, pero antes de que se complete.
'BeginDoubleClick:          Se activa después de que el usuario haga doble clic en un objeto del dibujo.
'BeginLISP:                 Se activa inmediatamente después de que AutoCAD recibe una solicitud para evaluar una expresión LISP.
'BeginPlot:                 Se activa inmediatamente después de que AutoCAD recibe una solicitud para imprimir un dibujo.
'BeginRightClick:           Se activa después de que el usuario hace clic con el botón derecho en la ventana Dibujo.
'BeginSave:                 Se activa inmediatamente después de que AutoCAD recibe una solicitud para guardar el dibujo.
'BeginShortcutMenuCommand:  Se activa después de que el usuario hace clic con el botón derecho en la ventana Dibujo y antes de que aparezca el menú contextual en el modo Comando..
'BeginShortcutMenuDefault:  Se activa después de que el usuario haga clic con el botón derecho en la ventana Dibujo y antes de que aparezca el menú contextual en el modo Predeterminado.
'BeginShortcutMenuEdit:     Se activa después de que el usuario hace clic con el botón derecho en la ventana Dibujo y antes de que aparezca el menú contextual en el modo Editar.
'BeginShortcutMenuGrip:     Se activa después de que el usuario haga clic con el botón derecho en la ventana Dibujo y antes de que aparezca el menú contextual en el modo Agarre.
'BeginShortcutMenuOsnap:    Se activa después de que el usuario hace clic con el botón derecho en la ventana Dibujo y antes de que aparezca el menú contextual en modo Osnap.
'Deactivate:                Se activa cuando la ventana de dibujo está desactivada.
'EndCommand:                Se activa inmediatamente después de que se completa un comando.
'EndLISP:                   Se activa al finalizar la evaluación de una expresión LISP.
'EndPlot:                   Se activa después de que se haya enviado un documento a la impresora.
'EndSave:                   Se activa cuando AutoCAD ha terminado de guardar el dibujo.
'EndShortcutMenu:           Se activa después de que aparece el menú contextual.
'LayoutSwitched:            Se activa después de que el usuario cambia a un diseño diferente (presentacion).
'LISPCancelled:             Se activa cuando se cancela la evaluación de una expresión LISP.
'ObjectAdded:               Se activa cuando se agrega un objeto al dibujo.
'ObjectErased:              Se activa cuando un objeto ha sido borrado del dibujo.
'ObjectModified:            Se activa cuando un objeto en el dibujo ha sido modificado.
'SelectionChanged:          Se activa cuando cambia el conjunto de selección actual de pickfirst.
'WindowChanged:             Se activa cuando hay un cambio en la ventana de documento.
'WindowMovedOrResized:      Se activa justo después de que la ventana Dibujo se haya movido o cambiado de tamaño.