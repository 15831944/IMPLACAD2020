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
    Public Shared Sub Subscribe_AXEditor()
        AddHandler AXEditor().SelectionRemoved, AddressOf AXEditor_SelectionRemoved_Handler
        AddHandler AXEditor().SelectionAdded, AddressOf AXEditor_SelectionAdded_Handler
        AddHandler AXEditor().Rollover, AddressOf AXEditor_Rollover_Handler
        AddHandler AXEditor().PromptingForString, AddressOf AXEditor_PromptingForString_Handler
        AddHandler AXEditor().PromptingForSelection, AddressOf AXEditor_PromptingForSelection_Handler
        AddHandler AXEditor().PromptingForPoint, AddressOf AXEditor_PromptingForPoint_Handler
        AddHandler AXEditor().PromptingForNestedEntity, AddressOf AXEditor_PromptingForNestedEntity_Handler
        AddHandler AXEditor().PromptingForKeyword, AddressOf AXEditor_PromptingForKeyword_Handler
        AddHandler AXEditor().PromptingForInteger, AddressOf AXEditor_PromptingForInteger_Handler
        AddHandler AXEditor().PromptingForEntity, AddressOf AXEditor_PromptingForEntity_Handler
        AddHandler AXEditor().PromptingForDouble, AddressOf AXEditor_PromptingForDouble_Handler
        AddHandler AXEditor().PromptingForDistance, AddressOf AXEditor_PromptingForDistance_Handler
        AddHandler AXEditor().PromptingForCorner, AddressOf AXEditor_PromptingForCorner_Handler
        AddHandler AXEditor().PromptingForAngle, AddressOf AXEditor_PromptingForAngle_Handler
        AddHandler AXEditor().PromptedForString, AddressOf AXEditor_PromptedForString_Handler
        AddHandler AXEditor().PromptedForSelection, AddressOf AXEditor_PromptedForSelection_Handler
        AddHandler AXEditor().PromptedForPoint, AddressOf AXEditor_PromptedForPoint_Handler
        AddHandler AXEditor().PromptedForNestedEntity, AddressOf AXEditor_PromptedForNestedEntity_Handler
        AddHandler AXEditor().PromptedForKeyword, AddressOf AXEditor_PromptedForKeyword_Handler
        AddHandler AXEditor().PromptedForInteger, AddressOf AXEditor_PromptedForInteger_Handler
        AddHandler AXEditor().PromptedForEntity, AddressOf AXEditor_PromptedForEntity_Handler
        AddHandler AXEditor().PromptedForDouble, AddressOf AXEditor_PromptedForDouble_Handler
        AddHandler AXEditor().PromptedForDistance, AddressOf AXEditor_PromptedForDistance_Handler
        AddHandler AXEditor().PromptedForCorner, AddressOf AXEditor_PromptedForCorner_Handler
        AddHandler AXEditor().PromptedForAngle, AddressOf AXEditor_PromptedForAngle_Handler
        AddHandler AXEditor().PromptForSelectionEnding, AddressOf AXEditor_PromptForSelectionEnding_Handler
        AddHandler AXEditor().PromptForEntityEnding, AddressOf AXEditor_PromptForEntityEnding_Handler
        AddHandler AXEditor().PointMonitor, AddressOf AXEditor_PointMonitor_Handler
        AddHandler AXEditor().PointFilter, AddressOf AXEditor_PointFilter_Handler
        AddHandler AXEditor().LeavingQuiescentState, AddressOf AXEditor_LeavingQuiescentState_Handler
        AddHandler AXEditor().EnteringQuiescentState, AddressOf AXEditor_EnteringQuiescentState_Handler
        AddHandler AXEditor().DraggingEnded, AddressOf AXEditor_DraggingEnded_Handler
        AddHandler AXEditor().Dragging, AddressOf AXEditor_Dragging_Handler
    End Sub

    Public Shared Sub Unsubscribe_AXEditor()
        Try
            'RemoveHandler AXEditor().SelectionRemoved, AddressOf AXEditor_SelectionRemoved_Handler
            'RemoveHandler AXEditor().SelectionAdded, AddressOf AXEditor_SelectionAdded_Handler
            'RemoveHandler AXEditor().Rollover, AddressOf AXEditor_Rollover_Handler
            'RemoveHandler AXEditor().PromptingForString, AddressOf AXEditor_PromptingForString_Handler
            'RemoveHandler AXEditor().PromptingForSelection, AddressOf AXEditor_PromptingForSelection_Handler
            'RemoveHandler AXEditor().PromptingForPoint, AddressOf AXEditor_PromptingForPoint_Handler
            'RemoveHandler AXEditor().PromptingForNestedEntity, AddressOf AXEditor_PromptingForNestedEntity_Handler
            'RemoveHandler AXEditor().PromptingForKeyword, AddressOf AXEditor_PromptingForKeyword_Handler
            'RemoveHandler AXEditor().PromptingForInteger, AddressOf AXEditor_PromptingForInteger_Handler
            'RemoveHandler AXEditor().PromptingForEntity, AddressOf AXEditor_PromptingForEntity_Handler
            'RemoveHandler AXEditor().PromptingForDouble, AddressOf AXEditor_PromptingForDouble_Handler
            'RemoveHandler AXEditor().PromptingForDistance, AddressOf AXEditor_PromptingForDistance_Handler
            'RemoveHandler AXEditor().PromptingForCorner, AddressOf AXEditor_PromptingForCorner_Handler
            'RemoveHandler AXEditor().PromptingForAngle, AddressOf AXEditor_PromptingForAngle_Handler
            'RemoveHandler AXEditor().PromptedForString, AddressOf AXEditor_PromptedForString_Handler
            'RemoveHandler AXEditor().PromptedForSelection, AddressOf AXEditor_PromptedForSelection_Handler
            'RemoveHandler AXEditor().PromptedForPoint, AddressOf AXEditor_PromptedForPoint_Handler
            'RemoveHandler AXEditor().PromptedForNestedEntity, AddressOf AXEditor_PromptedForNestedEntity_Handler
            'RemoveHandler AXEditor().PromptedForKeyword, AddressOf AXEditor_PromptedForKeyword_Handler
            'RemoveHandler AXEditor().PromptedForInteger, AddressOf AXEditor_PromptedForInteger_Handler
            'RemoveHandler AXEditor().PromptedForEntity, AddressOf AXEditor_PromptedForEntity_Handler
            'RemoveHandler AXEditor().PromptedForDouble, AddressOf AXEditor_PromptedForDouble_Handler
            'RemoveHandler AXEditor().PromptedForDistance, AddressOf AXEditor_PromptedForDistance_Handler
            'RemoveHandler AXEditor().PromptedForCorner, AddressOf AXEditor_PromptedForCorner_Handler
            'RemoveHandler AXEditor().PromptedForAngle, AddressOf AXEditor_PromptedForAngle_Handler
            'RemoveHandler AXEditor().PromptForSelectionEnding, AddressOf AXEditor_PromptForSelectionEnding_Handler
            'RemoveHandler AXEditor().PromptForEntityEnding, AddressOf AXEditor_PromptForEntityEnding_Handler
            'RemoveHandler AXEditor().PointMonitor, AddressOf AXEditor_PointMonitor_Handler
            'RemoveHandler AXEditor().PointFilter, AddressOf AXEditor_PointFilter_Handler
            'RemoveHandler AXEditor().LeavingQuiescentState, AddressOf AXEditor_LeavingQuiescentState_Handler
            'RemoveHandler AXEditor().EnteringQuiescentState, AddressOf AXEditor_EnteringQuiescentState_Handler
            'RemoveHandler AXEditor().DraggingEnded, AddressOf AXEditor_DraggingEnded_Handler
            'RemoveHandler AXEditor().Dragging, AddressOf AXEditor_Dragging_Handler
        Catch ex As system.Exception
            Debug.Print(ex.ToString)
        End Try
    End Sub

    Public Shared Sub AXEditor_Dragging_Handler(ByVal sender As Object, ByVal e As DraggingEventArgs)
        'AXDoc.Editor.WriteMessage("AXEditor_Dragging_Handler")
        If logeventos Then PonLogEv("AXEditor_Dragging_Handler")
    End Sub

    Public Shared Sub AXEditor_DraggingEnded_Handler(ByVal sender As Object, ByVal e As DraggingEndedEventArgs)
        'AXDoc.Editor.WriteMessage("AXEditor_DraggingEnded_Handler")
        If logeventos Then PonLogEv("AXEditor_DraggingEnded_Handler")
    End Sub


    ' AutoCAD entre en estado inactivo (Esta trabajando) is Quiescent = es inactivo
    Public Shared Sub AXEditor_EnteringQuiescentState_Handler(ByVal sender As Object, ByVal e As EventArgs)
        'AXDoc.Editor.WriteMessage("AXEditor_EnteringQuiescentState_Handler")
        'If logeventos Then PonLogEv("AXEditor_EnteringQuiescentState_Handler")
    End Sub

    ' AutoCAD sale de inactivo (Esta libre) is Quiescent = es inactivo
    Public Shared Sub AXEditor_LeavingQuiescentState_Handler(ByVal sender As Object, ByVal e As EventArgs)
        'AXDoc.Editor.WriteMessage("AXEditor_LeavingQuiescentState_Handler")
        'If logeventos Then PonLogEv("AXEditor_LeavingQuiescentState_Handler")
    End Sub

    Public Shared Sub AXEditor_PointFilter_Handler(ByVal sender As Object, ByVal e As PointFilterEventArgs)
        'AXDoc.Editor.WriteMessage("AXEditor_PointFilter_Handler")
        'If logeventos Then PonLogEv("AXEditor_PointFilter_Handler")
    End Sub

    Public Shared Sub AXEditor_PointMonitor_Handler(ByVal sender As Object, ByVal e As PointMonitorEventArgs)
        'AXDoc.Editor.WriteMessage("AXEditor_PointMonitor_Handler")
        If logeventos Then PonLogEv("AXEditor_PointMonitor_Handler")
    End Sub

    Public Shared Sub AXEditor_PromptForEntityEnding_Handler(ByVal sender As Object, ByVal e As PromptForEntityEndingEventArgs)
        'AXDoc.Editor.WriteMessage("AXEditor_PromptForEntityEnding_Handler")
        If logeventos Then PonLogEv("AXEditor_PromptForEntityEnding_Handler")
    End Sub

    Public Shared Sub AXEditor_PromptForSelectionEnding_Handler(ByVal sender As Object, ByVal e As PromptForSelectionEndingEventArgs)
        'AXDoc.Editor.WriteMessage("AXEditor_PromptForSelectionEnding_Handler")
        If logeventos Then PonLogEv("AXEditor_PromptForSelectionEnding_Handler")
    End Sub

    Public Shared Sub AXEditor_PromptedForAngle_Handler(ByVal sender As Object, ByVal e As PromptDoubleResultEventArgs)
        'AXDoc.Editor.WriteMessage("AXEditor_PromptedForAngle_Handler")
        If logeventos Then PonLogEv("AXEditor_PromptedForAngle_Handler")
    End Sub

    Public Shared Sub AXEditor_PromptedForCorner_Handler(ByVal sender As Object, ByVal e As PromptPointResultEventArgs)
        'AXDoc.Editor.WriteMessage("AXEditor_PromptedForCorner_Handler")
        If logeventos Then PonLogEv("AXEditor_PromptedForCorner_Handler")
    End Sub

    Public Shared Sub AXEditor_PromptedForDistance_Handler(ByVal sender As Object, ByVal e As PromptDoubleResultEventArgs)
        'AXDoc.Editor.WriteMessage("AXEditor_PromptedForDistance_Handler")
        If logeventos Then PonLogEv("AXEditor_PromptedForDistance_Handler")
    End Sub

    Public Shared Sub AXEditor_PromptedForDouble_Handler(ByVal sender As Object, ByVal e As PromptDoubleResultEventArgs)
        'AXDoc.Editor.WriteMessage("AXEditor_PromptedForDouble_Handler")
        If logeventos Then PonLogEv("AXEditor_PromptedForDouble_Handler")
    End Sub

    Public Shared Sub AXEditor_PromptedForEntity_Handler(ByVal sender As Object, ByVal e As PromptEntityResultEventArgs)
        'AXDoc.Editor.WriteMessage("AXEditor_PromptedForEntity_Handler")
        If logeventos Then PonLogEv("AXEditor_PromptedForEntity_Handler")
    End Sub

    Public Shared Sub AXEditor_PromptedForInteger_Handler(ByVal sender As Object, ByVal e As PromptIntegerResultEventArgs)
        'AXDoc.Editor.WriteMessage("AXEditor_PromptedForInteger_Handler")
        If logeventos Then PonLogEv("AXEditor_PromptedForInteger_Handler")
    End Sub

    Public Shared Sub AXEditor_PromptedForKeyword_Handler(ByVal sender As Object, ByVal e As PromptStringResultEventArgs)
        'AXDoc.Editor.WriteMessage("AXEditor_PromptedForKeyword_Handler")
        If logeventos Then PonLogEv("AXEditor_PromptedForKeyword_Handler")
    End Sub

    Public Shared Sub AXEditor_PromptedForNestedEntity_Handler(ByVal sender As Object, ByVal e As PromptNestedEntityResultEventArgs)
        'AXDoc.Editor.WriteMessage("AXEditor_PromptedForNestedEntity_Handler")
        If logeventos Then PonLogEv("AXEditor_PromptedForNestedEntity_Handler")
    End Sub

    Public Shared Sub AXEditor_PromptedForPoint_Handler(ByVal sender As Object, ByVal e As PromptPointResultEventArgs)
        'AXDoc.Editor.WriteMessage("AXEditor_PromptedForPoint_Handler")
        If logeventos Then PonLogEv("AXEditor_PromptedForPoint_Handler")
    End Sub

    Public Shared Sub AXEditor_PromptedForSelection_Handler(ByVal sender As Object, ByVal e As PromptSelectionResultEventArgs)
        'AXDoc.Editor.WriteMessage("AXEditor_PromptedForSelection_Handler")
        If logeventos Then PonLogEv("AXEditor_PromptedForSelection_Handler")
    End Sub

    Public Shared Sub AXEditor_PromptedForString_Handler(ByVal sender As Object, ByVal e As PromptStringResultEventArgs)
        'AXDoc.Editor.WriteMessage("AXEditor_PromptedForString_Handler")
        If logeventos Then PonLogEv("AXEditor_PromptedForString_Handler")
    End Sub

    Public Shared Sub AXEditor_PromptingForAngle_Handler(ByVal sender As Object, ByVal e As PromptAngleOptionsEventArgs)
        'AXDoc.Editor.WriteMessage("AXEditor_PromptingForAngle_Handler")
        If logeventos Then PonLogEv("AXEditor_PromptingForAngle_Handler")
    End Sub

    Public Shared Sub AXEditor_PromptingForCorner_Handler(ByVal sender As Object, ByVal e As PromptPointOptionsEventArgs)
        'AXDoc.Editor.WriteMessage("AXEditor_PromptingForCorner_Handler")
        If logeventos Then PonLogEv("AXEditor_PromptingForCorner_Handler")
    End Sub

    Public Shared Sub AXEditor_PromptingForDistance_Handler(ByVal sender As Object, ByVal e As PromptDistanceOptionsEventArgs)
        'AXDoc.Editor.WriteMessage("AXEditor_PromptingForDistance_Handler")
        If logeventos Then PonLogEv("AXEditor_PromptingForDistance_Handler")
    End Sub

    Public Shared Sub AXEditor_PromptingForDouble_Handler(ByVal sender As Object, ByVal e As PromptDoubleOptionsEventArgs)
        'AXDoc.Editor.WriteMessage("AXEditor_PromptingForDouble_Handler")
        If logeventos Then PonLogEv("AXEditor_PromptingForDouble_Handler")
    End Sub

    Public Shared Sub AXEditor_PromptingForEntity_Handler(ByVal sender As Object, ByVal e As PromptEntityOptionsEventArgs)
        'AXDoc.Editor.WriteMessage("AXEditor_PromptingForEntity_Handler")
        If logeventos Then PonLogEv("AXEditor_PromptingForEntity_Handler")
    End Sub

    Public Shared Sub AXEditor_PromptingForInteger_Handler(ByVal sender As Object, ByVal e As PromptIntegerOptionsEventArgs)
        'AXDoc.Editor.WriteMessage("AXEditor_PromptingForInteger_Handler")
        If logeventos Then PonLogEv("AXEditor_PromptingForInteger_Handler")
    End Sub

    Public Shared Sub AXEditor_PromptingForKeyword_Handler(ByVal sender As Object, ByVal e As PromptKeywordOptionsEventArgs)
        'AXDoc.Editor.WriteMessage("AXEditor_PromptingForKeyword_Handler")
        If logeventos Then PonLogEv("AXEditor_PromptingForKeyword_Handler")
    End Sub

    Public Shared Sub AXEditor_PromptingForNestedEntity_Handler(ByVal sender As Object, ByVal e As PromptNestedEntityOptionsEventArgs)
        'AXDoc.Editor.WriteMessage("AXEditor_PromptingForNestedEntity_Handler")
        If logeventos Then PonLogEv("AXEditor_PromptingForNestedEntity_Handler")
    End Sub

    Public Shared Sub AXEditor_PromptingForPoint_Handler(ByVal sender As Object, ByVal e As PromptPointOptionsEventArgs)
        'AXDoc.Editor.WriteMessage("AXEditor_PromptingForPoint_Handler")
        If logeventos Then PonLogEv("AXEditor_PromptingForPoint_Handler")
    End Sub

    Public Shared Sub AXEditor_PromptingForSelection_Handler(ByVal sender As Object, ByVal e As PromptSelectionOptionsEventArgs)
        'AXDoc.Editor.WriteMessage("AXEditor_PromptingForSelection_Handler")
        If logeventos Then PonLogEv("AXEditor_PromptingForSelection_Handler")
    End Sub

    Public Shared Sub AXEditor_PromptingForString_Handler(ByVal sender As Object, ByVal e As PromptStringOptionsEventArgs)
        'AXDoc.Editor.WriteMessage("AXEditor_PromptingForString_Handler")
        If logeventos Then PonLogEv("AXEditor_PromptingForString_Handler")
    End Sub

    Public Shared Sub AXEditor_Rollover_Handler(ByVal sender As Object, ByVal e As RolloverEventArgs)
        'AXDoc.Editor.WriteMessage("AXEditor_Rollover_Handler")
        If logeventos Then PonLogEv("AXEditor_Rollover_Handler")
    End Sub

    Public Shared Sub AXEditor_SelectionAdded_Handler(ByVal sender As Object, ByVal e As SelectionAddedEventArgs)
        'AXDoc.Editor.WriteMessage("AXEditor_SelectionAdded_Handler")
        If logeventos Then PonLogEv("AXEditor_SelectionAdded_Handler")
    End Sub

    Public Shared Sub AXEditor_SelectionRemoved_Handler(ByVal sender As Object, ByVal e As SelectionRemovedEventArgs)
        'AXDoc.Editor.WriteMessage("AXEditor_SelectionRemoved_Handler")
        If logeventos Then PonLogEv("AXEditor_SelectionRemoved_Handler")
    End Sub
End Class

