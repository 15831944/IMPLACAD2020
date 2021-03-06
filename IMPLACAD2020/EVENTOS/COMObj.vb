﻿Imports System.Linq
Imports Autodesk.AutoCAD.Interop.Common

'
Partial Public Class Eventos

    Public Shared Sub Subscribre_COMObj(ByRef pObject As AcadObject)
        If pObject Is Nothing Then Exit Sub
        If lTypesCOMObj.Contains(pObject.ObjectName) = False Then Exit Sub
        '
        If TypeOf pObject Is AcadBlockReference Then
            'AddHandler pObject.Modified, AddressOf AcadBlockReference_Modified
        ElseIf TypeOf pObject Is AcadCircle Then
            'AddHandler pObject.Modified, AddressOf AcadCircle_Modified
        End If
    End Sub

    Public Shared Sub Unsubscribre_COMObj(pObject As AcadObject)
        Try
            'AXDoc.Editor.WriteMessage("COMDoc_Activate")
            If logeventos Then PonLogEv("COMDoc_Activate")
            If pObject Is Nothing Then Exit Sub
            If lTypesCOMObj.Contains(pObject.ObjectName) = False Then Exit Sub
            '
            If TypeOf pObject Is AcadBlockReference Then
                'RemoveHandler pObject.Modified, AddressOf AcadBlockReference_Modified
            ElseIf TypeOf pObject Is AcadCircle Then
                'RemoveHandler pObject.Modified, AddressOf AcadCircle_Modified
            End If
        Catch ex As System.Exception
            Debug.Print(ex.ToString)
        End Try
    End Sub

    Public Shared Sub AcadCircle_Modified(pObject As AcadObject)
        'AXDoc.Editor.WriteMessage("AcadCircle_Modified")
        If logeventos Then PonLogEv("AcadCircle_Modified")
        'Try
        '    Dim oCi As AcadCircle = CType(pObject, AcadCircle)
        '    EvDocM.CurrentDocument.Editor.WriteMessage(vbLf & "COM Radio: " & oCi.Radius)
        'Catch ex As System.Exception

        'End Try
    End Sub

    Public Shared Sub AcadBlockReference_Modified(pObject As AcadObject)
        'AXDoc.Editor.WriteMessage("AcadBlockReference_Modified")
        If logeventos Then PonLogEv("AcadBlockReference_Modified")
        'Try
        '    Dim oBl As AcadBlockReference = CType(pObject, AcadBlockReference)
        '    EvDocM.CurrentDocument.Editor.WriteMessage(vbLf & "COM Nombre: " & oBl.EffectiveNamef)
        'Catch ex As System.Exception

        'End Try
    End Sub

End Class