Imports Autodesk.AutoCAD.Interop.Common

Public Class Zona
    Public Property Handle As String = ""
    Public Property Nombre As String = ""

    Public Sub New(oP As AcadLWPolyline)
        Me.Handle = oP.Handle
        Dim NZona As String = XData.XLeeDato(oP.Handle, "Zona")
        Me.Nombre = If(NZona = "", Me.Handle, NZona)
    End Sub

    Public Function oPol() As AcadLWPolyline
        Return CType(oApp.ActiveDocument.HandleToObject(Me.Handle), AcadLWPolyline)
    End Function

End Class