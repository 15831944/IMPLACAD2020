Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Interop.Common

Module modPolilineas

    Public Function Polilinea_DameBloquesEtiquetasDentro(oPol As AcadLWPolyline) As List(Of BlockReference)
        Dim resultado As New List(Of BlockReference)
        '
        Dim objetos As List(Of Long) = DameEntitiesDentroPolilinea(oPol)
        '
        For Each id As Long In objetos
            Dim oE As AcadObject = oApp.ActiveDocument.ObjectIdToObject(id)
            If TypeOf oE IsNot AcadBlockReference Then Continue For
            '
            Dim texto As String = XData.XLeeDato(oE, "Clase")
            If texto = "etiqueta" Then    ' "Clase=etiqueta"
                resultado.Add(oE)
            End If
            oE = Nothing
        Next
        objetos = Nothing
        'HazZoomObjeto(oEnt, 0.5)
        Return resultado
    End Function

End Module