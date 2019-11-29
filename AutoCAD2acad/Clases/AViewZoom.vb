Imports System.Diagnostics
Imports System.Collections
Imports System.Windows.Forms
Imports System.Drawing


Imports Autodesk.AutoCAD.Interop
Imports Autodesk.AutoCAD.Interop.Common
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.ApplicationServices
Imports oAppS = Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.EditorInput

Namespace A2acad
    Partial Public Class A2acad
        Public Sub Zoom_Seleccion()
            Autodesk.AutoCAD.Internal.Utils.ZoomObjects(False)
            oAppA.ZoomScaled(0.6, AcZoomScaleType.acZoomScaledRelative)
        End Sub
    End Class
End Namespace
