Imports System.Diagnostics
Imports System.Collections
Imports System.Windows.Forms
Imports System.Drawing
Imports System.IO


Imports Autodesk.AutoCAD.Interop
Imports Autodesk.AutoCAD.Interop.Common
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.EditorInput

Namespace A2acad
    Public Class A2acad
        Public Function PropiedadCustomDocumento_Existe(quePro As String) As Boolean
            Dim resultado As Boolean = False
            Dim pKey As String = ""
            Dim pValue As String = ""
            Dim oSInf As AcadSummaryInfo = oAppA.ActiveDocument.SummaryInfo
            If oSInf.NumCustomInfo = 0 Then
                resultado = False
            Else
                For x As Integer = 0 To oSInf.NumCustomInfo - 1
                    oSInf.GetCustomByIndex(x, pKey, pValue)
                    If pKey.ToUpper = quePro.ToUpper Then
                        resultado = True
                        Exit For
                    End If
                Next
            End If
            Return resultado
        End Function

        Public Function PropiedadCustomDocumento_Lee(quePro As String) As String
            Dim queValor As String = ""
            If PropiedadCustomDocumento_Existe(quePro) = True Then
                Try
                    '' Leemos el valor o ""
                    oAppA.ActiveDocument.SummaryInfo.GetCustomByKey(quePro, queValor)
                Catch ex As System.Exception
                    queValor = ""
                    Debug.Print(ex.ToString)
                End Try
            End If
            ''
            Return queValor
        End Function

        Public Sub PropiedadCustomDocumento_Escribe(quePro As String, queVal As String, Optional crear As Boolean = True)
            Dim valoractual As String = ""
            '' Leemos el valor de estado.
            If PropiedadCustomDocumento_Existe(quePro) = True Then
                oAppA.ActiveDocument.SummaryInfo.GetCustomByKey(quePro, valoractual)
                If valoractual <> queVal Then
                    Try
                        oAppA.ActiveDocument.SummaryInfo.SetCustomByKey(quePro, queVal)
                    Catch ex As Exception
                        Debug.Print(ex.ToString)
                    End Try
                End If
            Else
                If crear = True Then
                    Try
                        '' Si no existía, la creamos
                        oAppA.ActiveDocument.SummaryInfo.AddCustomInfo(quePro, queVal)
                    Catch ex As Exception
                        Debug.Print(ex.ToString)
                    End Try
                End If
            End If
            ''
        End Sub
    End Class
End Namespace