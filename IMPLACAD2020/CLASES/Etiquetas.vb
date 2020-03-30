Imports System.Linq
Imports ClosedXML.Excel

Public Class Etiquetas
    Public Property Campos As New List(Of String)
    Public Property LReferencias As New List(Of Etiqueta)
    Public Property LReferenciasString As New List(Of String)
    Public Property DReferencias As New Dictionary(Of String, Etiqueta)
    Public Property LConjuntos As New List(Of String)
    Public Property DConjuntos As New Dictionary(Of String, Etiqueta)
    Public Property DSustituciones As New Dictionary(Of String, String)
    Public Property LRows As New List(Of IXLRow)
    Public Sub New()
        LRows = modClosedXML.Excel_LeeFilas(appXLSX)
        For Each fila As IXLRow In LRows.AsParallel
            If fila.RowNumber = 1 Then
                For Each oCe As IXLCell In fila.CellsUsed
                    Campos.Add(Convert.ToString(oCe.Value))
                Next
            Else
                Dim oEti As New Etiqueta(fila)
                Dim ref As String = oEti.REFERENCIA
                ' Para que no incluya las que son sólo SUSTITUCIONES (No tienen TIPO)
                If oEti.TIPO <> "" Then
                    LReferencias.Add(oEti)
                    If LReferenciasString.Contains(ref) = False Then LReferenciasString.Add(ref)
                    If DReferencias.ContainsKey(ref) = False Then DReferencias.Add(ref, oEti)
                    If oEti.TIPO3.ToUpper.Contains("CONJUNTO") OrElse oEti.LARGO = "0" OrElse oEti.ANCHO = "0" Then
                        If LConjuntos.Contains(ref) = False Then
                            LConjuntos.Add(ref)
                            DConjuntos.Add(ref, oEti)
                        End If
                    End If
                End If
                If oEti.SUSTITUCION IsNot Nothing AndAlso oEti.SUSTITUCION.Trim <> "" Then
                    If DSustituciones.ContainsKey(oEti.REFERENCIA) = False Then
                        DSustituciones.Add(oEti.REFERENCIA, oEti.SUSTITUCION)
                    End If
                End If
                oEti = Nothing
            End If
            System.Windows.Forms.Application.DoEvents()
        Next
    End Sub
End Class


