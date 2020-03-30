Imports System.Linq
Imports ClosedXML.Excel

Public Class Etiqueta
    'Public IMPLACAD_DATA As String = "C:\ProgramData\IMPLACAD\"
    'REFERENCIA	TIPO	TIPO1	TIPO2	TIPO3	LARGO	ANCHO	STOCK	DESCRIPCION
    Public Property REFERENCIA As String = ""
    Public Property TIPO As String = ""
    Public Property TIPO1 As String = ""
    Public Property TIPO2 As String = ""
    Public Property TIPO3 As String = ""
    Public Property LARGO As String = ""
    Public Property ANCHO As String = ""
    Public Property STOCK As Boolean = True
    Public Property DESCRIPCION As String = ""
    Public Property COMPONENTES As String
    Public Property SUSTITUCION As String
    Public Property DWG As String = ""
    Public Property PNG As String = ""

    Public Sub New()
    End Sub
    Public Sub New(fila As IXLRow)
        Dim nFila As Integer = fila.RowNumber
        Me.STOCK = True
        Dim TempStock As String = ""
        For Each oCell As IXLCell In fila.Cells.AsParallel
            Dim cabecera As String = oCell.WorksheetColumn.FirstCell.Value.ToString.Trim
            Dim valor As String = IIf(oCell.Value Is Nothing, "", oCell.Value.ToString.Trim)
            Select Case cabecera.ToUpper
                Case "REFERENCIA" : Me.REFERENCIA = Convert.ToString(valor).Trim
                Case "TIPO" : Me.TIPO = Convert.ToString(valor).Trim
                Case "TIPO1" : Me.TIPO1 = Convert.ToString(valor).Trim
                Case "TIPO2" : Me.TIPO2 = Convert.ToString(valor).Trim
                Case "TIPO3" : Me.TIPO3 = Convert.ToString(valor).Trim
                Case "LARGO" : Me.LARGO = Convert.ToString(valor).Trim
                Case "ANCHO" : Me.ANCHO = Convert.ToString(valor).Trim
                Case "STOCK" : TempStock = Convert.ToString(valor).Trim
                Case "DESCRIPCION" : Me.DESCRIPCION = Convert.ToString(valor).Trim
                Case "COMPONENTES" : Me.COMPONENTES = Convert.ToString(valor).Trim.Replace(" ", "")
                Case "SUSTITUCION" : Me.SUSTITUCION = Convert.ToString(valor).Trim
            End Select
        Next
        If TempStock.ToUpper = "FALSE" Then Me.STOCK = False
        If Me.TIPO <> "" Then
            PonFullPath_DWGPNG()
        End If
    End Sub
    Public Sub New(ref As String, t As String, t1 As String, t2 As String, t3 As String, l As String, a As String, s As String, d As String)
        'REFERENCIA	TIPO	TIPO1	TIPO2	TIPO3	LARGO	ANCHO	STOCK	DESCRIPCION	PNG	DWG
        Me.REFERENCIA = ref
        Me.TIPO = t
        Me.TIPO1 = t1
        Me.TIPO2 = t2
        Me.TIPO3 = t3
        Me.LARGO = l
        Me.ANCHO = a
        Me.STOCK = True
        If s.ToUpper = "FALSE" Then Me.STOCK = False
        Me.DESCRIPCION = d
        PonFullPath_DWGPNG()
    End Sub
    '
    Private Sub PonFullPath_DWGPNG()
        DWG = IMPLACAD_DATA & IIf(IMPLACAD_DATA.EndsWith("\") = False, sep, "") & TIPO
        DWG &= IIf(TIPO1 <> "", sep & TIPO1, "")
        DWG &= IIf(TIPO2 <> "", sep & TIPO2, "")
        DWG &= IIf(TIPO3 <> "", sep & TIPO3, "")
        DWG &= sep & REFERENCIA & ".dwg"
        ' Si no existe el fichero DWG. Lo buscamos
        If IO.File.Exists(DWG) = False Then
            Dim fi As String() = IO.Directory.GetFiles(IMPLACAD_DATA, REFERENCIA & ".dwg", SearchOption.AllDirectories)
            If fi IsNot Nothing AndAlso fi.Count > 0 Then
                DWG = fi.First
            End If
        End If
        '
        PNG = IO.Path.ChangeExtension(DWG, ".png")
        ' Si no existe el fichero PNG. Lo buscamos
        If IO.File.Exists(PNG) = False Then
            Dim fi As String() = IO.Directory.GetFiles(IMPLACAD_DATA, REFERENCIA & ".png", SearchOption.AllDirectories)
            If fi IsNot Nothing AndAlso fi.Count > 0 Then
                PNG = fi.First
            End If
        End If
    End Sub
End Class
