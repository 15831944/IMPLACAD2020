Imports System.Windows.Forms
Imports Autodesk.AutoCAD.Interop
Imports Autodesk.AutoCAD.Interop.Common
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.ApplicationServices
Imports uau = UtilesAlberto.Utiles

Public Class frmZonas
    Public LEtiquetas As List(Of AcadBlockReference) = Nothing
    Private Sub frmZonas_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Text = "Administrar Zonas"
        LZona = New List(Of Zona)
        LlenaZonas()
        CapaCreaActivaZonas()
        BtnEditar.Enabled = False
    End Sub

    Private Sub frmZonas_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        FrmZ.Dispose()
        FrmZ = Nothing
    End Sub

    Public Sub LlenaZonas()
        LLwPolyline = DamePolilineasZonasImplacad(capa:="Zonas")
        If LLwPolyline IsNot Nothing AndAlso LLwPolyline.Count > 0 Then
            TvZonas.Nodes.Clear()
        Else
            Exit Sub
        End If
        '
        For Each OPol As AcadLWPolyline In LLwPolyline
            ' Solo Polilineas cerradas
            If OPol.Closed = False Then Continue For
            '
            Dim OZona As New Zona(OPol)
            Dim ONode As New TreeNode
            ONode.Text = If(OZona.Nombre <> "", OZona.Nombre, OZona.Handle)
            ONode.Tag = OZona.Handle
            TvZonas.Nodes.Add(ONode)
            ONode = Nothing
        Next
    End Sub

    Private Sub TvZonas_AfterSelect(sender As Object, e As TreeViewEventArgs) Handles TvZonas.AfterSelect
        If TvZonas.SelectedNode Is Nothing Then
            LblContar.Text = "XX etiquetas en Zona XXX"
            BtnEditar.Enabled = False
            Exit Sub
        End If
        BtnEditar.Enabled = True
        Dim h As String = TvZonas.SelectedNode.Tag.ToString ' Handled de la polilinea
        Dim oEnt As AcadObject = oApp.ActiveDocument.HandleToObject(h)
        'HazZoomObjeto(oEnt, 0.5)
        Dim objetos As List(Of Long) = DameEntitiesDentroPolilinea(CType(oEnt, AcadLWPolyline))
        LEtiquetas = New List(Of AcadBlockReference)
        '
        For Each id As Long In objetos
            Dim oE As AcadObject = oApp.ActiveDocument.ObjectIdToObject(id)
            Dim texto As String = XData.XLeeDato(oE, "Clase")
            If TypeOf oE Is AcadBlockReference And texto = "etiqueta" Then    ' "Clase=etiqueta"
                LEtiquetas.Add(oE)
            End If
            oE = Nothing
        Next

        LblContar.Text = LEtiquetas.Count & " etiquetas en Zona " & TvZonas.SelectedNode.Text
        objetos = Nothing
        HazZoomObjeto(oEnt, 0.5)
    End Sub

    Private Sub BtnCrear_Click(sender As Object, e As EventArgs) Handles BtnCrear.Click
        Dim UltimoHandle As String = oApp.ActiveDocument.ModelSpace.Item(oApp.ActiveDocument.ModelSpace.Count - 1).Handle
        Try
            Me.Visible = False
            Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.CurrentDocument.Editor.WriteMessage(vbLf)
            Dim command As String = "_pline "
            oApp.ActiveDocument.SendCommand(command)
            '
            '
            Dim CreadoHandle As String = oApp.ActiveDocument.ModelSpace.Item(oApp.ActiveDocument.ModelSpace.Count - 1).Handle
            ' Se ha creado un nuevo objeto. Ya no es igual el Handle que el ultimo antes de hacer nada.
            If CreadoHandle <> UltimoHandle Then
                Dim UltimaPolilinea As AcadLWPolyline = oApp.ActiveDocument.HandleToObject(CreadoHandle)
                If UltimaPolilinea.Closed Then
                    Dim NuevoNombreZona As String = InputBox("Nombre nueva Zona:", "ADMINISTRAR ZONAS")
                    If NuevoNombreZona <> "" AndAlso uau.Fichero_EsNombreCorrecto(NuevoNombreZona) Then
                        XData.XNuevo(UltimaPolilinea, "Zona=" & NuevoNombreZona)
                        'MsgBox(UltimaPolilinea.Length)
                        LlenaZonas()
                    End If
                Else
                    MsgBox("No esta cerrada")
                End If
            Else
                MsgBox("No se ha creado nada...")
            End If
        Catch ex As Exception
            Debug.Print(ex.ToString)
        Finally
            Me.Visible = True
        End Try
    End Sub

    Private Sub BtnEditar_Click(sender As Object, e As EventArgs) Handles BtnEditar.Click
        If TvZonas.SelectedNode Is Nothing Then Exit Sub
        Dim nombre As String = InputBox("Nombre Zona", "EDITAR ZONA", TvZonas.SelectedNode.Text)
        If nombre <> TvZonas.SelectedNode.Text Then
            Dim h As String = TvZonas.SelectedNode.Tag.ToString
            XData.XPonDato(h, "Zona", nombre)
            LlenaZonas()
        End If
    End Sub
End Class