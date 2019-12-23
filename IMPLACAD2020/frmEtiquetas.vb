Imports System
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.ApplicationServices
'Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Interop
Imports Autodesk.AutoCAD.Interop.Common
Imports System.Linq

Public Class frmEtiquetas
    Dim oTabla As DataTable
    Dim ultimoCaminoDWG As String
    Dim oEti As Etiqueta = Nothing
    Dim inicio As Boolean = True
    'Dim oTablaRef As DataTable

    Private Sub frmEtiquetas_FormClosed(sender As Object, e As Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Me.Dispose()
        frmE = Nothing
    End Sub

    Private Sub frmEtiquetas_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'oEs = New Etiquetas
        oTabla = UtilesAlberto.Utiles.Excel_DameDatatable(appXLSX)
        oTabla.TableName = "ETIQUETAS"
        Dim claves(0) As DataColumn
        claves(0) = oTabla.Columns("REFERENCIA")
        oTabla.PrimaryKey = claves

        'Dim t As System.Threading.Tasks = New System.Threading.Tasks(AddressOf RellenaListados(oTabla))
        RellenaListados(oTabla)
        lblCambiar.Text = ""
        btnCambiar.Visible = False
        btnInsertar.Visible = True
        bloqueEditar = Nothing
        Me.Text = My.Application.Info.ProductName & " - v." & My.Application.Info.Version.ToString
        Me.cbTIPO.SelectedItem = Tipo
        Me.cbTIPO1.SelectedItem = Tipo1
        Me.cbTIPO2.SelectedItem = Tipo2
        Me.cbTIPO3.SelectedItem = Tipo3
    End Sub

    Private Sub btnCerrar_Click(sender As Object, e As EventArgs) Handles btnCerrar.Click
        cIni.IniWrite(nIni, "OPCIONES", "Tipo", cbTIPO.Text)
        cIni.IniWrite(nIni, "OPCIONES", "Tipo1", cbTIPO1.Text)
        cIni.IniWrite(nIni, "OPCIONES", "Tipo2", cbTIPO2.Text)
        cIni.IniWrite(nIni, "OPCIONES", "Tipo3", cbTIPO3.Text)
        Me.Close()
    End Sub

    Private Sub btnInsertar_Click(sender As Object, e As EventArgs) Handles btnInsertar.Click
        If IO.Directory.Exists(IMPLACAD_DATA) = False Then
            MsgBox("No existe el directorio base --> " & IMPLACAD_DATA & vbCrLf & vbCrLf &
                   "Imposible continuar. Verifique fichero .ini con configuración")
        End If
        Dim PathDwg As String = oEti.DWG
        Dim referencia As String = oEti.REFERENCIA
        If IO.File.Exists(PathDwg) = True Then
            ultimoCaminoDWG = PathDwg
            Try
                Me.Visible = False
                If oApp Is Nothing Then _
                   oApp = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
                'oDoc = oApp.ActiveDocument
                ConfiguraDibujo()
                '' Poner capa 0 como activa.
                CapaCeroActiva()
                CapaZonaCoberturaACTDES(CapaEstado.Desactivar)      '' Crear y desactivar las 2 zonas de cobertura.
                Dim inserta As AcadBlockReference = Nothing
                Dim ultimoBlo As Autodesk.AutoCAD.Interop.Common.AcadEntity = Nothing
                oApp.ActiveDocument.SendCommand(
                "(Command ""-insert"" """ & ultimoCaminoDWG.Replace("\", "\\") & """ pause ""0.001"" ""0.001"" pause)" & vbCr)
                ''
                Try
                    ultimoBlo = oApp.ActiveDocument.ModelSpace.Item(oApp.ActiveDocument.ModelSpace.Count - 1)
                Catch ex As System.Exception
                    Me.Visible = True
                    Exit Sub
                End Try
                ''
                '' Si no insertamos nada (anulamos inserción) salimos sin más
                If ultimoBlo Is Nothing Then
                    Me.Visible = True
                    Exit Sub
                ElseIf Not (TypeOf ultimoBlo Is AcadBlockReference) Then
                    Me.Visible = True
                    Exit Sub
                ElseIf (TypeOf ultimoBlo Is AcadBlockReference) AndAlso CType(ultimoBlo, AcadBlockReference).Name <> referencia Then
                    Me.Visible = True
                    Exit Sub
                ElseIf (TypeOf ultimoBlo Is AcadBlockReference) AndAlso CType(ultimoBlo, AcadBlockReference).Name = referencia Then
                    inserta = oApp.ActiveDocument.ObjectIdToObject(ultimoBlo.ObjectID)
                    '' Ponerle XData al bloque
                    XData.XNuevo(inserta, "Clase=etiqueta") ';Tipo=" & txtTIPO.Text & ";Tipo1=" & txtTIPO1.Text & ";Tipo2=" & txtTIPO2.Text & ";Tipo3=" & txtTIPO3.Text)
                End If
                ''
                Me.Visible = True
                CapaZonaCoberturaACTDES(CapaEstado.Activar)
                oApp.ActiveDocument.Regen(AcRegenType.acActiveViewport)
                'End If
            Catch ex As System.Exception
                'Me.WindowState = Windows.Forms.FormWindowState.Normal
                MsgBox("Error insertando " & ultimoCaminoDWG)
            Finally
                'Me.Visible = True
                XRef_IMGListar(False)
            End Try
        Else
            MsgBox("No hay DWG asociado...")
        End If
    End Sub


    Private Sub btnCambiar_Click(sender As Object, e As EventArgs) Handles btnCambiar.Click
        If IO.Directory.Exists(IMPLACAD_DATA) = False Then
            MsgBox("No existe el directorio base --> " & IMPLACAD_DATA & vbCrLf & vbCrLf &
                   "Imposible continuar. Verifique fichero .ini con configuración")
        End If
        Dim PathDwg As String = oEti.DWG
        If IO.File.Exists(PathDwg) = True Then
            ultimoCaminoDWG = PathDwg
            Try
                Me.Visible = False
                Dim oApp As Autodesk.AutoCAD.Interop.AcadApplication =
                    CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication,
    Autodesk.AutoCAD.Interop.AcadApplication)
                '' Poner capa 0 como activa.
                If oApp.ActiveDocument.ActiveLayer.Name <> "0" Then
                    oApp.ActiveDocument.ActiveLayer = oApp.ActiveDocument.Layers.Item("0")
                End If
                '
                ' Insertar el bloque seleccionado en el punto de inserción del existente.
                'Dim resultado As PromptPointResult = Application.DocumentManager.CurrentDocument.Editor.GetPoint("Punto de insercion :")
                'If resultado IsNot Nothing Then
                Dim puntoInserta As Object = bloqueEditar.InsertionPoint
                Dim giroInserta As Double = bloqueEditar.Rotation
                Dim escalaInserta As Double = bloqueEditar.XScaleFactor
                ''
                Dim inserta As AcadBlockReference = Nothing
                ''
                '' Borramos el bloque anterior.
                bloqueEditar.Delete()
                bloqueEditar = Nothing
                ''
                inserta = oApp.ActiveDocument.ActiveLayout.Block.InsertBlock(puntoInserta, ultimoCaminoDWG, escalaInserta, escalaInserta, escalaInserta, giroInserta)
                ''
                XRef_IMGListar(False)
                '' Configurar las capas de cobertura desactivada e inutilizada.
                Dim oLayer As AcadLayer = Nothing
                Try
                    oLayer = oApp.ActiveDocument.Layers.Item("Zona Cobertura Evacuación")
                    If oLayer.LayerOn = True Then oLayer.LayerOn = False
                    If oLayer.Freeze = False Then oLayer.Freeze = True
                    oLayer = oApp.ActiveDocument.Layers.Item("Zona Cobertura Extincion")
                    If oLayer.LayerOn = True Then oLayer.LayerOn = False
                    If oLayer.Freeze = False Then oLayer.Freeze = True
                Catch ex As System.Exception
                    ''
                End Try
                oLayer = Nothing
                ''
                '' Ponerle XData al bloque
                '' Tipo;Tipo1;Tipo2;Tipo3
                XData.XNuevo(inserta, "Clase=etiqueta") ';Tipo=" & txtTIPO.Text & ";Tipo1=" & txtTIPO1.Text & ";Tipo2=" & txtTIPO2.Text & ";Tipo3=" & txtTIPO3.Text)
                ''
                oApp.ActiveDocument.Regen(AcRegenType.acActiveViewport)
            Catch ex As System.Exception
                'Me.WindowState = Windows.Forms.FormWindowState.Normal
                MsgBox("Error insertando " & ultimoCaminoDWG)
            Finally
                Me.Visible = True
            End Try
            lblCambiar.Text = ""
            btnCambiar.Visible = False
            btnInsertar.Visible = True
        End If
    End Sub

    Private Sub lbETIQUETAS_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lbETIQUETAS.SelectedIndexChanged
        VaciaFicha()

        If IO.Directory.Exists(IMPLACAD_DATA) = False Then
            MsgBox("No existe el directorio base --> " & IMPLACAD_DATA & vbCrLf & vbCrLf &
                   "Imposible continuar. Verifique fichero .ini con configuración")
        End If
        If lbETIQUETAS.SelectedIndex = -1 Then Exit Sub
        oEti = oEtis.DReferencias(lbETIQUETAS.Text)
        If oEti Is Nothing Then Exit Sub
        '
        Me.txtREFERENCIA.Text = oEti.REFERENCIA
        Me.txtTIPO.Text = oEti.TIPO
        Me.txtTIPO1.Text = oEti.TIPO1
        Me.txtTIPO2.Text = oEti.TIPO2
        Me.txtTIPO3.Text = oEti.TIPO3
        Me.txtDESCRIPCION.Text = oEti.DESCRIPCION
        Me.cbStock.Checked = oEti.STOCK
        Me.lbDim.Text = "Dimensiones = " & oEti.LARGO.ToString & " x " & oEti.ANCHO.ToString
        If IO.File.Exists(oEti.PNG) Then
            pbIMAGEN.Image = Drawing.Image.FromFile(oEti.PNG)
        Else
            pbIMAGEN.Image = Nothing
        End If
        Dim PathDwg As String = oEti.DWG
        If IO.File.Exists(PathDwg) = True Then
            btnInsertar.Enabled = True
        Else
            btnInsertar.Enabled = False
        End If
    End Sub
    ''
    Public Sub VaciaFicha()
        Me.txtREFERENCIA.Text = ""
        Me.txtTIPO.Text = ""
        Me.txtTIPO1.Text = ""
        Me.txtTIPO2.Text = ""
        Me.txtTIPO3.Text = ""
        Me.txtDESCRIPCION.Text = ""
        Me.cbStock.Checked = False
        Me.lbDim.Text = ""
        pbIMAGEN.Image = Nothing
    End Sub
    ''
#Region "RELLENAR LISTADOS"
    Public Sub RellenaListados(ByVal queTabla As DataTable)
        If LConjuntos Is Nothing Then LConjuntos = New List(Of String)
        If queTabla.Rows.Count = 0 Then Exit Sub
        ''
        '' Borrar los listados y añadir * en los combobox.
        Me.lbETIQUETAS.Items.Clear()
        cbTIPO.Items.Clear() : cbTIPO.Items.Add("*")
        cbTIPO1.Items.Clear() : cbTIPO1.Items.Add("*")
        cbTIPO2.Items.Clear() : cbTIPO2.Items.Add("*")
        cbTIPO3.Items.Clear() : cbTIPO3.Items.Add("*")
        ''
        '' Recorrer cada fila para sacar Referencia, Tipo, Tipo1, Tipo2 y Tipo3
        For Each quefila As DataRow In queTabla.Rows
            'REFERENCIA	TIPO	TIPO1	TIPO2	TIPO3	LARGO	ANCHO	STOCK	DESCRIPCION	PNG	DWG
            Dim Referencia As String = quefila("REFERENCIA").ToString
            Dim Tipo As String = quefila("TIPO").ToString
            Dim Tipo1 As String = quefila("TIPO1").ToString
            Dim Tipo2 As String = quefila("TIPO2").ToString
            Dim Tipo3 As String = quefila("TIPO3").ToString
            Dim Largo As String = quefila("LARGO").ToString
            Dim Ancho As String = quefila("ANCHO").ToString
            Dim STOCK As String = quefila("STOCK").ToString
            Dim DESCRIPCION As String = quefila("DESCRIPCION").ToString
            Dim oEti As New Etiqueta(Referencia, Tipo, Tipo1, Tipo2, Tipo3, Largo, Ancho, STOCK, DESCRIPCION)
            'oEti = Nothing
            If (Tipo3.ToUpper = "CONJUNTO" OrElse Largo = "0" OrElse Ancho = "0") AndAlso
                LConjuntos.Contains(Referencia) = False Then
                LConjuntos.Add(Referencia)
            End If
            '
            If Me.lbETIQUETAS.Items.Contains(Referencia) = False Then _
                Me.lbETIQUETAS.Items.Add(Referencia)
            If cbTIPO.Items.Contains(Tipo) = False Then cbTIPO.Items.Add(Tipo)
            If cbTIPO1.Items.Contains(Tipo1) = False Then cbTIPO1.Items.Add(Tipo1)
            If cbTIPO2.Items.Contains(Tipo2) = False Then cbTIPO2.Items.Add(Tipo2)
            If cbTIPO3.Items.Contains(Tipo3) = False Then cbTIPO3.Items.Add(Tipo3)
        Next
        ''
        '' Ordenar listados y poner valores por defecto.
        Me.lbETIQUETAS.Sorted = True
        Me.lbETIQUETAS.SelectedIndex = -1
        cbTIPO.Sorted = True : cbTIPO.SelectedItem = "*"
        cbTIPO1.Sorted = True : cbTIPO1.SelectedItem = "*"
        cbTIPO2.Sorted = True : cbTIPO2.SelectedItem = "*"
        cbTIPO3.Sorted = True : cbTIPO3.SelectedItem = "*"
        inicio = False
        lbCuantos.Text = lbETIQUETAS.Items.Count & " Etiquetas"
        VaciaFicha()
    End Sub
#End Region

#Region "FILTRAR DATOS"
    Public Sub FiltraListado()
        If oEtis Is Nothing Then Exit Sub
        If oEtis.LReferencias Is Nothing OrElse oEtis.LReferencias.Count = 0 Then Exit Sub
        '
        ' Borrar listado etiquetas solo.
        Me.lbETIQUETAS.Items.Clear()
        Dim filtro As String = ""
        Dim fTipo As String = ""
        Dim fTipo1 As String = ""
        Dim fTipo2 As String = ""
        Dim fTipo3 As String = ""
        '' Todos los objetos Etiqueta en oFiltro
        Dim oFiltro As List(Of Etiqueta) = oEtis.LReferencias.Where(Function(p) p.REFERENCIA <> "").ToList
        Dim oFiltroFin As List(Of Etiqueta) = oFiltro.ToList
        '
        If cbTIPO.Text <> "*" Then
            oFiltroFin = oFiltro.Where(Function(p) p.TIPO.Contains(cbTIPO.Text)).ToList
            oFiltro = oFiltroFin
        End If
        If cbTIPO1.Text <> "*" Then
            oFiltroFin = oFiltro.Where(Function(p) p.TIPO1.Contains(cbTIPO1.Text)).ToList
            oFiltro = oFiltroFin
        End If
        If cbTIPO2.Text <> "*" AndAlso cbTIPO2.Text <> "" Then
            oFiltroFin = oFiltro.Where(Function(p) p.TIPO2.Contains(cbTIPO2.Text)).ToList
            oFiltro = oFiltroFin
        End If
        If cbTIPO3.Text <> "*" AndAlso cbTIPO3.Text <> "" Then
            oFiltroFin = oFiltro.Where(Function(p) p.TIPO3.Contains(cbTIPO3.Text)).ToList
            oFiltro = oFiltroFin
        End If
        '
        Me.lbETIQUETAS.Items.AddRange((From x In oFiltro.ToList Select x.REFERENCIA).ToArray)
        '
        'For Each eti As Etiqueta In oFiltro.ToList
        '    Me.lbETIQUETAS.Items.Add(eti.REFERENCIA)
        'Next
        ' Ordenar listados y poner valores por defecto.
        Me.lbETIQUETAS.Sorted = True
        Me.lbETIQUETAS.SelectedIndex = -1
        lbCuantos.Text = lbETIQUETAS.Items.Count & " Etiquetas"
        VaciaFicha()
    End Sub
    ''

    Public Sub FiltraListado_Tabla()
        If oTabla.Rows.Count = 0 Then Exit Sub
        ''
        '' Borrar listado etiquetas solo.
        Me.lbETIQUETAS.Items.Clear()
        Dim filtro As String = ""
        Dim fTipo As String = ""
        Dim fTipo1 As String = ""
        Dim fTipo2 As String = ""
        Dim fTipo3 As String = ""
        ''
        If cbTIPO.Text = "*" Then
            fTipo &= "TIPO like '%' AND "
        Else
            fTipo &= "TIPO = '" & cbTIPO.Text & "' AND "
        End If
        If cbTIPO1.Text = "*" Then
            fTipo1 &= "TIPO1 like '%' AND "
        Else
            fTipo1 &= "TIPO1 = '" & cbTIPO1.Text & "' AND "
        End If
        If cbTIPO2.Text = "*" Then
            fTipo2 &= "TIPO2 like '%' AND "
        Else
            fTipo2 &= "TIPO2 = '" & cbTIPO2.Text & "' AND "
        End If
        If cbTIPO3.Text = "*" Then
            fTipo3 &= "TIPO3 like '%'"
        Else
            fTipo3 &= "TIPO3 = '" & cbTIPO3.Text & "'"
        End If
        filtro = fTipo & fTipo1 & fTipo2 & fTipo3
        Dim filas() As DataRow = oTabla.Select(filtro)
        ''
        '' Recorrer cada fila para sacar Referencia, Tipo, Tipo1, Tipo2 y Tipo3
        For Each quefila As DataRow In filas
            Dim Referencia As String = quefila("REFERENCIA").ToString
            If Me.lbETIQUETAS.Items.Contains(Referencia) = False Then _
                Me.lbETIQUETAS.Items.Add(Referencia)
        Next
        ''
        '' Ordenar listados y poner valores por defecto.
        Me.lbETIQUETAS.Sorted = True
        Me.lbETIQUETAS.SelectedIndex = -1
        lbCuantos.Text = lbETIQUETAS.Items.Count & " Etiquetas"
        VaciaFicha()
    End Sub
    ''
    Private Sub cbTIPO_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbTIPO.SelectedIndexChanged, cbTIPO1.SelectedIndexChanged, cbTIPO2.SelectedIndexChanged, cbTIPO3.SelectedIndexChanged
        If inicio = True Then Exit Sub
        FiltraListado()
        txtBusca.Text = ""
        If cIni Is Nothing Then cIni = New clsINI
        Dim ocb As System.Windows.Forms.ComboBox = sender
        If ocb.Name = "cbTIPO" Then
            cIni.IniWrite(nIni, "OPCIONES", "Tipo", cbTIPO.SelectedItem.ToString)
        ElseIf ocb.Name = "cbTIPO1" Then
            cIni.IniWrite(nIni, "OPCIONES", "Tipo1", cbTIPO1.SelectedItem.ToString)
        ElseIf ocb.Name = "cbTIPO2" Then
            cIni.IniWrite(nIni, "OPCIONES", "Tipo2", cbTIPO2.SelectedItem.ToString)
        ElseIf ocb.Name = "cbTIPO3" Then
            cIni.IniWrite(nIni, "OPCIONES", "Tipo3", cbTIPO3.SelectedItem.ToString)
        End If
    End Sub
#End Region

    Private Sub btnSel_Click(sender As Object, e As EventArgs) Handles btnSel.Click
        '' Selecciona todos los bloques.
        modImplacad.SeleccionaBloquesImplacad()
    End Sub

    Private Sub btnEditar_Click(sender As Object, e As EventArgs) Handles btnEditar.Click
        Me.Visible = False
        ''
        Try
            bloqueEditar = modImplacad.SeleccionaDameBloque()
            'While bloqueEditar Is Nothing
            'bloqueEditar = modImplacad.SeleccionaDameBloque()
            'End While
            If bloqueEditar IsNot Nothing Then
                'bloqueEditar = obl.Name
                'bloqueID = obl.ObjectID
                'MsgBox("Nombre = " & bloqueEditar & vbCrLf & "ID = " & bloqueID)
                cbTIPO.SelectedItem = "*"
                cbTIPO1.SelectedItem = "*"
                cbTIPO2.SelectedItem = "*"
                cbTIPO3.SelectedItem = "*"
                Me.lbETIQUETAS.SelectedItem = bloqueEditar.Name
                lblCambiar.Text = "Se cambiará el bloque seleccionado : " & bloqueEditar.Name
                btnCambiar.Visible = True
                btnInsertar.Visible = False
            Else
                Me.lbETIQUETAS.SelectedIndex = -1
                lblCambiar.Text = ""
                btnCambiar.Visible = False
                btnInsertar.Visible = True
            End If
        Catch ex As System.Exception
            ''
            MsgBox(ex.Message)
        End Try
        ''
        Me.Visible = True
    End Sub

    Private Sub txtBusca_TextChanged(sender As Object, e As EventArgs) Handles txtBusca.TextChanged
        'If txtBusca.Text = "" Then Exit Sub
        'Dim indice As Integer = lbETIQUETAS.FindString(txtBusca.Text.ToUpper, -1)
        'If indice > -1 Then
        '    lbETIQUETAS.SetSelected(indice, True)
        '    lbETIQUETAS.SelectedIndex = indice
        'End If
        txtBusca.Text = txtBusca.Text.ToUpper
    End Sub

    Private Sub BtnBuscar_Click(sender As Object, e As EventArgs) Handles BtnBuscar.Click
        ' Busca solo en Resultados
        'If txtBusca.Text = "" Then Exit Sub
        'Dim indice As Integer = lbETIQUETAS.FindString(txtBusca.Text.ToUpper, -1)
        'If indice > -1 Then
        '    lbETIQUETAS.SetSelected(indice, True)
        '    lbETIQUETAS.SelectedIndex = indice
        'End If
        '
        ' Busca en todas las etiquetas
        inicio = True
        lbETIQUETAS.Items.Clear()
        cbTIPO.Text = "*"
        cbTIPO1.Text = "*"
        cbTIPO2.Text = "*"
        cbTIPO3.Text = "*"
        txtBusca.Text = txtBusca.Text.ToUpper
        If oEtis Is Nothing Then Exit Sub
        If oEtis.LReferencias Is Nothing OrElse oEtis.LReferencias.Count = 0 Then Exit Sub
        '
        ' Borrar listado etiquetas solo.
        Me.lbETIQUETAS.Items.Clear()
        ' Todos los objetos Etiqueta
        Dim oFiltro As List(Of Etiqueta) = oEtis.LReferencias.Where(Function(p) p.REFERENCIA.Contains(txtBusca.Text)).ToList
        '
        If oFiltro IsNot Nothing AndAlso oFiltro.Count > 0 Then
            Me.lbETIQUETAS.Items.AddRange((From x In oFiltro.ToList Select x.REFERENCIA).ToArray)
        End If
        '
        ' Ordenar listados y poner valores por defecto.
        Me.lbETIQUETAS.Sorted = True
        Me.lbETIQUETAS.SelectedIndex = -1
        lbCuantos.Text = lbETIQUETAS.Items.Count & " Etiquetas"
        VaciaFicha()
        inicio = False
    End Sub
End Class