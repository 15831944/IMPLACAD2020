Imports System
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.ApplicationServices
'Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Interop
Imports Autodesk.AutoCAD.Interop.Common

'Imports Microsoft.SqlServer.Server
Imports System.Data.SqlServerCe

Public Class frmEtiquetas
    Dim oTabla As DataTable
    Dim ultimoCaminoDWG As String
    Dim ultimaFila As DataRow = Nothing
    Dim inicio As Boolean = True
    'Dim oTablaRef As DataTable

    Private Sub frmEtiquetas_FormClosed(sender As Object, e As Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Me.Dispose()
        frmE = Nothing
    End Sub

    Private Sub frmEtiquetas_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Dim connString As String = "Data Source='IMPLACAD.sdf'; LCID=1033; Password=Bgik8$l.; Encrypt = TRUE;"
        Dim source As String = My.Application.Info.DirectoryPath & "\" & My.Application.Info.AssemblyName & ".sdf"
        'Dim connString As String = "Data Source='IMPLACAD.sdf'; LCID=1033; Password=; Encrypt = TRUE;"
        Dim connString As String = "Data Source='" & source & "'; LCID=1033; Password=; Encrypt = TRUE;"

        'Dim engine As New SqlCeEngine(connString)
        Dim conexion As New SqlCeConnection(connString)
        conexion.Open()
        ''
        Dim oCom As New SqlCeCommand("Select ID, REFERENCIA, TIPO, TIPO1, TIPO2, TIPO3, LARGO, ANCHO, STOCK, DESCRIPCION, PNG, DWG from ETIQUETAS", conexion)
        'Dim oCom1 As New SqlCeCommand("Select REFERENCIA from ETIQUETAS", conexion)
        'MsgBox(oCom.ExecuteNonQuery())
        Dim oDa As New SqlCeDataAdapter(oCom)
        'Dim oDa1 As New SqlCeDataAdapter(oCom1)
        'oTabla = New Data.DataTable("ETIQUETAS")
        Dim oDs As New DataSet
        'oTablaRef = New Data.DataTable("REFERENCIAS")
        oDa.Fill(oDs, "ETIQUETAS")
        oTabla = oDs.Tables("ETIQUETAS")
        Dim claves(0) As DataColumn
        claves(0) = oTabla.Columns("REFERENCIA")
        oTabla.PrimaryKey = claves
        'oDa1.Fill(oTablaRef)
        'MsgBox(oTabla.Rows.Count)
        conexion.Close()
        'Me.lbETIQUETAS.DisplayMember = oTabla.Columns("REFERENCIA").ColumnName
        'Me.lbETIQUETAS.DataSource = oTabla
        RellenaListados(oTabla)   ', "REFERENCIA")
        lblCambiar.Text = ""
        btnCambiar.Visible = False
        btnInsertar.Visible = True
        bloqueEditar = Nothing
        Me.Text = My.Application.Info.ProductName & " - v." & My.Application.Info.Version.ToString
        INICargar()
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
        If IO.Directory.Exists(dirBase) = False Then
            MsgBox("No existe el directorio base --> " & dirBase & vbCrLf & vbCrLf & _
                   "Imposible continuar. Verifique fichero .ini con configuración")
        End If
        Dim dato As Object = dirBase & ultimaFila("DWG")
        Dim referencia As Object = ultimaFila("REFERENCIA")
        If IsDBNull(dato) = False AndAlso IO.File.Exists(dato.ToString) = True Then
            'Dim dwg As System.Byte()
            'dwg = CType(ultimaFila("DWG"), System.Byte())
            'ultimoCaminoDWG = "C:\Temp\" & Me.txtREFERENCIA.Text & ".dwg"
            ultimoCaminoDWG = dato.ToString
            'Bytes2EscribeFichero(dwg, ultimoCaminoDWG)
            'MsgBox("Se han escrito en C:\Temp " & dwg.GetLength(0) & " bites")
            'application
            'Application.DocumentManager.CurrentDocument.Editor
            Try
                Me.Visible = False
                If oApp Is Nothing Then _
                   oApp = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
                'oDoc = oApp.ActiveDocument
                ConfiguraDibujo()
                '' Poner capa 0 como activa.
                CapaCeroActiva()
                CapaZonaCoberturaACTDES(CapaEstado.Desactivar)      '' Crear y desactivar las 2 zonas de cobertura.
                'oApp.ActiveDocument.ActiveLayer = oApp.ActiveDocument.Layers.Item("0")
                ''
                'Me.Visible = False
                ''
                '' Insertar el bloque
                'Dim resultado As PromptPointResult = Application.DocumentManager.CurrentDocument.Editor.GetPoint("Punto de insercion :")
                'If resultado IsNot Nothing Then
                'Dim oPoint As Point3d = resultado.Value
                'Dim puntoInserta(0 To 2) As Double
                'puntoInserta(0) = oPoint.X
                'puntoInserta(1) = oPoint.Y
                'puntoInserta(2) = oPoint.Z
                Dim inserta As AcadBlockReference = Nothing
                Dim ultimoBlo As Autodesk.AutoCAD.Interop.Common.AcadEntity = Nothing
                'inserta = oApp.ActiveDocument.ActiveLayout.Block.InsertBlock(puntoInserta, ultimoCaminoDWG, escalaTotal, escalaTotal, escalaTotal, 0)
                'oApp.ActiveDocument.SendCommand("-insert " & ultimoCaminoDWG & " ") 'pause 1  pause")
                'oApp.ActiveDocument.SendCommand( _
                '    "-insert" & vbCr & _
                '    ultimoCaminoDWG & vbCr & _
                '    "pause" & vbCr & _
                '     vbCr & vbCr & _
                '     "pause")
                oApp.ActiveDocument.SendCommand( _
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
                ''
                '' Configurar la capa de cobertura desactivada e inutilizada.
                'Dim oLayer As AcadLayer = Nothing
                'Try
                '    oLayer = oApp.ActiveDocument.Layers.Item("Zona Cobertura Evacuación")
                '    If oLayer.LayerOn = True Then oLayer.LayerOn = False
                '    If oLayer.Freeze = False Then oLayer.Freeze = True
                'Catch ex As System.Exception
                '    ''
                'End Try
                'Try
                '    oLayer = oApp.ActiveDocument.Layers.Item("Zona Cobertura Extincion")
                '    If oLayer.LayerOn = True Then oLayer.LayerOn = False
                '    If oLayer.Freeze = False Then oLayer.Freeze = True
                'Catch ex As System.Exception
                '    ''
                'End Try
                'oLayer = Nothing
                ''''
                '' Girar el bloque insertado
                'Try
                '    Dim cadenaPunto As String = inserta.InsertionPoint(0) & "," & inserta.InsertionPoint(1) & "," & inserta.InsertionPoint(2)
                '    oApp.ActiveDocument.SendCommand("GIRA LT  " & cadenaPunto & " pause ")
                'Catch ex As System.Exception
                '    '' no hacemos nada.
                '    Debug.Print(ex.Message)
                'End Try
                CapaZonaCoberturaACTDES(CapaEstado.Activar)
                oApp.ActiveDocument.Regen(AcRegenType.acActiveViewport)
                'End If
            Catch ex As System.Exception
                'Me.WindowState = Windows.Forms.FormWindowState.Normal
                MsgBox("Error insertando " & ultimoCaminoDWG)
            Finally
                'Me.Visible = True
            End Try
        Else
            MsgBox("No hay DWG asociado...")
        End If
        'Me.DialogResult = Windows.Forms.DialogResult.OK
        'Me.Close()
        'CapaZonaCoberturaACTDES(CapaEstado.Desactivar)      '' Crear y desactivar las 2 zonas de cobertura.
    End Sub


    Private Sub btnCambiar_Click(sender As Object, e As EventArgs) Handles btnCambiar.Click
        If IO.Directory.Exists(dirBase) = False Then
            MsgBox("No existe el directorio base --> " & dirBase & vbCrLf & vbCrLf & _
                   "Imposible continuar. Verifique fichero .ini con configuración")
        End If
        Dim dato As Object = dirBase & ultimaFila("DWG")
        If IsDBNull(dato) = False AndAlso IO.File.Exists(dato.ToString) = True Then
            'Dim dwg As System.Byte()
            'dwg = CType(ultimaFila("DWG"), System.Byte())
            'ultimoCaminoDWG = "C:\Temp\" & Me.txtREFERENCIA.Text & ".dwg"
            ultimoCaminoDWG = dato.ToString
            'Bytes2EscribeFichero(dwg, ultimoCaminoDWG)
            'MsgBox("Se han escrito en C:\Temp " & dwg.GetLength(0) & " bites")
            'application
            'Application.DocumentManager.CurrentDocument.Editor
            Try
                Me.Visible = False
                Dim oApp As Autodesk.AutoCAD.Interop.AcadApplication = _
                    CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, 
    Autodesk.AutoCAD.Interop.AcadApplication)
                '' Poner capa 0 como activa.
                If oApp.ActiveDocument.ActiveLayer.Name <> "0" Then
                    oApp.ActiveDocument.ActiveLayer = oApp.ActiveDocument.Layers.Item("0")
                End If
                ''
                'Me.Visible = False
                ''
                '' Insertar el bloque seleccionado en el punto de inserción del existente.
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
        If IO.Directory.Exists(dirBase) = False Then
            MsgBox("No existe el directorio base --> " & dirBase & vbCrLf & vbCrLf & _
                   "Imposible continuar. Verifique fichero .ini con configuración")
        End If
        If lbETIQUETAS.SelectedIndex = -1 Then Exit Sub
        ultimaFila = oTabla.Rows.Find(lbETIQUETAS.Text)
        If ultimaFila Is Nothing Then Exit Sub
        Me.txtREFERENCIA.Text = ultimaFila("REFERENCIA").ToString
        Me.txtTIPO.Text = ultimaFila("TIPO").ToString
        Me.txtTIPO1.Text = ultimaFila("TIPO1").ToString
        Me.txtTIPO2.Text = ultimaFila("TIPO2").ToString
        Me.txtTIPO3.Text = ultimaFila("TIPO3").ToString
        Me.txtDESCRIPCION.Text = ultimaFila("DESCRIPCION").ToString
        'Me.cbStock.Enabled = CBool(ultimaFila("STOCK").ToString)
        Me.cbStock.Checked = CBool(ultimaFila("STOCK").ToString)
        Me.lbDim.Text = "Dimensiones = " & ultimaFila("LARGO").ToString & " x " & ultimaFila("ANCHO").ToString
        If IO.File.Exists(dirBase & ultimaFila("PNG")) Then
            ''
            '' Si la imagen está almacenada en la BD
            ''Me.pbIMAGEN.Image = Bytes2Image(ultimaFila("PNG"))
            ''
            '' Si la imagen la cargamos del disco duro.
            pbIMAGEN.Image = Drawing.Image.FromFile(dirBase & ultimaFila("PNG").ToString)
        Else
            pbIMAGEN.Image = Nothing
        End If
        Dim dato As Object = dirBase & ultimaFila("DWG")
        If IsDBNull(dato) = False AndAlso IO.File.Exists(dato.ToString) = True Then
            btnInsertar.Enabled = True
            '    Dim dwg As System.Byte()
            '    dwg = CType(ultimaFila("DWG"), System.Byte())
            '    ultimoCaminoDWG = "C:\Temp\" & Me.txtREFERENCIA.Text & ".dwg"
            '    Bytes2EscribeFichero(dwg, ultimoCaminoDWG)
            '    'MsgBox("Se han escrito en C:\Temp " & dwg.GetLength(0) & " bites")
        Else
            '    'MsgBox("No hay DWG asociado...")
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
            Dim Referencia As String = quefila("REFERENCIA").ToString
            Dim Tipo As String = quefila("TIPO").ToString
            Dim Tipo1 As String = quefila("TIPO1").ToString
            Dim Tipo2 As String = quefila("TIPO2").ToString
            Dim Tipo3 As String = quefila("TIPO3").ToString
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
        If txtBusca.Text = "" Then Exit Sub
        Dim indice As Integer = lbETIQUETAS.FindString(txtBusca.Text.ToUpper, -1)
        If indice > -1 Then
            lbETIQUETAS.SetSelected(indice, True)
            lbETIQUETAS.SelectedIndex = indice
        End If
    End Sub
End Class