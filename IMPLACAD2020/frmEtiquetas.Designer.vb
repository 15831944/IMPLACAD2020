<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmEtiquetas
    Inherits System.Windows.Forms.Form

    'Form reemplaza a Dispose para limpiar la lista de componentes.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Requerido por el Diseñador de Windows Forms
    Private components As System.ComponentModel.IContainer

    'NOTA: el Diseñador de Windows Forms necesita el siguiente procedimiento
    'Se puede modificar usando el Diseñador de Windows Forms.  
    'No lo modifique con el editor de código.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.btnCerrar = New System.Windows.Forms.Button()
        Me.btnInsertar = New System.Windows.Forms.Button()
        Me.txtREFERENCIA = New System.Windows.Forms.TextBox()
        Me.txtDESCRIPCION = New System.Windows.Forms.TextBox()
        Me.lbETIQUETAS = New System.Windows.Forms.ListBox()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.cbTIPO3 = New System.Windows.Forms.ComboBox()
        Me.cbTIPO2 = New System.Windows.Forms.ComboBox()
        Me.cbTIPO1 = New System.Windows.Forms.ComboBox()
        Me.cbTIPO = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txtTIPO = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.txtTIPO1 = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.txtTIPO2 = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.txtTIPO3 = New System.Windows.Forms.TextBox()
        Me.cbStock = New System.Windows.Forms.CheckBox()
        Me.lbDim = New System.Windows.Forms.Label()
        Me.lbCuantos = New System.Windows.Forms.Label()
        Me.btnSel = New System.Windows.Forms.Button()
        Me.btnEditar = New System.Windows.Forms.Button()
        Me.btnCambiar = New System.Windows.Forms.Button()
        Me.lblCambiar = New System.Windows.Forms.Label()
        Me.txtBusca = New System.Windows.Forms.TextBox()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.BtnBuscar = New System.Windows.Forms.Button()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.pbIMAGEN = New System.Windows.Forms.PictureBox()
        Me.GroupBox1.SuspendLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.pbIMAGEN, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btnCerrar
        '
        Me.btnCerrar.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnCerrar.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnCerrar.Location = New System.Drawing.Point(728, 556)
        Me.btnCerrar.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btnCerrar.Name = "btnCerrar"
        Me.btnCerrar.Size = New System.Drawing.Size(100, 28)
        Me.btnCerrar.TabIndex = 0
        Me.btnCerrar.Text = "Cerrar"
        Me.btnCerrar.UseVisualStyleBackColor = True
        '
        'btnInsertar
        '
        Me.btnInsertar.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnInsertar.Enabled = False
        Me.btnInsertar.Location = New System.Drawing.Point(592, 556)
        Me.btnInsertar.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btnInsertar.Name = "btnInsertar"
        Me.btnInsertar.Size = New System.Drawing.Size(100, 28)
        Me.btnInsertar.TabIndex = 1
        Me.btnInsertar.Text = "Insertar"
        Me.btnInsertar.UseVisualStyleBackColor = True
        '
        'txtREFERENCIA
        '
        Me.txtREFERENCIA.Location = New System.Drawing.Point(264, 37)
        Me.txtREFERENCIA.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.txtREFERENCIA.MaxLength = 15
        Me.txtREFERENCIA.Name = "txtREFERENCIA"
        Me.txtREFERENCIA.Size = New System.Drawing.Size(233, 22)
        Me.txtREFERENCIA.TabIndex = 2
        '
        'txtDESCRIPCION
        '
        Me.txtDESCRIPCION.Location = New System.Drawing.Point(264, 394)
        Me.txtDESCRIPCION.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.txtDESCRIPCION.MaxLength = 500
        Me.txtDESCRIPCION.Multiline = True
        Me.txtDESCRIPCION.Name = "txtDESCRIPCION"
        Me.txtDESCRIPCION.Size = New System.Drawing.Size(575, 64)
        Me.txtDESCRIPCION.TabIndex = 3
        '
        'lbETIQUETAS
        '
        Me.lbETIQUETAS.FormattingEnabled = True
        Me.lbETIQUETAS.ItemHeight = 16
        Me.lbETIQUETAS.Location = New System.Drawing.Point(16, 239)
        Me.lbETIQUETAS.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.lbETIQUETAS.Name = "lbETIQUETAS"
        Me.lbETIQUETAS.Size = New System.Drawing.Size(219, 292)
        Me.lbETIQUETAS.TabIndex = 5
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.Label10)
        Me.GroupBox1.Controls.Add(Me.Label9)
        Me.GroupBox1.Controls.Add(Me.Label8)
        Me.GroupBox1.Controls.Add(Me.Label7)
        Me.GroupBox1.Controls.Add(Me.cbTIPO3)
        Me.GroupBox1.Controls.Add(Me.cbTIPO2)
        Me.GroupBox1.Controls.Add(Me.cbTIPO1)
        Me.GroupBox1.Controls.Add(Me.cbTIPO)
        Me.GroupBox1.Location = New System.Drawing.Point(16, 12)
        Me.GroupBox1.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Padding = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.GroupBox1.Size = New System.Drawing.Size(220, 175)
        Me.GroupBox1.TabIndex = 9
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Filtros - TIPO"
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(8, 146)
        Me.Label10.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(44, 17)
        Me.Label10.TabIndex = 16
        Me.Label10.Text = "Tipo3"
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(8, 107)
        Me.Label9.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(44, 17)
        Me.Label9.TabIndex = 15
        Me.Label9.Text = "Tipo2"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(8, 69)
        Me.Label8.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(44, 17)
        Me.Label8.TabIndex = 14
        Me.Label8.Text = "Tipo1"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(8, 27)
        Me.Label7.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(36, 17)
        Me.Label7.TabIndex = 13
        Me.Label7.Text = "Tipo"
        '
        'cbTIPO3
        '
        Me.cbTIPO3.FormattingEnabled = True
        Me.cbTIPO3.Location = New System.Drawing.Point(65, 143)
        Me.cbTIPO3.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.cbTIPO3.Name = "cbTIPO3"
        Me.cbTIPO3.Size = New System.Drawing.Size(124, 24)
        Me.cbTIPO3.TabIndex = 12
        '
        'cbTIPO2
        '
        Me.cbTIPO2.FormattingEnabled = True
        Me.cbTIPO2.Location = New System.Drawing.Point(65, 103)
        Me.cbTIPO2.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.cbTIPO2.Name = "cbTIPO2"
        Me.cbTIPO2.Size = New System.Drawing.Size(124, 24)
        Me.cbTIPO2.TabIndex = 11
        '
        'cbTIPO1
        '
        Me.cbTIPO1.FormattingEnabled = True
        Me.cbTIPO1.Location = New System.Drawing.Point(65, 63)
        Me.cbTIPO1.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.cbTIPO1.Name = "cbTIPO1"
        Me.cbTIPO1.Size = New System.Drawing.Size(124, 24)
        Me.cbTIPO1.TabIndex = 10
        '
        'cbTIPO
        '
        Me.cbTIPO.FormattingEnabled = True
        Me.cbTIPO.Location = New System.Drawing.Point(65, 23)
        Me.cbTIPO.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.cbTIPO.Name = "cbTIPO"
        Me.cbTIPO.Size = New System.Drawing.Size(124, 24)
        Me.cbTIPO.TabIndex = 9
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(260, 15)
        Me.Label1.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(77, 17)
        Me.Label1.TabIndex = 10
        Me.Label1.Text = "Referencia"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(260, 81)
        Me.Label2.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(36, 17)
        Me.Label2.TabIndex = 12
        Me.Label2.Text = "Tipo"
        '
        'txtTIPO
        '
        Me.txtTIPO.Location = New System.Drawing.Point(264, 103)
        Me.txtTIPO.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.txtTIPO.MaxLength = 15
        Me.txtTIPO.Name = "txtTIPO"
        Me.txtTIPO.Size = New System.Drawing.Size(233, 22)
        Me.txtTIPO.TabIndex = 11
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(260, 140)
        Me.Label3.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(44, 17)
        Me.Label3.TabIndex = 14
        Me.Label3.Text = "Tipo1"
        '
        'txtTIPO1
        '
        Me.txtTIPO1.Location = New System.Drawing.Point(264, 162)
        Me.txtTIPO1.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.txtTIPO1.MaxLength = 15
        Me.txtTIPO1.Name = "txtTIPO1"
        Me.txtTIPO1.Size = New System.Drawing.Size(233, 22)
        Me.txtTIPO1.TabIndex = 13
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(260, 199)
        Me.Label4.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(44, 17)
        Me.Label4.TabIndex = 16
        Me.Label4.Text = "Tipo2"
        '
        'txtTIPO2
        '
        Me.txtTIPO2.Location = New System.Drawing.Point(264, 222)
        Me.txtTIPO2.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.txtTIPO2.MaxLength = 15
        Me.txtTIPO2.Name = "txtTIPO2"
        Me.txtTIPO2.Size = New System.Drawing.Size(233, 22)
        Me.txtTIPO2.TabIndex = 15
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(260, 370)
        Me.Label5.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(82, 17)
        Me.Label5.TabIndex = 17
        Me.Label5.Text = "Descripción"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(260, 257)
        Me.Label6.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(44, 17)
        Me.Label6.TabIndex = 19
        Me.Label6.Text = "Tipo3"
        '
        'txtTIPO3
        '
        Me.txtTIPO3.Location = New System.Drawing.Point(264, 279)
        Me.txtTIPO3.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.txtTIPO3.MaxLength = 15
        Me.txtTIPO3.Name = "txtTIPO3"
        Me.txtTIPO3.Size = New System.Drawing.Size(233, 22)
        Me.txtTIPO3.TabIndex = 18
        '
        'cbStock
        '
        Me.cbStock.AutoSize = True
        Me.cbStock.Enabled = False
        Me.cbStock.Location = New System.Drawing.Point(264, 324)
        Me.cbStock.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.cbStock.Name = "cbStock"
        Me.cbStock.Size = New System.Drawing.Size(86, 21)
        Me.cbStock.TabIndex = 20
        Me.cbStock.Text = "En Stock"
        Me.cbStock.UseVisualStyleBackColor = True
        '
        'lbDim
        '
        Me.lbDim.Location = New System.Drawing.Point(507, 348)
        Me.lbDim.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lbDim.Name = "lbDim"
        Me.lbDim.Size = New System.Drawing.Size(333, 21)
        Me.lbDim.TabIndex = 21
        Me.lbDim.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lbCuantos
        '
        Me.lbCuantos.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.lbCuantos.AutoSize = True
        Me.lbCuantos.Location = New System.Drawing.Point(24, 567)
        Me.lbCuantos.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lbCuantos.Name = "lbCuantos"
        Me.lbCuantos.Size = New System.Drawing.Size(36, 17)
        Me.lbCuantos.TabIndex = 22
        Me.lbCuantos.Text = "       "
        '
        'btnSel
        '
        Me.btnSel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnSel.BackColor = System.Drawing.SystemColors.ActiveCaption
        Me.btnSel.Location = New System.Drawing.Point(444, 556)
        Me.btnSel.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btnSel.Name = "btnSel"
        Me.btnSel.Size = New System.Drawing.Size(100, 28)
        Me.btnSel.TabIndex = 23
        Me.btnSel.Text = "Selecciona"
        Me.btnSel.UseVisualStyleBackColor = False
        Me.btnSel.Visible = False
        '
        'btnEditar
        '
        Me.btnEditar.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnEditar.Location = New System.Drawing.Point(264, 542)
        Me.btnEditar.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btnEditar.Name = "btnEditar"
        Me.btnEditar.Size = New System.Drawing.Size(137, 46)
        Me.btnEditar.TabIndex = 24
        Me.btnEditar.Text = "Seleccionar Bloque a Cambiar"
        Me.btnEditar.UseVisualStyleBackColor = True
        '
        'btnCambiar
        '
        Me.btnCambiar.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnCambiar.Location = New System.Drawing.Point(579, 555)
        Me.btnCambiar.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btnCambiar.Name = "btnCambiar"
        Me.btnCambiar.Size = New System.Drawing.Size(100, 28)
        Me.btnCambiar.TabIndex = 25
        Me.btnCambiar.Text = "Cambiar"
        Me.btnCambiar.UseVisualStyleBackColor = True
        Me.btnCambiar.Visible = False
        '
        'lblCambiar
        '
        Me.lblCambiar.AutoSize = True
        Me.lblCambiar.Location = New System.Drawing.Point(269, 496)
        Me.lblCambiar.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblCambiar.Name = "lblCambiar"
        Me.lblCambiar.Size = New System.Drawing.Size(59, 17)
        Me.lblCambiar.TabIndex = 26
        Me.lblCambiar.Text = "Label11"
        '
        'txtBusca
        '
        Me.txtBusca.Location = New System.Drawing.Point(81, 202)
        Me.txtBusca.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.txtBusca.Name = "txtBusca"
        Me.txtBusca.Size = New System.Drawing.Size(111, 22)
        Me.txtBusca.TabIndex = 27
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(20, 206)
        Me.Label11.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(56, 17)
        Me.Label11.TabIndex = 28
        Me.Label11.Text = "Buscar:"
        '
        'BtnBuscar
        '
        Me.BtnBuscar.BackgroundImage = Global.IMPLACAD.My.Resources.Resources.Search_256x256
        Me.BtnBuscar.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.BtnBuscar.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtnBuscar.Location = New System.Drawing.Point(199, 197)
        Me.BtnBuscar.Name = "BtnBuscar"
        Me.BtnBuscar.Size = New System.Drawing.Size(32, 32)
        Me.BtnBuscar.TabIndex = 30
        Me.BtnBuscar.UseVisualStyleBackColor = True
        '
        'PictureBox1
        '
        Me.PictureBox1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.PictureBox1.Image = Global.IMPLACAD.My.Resources.Resources.LOGO_IMPLASER
        Me.PictureBox1.Location = New System.Drawing.Point(16, 542)
        Me.PictureBox1.Margin = New System.Windows.Forms.Padding(4)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(220, 53)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.PictureBox1.TabIndex = 29
        Me.PictureBox1.TabStop = False
        '
        'pbIMAGEN
        '
        Me.pbIMAGEN.Location = New System.Drawing.Point(507, 15)
        Me.pbIMAGEN.Margin = New System.Windows.Forms.Padding(4)
        Me.pbIMAGEN.Name = "pbIMAGEN"
        Me.pbIMAGEN.Size = New System.Drawing.Size(333, 330)
        Me.pbIMAGEN.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.pbIMAGEN.TabIndex = 4
        Me.pbIMAGEN.TabStop = False
        '
        'frmEtiquetas
        '
        Me.AcceptButton = Me.btnInsertar
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnCerrar
        Me.ClientSize = New System.Drawing.Size(856, 599)
        Me.Controls.Add(Me.BtnBuscar)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.Label11)
        Me.Controls.Add(Me.txtBusca)
        Me.Controls.Add(Me.lblCambiar)
        Me.Controls.Add(Me.btnCambiar)
        Me.Controls.Add(Me.btnEditar)
        Me.Controls.Add(Me.btnSel)
        Me.Controls.Add(Me.lbCuantos)
        Me.Controls.Add(Me.lbDim)
        Me.Controls.Add(Me.cbStock)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.txtTIPO3)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.txtTIPO2)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.txtTIPO1)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.txtTIPO)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.lbETIQUETAS)
        Me.Controls.Add(Me.pbIMAGEN)
        Me.Controls.Add(Me.txtDESCRIPCION)
        Me.Controls.Add(Me.txtREFERENCIA)
        Me.Controls.Add(Me.btnInsertar)
        Me.Controls.Add(Me.btnCerrar)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmEtiquetas"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.Text = "Insertar/Editar Etiquetas"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.pbIMAGEN, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnCerrar As System.Windows.Forms.Button
    Friend WithEvents btnInsertar As System.Windows.Forms.Button
    Friend WithEvents txtREFERENCIA As System.Windows.Forms.TextBox
    Friend WithEvents txtDESCRIPCION As System.Windows.Forms.TextBox
    Friend WithEvents pbIMAGEN As System.Windows.Forms.PictureBox
    Friend WithEvents lbETIQUETAS As System.Windows.Forms.ListBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents cbTIPO2 As System.Windows.Forms.ComboBox
    Friend WithEvents cbTIPO1 As System.Windows.Forms.ComboBox
    Friend WithEvents cbTIPO As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtTIPO As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtTIPO1 As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txtTIPO2 As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents txtTIPO3 As System.Windows.Forms.TextBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents cbTIPO3 As System.Windows.Forms.ComboBox
    Friend WithEvents cbStock As System.Windows.Forms.CheckBox
    Friend WithEvents lbDim As System.Windows.Forms.Label
    Friend WithEvents lbCuantos As System.Windows.Forms.Label
    Friend WithEvents btnSel As System.Windows.Forms.Button
    Friend WithEvents btnEditar As System.Windows.Forms.Button
    Friend WithEvents btnCambiar As System.Windows.Forms.Button
    Friend WithEvents lblCambiar As System.Windows.Forms.Label
    Friend WithEvents txtBusca As System.Windows.Forms.TextBox
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents BtnBuscar As Windows.Forms.Button
End Class
