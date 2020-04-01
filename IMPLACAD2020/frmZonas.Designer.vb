<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmZonas
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmZonas))
        Me.TvZonas = New System.Windows.Forms.TreeView()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.LblContar = New System.Windows.Forms.Label()
        Me.BtnCrear = New System.Windows.Forms.Button()
        Me.BtnEditar = New System.Windows.Forms.Button()
        Me.BtnBorrar = New System.Windows.Forms.Button()
        Me.BtnTablaSeleccion = New System.Windows.Forms.Button()
        Me.BtnTablaTodas = New System.Windows.Forms.Button()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.CheckBox1 = New System.Windows.Forms.CheckBox()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'TvZonas
        '
        Me.TvZonas.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.TvZonas.Location = New System.Drawing.Point(12, 32)
        Me.TvZonas.Name = "TvZonas"
        Me.TvZonas.Size = New System.Drawing.Size(173, 379)
        Me.TvZonas.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(16, 8)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(56, 17)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Zonas :"
        '
        'LblContar
        '
        Me.LblContar.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.LblContar.AutoSize = True
        Me.LblContar.Location = New System.Drawing.Point(14, 424)
        Me.LblContar.Name = "LblContar"
        Me.LblContar.Size = New System.Drawing.Size(176, 17)
        Me.LblContar.TabIndex = 2
        Me.LblContar.Text = "XX etiquetas en Zona XXX"
        '
        'BtnCrear
        '
        Me.BtnCrear.Location = New System.Drawing.Point(211, 32)
        Me.BtnCrear.Name = "BtnCrear"
        Me.BtnCrear.Size = New System.Drawing.Size(114, 32)
        Me.BtnCrear.TabIndex = 3
        Me.BtnCrear.Text = "Crear"
        Me.BtnCrear.UseVisualStyleBackColor = True
        '
        'BtnEditar
        '
        Me.BtnEditar.Location = New System.Drawing.Point(211, 84)
        Me.BtnEditar.Name = "BtnEditar"
        Me.BtnEditar.Size = New System.Drawing.Size(114, 32)
        Me.BtnEditar.TabIndex = 4
        Me.BtnEditar.Text = "Editar"
        Me.BtnEditar.UseVisualStyleBackColor = True
        '
        'BtnBorrar
        '
        Me.BtnBorrar.Location = New System.Drawing.Point(211, 136)
        Me.BtnBorrar.Name = "BtnBorrar"
        Me.BtnBorrar.Size = New System.Drawing.Size(114, 32)
        Me.BtnBorrar.TabIndex = 5
        Me.BtnBorrar.Text = "Borrar"
        Me.BtnBorrar.UseVisualStyleBackColor = True
        '
        'BtnTablaSeleccion
        '
        Me.BtnTablaSeleccion.Image = Global.IMPLACAD.My.Resources.Resources.TABLA_DATOS32
        Me.BtnTablaSeleccion.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.BtnTablaSeleccion.Location = New System.Drawing.Point(211, 191)
        Me.BtnTablaSeleccion.Name = "BtnTablaSeleccion"
        Me.BtnTablaSeleccion.Size = New System.Drawing.Size(114, 49)
        Me.BtnTablaSeleccion.TabIndex = 6
        Me.BtnTablaSeleccion.Text = "Tabla Selección"
        Me.BtnTablaSeleccion.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BtnTablaSeleccion.UseVisualStyleBackColor = True
        '
        'BtnTablaTodas
        '
        Me.BtnTablaTodas.Image = Global.IMPLACAD.My.Resources.Resources.TABLA_DATOS32
        Me.BtnTablaTodas.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.BtnTablaTodas.Location = New System.Drawing.Point(211, 267)
        Me.BtnTablaTodas.Name = "BtnTablaTodas"
        Me.BtnTablaTodas.Size = New System.Drawing.Size(114, 49)
        Me.BtnTablaTodas.TabIndex = 7
        Me.BtnTablaTodas.Text = "Tablas     TODAS"
        Me.BtnTablaTodas.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BtnTablaTodas.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.CheckBox1)
        Me.GroupBox1.Location = New System.Drawing.Point(210, 340)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(114, 70)
        Me.GroupBox1.TabIndex = 8
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Capas :"
        '
        'CheckBox1
        '
        Me.CheckBox1.AutoSize = True
        Me.CheckBox1.Location = New System.Drawing.Point(6, 21)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.Size = New System.Drawing.Size(70, 21)
        Me.CheckBox1.TabIndex = 0
        Me.CheckBox1.Text = "Zonas"
        Me.CheckBox1.UseVisualStyleBackColor = True
        '
        'frmZonas
        '
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None
        Me.ClientSize = New System.Drawing.Size(352, 453)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.BtnTablaTodas)
        Me.Controls.Add(Me.BtnTablaSeleccion)
        Me.Controls.Add(Me.BtnBorrar)
        Me.Controls.Add(Me.BtnEditar)
        Me.Controls.Add(Me.BtnCrear)
        Me.Controls.Add(Me.LblContar)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.TvZonas)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MaximumSize = New System.Drawing.Size(370, 500)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(370, 500)
        Me.Name = "frmZonas"
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "frmZonas"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents TvZonas As Windows.Forms.TreeView
    Friend WithEvents Label1 As Windows.Forms.Label
    Friend WithEvents LblContar As Windows.Forms.Label
    Friend WithEvents BtnCrear As Windows.Forms.Button
    Friend WithEvents BtnEditar As Windows.Forms.Button
    Friend WithEvents BtnBorrar As Windows.Forms.Button
    Friend WithEvents BtnTablaSeleccion As Windows.Forms.Button
    Friend WithEvents BtnTablaTodas As Windows.Forms.Button
    Friend WithEvents GroupBox1 As Windows.Forms.GroupBox
    Friend WithEvents CheckBox1 As Windows.Forms.CheckBox
End Class
