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
        Me.SuspendLayout()
        '
        'TvZonas
        '
        Me.TvZonas.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.TvZonas.Location = New System.Drawing.Point(12, 32)
        Me.TvZonas.Name = "TvZonas"
        Me.TvZonas.Size = New System.Drawing.Size(173, 376)
        Me.TvZonas.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.AutoSize = True
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
        Me.LblContar.Location = New System.Drawing.Point(14, 421)
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
        Me.BtnEditar.Location = New System.Drawing.Point(211, 83)
        Me.BtnEditar.Name = "BtnEditar"
        Me.BtnEditar.Size = New System.Drawing.Size(114, 32)
        Me.BtnEditar.TabIndex = 4
        Me.BtnEditar.Text = "Editar"
        Me.BtnEditar.UseVisualStyleBackColor = True
        '
        'BtnBorrar
        '
        Me.BtnBorrar.Location = New System.Drawing.Point(211, 168)
        Me.BtnBorrar.Name = "BtnBorrar"
        Me.BtnBorrar.Size = New System.Drawing.Size(114, 32)
        Me.BtnBorrar.TabIndex = 5
        Me.BtnBorrar.Text = "Borrar"
        Me.BtnBorrar.UseVisualStyleBackColor = True
        '
        'frmZonas
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 450)
        Me.Controls.Add(Me.BtnBorrar)
        Me.Controls.Add(Me.BtnEditar)
        Me.Controls.Add(Me.BtnCrear)
        Me.Controls.Add(Me.LblContar)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.TvZonas)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmZonas"
        Me.Text = "frmZonas"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents TvZonas As Windows.Forms.TreeView
    Friend WithEvents Label1 As Windows.Forms.Label
    Friend WithEvents LblContar As Windows.Forms.Label
    Friend WithEvents BtnCrear As Windows.Forms.Button
    Friend WithEvents BtnEditar As Windows.Forms.Button
    Friend WithEvents BtnBorrar As Windows.Forms.Button
End Class
