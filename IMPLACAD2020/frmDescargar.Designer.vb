<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmDescargar
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
        Me.pb1 = New System.Windows.Forms.ProgressBar()
        Me.lblInfo = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'pb1
        '
        Me.pb1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.pb1.Location = New System.Drawing.Point(12, 31)
        Me.pb1.Name = "pb1"
        Me.pb1.Size = New System.Drawing.Size(460, 20)
        Me.pb1.Style = System.Windows.Forms.ProgressBarStyle.Continuous
        Me.pb1.TabIndex = 0
        '
        'lblInfo
        '
        Me.lblInfo.Location = New System.Drawing.Point(12, 9)
        Me.lblInfo.Name = "lblInfo"
        Me.lblInfo.Size = New System.Drawing.Size(460, 14)
        Me.lblInfo.TabIndex = 1
        Me.lblInfo.Text = "Descargando xxx de xxx"
        Me.lblInfo.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'frmDescargar
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(484, 61)
        Me.ControlBox = False
        Me.Controls.Add(Me.lblInfo)
        Me.Controls.Add(Me.pb1)
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmDescargar"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Actualizando..."
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents pb1 As System.Windows.Forms.ProgressBar
    Friend WithEvents lblInfo As System.Windows.Forms.Label
End Class
