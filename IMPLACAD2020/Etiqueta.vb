Public Class Etiqueta
    'Public IMPLACAD_DATA As String = "C:\ProgramData\IMPLACAD\"
    'REFERENCIA	TIPO	TIPO1	TIPO2	TIPO3	LARGO	ANCHO	STOCK	DESCRIPCION	PNG	DWG
    Public Property REFERENCIA As String = ""
    Public Property TIPO As String = ""
    Public Property TIPO1 As String = ""
    Public Property TIPO2 As String = ""
    Public Property TIPO3 As String = ""
    Public Property LARGO As String = ""
    Public Property ANCHO As String = ""
    Public Property STOCK As String = ""
    Public Property DESCRIPCION As String = ""
    Private Property _PNG As String = ""
    Private Property _DWG As String = ""
    Public Property PNG As String
        Set(value As String)
            If value.StartsWith(IMPLACAD_DATA) Then
                _PNG = value
            Else
                _PNG = IO.Path.Combine(IMPLACAD_DATA, value)
            End If
        End Set
        Get
            If _PNG.StartsWith(IMPLACAD_DATA) = False Then
                Return IO.Path.Combine(IMPLACAD_DATA, _PNG)
            Else
                Return _PNG
            End If
        End Get
    End Property
    Public Property DWG As String
        Set(value As String)
            If value.StartsWith(IMPLACAD_DATA) Then
                _DWG = value
            Else
                _DWG = IO.Path.Combine(IMPLACAD_DATA, value)
            End If
        End Set
        Get
            If _DWG.StartsWith(IMPLACAD_DATA) = False Then
                Return IO.Path.Combine(IMPLACAD_DATA, _DWG)
            Else
                Return _DWG
            End If
        End Get
    End Property
    Public Sub New(ref As String, t As String, t1 As String, t2 As String, t3 As String, l As String, a As String, s As String, d As String, pn As String, dw As String)
        'REFERENCIA	TIPO	TIPO1	TIPO2	TIPO3	LARGO	ANCHO	STOCK	DESCRIPCION	PNG	DWG
        Me.REFERENCIA = ref
        Me.TIPO = t
        Me.TIPO1 = t1
        Me.TIPO2 = t2
        Me.TIPO3 = t3
        Me.LARGO = l
        Me.ANCHO = a
        Me.STOCK = s
        Me.DESCRIPCION = d
        Me.PNG = pn
        Me.DWG = dw
    End Sub
End Class
