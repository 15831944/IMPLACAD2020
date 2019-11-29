Public Class Parameter
    Private Parameter As String = ""
    Private Value As Object = Nothing
    Private Units As String = ""
    Private Description As String = ""

    Public Property _Parameter As String
        Get
            Return Parameter
        End Get
        Set(value As String)
            Parameter = value
        End Set
    End Property

    Public Property _Value As Object
        Get
            Return Value
        End Get
        Set(value As Object)
            Me.Value = value
        End Set
    End Property

    Public Property _Units As String
        Get
            Return Units
        End Get
        Set(value As String)
            Units = value
        End Set
    End Property

    Public Property _Description As String
        Get
            Return Description
        End Get
        Set(value As String)
            Description = value
        End Set
    End Property
    '
    Public Overrides Function toString() As String
        Dim resultado As String =
            _Parameter & vbCrLf & _Value.ToString & vbCrLf & _Units & vbCrLf & _Description
        Return resultado
    End Function
    '
    ' *** Para ordenar un Dim lPar as List(Of Parameter)
    ' Descendiente
    'Dim comparador As New ParameterSortDescending
    'lPar.Sort(comparador)
    ' Ascendiente
    'Dim comparador As New ParameterSortAscending
    'lPar.Sort(comparador)
    ' *** Solo quedaría recorrer lPar y mostrar sus valores ordenados

End Class

Public Class ParameterSortDescending
    Implements IComparer(Of Parameter)

    Public Function Compare(p1 As Parameter, p2 As Parameter) As Integer Implements IComparer(Of Parameter).Compare

        Dim resultado As Integer = 0        ' Sin iguales
        If (p1._Parameter < p2._Parameter) Then
            resultado = 1
        ElseIf (p1._Parameter > p2._Parameter) Then
            resultado = -1
        ElseIf (p1._Parameter = p2._Parameter) Then
            resultado = 0
        End If
        Return resultado
    End Function
End Class
Public Class ParameterSortAscending
    Implements IComparer(Of Parameter)

    Public Function Compare(p1 As Parameter, p2 As Parameter) As Integer Implements IComparer(Of Parameter).Compare
        ', Optional Descending As Boolean = True
        'Throw New NotImplementedException()
        Dim resultado As Integer = 0        ' Sin iguales
        If (p1._Parameter > p2._Parameter) Then
            resultado = 1
        ElseIf (p1._Parameter < p2._Parameter) Then
            resultado = -1
        ElseIf (p1._Parameter = p2._Parameter) Then
            resultado = 0
        End If
        Return resultado
    End Function
End Class
