Partial Public Class Utiles
    '' Le pasamos una cadena y nos devuelve sólo el número, sin las unidades
    '' Ej.: Le pasamos 1000 cm y nos devuelve 1000
    ''' <summary>
    ''' Returns only the numbers (Remove the units from the end, if it has)
    ''' </summary>
    ''' <param name="oString"></param>
    ''' <returns>Only numbers (Without units, as String)</returns>
    Public Shared Function String_WhithoutUnits(oString As String) As String
        Dim resultado As String = ""
        '
        If oString = "" Then
            resultado = oString.Trim
        ElseIf IsNumeric(oString) = True Then
            resultado = oString.Trim
        ElseIf IsNumeric(oString.Chars(0)) = False Then
            resultado = oString.Trim
        ElseIf oString.Contains(" ") = True Then
            Dim valores() As String = oString.Split(" ")
            If IsNumeric(valores(0)) = True Then
                resultado = valores(0).ToString.Trim
            Else
                resultado = oString.Trim
            End If
        End If
        '
        ' No hay números al inicio de la cadena. Los buscamos
        'For x As Integer = 0 To oString.Length - 1
        '    If IsNumeric(oString(x)) Then
        '        Continue For
        '    ElseIf oString(x) = "" Then
        '        Continue For
        '    Else
        '        resultado = oString.Substring(0, x - 1)
        '        Exit For
        '    End If
        'Next
        Return resultado
    End Function
End Class
