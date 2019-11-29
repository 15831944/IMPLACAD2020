Imports System.Text
Imports System.Linq
Partial Public Class Utiles
    ''
    ''' <summary>
    ''' Returns only the characters from the beginning of the string. Until he finds a number.
    ''' </summary>
    ''' <param name="txtString"></param>
    ''' <returns>Only the chars</returns>
    Public Shared Function String_GetStart(ByVal txtString As String) As String
        Dim resultado As String = ""
        For x As Integer = 0 To txtString.Length - 1
            If IsNumeric(txtString.Chars(x)) = False Then
                resultado &= txtString.Chars(x).ToString
            Else
                Exit For
            End If
        Next
        Return resultado
    End Function
    ''
    ''' <summary>
    ''' Returns only the numbers from the beginning of the string. Until he finds a char.
    ''' </summary>
    ''' <param name="txtString"></param>
    ''' <returns>Only the numbers (As string)</returns>
    Public Shared Function Number_GetStart(ByVal txtString As String) As String
        Dim resultado As String = ""
        For x As Integer = 0 To txtString.Length - 1
            If IsNumeric(txtString.Chars(x)) = True Then
                resultado &= txtString.Chars(x).ToString
            Else
                Exit For
            End If
        Next
        Return resultado
    End Function

End Class
