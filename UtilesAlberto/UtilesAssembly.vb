Imports System.Reflection

Partial Public Class Utiles
    Public Shared Function _appFullUtiles(asm As Assembly) As String
        Return asm.Location
    End Function
End Class
