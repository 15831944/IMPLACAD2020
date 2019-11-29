Imports System.Exception
<Serializable>
Public Class ActivationException
    Inherits Exception

    Public Sub New()
    End Sub

    Public Sub New(message As String)
        MyBase.New(message)
    End Sub

    Public Sub New(message As String, inner As Exception)
        MyBase.New(message, inner)
    End Sub

    Protected Sub New(serializationInfo As Runtime.Serialization.SerializationInfo, streamingContext As Runtime.Serialization.StreamingContext)
        Throw New NotImplementedException()
    End Sub
End Class
