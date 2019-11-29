Imports System.Collections
Imports System.Collections.Generic
Imports UtilesAlberto

Partial Public Class Utiles
    Public Const sepItems As String = "·"
    Public Const sepValues As String = "|"

    Public Shared Function Array_ChangeToArrayString(arrObjects As Object()) As String()
        Dim resultado() As String = Array.ConvertAll(arrObjects,
                                               New Converter(Of Object, String)(Function(x)
                                                                                    Return CType(x, String)
                                                                                End Function))
        Return resultado
    End Function
    Public Shared Function List_ChangeToArrayString(ListObjects As List(Of Object)) As String()
        Dim resultado() As String = Array.ConvertAll(ListObjects.ToArray,
                                               New Converter(Of Object, String)(Function(x)
                                                                                    Return CType(x, String)
                                                                                End Function))
        Return resultado
    End Function

    ''' <summary>
    ''' Add Item to Array
    ''' </summary>
    ''' <param name="oArray"></param>
    ''' <param name="oValue"></param>
    ''' <remarks></remarks>
    Public Sub Array_AddItem(ByRef oArray() As Object, ByVal oValue As Object)
        ReDim Preserve oArray(oArray.GetUpperBound(0) + 1)
        oArray(oArray.GetUpperBound(0)) = oValue
    End Sub

    Public Shared Function Dictionary_ToString(Dic As Dictionary(Of String, String), Optional sItems As String = sepItems, Optional sValues As String = sepValues) As String
        Dim resultado As String = ""
        Dim union = From k In Dic.Keys
                    Select k & sItems & Dic(k)

        resultado = String.Join(sValues, union.ToArray)
        Return resultado
    End Function

    Public Shared Function String_ToDictionary(Str As String, Optional sItems As String = sepItems, Optional sValues As String = sepValues) As Dictionary(Of String, String)
        Dim resultado As New Dictionary(Of String, String)
        Dim arrValues As String() = Str.Split(sValues)
        For Each Value In arrValues
            Dim arrItems() As String = Value.Split(sItems)
            If arrItems.Count = 2 Then
                resultado.Add(arrItems(0), arrItems(1))
            ElseIf arrItems.Count = 1 Then
                resultado.Add(arrItems(0), "")
            End If
        Next
        '
        Return resultado
    End Function

    Public Shared Function Array_ToString(Arr As String(), Optional sItems As String = sepItems, Optional sValues As String = sepValues) As String
        Dim resultado As String = ""
        resultado = String.Join(sValues, Arr)
        Return resultado
    End Function

    Public Shared Function String_ToArray(Str As String, Optional sItems As String = sepItems, Optional sValues As String = sepValues) As String()
        Dim resultado As New List(Of String)
        Dim arrValues As String() = Str.Split(sValues)
        For Each Value In arrValues
            Dim arrItems() As String = Value.Split(sItems)
            resultado.Add(String.Join(sItems, arrItems))
        Next
        '
        Return resultado.ToArray
    End Function
    Public Shared Function List_ToString(lPar As List(Of Parameter), Optional sItems As String = sepItems, Optional sValues As String = sepValues) As String
        Dim resultado As String = ""
        For Each oP As Parameter In lPar
            Dim _P As String = ""
            Dim _V As String = ""
            Dim _U As String = ""
            Dim _D As String = ""
            On Error Resume Next
            _P = oP._Parameter
            _V = oP._Value.ToString
            _U = oP._Units
            _D = oP._Description

            resultado &=
                sValues &
                _P & sItems &
                _V & sItems &
                _U & sItems &
                _D
        Next
        '
        resultado = IIf(resultado <> "", resultado.Substring(1), "")
        Return resultado
    End Function
    Public Shared Function String_ToList(Str As String, Optional sItems As String = sepItems, Optional sValues As String = sepValues) As List(Of Parameter)
        Dim resultado As New List(Of Parameter)
        Dim arrValues As String() = Str.Split(sValues)
        For Each Item In arrValues
            Dim arrItems() As String = Item.Split(sItems)
            '
            Dim oPar As New Parameter
            oPar._Parameter = arrItems(0)
            If arrItems.Length > 1 Then oPar._Value = arrItems(1)
            If arrItems.Length > 2 Then oPar._Units = arrItems(2)
            If arrItems.Length > 3 Then oPar._Description = arrItems(3)
            resultado.Add(oPar)
            oPar = Nothing
        Next
        '
        Return resultado
    End Function

    Public Shared Function Dictionary_Imprime(Dic As Dictionary(Of String, String), Optional sItems As String = sepItems) As String
        Dim resultado As String = ""
        Array.ForEach(Dic.Keys.ToArray, Sub(a As String)
                                            resultado &= a & sepItems & Dic(a) & vbCrLf
                                        End Sub)
        Return resultado
    End Function
    Public Shared Function Array_Imprime(Arr As String(), Optional sItems As String = sepItems) As String
        Dim resultado As String = ""
        Array.ForEach(Arr, Sub(a As String)
                               resultado &= String.Join(sepItems, a) & vbCrLf
                           End Sub)
        Return resultado
    End Function
    Public Shared Function List_Imprime(lPar As List(Of Parameter), Optional sItems As String = sepItems) As String
        Dim resultado As String = ""
        lPar.ForEach(Sub(p As Parameter)
                         resultado &= p._Parameter & sItems & p._Value.ToString & sItems & p._Units & sItems & p._Description & vbCrLf
                     End Sub)
        Return resultado
    End Function
End Class