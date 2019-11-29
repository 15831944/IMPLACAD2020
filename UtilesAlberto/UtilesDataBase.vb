Imports System.Data

Partial Public Class Utiles

    Public Shared Function DataRow_Cambia(ByRef queDR As DataRow) As DataRow
        Dim resultado As DataRow = queDR.Table.NewRow
        Dim columnas As DataColumnCollection = queDR.Table.Columns

        For Each col As DataColumn In columnas
            If IsDBNull(queDR.Item(col)) Then
                'Debug.Print(col.ColumnName & "( " & col.DataType.Name & " )")
                Select Case col.DataType.Name
                    Case "String"
                        resultado.Item(col) = ""
                    Case "Decimal"
                        resultado.Item(col) = 0
                    Case "Boolean"
                        resultado.Item(col) = False
                    Case Else
                        Try
                            resultado.Item(col) = ""
                        Catch ex As Exception
                            resultado.Item(col) = 0
                        End Try
                End Select
                'Debug.Print(vbTab & "Valor : ( " & resultado.Item(col).ToString & " )")
            Else
                If col.DataType.Name = "String" Then
                    Dim valor As String = Trim(queDR.Item(col).ToString)
                    If IsNumeric(valor) Then valor = valor.Replace(".", ",")
                    resultado.Item(col) = valor
                Else
                    resultado.Item(col) = queDR.Item(col)
                End If
            End If
        Next
        DataRow_Cambia = resultado
        Exit Function
    End Function
End Class
