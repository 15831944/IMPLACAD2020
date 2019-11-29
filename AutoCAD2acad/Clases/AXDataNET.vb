Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices

Namespace A2acad
    Partial Public Class A2acad

        Public Function XLeeDatoNET(ByVal pId As ObjectId, ByVal queNombre As String, ByVal BlkRefIsErased As Boolean) As String
            Dim doc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim db As Database = doc.Database
            Dim strResult As String = ""
            Try
                Using acTrans As Transaction = db.TransactionManager.StartOpenCloseTransaction()

                    Dim ent As Entity = acTrans.GetObject(pId, OpenMode.ForRead, BlkRefIsErased)
                    Dim XdataOut() As TypedValue = ent.XData.AsArray()
                    If (Not XdataOut Is Nothing) Then

                        Dim todo As String = XdataOut(1).Value.ToString()  '' Cadena de texto con todas los datos.
                        If queNombre = "" Then
                            strResult = todo
                        ElseIf queNombre = regAPPA Then
                            strResult = XdataOut(0).Value.ToString()
                        ElseIf todo.Contains(queNombre) Then
                            Dim valoresdatos() As String = todo.Split(";"c)      '' cada elemento nombre=valor
                            For x As Integer = 0 To UBound(valoresdatos)
                                Dim nombre As String = valoresdatos(x).Split("="c)(0)
                                Dim valor As String = valoresdatos(x).Split("="c)(1)
                                If nombre.Equals(queNombre) Then
                                    strResult = valor
                                    Exit For
                                End If
                            Next
                        Else
                            strResult = ""
                        End If
                    End If
                    acTrans.Commit()
                End Using

                VaciaMemoria()
                Return strResult
            Catch ex As System.Exception
                Return ""
            End Try
        End Function

        Public Sub XPonDatoNET(ByVal pId As ObjectId, queNombre As String, queValor As String, ByVal BlkRefIsErased As Boolean)

            Dim doc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim db As Database = doc.Database

            Using acTrans As Transaction = db.TransactionManager.StartOpenCloseTransaction()
                Dim ent As Entity = acTrans.GetObject(pId, OpenMode.ForRead, BlkRefIsErased)
                Dim XDataOut() As TypedValue = ent.XData.AsArray()

                If XDataOut Is Nothing Then
                    'No existe el XDATA
                    ent.UpgradeOpen()
                    Using rb As New ResultBuffer
                        rb.Add(New TypedValue(DxfCode.ExtendedDataRegAppName, regAPPA))
                        rb.Add(New TypedValue(DxfCode.ExtendedDataAsciiString, queNombre & "=" & queValor))

                        ent.XData = rb

                    End Using
                Else
                    Dim todo As String = XDataOut(1).Value.ToString()
                    Dim arrayTodo() As String = todo.Split(";")
                    For i As Integer = 0 To arrayTodo.Length - 1
                        'Analiza cada uno de los datos
                        If arrayTodo(i).Contains(queNombre & "=") Then
                            arrayTodo(i) = queNombre & "=" & queValor
                            Exit For
                        End If
                    Next
                    todo = String.Join(";", arrayTodo)

                    ent.UpgradeOpen()
                    Using rb As New ResultBuffer
                        rb.Add(New TypedValue(DxfCode.ExtendedDataRegAppName, regAPPA))
                        rb.Add(New TypedValue(DxfCode.ExtendedDataAsciiString, todo))

                        ent.XData = rb

                    End Using
                    Dim strDato As String = XLeeDatoNET(pId, "ELEMENTO", BlkRefIsErased)
                End If


                acTrans.Commit()
            End Using
        End Sub
    End Class
End Namespace