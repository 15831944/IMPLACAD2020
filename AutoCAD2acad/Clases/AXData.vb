Imports System.Diagnostics
Imports System.Collections
Imports System.IO
Imports System.Windows.Forms

Imports Autodesk.AutoCAD.Interop
Imports Autodesk.AutoCAD.Interop.Common
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.EditorInput

Namespace A2acad
    Partial Public Class A2acad
        ''***** XDatos que pondremos por defecto al crear los objetos.
        '' El dato(0) sera igual a la aplicación a registrar (appReg = 2acad, por defecto)
        '' El dato(1) siempre será texto. Aquí pondremos todas las variables (nombre=valor;nombre1=valor1;etc)
        '' Habrá que convertir a double los valores que sean numéricos
        ''   para realizar operaciones FormatNumber(cdbl(valor, valor1),2,tristate)
        Public xTObj As Short() = New Short() {1001, 1000}
        Public xDObj As Object() = New Object() {"", ""}        'nombre=valor;nombre1=valor1 {regAPPA, ""}
        Public Sub XNuevo(ByRef objA As AcadObject,
                  Optional queRegApp As String = "",
                  Optional queValor As String = "")
            'Using Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument()
            If queRegApp = "" Then
                    xDObj(0) = regAPPA
                Else
                    xDObj(0) = queRegApp
                End If
                '
                If queValor <> "" Then
                    xDObj(1) = queValor
                End If
                '' Pone los XData. Solo si un objeto no tiene XData.
                XPonTiposDatosArrays(objA, xTObj, xDObj)
            'End Using
        End Sub

        Public Sub XBorrar(ByVal objA As AcadObject)
            Dim DataType(0) As Short
            Dim Data(0) As Object
            DataType(0) = 1001 : Data(0) = regAPPA
            objA.SetXData(DataType, Data)
            objA.Update()
        End Sub

        Public Sub XLeeMensaje(ByVal obj As AcadObject)
            '' Leer xdata del objeto
            Dim xdatos As Object = Nothing
            Dim xtipos As Object = Nothing
            obj.GetXData(regAPPA, xtipos, xdatos)

            Dim mensaje As String = ""
            For x As Integer = 0 To UBound(xdatos)
                mensaje &= xdatos(x) & vbCrLf
            Next
            If mensaje = "" Then
                MessageBox.Show("El objeto no tiene XData...")
            Else
                MessageBox.Show(mensaje)
            End If
        End Sub

        Public Function XLeeTiposDatos(ByVal objA As AcadObject) As Object()         'Devuelve todos los datos de SERICAD 0-SERICAD, 1-CAPA
            Dim resul(1) As Object
            ' Leer xdata del objeto
            Dim xtipos() As Short = Nothing
            Dim xdatos() As Object = Nothing
            objA.GetXData(regAPPA, xtipos, xdatos)
            '' Si el objeto no tiene XData
            If xdatos Is Nothing Then
                XNuevo(objA)
                objA.GetXData(regAPPA, xtipos, xdatos)
            End If
            resul(0) = xtipos : resul(1) = xdatos
            Return resul
        End Function

        Public Function XLeeDatosHandle(ByVal oHandle As String) As Object()         'Devuelve todos los Xdata en un array
            Dim objA As AcadObject = Nothing
            ' Leer xdata del objeto
            Dim xtipos() As Short = Nothing
            Dim resultado As Object() = Nothing
            If oHandle = "" Then
                Return resultado
                Exit Function
            End If
            Try
                objA = oAppA.ActiveDocument.HandleToObject(oHandle)
            Catch ex As Exception
                Return resultado
                Exit Function
            End Try
            '
            Try
                objA.GetXData(regAPPA, xtipos, resultado)
                '' Si el objeto no tiene XData
                If resultado Is Nothing Then
                    XNuevo(objA)
                    objA.GetXData(regAPPA, xtipos, resultado)
                End If
            Catch ex As Exception

            End Try
            Return resultado
        End Function

        Public Function XLeeDato(ByVal oHandle As String, ByVal queNombre As String) As String
            Dim resultado As String = ""
            Try
                Dim xtipos() As Short = Nothing
                Dim xdatos() As Object = Nothing
                Dim objA As AcadObject = Nothing
                Try
                    objA = oAppA.ActiveDocument.HandleToObject(oHandle)
                Catch ex As Exception
                    Return ""
                    Exit Function
                End Try
                '
                objA.GetXData(regAPPA, xtipos, xdatos)
                '' Si el objeto no tiene XData
                If xdatos Is Nothing Then
                    XNuevo(objA)
                    objA.GetXData(regAPPA, xtipos, xdatos)
                End If
                '
                Dim todo As String = xdatos(1)  '' Cadena de texto con todas los datos.
                If queNombre = "" Then
                    resultado = todo
                ElseIf queNombre = regAPPA Then
                    resultado = xdatos(0)
                ElseIf todo.Contains(queNombre) Then
                    Dim valoresdatos() As String = todo.Split("|"c)      '' cada elemento nombre=valor
                    For x As Integer = 0 To UBound(valoresdatos)
                        Dim nombre As String = valoresdatos(x).Split("="c)(0)
                        Dim valor As String = valoresdatos(x).Split("="c)(1)
                        If nombre.Equals(queNombre) Then
                            resultado = valor
                            Exit For
                        End If
                    Next
                Else
                    resultado = ""
                End If
            Catch ex As Exception

            End Try
            Return resultado
        End Function

        Public Function XLeeDato(ByVal oId As Long, ByVal queNombre As String) As String
            Dim resultado As String = ""
            Dim objA As AcadObject = Nothing
            Try
                objA = oAppA.ActiveDocument.HandleToObject(oId)
            Catch ex As Exception
                Return ""
                Exit Function
            End Try
            '
            resultado = XLeeDato(objA.Handle, queNombre)
            Return resultado
        End Function

        Public Function XLeeDato(ByVal objA As AcadObject, ByVal queNombre As String) As String
            Dim resultado As String = ""
            resultado = XLeeDato(objA.Handle, queNombre)
            Return resultado
        End Function

        Public Function XLeeDatoTodosNombresValores(ByVal objA As AcadObject) As String
            Dim resultado As String = ""
            Try
                Dim xtipos() As Short = Nothing
                Dim xdatos() As Object = Nothing
                objA.GetXData(regAPPA, xtipos, xdatos)
                '' Si el objeto no tiene XData
                If xdatos Is Nothing Then
                    XNuevo(objA)
                    objA.GetXData(regAPPA, xtipos, xdatos)
                End If
                '
                Dim todo As String = xdatos(1)  '' Cadena de texto con todas los datos.
                resultado = xdatos(0) & " = " & xdatos(0) & vbCrLf
                '
                If todo <> "" Then
                    Dim valoresdatos() As String = todo.Split(";"c)      '' cada elemento nombre=valor
                    For x As Integer = 0 To UBound(valoresdatos)
                        Dim nombre As String = valoresdatos(x).Split("="c)(0)
                        Dim valor As String = valoresdatos(x).Split("="c)(1)
                        resultado &= vbTab & nombre & " = " & valor & vbCrLf
                    Next
                End If
            Catch ex As Exception

            End Try

            '
            Return resultado
        End Function

        Public Sub XPonTiposDatosArrays(ByVal objA As AcadObject, ByVal tipo As Short(), ByVal dato As Object())
            '' Poner xdata al objeto. Solo para SERICAD
            Using Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument()
                Call objA.SetXData(tipo, dato)
            End Using
            Try
                objA.Update()
            Catch ex As Exception
            End Try
        End Sub

        Public Sub XPonDatosTodoString(ByVal objA As AcadObject, ByVal datos As String)
            Dim xtipos() As Short = Nothing
            Dim xdatos() As Object = Nothing
            objA.GetXData(regAPPA, xtipos, xdatos)

            '' Si el objeto no tiene XData o si tienen un numéro de elementos diferente.
            If xdatos Is Nothing Then
                XNuevo(objA)
                xdatos(1) = datos
                Using Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument()
                    objA.SetXData(xtipos, xdatos)
                End Using
            End If
            objA.Update()
        End Sub

        'Public Sub XPonDato(ByVal objA As Object, queNombre As String, queValor As String,
        '                Optional crear As Boolean = True)
        '    If objA Is Nothing Then Exit Sub
        '    '
        '    Dim xtipos() As Short = Nothing
        '    Dim xdatos() As Object = Nothing
        '    objA.GetXData(regAPPA, xtipos, xdatos)
        '    '' Si el objeto no tiene XData
        '    If xdatos Is Nothing Then
        '        XNuevo(CType(objA, AcadObject))
        '        objA.GetXData(regAPPA, xtipos, xdatos)
        '    End If
        '    '
        '    Dim encontrado As Boolean = False
        '    Dim todo As String = xdatos(1).ToString  '' Cadena de texto con todas los datos.
        '    '
        '    If queNombre = "" Then
        '        ' No hacemos nada
        '    ElseIf queNombre = regAPPA And xdatos(0) <> regAPPA Then
        '        xdatos(0) = regAPPA
        '    ElseIf todo.Contains(queNombre) Then
        '        Dim valoresdatos() As String = todo.Split(";"c)      '' cada elemento nombre=valor
        '        For x As Integer = 0 To UBound(valoresdatos)
        '            Dim nombrevalor As String = valoresdatos(x)
        '            Dim nombre As String = nombrevalor.Split("="c)(0)
        '            Dim valor As String = nombrevalor.Split("="c)(1)
        '            If nombre.Equals(queNombre) Then
        '                Dim nombrevalornuevo = nombre & "=" & queValor
        '                todo = todo.Replace(nombrevalor, nombrevalornuevo)
        '                xdatos(1) = todo
        '                CType(objA, AcadObject).SetXData(xtipos, xdatos)
        '                encontrado = True
        '                Exit For
        '            End If
        '        Next
        '    End If
        '    '
        '    If encontrado = False And crear = True Then
        '        If xdatos(1) = "" Then
        '            xdatos(1) = queNombre & "=" & queValor
        '        Else
        '            xdatos(1) &= ";" & queNombre & "=" & queValor
        '        End If
        '        CType(objA, AcadObject).SetXData(xtipos, xdatos)
        '    End If
        '    objA.Update()
        'End Sub

        'Public Sub XPonDato(ByVal objA As AcadEntity, queNombre As String, queValor As String,
        '                Optional crear As Boolean = True)
        '    Dim xtipos() As Short = Nothing
        '    Dim xdatos() As Object = Nothing
        '    objA.GetXData(regAPPA, xtipos, xdatos)
        '    '' Si el objeto no tiene XData
        '    If xdatos Is Nothing Then
        '        XNuevo(objA)
        '        objA.GetXData(regAPPA, xtipos, xdatos)
        '    End If
        '    '
        '    Dim encontrado As Boolean = False
        '    Dim todo As String = xdatos(1).ToString  '' Cadena de texto con todas los datos.
        '    '
        '    If queNombre = "" Then
        '        ' No hacemos nada
        '    ElseIf queNombre = regAPPA And xdatos(0) <> regAPPA Then
        '        xdatos(0) = regAPPA
        '    ElseIf todo.Contains(queNombre) Then
        '        Dim valoresdatos() As String = todo.Split(";"c)      '' cada elemento nombre=valor
        '        For x As Integer = 0 To UBound(valoresdatos)
        '            Dim nombrevalor As String = valoresdatos(x)
        '            Dim nombre As String = nombrevalor.Split("="c)(0)
        '            Dim valor As String = nombrevalor.Split("="c)(1)
        '            If nombre.Equals(queNombre) Then
        '                Dim nombrevalornuevo = nombre & "=" & queValor
        '                todo = todo.Replace(nombrevalor, nombrevalornuevo)
        '                xdatos(1) = todo
        '                objA.SetXData(xtipos, xdatos)
        '                encontrado = True
        '                Exit For
        '            End If
        '        Next
        '    End If
        '    '
        '    If encontrado = False And crear = True Then
        '        If xdatos(1) = "" Then
        '            xdatos(1) = queNombre & "=" & queValor
        '        Else
        '            xdatos(1) &= ";" & queNombre & "=" & queValor
        '        End If
        '        objA.SetXData(xtipos, xdatos)
        '    End If
        '    objA.Update()
        'End Sub
        'Public Sub XPonDato(ByVal objA As AcadObject, queNombre As String, queValor As String,
        '                Optional crear As Boolean = True)
        '    Dim xtipos() As Short = Nothing
        '    Dim xdatos() As Object = Nothing
        '    objA.GetXData(regAPPA, xtipos, xdatos)
        '    '' Si el objeto no tiene XData
        '    If xdatos Is Nothing Then
        '        XNuevo(objA)
        '        objA.GetXData(regAPPA, xtipos, xdatos)
        '    End If
        '    '
        '    Dim encontrado As Boolean = False
        '    Dim todo As String = xdatos(1).ToString  '' Cadena de texto con todas los datos.
        '    '
        '    If queNombre = "" Then
        '        ' No hacemos nada
        '    ElseIf queNombre = regAPPA And xdatos(0) <> regAPPA Then
        '        xdatos(0) = regAPPA
        '    ElseIf todo.Contains(queNombre) Then
        '        Dim valoresdatos() As String = todo.Split(";"c)      '' cada elemento nombre=valor
        '        For x As Integer = 0 To UBound(valoresdatos)
        '            Dim nombrevalor As String = valoresdatos(x)
        '            Dim nombre As String = nombrevalor.Split("="c)(0)
        '            Dim valor As String = nombrevalor.Split("="c)(1)
        '            If nombre.Equals(queNombre) Then
        '                Dim nombrevalornuevo = nombre & "=" & queValor
        '                todo = todo.Replace(nombrevalor, nombrevalornuevo)
        '                xdatos(1) = todo
        '                objA.SetXData(xtipos, xdatos)
        '                encontrado = True
        '                Exit For
        '            End If
        '        Next
        '    End If
        '    '
        '    If encontrado = False And crear = True Then
        '        If xdatos(1) = "" Then
        '            xdatos(1) = queNombre & "=" & queValor
        '        Else
        '            xdatos(1) &= ";" & queNombre & "=" & queValor
        '        End If
        '        objA.SetXData(xtipos, xdatos)
        '    End If
        '    objA.Update()
        'End Sub

        Public Sub XPonDato(oId As Long, queNombre As String, queValor As String,
                        Optional crear As Boolean = True)
            If IsNumeric(oId) = False Then Exit Sub
            '
            Dim xtipos() As Short = Nothing
            Dim xdatos() As Object = Nothing
            Try
                Dim objA As AcadObject = oAppA.ActiveDocument.ObjectIdToObject(oId)
                objA.GetXData(regAPPA, xtipos, xdatos)
                '' Si el objeto no tiene XData
                If xdatos Is Nothing Then
                    XNuevo(objA)
                    objA.GetXData(regAPPA, xtipos, xdatos)
                End If
                '
                Dim encontrado As Boolean = False
                Dim todo As String = xdatos(1).ToString  '' Cadena de texto con todas los datos.
                '
                If queNombre = "" Then
                    ' No hacemos nada
                ElseIf queNombre = regAPPA And xdatos(0) <> regAPPA Then
                    xdatos(0) = regAPPA
                ElseIf todo.Contains(queNombre) Then
                    Dim valoresdatos() As String = todo.Split(";"c)      '' cada elemento nombre=valor
                    For x As Integer = 0 To UBound(valoresdatos)
                        Dim nombrevalor As String = valoresdatos(x)
                        Dim nombre As String = nombrevalor.Split("="c)(0)
                        Dim valor As String = nombrevalor.Split("="c)(1)
                        If nombre.Equals(queNombre) Then
                            Dim nombrevalornuevo = nombre & "=" & queValor
                            todo = todo.Replace(nombrevalor, nombrevalornuevo)
                            xdatos(1) = todo

                            Using Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument()
                                objA.SetXData(xtipos, xdatos)
                            End Using
                            encontrado = True
                            Exit For
                        End If
                    Next
                End If
                '
                If encontrado = False And crear = True Then
                    If xdatos(1).ToString = "" Then
                        xdatos(1) = queNombre & "=" & queValor
                    Else
                        xdatos(1) &= ";" & queNombre & "=" & queValor
                    End If
                    objA.SetXData(xtipos, xdatos)
                End If
                objA.Update()
            Catch ex As Exception

            End Try
        End Sub

        Public Sub XPonDato(oHandle As String, queNombre As String, queValor As String,
                        Optional crear As Boolean = True)
            Dim xtipos() As Short = Nothing
            Dim xdatos() As Object = Nothing
            Try
                Dim objA As AcadObject = oAppA.ActiveDocument.HandleToObject(oHandle)
                objA.GetXData(regAPPA, xtipos, xdatos)
                '' Si el objeto no tiene XData
                If xdatos Is Nothing Then
                    XNuevo(objA)
                    objA.GetXData(regAPPA, xtipos, xdatos)
                End If
                '
                Dim encontrado As Boolean = False
                Dim todo As String = xdatos(1).ToString  '' Cadena de texto con todas los datos.
                '
                If queNombre = "" Then
                    ' No hacemos nada
                ElseIf queNombre = regAPPA And xdatos(0) <> regAPPA Then
                    xdatos(0) = regAPPA
                ElseIf todo.Contains(queNombre) Then
                    Dim valoresdatos() As String = todo.Split("|"c)      '' cada elemento nombre=valor
                    For x As Integer = 0 To UBound(valoresdatos)
                        Dim nombrevalor As String = valoresdatos(x)
                        Dim nombre As String = nombrevalor.Split("="c)(0)
                        Dim valor As String = nombrevalor.Split("="c)(1)
                        If nombre.Equals(queNombre) Then
                            Dim nombrevalornuevo = nombre & "=" & queValor
                            todo = todo.Replace(nombrevalor, nombrevalornuevo)
                            xdatos(1) = todo
                            Using Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument()
                                objA.SetXData(xtipos, xdatos)
                            End Using
                            encontrado = True
                            Exit For
                        End If
                    Next
                End If
                '
                If encontrado = False And crear = True Then
                    If xdatos(1).ToString = "" Then
                        xdatos(1) = queNombre & "=" & queValor
                    Else
                        xdatos(1) &= "|" & queNombre & "=" & queValor
                    End If
                    objA.SetXData(xtipos, xdatos)
                End If
                objA.Update()
            Catch ex As Exception

            End Try

        End Sub
        '
        Public Function XEsApp(ByVal objA As AcadObject) As Boolean
            Dim resultado As Boolean = False
            Dim xtipos() As Short = Nothing
            Dim xdatos() As Object = Nothing
            Try
                objA.GetXData(regAPPA, xtipos, xdatos)

                If xdatos Is Nothing Then
                    resultado = False
                ElseIf xdatos(0).ToString = regAPPA Then
                    resultado = True
                End If
            Catch ex As Exception
                '
            End Try
            '
            Return resultado
        End Function
        '' EJEMPLOS XData....
        'DataType(0) = 1001 : Data(0) = "Aplicacion Reg."       ' Aplicacion Registrada
        'DataType(1) = 1000 : Data(1) = "Un texto...."          ' string
        'DataType(2) = 1003 : Data(2) = "0"                     ' layer
        'DataType(3) = 1040 : Data(3) = 1.23479137438413E+40    ' real
        'DataType(4) = 1041 : Data(4) = 1237324938              ' distance
        'DataType(5) = 1070 : Data(5) = 32767                   ' 16 bit Integer
        'DataType(6) = 1071 : Data(6) = 32767                   ' 32 bit Integer
        'DataType(7) = 1042 : Data(7) = 10                      ' scaleFactor
        '' Faltarían, si se quiere, las cadenas "{" y "}" de inicio y fin de lista. (1002)

        Public Function DameBloquesTODOS(Optional queCapas As String = "*", Optional queNombre As String = "*", Optional elID As Boolean = False) As ArrayList
            Dim resultado As New ArrayList
            '
            Dim F1(3) As Short
            Dim F2(3) As Object
            Dim vF1 As Object = Nothing
            Dim vF2 As Object = Nothing
            ' Las 2 maneras valen igual. AcDbBlckReference es mejor (Solo coge bloques) INSERT coge sombreados también.
            F1(0) = 100 : F2(0) = "AcDbBlockReference,AcDbEntity"
            F1(1) = 0 : F2(1) = "INSERT"
            F1(2) = 8 : F2(2) = queCapas
            F1(3) = 2 : F2(3) = queNombre
            '
            vF1 = F1
            vF2 = F2
            ''
            Try
                oSel = oAppA.ActiveDocument.SelectionSets.Item(nSel)
            Catch ex As System.Exception
                oSel = oAppA.ActiveDocument.SelectionSets.Add(nSel)
            End Try
            ''
            oSel.Clear()
            '
            oAppA.ActiveDocument.SetVariable("pickadd", 2)   '' Solo una selección. Se quita lo que hubiera
            Try
                oSel.Select(AcSelect.acSelectionSetAll, , , vF1, vF2)
            Catch ex As System.Exception
                Debug.Print(ex.Message)
            End Try
            ''
            If oSel.Count > 0 Then
                For Each oEnt As AcadEntity In oSel
                    If Not (TypeOf oEnt Is AcadBlockReference) Then Continue For
                    'Dim oBl As AcadBlockReference = oAppA.ActiveDocument.ObjectIdToObject(oEnt.ObjectID)
                    If elID Then
                        resultado.Add(CType(oEnt, AcadBlockReference).ObjectID) 'oBl.ObjectID)
                    Else
                        resultado.Add(CType(oEnt, AcadBlockReference))  'oBl)
                    End If
                    'oBl = Nothing
                Next
            End If
            '
            oAppA.ActiveDocument.SetVariable("pickadd", 2)   '' La seleccion actual se suma a la que hubiera.
            Return resultado
        End Function
        Public Function DameBloquesTODOS_Capa(Optional queCapas As String = "*", Optional elID As Boolean = False) As ArrayList
            Dim resultado As New ArrayList
            '
            Dim F1(2) As Short
            Dim F2(2) As Object
            Dim vF1 As Object = Nothing
            Dim vF2 As Object = Nothing
            ' Las 2 maneras valen igual. AcDbBlckReference es mejor (Solo coge bloques) INSERT coge sombreados también.
            F1(0) = 100 : F2(0) = "AcDbBlockReference,AcDbEntity"
            F1(1) = 0 : F2(1) = "INSERT"
            F1(2) = 8 : F2(2) = queCapas
            '
            vF1 = F1
            vF2 = F2
            ''
            Try
                oSel = oAppA.ActiveDocument.SelectionSets.Item(nSel)
            Catch ex As System.Exception
                oSel = oAppA.ActiveDocument.SelectionSets.Add(nSel)
            End Try
            ''
            oSel.Clear()
            '
            oAppA.ActiveDocument.SetVariable("pickadd", 2)   '' Solo una selección. Se quita lo que hubiera
            Try
                oSel.Select(AcSelect.acSelectionSetAll, , , vF1, vF2)
            Catch ex As System.Exception
                Debug.Print(ex.Message)
            End Try
            ''
            If oSel.Count > 0 Then
                For Each oEnt As AcadEntity In oSel
                    If Not (TypeOf oEnt Is AcadBlockReference) Then Continue For
                    'Dim oBl As AcadBlockReference = oAppA.ActiveDocument.ObjectIdToObject(oEnt.ObjectID)
                    If elID Then
                        resultado.Add(CType(oEnt, AcadBlockReference).ObjectID) 'oBl.ObjectID)
                    Else
                        resultado.Add(CType(oEnt, AcadBlockReference))  'oBl)
                    End If
                    'oBl = Nothing
                Next
            End If
            '
            oAppA.ActiveDocument.SetVariable("pickadd", 2)   '' La seleccion actual se suma a la que hubiera.
            Return resultado
        End Function

        Public Function DamePolilineasTODAS_CAPA(Optional queCapas As String = "*", Optional elID As Boolean = False) As ArrayList
            Dim resultado As New ArrayList
            '
            Dim F1(2) As Short
            Dim F2(2) As Object
            Dim vF1 As Object = Nothing
            Dim vF2 As Object = Nothing
            '
            F1(0) = 100 : F2(0) = "AcDbPolyline"
            F1(1) = 0 : F2(1) = "*POLYLINE*"    '"LWPOLYLINE"
            F1(2) = 8 : F2(2) = queCapas
            vF1 = F1
            vF2 = F2
            '
            Try
                oSel = oAppA.ActiveDocument.SelectionSets.Item(nSel)
            Catch ex As System.Exception
                oSel = oAppA.ActiveDocument.SelectionSets.Add(nSel)
            End Try
            oSel.Clear()
            '
            oAppA.ActiveDocument.SetVariable("pickadd", 2)   '' Solo una selección. Se quita lo que hubiera
            '
            Try
                oSel.Select(AcSelect.acSelectionSetAll,,, vF1, vF2)
            Catch ex As System.Exception
                'Debug.Print(ex.Message)
            End Try
            'MsgBox(oSel.Count)
            If oSel IsNot Nothing AndAlso oSel.Count > 0 Then
                For Each oEnt As AcadEntity In oSel
                    If TypeOf oEnt IsNot AcadLWPolyline Then Continue For
                    '
                    If elID Then
                        resultado.Add(CType(oEnt, AcadLWPolyline).ObjectID)
                    Else
                        resultado.Add(CType(oEnt, AcadLWPolyline))
                    End If
                Next
            Else
                resultado = Nothing
            End If
            '
            oAppA.ActiveDocument.SetVariable("pickadd", 2)   '' La seleccion actual se suma a la que hubiera.
            Return resultado
        End Function

        'Nuevo por afleta
        Public Function DameBloquesTODOS_XData(queNombre As String, queValor As String, Optional queCapas As String = "*", Optional elID As Boolean = False) As ArrayList
            Dim resultado As New ArrayList
            '
            Dim F1(2) As Short
            Dim F2(2) As Object
            Dim vF1 As Object = Nothing
            Dim vF2 As Object = Nothing
            ' Las 2 maneras valen igual. AcDbBlckReference es mejor (Solo coge bloques) INSERT coge sombreados también.
            F1(0) = 100 : F2(0) = "AcDbBlockReference,AcDbEntity"
            F1(1) = 0 : F2(1) = "INSERT"
            F1(2) = 8 : F2(2) = queCapas
            '
            vF1 = F1
            vF2 = F2
            ''
            Try
                oSel = oAppA.ActiveDocument.SelectionSets.Item(nSel)
            Catch ex As System.Exception
                oSel = oAppA.ActiveDocument.SelectionSets.Add(nSel)
            End Try
            ''
            oSel.Clear()
            '
            oAppA.ActiveDocument.SetVariable("pickadd", 2)   '' Solo una selección. Se quita lo que hubiera
            Try
                oSel.Select(AcSelect.acSelectionSetAll, , , vF1, vF2)
            Catch ex As System.Exception
                Debug.Print(ex.Message)
            End Try
            ''
            If oSel.Count > 0 Then
                For Each oEnt As AcadEntity In oSel
                    If Not (TypeOf oEnt Is AcadBlockReference) Then Continue For

                    Dim strDato = Me.XLeeDato(oEnt.Handle, queNombre)
                    If strDato.Trim().ToUpper() = queValor.Trim().ToUpper() Then
                        If elID Then
                            resultado.Add(CType(oEnt, AcadBlockReference).ObjectID) 'oBl.ObjectID)
                        Else
                            resultado.Add(CType(oEnt, AcadBlockReference))  'oBl)
                        End If
                    End If
                    'oBl = Nothing
                Next
            End If
            '
            oAppA.ActiveDocument.SetVariable("pickadd", 2)   '' La seleccion actual se suma a la que hubiera.
            Return resultado
        End Function
        '

        'Nuevo por afleta
        Public Function DameTODOS_XData(queNombre As String, queValor As String, Optional elID As Boolean = False) As ArrayList
            Dim resultado As New ArrayList
            '
            Dim F1(2) As Short
            Dim F2(2) As Object
            Dim vF1 As Object = Nothing
            Dim vF2 As Object = Nothing
            ' Las 2 maneras valen igual. AcDbBlckReference es mejor (Solo coge bloques) INSERT coge sombreados también.
            F1(0) = 100 : F2(0) = "AcDbEntity"
            F1(1) = 0 : F2(1) = "*"
            F1(2) = 8 : F2(2) = "*"
            '
            vF1 = F1
            vF2 = F2
            ''
            Try
                oSel = oAppA.ActiveDocument.SelectionSets.Item(nSel)
            Catch ex As System.Exception
                oSel = oAppA.ActiveDocument.SelectionSets.Add(nSel)
            End Try
            ''
            oSel.Clear()
            '
            oAppA.ActiveDocument.SetVariable("pickadd", 2)   '' Solo una selección. Se quita lo que hubiera
            Try
                oSel.Select(AcSelect.acSelectionSetAll, , , vF1, vF2)
            Catch ex As System.Exception
                Debug.Print(ex.Message)
            End Try
            ''
            If oSel.Count > 0 Then
                For Each oEnt As AcadEntity In oSel
                    If Not (TypeOf oEnt Is AcadBlockReference) Then Continue For

                    Dim strDato = Me.XLeeDato(oEnt.Handle, queNombre)
                    If strDato.Trim().ToUpper() = queValor.Trim().ToUpper() Then
                        If elID Then
                            resultado.Add(CType(oEnt, AcadBlockReference).ObjectID) 'oBl.ObjectID)
                        Else
                            resultado.Add(CType(oEnt, AcadBlockReference))  'oBl)
                        End If
                    End If
                    'oBl = Nothing
                Next
            End If
            '
            oAppA.ActiveDocument.SetVariable("pickadd", 2)   '' La seleccion actual se suma a la que hubiera.
            Return resultado
        End Function
    End Class
End Namespace