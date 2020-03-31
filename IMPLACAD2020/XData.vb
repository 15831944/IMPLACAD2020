Imports System.Diagnostics
Imports System.Collections
Imports System.Windows.Forms

Imports Autodesk.AutoCAD.Interop
Imports Autodesk.AutoCAD.Interop.Common

Module XData
    ''***** XDatos que pondremos por defecto al crear los objetos Polilinea, Bloque y Muro.
    '' habrá que convertir a double para realizar operaciones FormatNumber(altura,2,tristate)
    ' 0 podrá tener los valores INSERT, POLYLINE, CIRCLE, LWPOLILINE, etc.
    ' 1001=regApp
    ' 1003 = Capa
    ' 1000=Clase;Tipo;Tipo1;Tipo2;Tipo3
    Public xTTodo As Short() = New Short() {1001, 1000}
    Public xDTodo As Object() = New Object() {regAPP, ""}
    'Public xDPol As Object() = New Object() {regAPP, 0, "0", "Clase=polilinea"}
    'Public xTBlo As Short() = New Short() {1001, 1040, 1003, 1000}
    'Public xDBlo As Object() = New Object() {regAPP, 0, "0", "Clase=etiqueta"}
    'Public xTTab As Short() = New Short() {1001, 1040, 1003, 1000}
    'Public xDTab As Object() = New Object() {regAPP, 0, "0", "Clase=tabla"}
    'Public xDTabesc As Object() = New Object() {regAPP, 0, "0", "Clase=tablaescaleras"}
    'Public xTBSue As Short() = New Short() {1001, 1040, 1003, 1000}
    'Public xDBSue As Object() = New Object() {regAPP, 0, "0", "Clase=balizasuelo"}
    'Public xTBPar As Short() = New Short() {1001, 1040, 1003, 1000}
    'Public xDBPar As Object() = New Object() {regAPP, 0, "0", "Clase=balizapared"}
    'Public xTBEsc As Short() = New Short() {1001, 1040, 1003, 1000}
    'Public xDBEsc As Object() = New Object() {regAPP, 0, "0", "Clase=balizaescalera"}
    ''
    Public Sub XNuevo(ByRef objA As AcadObject, Optional queTipos As String = "")
        If queTipos <> "" Then
            xDTodo(1) = queTipos
        End If
        '' Pone los XData para IMPLACAD. Solo si un objeto no tiene XData.
        XPonTodo(objA, xTTodo, xDTodo)
    End Sub

    Public Sub XBorrar(ByVal objA As AcadObject)
        Dim DataType(0) As Short
        Dim Data(0) As Object
        DataType(0) = 1001 : Data(0) = regAPP
        'If TypeOf objA Is AcadLWPolyline Then
        '    objA = CType(objA, AcadLWPolyline)
        'ElseIf TypeOf objA Is AcadBlockReference Then
        '    objA = CType(objA, AcadBlockReference)
        'ElseIf TypeOf objA Is AcadTable Then
        '    objA = CType(objA, AcadTable)
        'End If
        objA.SetXData(DataType, Data)
        objA.Update()
    End Sub

    Public Sub XLeeMensaje(ByVal obj As AcadObject)
        '' Leer xdata del objeto
        Dim xtipos As Object = Nothing
        Dim xdatos As Object = Nothing
        obj.GetXData(regAPP, xtipos, xdatos)

        Dim mensaje As String = ""
        For x As Integer = 0 To UBound(xdatos)
            mensaje &= xdatos(x) & vbCrLf
        Next
        If mensaje = "" Then
            Dim r As DialogResult = MessageBox.Show("El objeto no tiene XData...")
        Else
            MessageBox.Show(mensaje)
        End If
    End Sub

    Public Function XLeeTiposDatos(ByVal objA As AcadObject, Optional ByVal app As String = "") As Object()         'Devuelve todos los datos de SERICAD 0-SERICAD, 1-CAPA
        If app = "" Then app = regAPP
        Dim resul(1) As Object
        ' Leer xdata del objeto
        Dim xtipos() As Short = Nothing
        Dim xdatos() As Object = Nothing
        objA.GetXData(app, xtipos, xdatos)
        '' Si el objeto no tiene XData
        If xdatos Is Nothing Then
            XNuevo(objA)
            objA.GetXData(app, xtipos, xdatos)
        End If
        resul(0) = xtipos : resul(1) = xdatos
        Return resul
    End Function

    Public Function XLeeDatos(ByVal objA As AcadObject, Optional ByVal app As String = "") As Object         'Devuelve todos los datos de SERICAD 0-SERICAD, 1-CAPA
        If app = "" Then app = regAPP
        ' Leer xdata del objeto
        Dim xtipos() As Short = Nothing
        Dim xdatos() As Object = Nothing
        objA.GetXData(app, xtipos, xdatos)
        '' Si el objeto no tiene XData
        If xdatos Is Nothing Then
            XNuevo(objA)
            objA.GetXData(app, xtipos, xdatos)
        End If
        Return xdatos
    End Function
    Public Function XLeeDatos(ByVal oId As Long, Optional ByVal app As String = "") As Object         'Devuelve todos los datos de SERICAD 0-SERICAD, 1-CAPA
        Return XLeeDatos(CType(oApp.ActiveDocument.ObjectIdToObject(oId), AcadObject), app)
    End Function

    Public Function XLeeDatos(ByVal oHandle As String, Optional ByVal app As String = "") As Object         'Devuelve todos los datos de SERICAD 0-SERICAD, 1-CAPA
        Return XLeeDatos(CType(oApp.ActiveDocument.HandleToObject(oHandle), AcadObject), app)
    End Function
    Public Function XLeeDatos(ByVal objA As AcadEntity, Optional ByVal app As String = "") As Object         'Devuelve todos los datos de SERICAD 0-SERICAD, 1-CAPA
        Return XLeeDatos(CType(objA, AcadObject), app)
    End Function
    '
    Public Function XLeeDato(ByVal objA As AcadObject, ByVal queNombre As String,
                              Optional SoloValor As Boolean = True) As String
        Dim resultado As String = ""
        Try
            Dim xtipos() As Short = Nothing
            Dim xdatos() As Object = Nothing
            '
            objA.GetXData(regAPP, xtipos, xdatos)
            '' Si el objeto no tiene XData
            If xtipos Is Nothing OrElse xdatos Is Nothing Then
                Return resultado
            End If
            ' Cambiar los XData viejos (4 datos) al nuevo con (2 datos. RegApp y Textos)
            If xdatos.Length = 4 Then
                ' Coger el texto que había en el 4º elemento (3)
                Dim Txt4 As String = xdatos(3)
                Txt4 = Txt4.Replace("vacio", "")
                ' Crear XData con 1001 y 1000 (poner ya el valor de 1000 = Txt4)
                XNuevo(objA, Txt4)
                objA.GetXData(regAPP, xtipos, xdatos)
            End If
            ' Tiene XData
            'Dim App As String = xdatos(0)   ' Nombre de la App registrada
            Dim todo As String = ""  ' xdatos(1)  '' Cadena de texto con todas los datos de texto, concatenados con |.
            Dim indice As Integer = -1
            For x As Integer = 0 To xtipos.Length - 1
                Dim oTipo As Short = xtipos(x)
                If oTipo = 1000 Then
                    todo = xdatos(x)
                    indice = x
                    Exit For
                End If
            Next
            '
            If queNombre = "" Then
                resultado = todo
            ElseIf queNombre = regAPP Then
                resultado = xdatos(0)
            ElseIf todo.Contains(queNombre) Then
                Dim valoresdatos() As String = todo.Split("|"c)      '' cada elemento nombre=valor|nombre1=valor1
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
        If resultado <> "" AndAlso SoloValor = True AndAlso resultado.Contains("=") Then
            resultado = resultado.Split("=")(1)
        End If
        Return resultado
    End Function

    Public Function XLeeDato(ByVal oId As Long, ByVal queNombre As String) As String
        Return XLeeDato(oApp.ActiveDocument.ObjectIdToObject(oId), queNombre)
    End Function

    Public Function XLeeDato(ByVal oHandle As String, ByVal queNombre As String) As String
        Return XLeeDato(oApp.ActiveDocument.HandleToObject(oHandle), queNombre)
    End Function
    Public Function XLeeDato(ByVal oEnt As AcadEntity, ByVal queNombre As String) As String
        Return XLeeDato(CType(oEnt, AcadObject), queNombre)
    End Function

    Public Sub XPonTodo(ByVal objA As AcadObject, ByVal tipo As Short(), ByVal dato As Object())
        '' Poner xdata al objeto. Solo para SERICAD
        Call objA.SetXData(tipo, dato)
        '' Tipo de objeto al que vamos a poner los datos
        If TypeOf objA Is AcadLWPolyline Then
            objA = CType(objA, AcadLWPolyline)
        ElseIf TypeOf objA Is AcadBlockReference Then
            objA = CType(objA, AcadBlockReference)
        End If
        objA.Update()
    End Sub

    Public Sub XPonDato(objA As AcadObject, queNombre As String, queValor As String,
                        Optional crear As Boolean = True)
        Dim xtipos() As Short = Nothing
        Dim xdatos() As Object = Nothing
        Try
            objA.GetXData(regAPP, xtipos, xdatos)
            '' Si el objeto no tiene XData
            If xdatos Is Nothing Then
                XNuevo(objA)
                objA.GetXData(regAPP, xtipos, xdatos)
            End If
            '
            Dim encontrado As Boolean = False
            Dim todo As String = ""  ' xdatos(1)  '' Cadena de texto con todas los datos de texto, concatenados con |.
            Dim indice As Integer = -1
            For x As Integer = 0 To xtipos.Length - 1
                Dim oTipo As Short = xtipos(x)
                If oTipo = 1000 Then
                    todo = xdatos(x)
                    todo.Replace("vacio|", "")
                    todo.Replace("vacio;", "")
                    todo.Replace("vacio", "")
                    indice = x
                    Exit For
                End If
            Next
            '
            If queNombre = "" Then
                ' No hacemos nada
            ElseIf queNombre = regAPP And xdatos(0) <> regAPP Then
                xdatos(0) = regAPP
            ElseIf todo.Contains(queNombre) Then
                Dim valoresdatos() As String = todo.Split("|"c)      '' cada elemento nombre=valor|nombre1=valor1
                For x As Integer = 0 To UBound(valoresdatos)
                    Dim nombrevalor As String = valoresdatos(x)
                    Dim nombre As String = nombrevalor.Split("="c)(0)
                    Dim valor As String = nombrevalor.Split("="c)(1)
                    If nombre.Equals(queNombre) Then
                        Dim nombrevalornuevo = nombre & "=" & queValor
                        todo = todo.Replace(nombrevalor, nombrevalornuevo)
                        xdatos(indice) = todo

                        Using Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument()
                            objA.SetXData(xtipos, xdatos)
                        End Using
                        encontrado = True
                        Exit For
                    End If
                Next
            End If
            '
            If encontrado = False And crear = True AndAlso indice > -1 Then
                If xdatos(indice).ToString = "" OrElse xdatos(indice).ToString = "vacio" Then
                    xdatos(indice) = queNombre & "=" & queValor
                Else
                    xdatos(indice) &= "|" & queNombre & "=" & queValor
                End If
                objA.SetXData(xtipos, xdatos)
            End If
            objA.Update()
        Catch ex As Exception
        End Try
    End Sub

    Public Sub XPonDato(oHandle As String, queNombre As String, queValor As String,
                        Optional crear As Boolean = True)
        XPonDato(oApp.ActiveDocument.HandleToObject(oHandle), queNombre, queValor, crear)
    End Sub

    Public Sub XPonDato(oId As Long, queNombre As String, queValor As String,
                        Optional crear As Boolean = True)
        XPonDato(oApp.ActiveDocument.ObjectIdToObject(oId), queNombre, queValor, crear)
    End Sub

    Public Sub XPonValoresArr(ByVal objA As AcadObject, ByVal datos() As Object, Optional ByVal app As String = "")
        'XDataBorrar(objA)
        If app = "" Then app = regAPP
        Dim xtipos() As Short = Nothing
        Dim xdatos() As Object = Nothing
        objA.GetXData(app, xtipos, xdatos)

        '' Si el objeto no tiene XData o si tienen un numéro de elementos diferente.
        If xdatos Is Nothing Then
            XNuevo(objA)
            objA.GetXData(app, xtipos, xdatos)
        ElseIf UBound(xdatos) <> UBound(datos) Then
            XNuevo(objA)
            objA.GetXData(app, xtipos, xdatos)
        End If
        '' Recorremos los datos y solo pondremos los que lleven algún alor.
        For x As Integer = LBound(datos) To UBound(datos)
            If IsNumeric(datos(x)) = True Then
                xdatos(x) = datos(x)
            ElseIf datos(x) <> "" Then
                xdatos(x) = datos(x)
            End If
        Next

        '' Poner xdata al objeto
        Call objA.SetXData(xtipos, xdatos)
        '' Tipo de objeto al que vamos a poner los datos
        If TypeOf objA Is AcadLWPolyline Then
            objA = CType(objA, AcadLWPolyline)
            'XPonTodo(objA, xTPol, xDPol)
        ElseIf TypeOf objA Is AcadBlockReference Then
            objA = CType(objA, AcadBlockReference)
            'XPonTodo(objA, xTBlo, xDBlo)
        End If
        objA.Update()
    End Sub

    Public Function XEsApp(ByVal objA As AcadObject, Optional ByVal app As String = "") As Boolean
        Dim resultado As Boolean = False
        If app = "" Then app = regAPP
        Dim xtipos() As Short = Nothing
        Dim xdatos() As Object = Nothing
        objA.GetXData(app, xtipos, xdatos)

        If xdatos Is Nothing Then
            resultado = False
        ElseIf xdatos(0) = app Then
            resultado = True
        End If
        Return resultado
    End Function
End Module

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