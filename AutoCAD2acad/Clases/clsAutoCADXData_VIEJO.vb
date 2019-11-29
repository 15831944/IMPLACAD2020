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
Partial Public Class clsAutoCAD2acad
    ''***** XDatos que pondremos por defecto al crear los objetos Polilinea, Bloque y Muro.
    '' El dato(3) corresponde a la altura de la zona. La pondremos como cadena para evitar conversiones
    '' habrá que convertir a double para realizar operaciones FormatNumber(altura,2,tristate)
    Public xTPol As Short() = New Short() {1001, 1040, 1003, 1000, 1000}
    Public xDPol As Object() = New Object() {regAPP, 0, "", "0", "nZ=1;muros=0"}
    Public xTBlo As Short() = New Short() {1001, 1040, 1003, 1000, 1000}
    Public xDBlo As Object() = New Object() {regAPP, 0, "", "0", "codPom=;nombre="}
    Public xBase5 As Object() = New Object() {"", "", "", "", ""}
    Public xBase19 As Object() = New Object() {"", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""}

    Public Sub XNuevo(ByRef objA As AcadObject)
        '' Pone los XData para SERICAD. Solo si un objeto no tiene XData.
        '' Tipo de objeto al que vamos a poner los datos
        If TypeOf objA Is AcadLWPolyline Then
            XPonTodo(objA, xTPol, xDPol)
        ElseIf TypeOf objA Is AcadBlockReference Then
            XPonTodo(objA, xTBlo, xDBlo)
        End If
    End Sub

    Public Sub XBorrar(ByVal objA As AcadObject)
        Dim DataType(0) As Short
        Dim Data(0) As Object
        DataType(0) = 1001 : Data(0) = regAPP
        If TypeOf objA Is AcadLWPolyline Then
            objA = CType(objA, AcadLWPolyline)
        ElseIf TypeOf objA Is AcadBlockReference Then
            objA = CType(objA, AcadBlockReference)
        End If
        objA.SetXData(DataType, Data)
        objA.Update()
    End Sub

    Public Sub XLeeMensaje(ByVal obj As AcadObject)
        '' Leer xdata del objeto
        Dim xdatos As Object = Nothing
        Dim xtipos As Object = Nothing
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
        XLeeTiposDatos = resul
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
        XLeeDatos = xdatos
    End Function

    Public Function XLeeDato(ByVal objA As AcadObject, ByVal queDato As xT, Optional ByVal app As String = "") As Object     'Devuelve sólo la zona de SERICAD
        Dim resul As Object = Nothing
        If app = "" Then app = regAPP
        Dim xtipos() As Short = Nothing
        Dim xdatos() As Object = Nothing
        objA.GetXData(regAPP, xtipos, xdatos)
        '' Si el objeto no tiene XData
        If xdatos Is Nothing Then
            XNuevo(objA)
            objA.GetXData(app, xtipos, xdatos)
        End If

        If queDato > UBound(xdatos) Then
            resul = ""
        Else
            resul = xdatos(queDato)
        End If
        XLeeDato = resul
    End Function

    Public Function XLeeDatoValores(ByVal objA As AcadObject, ByVal tipo As xValor, Optional ByVal app As String = "") As String
        Dim resul As String = ""
        'XDataBorrar(objA)
        If app = "" Then app = regAPP
        Dim xtipos() As Short = Nothing
        Dim xdatos() As Object = Nothing
        objA.GetXData(app, xtipos, xdatos)
        '' Si el objeto no tiene XData
        If xdatos Is Nothing Then
            XNuevo(objA)
            objA.GetXData(app, xtipos, xdatos)
        End If

        '' Por si no tuviese el dato de xdatos(4) que corresponde
        '' a cadenas (dato1=valor1;dato2=valor2)
        If UBound(xdatos) < 4 Then
            XLeeDatoValores = resul
            Exit Function
        End If
        '' Por si los valores no son cadenas con formato predefinido.
        Dim Valores As String = xdatos(4).ToString
        If Valores.Contains("=") = False Or Valores.Contains(";") = False Then
            XLeeDatoValores = resul
            Exit Function
        End If
        '' Todas las cadenas entre ;
        Dim cadenas() As String = Valores.Split(";")
        '' Por si no tenemos la cadena nº que indica tipo
        If UBound(cadenas) < tipo Then
            XLeeDatoValores = resul
            Exit Function
        End If
        '' Cadena simple que indica tipo (nombre=valor)
        Dim cadenaDato As String = cadenas(tipo)
        '' Si no tiene =
        If cadenaDato.Contains("=") = False Then
            XLeeDatoValores = resul
            Exit Function
        End If
        '' Dato final (0) nombre / (1) dato
        Dim final() As String = cadenaDato.Split("=")
        resul = final(1).ToString

        XLeeDatoValores = resul
        Exit Function
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

    Public Sub XPonDato(ByVal objA As AcadObject, ByVal tipo As xT, ByVal dato As Object, Optional ByVal app As String = regAPP)
        'XDataBorrar(objA)
        If app = "" Then app = regAPP
        '' Si es la altura lo que vamos a cambiar. Lo ponemos en formato 3,00
        If tipo = 3 Then dato = FormatNumber(dato, 2, TriState.True)
        Dim xtipos() As Short = Nothing
        Dim xdatos() As Object = Nothing
        objA.GetXData(app, xtipos, xdatos)

        '' Si el objeto no tiene XData
        If xdatos Is Nothing Then
            XNuevo(objA)
            objA.GetXData(app, xtipos, xdatos)
        End If

        If tipo > UBound(xdatos) Then
            Exit Sub
        Else
            xdatos(tipo) = dato
        End If

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


    Public Sub XPonDatosArr(ByVal objA As AcadObject, ByVal datos() As Object, Optional ByVal app As String = "")
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

    Public Sub XPonDatoValores(ByVal objA As AcadObject, ByVal tipo As xValor, ByVal dato As Object, Optional ByVal app As String = "")
        'XDataBorrar(objA)
        If app = "" Then app = regAPP
        Dim xtipos() As Short = Nothing
        Dim xdatos() As Object = Nothing
        objA.GetXData(app, xtipos, xdatos)
        '' Si el objeto no tiene XData
        If xdatos Is Nothing Then
            XNuevo(objA)
            objA.GetXData(app, xtipos, xdatos)
        End If
        '' Por si no tuviese el dato de xdatos(4) que corresponde
        '' a cadenas (dato1=valor1;dato2=valor2)
        If UBound(xdatos) < 4 Then Exit Sub 'ReDim Preserve xdatos(4)
        '' Por si los valores no son cadenas con formato predefinido.
        Dim Valores As String = xdatos(4).ToString
        If Valores.Contains("=") = False Or Valores.Contains(";") = False Then Exit Sub
        '' Todas las cadenas entre ;
        Dim cadenas() As String = Valores.Split(";")
        '' Por si no tenemos la cadena nº que indica tipo
        If UBound(cadenas) < tipo Then Exit Sub
        '' Cadena simple que indica tipo (nombre=valor)
        Dim cadenaDato As String = cadenas(tipo)
        '' Si no tiene =
        If cadenaDato.Contains("=") = False Then Exit Sub
        '' Dato final (0) nombre / (1) dato
        Dim final() As String = cadenaDato.Split("=")
        Dim resultado As String = final(0) & "=" & dato.ToString

        Valores = Valores.Replace(cadenaDato, resultado)
        xdatos(4) = Valores

        '' Poner xdata al objeto
        Call objA.SetXData(xtipos, xdatos)
        '' Tipo de objeto al que vamos a poner los datos
        If TypeOf objA Is AcadLWPolyline Then
            objA = CType(objA, AcadLWPolyline)
        ElseIf TypeOf objA Is AcadBlockReference Then
            objA = CType(objA, AcadBlockReference)
        End If
        objA.Update()
    End Sub

    Public Sub XPonTodoValoresCadena(ByVal objA As AcadObject, ByVal dato As Object, Optional ByVal app As String = "")
        'XDataBorrar(objA)
        If app = "" Then app = regAPP
        Dim xtipos() As Short = Nothing
        Dim xdatos() As Object = Nothing
        objA.GetXData(app, xtipos, xdatos)
        '' Si el objeto no tiene XData
        If xdatos Is Nothing Then
            XNuevo(objA)
            objA.GetXData(app, xtipos, xdatos)
        End If
        '' Por si no tuviese el dato de xdatos(4) que corresponde
        '' a cadenas (dato1=valor1;dato2=valor2)
        If UBound(xdatos) < 4 Then Exit Sub 'ReDim Preserve xdatos(4)
        '' Por si los valores no son cadenas con formato predefinido.
        xdatos(4) = dato

        '' Poner xdata al objeto
        Call objA.SetXData(xtipos, xdatos)
        '' Tipo de objeto al que vamos a poner los datos
        If TypeOf objA Is AcadLWPolyline Then
            objA = CType(objA, AcadLWPolyline)
        ElseIf TypeOf objA Is AcadBlockReference Then
            objA = CType(objA, AcadBlockReference)
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
            XEsApp = resultado
            Exit Function
        End If

        If xdatos(xT.APLICACION) = regAPP Then
            resultado = True
            XEsApp = resultado
            Exit Function
        End If
        Return resultado
    End Function

    Public Enum xT As Integer
        APLICACION = 0
        OBJECTID = 1
        CAPA = 2
        ALTURA = 3
        VALORES = 4
        mat1 = 5
        x1 = 6
        y1 = 7
        mat2 = 8
        x2 = 9
        y2 = 10
        mat3 = 11
        x3 = 12
        y3 = 13
        mat4 = 14
        x4 = 15
        y4 = 16
        mat5 = 17
        x5 = 18
        y5 = 19
    End Enum

    Public Enum xValor As Integer
        nZona = 0
        muros = 1
        codPom = 0
        nombre = 1
        area = 0
        dato1 = 1
        dato2 = 2
        dato3 = 3
    End Enum

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
End Class
