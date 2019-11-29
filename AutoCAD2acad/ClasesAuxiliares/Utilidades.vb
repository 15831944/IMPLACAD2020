Option Compare Text
Imports Microsoft.Win32
Imports System.Runtime.InteropServices
Imports System.Data
Imports System.Drawing
Imports System.Text
Imports System.IO

'' Agregar Referencias NET: "System.Speech" y COM: "Microsoft Speech Object Library"
Module Utilidades
    '' ****************************************
    Public dirApp As String = My.Application.Info.DirectoryPath & "\"
    Public dirAppSin As String = My.Application.Info.DirectoryPath
    Public recursos As String = dirApp & "RECURSOS\"
    Public nApp As String = My.Application.Info.AssemblyName
    Public nIni As String = dirApp & nApp & ".ini"
    Public CuantoTiempo As Stopwatch
    Public Const nFijo As String = "2aCAD"
    '' ***** COLECCIONES *****
    Public colFicherosProc As Hashtable = Nothing   ' Coleccion de FullFilename/Procesado (True/False)
    '' ****************************************
    Public contar As Integer
    Public WithEvents oTt As Windows.Forms.ToolTip
    Public PermisosFi As System.Security.Permissions.FileIOPermission
    Public Licencia As String = ""
    Public miMaquina As String = ""
    Public miMaquina1 As String = ""
    Public FechaHoy As Date = Nothing
    '' Variable principal para saber si el usuario conectado tiene o no soporte CAD y licencia para la aplicaciÛn, sin conexiÛn a internet.
    Public TieneSoporte As Boolean = True   ' Para que no haga falta conectar con BD. False
    Public esComercial As Boolean = False
    'ALBERTOPORTATIL =  (Complejo="k)rRl!ky|Pz(ydvêyFxE~Óg≥~Ho®vﬁ"    /   Simple="FOGHWWTSTUYDYLQ")
    '31/12/2011 =       (Complejo=":ª7Ä9ı6u;´5È<ô6‰Cß?π"    /   Simple="8444727364"
    'LICENCIA =         (Complejo="v˘o´mzk”x©iÅsΩgÍ"    /   Simple="QLHHSFND")
    'QLHHSFND=          (Complejo=""    /   Simple="FOGHWWTSTUYDYLQ∑8444727364")
    '' Argumentos de la linea de comandos: 2aCAD Bgik8$l. (sÛlo para modo Debug)
    Public Const clave1 As String = "<°gÄMOGßN~"
    Public Const superClave As String = "LWmEsÙqlBC*˚vW4É"
    Public superAcceso As Boolean = False
    Public ConfReg As System.Globalization.CultureInfo
    Public carExp As ArrayList = New ArrayList(New String() {"_", "*", "\", "^"})
    'Public colFicherosProc As Hashtable

    Public Enum EstructuraDirs
        DirPrincipal = FileIO.SearchOption.SearchTopLevelOnly
        DirTodos = FileIO.SearchOption.SearchAllSubDirectories
    End Enum

    Public Function DameNombreEquipo() As String
        Return My.Computer.Name
    End Function

    ''' <summary>
    ''' Si le damos una cadena completa (unidad:\directorio\fichero.extension) nos devuelve la parte que le indiquemos.
    ''' </summary>
    ''' <param name="queCamino">Cadena completa con el camino a procesar DIR+FICHERO+EXT</param>
    ''' <param name="queParte">Que queremos que nos devuelva</param>
    ''' <param name="queExtension">"" o extensiÛn (Ej: ".bak"), si queremos cambiarla</param>
    ''' <returns>Retorna la cadena de texto con la opciÛn indicada</returns>
    ''' <remarks></remarks>
    Public Function DameParteCamino(ByVal queCamino As String, Optional ByVal queParte As ParteCamino = 0, Optional ByVal queExtension As String = "") As String
        Dim resultado As String = ""

        Select Case queParte
            Case 0  'ParteCamino.SoloCambiaExtension (dwg) Sin punto
                If queExtension <> "" And IO.Path.HasExtension(queCamino) Then
                    queCamino = IO.Path.ChangeExtension(queCamino, queExtension)
                End If
                resultado = queCamino
            Case 1  'ParteCamino.CaminoSinFichero
                resultado = IO.Path.GetDirectoryName(queCamino)
            Case 2  'ParteCamino.CaminoSinFicheroBarra
                resultado = IO.Path.GetDirectoryName(queCamino) & "\"
            Case 3  'ParteCamino.CaminoConFicheroSinExtension
                resultado = IO.Path.ChangeExtension(queCamino, Nothing)
            Case 4  'ParteCamino.CaminoConFicheroSinExtension
                resultado = IO.Path.ChangeExtension(queCamino, Nothing) & "\"
            Case 5  'ParteCamino.SoloFicheroConExtension
                resultado = IO.Path.GetFileName(queCamino)
            Case 6  'ParteCamino.SoloFicheroSinExtension
                resultado = IO.Path.GetFileNameWithoutExtension(queCamino)
            Case 7  'ParteCamino.SoloExtension
                resultado = IO.Path.GetExtension(queCamino)
            Case 8  'ParteCamino.SoloRaiz
                resultado = IO.Path.GetPathRoot(queCamino)
            Case 9  'ParteCamino.SoloNombreDirectorio
                Dim trozos() As String = queCamino.Split("\")
                resultado = trozos(trozos.GetUpperBound(0) - 1)
            Case 10  'ParteCamino.PenultimoDirectorioSinBarra
                Dim trozos() As String = queCamino.Split("\")
                If trozos.GetUpperBound(0) > 2 Then
                    Dim final(trozos.GetUpperBound(0) - 1) As String
                    Array.Copy(trozos, final, trozos.GetUpperBound(0) - 1)
                    resultado = String.Join("\", final)
                ElseIf trozos.GetUpperBound(0) > 1 Then
                    resultado = trozos(0)
                Else
                    resultado = "C:"
                End If
                If resultado.EndsWith("\") Then resultado = Mid(resultado, 1, resultado.Length - 1)
            Case 11  'ParteCamino.PenultimoDirectorioConBarra
                Dim trozos() As String = queCamino.Split("\")
                If trozos.GetUpperBound(0) > 1 Then
                    Dim final(trozos.GetUpperBound(0) - 2) As String
                    Array.Copy(trozos, final, trozos.GetUpperBound(0) - 1)
                    resultado = String.Join("\", final)
                ElseIf trozos.GetUpperBound(0) > 1 Then
                    resultado = trozos(0)
                Else
                    resultado = "C:"
                End If
                If resultado.EndsWith("\") = False Then resultado &= "\"
            Case 12  'ParteCamino.AntePenultimoDirectorioSinBarra
                Dim trozos() As String = queCamino.Split("\")
                If trozos.GetUpperBound(0) > 2 Then
                    Dim final(trozos.GetUpperBound(0) - 2) As String
                    Array.Copy(trozos, final, trozos.GetUpperBound(0) - 2)
                    resultado = String.Join("\", final)
                ElseIf trozos.GetUpperBound(0) > 1 Then
                    resultado = trozos(0)
                Else
                    resultado = "C:"
                End If
                If resultado.EndsWith("\") Then resultado = Mid(resultado, 1, resultado.Length - 1)
            Case 13  'ParteCamino.AntePenultimoDirectorioConBarra
                Dim trozos() As String = queCamino.Split("\")
                If trozos.GetUpperBound(0) > 2 Then
                    Dim final(trozos.GetUpperBound(0) - 2) As String
                    Array.Copy(trozos, final, trozos.GetUpperBound(0) - 2)
                    resultado = String.Join("\", final)
                ElseIf trozos.GetUpperBound(0) > 1 Then
                    resultado = trozos(0)
                Else
                    resultado = "C:"
                End If
                If resultado.EndsWith("\") = False Then resultado &= "\"
        End Select
        DameParteCamino = resultado
        Exit Function
    End Function

    Public Enum ParteCamino
        SoloCambiaExtension = 0
        CaminoSinFichero = 1
        CaminoSinFicheroBarra = 2
        CaminoConFicheroSinExtension = 3
        CaminoConFicheroSinExtensionBarra = 4
        SoloFicheroConExtension = 5
        SoloFicheroSinExtension = 6
        SoloExtension = 7
        SoloRaiz = 8
        SoloNombreDirectorio = 9
        PenultimoDirectorioSinBarra = 10
        PenultimoDirectorioConBarra = 11
        AntePenultimoDirectorioSinBarra = 10
        AntePenultimoDirectorioConBarra = 11
    End Enum
    ''' <summary>
    ''' Buscar todos los fichero dentro en directorio (opciÛn subdirectorios y m·scara con extensiÛn) 
    ''' </summary>
    ''' <param name="strD1">Directorio donde se va a buscar</param>
    ''' <param name="queEstructura">Tipo de b˙squeda (Dir indicado solo o TambiÈn SubDir)</param>
    ''' <param name="extension">ExtensiÛn opcional a buscar (Por defecto Todos *.* si no ponemos nada)</param>
    ''' <returns>Devuelve un colecciÛn con todos los nombres completos encontrado, que habr· que recorrer</returns>
    ''' <remarks></remarks>
    Public Function BuscaFicheros(ByVal strD1 As Object, ByVal queEstructura As EstructuraDirs, Optional ByVal extension As String = "*.*") As Collection
        Dim colLista As New Collection
        Dim strD As String = CType(strD1, String)

        For Each foundF As String In My.Computer.FileSystem.GetFiles( _
        strD, queEstructura, extension)
            'If todo = True Then
            'colLista.Add(foundF)
            'contar += 1
            'Else
            ' SÛlo fichero de Inventor
            'If (foundF.EndsWith(".iam") Or foundF.EndsWith(".ipt") Or foundF.EndsWith(".idw") Or foundF.EndsWith(".ipn")) And Not (foundF.Contains("OldVersions")) Then
            If colLista.Contains(foundF) = False Then colLista.Add(foundF)
            contar += 1
            'End If
            'End If
        Next
        'MessageBox.Show(colLista.Count & " / " & colLista.Item(colLista.Count).ToString)
        BuscaFicheros = colLista
    End Function
    ''' <summary>
    ''' Buscar todos los fichero dentro en directorio (opciÛn subdirectorios y m·scara con extensiÛn)
    ''' Llena el Hashtable "colListaProc" KEY=fullPath, VALUE=False (No procesado aun) 
    ''' </summary>
    ''' <param name="strD1">Directorio donde se va a buscar</param>
    ''' <param name="queEstructura">Tipo de b˙squeda (Dir indicado solo o TambiÈn SubDir)</param>
    ''' <param name="extension">ExtensiÛn opcional a buscar (Por defecto Todos *.* si no ponemos nada)</param>
    ''' <remarks></remarks>
    Public Sub BuscaFicherosProc(ByVal strD1 As Object, ByVal queEstructura As EstructuraDirs, exclusiones() As String, Optional ByVal extension As String = "*.*")
        If colFicherosProc Is Nothing Then colFicherosProc = New Hashtable
        Dim strD As String = CType(strD1, String)

        For Each foundF As String In My.Computer.FileSystem.GetFiles( _
        strD, queEstructura, extension)
            'If todo = True Then
            'colLista.Add(foundF)
            'contar += 1
            'Else
            ' SÛlo fichero de Inventor
            'If (foundF.EndsWith(".iam") Or foundF.EndsWith(".ipt") Or foundF.EndsWith(".idw") Or foundF.EndsWith(".ipn")) And Not (foundF.Contains("OldVersions")) Then
            Dim excluir As Boolean = False
            For Each exclu As String In exclusiones
                If foundF.Contains(exclu) = True Then excluir = True
            Next
            If colFicherosProc.Contains(foundF) = False AndAlso excluir = False Then colFicherosProc.Add(foundF, False)
            contar += 1
            'End If
            'End If
            System.Windows.Forms.Application.DoEvents()
        Next
        'MessageBox.Show(colLista.Count & " / " & colLista.Item(colLista.Count).ToString)
    End Sub
    Public Function DataRowCambia(ByRef queDR As DataRow) As DataRow
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
        DataRowCambia = resultado
        Exit Function
    End Function

    Public Sub Retardo(ByVal segundos As Integer)
        Const NSPerSecond As Long = 10000000
        Dim ahora As Long = Date.Now.Ticks
        Debug.Print(Date.Now.Ticks)
        Do
            ' No hacemos nada
            'My.Application.DoEvents()
        Loop While Date.Now.Ticks < ahora + (segundos * NSPerSecond)
        Debug.Print(Date.Now.Ticks)
    End Sub

    Public Function HashtableCambia(ByRef queHorigen As Hashtable) As Hashtable
        Dim resultado As Hashtable = New Hashtable
        ' Se crea un Array de elementos DictionaryEntry con el tamaÒo del Hashtable a recorrer
        Dim arrayCopia(queHorigen.Count - 1) As DictionaryEntry

        ' Se copia el contenido de la tabla sobre el array que acabamos de crear()
        queHorigen.CopyTo(arrayCopia, 0)

        ' Se borra completamente la tabla inicial
        queHorigen.Clear()

        '' Pasamos a centimetros todos los valores que vengan en milimetros. Par·metros que empiezan por "di_"
        ' Se recorre el arrayCopia con un bucle For... Next a la vez que se aÒaden los nuevos elementos modificados a la tabla
        For i As Integer = 0 To arrayCopia.Length - 1
            resultado.Add(arrayCopia(i).Key, arrayCopia(i).Value)
        Next
        HashtableCambia = resultado
        Exit Function
    End Function

    Public Enum D2aCAD
        Ninguna = 0
        Bilbao = 1
        CastellÛn = 2
        Madrid = 3
        Valencia = 4
        Zaragoza = 5
    End Enum
    '' Direcciones de 2aCAD Delegaciones (8 lÌneas m·ximo)
    Public Function Direcciones2aCAD(ByVal queDire As D2aCAD) As String
        Dim resultado As String = ""
        Select Case queDire
            Case D2aCAD.Bilbao
                resultado &= "Bilbao"
                resultado &= "C/Uribitarte, 22 - 5∫" & vbCrLf
                resultado &= "48001 - Bilbao" & vbCrLf & vbCrLf
                resultado &= "TelÈfono: 94 435 53 94" & vbCrLf
                resultado &= "Email: bilbao@2acad.com" & vbCrLf
            Case D2aCAD.CastellÛn
                resultado &= "CastellÛn" & vbCrLf
                resultado &= "C/ MarÌa Teresa Gonz·lez, 26 Entlo." & vbCrLf
                resultado &= "12005 - CastellÛn" & vbCrLf & vbCrLf
                resultado &= "TelÈfono: 964 72 48 70" & vbCrLf
                resultado &= "Fax: 964 72 48 71" & vbCrLf
                resultado &= "Email: castellon@2acad.com" & vbCrLf
            Case D2aCAD.Madrid
                resultado &= "Madrid" & vbCrLf
                resultado &= "Rodriguez San Pedro, 13 Planta 2™" & vbCrLf
                resultado &= "28015 - Madrid" & vbCrLf & vbCrLf
                resultado &= "TelÈfono: 91 564 81 94" & vbCrLf
                resultado &= "Email: madrid@2acad.com" & vbCrLf
            Case D2aCAD.Valencia
                resultado &= "Valencia" & vbCrLf
                resultado &= "Ronda Narciso Monturiol, 6" & vbCrLf
                resultado &= "Parque TecnolÛgico de Paterna" & vbCrLf
                resultado &= "46980 - Paterna (Valencia)" & vbCrLf & vbCrLf
                resultado &= "TelÈfono: 96 313 40 35" & vbCrLf
                resultado &= "Fax: 96 383 78 97" & vbCrLf
                resultado &= "Email: valencia@2acad.com" & vbCrLf
            Case D2aCAD.Zaragoza
                resultado &= "Zaragoza" & vbCrLf
                resultado &= "C/Bari, 57" & vbCrLf
                resultado &= "Centro TecnolÛgico TICXXI PLA-ZA" & vbCrLf
                resultado &= "50197 - Zaragoza" & vbCrLf & vbCrLf
                resultado &= "TelÈfono: 976 45 81 45/51 (CAD directo)" & vbCrLf
                resultado &= "Fax: 976 75 20 15" & vbCrLf
                resultado &= "Email: zaragoza@2acad.com" & vbCrLf
        End Select
        Return resultado
    End Function

    Public Function PermisosTodo() As String
        Dim resultado As String = ""
        PermisosFi = New System.Security.Permissions.FileIOPermission(Security.Permissions.PermissionState.Unrestricted)
        PermisosFi.AllLocalFiles = Security.Permissions.FileIOPermissionAccess.AllAccess
        Try
            PermisosFi.Demand()
            'Permisos.SetPathList(Security.Permissions.FileIOPermissionAccess.AllAccess, nDir)
        Catch s As System.Security.SecurityException
            Console.WriteLine(s.Message)
            MsgBox(s.Message)
        End Try
        resultado = PermisosFi.ToString
        Return resultado
    End Function
    ''
    ''
    Public Sub SoloLecturaDirectorioFichero(queDir As String, queFi As String, Optional mascara As String = "*.*", Optional recursivo As Boolean = True, Optional sololectura As Boolean = False)
        ''
        If queDir = "" Then
            ''
            If queFi = "" OrElse IO.File.Exists(queFi) = False Then Exit Sub
            ''
            Dim fInfo As New System.IO.FileInfo(queFi)
            fInfo.IsReadOnly = sololectura
        ElseIf queDir <> "" AndAlso IO.Directory.Exists(queDir) Then
            For Each queF As String In IO.Directory.GetFiles(queDir, mascara, IIf(recursivo, IO.SearchOption.AllDirectories, IO.SearchOption.TopDirectoryOnly))
                Dim fInfo As New System.IO.FileInfo(queF)
                fInfo.IsReadOnly = sololectura
            Next
        End If
    End Sub
''
    Public Sub BorraTodosFicheros(ByVal queDir As String, ByVal dirTambien As Boolean)
        If IO.Directory.Exists(queDir) = False Then
            IO.Directory.CreateDirectory(queDir)
            Exit Sub
        End If

        For Each fichero As String In IO.Directory.GetFiles(queDir, "*.*", IO.SearchOption.AllDirectories)
            IO.File.Delete(fichero)
        Next

        If dirTambien = True Then IO.Directory.Delete(queDir, True)
    End Sub

    Public Function DameDesarrollo(Optional ByVal todo As Boolean = False) As String
        If todo = True Then
            Return My.Application.Info.DirectoryPath & "\" & My.Application.Info.AssemblyName
        Else
            Return My.Application.Info.AssemblyName
        End If
    End Function


    Public Function DameLetrasInicialesSplit(ByVal cadena As String) As String
        Dim resultado As String = ""

        If cadena.Contains("∑") Then
            '' Si encontramos FU∑90
            Dim strcodArmado() As String = cadena.Split("∑")
            resultado = strcodArmado(0)  'FU
        Else
            '' Si encontramos FU90
            For x As Integer = 0 To cadena.Length - 1
                If IsNumeric(cadena.Chars(x)) = False Then
                    resultado &= cadena.Chars(x).ToString
                Else
                    Exit For
                End If
            Next
        End If

        Return resultado
        Exit Function
    End Function

    Public Function DameLetrasIniciales(ByVal cadena As String) As String
        Dim resultado As String = ""
        For x As Integer = 0 To cadena.Length - 1
            If IsNumeric(cadena.Chars(x)) = False Then
                resultado &= cadena.Chars(x).ToString
            Else
                Exit For
            End If
        Next
        Return resultado
    End Function


    Public Function DameNumerosIniciales(ByVal cadena As String) As String
        Dim resultado As String = ""
        For x As Integer = 0 To cadena.Length - 1
            If IsNumeric(cadena.Chars(x)) = True Then
                resultado &= cadena.Chars(x).ToString
            Else
                Exit For
            End If
        Next
        Return resultado
    End Function

    Public Sub ConfiguracionRegionalEspaÒa()
        Try
            '' Poner la configuraciÛn regional en EspaÒa. Para las conversiones de . decimal en ,
            If ConfReg Is Nothing Then ConfReg = New System.Globalization.CultureInfo("es-ES")
            System.Threading.Thread.CurrentThread.CurrentCulture = ConfReg
        Catch ex As Exception
            '' No hacemos nada.
        End Try
    End Sub

    Public Function ArrayLimpiaOldVersions(ByRef queArr As ArrayList) As ArrayList
        Dim queArrNuevo As New ArrayList

        For Each queFi As String In queArr
            If queFi.ToLower.Contains("oldversions") = False Then queArrNuevo.Add(queFi)
        Next
        Return queArrNuevo
    End Function

    Public Function DistanciaDame(ByVal x1 As Double, ByVal y1 As Double, ByVal z1 As Double, ByVal x2 As Double, ByVal y2 As Double, ByVal z2 As Double) As Double
        Return (((x2 - x1) ^ 2) + ((y2 - y1) ^ 2) + ((z2 - z1) ^ 2)) ^ 0.5
        Exit Function
    End Function

    Public Sub PonLogNombre(ByVal quetexto As String, Optional ByVal borrar As Boolean = False, Optional queF As String = "")
        If queF = "" Then queF = dirApp & nApp & "_log.txt"
        If borrar = True And IO.File.Exists(queF) Then IO.File.Delete(queF)
        If quetexto.EndsWith(vbCrLf) = False Then quetexto &= vbCrLf
        IO.File.AppendAllText(queF, Date.Now & vbTab & quetexto)
    End Sub

    Public Function GraRad(ByVal queGra As Double) As Double
        Return (queGra * Math.PI) / 180
    End Function

    Public Function RadGra(ByVal queRad As Double) As Double
        Return (queRad * 180) / Math.PI
    End Function

    ''' <summary>
    ''' Ampliar una matriz de String. SÛlo una indicando su nombre y el valor a aÒadir (Tiene que ser un string)
    ''' </summary>
    ''' <param name="queMatriz"></param>
    ''' <param name="queDato"></param>
    ''' <remarks></remarks>
    Public Sub MatrizPonDatoString(ByRef queMatriz() As String, ByVal queDato As String)
        ReDim Preserve queMatriz(queMatriz.GetUpperBound(0) + 1)
        queMatriz(queMatriz.GetUpperBound(0)) = queDato
    End Sub

    ''' <summary>
    ''' Ampliar una matriz de Object. SÛlo una indicando su nombre y el valor a aÒadir (Tiene que ser un Object, Clase o Similar)
    ''' </summary>
    ''' <param name="queMatriz"></param>
    ''' <param name="queDato"></param>
    ''' <remarks></remarks>
    Public Sub MatrizPonDatoObject(ByRef queMatriz() As Object, ByVal queDato As Object)
        ReDim Preserve queMatriz(queMatriz.GetUpperBound(0) + 1)
        queMatriz(queMatriz.GetUpperBound(0)) = queDato
    End Sub
    ''
    '' Le pasamos una cadena que contiene letras al inicio y n˙meros.
    '' Nos devolver· los n˙meros, rellenos con 0 a la izquierda hasta nTotal car·cteres
    '' Ejemplo: Le pasamos NumeroDameRelleno("V024", "0", 4) y nos devolver· "0024"
    Public Function NumeroDameRelleno(queCadena As String, CharRelleno As String, nTotal As Integer) As String
        Dim resultado As String = ""
        For x As Integer = 0 To queCadena.Length - 1
            If IsNumeric(queCadena(x)) = True Then
                resultado = queCadena.Substring(x)
                Exit For
            End If
        Next

        If resultado <> "" And resultado.Length < nTotal Then
            resultado = resultado.PadLeft(nTotal, CharRelleno)
        End If
        Return resultado
    End Function


    '' Le pasamos una cadena y nos devuelve sÛlo el n˙mero, sin las unidades
    '' Ej.: Le pasamos 1000 cm y nos devuelve 1000
    Public Function QuitaTextoUnidades(queCadena As String) As String
        If queCadena = "" Then
            Return queCadena
            Exit Function
        ElseIf IsNumeric(queCadena) = True Then
            Return queCadena
            Exit Function
        ElseIf IsNumeric(queCadena.Chars(0)) = False Then
            Return queCadena
            Exit Function
        ElseIf queCadena.Contains(" ") = True Then
            Dim valores() As String = queCadena.Split(" ")
            If IsNumeric(valores(0)) = True Then
                Return valores(0)
                Exit Function
            Else
                Return queCadena
                Exit Function
            End If
        End If

        Dim resultado As String = queCadena
        For x As Integer = 0 To queCadena.Length - 1
            If IsNumeric(queCadena(x)) Then
                Continue For
            ElseIf queCadena(x) = "" Then
                Continue For
            Else
                resultado = queCadena.Substring(0, x - 1)
                Exit For
            End If
        Next

        Return resultado
    End Function

    ''' <summary>
    ''' Nos da un cÛdigo ˙nico (0f8fad5b-d9cb-469f-a165-70867728950e) de 32 car·cteres m·ximo (sin los "-")
    ''' O del n˙mero de car·cteres (menor de 32) que le indiquemos. Si nChar>32 devuelve 32.
    ''' </summary>
    ''' <param name="nChar">N∫ de car·cteres de la cadena a devolver (32 si no lo indicamos)</param>
    ''' <returns>Cadena ˙nica de 32 car·cteres o nChar (siempre menor que 32) que le indiquemos</returns>
    ''' <remarks></remarks>
    Public Function DameIdUnico(Optional nChar As Integer = 32) As String
        'System.GUID.NewGuid nos da un codigo ˙nico (8-4-4-4-12) y despues aplicamos esta formula para formatearlo para ADAgeS
        'left(replace(newid(),'-',''),16)
        If nChar > 32 Then nChar = 32
        Return Strings.Left(Replace(System.Guid.NewGuid.ToString, "-", ""), nChar)
    End Function
    ''
    Public Function ConvertImageToByteArray(ByVal imageIn As System.Drawing.Image) As Byte()
        Dim ms As New IO.MemoryStream()
        imageIn.Save(ms, System.Drawing.Imaging.ImageFormat.Png)
        Return ms.ToArray()
    End Function
    ''
    Public Function ConvertByteArrayToImage(ByVal byteArrayIn As Byte()) As Image
        Dim ms As New IO.MemoryStream(byteArrayIn)
        Dim returnImage As Image = Image.FromStream(ms)
        Return returnImage
    End Function
    ''
    Public Function EsRutaValida(ByVal queRuta As String) As Boolean
        Dim resultado As Boolean = True
        '' Comprobar primero los car·cteres invalidos del directorio.
        For Each oChar As Char In IO.Path.GetInvalidPathChars
            If queRuta.Contains(oChar) Then
                resultado = False
                Exit For
            End If
        Next
        '' Comprobar la longitud de la cadena.
        If queRuta.Length > 255 Then resultado = False
        ''
        Return resultado
    End Function

    Public Function CarpetaFicheroRemplazaCaracteresIncorrectos(ByVal queTexto As String) As String
        Dim resultado As String = queTexto

        'Dim malos() As Char = New Char() _
        '{"/", "\", "<", ">", "[", "]", "{", "}", "ø", "?", "°", "!", ":", "|", "*"}   ', ",", "."}
        For Each queChar As Char In IO.Path.GetInvalidFileNameChars ' malos
            If queTexto.Contains(queChar) Then
                resultado = resultado.Replace(queChar, "_")
            End If
        Next

        Return resultado
    End Function
    ''
    Public Function CarpetaFicheroDameSoloNombreCorrecto(ByVal nCarpeta As String) As String
        Dim resultado As String = nCarpeta
        ''
        resultado = CarpetaFicheroRemplazaCaracteresIncorrectos(resultado)
        ''
Repite:
        If resultado.EndsWith(".") Or
                resultado.EndsWith(",") Or
                resultado.EndsWith("'") Then
            resultado = resultado.Substring(0, resultado.Length - 1)
            GoTo Repite
        End If
        ''
        Return resultado
    End Function
    Public Function CarpetaFicheroTieneCaracteresIncorrectos(ByVal queTexto As String) As Boolean
        Dim resultado As Boolean = False

        ' Dim malos() As Char = New Char() _
        '{"/", "\", "<", ">", "[", "]", "{", "}", "ø", "?", "°", "!", ":", "|", "*"}   ', ","}
        For Each queChar As Char In IO.Path.GetInvalidFileNameChars ' malos
            If queTexto.Contains(queChar) Then
                resultado = True
                Exit For
            End If
        Next
        Return resultado
    End Function

    Public Function AleatorioUno(minimo As Integer, maximo As Integer) As Integer
        Randomize()
        'randomValue = CInt(Math.Floor((upperbound - lowerbound + 1) * Rnd())) + lowerbound
        Return CInt(Math.Floor((maximo - minimo + 1) * Rnd())) + minimo
    End Function

    Public Function AleatorioArrayListOrdenado(minimo As Integer, maximo As Integer, cuantos As Integer, Optional ordenado As Boolean = True) As ArrayList
        Randomize()
        'randomValue = CInt(Math.Floor((upperbound - lowerbound + 1) * Rnd())) + lowerbound
        Dim arrNumeros As New ArrayList
        For x = 1 To cuantos
            Dim numero As Integer = AleatorioUno(minimo, maximo)
            If arrNumeros.Contains(numero) Then
                x -= 1
                Continue For
            Else
                arrNumeros.Add(numero)
                'txtDatos.Text &= numero.ToString & vbCrLf
            End If
        Next
        If ordenado = True Then arrNumeros.Sort()
        Return arrNumeros
    End Function
	''
    Public Sub CierraProceso(queProceso As String)
        'Dim proceso As System.Diagnostics.Process
        ' '' Cargar el proceso.
        'Try
        '    proceso = Process.GetProcessesByName(queProceso)
        '    proceso.Kill()
        'Catch ex As Exception
        '    proceso = Nothing
        '    Exit Sub
        'End Try
        ' ''
        ' '' Eliminar todas las instancias del proceso y subprocesos.
        'Try
        '    For Each proc As System.Diagnostics.Process In proceso
        '        proc.CloseMainWindow()
        '        proc.WaitForExit()
        '    Next
        'Catch ex As System.NullReferenceException
        '    'MessageBox.Show("No hay instancias de " & queProceso & " en ejecuciÛn.")
        'End Try
        'proceso = Nothing
        ''
        ''
        Try
            Dim proc As System.Diagnostics.Process
            ''
            For Each proc In System.Diagnostics.Process.GetProcessesByName(queProceso)
                proc.Kill()
            Next
        Catch ex As Exception
            '    'MessageBox.Show("No hay instancias de " & queProceso & " en ejecuciÛn.")
        End Try
    End Sub
    '' Devolver· la misma carpeta (Si es local) O la carpeta de red (\\servidor\carpeta\, si es de red)
    Public Function GetUNCPath(ByVal sFilePath As String) As String
        Dim allDrives() As DriveInfo = DriveInfo.GetDrives()
        Dim d As DriveInfo
        Dim DriveType, Ctr As Integer
        Dim DriveLtr, UNCName As String
        Dim StrBldr As New StringBuilder

        If sFilePath.StartsWith("\\") Then Return sFilePath

        UNCName = Space(160)
        GetUNCPath = ""

        DriveLtr = sFilePath.Substring(0, 3)

        For Each d In allDrives
            If d.Name.ToUpper = DriveLtr.ToUpper Then
                DriveType = d.DriveType
                Exit For
            End If
        Next

        If DriveType = 4 Then
            Ctr = WNetGetConnection(sFilePath.Substring(0, 2), UNCName, UNCName.Length)

            If Ctr = 0 Then
                UNCName = UNCName.Trim
                For Ctr = 0 To UNCName.Length - 1
                    Dim SingleChar As Char = UNCName(Ctr)
                    Dim asciiValue As Integer = Asc(SingleChar)
                    If asciiValue > 0 Then
                        StrBldr.Append(SingleChar)
                    Else
                        Exit For
                    End If
                Next
                StrBldr.Append(sFilePath.Substring(2))
                GetUNCPath = StrBldr.ToString
            End If
        Else
            GetUNCPath = sFilePath
        End If
    End Function
    ''
    'Public Function Eval(ByVal Expresion As String) As Object
    '    Dim vbcp As New Microsoft.VisualBasic.VBCodeProvider
    '    Dim vbc As System.CodeDom.Compiler.ICodeCompiler = vbcp.CreateCompiler()
    '    Dim cpar As New System.CodeDom.Compiler.CompilerParameters
    '    Dim res As System.CodeDom.Compiler.CompilerResults

    '    cpar.GenerateExecutable = False ' Generar DLL
    '    cpar.GenerateInMemory = True ' Generar en memoria
    '    cpar.IncludeDebugInformation = True

    '    ' Agregar referencias
    '    cpar.ReferencedAssemblies.Add("Microsoft.VisualBasic.dll")

    '    ' Compilar
    '    res = vbc.CompileAssemblyFromSource(cpar, _
    '        "Imports Microsoft.VisualBasic" & vbCrLf & _
    '        "Namespace MiNamespace" & vbCrLf & _
    '        " Public Class MiClase" & vbCrLf & _
    '        "  Public Shared Function Eval() As Object " & vbCrLf & _
    '        "   Return " & Expresion & vbCrLf & _
    '        "  End Function" & vbCrLf & _
    '        " End Class" & vbCrLf & _
    '        "End Namespace")

    '    If res.Errors.Count = 0 Then
    '        ' Obtengo el Type de la clase recien compilada
    '        Dim miclase As System.Type
    '        miclase = res.CompiledAssembly.GetType("MiNamespace.MiClase")
    '        ' Obtengo el metodo Eval de la clase
    '        Dim funceval As System.Reflection.MethodInfo
    '        funceval = miclase.GetMethod("Eval")
    '        ' Ejecuto la funcion Eval recien creada
    '        Return funceval.Invoke(Nothing, Nothing)
    '    Else
    '        Return res.Errors(0).ErrorText
    '    End If
    'End Function
    ''
    Declare Function WNetGetConnection Lib "mpr.dll" Alias "WNetGetConnectionA" (ByVal lpszLocalName As String, _
           ByVal lpszRemoteName As String, ByRef cbRemoteName As Integer) As Integer
End Module
