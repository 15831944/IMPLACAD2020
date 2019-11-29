Imports System.IO
Imports System.Linq
Imports System.Xml
Imports System.Xml.Linq

Module modXml
    '' EJEMPLO CON LINQ
    'Dim xRoot As xElement = xDocument.Load("Nombre.xml").Root
    'Dim Puerto As Integer = xRoot.<Puerto>.Value
    'Dim Longitud As Integer = xRoot.<Longitud>.Value
    'Dim EOF As String = xRoot.<EOF>.Value
    ''
    '' Utilizando XMLTextReader (FICHERO: \offlineBDIdata\languages.xml)
    Public Function LeeXML_XMLTextReader(queFichero As String) As SortedList
        Dim resultado As New SortedList
        Dim m_xmlr As XmlTextReader
        'Creamos el TextReader
        m_xmlr = New XmlTextReader(queFichero)
        'Desabilitamos las lineas en blanco, 
        'ya no las necesitamos
        m_xmlr.WhitespaceHandling = WhitespaceHandling.None
        'Leemos el archivo y avanzamos al tag de usuarios
        m_xmlr.Read()
        'Leemos el tag usuarios
        m_xmlr.Read()
        'Creamos la secuancia que nos permite 
        'leer el archivo
        While Not m_xmlr.EOF
            'Avanzamos al siguiente tag
            m_xmlr.Read()
            'si no tenemos el elemento inicial 
            'debemos salir del ciclo
            If Not m_xmlr.IsStartElement() Then
                Exit While
            End If
            'Obtenemos el elemento codigo
            Dim mLenguajes = m_xmlr.GetAttribute("languages")
            'Read elements firstname and lastname
            m_xmlr.Read()
            'Obtenemos el elemento del Nombre del Usuario
            Dim mCodigo = m_xmlr.ReadElementString("languageCode")

            'Obtenemos el elemento del Apellido del Usuario
            Dim mDescripcion = m_xmlr.ReadElementString("description")
            'Escribimos el resultado en la consola, pero tambien podriamos utilizarlos donde deseemos    
            'Console.WriteLine("Lenguaje: " & mLenguajes _
            ' & " languageCode: " & mCodigo _
            ' & " description: " & mDescripcion)
            'Console.Write(vbCrLf)
            If resultado.ContainsKey(mCodigo) = False Then resultado.Add(mCodigo, mDescripcion)
        End While

        'Cerramos la lactura del archivo
        m_xmlr.Close()
        ''
        Return resultado
    End Function
    ''
    Public Function LeeXML_XMLDocument(queFichero As String,
                                       queNodes As String) As String
        Dim resultado As String = ""
        Try
            Dim documento As XmlDocument
            Dim xLista As XmlNodeList
            Dim xNode As XmlNode
            'Creamos el "Document"
            documento = New XmlDocument()
            'Cargamos el archivo
            documento.Load(queFichero)
            '' Obtenemos el nodo raiz del documento
            Dim raiz As XmlElement = documento.DocumentElement
            'Obtenemos la lista de los nodos "name"
            xLista = documento.SelectNodes(queNodes)    '"/languages/language")
            'Iniciamos el ciclo de lectura
            For Each xNode In xLista
                'Obtenemos el atributo del codigo
                'Dim mCodigo = m_node.Attributes.GetNamedItem("codigo").Value       '' Si language tuviese un atributo codigo=XXXX
                'Obtenemos el Elemento languageCode
                Dim mCodigo = xNode.ChildNodes.Item(0).InnerText
                'Obtenemos el Elemento description
                Dim mDescripcion = xNode.ChildNodes.Item(1).InnerText
                'Escribimos el resultado en la consola, pero tambien podriamos utilizarlos donde deseemos
                'Console.Write("Codigo usuario: " & mCodigo _
                '  & " Nombre: " & mNombre _
                '  & " Apellido: " & mApellido)
                'Console.Write(vbCrLf)
                resultado &= "Codigo = " & mCodigo & "  /  Descripcion = " & mDescripcion & vbCrLf
            Next
        Catch ex As Exception
            'Error trapping
            Debug.Print(ex.ToString())
        End Try
        ''
        Return resultado
    End Function
    ''
    Public Function BuscaNodo(InnerText As String) As XmlNode
        Dim resultado As XmlNode = Nothing

        Return resultado
    End Function
    ''

    ''
    ''
    Public Function EscribeBorraXMLNode_Linq(queFichero As String,
                                             elementoslistar As String,
                                             NodoBusco As String,
                                             DatoBusco As String,
                                             DatoValor As String,
                                             Optional AttNodo As String = "",
                                             Optional AttNodoValor As String = "",
                                             Optional AttHijo As String = "",
                                             Optional AttHijoValor As String = "",
                                             Optional borrarNodo As Boolean = False) As String
        Dim resultado As String = ""
        '' Cargamos el documento.
        Dim documento As New XmlDocument
        documento.Load(queFichero)
        '' Obtenemos el nodo raiz del documento
        Dim raiz As XmlElement = documento.DocumentElement                              'Ej: empleados
        '' Obtenemos la lista completa de los nodos que cuelgan.
        Dim listaelementos As XmlNodeList = documento.SelectNodes(elementoslistar)     'Ej: "empleados/empleado"
        ''
        Dim cambiado As Boolean = False
        '' Iteramos con listaelementos para localizar el que queremos borrar/cambiar
        For Each item As XmlNode In listaelementos
            If item.InnerText = NodoBusco Then
                '' Borrar Nodo
                If borrarNodo = True Then
                    Try
                        raiz.RemoveChild(item)
                        cambiado = True
                    Catch ex As Exception
                        resultado &= "Error borrando Elemento --> " & NodoBusco & vbCrLf & vbCrLf
                    End Try
                    Exit For
                End If
                ''
                '' Cambiar dato. Si queDato <> ""
                If DatoBusco <> "" Then
                    For Each itemHijo As XmlNode In item.ChildNodes
                        '' Cambiar dato de los los hijos
                        If itemHijo.InnerText = DatoBusco Then
                            If itemHijo.Value <> DatoValor Then
                                itemHijo.Value = DatoValor
                                cambiado = True
                            End If
                            Exit For
                        End If
                        ''
                        '' Cambiar atributo de los hijos. Si AttHijo <> ""
                        If AttHijo <> "" Then
                            Dim itemHijoAtt As XmlAttribute = Nothing
                            Try
                                itemHijoAtt = itemHijo.Attributes(AttHijo)
                            Catch ex As Exception
                                '' No existía el atributo. No hacemos nada.
                                resultado &= "No existe el atributo --> " & AttHijo & " en " & DatoBusco & vbCrLf & vbCrLf
                            End Try
                            ''
                            If itemHijoAtt IsNot Nothing AndAlso itemHijoAtt.Value <> AttHijoValor Then
                                itemHijoAtt.Value = AttHijoValor
                                cambiado = True
                            End If
                        End If
                    Next
                End If
                '' Cambiar atributo del Nodo. Si AttNodo <> ""
                If AttNodo <> "" Then
                    Dim itemAtt As XmlAttribute = Nothing
                    Try
                        itemAtt = item.Attributes(AttNodo)
                    Catch ex As Exception
                        '' No existía el atributo. No hacemos nada.
                        resultado &= "No existe el atributo --> " & AttNodo & " en " & NodoBusco & vbCrLf & vbCrLf
                    End Try
                    ''
                    If itemAtt IsNot Nothing AndAlso itemAtt.Value <> AttNodoValor Then
                        itemAtt.Value = AttNodoValor
                        cambiado = True
                    End If
                End If
            End If
        Next
        ''
        If cambiado = True Then
            documento.Save(queFichero)
        End If
        ''
        Return resultado
    End Function
    ''
    Public Function LeeXMLLenguajes_Linq(queFichero As String) As SortedList
        Dim resultado As New SortedList
        ''
        Dim FeedXML As System.Xml.Linq.XDocument = XDocument.Load(queFichero)
        ''
        Dim Feeds = From Feed In FeedXML.Descendants("language")
                    Where (Feed.Attribute("status") Is Nothing) OrElse (Feed.Attribute("status").Value <> "disabled")
                    Select languageCode = Feed.Element("languageCode").Value,
                            description = Feed.Element("description").Value
        ''
        For Each item In Feeds
            If resultado.Contains(item.languageCode) = False Then resultado.Add(item.languageCode, item.description)
        Next
        ''
        Return resultado
    End Function
    ''
    ''
    Public Function LeeXMLCompanies_Linq(queFichero As String) As SortedList
        Dim resultado As New SortedList
        ''
        Dim companiesXML As System.Xml.Linq.XDocument = XDocument.Load(queFichero)
        ''
        Dim companies = From company In companiesXML.Descendants("company")
                        Where (company.Attribute("status") Is Nothing) OrElse (company.Attribute("status").Value <> "disabled")
                        Select companyCode = company.Element("companyCode").Value,
                                description = company.Element("description").Value,
                                languageCode = company.Element("languageCode").Value,
                                countryCode = company.Element("countryCode").Value
        ''
        For Each item In companies
            'If resultado.Contains(item.languageCode) = False Then
            '    resultado.Add(item.languageCode, item.description)
            'End If
        Next
        ''
        Return resultado
    End Function
    ''
    Public Sub RellenaCompaniesXML_Linq(queFichero As String)
        Dim companiesXML As System.Xml.Linq.XDocument = XDocument.Load(queFichero)
        ''
        Dim companies = From company In companiesXML.Descendants("company")
                        Where (company.Attribute("status") Is Nothing) OrElse (company.Attribute("status").Value <> "disabled")
                        Select companyCode = company.Element("companyCode").Value,
                                description = company.Element("description").Value,
                                languageCode = company.Element("languageCode").Value,
                                countryCode = company.Element("countryCode").Value
        ''
        'colMercadosCod = New SortedList
        'colMercadosDes = New SortedList
        'arrMercadosCod = New ArrayList
        'arrMercadosDes = New ArrayList
        ''
        'For Each item In companies
        '    cComs = New clsCompanies(item.companyCode, item.description, item.languageCode, item.countryCode)
        '    ''
        '    If arrMercadosCod.Contains(item.companyCode) = False Then arrMercadosCod.Add(item.companyCode)
        '    If colMercadosCod.ContainsKey(item.companyCode) = False Then colMercadosCod.Add(item.companyCode, cComs)
        '    If arrMercadosDes.Contains(item.description) = False Then arrMercadosDes.Add(item.description)
        '    If colMercadosDes.ContainsKey(item.description) = False Then colMercadosDes.Add(item.description, cComs)
        '    arrMercadosCod.Sort()
        '    arrMercadosDes.Sort()
        'Next
    End Sub
    ''
    '' ***** Asociado a clsProyecto (Para crear o modificar nodos)
    ''
    '' Crear fichero XML, si no exite
    Public Sub XMLLinq_CreaXDocument(fullPath As String)
        If IO.File.Exists(fullPath) = True Then Exit Sub
        '' Crear el documento y el XElement raiz "OFERTAS"
        Dim xdoc As XDocument = New XDocument(New XDeclaration("1.0", "utf-8", "yes"), New XElement("OFERTAS"))
        xdoc.Save(fullPath)
    End Sub
    '' Crear/Modificar un nodo con los datos de clsProyecto
    'Public Sub XMLLinq_CreaModificaXElement(fullPath As String, cPro As clsProyecto)
    '    Dim modificado As Boolean = False
    '    ''
    '    If IO.File.Exists(fullPath) = False Then XMLLinq_CreaXDocument(fullPath)
    '    ''
    '    Dim xdoc As XDocument = XDocument.Load(fullPath)
    '    '' Si ya está, lo modificamos. Y si no está lo creamos.
    '    If XMLLinq_DameXElementDatoPorClaveProyecto(fullPath, cPro.claveproyecto, "claveProyecto") = "" Then
    '        '' Crearlo
    '        Dim xRoot As XElement = xdoc.Element("OFERTAS")    ' xdoc.Root    'xdoc.element("OFERTAS")
    '        xRoot.Add(New XElement("directorio", cPro.directorio))
    '        xRoot.Add(New XElement("nombre", cPro.nombre))
    '        xRoot.Add(New XElement("descripcion", cPro.descripcion))
    '        xRoot.Add(New XElement("tipo", cPro.tipo))
    '        xRoot.Add(New XElement("claveproyecto", cPro.claveproyecto))
    '        ''
    '        xRoot.Add(New XElement("fechaCreacion", cPro.fechaCreacion))
    '        xRoot.Add(New XElement("fechaModificacion", cPro.fechaModificacion))
    '        xRoot.Add(New XElement("usuarioCreacion", cPro.usuarioCreacion))
    '        xRoot.Add(New XElement("usuarioModificacion", cPro.usuarioModificacion))
    '        xRoot.Add(New XElement("maquinaCrea", cPro.maquinaCrea))
    '        xRoot.Add(New XElement("maquinaModifica", cPro.maquinaModifica))
    '        modificado = True
    '    Else
    '        '' Modificarlo
    '        Dim modificar As XElement = XMLLinq_DameXElementPorClaveProyecto(fullPath, cPro.claveproyecto)
    '        ''
    '        If modificar.Element("directorio").Value <> cPro.directorio Then
    '            modificar.Element("directorio").Value = cPro.directorio
    '            modificado = True
    '        End If
    '        ''
    '        If modificar.Element("nombre").Value <> cPro.nombre Then
    '            modificar.Element("nombre").Value = cPro.nombre
    '            modificado = True
    '        End If
    '        ''
    '        If modificar.Element("descripcion").Value <> cPro.descripcion Then
    '            modificar.Element("descripcion").Value = cPro.descripcion
    '            modificado = True
    '        End If
    '        ''
    '        If modificar.Element("tipo").Value <> cPro.tipo Then
    '            modificar.Element("tipo").Value = cPro.tipo
    '            modificado = True
    '        End If
    '        ''
    '        If modificar.Element("claveproyecto").Value <> cPro.claveproyecto Then
    '            modificar.Element("claveproyecto").Value = cPro.claveproyecto
    '            modificado = True
    '        End If
    '        ''
    '        If modificar.Element("fechaCreacion").Value <> cPro.fechaCreacion Then
    '            modificar.Element("fechaCreacion").Value = cPro.fechaCreacion
    '            modificado = True
    '        End If
    '        ''
    '        If modificar.Element("fechaModificacion").Value <> cPro.fechaModificacion Then
    '            modificar.Element("fechaModificacion").Value = cPro.fechaModificacion
    '            modificado = True
    '        End If
    '        ''
    '        If modificar.Element("usuarioCreacion").Value <> cPro.usuarioCreacion Then
    '            modificar.Element("usuarioCreacion").Value = cPro.usuarioCreacion
    '            modificado = True
    '        End If
    '        ''
    '        If modificar.Element("usuarioModificacion").Value <> cPro.usuarioModificacion Then
    '            modificar.Element("usuarioModificacion").Value = cPro.usuarioModificacion
    '            modificado = True
    '        End If
    '        ''
    '        If modificar.Element("maquinaCrea").Value <> cPro.maquinaCrea Then
    '            modificar.Element("maquinaCrea").Value = cPro.maquinaCrea
    '            modificado = True
    '        End If
    '        ''
    '        If modificar.Element("maquinaModifica").Value <> cPro.maquinaModifica Then
    '            modificar.Element("maquinaModifica").Value = cPro.maquinaModifica
    '            modificado = True
    '        End If
    '    End If
    '    ''
    '    If modificado = True Then
    '        xdoc.Save(fullPath)
    '    End If
    'End Sub
    ''
    '' Editar un nodo en clsProyecto
    ''
    '' Buscar un nodo en clsProyecto
    'Public Function XMLLinq_DameclsProyectoPorClaveProyecto(fullPath As String, claveProyeto As String) As clsProyecto
    '    Dim resultado As clsProyecto = Nothing
    '    '' Cargar el fichero XML
    '    Dim proyectos As XDocument = XDocument.Load(fullPath)
    '    ''
    '    '' Obtener todos los proyecto cuya claveproyecto = claveProyecto
    '    Dim contactosL As IList(Of XElement) =
    '        From c As XElement In proyectos.Elements("OFERTAS").Elements("Accion")
    '        Where c.Element("claveProyecto").Value = claveProyeto
    '        Select c
    '    ''
    '    If contactosL.Count > 0 Then
    '        resultado = New clsProyecto(contactosL.First)
    '    End If
    '    ''
    '    Return resultado
    'End Function
    ''
    Public Function XMLLinq_DameXElementPorClaveProyecto(fullPath As String, claveProyeto As String) As XElement
        Dim resultado As XElement = Nothing
        '' Cargar el fichero XML
        Dim proyectos As XDocument = XDocument.Load(fullPath)
        ''
        '' Obtener todos los proyecto cuya claveproyecto = claveProyecto
        Dim contactosL As IList(Of XElement) =
            From c As XElement In proyectos.Elements("OFERTAS").Elements("Accion")
            Where c.Element("claveProyecto").Value = claveProyeto
            Select c
        ''
        If contactosL.Count > 0 Then
            resultado = contactosL.First
        End If
        ''
        Return resultado
    End Function
    Public Function XMLLinq_DameXElementDatoPorClaveProyecto(fullPath As String, claveProyeto As String, queDato As String) As String
        Dim resultado As String = Nothing
        '' Cargar el fichero XML
        Dim proyectos As XDocument = XDocument.Load(fullPath)
        ''
        '' Obtener todos los proyecto cuya claveproyecto = claveProyecto
        Dim contactosL As IList(Of XElement) =
            From c As XElement In proyectos.Elements("OFERTAS").Elements("Accion")
            Where c.Element("claveProyecto").Value = claveProyeto
            Select c
        ''
        If contactosL.Count > 0 Then
            resultado = contactosL.First.Element(queDato).Value
        End If
        ''
        Return resultado
    End Function
    ''
    '' Borrar un nodo de clsProyecto
    ''
    'Public Sub RellenaArticulosXML_Linq(queFichero As String)
    '    colArticulos = New SortedList
    '    ''
    '    Dim ArticulosXML As System.Xml.Linq.XDocument = XDocument.Load(queFichero)
    '    ''
    '    For Each articulo As System.Xml.Linq.XElement In ArticulosXML.Descendants("article")
    '        cArts = New clsArticulos
    '        '' ARTICLECODE  (<articleCode>0000101</articleCode>)
    '        Try
    '            cArts.artcode = articulo.Element("articleCode").Value
    '        Catch ex As Exception
    '            cArts.artcode = ""
    '        End Try
    '        '' ACTIVE   (<active>True</active>)
    '        'Try
    '        '    cArts.active = CBool(articulo.Element("active"))
    '        'Catch ex As Exception
    '        '    cArts.active = False
    '        'End Try
    '        ''
    '        If colArticulos.ContainsKey(cArts.artcode) = True Or cArts.artcode = "" Then
    '            Continue For
    '        End If
    '        ''
    '        '' ** description ** Rellenar el Hashtable de descripciones.
    '        '<description language="local">TAPE H=3,20x0,40</description>
    '        '<description language="es">TAPE H=3,2x0,4</description>
    '        '<description language="es-ES">TAPE H=3,2x0,4</description>
    '        '<description language="es-PE">TAPE H=3,2x0,4</description>
    '        '<description language="es-CL">TAPE H=3,2x0,4</description>
    '        '<description language="es-MX">TAPE H=3,2x0,4</description>
    '        '<description language="en">COMPENSATION PLATE H=3,2x0,4</description>
    '        '<description language="en-US">TAPE H=3,2x0,4</description>
    '        '<description language="de-DE">TAPE H=3,2x0,4</description>
    '        '<description language="fr"></description>
    '        '<description language="fr-FR">TAPE H=3,2x0,4</description>
    '        '<description language="it-IT">TAPE H=3,2x0,4</description>
    '        '<description language="pl-PL">TAPE H=3,2x0,4</description>
    '        '<description language="pt-PT">TAPE H=3,20x0,40</description>
    '        '<description language="cz-CZ">TAPE H=3,2x0,4</description>
    '        '<description language="sk-SK">COMPENSATION PLATE H=3,2x0,4</description>
    '        '<description language="ro-RO">TAPE H=3,2x0,4</description>
    '        cArts.colDescripciones = New Hashtable
    '        For Each des As XElement In articulo.Elements("description")
    '            Dim queDes As String = ""
    '            Try
    '                queDes = des.Value
    '            Catch ex As Exception
    '                queDes = ""
    '            End Try
    '            cArts.colDescripciones.Add(des.Attribute("language").Value, queDes)
    '        Next
    '        '' Para asegurarnos que están "de-DE" y "en-DE"
    '        If cArts.colDescripciones.ContainsKey("de-DE") = False Then
    '            cArts.colDescripciones.Add("de-DE", "")
    '        End If
    '        If cArts.colDescripciones.ContainsKey("en-DE") = False Then
    '            cArts.colDescripciones.Add("en-DE", "")
    '        End If
    '        ''
    '        '' ** adicionalData **
    '        '   <aditionalData>
    '        '       <material>Acero</material>
    '        '       <weight unit="kg">57</weight>
    '        '       <formArea unit="m2">1.024</formArea>
    '        '   </aditionalData>
    '        For Each aData As XElement In articulo.Elements("aditionalData").Descendants
    '            Dim dato As String = aData.Value
    '            ''
    '            If aData.Name = "material" Then
    '                cArts.material = dato
    '            ElseIf aData.Name = "weight" Then
    '                cArts.pesoUnid = aData.Attribute("unit").Value
    '                ''
    '                If IsNumeric(dato) AndAlso dato <> "" Then
    '                    dato = FormatNumber(dato.ToString, 2)
    '                End If
    '                ''
    '                cArts.peso = dato
    '            ElseIf aData.Name = "formArea" Then
    '                cArts.areaUnid = aData.Attribute("unit").Value
    '                ''
    '                If IsNumeric(dato) AndAlso dato <> "" Then
    '                    dato = FormatNumber(dato.ToString, 2)
    '                End If
    '                cArts.area = dato
    '            End If

    '            '' Si las unidades de longitud son diferentes de milímetros, convierte a milimetros.
    '            If cArts.areaUnid.ToUpper <> "M2" And IsNumeric(cArts.peso) Then
    '                '' Revit internamente tiene las unidades de peso en Kilos.
    '                Select Case cArts.pesoUnid.ToUpper
    '                    Case "LB"
    '                        cArts.peso = UnitUtils.Convert(CDbl(FormatNumber(cArts.peso.Trim, 2)), DisplayUnitType.DUT_POUNDS_MASS, DisplayUnitType.DUT_KILOGRAMS_MASS).ToString
    '                    Case Else
    '                End Select
    '                'cArts.peso = UnitUtils.ConvertToInternalUnits(CDbl(FormatNumber(cArts.peso.Trim, 2)), DisplayUnitType.DUT_KILOGRAMS_MASS).ToString
    '                cArts.pesoUnid = "kg"
    '                cArts.areaUnid = "m2"
    '                cArts.area = "1"
    '            End If
    '            ''
    '        Next
    '        '' TOTALES DE PESO y AREA
    '        If cArts.peso <> "" And cArts.pesoUnid <> "" Then
    '            cArts.pesotodo = cArts.peso & " " & cArts.pesoUnid
    '        ElseIf cArts.peso <> "" And cArts.pesoUnid = "" Then
    '            cArts.pesotodo = cArts.peso
    '        Else
    '            cArts.pesotodo = ""
    '        End If
    '        ''
    '        If cArts.area <> "" And cArts.areaUnid <> "" Then
    '            cArts.areatodo = cArts.area & " " & cArts.areaUnid
    '        ElseIf cArts.area <> "" And cArts.areaUnid = "" Then
    '            cArts.areatodo = cArts.area
    '        Else
    '            cArts.areatodo = ""
    '        End If
    '        ''
    '        '' Lo añadimos al Hashtable directamente. Ya que si existiera lo habríamos saltado antes.
    '        colArticulos.Add(cArts.artcode, cArts)
    '    Next
    'End Sub
    ''
    ''
    '    Public Sub RellenaFamCodeXML_Linq(queFicheros() As String)
    '        '<dynamicFamilies>
    '        '        <dynamicFamily>
    '        '           <FAMILY_CODE>3_LAYER_PLYWOOD</FAMILY_CODE>
    '        '           <ITEM_LENGTH unit="">0</ITEM_LENGTH>
    '        '           <ITEM_WIDTH unit="">0</ITEM_WIDTH>
    '        '           <ITEM_HEIGHT unit="mm">21</ITEM_HEIGHT>
    '        '           <ITEM_GENERIC>2</ITEM_GENERIC>
    '        '           <ITEM_CODE>7251G00</ITEM_CODE>
    '        '           <ITEM_DESCRIPTION language="en">3 LAYER PLYWOOD 21 (M2)</ITEM_DESCRIPTION>
    '        '           <ITEM_DESCRIPTION language="es">TRICAPA 21 (M2)</ITEM_DESCRIPTION>
    '        '       </dynamicFamily>
    '        '</dynamicFamilies>
    '        colFamCode = New Hashtable
    '        ''
    '        '' Comprobar si existen los ficheros y poner 120_dynamicfamilies.xml si no existe ninguno.
    '        Dim fiOrigen As String = IO.Path.Combine(_dirApp, ficherosXML & "\120_dynamicfamilies.xml")
    '        'Dim encontrado As Boolean = False
    '        For Each queF As String In queFicheros
    '            If queF = "" Then Continue For
    '            ''
    '            If IO.File.Exists(queF) = False Then
    '                IO.File.Copy(fiOrigen, queF, False)
    '            End If
    '        Next
    '        'If encontrado = False Then
    '        'queFicheros(0) = IO.Path.Combine(_dirApp, ficherosXML & "\120_dynamicfamilies.xml")
    '        'End If
    '        ''
    '        For Each queFichero As String In queFicheros
    '            '' Recibimos array con los 3 ficheros a procesar. Si no existe el fichero, continuar.
    '            If IO.File.Exists(queFichero) = False Then Continue For
    '            ''
    '            Dim mercado As String = Split(IO.Path.GetFileNameWithoutExtension(queFichero), "_"c)(0)
    '            '' Cargar dynamicFamilies
    '            Dim dynamicFamilies As System.Xml.Linq.XDocument = XDocument.Load(queFichero)
    '            ''
    '            For Each articulo As System.Xml.Linq.XElement In dynamicFamilies.Descendants("dynamicFamily")
    '                cFC = New clsFamCode
    '                cFC.IMARKET = mercado
    '                ''
    '                '' FAMILY_CODE  (<FAMILY_CODE>3_LAYER_PLYWOOD</FAMILY_CODE>)
    '                Try
    '                    cFC.FCODE = articulo.Element("FAMILY_CODE").Value
    '                Catch ex As Exception
    '                    cFC.FCODE = ""
    '                End Try
    '                Dim unidTemp As String = ""
    '                ''
    '                '' ITEM_LENGTH  (<ITEM_LENGTH unit="">0</ITEM_LENGTH>)
    '                Try
    '                    cFC.ILENGTH = articulo.Element("ITEM_LENGTH").Value
    '                Catch ex As Exception
    '                    cFC.ILENGTH = ""
    '                End Try
    '                '' ITEM_LENGTH (Atributo unit)
    '                Try
    '                    cFC.IUnid = articulo.Element("ITEM_LENGTH").Attribute("unit").Value
    '                Catch ex As Exception
    '                    cFC.IUnid = ""
    '                End Try
    '                ''
    '                '' ITEM_WIDTH  (<ITEM_WIDTH unit="">0</ITEM_WIDTH>)
    '                Try
    '                    cFC.IWIDTH = articulo.Element("ITEM_WIDTH").Value
    '                Catch ex As Exception
    '                    cFC.IWIDTH = ""
    '                End Try
    '                '' ITEM_WIDTH (Atributo unit)
    '                Try
    '                    unidTemp = articulo.Element("ITEM_WIDTH").Attribute("unit").Value
    '                    If cFC.IUnid = "" And unidTemp <> "" Then cFC.IUnid = unidTemp
    '                    unidTemp = ""
    '                Catch ex As Exception
    '                    ''
    '                End Try
    '                ''
    '                '' ITEM_HEIGHT  (<ITEM_HEIGHT unit="">0</ITEM_HEIGHT>)
    '                Try
    '                    cFC.IHEIGHT = articulo.Element("ITEM_HEIGHT").Value
    '                Catch ex As Exception
    '                    cFC.IHEIGHT = ""
    '                End Try
    '                '' ITEM_LENGTH (Atributo unit)
    '                Try
    '                    unidTemp = articulo.Element("ITEM_HEIGHT").Attribute("unit").Value
    '                    If cFC.IUnid = "" And unidTemp <> "" Then cFC.IUnid = unidTemp
    '                Catch ex As Exception
    '                    ''
    '                End Try
    '                '' Si finalmente no tiene unidades, ponemos milimetros
    '                If cFC.IUnid = "" Then cFC.IUnid = "mm"
    '                ''
    '                '' ITEM_GENERIC  (<ITEM_GENERIC>2</ITEM_GENERIC>)
    '                Try
    '                    cFC.IGENERIC = articulo.Element("ITEM_GENERIC").Value
    '                Catch ex As Exception
    '                    cFC.IGENERIC = ""
    '                End Try
    '                ''
    '                '' ITEM_CODE  (<ITEM_CODE>7251G00</ITEM_CODE>)
    '                Try
    '                    cFC.ICODE = articulo.Element("ITEM_CODE").Value
    '                Catch ex As Exception
    '                    cFC.ICODE = ""
    '                End Try
    '                ''
    '                ''<ITEM_DESCRIPTION language="en">3 LAYER PLYWOOD 2500x500x27</ITEM_DESCRIPTION>
    '                ''<ITEM_DESCRIPTION language="es">TRICAPA 2500x500x27</ITEM_DESCRIPTION>
    '                For Each des As XElement In articulo.Elements("ITEM_DESCRIPTION")
    '                    If des.Attribute("language").Value.ToUpper = "EN" Then
    '                        cFC.IDESCRIPTIONEN = des.Value
    '                    ElseIf des.Attribute("language").Value.ToUpper = "ES" Then
    '                        cFC.IDESCRIPTIONES = des.Value
    '                    End If
    '                Next
    '                '' Fundamental ejecutar esta función. Pone las medidas en mm con 2 decimales y clave.
    '                cFC.PonMedidas()
    '                '' 
    '                '' Finalmente añadimos la clase cFC a la colección colFamCode (key=clave, value=cFC)
    '                If colFamCode.ContainsKey(cFC.clave) = False Then
    '                    colFamCode.Add(cFC.clave, cFC)
    '                End If
    '                '' Y también lo añadiremos con la clave de genérico. Para facilitar búsquedas
    '                'If colFamCode.ContainsKey(cFC.clave) = False Then
    '                'colFamCode.Add(cFC.clave, cFC)
    '                'End If
    '            Next
    '        Next
    '    End Sub
End Module
'' FICHERO: \offlineBDIdata\languages.xml
'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
'<languages>
'  <language>
'    <languageCode>es</languageCode>
'    <description>Español</description>
'  </language>
'  <language>
'    <languageCode>es-ES</languageCode>
'    <description>Español(España)</description>
'  </language>
'  <language>
'    <languageCode>es-PE</languageCode>
'    <description>Español(Perú)</description>
'  </language>
'  <language>
'    <languageCode>es-CL</languageCode>
'    <description>Español(Chile)</description>
'  </language>
'  <language>
'    <languageCode>es-MX</languageCode>
'    <description>Español(México)</description>
'  </language>
'  <language>
'    <languageCode>en</languageCode>
'    <description>English</description>
'  </language>
'  <language>
'    <languageCode>en-US</languageCode>
'    <description>English(USA)</description>
'  </language>
'  <language>
'    <languageCode>en-DE</languageCode>
'    <description>English(Germany)</description>
'  </language>
'  <language>
'    <languageCode>fr-FR</languageCode>
'    <description>Français(France)</description>
'  </language>
'  <language>
'    <languageCode>it-IT</languageCode>
'    <description>Italiano(Italia)</description>
'  </language>
'  <language>
'    <languageCode>pl-PL</languageCode>
'    <description>Polski(Polska)</description>
'  </language>
'  <language>
'    <languageCode>pt-PT</languageCode>
'    <description>Português(Portugal)</description>
'  </language>
'  <language>
'    <languageCode>cz-CZ</languageCode>
'    <description>Czech(Czech Republic)</description>
'  </language>
'  <language>
'    <languageCode>sk-SK</languageCode>
'    <description>Slovak(Slovakia)</description>
'  </language>
'  <language>
'    <languageCode>ro-RO</languageCode>
'    <description>Romanian(Romania)</description>
'  </language>
'</languages>

'"C:\XMLPrueba.xml"
'<?xml version="1.0" encoding="UTF-8"?>
'<usuarios>
'  <name codigo="mtorres">
'    <nombre>Maria  </nombre>
'    <apellido>Torres  </apellido>
'  </name>
'  <name codigo="cortiz">
'    <nombre>Carlos  </nombre>
'    <apellido>Ortiz  </apellido>
'  </name>
'</usuarios>
''
'' Utilizando XMLTextReader
'Public Sub LeeXML_XMLTextReader()
'    Dim m_xmlr As XmlTextReader
'    'Creamos el TextReader
'    m_xmlr = New XmlTextReader("C:\XMLPrueba.xml")

'    'Desabilitamos las lineas en blanco, 
'    'ya no las necesitamos
'    m_xmlr.WhitespaceHandling = WhitespaceHandling.None

'    'Leemos el archivo y avanzamos al tag de usuarios
'    m_xmlr.Read()

'    'Leemos el tag usuarios
'    m_xmlr.Read()

'    'Creamos la secuancia que nos permite 
'    'leer el archivo
'    While Not m_xmlr.EOF
'        'Avanzamos al siguiente tag
'        m_xmlr.Read()

'        'si no tenemos el elemento inicial 
'        'debemos salir del ciclo
'        If Not m_xmlr.IsStartElement() Then
'            Exit While
'        End If

'        'Obtenemos el elemento codigo
'        Dim mCodigo = m_xmlr.GetAttribute("codigo")
'        'Read elements firstname and lastname

'        m_xmlr.Read()
'        'Obtenemos el elemento del Nombre del Usuario
'        Dim mNombre = m_xmlr.ReadElementString("nombre")

'        'Obtenemos el elemento del Apellido del Usuario
'        Dim mApellido = m_xmlr.ReadElementString("apellido")

'        'Escribimos el resultado en la consola, 
'        'pero tambien podriamos utilizarlos en
'        'donde deseemos    
'        Console.WriteLine("Codigo usuario: " & mCodigo _
'          & " Nombre: " & mNombre _
'          & " Apellido: " & mApellido)
'        Console.Write(vbCrLf)
'    End While

'    'Cerramos la lactura del archivo
'    m_xmlr.Close()

'End Sub
''
''
'' Utilizando XMLDocument
'Sub LeeXML_XMLDocument
'    Try
'        Dim m_xmld As XmlDocument
'        Dim m_nodelist As XmlNodeList
'        Dim m_node As XmlNode

'        'Creamos el "Document"
'        m_xmld = New XmlDocument()

'        'Cargamos el archivo
'        m_xmld.Load("C:\XMLPrueba.xml")

'        'Obtenemos la lista de los nodos "name"
'        m_nodelist = m_xmld.SelectNodes("/usuarios/name")

'        'Iniciamos el ciclo de lectura
'        For Each m_node In m_nodelist
'            'Obtenemos el atributo del codigo
'            Dim mCodigo = m_node.Attributes.GetNamedItem("codigo").Value

'            'Obtenemos el Elemento nombre
'            Dim mNombre = m_node.ChildNodes.Item(0).InnerText

'            'Obtenemos el Elemento apellido
'            Dim mApellido = m_node.ChildNodes.Item(1).InnerText

'            'Escribimos el resultado en la consola, 
'            'pero tambien podriamos utilizarlos en
'            'donde deseemos
'            Console.Write("Codigo usuario: " & mCodigo _
'              & " Nombre: " & mNombre _
'              & " Apellido: " & mApellido)
'            Console.Write(vbCrLf)

'        Next
'    Catch ex As Exception
'        'Error trapping
'        Console.Write(ex.ToString())
'    End Try
'End Sub