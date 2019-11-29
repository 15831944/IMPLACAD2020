Imports System
Imports System.Windows.Forms
''
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.Windows
Imports Autodesk.AutoCAD.Interop

Module modIMO
    'Public colPro As Dictionary(Of String, clsDatosIni)
    Public _showTipsOnDisabled As Boolean = False
    Public _visibleToolbars As List(Of String)
    '' OFERTAS
    Public Sub cDatosIni_Crea_Llena_DesdeDirINI(queDir As String, qTipo As TIPOPRO)
        If IO.Directory.Exists(queDir) = False Then Exit Sub
        ''
        If cIni Is Nothing Then cIni = New clsINI
        ''
        '' Variables para leer datos del .Ini (Si existe)
        Dim queIni As String = IO.Path.Combine(queDir, nPlantillaIni)
        ''
        Dim año As String = ""
        Dim tipo As String = ""
        Dim cliente As String = ""
        ''
        Dim accion As String = ""
        Dim obra As String = ""
        Dim modelo As String = ""
        ''
        Dim descripcion As String = ""
        Dim directorio As String = queDir
        ''
        Dim fechaCreacion As String = ""
        Dim fechaModificacion As String = ""
        Dim usuarioCreacion As String = ""
        Dim usuarioModificacion As String = ""
        Dim maquinaCrea As String = ""
        Dim maquinaModifica As String = ""
        Dim parametros As New Hashtable
        ''
        '' Rellenamos los datos, cogidos del nombre del directorio.
        '' Según el tipo de proyecto activo.
        If qTipo = TIPOPRO.OFERTAS Then
            año = DirectorioDameDatos(queDir, DATODIR.PadreP_Solo)
            tipo = DirectorioDameDatos(pathDirOf, DATODIR.NombreSolo)
            cliente = DirectorioDameDatos(queDir, DATODIR.Padre_Solo)
            ''
            accion = DirectorioDameDatos(queDir, DATODIR.NombreSolo)
            obra = ""
            modelo = ""
        ElseIf qTipo = TIPOPRO.INSTALACIONES Then
            año = ""
            tipo = DirectorioDameDatos(pathDirIn, DATODIR.NombreSolo)
            cliente = DirectorioDameDatos(queDir, DATODIR.Padre_Solo)
            ''
            accion = ""
            obra = DirectorioDameDatos(queDir, DATODIR.NombreSolo)
            modelo = ""
        ElseIf qTipo = TIPOPRO.MSTANDARD Then
            año = ""
            tipo = DirectorioDameDatos(pathDirMs, DATODIR.NombreSolo)
            cliente = ""
            ''
            accion = ""
            obra = ""
            modelo = DirectorioDameDatos(queDir, DATODIR.NombreSolo)
        End If
        '' Objeto DirectoryInfo, para leer datos del directorio.
        'Dim oDInf As New IO.DirectoryInfo(queDir)
        ''
        If IO.File.Exists(queIni) Then
            descripcion = cIni.IniGet(queIni, "OPCIONES", "descripcion")
            fechaCreacion = cIni.IniGet(queIni, "OPCIONES", "fechaCreacion")
            fechaModificacion = cIni.IniGet(queIni, "OPCIONES", "fechaModificacion")
            usuarioCreacion = cIni.IniGet(queIni, "OPCIONES", "usuarioCreacion")
            usuarioModificacion = cIni.IniGet(queIni, "OPCIONES", "usuarioModificacion")
            maquinaCrea = cIni.IniGet(queIni, "OPCIONES", "maquinaCrea")
            maquinaModifica = cIni.IniGet(queIni, "OPCIONES", "maquinaModifica")
            '' PARAMETROS
            Dim txtparametros As String() = cIni.IniGetSection(queIni, "PARAMETROS")
            If txtparametros Is Nothing OrElse txtparametros.Length < 1 Then
                'mensaje &= "Bloques LV no indicados... Corrija fichero .ini" & vbCrLf
            Else
                For i = 0 To UBound(txtparametros) - 1 Step 2
                    Dim nPar As String = txtparametros(i)
                    Dim vPar As String = txtparametros(i + 1)
                    parametros.Add(nPar, vPar)
                Next
            End If
        Else
            descripcion = ""
            '' OPCIONES
            fechaCreacion = DirectorioDameDatos(queDir, DATODIR.FechaCre)
            fechaModificacion = DirectorioDameDatos(queDir, DATODIR.FechaMod)
            usuarioCreacion = usuario
            usuarioModificacion = usuario
            maquinaCrea = maquina
            maquinaModifica = maquina
        End If
        ''
        ''año, tipo, cliente, accion, obra, modelo, descripcion, directorio,
        ''fechaCreacion, fechaModificacion, usuarioCreacion, usuarioModificacion, maquinaCrea, maquinaModifica
        cDatosIni = New clsDatosIni(año, tipo, cliente,
                                        accion, obra, modelo,
                                        descripcion,
                                        fechaCreacion, fechaModificacion, usuarioCreacion, usuarioModificacion, maquinaCrea, maquinaModifica,
                                    parametros)
        ''
        If IO.File.Exists(queIni) = False Then cDatosIni.FicheroINI_Llena_Desde_cDatosIni()
    End Sub
    ''
    Public Function DirectorioDameDatos(queDir As String, queDato As DATODIR) As String
        Dim resultado As String = ""
        ''
        If IO.Directory.Exists(queDir) = False Then
            Return resultado
            Exit Function
        End If
        ''
        '' C:\DIRBASE\PROYECTOS\M.STANDARD\MAQUINAS ST\MODELO1
        Dim oDInf As New IO.DirectoryInfo(queDir)
        Select Case queDato
            Case DATODIR.NombreSolo         '' MODELO1
                resultado = oDInf.Name
            Case DATODIR.Padre              'C:\DIRBASE\PROYECTOS\M.STANDARD\MAQUINAS ST
                resultado = oDInf.Parent.FullName
            Case DATODIR.Padre_Solo         '' MAQUINAS ST
                resultado = oDInf.Parent.FullName
                resultado = DirectorioDameDatos(resultado, DATODIR.NombreSolo)
            Case DATODIR.PadreP             '' C:\DIRBASE\PROYECTOS\M.STANDARD
                resultado = oDInf.Parent.Parent.FullName
            Case DATODIR.PadreP_Solo        '' M.STANDARD
                resultado = oDInf.Parent.Parent.FullName
                resultado = DirectorioDameDatos(resultado, DATODIR.NombreSolo)
            Case DATODIR.PadrePP            '' C:\DIRBASE\PROYECTOS
                resultado = oDInf.Parent.Parent.Parent.FullName
            Case DATODIR.PadrePP_Solo       '' PROYECTOS
                resultado = oDInf.Parent.Parent.Parent.FullName
                resultado = DirectorioDameDatos(resultado, DATODIR.NombreSolo)
            Case DATODIR.PadrePPP           '' C:\DIRBASE\
                resultado = oDInf.Parent.Parent.Parent.Parent.FullName
            Case DATODIR.PadrePPP_Solo      '' DIRBASE
                resultado = oDInf.Parent.Parent.Parent.Parent.FullName
                resultado = DirectorioDameDatos(resultado, DATODIR.NombreSolo)
            Case DATODIR.Raiz               '' C:\
                resultado = oDInf.Root.ToString
            Case DATODIR.FechaCre
                resultado = oDInf.CreationTime.ToShortDateString
            Case DATODIR.FechaMod
                resultado = oDInf.LastWriteTime.ToShortDateString
        End Select
        ''
        oDInf = Nothing
        ''
        Return resultado
    End Function
    ''
    '' C:\DIRBASE\PROYECTOS\M.STANDARD\MAQUINAS ST\MODELO1
    Public Enum DATODIR
        NombreSolo      '' MODELO1
        Padre           '' C:\DIRBASE\PROYECTOS\M.STANDARD\MAQUINAS ST
        Padre_Solo      '' MAQUINAS ST
        PadreP          '' C:\DIRBASE\PROYECTOS\M.STANDARD
        PadreP_Solo     '' M.STANDARD
        PadrePP         '' C:\DIRBASE\PROYECTOS
        PadrePP_Solo    '' PROYECTOS
        PadrePPP        '' C:\DIRBASE
        PadrePPP_Solo   '' DIRBASE
        Raiz            '' C:\
        FechaCre        '' IO.DirectoryInfo.CreationTime
        FechaMod        '' IO.DirectoryInfo.LastWriteTime
    End Enum
    ''
    Public Sub INIAccionEscribeOF(queIni As String, Optional esnuevo As Boolean = True)
        If cIni Is Nothing Then cIni = New clsINI
        ''
        Dim directorio As String = IO.Path.GetDirectoryName(queIni) 'IO.Path.Combine(pathDirIn, queF.lblDir.Text)
        ''
        'Dim año As String = frmOf.txtAño.Text
        'Dim tipo As String = frmOf.txtTipo.Text  ' TIPOPRO.OFERTAS.ToString
        'Dim cliente As String = DirectorioDameDatos(ultimoClienteOf, DATODIR.NombreSolo)    '.Split("·"c)(0)
        ''
        'Dim accion As String = IO.Path.GetFileNameWithoutExtension(queIni)
        'Dim obra As String = ""
        'Dim modelo As String = ""
        ''
        Dim descripcion As String = frmOf.txtDesc.Text
        ''
        Dim fechaCreacion As String = frmOf.txtDateC.Text
        Dim fechaModificacion As String = frmOf.txtDateM.Text
        Dim usuarioCreacion As String = frmOf.txtUserC.Text
        Dim usuarioModificacion As String = frmOf.txtUserM.Text
        Dim maquinaCrea As String = frmOf.txtPcC.Text
        Dim maquinaModifica As String = frmOf.txtPcM.Text
        ''
        ''
        If IO.File.Exists(queIni) = False And esnuevo Then
            IO.File.Copy(plantillaIni, queIni)
        End If
        ''
        'cIni.IniWrite(queIni, "OPCIONES", "año", año)
        'cIni.IniWrite(queIni, "OPCIONES", "tipo", tipo)
        'cIni.IniWrite(queIni, "OPCIONES", "cliente", cliente)
        'cIni.IniWrite(queIni, "OPCIONES", "accion", accion)
        'cIni.IniWrite(queIni, "OPCIONES", "obra", "")
        'cIni.IniWrite(queIni, "OPCIONES", "modelo", "")
        ''
        cIni.IniWrite(queIni, "OPCIONES", "descripcion", descripcion)
        cIni.IniWrite(queIni, "OPCIONES", "fechaCreacion", fechaCreacion)
        cIni.IniWrite(queIni, "OPCIONES", "fechaModificacion", fechaModificacion)
        cIni.IniWrite(queIni, "OPCIONES", "usuarioCreacion", usuarioCreacion)
        cIni.IniWrite(queIni, "OPCIONES", "usuarioModificacion", usuarioModificacion)
        cIni.IniWrite(queIni, "OPCIONES", "maquinaCrea", maquinaCrea)
        cIni.IniWrite(queIni, "OPCIONES", "maquinaModifica", maquinaModifica)
        ''
        cDatosIni_Crea_Llena_DesdeDirINI(IO.Path.GetDirectoryName(queIni), TIPOPRO.OFERTAS)
    End Sub
    ''
    Public Sub INIAccionEscribeIN(queIni As String, Optional esnuevo As Boolean = True)
        If cIni Is Nothing Then cIni = New clsINI
        ''
        Dim directorio As String = IO.Path.GetDirectoryName(queIni) 'IO.Path.Combine(pathDirIn, queF.lblDir.Text)
        ''
        'Dim año As String = frmIn.txtAño.Text
        'Dim tipo As String = frmIn.txtTipo.Text  ' TIPOPRO.OFERTAS.ToString
        'Dim cliente As String = DirectorioDameDatos(ultimoClienteOf, DATODIR.NombreSolo)    '.Split("·"c)(0)
        ''
        'Dim accion As String = IO.Path.GetFileNameWithoutExtension(queIni)
        'Dim obra As String = ""
        'Dim modelo As String = ""
        ''
        Dim descripcion As String = frmIn.txtDesc.Text
        ''
        Dim fechaCreacion As String = frmIn.txtDateC.Text
        Dim fechaModificacion As String = frmIn.txtDateM.Text
        Dim usuarioCreacion As String = frmIn.txtUserC.Text
        Dim usuarioModificacion As String = frmIn.txtUserM.Text
        Dim maquinaCrea As String = frmIn.txtPcC.Text
        Dim maquinaModifica As String = frmIn.txtPcM.Text
        ''
        ''
        If IO.File.Exists(queIni) = False And esnuevo Then
            IO.File.Copy(plantillaIni, queIni)
        End If
        ''
        'cIni.IniWrite(queIni, "OPCIONES", "tipo", tipo)
        'cIni.IniWrite(queIni, "OPCIONES", "año", año)
        'cIni.IniWrite(queIni, "OPCIONES", "cliente", cliente)
        'cIni.IniWrite(queIni, "OPCIONES", "accion", accion)
        'cIni.IniWrite(queIni, "OPCIONES", "obra", "")
        'cIni.IniWrite(queIni, "OPCIONES", "modelo", "")
        cIni.IniWrite(queIni, "OPCIONES", "descripcion", descripcion)
        cIni.IniWrite(queIni, "OPCIONES", "fechaCreacion", fechaCreacion)
        cIni.IniWrite(queIni, "OPCIONES", "fechaModificacion", fechaModificacion)
        cIni.IniWrite(queIni, "OPCIONES", "usuarioCreacion", usuarioCreacion)
        cIni.IniWrite(queIni, "OPCIONES", "usuarioModificacion", usuarioModificacion)
        cIni.IniWrite(queIni, "OPCIONES", "maquinaCrea", maquinaCrea)
        cIni.IniWrite(queIni, "OPCIONES", "maquinaModifica", maquinaModifica)
        ''
        cDatosIni_Crea_Llena_DesdeDirINI(IO.Path.GetDirectoryName(queIni), TIPOPRO.INSTALACIONES)
    End Sub
    ''

    Public Sub INIAccionEscribeMS(queIni As String, Optional esnuevo As Boolean = True)
        If cIni Is Nothing Then cIni = New clsINI
        ''
        Dim directorio As String = IO.Path.GetDirectoryName(queIni) 'IO.Path.Combine(pathDirMs, queF.lblDir.Text) '
        ''
        'Dim año As String = frmMs.txtAño.Text
        'Dim tipo As String = frmMs.txtTipo.Text  ' "M.STANDARD"  
        'Dim cliente As String = ""  'DirectorioDameDatos(ultimoClienteOf, DATODIR.NombreSolo)    '.Split("·"c)(0)
        ''
        'Dim accion As String = ""   'IO.Path.GetFileNameWithoutExtension(queIni)
        'Dim obra As String = ""
        'Dim modelo As String = frmMs.txtModelo.Text
        ''
        Dim descripcion As String = frmMs.txtDesc.Text
        ''
        Dim fechaCreacion As String = frmMs.txtDateC.Text
        Dim fechaModificacion As String = frmMs.txtDateM.Text
        Dim usuarioCreacion As String = frmMs.txtUserC.Text
        Dim usuarioModificacion As String = frmMs.txtUserM.Text
        Dim maquinaCrea As String = frmMs.txtPcC.Text
        Dim maquinaModifica As String = frmMs.txtPcM.Text
        ''
        ''
        If IO.File.Exists(queIni) = False And esnuevo Then
            IO.File.Copy(plantillaIni, queIni)
        End If
        ''
        'cIni.IniWrite(queIni, "OPCIONES", "año", año)
        'cIni.IniWrite(queIni, "OPCIONES", "tipo", tipo)
        'cIni.IniWrite(queIni, "OPCIONES", "cliente", cliente)
        ''
        'cIni.IniWrite(queIni, "OPCIONES", "accion", accion)
        'cIni.IniWrite(queIni, "OPCIONES", "obra", obra)
        'cIni.IniWrite(queIni, "OPCIONES", "modelo", modelo)
        ''
        cIni.IniWrite(queIni, "OPCIONES", "descripcion", descripcion)
        ''
        cIni.IniWrite(queIni, "OPCIONES", "fechaCreacion", fechaCreacion)
        cIni.IniWrite(queIni, "OPCIONES", "fechaModificacion", fechaModificacion)
        cIni.IniWrite(queIni, "OPCIONES", "usuarioCreacion", usuarioCreacion)
        cIni.IniWrite(queIni, "OPCIONES", "usuarioModificacion", usuarioModificacion)
        cIni.IniWrite(queIni, "OPCIONES", "maquinaCrea", maquinaCrea)
        cIni.IniWrite(queIni, "OPCIONES", "maquinaModifica", maquinaModifica)
        ''
        cDatosIni_Crea_Llena_DesdeDirINI(IO.Path.GetDirectoryName(queIni), TIPOPRO.MSTANDARD)
    End Sub
    ''
    Public Sub RibbonActiva(activar As Boolean)
        '' rellenar rps con el RibbonPaletteSet actual
        Dim rps As Autodesk.AutoCAD.Ribbon.RibbonPaletteSet
        rps = Autodesk.AutoCAD.Ribbon.RibbonServices.CreateRibbonPaletteSet()
        ''
        rps.RibbonControl.IsEnabled = activar
        ''
        If activar = False Then
            ' Store the current setting for "Show tooltips when the ribbon is disabled"
            ' And then modify the setting
            _showTipsOnDisabled = ComponentManager.ToolTipSettings.ShowOnDisabled
            ComponentManager.ToolTipSettings.ShowOnDisabled = activar
        Else
            ' Restore the setting for "Show tooltips when the ribbon is disabled"
            ComponentManager.ToolTipSettings.ShowOnDisabled = _showTipsOnDisabled
        End If
        ''      
        '' Enable Or disable background tab rendering
        rps.RibbonControl.IsBackgroundTabRenderingEnabled = activar
    End Sub
    ''
    Public Sub ToolbarActiva(activar As Boolean)
        _visibleToolbars = New List(Of String)
        ''
        ''Clear the list Of toolbars that were previously visible, If we're disabling
        If activar = False Then _visibleToolbars.Clear()
        ''
        '' Use dynamic .NET to avoid the COM typelib reference
        '' (all the implicit "vars" below will also be resolved to dynamics)
        Dim app As Object = Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication
        ''
        '' Iterate through all the menu groups
        Dim mgs As Object = app.MenuGroups
        ''
        For Each mg As Object In mgs
            '' Iterate through a menu group's toolbars
            Dim tbs As Object = mg.Toolbars
            ''
            For Each tb As Object In tbs
                '' If we're disabling, check whether the toolbar is visible
                '' And, if so, add its ID to the list before turning it off
                If activar = False Then
                    If tb.Visible = True Then
                        _visibleToolbars.Add(tb.TagString)
                        tb.Visible = activar
                    End If
                Else
                    '' If we're enabling, check whether the toolbar is in our list
                    '' before turning it on
                    If _visibleToolbars.Contains(tb.TagString) Then
                        tb.Visible = activar
                    End If
                End If
            Next
        Next
    End Sub
    ''
    Public Sub TabPageActiva(ByRef queTc As TabControl, queTab As String, ByRef colTabs As Collection)
        Dim encontrado As Boolean = False
        For Each oTp As TabPage In queTc.TabPages
            If oTp.Name = queTab Then
                encontrado = True
                Exit For
            End If
        Next
        ''
        If encontrado = False Then
            queTc.TabPages.Insert(1, colTabs(queTab))
        End If
    End Sub
    ''
    Public Sub TabPageDesactiva(ByRef queTc As TabControl, queTab As String, ByRef colTabs As Collection)
        Exit Sub  '' Anulamos este procedimiento
        ''
        Dim encontrado As Boolean = False
        Dim tpTemp As TabPage = Nothing
        For Each oTp As TabPage In queTc.TabPages
            If oTp.Name = queTab Then
                If colTabs.Contains(queTab) = False Then
                    colTabs.Add(oTp, oTp.Name)
                End If
                tpTemp = queTc.TabPages.Item(oTp.Name)
                encontrado = True
                Exit For
            End If
        Next
        ''
        If encontrado = True Then
            '' Quitar TabPage (Proyecto) hasta que activemos un proyecto
            queTc.Controls.Remove(tpTemp)
        End If
    End Sub
    ''
    Public Function DameFechaCortaUnido(Optional ponHora As Boolean = False) As String
        Dim resultado As String = ""
        Dim año As String = Date.Now.Year.ToString
        Dim mes As String = Date.Now.Month.ToString.PadLeft(2, "0"c)
        Dim dia As String = Date.Now.Day.ToString.PadLeft(2, "0"c)
        ''
        Dim hora As String = Date.Now.Hour.ToString.PadLeft(2, "0"c)
        Dim minuto As String = Date.Now.Minute.ToString.PadLeft(2, "0"c)
        ''
        resultado = año & mes & dia
        ''
        If ponHora = True Then
            resultado &= "·" & hora & minuto
        End If
        ''
        Return resultado
    End Function
    ''
    'Public Sub AccionesRellena()
    '    ''
    '    tvAcciones.Nodes.Clear()
    '    tvAcciones.ImageList = imgList1
    '    For x As Integer = 0 To lbacciones.Items.Count - 1
    '        Dim accion As String = lbacciones.Items.Item(x)
    '        Dim queDir As String = ultimoPro.directorio & "\" & accion
    '        ''
    '        Dim oNode As New TreeNode
    '        oNode.Text = accion
    '        oNode.Name = accion
    '        oNode.Tag = queDir
    '        oNode.ContextMenuStrip = contexMAcciones
    '        oNode.ImageKey = "carpeta"
    '        oNode.SelectedImageKey = "carpeta"
    '        '' Ahora le ponemos los hijos (Carpetas y ficheros DWG)
    '        PonHijosNodo(oNode)
    '        Dim indice As Integer = tvAcciones.Nodes.Add(oNode)
    '        oNode = Nothing
    '    Next
    'End Sub
    Public Sub PonHijosNodo(ByRef queNode As TreeNode)
        If ultimofrm Is Nothing Then Exit Sub
        ''
        Dim queDir As String = queNode.Tag.ToString
        ''
        '' Primer buscamos los DWG de esta carpeta.
        For Each xDwg As String In IO.Directory.GetFiles(queDir, "*.dwg", IO.SearchOption.TopDirectoryOnly)
            Dim oNode As New TreeNode
            oNode.Text = IO.Path.GetFileName(xDwg)
            oNode.Name = oNode.Text
            oNode.Tag = xDwg
            ''
            If ultimofrm.Name = "frmOfertas" Then
                oNode.ContextMenuStrip = frmOf.contextMPlanos
            ElseIf ultimofrm.Name = "frmInstalaciones" Then
                oNode.ContextMenuStrip = frmIn.contextMPlanos
            ElseIf ultimofrm.Name = "frmMStandard" Then
                oNode.ContextMenuStrip = frmMs.contextMPlanos
            End If
            ''
            oNode.ImageKey = "dwg"
            oNode.SelectedImageKey = "dwg"
            Dim indice As Integer = queNode.Nodes.Add(oNode)
            oNode = Nothing
        Next
        '' Ahora el resto de ficheros
        For Each queFi As String In IO.Directory.GetFiles(queDir, "*.*", IO.SearchOption.TopDirectoryOnly)
            '' No poner los .dwg (Se han cargado ya)
            If IO.Path.GetExtension(queFi).ToLower = ".dwg" Then Continue For
            '' No cargar el fichero datos.ini
            If IO.Path.GetFileName(queFi).ToLower = "datos.ini" Then Continue For
            ''
            Dim oNode As New TreeNode
            oNode.Text = IO.Path.GetFileName(queFi)
            oNode.Name = oNode.Text
            oNode.Tag = queFi
            ''
            If ultimofrm.Name = "frmOfertas" Then
                oNode.ContextMenuStrip = frmOf.ContextMOtros
            ElseIf ultimofrm.Name = "frmInstalaciones" Then
                oNode.ContextMenuStrip = frmIn.ContextMOtros
            ElseIf ultimofrm.Name = "frmMStandard" Then
                oNode.ContextMenuStrip = frmMs.ContextMOtros
            End If
            ''
            Dim queimg As String = ".dwg"
            Select Case IO.Path.GetExtension(queFi).ToLower
                Case ".dxf", ".doc", ".docx", ".xls", ".xlsx"
                    queimg = "documento"
                Case ".bmp", ".png", ".jpg"
                    queimg = "imagen"
                Case Else
                    queimg = "documento"
            End Select
            oNode.ImageKey = queimg
            oNode.SelectedImageKey = queimg
            Dim indice As Integer = queNode.Nodes.Add(oNode)
            oNode = Nothing
        Next
        ''
        '' Ahora buscados las carpetas que tenga.
        For Each xDir As String In IO.Directory.GetDirectories(queDir, "*.*", IO.SearchOption.TopDirectoryOnly)
            Dim oNode As New TreeNode
            oNode.Text = IO.Path.GetFileName(xDir)
            oNode.Name = oNode.Text
            oNode.Tag = xDir
            ''
            If ultimofrm.Name = "frmOfertas" Then
                oNode.ContextMenuStrip = frmOf.contextMCarpetas
            ElseIf ultimofrm.Name = "frmInstalaciones" Then
                oNode.ContextMenuStrip = frmIn.contextMCarpetas
            ElseIf ultimofrm.Name = "frmMStandard" Then
                oNode.ContextMenuStrip = frmMs.contextMCarpetas
            End If
            ''
            oNode.ImageKey = "carpeta"
            oNode.SelectedImageKey = "carpeta"
            '' Ahora le ponemos los hijos (Carpetas y ficheros DWG)
            Dim indice As Integer = queNode.Nodes.Add(oNode)
            ''
            '' Llamar recursivamente a esta función. Sobre las carpetas.
            PonHijosNodo(oNode)
            oNode = Nothing
        Next
    End Sub
    ''
    Public Function DocumentoEstaYaAbierto(ByRef queApp As Autodesk.AutoCAD.Interop.AcadApplication, queDwg As String, Optional activar As Boolean = True) As Boolean
        Dim resultado As Boolean = False
        ''
        For Each queDoc As Autodesk.AutoCAD.Interop.AcadDocument In oApp.Documents
            If queDoc.FullName = queDwg Then
                If activar = True Then oApp.ActiveDocument = queDoc
                resultado = True
                Exit For
            End If
        Next
        ''
        Return resultado
    End Function
    ''
    Public Function DocumentoAbiertoDame(ByRef queApp As Autodesk.AutoCAD.Interop.AcadApplication, queDwg As String) As Autodesk.AutoCAD.Interop.AcadDocument
        Dim resultado As Autodesk.AutoCAD.Interop.AcadDocument = Nothing
        ''
        For Each queDoc As Autodesk.AutoCAD.Interop.AcadDocument In oApp.Documents
            If queDoc.FullName = queDwg Then
                resultado = queDoc
                Exit For
            End If
        Next
        ''
        Return resultado
    End Function
    ''
    Public Function DocumentoAbiertoCierra(ByRef queApp As Autodesk.AutoCAD.Interop.AcadApplication, queDwg As String) As Boolean
        Dim resultado As Boolean = Nothing
        ''
        For Each queDoc As Autodesk.AutoCAD.Interop.AcadDocument In oApp.Documents
            If queDoc.FullName = queDwg Then
                queDoc.Close(False)
                resultado = True
                Exit For
            End If
        Next
        ''
        Return resultado
    End Function
    ''
    Public Function FicheroGuardaDWG(carpetainicio As String) As String
        Dim resultado As String = ""
        ''
        Using oFg As New SaveFileDialog
            If IO.Directory.Exists(carpetainicio) Then oFg.InitialDirectory = carpetainicio
            oFg.AddExtension = True
            oFg.Filter = "Fichero AutoCAD (*.dwg)|*.dwg"
            'oFg.FilterIndex = 1
            If oFg.ShowDialog() = DialogResult.OK Then
                resultado = oFg.FileName
            End If
        End Using
        ''
        Return resultado
    End Function
    ''
    Public Function FicheroAbreDWG(carpetainicio As String) As String
        Dim resultado As String = ""
        ''
        Using oFo As New OpenFileDialog
            If IO.Directory.Exists(carpetainicio) Then oFo.InitialDirectory = carpetainicio
            oFo.Filter = "Fichero AutoCAD (*.dwg)|*.dwg"
            'oFg.FilterIndex = 1
            If oFo.ShowDialog() = DialogResult.OK Then
                resultado = oFo.FileName
            End If
        End Using
        ''
        Return resultado
    End Function
    ''
    ''
    Public Function FicheroGuardaOtros(carpetainicio As String) As String
        Dim resultado As String = ""
        ''
        Using oFg As New SaveFileDialog
            If IO.Directory.Exists(carpetainicio) Then oFg.InitialDirectory = carpetainicio
            'oFg.AddExtension = True
            oFg.Filter = "Todos (*.*)|*.*"
            'oFg.FilterIndex = 1
            If oFg.ShowDialog() = DialogResult.OK Then
                resultado = oFg.FileName
            End If
        End Using
        ''
        Return resultado
    End Function
    ''
    Public Function FicheroAbreOtros(carpetainicio As String) As String
        Dim resultado As String = ""
        ''
        Using oFo As New OpenFileDialog
            If IO.Directory.Exists(carpetainicio) Then oFo.InitialDirectory = carpetainicio
            oFo.Filter = "Todos (*.*)|*.*"
            'oFg.FilterIndex = 1
            If oFo.ShowDialog() = DialogResult.OK Then
                resultado = oFo.FileName
            End If
        End Using
        ''
        Return resultado
    End Function
    Public Sub BotonesActivaAccion(nombreFrom As String, queB As QUEBOTONES)
        Dim queF As Object = Nothing
        Select Case nombreFrom
            Case "frmOfertas"
                If frmOf IsNot Nothing Then
                    ' Primero desactivamos los 4 botones de acciones
                    frmOf.Nuevo_button.Enabled = False
                    frmOf.Editar_button.Enabled = False
                    frmOf.Guardar_button.Enabled = False
                    frmOf.Cancelar_button.Enabled = False
                    ' Ahora Activamos los botones necesarios, según queB
                    If queB = QUEBOTONES.accNuevo Then
                        frmOf.Nuevo_button.Enabled = True
                    ElseIf queB = QUEBOTONES.accNuevoEditar Then
                        frmOf.Nuevo_button.Enabled = True
                        frmOf.Editar_button.Enabled = True
                    ElseIf queB = QUEBOTONES.accEditando Then
                        frmOf.Guardar_button.Enabled = True
                        frmOf.Cancelar_button.Enabled = True
                    ElseIf queB = QUEBOTONES.accNada Then
                        '' No hacemos nada, ya están desactivados
                    End If
                End If
            Case "frmInstalaciones"
                If frmIn IsNot Nothing Then
                    ' Primero desactivamos los 4 botones
                    frmIn.Nuevo_button.Enabled = False
                    frmIn.Editar_button.Enabled = False
                    frmIn.Guardar_button.Enabled = False
                    frmIn.Cancelar_button.Enabled = False
                    ' Ahora Activamos los botones necesarios, según queB
                    If queB = QUEBOTONES.accNuevo Then
                        frmIn.Nuevo_button.Enabled = True
                    ElseIf queB = QUEBOTONES.accNuevoEditar Then
                        frmIn.Nuevo_button.Enabled = True
                        frmIn.Editar_button.Enabled = True
                    ElseIf queB = QUEBOTONES.accEditando Then
                        frmIn.Guardar_button.Enabled = True
                        frmIn.Cancelar_button.Enabled = True
                    ElseIf queB = QUEBOTONES.accNada Then
                        '' No hacemos nada, ya están desactivados
                    End If
                End If
            Case "frmMStandard"
                If frmMs IsNot Nothing Then
                    ' Primero desactivamos los 4 botones
                    frmMs.Nuevo_button.Enabled = False
                    frmMs.Editar_button.Enabled = False
                    frmMs.Guardar_button.Enabled = False
                    frmMs.Cancelar_button.Enabled = False
                    ' Ahora Activamos los botones necesarios, según queB
                    If queB = QUEBOTONES.accNuevo Then
                        frmMs.Nuevo_button.Enabled = True
                    ElseIf queB = QUEBOTONES.accNuevoEditar Then
                        frmMs.Nuevo_button.Enabled = True
                        frmMs.Editar_button.Enabled = True
                    ElseIf queB = QUEBOTONES.accEditando Then
                        frmMs.Guardar_button.Enabled = True
                        frmMs.Cancelar_button.Enabled = True
                    ElseIf queB = QUEBOTONES.accNada Then
                        '' No hacemos nada, ya están desactivados
                    End If
                End If
        End Select
    End Sub
    ''
    Public Sub TreeViewActivaNode(ByRef queTv As TreeView, queTexto As String)
        queTv.SelectedNode = Nothing
        For Each oNode As System.Windows.Forms.TreeNode In queTv.Nodes
            If oNode.Text.ToUpper = queTexto.ToUpper Then
                queTv.SelectedNode = oNode
                Exit For
            End If
        Next
    End Sub
    ''
    Public Sub ComboboxActivaItem(ByRef queCb As ComboBox, queTexto As String)
        queCb.SelectedIndex = -1
        For x As Integer = 0 To queCb.Items.Count - 1
            If queCb.Items.Item(x).ToString.ToUpper = queTexto.ToUpper Then
                queCb.SelectedIndex = x
                Exit For
            End If
        Next
    End Sub
    ''
    Public Sub ListboxActivaItem(ByRef queLb As ListBox, queTexto As String)
        queLb.SelectedIndex = -1
        For x As Integer = 0 To queLb.Items.Count - 1
            If queLb.Items.Item(x).ToString.ToUpper = queTexto.ToUpper Then
                queLb.SelectedIndex = x
                Exit For
            End If
        Next
    End Sub
    ''
    Public Sub ParametrosIniActualiza(queLV As ListView)
        ''
        '' Cambiar el Hashtable de la clase clsAccion con los nuevos parámetros (Creados o Cambiados)
        '' y cambiarlos en los Dictinary colPro y colProOf
        Dim colP As New Hashtable
        For Each lvI As ListViewItem In queLV.Items
            colP.Add(lvI.Name, lvI.SubItems(1).Text)
        Next
        ''
        '' Cambiar el .ini. Actualizar, Añadir y Borrar los parámetros
        ''
        If cIni Is Nothing Then cIni = New clsINI
        ''
        cDatosIni.ParametrosPonTodos(colP)
        cDatosIni.FicheroINI_Llena_Desde_cDatosIni(ultFicheroIni)
    End Sub
End Module

''
#Region "ENUMERACIONES"
Public Enum TIPOPRO
    OFERTAS
    INSTALACIONES
    MSTANDARD
End Enum
Public Enum QUEBOTONES
    cliNuevo
    cliNuevoEditar
    cliNada
    accNuevo
    accNuevoEditar
    accEditando
    accNada
End Enum
#End Region
