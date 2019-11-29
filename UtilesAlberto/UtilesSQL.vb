Imports System
Imports System.IO
Imports System.Xml
Imports System.ComponentModel
Imports System.Data
Imports System.Drawing
Imports System.Windows.Forms
Imports System.Data.SqlClient

'Pedidos_Clientes           *Tabla Pedidos Clientes ADAgeS
'Detalle_Pedidos_Clientes   *Tabla Detalle Pedidos Clientes ADAgeS
'Clientes                   *Tabla Clientes ADAgeS
'Articulos / Precios_Tarifas_Cliente *Para buscar la referencia del articulo del cliente REF_CLI y CLA_CLI
'***********************************
'System.GUID nos da un codigo único y despues aplicamos esta formula para formatearlo para ADAgeS
'left(replace(newid(),'-',''),16)

'* FORMULA PARA GENERAR LAS CLAVES PARA NUEVAS ALTAS
'* Hacer primero un "Select newid()" para generar un valor a una variable string
'* Y a esta variable le aplicamos esto que genera el código unico.
'***********************************

'Private strCadenaC As String = "User ID =sa;Password=castilla;Initial Catalog=raptor_xp_antonio;Data Source=adaserver;"
'Private strServidor As String = "adaserver"
'Private strBD As String = "raptor_xp_antonio"
'Private strU As String = "sa"
'Private strC As String = "castilla"

'' Cadena de conexión con seguridad integrada y no integrada *******************
'If My.Settings.Seguridad = False Then
'CADENA DE CONEXION ESTANDAR ************************************
'strCadenaC = "User ID =" & strU & _
'";Password=" & strC & _
'";Initial Catalog=" & strBd & _
'";Data Source=" & strServidor '& ";"
'Ejemplo de cadena de conexion:
'Else
'CADENA DE CONEXION CON SEGURIDAD DE WINDOWS INTEGRADA **********
'strCadenaC = "Data Source=" & strServidor & _
'"; Initial Catalog=" & strBd & _
'"; Integrated Security=SSPI"
'Ejemplo de cadena de conexion: "Data Source=Aron1; Initial Catalog=pubs; Integrated Security=SSPI;"
'End If
'*******************************************************************************

Partial Public Class Utiles
    Public Shared Function ConectSQL_Start(_server As String,
                                           _database As String,
                                           _user As String,
                                           _password As String,
                                           _bolSegInt As Boolean,
                                        cfg As Conf) As String
        'If cfg Is Nothing Then cfg = New Conf()
        Dim resultado As String = ""
        '"Data Source=ANGEL-HP\SQLEXPRESS;Initial Catalog=MKKitHispania;User ID=sa;Password=castilla;multipleactiveresultsets=True;"
        If _bolSegInt = True Then
            cfg._connectionString = "Data Source=" & _server & "; Initial Catalog=" & _database & "; Integrated Security=SSPI;"
        Else
            cfg._connectionString = "Data Source=" & _server & "; Initial Catalog=" & _database & ";User ID =" & _user & ";Password=" & _password & ";"
        End If
        '
        Dim Cnn As SqlConnection = Nothing
        Dim DAdapter As SqlDataAdapter = Nothing
        Dim CBuilder As SqlCommandBuilder = Nothing
        Dim DSet As DataSet = Nothing
        Dim Cmd As SqlCommand
        Try
            Cnn = New SqlConnection(cfg._connectionString)
            AddHandler Cnn.InfoMessage, AddressOf Cnn_InfoMessage
            ' Selecctionar todas las tablas de una BD SQL
            ' select * from INFORMATION_SCHEMA.TABLES
            DAdapter = New SqlDataAdapter("select * from INFORMATION_SCHEMA.TABLES", Cnn)
            DSet = New DataSet
            DAdapter.Fill(DSet, "TABLAS")
            RemoveHandler Cnn.InfoMessage, AddressOf Cnn_InfoMessage
        Catch ex As Exception
            resultado = ex.ToString
        Finally
            If Cnn IsNot Nothing AndAlso Cnn.State <> ConnectionState.Closed Then Cnn.Close()
            Cnn = Nothing
            DAdapter = Nothing
            CBuilder = Nothing
            DSet = Nothing
            Cmd = Nothing
            GC.Collect()
        End Try
        Return resultado
    End Function
    Public Shared Function ConectSQL_Start(_strCadenaC As String,
                                        cfg As Conf) As String
        'If cfg Is Nothing Then cfg = New Conf()
        Dim resultado As String = ""
        '
        If cfg._connectionString <> _strCadenaC Then cfg._connectionString = _strCadenaC
        '
        Dim DAdapter As SqlDataAdapter = Nothing
        Dim CBuilder As SqlCommandBuilder = Nothing
        Dim DSet As DataSet = Nothing
        Dim Cmd As SqlCommand = Nothing
        '
        Using Cnn As New SqlConnection(cfg._connectionString)
            Try
                Cnn.Open()
                AddHandler Cnn.InfoMessage, AddressOf Cnn_InfoMessage
                ' Selecctionar todas las tablas de una BD SQL
                ' select * from INFORMATION_SCHEMA.TABLES
                DAdapter = New SqlDataAdapter("select * from INFORMATION_SCHEMA.TABLES", Cnn)
                DSet = New DataSet
                DAdapter.Fill(DSet, "TABLAS")
                RemoveHandler Cnn.InfoMessage, AddressOf Cnn_InfoMessage
                Cnn.Close()
                resultado = ""
            Catch ex As Exception
                resultado = ex.ToString
            Finally
                If Cnn IsNot Nothing AndAlso Cnn.State <> ConnectionState.Closed Then Cnn.Close()
                DAdapter = Nothing
                CBuilder = Nothing
                DSet = Nothing
                Cmd = Nothing
                GC.Collect()
            End Try
        End Using
        Return resultado
    End Function

    'Ejemplo de Uso: DataSet_FillGet("SELECT * FROM Pedidos_Clientes", "DATOS")
    ''' <summary>Rellenamos el Dataset "resultado" con una consulta en la tabla indicada</summary>
    ''' <param name="sqlstring">Cadena de consulta (Ej: "SELECT CLI_PROVEE FROM CLIENTES WHERE CLI_PROVEE='55555'")</param>
    ''' <param name="ntabla">Tabla a generar</param>
    ''' <remarks>Rellenar dataset "resultado" con una tabla a partir de datos</remarks>
    Public Shared Function DataSet_FillGet(ByVal sqlstring As String, ByVal ntabla As String,
                                        cfg As Conf) As DataSet
        'If cfg Is Nothing Then cfg = New Conf()
        Dim Cnn As SqlConnection = Nothing
        Dim DAdapter As SqlDataAdapter = Nothing
        Dim CBuilder As SqlCommandBuilder = Nothing
        Dim Cmd As SqlCommand = Nothing

        Dim resultado As DataSet = Nothing

        Cnn = New SqlConnection(cfg._connectionString)
        Cmd = New SqlCommand(sqlstring, Cnn)
        DAdapter = New SqlDataAdapter(Cmd)
        Try
            resultado = New DataSet
            DAdapter.Fill(resultado, ntabla)
        Catch Exc As Exception
            MessageBox.Show(Exc.Message, "Pantalla ERROR", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        Finally
            If Cnn IsNot Nothing AndAlso Cnn.State <> ConnectionState.Closed Then Cnn.Close()
            Cnn = Nothing
            DAdapter = Nothing
            CBuilder = Nothing
            Cmd = Nothing
            GC.Collect()
        End Try
        Return resultado
    End Function

    '''<summary>Busca un registro para ver si existe y devuelve Primer campo o "" si no existe</summary>
    '''<param name="cadenaSQL">Cadena SQL que vamos a lanzar</param>
    Public Shared Function Row_Exist(ByVal cadenaSQL As String,
                                        cfg As Conf) As String
        'If cfg Is Nothing Then cfg = New Conf()
        'Dim Cnn As SqlConnection = Nothing
        Dim DAdapter As SqlDataAdapter = Nothing
        Dim CBuilder As SqlCommandBuilder = Nothing
        Dim DSet As DataSet = Nothing
        Dim Cmd As SqlCommand = Nothing

        'Buscaremos si existe el registro con el valor (valor) en la Base de Datos (tabla)
        ' devolviendo "" si no existe o el primer campo CLI_CLAVE (Clientes) 
        'Creamos la conexión a la Base de Datos SQL
        Using Cnn As New SqlConnection(cfg._connectionString)
            Dim resultado As String = ""
            Try
                'Abrimos la conexión. Aunque en teoría no haría falta.
                Cnn.Open()
                Cmd = New SqlCommand
                Cmd.Connection = Cnn
                Cmd.CommandText = cadenaSQL
                resultado = Cmd.ExecuteScalar
                If resultado = Nothing Then resultado = ""
                Cnn.Close()
                Return resultado.ToString
            Catch Exc As Exception
                MessageBox.Show(Exc.Message, "Pantalla ERROR ExisteRegistro", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Return ""
            Finally
                If Cnn IsNot Nothing AndAlso Cnn.State <> ConnectionState.Closed Then Cnn.Close()
                DAdapter = Nothing
                CBuilder = Nothing
                DSet = Nothing
                Cmd = Nothing
                GC.Collect()
            End Try
        End Using
    End Function


    Public Shared Function Tables_Read(cfg As Conf) As DataTable
        'If cfg Is Nothing Then cfg = New Conf()
        Dim DAdapter As SqlDataAdapter = Nothing
        Dim CBuilder As SqlCommandBuilder = Nothing
        Dim DSet As DataSet = Nothing
        Dim Cmd As SqlCommand = Nothing

        Dim resultado As DataTable = Nothing

        Using Cnn As New SqlConnection(cfg._connectionString)
            Try
                Cnn.Open()
                DSet = New DataSet
                DAdapter = New SqlDataAdapter("SELECT * FROM master.dbo.sysdatabases", Cnn)
                DAdapter.Fill(DSet, "BDs")
                resultado = DSet.Tables("BDs")
                Cnn.Close()
            Catch ex As Exception
                'MessageBox.Show("Imposible conectar con servidor ( " & servidor & " ) - SQL... ERROR" & vbCrLf & ex.Message)
                resultado = Nothing
            Finally
                If Cnn IsNot Nothing AndAlso Cnn.State <> ConnectionState.Closed Then Cnn.Close()
                DAdapter = Nothing
                CBuilder = Nothing
                DSet = Nothing
                Cmd = Nothing
                GC.Collect()
            End Try
        End Using
        '
        Return resultado
    End Function

    Public Shared Sub Combobox_FillTables(ByVal cb As ComboBox,
                                        cfg As Conf)
        'If cfg Is Nothing Then cfg = New Conf()
        Dim Cnn As SqlConnection = Nothing
        Dim DAdapter As SqlDataAdapter = Nothing
        Dim CBuilder As SqlCommandBuilder = Nothing
        Dim DSet As DataSet = Nothing
        Dim Cmd As SqlCommand = Nothing

        Dim dataT As DataTable = Tables_Read(cfg)

        If dataT IsNot Nothing Then
            cb.DataSource = dataT
            cb.DisplayMember = "NAME"
        End If
    End Sub


    Private Sub Table_InsertValues(ByVal arrValues() As Object,
                                        cfg As Conf)
        'If cfg Is Nothing Then cfg = New Conf()
        Dim Cnn As SqlConnection = Nothing
        Dim DAdapter As SqlDataAdapter = Nothing
        Dim CBuilder As SqlCommandBuilder = Nothing
        Dim DSet As DataSet = Nothing
        Dim Cmd As SqlCommand = Nothing


        'Creamos la conexión a la Base de Datos SQL
        'Dim cnn As New OleDbConnection(My.Settings.SERICADconexion)
        Dim resultado As Integer = 0
        Dim tabla As String = "Proyectos"   '"Proyectos"   '"Componentes"  '"Armados"
        'SELECT  PRO_CLAVE, PRO_NOMBRE, PRO_DESCRIP, PRO_FECHA, PRO_ENSAMB, PRO_CAMINO, PRO_IMAGEN
        'FROM dbo.Proyectos
        Dim strIns As String = "INSERT INTO " & tabla & " (" &
        "PRO_CLAVE, PRO_NOMBRE, PRO_DESCRIP, PRO_FECHA, PRO_ENSAMB, PRO_CAMINO, PRO_IMAGEN) " &
        "VALUES " &
        "(@CLAVE, @NOMBRE, @DESCRIP, @FECHA, @ENSAMB, @CAMINO, @IMAGEN)"

        Cmd = New SqlCommand(strIns, Cnn)
        Call Cmd.Parameters.AddWithValue("@CLAVE", arrValues(0))
        Call Cmd.Parameters.AddWithValue("@NOMBRE", arrValues(1))
        Call Cmd.Parameters.AddWithValue("@DESCRIP", arrValues(2))
        Call Cmd.Parameters.AddWithValue("@FECHA", arrValues(3))
        Call Cmd.Parameters.AddWithValue("@ENSAMB", arrValues(4))
        Call Cmd.Parameters.AddWithValue("@CAMINO", arrValues(5))
        Call Cmd.Parameters.AddWithValue("@IMAGEN", arrValues(6))

        Try
            'Abrimos la conexión. Aunque en teoría no haría falta.
            Cnn.Open()
            Dim t As Integer = CInt(Cmd.ExecuteScalar)
            Cnn.Close()
        Catch Exc As Exception
            MessageBox.Show("ERROR insertando datos..." & vbCrLf & Exc.Message, "Pantalla ERRORES", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        Finally
            If Cnn IsNot Nothing AndAlso Cnn.State <> ConnectionState.Closed Then Cnn.Close()
            Cnn = Nothing
            DAdapter = Nothing
            CBuilder = Nothing
            DSet = Nothing
            Cmd = Nothing
            GC.Collect()
        End Try
    End Sub

    Public Shared Sub Image_Insert(ByVal table As String, cId As String, ByVal cIdValor As String, CampoImagen As String, ByVal imagen() As Byte,
                                        cfg As Conf)
        'If cfg Is Nothing Then cfg = New Conf()
        Dim DAdapter As SqlDataAdapter = Nothing
        Dim CBuilder As SqlCommandBuilder = Nothing
        Dim DSet As DataSet = Nothing
        Dim Cmd As SqlCommand = Nothing

        Using Cnn As New SqlConnection(cfg._connectionString)
            DAdapter = New SqlDataAdapter("select * from " & table & " where " & cId & " = '" & cIdValor & "'", Cnn)
            CBuilder = New SqlCommandBuilder(DAdapter)
            'CBuilder.GetUpdateCommand()
            DAdapter.MissingSchemaAction = MissingSchemaAction.AddWithKey
            DSet = New DataSet
            Try
                Cnn.Open()
                DAdapter.Fill(DSet, "fila")
                'Llenamos el DataRow con la fila primera (la única que hay)
                Dim fila As DataRow = DSet.Tables("fila").Rows.Item(0)
                'Cargamos la imagen en el campo PRO_IMAGEN y actualizamos el Datase a través del DataAdapter.
                fila.Item(CampoImagen) = imagen
                'Obtenemos el comando para actualizar la tabla del CommanBuilder (cb)
                CBuilder.GetUpdateCommand()
                DAdapter.Update(DSet, "fila")
                Cnn.Close()
            Catch Exc As Exception
                MessageBox.Show("ERROR ImagenInsertar..." & vbCrLf & Exc.Message, "Pantalla ERRORES", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Finally
                If Cnn IsNot Nothing AndAlso Cnn.State <> ConnectionState.Closed Then Cnn.Close()
                DAdapter = Nothing
                CBuilder = Nothing
                DSet = Nothing
                Cmd = Nothing
                GC.Collect()
            End Try
        End Using
    End Sub

    Public Shared Function Image_Read(ByVal table As String, cId As String, ByVal cIdValor As String, CampoImagen As String,
                                        cfg As Conf) As System.Drawing.Image
        'If cfg Is Nothing Then cfg = New Conf()
        Dim DAdapter As SqlDataAdapter = Nothing
        Dim CBuilder As SqlCommandBuilder = Nothing
        Dim DSet As DataSet = Nothing
        Dim Cmd As SqlCommand = Nothing

        Dim resultado(-1) As Byte
        Using Cnn As New SqlConnection(cfg._connectionString)
            Try
                'Creamos la conexión a la Base de Datos SQL y seleccionamos a una tabla
                'cnn = New SqlConnection(strCadenaC)
                DAdapter = New SqlDataAdapter("select * from " & table & " where " & cId & " = '" & cIdValor & "'", Cnn)
                'Creamos el DataSet "ds" y lo llenamos con la tabla "fila" que sólo tiene una fila
                Dim ds As New DataSet
                Cnn.Open()
                DAdapter.Fill(ds, "fila")
                'Llenamos el DataRow con la fila primera (la única que hay)
                Dim fila As DataRow = ds.Tables("fila").Rows(0)
                'Cargamos la imagen en el campo PRO_IMAGEN y actualizamos el Datase a través del DataAdapter.
                If IsDBNull(fila.Item(CampoImagen)) = False Then resultado = fila.Item(CampoImagen)
                Cnn.Close()
            Catch Exc As Exception
                MessageBox.Show("ERROR ImagenLeer..." & vbCrLf & Exc.Message, "Pantalla ERRORES", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Finally
                If Cnn IsNot Nothing AndAlso Cnn.State <> ConnectionState.Closed Then Cnn.Close()
                DAdapter = Nothing
                CBuilder = Nothing
                DSet = Nothing
                Cmd = Nothing
                GC.Collect()
            End Try
        End Using
        '
        If resultado Is Nothing Then
            Return Nothing  ' My.Resources.SinImagen
        Else
            Return Image_ByteArrayToImage(resultado)
        End If
    End Function

    Public Shared Function Row_Read(ByVal table As String, command As String,
                                        cfg As Conf) As DataRow
        'If cfg Is Nothing Then cfg = New Conf()
        Dim DAdapter As SqlDataAdapter = Nothing
        Dim CBuilder As SqlCommandBuilder = Nothing
        Dim DSet As DataSet = Nothing
        Dim Cmd As SqlCommand = Nothing

        Dim resultado As DataRow = Nothing

        Using Cnn As New SqlConnection(cfg._connectionString)
            DAdapter = New SqlDataAdapter(command, Cnn)
            DSet = New DataSet
            '
            Try
                Cnn.Open()
                DAdapter.Fill(DSet, "fila")
                resultado = DSet.Tables("fila").Rows(0)
                Cnn.Close()
            Catch Exc As Exception
                resultado = Nothing
            Finally
                If Cnn IsNot Nothing AndAlso Cnn.State <> ConnectionState.Closed Then Cnn.Close()
                DAdapter = Nothing
                CBuilder = Nothing
                DSet = Nothing
                Cmd = Nothing
                GC.Collect()
            End Try
        End Using
        '
        Return resultado
    End Function

    ' Actualiza el valor de un campo. Según un criterio
    Public Shared Function Row_Update(table As String, cId As String, cIdValor As Object, cName As String, cValue As Object,
                                        cfg As Conf) As Integer
        'If cfg Is Nothing Then cfg = New Conf()
        Dim DAdapter As SqlDataAdapter = Nothing
        Dim CBuilder As SqlCommandBuilder = Nothing
        Dim DSet As DataSet = Nothing
        Dim Cmd As SqlCommand = Nothing
        '
        Dim resultado As Integer = 0
        Using Cnn As New SqlConnection(cfg._connectionString)
            Cmd = New SqlCommand
            Cmd.Connection = Cnn
            Cmd.CommandText =
                "UPDATE " & table &
                " SET " & cName & "= '" & cValue & "'" &
                " WHERE (" & cId & " = '" & cIdValor & "')"
            Try
                Cnn.Open()
                resultado = CInt(Cmd.ExecuteNonQuery)
                Cnn.Close()
            Catch ex As Exception
                resultado = 0
            Finally
                If Cnn IsNot Nothing AndAlso Cnn.State <> ConnectionState.Closed Then Cnn.Close()
                DAdapter = Nothing
                CBuilder = Nothing
                DSet = Nothing
                Cmd = Nothing
                GC.Collect()
            End Try
        End Using
        '
        Return resultado
    End Function
    ' Actualiza el valor de un campo. Según un criterio
    Public Shared Function Row_Update(table As String, cId As String, cIdValor As Object, cNames As String(), cValues As Object(),
                                        cfg As Conf) As Integer
        'If cfg Is Nothing Then cfg = New Conf()
        Dim DAdapter As SqlDataAdapter = Nothing
        Dim CBuilder As SqlCommandBuilder = Nothing
        Dim DSet As DataSet = Nothing
        Dim Cmd As SqlCommand = Nothing
        '
        ' *************** New Array. Union cNames and cValues
        If cNames.Length <> cValues.Length Then
            Return 0
            Exit Function
        End If
        '
        Dim newL As New List(Of String)
        For x As Integer = LBound(cNames) To UBound(cNames)
            newL.Add(cNames(x) & " = '" & cValues(x).ToString & "'")
        Next
        ' ********************************
        Dim resultado As Integer = 0
        Using Cnn As New SqlConnection(cfg._connectionString)
            Cmd = New SqlCommand
            Cmd.Connection = Cnn
            Cmd.CommandText =
                "UPDATE " & table &
                " SET " & String.Join(", ", newL.ToArray) &
                " WHERE (" & cId & " = '" & cIdValor & "')"
            Try
                Cnn.Open()
                resultado = CInt(Cmd.ExecuteNonQuery)
                Cnn.Close()
            Catch ex As Exception
                resultado = 0
            Finally
                If Cnn IsNot Nothing AndAlso Cnn.State <> ConnectionState.Closed Then Cnn.Close()
                DAdapter = Nothing
                CBuilder = Nothing
                DSet = Nothing
                Cmd = Nothing
                GC.Collect()
            End Try
        End Using
        '
        Return resultado
    End Function

    Private Sub FilaInserta(ByVal tabla As String, ByVal campos() As String, ByVal valores() As Object,
                                        cfg As Conf)
        'If cfg Is Nothing Then cfg = New Conf()
        Dim Cnn As SqlConnection = Nothing
        Dim DAdapter As SqlDataAdapter = Nothing
        Dim CBuilder As SqlCommandBuilder = Nothing
        Dim DSet As DataSet = Nothing
        Dim Cmd As SqlCommand = Nothing

        'Campo de imagen que vamos a actualizar (3 primeros char de la tabla + _IMAGEN"
        Dim CampoClave As String = tabla.Substring(0, 3).ToUpper & "_CLAVE"   '"Proyectos"   '"Componentes"  '"Armados"

        'Creamos la conexión a la Base de Datos SQL y seleccionamos a una tabla
        Cnn = New SqlConnection(cfg._connectionString)
        DAdapter = New SqlDataAdapter("select * from " & tabla, Cnn)
        'Creamos el commandbuilder para que nos genere todos los comandos: select, insert y update.
        CBuilder = New SqlCommandBuilder(DAdapter)
        'da.UpdateCommand
        DAdapter.MissingSchemaAction = MissingSchemaAction.AddWithKey
        'Creamos el DataSet "ds" y lo llenamos con la tabla "fila" que sólo tiene una fila
        DSet = New DataSet
        Try
            Cnn.Open()
            DAdapter.Fill(DSet, "fila")
            '' Creamos una fila nueva con los campos de la tabla
            Dim fila As DataRow = DSet.Tables("fila").NewRow
            '' Aquí pondremos valores a los campos
            For x = 0 To campos.Length - 1
                'If valores(x) <> "" Then
                fila.Item(campos(x)) = valores(x)
                'Else
                'fila.Item(campos(x)) = DBNull.Value
                'End If
            Next
            '*********************************************************************
            '' Obtenemos el comando para insertar la tabla del CommanBuilder (cb)
            CBuilder.GetInsertCommand()
            DSet.Tables("fila").Rows.Add(fila)
            '' Create a DataView using the DataTable.
            'View = New DataView(table)
            '' Set a DataGrid control's DataSource to the DataView.
            'DataGrid1.DataSource = View
            DAdapter.Update(DSet, "fila")
            Cnn.Close()
        Catch Exc As Exception
            MessageBox.Show("ERROR FILAINSERTAR..." & vbCrLf & Exc.Message, "Pantalla ERRORES", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        Finally
            If Cnn IsNot Nothing AndAlso Cnn.State <> ConnectionState.Closed Then Cnn.Close()
            Cnn = Nothing
            DAdapter = Nothing
            CBuilder = Nothing
            DSet = Nothing
            Cmd = Nothing
            GC.Collect()
        End Try
    End Sub

    Private Sub FilaActualiza(ByVal tabla As String, ByVal fila As DataRow, ByVal campoActualiza As String,
                                        cfg As Conf)
        'If cfg Is Nothing Then cfg = New Conf()
        Dim Cnn As SqlConnection = Nothing
        Dim DAdapter As SqlDataAdapter = Nothing
        Dim CBuilder As SqlCommandBuilder = Nothing
        Dim DSet As DataSet = Nothing
        Dim Cmd As SqlCommand = Nothing

        'Campo de imagen que vamos a actualizar (3 primeros char de la tabla + _IMAGEN"
        Dim CampoClave As String = tabla.Substring(0, 3).ToUpper & "_CLAVE"   '"Proyectos"   '"Componentes"  '"Armados"

        'Creamos la conexión a la Base de Datos SQL y seleccionamos a una tabla
        If Cnn Is Nothing Then Cnn = New SqlConnection(cfg._connectionString)
        DAdapter = New SqlDataAdapter("select * from " & tabla & " where " & CampoClave & " = '" & Trim(fila(CampoClave)) & "'", Cnn)
        'Creamos el commandbuilder para que nos genere todos los comandos: select, insert y update.
        CBuilder = New SqlCommandBuilder(DAdapter)
        'da.UpdateCommand
        DAdapter.MissingSchemaAction = MissingSchemaAction.AddWithKey
        'Creamos el DataSet "ds" y lo llenamos con la tabla "fila" que sólo tiene una fila
        Dim ds As New DataSet
        Try
            Cnn.Open()
            DAdapter.Fill(DSet, "filas")
            '' Obtenemos el comando para insertar la tabla del CommanBuilder (cb)
            'Call cb.GetUpdateCommand()
            ds.Tables("filas").Rows(0).Item(campoActualiza) = fila(campoActualiza)
            ds.AcceptChanges()
            'da.Update(ds, "filas")
            Cnn.Close()
        Catch Exc As Exception
            MessageBox.Show("ERROR FilaActualiza..." & vbCrLf & Exc.Message, "Pantalla ERRORES", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        Finally
            If Cnn IsNot Nothing AndAlso Cnn.State <> ConnectionState.Closed Then Cnn.Close()
            Cnn = Nothing
            DAdapter = Nothing
            CBuilder = Nothing
            DSet = Nothing
            Cmd = Nothing
            GC.Collect()
        End Try
    End Sub

    ''' <summary>
    ''' Devuelve una Tabla con los registros que cumplen el criterio especificado
    ''' </summary>
    ''' <param name="tabla">Tabla en la que buscar los registros</param>
    ''' <param name="campo">Campo en el que buscar</param>
    ''' <param name="valor">Valor buscado en el campo</param>
    ''' <returns>Devuelve un objeto DataTable</returns>
    ''' <remarks></remarks>
    Private Function FilasLeer(ByVal tabla As String,
                              ByVal campo As String, ByVal valor As Object,
                                        cfg As Conf) As DataTable
        'If cfg Is Nothing Then cfg = New Conf()
        Dim Cnn As SqlConnection = Nothing
        Dim DAdapter As SqlDataAdapter = Nothing
        Dim CBuilder As SqlCommandBuilder = Nothing
        Dim DSet As DataSet = Nothing
        Dim Cmd As SqlCommand = Nothing

        Dim dt As DataTable = Nothing

        'Creamos la conexión a la Base de Datos SQL y seleccionamos una tabla
        Cnn = New SqlConnection(cfg._connectionString)
        DAdapter = New SqlDataAdapter("select * from " & tabla & " where " & campo & " = '" & valor & "'", Cnn)
        'Creamos el DataSet "ds" y lo llenamos con la tabla "registro" que sólo tiene una fila
        DSet = New DataSet
        Try
            Cnn.Open()
            DAdapter.Fill(DSet, "filas")
            'Llenamos el DataRow con la fila primera (la única que hay)
            dt = DSet.Tables("filas")
            Cnn.Close()
        Catch Exc As Exception
            MessageBox.Show("ERROR FilasLeer..." & vbCrLf & Exc.Message, "Pantalla ERRORES", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        Finally
            If Cnn IsNot Nothing AndAlso Cnn.State <> ConnectionState.Closed Then Cnn.Close()
            Cnn = Nothing
            DAdapter = Nothing
            CBuilder = Nothing
            DSet = Nothing
            Cmd = Nothing
            GC.Collect()
        End Try
        Return dt
    End Function

    ''' <summary>
    ''' Devuelve un array de cadenas con los nombres de todas las columnas de una Tabla
    ''' </summary>
    ''' <param name="tabla">Tabla en la que buscar los registros</param>
    ''' <returns>Devuelve un array de string con los nombres de los campos</returns>
    ''' <remarks></remarks>
    Private Function ColumnasNombreLeer(ByVal tabla As String,
                                        cfg As Conf) As String()
        'If cfg Is Nothing Then cfg = New Conf()
        Dim Cnn As SqlConnection = Nothing
        Dim DAdapter As SqlDataAdapter = Nothing
        Dim CBuilder As SqlCommandBuilder = Nothing
        Dim DSet As DataSet = Nothing
        Dim Cmd As SqlCommand = Nothing

        Dim resultado() As String = Nothing
        Dim dt As DataTable = Nothing

        'Creamos la conexión a la Base de Datos SQL y seleccionamos una tabla
        Cnn = New SqlConnection(cfg._connectionString)
        DAdapter = New SqlDataAdapter("select top 1 * from " & tabla, Cnn)
        'Creamos el DataSet "ds" y lo llenamos con la tabla "registro" que sólo tiene una fila
        DSet = New DataSet
        Try
            Cnn.Open()
            DAdapter.Fill(DSet, "fila1")
            'Llenamos el DataRow con la fila primera (la única que hay)
            dt = DSet.Tables("fila1")
            Dim arrCampos As New ArrayList
            For Each campo As DataColumn In dt.Columns
                arrCampos.Add(campo.ColumnName)
            Next
            resultado = arrCampos.ToArray
            Cnn.Close()
        Catch Exc As Exception
            MessageBox.Show("ERROR FilasLeer..." & vbCrLf & Exc.Message, "Pantalla ERRORES", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        Finally
            If Cnn IsNot Nothing AndAlso Cnn.State <> ConnectionState.Closed Then Cnn.Close()
            Cnn = Nothing
            DAdapter = Nothing
            CBuilder = Nothing
            DSet = Nothing
            Cmd = Nothing
            GC.Collect()
        End Try
        Return resultado
    End Function

    '' Nos dará el total de registros de una tabla.
    Private Function DameTotalRegistros(ByVal tabla As String,
                                        cfg As Conf) As Integer
        'If cfg Is Nothing Then cfg = New Conf()
        Dim Cnn As SqlConnection = Nothing
        Dim DAdapter As SqlDataAdapter = Nothing
        Dim CBuilder As SqlCommandBuilder = Nothing
        Dim DSet As DataSet = Nothing
        Dim Cmd As SqlCommand = Nothing

        Dim resultado As Integer = 0
        'Dim cnn As New OleDbConnection
        Cnn = New SqlConnection(cfg._connectionString)
        Dim sqlCom As New SqlCommand    ' OleDbCommand
        sqlCom.Connection = Cnn
        sqlCom.CommandText = "select count(*) from " & tabla
        Try
            Cnn.Open()
            resultado = CInt(sqlCom.ExecuteScalar)
            Cnn.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        Finally
            If Cnn IsNot Nothing AndAlso Cnn.State <> ConnectionState.Closed Then Cnn.Close()
            Cnn = Nothing
            DAdapter = Nothing
            CBuilder = Nothing
            DSet = Nothing
            Cmd = Nothing
            GC.Collect()
        End Try
        Return resultado
        'Exit Function
    End Function
#Region "EVENTS"
    Private Shared Sub Cnn_InfoMessage(ByVal sender As Object, ByVal e As System.Data.SqlClient.SqlInfoMessageEventArgs)
        Dim mensaje As String = ""
        mensaje &= "El objeto " & sender.GetType.ToString & " da el siguiente aviso :" & vbCrLf & vbCrLf & vbCrLf
        For Each err As SqlError In e.Errors
            mensaje &= err.Number & vbTab & err.Procedure & vbTab & err.Message & vbCrLf
        Next
        MsgBox(mensaje)
    End Sub
#End Region
End Class
