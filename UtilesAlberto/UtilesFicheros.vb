Imports System.IO
Imports System.Text

Partial Public Class Utiles
    Public Shared oPermissionAll As System.Security.Permissions.FileIOPermission

    Public Shared Function FolderFile_SetPremissionAll(path As String) As Boolean
        Dim resultado As Boolean = False
        Try
            ' Create PermisosFi con permiso para todos los ficheros (Red y Locales)
            AllFiles_AllAccess()
            If oPermissionAll IsNot Nothing Then
                oPermissionAll.AddPathList(oPermissionAll.AllFiles, path)
            End If
        Catch ex As Exception
            Debug.Print(ex.ToString)
        End Try
        '
        Return resultado
    End Function
    Private Shared Sub AllFiles_AllAccess()
        oPermissionAll = New System.Security.Permissions.FileIOPermission(Security.Permissions.PermissionState.Unrestricted)
        oPermissionAll.AllFiles = Security.Permissions.FileIOPermissionAccess.AllAccess
        Try
            oPermissionAll.Demand()
            'Permisos.SetPathList(Security.Permissions.FileIOPermissionAccess.AllAccess, nDir)
        Catch s As System.Security.SecurityException
            Debug.Print(s.ToString)
        End Try
    End Sub
    Private Shared Sub AllLocalFiles_AllAccess()
        Dim PermisosFi As New System.Security.Permissions.FileIOPermission(Security.Permissions.PermissionState.Unrestricted)
        PermisosFi.AllLocalFiles = Security.Permissions.FileIOPermissionAccess.AllAccess
        Try
            PermisosFi.Demand()
            'Permisos.SetPathList(Security.Permissions.FileIOPermissionAccess.AllAccess, nDir)
        Catch s As System.Security.SecurityException
            Debug.Print(s.ToString)
        End Try
    End Sub
    Public Shared Sub Folder_ReadOnly(pathFo As String, rOnly As Boolean,
                                      Optional mascara As String = "*.*",
                                      Optional recursive As Boolean = False)
        If IO.Directory.Exists(pathFo) = False Then Exit Sub
        ''
        For Each queF As String In IO.Directory.GetFiles(pathFo, mascara, IIf(recursive, IO.SearchOption.AllDirectories, IO.SearchOption.TopDirectoryOnly))
            File_ReadOnly(queF, rOnly)
        Next
        Dim dInfo As New System.IO.DirectoryInfo(pathFo)
        If rOnly Then
            'dInfo.Attributes = dInfo.Attributes Or FileAttributes.ReadOnly
            dInfo.Attributes = Attribute_Add(dInfo.Attributes, FileAttributes.ReadOnly)
        Else
            dInfo.Attributes = Attribute_Remove(dInfo.Attributes, FileAttributes.ReadOnly)
        End If
    End Sub
    '
    Public Shared Sub File_ReadOnly(pathFi As String, rOnly As Boolean)
        If IO.File.Exists(pathFi) = False Then Exit Sub
        '
        Dim fInfo As New System.IO.FileInfo(pathFi)
        Dim attributes As FileAttributes = File.GetAttributes(pathFi)
        If ((attributes And FileAttributes.ReadOnly) = FileAttributes.ReadOnly) Then
            'Console.WriteLine("read-only file")
            If rOnly = False Then fInfo.IsReadOnly = False
        Else
            'Console.WriteLine("not read-only file")
            If rOnly = True Then fInfo.IsReadOnly = True
        End If
    End Sub
    '
    Public Shared Function Attribute_Add(ByVal attributes As FileAttributes, ByVal attributesToAdd As FileAttributes) As FileAttributes
        Return attributes Or attributesToAdd
    End Function

    Public Shared Function Attribute_Remove(ByVal attributes As FileAttributes, ByVal attributesToRemove As FileAttributes) As FileAttributes
        Return attributes And (Not attributesToRemove)
    End Function

    '' Devolverá la misma carpeta (Si es local) O la carpeta de red (\\servidor\carpeta\, si es de red)
    Public Shared Function UNCPath_Get(ByVal sFilePath As String) As String
        Dim resultado As String = ""
        Dim allDrives() As DriveInfo = DriveInfo.GetDrives()
        Dim d As DriveInfo
        Dim DriveType, Ctr As Integer
        Dim DriveLtr, UNCName As String
        Dim StrBldr As New StringBuilder

        If sFilePath.StartsWith("\\") Then Return sFilePath

        UNCName = Space(160)
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
                resultado = StrBldr.ToString
            End If
        Else
            resultado = sFilePath
        End If
        Return resultado
    End Function
    '
    Public Shared Function Path_IsValid(ByVal path As String) As Boolean
        Dim resultado As Boolean = True
        '' Comprobar primero los carácteres invalidos del directorio.
        For Each oChar As Char In IO.Path.GetInvalidPathChars
            If path.Contains(oChar) Then
                resultado = False
                Exit For
            End If
        Next
        '' Comprobar la longitud de la cadena.
        If path.Length > 255 Then resultado = False
        '
        Return resultado
    End Function

    Public Shared Function FileName_PutCorrectChars(ByVal queTexto As String) As String
        Dim resultado As String = queTexto

        Dim malos() As Char = New Char() _
                              {"/", "\", "<", ">", "[", "]", "{", "}", "¿", "?", "¡", "!", ":", "|", "*", "."}
        For Each queChar As Char In malos
            If queTexto.Contains(queChar) Then
                resultado = resultado.Replace(queChar, "_")
            End If
        Next

        Return resultado
    End Function

    Public Shared Function FileName_IsCorrectName(ByVal queTexto As String) As Boolean
        Dim resultado As Boolean = False

        Dim malos() As Char = New Char() _
                              {"/", "\", "<", ">", "[", "]", "{", "}", "¿", "?", "¡", "!", ":", "|", "*", "."}
        For Each queChar As Char In malos
            If queTexto.Contains(queChar) Then
                resultado = True
                Exit For
            End If
        Next
        Return resultado
    End Function

    ''' <summary>
    ''' Si le damos una cadena completa (unidad:\directorio\fichero.extension) nos devuelve la parte que le indiquemos.
    ''' </summary>
    ''' <param name="queCamino">Cadena completa con el camino a procesar DIR+FICHERO+EXT</param>
    ''' <param name="queParte">Que queremos que nos devuelva</param>
    ''' <param name="queExtension">"" o extensión (Ej: ".bak"), si queremos cambiarla</param>
    ''' <returns>Retorna la cadena de texto con la opción indicada</returns>
    ''' <remarks></remarks>
    Public Shared Function DameParteCamino(ByVal queCamino As String,
                                    Optional ByVal queParte As ParteCamino = 0,
                                    Optional ByVal queExtension As String = "") As String
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
                ' Directorio o Fichero
                If IO.File.Exists(queCamino) Then
                    resultado = trozos(trozos.GetUpperBound(0) - 1)
                ElseIf IO.Directory.Exists(queCamino) Then
                    resultado = trozos(trozos.GetUpperBound(0))
                End If
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
                    Dim final(trozos.GetUpperBound(0) - 1) As String
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
    ''' Buscar todos los fichero dentro en directorio (opción subdirectorios y máscara con extensión) 
    ''' </summary>
    ''' <param name="folder">Search start folder</param>
    ''' <param name="SearchOptions">Tipo de búsqueda (Dir indicado solo o También SubDir)</param>
    ''' <param name="extension">Extensión opcional a buscar (Por defecto Todos *.* si no ponemos nada)</param>
    ''' <returns>Devuelve un colección con todos los nombres completos encontrado, que habrá que recorrer</returns>
    ''' <remarks></remarks>
    Public Shared Function Files_ListInFolder(ByVal folder As String, ByVal SearchOptions As IO.SearchOption,
                                  Optional ByVal extension As String = "*.*",
                                  Optional message As Boolean = False) As List(Of String)
        Dim colLista As New List(Of String)

        For Each foundF As String In My.Computer.FileSystem.GetFiles(
        folder, SearchOptions, extension)
            If colLista.Contains(foundF) = False Then colLista.Add(foundF)
        Next
        If message = True Then
            MsgBox(String.Join(vbCrLf, colLista.ToArray))
        End If
        Return colLista
    End Function
    Public Shared Sub Files_Delete(ByVal folder As String, Optional folderdelete As Boolean = False)
        If IO.Directory.Exists(folder) = False Then
            IO.Directory.CreateDirectory(folder)
            Exit Sub
        End If

        For Each fichero As String In IO.Directory.GetFiles(folder, "*.*", IO.SearchOption.AllDirectories)
            Try
                IO.File.Delete(fichero)
            Catch ex As Exception

            End Try
        Next

        If folderdelete = True Then
            Try
                IO.Directory.Delete(folder, True)
            Catch ex As Exception

            End Try
        End If
    End Sub
    Public Shared Function FicheroEstaAbierto(filePath As String) As Boolean
        Dim rtnvalue As Boolean = False
        Try
            Dim fs As System.IO.FileStream = System.IO.File.OpenWrite(filePath)
            fs.Close()
        Catch ex As System.IO.IOException
            rtnvalue = True
        End Try
        Return rtnvalue
    End Function

    Public Shared Function FicheroEstaAbiertoMensaje(filePath As String, nApp As String) As MsgBoxResult
REPETIR:
        If FicheroEstaAbierto(filePath) = False Then
            Return MsgBoxResult.Ok
        End If
        Dim resultado As MsgBoxResult = MsgBoxResult.Ok
        '
        'If FicheroEstaAbierto(filePath) Then
        If MsgBox("El fichero " & filePath & " Esta abierto." & vbCrLf & vbCrLf &
                          "- Cerrarlo y Reintentar o Cancelar (No cargara " & nApp & ")", MsgBoxStyle.RetryCancel) = MsgBoxResult.Retry Then
            GoTo REPETIR
        Else
            Return MsgBoxResult.Cancel
            End If
        'End If
        Return resultado
    End Function
End Class
