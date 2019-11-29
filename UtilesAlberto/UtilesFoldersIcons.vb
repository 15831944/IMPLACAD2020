Partial Public Class Utiles
    '---------------------------------------------------------------------------
    '---------------------------------------------------------------------------
    ' -- Descripción        : Simple módulo de código para obtener Items de windows (Panel de control, mi pc etc..) y carpetas especiales junto con las imagenes e íconos
    ' -- Autor              : Luciano Lodola -- http://www.recursosvisualbasic.com.ar/
    '---------------------------------------------------------------------------
    '---------------------------------------------------------------------------
    ' -- Enum para Carpetas / Directorios especiales de Windows
    '---------------------------------------------------------------------------
    Const MAX_PATH = 260

    Enum ESPECIAL_FOLDERS
        e_ADMINTOOLS = &H30
        e_ALTSTARTUP = &H1D
        e_APPDATA = &H1A
        e_BITBUCKET = &HA
        e_COMMON_ADMINTOOLS = &H2F
        e_COMMON_ALTSTARTUP = &H1E
        e_COMMON_APPDATA = &H23
        e_COMMON_DESKTOPDIRECTORY = &H19
        e_COMMON_DOCUMENTS = &H2E
        e_COMMON_FAVORITES = &H1F
        e_COMMON_PROGRAMS = &H17
        e_COMMON_STARTMENU = &H16
        e_COMMON_STARTUP = &H18
        e_COMMON_TEMPLATES = &H2D
        e_CONNECTIONS = &H31
        e_CONTROLS = &H3
        e_COOKIES = &H21
        e_DESKTOP = &H0
        e_DESKTOPDIRECTORY = &H10
        e_DRIVES = &H11
        e_FAVORITES = &H6
        e_FLAG_DONT_VERIFY = &H4000
        e_FLAG_MASK = &HFF00&
        e_FLAG_PFTI_TRACKTARGET = e_FLAG_DONT_VERIFY
        e_FONTS = &H14
        e_INTERNET = &H1
        e_HISTORY = &H22
        e_INTERNET_CACHE = &H20
        e_LOCAL_APPDATA = &H1C
        e_MYPICTURES = &H27
        e_NETHOOD = &H13
        e_NETWORK = &H12
        e_PERSONAL = &H5
        e_PRINTERS = &H4
        e_PRINTHOOD = &H1B
        e_PROFILE = &H28
        e_PROGRAM_FILES = &H26
        e_PROGRAM_FILES_COMMON = &H2B
        e_PROGRAM_FILES_COMMONX86 = &H2C
        e_PROGRAM_FILESX86 = &H2A
        e_PROGRAMS = &H2
        e_RECENT = &H8
        e_SENDTO = &H9
        e_STARTMENU = &HB
        e_STARTUP = &H7
        e_SYSTEM = &H25
        e_SYSTEMX86 = &H29
        e_TEMPLATES = &H15
        e_WINDOWS = &H24
    End Enum

    Enum ESYSTEM_ITEMS
        MI_PC = 0
        PAPELERA = 1
        MIS_SITIOS_DE_RED = 2
        PANEL_DE_CONTROL = 3
        IMPRESORAS_FAX = 4
        CARPETAS_WEB = 5
        HERRAMIENTAS_ADMINISTRATIVAS = 6
        CONEXIONES_DE_RED = 7
        FUENTES = 8
    End Enum

    ' -- Tamaño de íconos de archivos y drives (para la función SHGetFileInfo)
    Public Enum EICON_SIZE
        eSmallIcon = 257
        eLargeIcon = 256
    End Enum
    Public Enum PictureTypeConstants
        vbPicTypeNone = 0
        vbPicTypeBitmap = 1
        vbPicTypeMetafile = 2
        vbPicTypeIcon = 3
        vbPicTypeEMetafile = 4
    End Enum

    ' -- Type para la función api ShellFileInfo
    Private Structure ShellFileInfoType
        Dim hIcon As Long
        Dim iIcon As Long
        Dim dwAttributes As Long
        Dim szDisplayName As String '* 260
        Dim szTypeName As String '* 80
    End Structure

    ' -- Para la función SHGetFileInfo
    Private Structure IconType
        Dim cbSize As Long
        Dim picType As PictureTypeConstants
        Dim hIcon As Long
    End Structure

    ' -- Para la función SHGetFileInfo
    Private Structure CLSIdType
        Dim id() As Byte
    End Structure
    Private Structure SHITEMID
        Dim cb As Long
        Dim abID As Byte
    End Structure

    Private Structure ITEMIDLIST
        Dim mkid As SHITEMID
    End Structure

    ' -- Funciones de windows
    '-------------------------------------------
    ' Función para recuperar las carpetas a partir del ID
    Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
    Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
    ' Función para obtener información varia de un archivo
    Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As ShellFileInfoType, ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long
    Private Declare Function OleCreatePictureIndirect Lib "oleaut32.dll" (pDicDesc As IconType, riid As CLSIdType, ByVal fown As Long, lpUnk As Object) As Long
    ' Función para obtener el nombre de usuario actual
    Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

    ' -- Colecciones, variables
    Private Shared mColSFolderCurrentUser As Collection
    Private Shared mColSFolderAllUsers As Collection
    Private Shared mColSystemPaths As Collection

    ' -------------------------------------------------------------------------------
    '\\ -- Carpetas del usuario actual
    ' -------------------------------------------------------------------------------
    Public Shared Function ColSFolderCurrentUser() As Collection
        FolderIcons_Initialize()
        Return mColSFolderCurrentUser
    End Function
    ' -------------------------------------------------------------------------------
    '\\ -- Carpetas de todos los usuarios
    ' -------------------------------------------------------------------------------
    Public Shared Function ColSFolderAllUsers() As Collection
        FolderIcons_Initialize()
        Return mColSFolderAllUsers
    End Function
    ' -------------------------------------------------------------------------------
    '\\ -- Otras carpetas de windows
    ' -------------------------------------------------------------------------------
    Public Shared Function ColSystemPaths() As Collection
        FolderIcons_Initialize()
        Return mColSystemPaths
    End Function
    ' -------------------------------------------------------------------------------
    '\\ -- Abrir un item del sistema con el comando shell de vb
    ' -------------------------------------------------------------------------------
    Public Shared Sub OpenItem(lOption As ESYSTEM_ITEMS)
        FolderIcons_Initialize()
        If lOption = ESYSTEM_ITEMS.MI_PC Then Shell("explorer " & "::{20D04FE0-3AEA-1069-A2D8-08002B30309D}", vbNormalFocus)
        If lOption = ESYSTEM_ITEMS.PAPELERA Then Shell("explorer " & "::{645FF040-5081-101B-9F08-00AA002F954E}", vbNormalFocus)
        If lOption = ESYSTEM_ITEMS.MIS_SITIOS_DE_RED Then Shell("explorer " & "::{208D2C60-3AEA-1069-A2D7-08002B30309D}", vbNormalFocus)
        If lOption = ESYSTEM_ITEMS.PANEL_DE_CONTROL Then Shell("explorer " & "::{20D04FE0-3AEA-1069-A2D8-08002B30309D}\::{21EC2020-3AEA-1069-A2DD-08002B30309D}", vbNormalFocus)
        If lOption = ESYSTEM_ITEMS.IMPRESORAS_FAX Then Shell("explorer " & "::{20D04FE0-3AEA-1069-A2D8-08002B30309D}\::{2227A280-3AEA-1069-A2DE-08002B30309D}", vbNormalFocus)
        If lOption = ESYSTEM_ITEMS.CARPETAS_WEB Then Shell("explorer " & "::{20D04FE0-3AEA-1069-A2D8-08002B30309D}\::{BDEADF00-C265-11D0-BCED-00A0C90AB50F}", vbNormalFocus)
        If lOption = ESYSTEM_ITEMS.HERRAMIENTAS_ADMINISTRATIVAS Then Shell("explorer " & "::{20D04FE0-3AEA-1069-A2D8-08002B30309D}\::{21EC2020-3AEA-1069-A2DD-08002B30309D}\::{D20EA4E1-3957-11d2-A40B-0C5020524153}", vbNormalFocus)
        If lOption = ESYSTEM_ITEMS.CONEXIONES_DE_RED Then Shell("explorer " & "::{7007ACC7-3202-11D1-AAD2-00805FC1270E}", vbNormalFocus)
        If lOption = ESYSTEM_ITEMS.FUENTES Then Shell("explorer " & "::{20D04FE0-3AEA-1069-A2D8-08002B30309D}\::{21EC2020-3AEA-1069-A2DD-08002B30309D}\::{D20EA4E1-3957-11d2-A40B-0C5020524152}", vbNormalFocus)
    End Sub
    ' -------------------------------------------------------------------------------
    '\\ -- Retornar íconos (16 y 32 pixeles ) de algunos items de windows : Mi pc, red, panel de control ...
    ' -------------------------------------------------------------------------------
    Public Shared Function GetIconSystemItems(lIconSize As EICON_SIZE, lItem As ESYSTEM_ITEMS) As stdole.IPictureDisp
        FolderIcons_Initialize()
        Dim resultado As stdole.IPictureDisp = Nothing
        Dim sFilename As String = ""
        If lItem = ESYSTEM_ITEMS.MI_PC Then
            sFilename = "::{20D04FE0-3AEA-1069-A2D8-08002B30309D}"
        End If
        If lItem = ESYSTEM_ITEMS.PAPELERA Then
            sFilename = "::{645FF040-5081-101B-9F08-00AA002F954E}"
        End If
        If lItem = ESYSTEM_ITEMS.MIS_SITIOS_DE_RED Then
            sFilename = "::{208D2C60-3AEA-1069-A2D7-08002B30309D}"
        End If
        If lItem = ESYSTEM_ITEMS.PANEL_DE_CONTROL Then
            sFilename = "::{20D04FE0-3AEA-1069-A2D8-08002B30309D}\::{21EC2020-3AEA-1069-A2DD-08002B30309D}"
        End If
        If lItem = ESYSTEM_ITEMS.IMPRESORAS_FAX Then
            sFilename = "::{20D04FE0-3AEA-1069-A2D8-08002B30309D}\::{2227A280-3AEA-1069-A2DE-08002B30309D}"
        End If
        If lItem = ESYSTEM_ITEMS.CARPETAS_WEB Then
            sFilename = "::{20D04FE0-3AEA-1069-A2D8-08002B30309D}\::{BDEADF00-C265-11D0-BCED-00A0C90AB50F}"
        End If
        If lItem = ESYSTEM_ITEMS.HERRAMIENTAS_ADMINISTRATIVAS Then
            sFilename = "::{20D04FE0-3AEA-1069-A2D8-08002B30309D}\::{21EC2020-3AEA-1069-A2DD-08002B30309D}\::{D20EA4E1-3957-11d2-A40B-0C5020524153}"
        End If
        If lItem = ESYSTEM_ITEMS.CONEXIONES_DE_RED Then
            sFilename = "::{7007ACC7-3202-11D1-AAD2-00805FC1270E}"
        End If
        If lItem = ESYSTEM_ITEMS.FUENTES Then
            sFilename = "::{20D04FE0-3AEA-1069-A2D8-08002B30309D}\::{21EC2020-3AEA-1069-A2DD-08002B30309D}\::{D20EA4E1-3957-11d2-A40B-0C5020524152}"
        End If
        '
        If sFilename <> "" Then
            resultado = LoadIcon(lIconSize, sFilename)
        End If
        '
        Return resultado
    End Function
    ' -------------------------------------------------------------------------------
    '\\ -- Retorna la ruta de las carpetas especiales de windows
    ' -------------------------------------------------------------------------------
    Public Shared Function GetPath(SpecialFolder As ESPECIAL_FOLDERS) As String
        FolderIcons_Initialize()
        Dim lRet As Long
        Dim IDL As ITEMIDLIST
        Dim Spath As String
        Dim resultado As String = ""

        ' -- Obtener el ID del path
        lRet = SHGetSpecialFolderLocation(100, SpecialFolder, IDL)

        If lRet = 0 Then
            ' -- Crear Buffer de caracteres para el path
            Spath = Space(512)
            ' -- Recuperamos el path desde el ID
            lRet = SHGetPathFromIDList(IDL.mkid.cb, Spath)
            ' -- Eliminr nulos y retornar
            resultado = Left(Spath, InStr(Spath, Chr(0)) - 1)
        End If
        Return resultado
    End Function
    ' -----------------------------------------------------------------------------------------
    ' \\ -- Devolver el ícono de un archivo o recurso del sistema ( un drive, Item del panel de control etc ..) como IPictureDisp
    ' -----------------------------------------------------------------------------------------
    Public Shared Function LoadIcon(lIconSize As EICON_SIZE, ByVal sFileName As String) As stdole.IPictureDisp
        FolderIcons_Initialize()
        Dim ret As Long
        Dim resultado As Object = Nothing
        Dim Icon As IconType
        Dim CLSID As CLSIdType
        Dim ShellInfo As ShellFileInfoType
        ' Inizialice ShellInfo
        With ShellInfo
            .hIcon = 0
            .iIcon = 0
            .dwAttributes = 0
            .szDisplayName = ""
            .szTypeName = ""
        End With
        ' Set ShellInfo Data
        Call SHGetFileInfo(sFileName, 0, ShellInfo, Len(ShellInfo), lIconSize)
        With Icon
            .cbSize = Len(Icon)
            .picType = 3
            .hIcon = ShellInfo.hIcon
        End With
        With CLSID
            ReDim Preserve .id(16)
            .id(8) = &HC0
            .id(15) = &H46
        End With
        ret = OleCreatePictureIndirect(Icon, CLSID, 1, resultado)
        Return resultado
    End Function
    ' -----------------------------------------------------------------------------------------
    ' \\ -- Iniciar
    ' -----------------------------------------------------------------------------------------
    Private Shared Sub FolderIcons_Initialize()
        If mColSFolderCurrentUser IsNot Nothing And
            mColSFolderAllUsers IsNot Nothing And
            mColSystemPaths IsNot Nothing Then
            Exit Sub
        End If
        '
        On Error GoTo error_handler
        Dim i As Long
        Dim sPathSpecialFolder As String
        Dim sPathCurrentUser As String
        Dim sPathAllUsers As String

        ' -- Inicio de variables
        mColSFolderCurrentUser = New Collection
        mColSFolderAllUsers = New Collection
        mColSystemPaths = New Collection

        ' -- Obtener ruta base del usuario actual
        sPathCurrentUser = Environ("USERPROFILE")

        ' -- Obtener ruta base de la carpeta de todos los usuarios
        sPathAllUsers = Environ("ALLUSERSPROFILE")

        ' -- Recorrer y Obtener las carpetas especiales de windows
        For i = 1 To 100
            sPathSpecialFolder = GetPath(i)
            ' -- Comprobar el path
            If sPathSpecialFolder <> "" Then
                ' -- Si es el usuario actual  agregarlo a la colección ...
                If InStr(sPathSpecialFolder, sPathCurrentUser) Then
                    mColSFolderCurrentUser.Add(sPathSpecialFolder)
                    ' -- Si es la de todos los usuarios ...
                ElseIf InStr(sPathSpecialFolder, sPathAllUsers) Then
                    mColSFolderAllUsers.Add(sPathSpecialFolder)
                    ' -- Sino otras carpetas ( windows, archivos de programa, ...)
                Else
                    mColSystemPaths.Add(sPathSpecialFolder, sPathSpecialFolder)
                End If
            End If
        Next
        ' -- Error
        Exit Sub
error_handler:

        If Err.Number = 457 Then
            Resume Next
        Else
            MsgBox(Err.Description, vbCritical, "N° de error: " & CStr(Err.Number))
        End If
    End Sub

    Public Shared Function GetFolderName(ByVal Spath As String)
        FolderIcons_Initialize()
        On Error GoTo error_handler
        If Right(Spath, 1) = "\" Then Spath = Left(Spath, Len(Spath) - 1)
        Dim arrFolders() As String
        arrFolders = Split(Spath, "\")
        Return arrFolders(UBound(arrFolders))
        Exit Function
error_handler:
    End Function
    ' -----------------------------------------------------------------------------------------
    ' \\ -- Terminar
    ' -----------------------------------------------------------------------------------------
    Public Shared Sub FolderIcons_Terminate()
        mColSFolderCurrentUser = Nothing
        mColSFolderAllUsers = Nothing
        mColSystemPaths = Nothing
    End Sub
End Class
