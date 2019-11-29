'------------------------------------------------------------------------------
' Clase con las funciones del API de Windows                        (14/Ago/05)
' Las funciones están declaradas como compartidas
' para usarlas sin crear una instancia.
'
' ©Guillermo 'guille' Som, 2005
'------------------------------------------------------------------------------
Imports Microsoft.VisualBasic
Imports System
Imports System.Runtime.InteropServices

Public Class clsAPI
    Public Shared ultimaVentana As IntPtr = IntPtr.Zero
    Public Shared ultimoProceso As Integer = 0

    <System.Runtime.InteropServices.DllImport("user32.dll", EntryPoint:="SetWindowLong")>
    Friend Shared Function SetWindowLong32(ByVal hWnd As IntPtr, ByVal nIndex As Integer, ByVal dwNewLong As Integer) As Integer
    End Function

    <System.Runtime.InteropServices.DllImport("user32.dll", EntryPoint:="SetWindowLongPtr")>
    Friend Shared Function SetWindowLongPtr64(ByVal hWnd As IntPtr, ByVal nIndex As Integer, ByVal dwNewLong As IntPtr) As IntPtr
    End Function

    Friend Shared Function SetWindowLongPtr(ByVal hWnd As IntPtr, ByVal nIndex As Integer, ByVal dwNewLong As IntPtr) As IntPtr
        If IntPtr.Size = 8 Then
            Return SetWindowLongPtr64(hWnd, nIndex, dwNewLong)
        Else
            Return New IntPtr(SetWindowLong32(hWnd, nIndex, dwNewLong.ToInt32))
        End If
    End Function

    ' Declaraciones para extraer iconos de los programas
    <DllImport("shell32.dll", CharSet:=CharSet.Unicode)>
    Friend Shared Function ExtractIconEx(
      ByVal lpszFile As String, ByVal nIconIndex As Integer,
      ByRef phiconLarge As Integer, ByRef phiconSmall As Integer,
      ByVal nIcons As UInteger) As Integer
    End Function

    <DllImport("shell32.dll", CharSet:=CharSet.Unicode)>
    Friend Shared Function ExtractIcon(
      ByVal hInst As Integer, ByVal lpszExeFileName As String,
      ByVal nIconIndex As Integer) As IntPtr
    End Function

    Friend Declare Unicode Function SHGetImageList Lib "shell32.dll" (ByVal iImageList As Long, ByRef riid As Long, ByRef ppv As Object) As Long
    Friend Declare Unicode Function SHGetImageListXP Lib "shell32.dll" Alias "#727" (ByVal iImageList As Long, ByRef riid As Long, ByRef ppv As Object) As Long


    ' Hace que una ventana sea hija (o esté contenida) en otra
    <System.Runtime.InteropServices.DllImport("user32.dll")>
    Friend Shared Function SetParent(
        ByVal hWndChild As IntPtr,
        ByVal hWndNewParent As IntPtr) As IntPtr
    End Function

    ' Nos da un array con el texto de la ventana.
    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Unicode)>
    Friend Shared Function GetWindowText(
        ByVal hwnd As IntPtr,
        ByVal lpString As Text.StringBuilder,
        ByVal cch As Integer) As Integer
    End Function

    ' Nos da la longitud del texto de la ventana.
    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto)>
    Friend Shared Function GetWindowTextLength(ByVal hwnd As IntPtr) As Integer
    End Function

    Public Shared Function GetText(ByVal hWnd As IntPtr) As String
        Dim length As Integer
        If hWnd.ToInt32 <= 0 Then
            Return Nothing
        End If
        length = GetWindowTextLength(hWnd)
        If length = 0 Then
            Return Nothing
        End If
        Dim sb As New System.Text.StringBuilder("", length + 1)

        GetWindowText(hWnd, sb, sb.Capacity)
        Return sb.ToString()
    End Function

    Friend Declare Unicode Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hdc As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As POINTAPI) As Long
    Friend Declare Unicode Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long

    'Public Shared Function text_length(Mytext As String) As Long
    '    Dim TextSize As POINTAPI
    '    GetTextExtentPoint32(GetWindowDC(hwnd), Mytext, Len(Mytext), TextSize)
    '    text_length = TextSize.X
    'End Function

    ' Capturar el Handle de la ventana activa.
    <System.Runtime.InteropServices.DllImport("user32.dll", CharSet:=CharSet.Auto, ExactSpelling:=True)>
    Friend Shared Function GetActiveWindow() As IntPtr
    End Function

    ' Capturar el Handle de la ventana activa y que está en primer plano.
    <System.Runtime.InteropServices.DllImport("user32.dll", CharSet:=CharSet.Auto, ExactSpelling:=True)>
    Friend Shared Function GetForegroundWindow() As IntPtr
    End Function
    '' El Id de proceso de la venta.
    <DllImport("user32.dll", SetLastError:=True)>
    Friend Shared Function GetWindowThreadProcessId(ByVal hwnd As IntPtr,
                          ByRef lpdwProcessId As Integer) As Integer
    End Function

    ' Devuelve el Handle (hWnd) de una ventana de la que sabemos el título
    <System.Runtime.InteropServices.DllImport("user32.dll", CharSet:=CharSet.Unicode)>
    Friend Shared Function FindWindow(
        ByVal lpClassName As String,
        ByVal lpWindowName As String) As IntPtr
    End Function


    'Busca el Handle o Hwnd del botón hijo de la ventana encontrada con FindWindow 
    <System.Runtime.InteropServices.DllImport("user32.dll", CharSet:=CharSet.Unicode)>
    Friend Shared Function FindWindowEx(
        ByVal hWnd1 As IntPtr,
        ByVal hWnd2 As IntPtr,
        ByVal className As String,
        ByVal Caption As String) As IntPtr
    End Function
    'ByVal lpsz1 As String, _
    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Unicode)>
    Friend Shared Function GetClassName(
        ByVal hWnd As IntPtr,
        ByVal lpClassName As System.Text.StringBuilder,
        ByVal nMaxCount As Integer) As Integer
    End Function

    ' Cambia el tamaño y la posición de una ventana
    <System.Runtime.InteropServices.DllImport("user32.dll")>
    Friend Shared Function MoveWindow(ByVal hWnd As IntPtr,
    ByVal x As Integer, ByVal y As Integer,
     ByVal nWidth As Integer, ByVal nHeight As Integer,
     ByVal bRepaint As Integer) As Integer
    End Function

    ' Destruye (cierra y vacia de la memoria) una ventana
    <System.Runtime.InteropServices.DllImport("user32.dll")>
    Friend Shared Function DestroyWindow(ByVal hwnd As System.IntPtr) As System.IntPtr
    End Function
    ''
    <DllImport("User32.dll")>
    Private Shared Function EnumChildWindows _
        (ByVal WindowHandle As IntPtr, ByVal Callback As EnumWindowProcess,
        ByVal lParam As IntPtr) As Boolean
    End Function

    Public Delegate Function EnumWindowProcess(ByVal Handle As IntPtr, ByVal Parameter As IntPtr) As Boolean

    Public Shared Function GetChildWindows(ByVal ParentHandle As IntPtr) As IntPtr()
        Dim ChildrenList As New List(Of IntPtr)
        Dim ListHandle As GCHandle = GCHandle.Alloc(ChildrenList)
        Try
            EnumChildWindows(ParentHandle, AddressOf EnumWindow, GCHandle.ToIntPtr(ListHandle))
        Finally
            If ListHandle.IsAllocated Then ListHandle.Free()
        End Try
        Return ChildrenList.ToArray
    End Function

    Private Shared Function EnumWindow(ByVal Handle As IntPtr, ByVal Parameter As IntPtr) As Boolean
        Dim ChildrenList As List(Of IntPtr) = GCHandle.FromIntPtr(Parameter).Target
        If ChildrenList Is Nothing Then Throw New Exception("GCHandle Target could not be cast as List(Of IntPtr)")
        ChildrenList.Add(Handle)
        Return True
    End Function

    Public Shared Sub WindowAPI_Click(ByVal hwnd As IntPtr)
        Dim retVal As Long
        retVal = SendMessage(hwnd, eMensajes.WM_LBUTTONDOWN, 0, 0)
        retVal = SendMessage(hwnd, eMensajes.WM_LBUTTONDOWN, 0, 0)
        retVal = SendMessage(hwnd, eMensajes.WM_KEYUP, VK_SPACE, 0)
        retVal = SendMessage(hwnd, eMensajes.WM_LBUTTONUP, 0, 0)
    End Sub

    '------------------------------------------------------------------------------
    ' APIS para incluir las ventanas en un PictureBox
    '------------------------------------------------------------------------------
    '
    ' Para mostrar una ventana según el handle (hwnd)
    ' ShowWindow() Commands
    Public Enum eShowWindow As Integer
        HIDE_eSW = 0&
        SHOWNORMAL_eSW = 1&
        NORMAL_eSW = 1&
        SHOWMINIMIZED_eSW = 2&
        SHOWMAXIMIZED_eSW = 3&
        MAXIMIZE_eSW = 3&
        SHOWNOACTIVATE_eSW = 4&
        SHOW_eSW = 5&
        MINIMIZE_eSW = 6&
        SHOWMINNOACTIVE_eSW = 7&
        SHOWNA_eSW = 8&
        RESTORE_eSW = 9&
        SHOWDEFAULT_eSW = 10&
        MAX_eSW = 10&
    End Enum

    <System.Runtime.InteropServices.DllImport("user32.dll")>
    Friend Shared Function ShowWindow(
      ByVal hWnd As System.IntPtr,
      ByVal nCmdShow As eShowWindow) As Integer
    End Function

    Public Enum eTeclas As Integer
        WM_SETHOTKEY = &H32
        WM_SHOWWINDOW = &H18
        HK_SHIFTA = &H141 'Shift + A
        HK_SHIFTB = &H142 'Shift + B
        HK_CONTROLA = &H241 'Control + A
        HK_ALTZ = &H45A
        VK_RETURN = &HD
    End Enum

    Public Enum eMensajes As Integer
        cero = 0
        WM_ACTIVATE = &H6
        WM_SETHOTKEY = &H32
        WM_SHOWWINDOW = &H18
        WM_SYSCOMMAND = &H112
        SC_CLOSE = &HF060&
        VK_SPACE = &H20
        ''Activar y hacer clic 
        ''SendMessage(hWnd, WM_ACTIVATE, WA_ACTIVE, 0)
        ''SendMessage(hWnd, BM_CLICK, 0, 0) 

        WA_ACTIVE = 1
        WM_NULL = &H0
        WM_CREATE = &H1
        WM_DESTROY = &H2
        WM_CLOSE = &H10    '16 EN DECIMAL
        WM_MOVE = &H3
        WM_SIZE = &H5
        WM_COPYDATA = &H4A
        WM_SETTEXT = &HC
        WM_KEYDOWN = &H100
        WM_KEYUP = &H101
        MOUSE_MOVE = &HF012&
        LB_FINDSTRING = &H18F
        'Declaración de las constantes
        WM_USER = &H400
        EM_GETSEL = WM_USER + 0
        EM_SETSEL = WM_USER + 1
        EM_REPLACESEL = WM_USER + 18
        EM_UNDO = WM_USER + 23
        'EM_LINEFROMCHAR = WM_USER + 25
        'EM_GETLINECOUNT = WM_USER + 10
        EM_GETLINECOUNT = &HBA
        EM_LINEFROMCHAR = &HC9
        EM_LINELENGTH = &HC1
        EM_LINEINDEX = &HBB
        ''
        WM_CUT = &H300
        WM_COPY = &H301
        WM_PASTE = &H302
        WM_CLEAR = &H303    '' Limpiar seleccion

        WM_LBUTTONDOWN = &H201
        WM_LBUTTONUP = &H202 ' izquierdo arriba  
        WM_LBUTTONDBLCLK = &H203 ' izquierdo doble click 
        WM_click = 245          ' Click en botón.

        MK_CONTROL = &H8
        MK_LBUTTON = &H1
        MK_MBUTTON = &H10
        MK_RBUTTON = &H2
        MK_SHIFT = &H4
        MK_XBUTTON1 = &H20
        MK_XBUTTON2 = &H40

        BM_CLICK = &HF5
        IDOK = 1
        CLIK_BUTTON = &H83


        '' ***** PARA CASILLAS DE OPCIONES
        BM_SETCHECK = &HF1  ' Para poner estado en casillas de opción.
        BST_CHECKED = &H1
        BST_UNCHECKED = &H0  ' Casilla de opción desmarcada
        'Establecer como marcado
        'SendMessage (theHandle, BM_SETCHECK, BST_CHECKED, 0) 
    End Enum

#Region "  Constantes de SendMessage:"
    'Constantes de SendMessage:
    Public Const WM_CLOSE = &H10    '16 EN DECIMAL
    Public Const WM_KEYDOWN = &H100
    Public Const WM_KEYUP = &H101
    Public Const WM_SETTEXT = &HC
    Public Const WM_GETTEXT = &HD
    Public Const WM_GETTEXTLENGTH = &HE
    Public Const WM_CHAR = &H102
    Public Const WM_COMMAND = &H111
    Private Const GW_HWNDFIRST = 0&
    Private Const GW_HWNDNEXT = 2&
    Private Const GW_CHILD = 5&
    'Public Const GWL_HWNDPARENT = (-rie_gafas

    Public Const VK_LBUTTON = &H1
    Public Const VK_RBUTTON = &H2
    Public Const VK_CTRLBREAK = &H3
    Public Const VK_MBUTTON = &H4
    Public Const VK_BACKSPACE = &H8
    Public Const VK_TAB = &H9
    Public Const VK_ENTER = &HD
    Public Const VK_SHIFT = &H10
    Public Const VK_CONTROL = &H11
    Public Const VK_ALT = &H12
    Public Const VK_PAUSE = &H13
    Public Const VK_CAPSLOCK = &H14
    Public Const VK_ESCAPE = &H1B
    Public Const VK_SPACE = &H20
    Public Const VK_PAGEUP = &H21
    Public Const VK_PAGEDOWN = &H22
    Public Const VK_END = &H23
    Public Const VK_HOME = &H24
    Public Const VK_LEFT = &H25
    Public Const VK_UP = &H26
    Public Const VK_RIGHT = &H27
    Public Const VK_DOWN = &H28
    Public Const VK_PRINTSCREEN = &H2C
    Public Const VK_INSERT = &H2D
    Public Const VK_DELETE = &H2E
    Public Const VK_0 = &H30
    Public Const VK_1 = &H31
    Public Const VK_2 = &H32
    Public Const VK_3 = &H33
    Public Const VK_4 = &H34
    Public Const VK_5 = &H35
    Public Const VK_6 = &H36
    Public Const VK_7 = &H37
    Public Const VK_8 = &H38
    Public Const VK_9 = &H39
    Public Const VK_A = &H41
    Public Const VK_B = &H42
    Public Const VK_C = &H43
    Public Const VK_D = &H44
    Public Const VK_E = &H45
    Public Const VK_F = &H46
    Public Const VK_G = &H47
    Public Const VK_H = &H48
    Public Const VK_I = &H49
    Public Const VK_J = &H4A
    Public Const VK_K = &H4B
    Public Const VK_L = &H4C
    Public Const VK_M = &H4D
    Public Const VK_n = &H4E
    Public Const VK_O = &H4F
    Public Const VK_P = &H50
    Public Const VK_Q = &H51
    Public Const VK_R = &H52
    Public Const VK_S = &H53
    Public Const VK_T = &H54
    Public Const VK_U = &H55
    Public Const VK_V = &H56
    Public Const VK_W = &H57
    Public Const VK_X = &H58
    Public Const VK_Y = &H59
    Public Const VK_Z = &H5A
    Public Const VK_LWINDOWS = &H5B
    Public Const VK_RWINDOWS = &H5C
    Public Const VK_APPSPOPUP = &H5D
    Public Const VK_NUMPAD_0 = &H60
    Public Const VK_NUMPAD_1 = &H61
    Public Const VK_NUMPAD_2 = &H62
    Public Const VK_NUMPAD_3 = &H63
    Public Const VK_NUMPAD_4 = &H64
    Public Const VK_NUMPAD_5 = &H65
    Public Const VK_NUMPAD_6 = &H66
    Public Const VK_NUMPAD_7 = &H67
    Public Const VK_NUMPAD_8 = &H68
    Public Const VK_NUMPAD_9 = &H69
    Public Const VK_NUMPAD_MULTIPLY = &H6A
    Public Const VK_NUMPAD_ADD = &H6B
    Public Const VK_NUMPAD_PLUS = &H6B
    Public Const VK_NUMPAD_SUBTRACT = &H6D
    Public Const VK_NUMPAD_MINUS = &H6D
    Public Const VK_NUMPAD_MOINS = &H6D
    Public Const VK_NUMPAD_DECIMAL = &H6E
#End Region


    '' No espera a que termine. SendMessage si espera.
    <System.Runtime.InteropServices.DllImport("user32.dll")>
    Friend Shared Function PostMessage(
        ByVal hwnd As System.IntPtr,
        ByVal wMsg As eMensajes,
        ByVal wParam As Integer,
        ByVal lParam As Integer) As Integer
    End Function

    <DllImport("user32.dll", SetLastError:=True)>
    Friend Shared Function PostMessage(
        ByVal hWnd As IntPtr,
        <MarshalAs(UnmanagedType.U4)> ByVal Msg As UInteger,
        ByVal wParam As IntPtr,
        ByVal lParam As IntPtr) As Boolean
    End Function

    '' hWnd = handle of destination window
    '' wMsg = message to send
    '' wParam = first message parameter
    '' lParam = second message parameter
    <System.Runtime.InteropServices.DllImport("user32.DLL")>
    Friend Shared Function SendMessage(
        ByVal hWnd As System.IntPtr,
        ByVal wMsg As eMensajes,
        ByVal wParam As Integer,
        ByVal lParam As Integer) As Integer
    End Function
    '' Si queremos usar un String en el último parámetro y queremos usarla al mismo tiempo que la anterior,
    '' sólo hay que declararla nuevamente con los parámetros diferentes, (sin necesidad de cambiar el nombre)
    <System.Runtime.InteropServices.DllImport("user32.DLL")>
    Friend Shared Function SendMessage(
        ByVal hWnd As System.IntPtr,
        ByVal wMsg As eMensajes,
        ByVal wParam As Integer,
        ByVal lParam As Object
        ) As Integer
    End Function
    '' Si queremos usar un String en el último parámetro y queremos usarla al mismo tiempo que la anterior,
    '' sólo hay que declararla nuevamente con los parámetros diferentes, (sin necesidad de cambiar el nombre)
    <System.Runtime.InteropServices.DllImport("user32.DLL", CharSet:=CharSet.Unicode)>
    Friend Shared Function SendMessage(
        ByVal hWnd As System.IntPtr,
        ByVal wMsg As eMensajes,
        ByVal wParam As Integer,
        ByVal lParam As String
        ) As Integer
    End Function
    ''
    '<System.Runtime.InteropServices.DllImport("user32.DLL", CharSet:=CharSet.Auto)> _
    Friend Declare Unicode Function SendMessage_Long Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef LParam As Long) As Integer

    '
    ' Para cambiar el tamaño de una ventana y asignar los valores máximos y mínimos del tamaño
    Public Structure POINTAPI
        Dim x As Long
        Dim y As Long
    End Structure

    Public Structure RECTAPI
        Dim Left As Long
        Dim Top As Long
        Dim Right As Long
        Dim Bottom As Long
    End Structure

    Public Structure WINDOWPLACEMENT
        Dim Length As Long
        Dim Flags As Long
        Dim ShowCmd As Long
        Dim ptMinPosition As POINTAPI
        Dim ptMaxPosition As POINTAPI
        Dim rcNormalPosition As RECTAPI
    End Structure

    <System.Runtime.InteropServices.DllImport("user32.dll")>
    Friend Shared Function GetWindowPlacement(
        ByVal hWnd As Long,
        ByRef lpwndpl As WINDOWPLACEMENT) As Long
    End Function

    'Crear una ventana flotante al estilo de los tool-bar
    'Cuando se minimiza la ventana padre, también lo hace ésta.
    Public Const SWW_hParent = -8
    '************************************************

    'En Form_Load (suponiendo que la ventana padre es Form1)
    'If SetWindowWord(hWnd, SWW_hParent, form1.hWnd) Then
    'End If
    '***********************************************
    '
    <System.Runtime.InteropServices.DllImport("user32.dll", CharSet:=CharSet.Unicode)>
    Friend Shared Function SetWindowWord(
        ByVal hwnd As Long,
        ByVal nIndex As Long,
        ByVal wNewWord As Long) As Long
    End Function
    '

    Public Shared Sub ActivaAppAPI(ByVal queApp As String, Optional ByVal MoverC As Boolean = True)
        Dim intP As IntPtr = FindWindow(Nothing, queApp)
        If intP <> IntPtr.Zero Then
            SetForegroundWindow(intP)
            AppActivate(queApp)
            If MoverC = True Then System.Windows.Forms.Cursor.Position = New System.Drawing.Point(100, System.Windows.Forms.Screen.PrimaryScreen.WorkingArea.Height / 2)
        End If
    End Sub


    Public Const SW_HIDE = 0&
    Public Const SW_SHOWNORMAL = 1&
    Public Const SW_SHOWMINIMIZED = 2&
    Public Const SW_MAXIMIZE = 3&
    Public Const SW_SHOWMAXIMIZED = 3&
    Public Const SW_SHOWNOACTIVATE = 4&
    Public Const SW_SHOW = 5&
    Public Const SW_MINIMIZE = 6&
    Public Const SW_SHOWMINNOACTIVE = 7&
    Public Const SW_SHOWNA = 8&
    Public Const SW_RESTORE = 9&
    Public Const SW_SHOWDEFAULT = 10&
    Public Const SW_MAX = 10&


    'Public Declare Function ShowWindow Lib "user32" () _
    '(ByVal hWnd As Long, ByVal nCmdShow As eShowWindow) As Long

    ' Activate an application window.
    '<System.Runtime.InteropServices.DllImport("user32.dll")> _
    'Public Declare Function SetForegroundWindow Lib "user32" Alias "SetForegroundWindow" (ByVal hWnd As System.IntPtr) As Integer

    <System.Runtime.InteropServices.DllImport("user32.dll")>
    Friend Shared Function SetForegroundWindow(
        ByVal hWnd As System.IntPtr) As Integer
    End Function
    ''
    <DllImport("user32.dll", SetLastError:=True)>
    Friend Shared Function IsWindowVisible(ByVal hWnd As IntPtr) As <MarshalAs(UnmanagedType.Bool)> Boolean
    End Function
    ''
    <DllImport("user32.dll", SetLastError:=True)>
    Friend Shared Function GetShellWindow() As IntPtr
    End Function
    ' Para que la ventanta tenga el foco del teclado
    '<System.Runtime.InteropServices.DllImport("user32.dll")> _
    'Public Declare Function SetFocus Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long

    <System.Runtime.InteropServices.DllImport("user32.dll")>
    Friend Shared Function SetFocus(
        ByVal hwnd As Long) As Long
    End Function


    Friend Shared Function DameIntPtr(ByVal queApp As String) As IntPtr
        Dim intP As IntPtr = Nothing
        intP = FindWindow(Nothing, queApp)
        DameIntPtr = intP
    End Function


    Private Structure COPYDATASTRUCT
        Dim dwData As Long
        Dim cbData As Long
        Dim lpData As Long
    End Structure

    Friend Declare Unicode Sub CopyMemory _
    Lib "kernel32" _
    Alias "RtlMoveMemory" (ByVal hpvDest As Object,
    ByVal hpvSource As Object,
    ByVal cbCopy As Long)

    Friend Shared Function SendCmd(ByVal sCommand As String, ByVal ventana As Long) As Long

        'On Error Resume Next

        Dim tCDS As COPYDATASTRUCT
        Dim abytDATA(0 To 255) As Byte

        Call CopyMemory(abytDATA(1), sCommand, Len(sCommand))

        With tCDS
            .dwData = 1
            .cbData = Len(sCommand) + 1
            .lpData = (abytDATA(1)) 'VarPtr(abytDATA(1))
        End With

        SendCmd = SendMessage(ventana, eMensajes.WM_COPYDATA, vbNull, tCDS)

    End Function

    Public Enum ShellSpecialFolderConstants As Integer
        ssfDESKTOP = &H0
        ssfPROGRAMS = &H2
        ssfCONTROLS = &H3
        ssfPRINTERS = &H4
        ssfPERSONAL = &H5
        ssfFAVORITES = &H6
        ssfSTARTUP = &H7
        ssfRECENT = &H8
        ssfSENDTO = &H9
        ssfBITBUCKET = &HA
        ssfSTARTMENU = &HB
        ssfDESKTOPDIRECTORY = &H10
        ssfDRIVES = &H11
        ssfNETWORK = &H12
        ssfNETHOOD = &H13
        ssfFONTS = &H14
        ssfTEMPLATES = &H15
        ssfCOMMONSTARTMENU = &H16
        ssfCOMMONPROGRAMS = &H17
        ssfCOMMONSTARTUP = &H18
        ssfCOMMONDESKTOPDIR = &H19
        ssfAPPDATA = &H1A
        ssfPRINTHOOD = &H1B
        ssfLOCALAPPDATA = &H1C
        ssfALTSTARTUP = &H1D
        ssfCOMMONALTSTARTUP = &H1E
        ssfCOMMONFAVORITES = &H1F
        ssfINTERNETCACHE = &H20
        ssfCOOKIES = &H21
        ssfHISTORY = &H22
        ssfCOMMONAPPDATA = &H23
        ssfWINDOWS = &H24
        ssfSYSTEM = &H25
        ssfPROGRAMFILES = &H26
        ssfMYPICTURES = &H27
        ssfPROFILE = &H28
    End Enum


    Public Structure RECT
        Dim Left As Long
        Dim Top As Long
        Dim Right As Long
        Dim Bottom As Long
    End Structure

    Public Shared Function CierraProcesoViejo(nombrePro As String) As Boolean
        Dim resultado As Boolean = False
        ''
        '' Cerrar firefox, Google Chrome o Internet Explorer.
        Dim oProc As Process() = Process.GetProcessesByName(nombrePro)
        If oProc IsNot Nothing AndAlso oProc.Count > 0 Then
            For Each oPr As Process In oProc
                Try
                    resultado = oPr.CloseMainWindow
                    oPr.Close()
                    If resultado = False Then oPr.Kill()
                Catch ex As Exception
                    oPr.Kill()
                End Try
            Next
        End If
        oProc = Nothing
        Return resultado
    End Function
    ''
    Public Shared Function ActivaProceso(nombrePro As String) As Boolean
        Dim resultado As Boolean = False
        ''
        '' Cerrar firefox, Google Chrome o Internet Explorer.
        Dim oProc As Process() = Process.GetProcesses   '.GetProcessesByName(nombrePro)
        If oProc IsNot Nothing AndAlso oProc.Count > 0 Then
            For Each oPr As Process In oProc
                If oPr.ProcessName.ToUpper.Contains(nombrePro.ToUpper) = False Then Continue For
                Try
                    Call SetForegroundWindow(oPr.MainWindowHandle)
                Catch ex As Exception
                    'oPr.Kill()
                End Try
            Next
        End If
        oProc = Nothing
        Return resultado
    End Function
    ''
    Public Shared Function CierraProcesoParteNombre(nombrePro As String) As Boolean
        Dim resultado As Boolean = False
        ''
        '' Cerrar firefox, Google Chrome o Internet Explorer.
        Dim oProc As Process() = Process.GetProcesses   '.GetProcessesByName(nombrePro)
        If oProc IsNot Nothing AndAlso oProc.Count > 0 Then
            For Each oPr As Process In oProc
                If oPr.ProcessName.ToUpper.Contains(nombrePro.ToUpper) = False Then Continue For
                Try
                    resultado = oPr.CloseMainWindow
                    oPr.Close()
                    If resultado = False Then oPr.Kill()
                Catch ex As Exception
                    'oPr.Kill()
                End Try
            Next
        End If
        oProc = Nothing
        Return resultado
    End Function
    ''
    ''
    Public Shared Function DameProcesos(queFi As String) As String
        Dim mensaje As String = ""
        Dim texto As String = "Proceso actual = " &
                             Process.GetCurrentProcess.ProcessName & "(" &
                             Process.GetCurrentProcess.MainWindowTitle & ")" & vbCrLf
        mensaje &= texto
        IO.File.WriteAllText(queFi, texto) 'quePro.MainWindowTitle)
        Dim nivel As Integer = 1
        ''
        ''Debug.Print(clsAPI.GetText(vInforme))   ' & " / " & clsAPI.g)
        For Each procHijo As Process In Process.GetProcesses
            texto = vbTab & procHijo.ProcessName & "(" & procHijo.MainWindowTitle & ")" & vbCrLf
            mensaje &= texto
            IO.File.AppendAllText(queFi, texto)
            System.Windows.Forms.Application.DoEvents()
        Next
        ''
        Return mensaje
    End Function
    ''
    ''
    Public Shared Function DameProcesoActual(quePro As Process, queFi As String) As String
        If quePro Is Nothing Then quePro = Process.GetCurrentProcess
        IO.File.WriteAllText(queFi, quePro.ProcessName) 'quePro.MainWindowTitle)
        Dim nivel As Integer = 1
        ''
        ''Debug.Print(clsAPI.GetText(vInforme))   ' & " / " & clsAPI.g)
        For Each procHijo As ProcessThread In quePro.Threads
            IO.File.AppendAllText(queFi, StrDup(nivel, vbTab) & procHijo.ToString & vbCrLf)
            System.Windows.Forms.Application.DoEvents()
        Next
        ''
        Return ""
    End Function
    ''
    ''
    Public Shared Function DameVentanasHija(ventanaPadre As IntPtr, queFi As String) As String
        Dim mensaje As String = clsAPI.GetText(ventanaPadre) & vbCrLf
        IO.File.WriteAllText(queFi, mensaje)
        Dim nivel As Integer = 1
        ''
        ''Debug.Print(clsAPI.GetText(vInforme))   ' & " / " & clsAPI.g)
        For Each vhija As IntPtr In clsAPI.GetChildWindows(ventanaPadre)
            DameVentanasHijaRecursivo(vhija, nivel, queFi)
            System.Windows.Forms.Application.DoEvents()
        Next
        ''
        Return mensaje
    End Function
    ''
    ''
    Public Shared Sub DameVentanasHijaRecursivo(ventanaPadre As IntPtr, nivel As Integer, queFi As String)
        Dim queTexto As String = clsAPI.GetText(ventanaPadre)
        If queTexto IsNot Nothing AndAlso queTexto <> "" Then
            IO.File.AppendAllText(queFi, StrDup(nivel, vbTab) & queTexto & vbCrLf)
        End If
        ''
        ''Debug.Print(clsAPI.GetText(vInforme))   ' & " / " & clsAPI.g)
        For Each vhija As IntPtr In clsAPI.GetChildWindows(ventanaPadre)
            DameVentanasHijaRecursivo(vhija, nivel + 1, queFi)
            System.Windows.Forms.Application.DoEvents()
        Next
    End Sub

    ''
    Public Shared Function PulsaBotonForm(nombreBoton As String, Optional esperar As Boolean = False) As Long
        ''
        '' Para esperar, utilizar PostMessage en vez de SendMessage
        Dim retVal As Long = 0
        ''
        'Dim oProceso As Process = Process.GetCurrentProcess
        Dim ventana0 As System.IntPtr = clsAPI.GetForegroundWindow  '  clsAPI.FindWindow("Afx:000000013F1C0000:8:0000000000010003:0000000000000000:0000000000500D79", "Autodesk Inventor Professional 2015") 'oProIn.Handle ' clsAPI.FindWindow("#32769 (Escritorio)", "")
        ultimaVentana = ventana0
        If ventana0 = IntPtr.Zero Then
            Return retVal
            Exit Function
        End If
        ''Debug.Print(clsAPI.GetText(vInforme))   ' & " / " & clsAPI.g)
        For Each vhija As IntPtr In clsAPI.GetChildWindows(ventana0)
            Dim texto As String = clsAPI.GetText(vhija)
            If Not (texto Is Nothing) AndAlso texto <> "" AndAlso (texto = nombreBoton Or texto.StartsWith(nombreBoton)) Then
                'retVal = clsAPI.SendMessage(vhija, clsAPI.eMensajes.WM_LBUTTONDOWN, 0, 0)
                'retVal = clsAPI.SendMessage(vhija, clsAPI.eMensajes.WM_LBUTTONDOWN, 0, 0)
                'retVal = clsAPI.SendMessage(vhija, clsAPI.eMensajes.WM_KEYDOWN, clsAPI.eMensajes.VK_SPACE, 0)
                'retVal = clsAPI.SendMessage(vhija, clsAPI.eMensajes.WM_LBUTTONUP, 0, 0)
                'If esperar Then
                'retVal = clsAPI.PostMessage(vhija, clsAPI.eMensajes.WM_click, 0, 0)
                'Else
                'retVal = clsAPI.SendMessage(vhija, clsAPI.eMensajes.WM_click, 0, 0)
                'End If
                '' PostMessage espera a que se complete la tarea.
                ''retVal = clsAPI.SendMessage(ventana0, eMensajes.WM_SYSCOMMAND, clsAPI.eMensajes.CLIK_BUTTON, vhija)
                'System.Windows.Forms.SendKeys.Send("{ENTER}")
                'System.Windows.Forms.SendKeys.Send("{ENTER}")
                'retVal = clsAPI.SendMessage(vhija, eMensajes.WM_ACTIVATE, eMensajes.WA_ACTIVE, 0)
                'retVal = clsAPI.SendMessage(vhija, eMensajes.BM_CLICK, 0, 0)
                retVal = clsAPI.SendMessage(vhija, eMensajes.WM_KEYDOWN, eMensajes.VK_SPACE, 0)
                retVal = clsAPI.SendMessage(vhija, eMensajes.WM_KEYUP, eMensajes.VK_SPACE, 0)
                Exit For
            End If
            For Each vhija1 As IntPtr In clsAPI.GetChildWindows(vhija)
                Dim texto1 As String = clsAPI.GetText(vhija1)
                If Not (texto1 Is Nothing) AndAlso texto1 <> "" AndAlso (texto1 = nombreBoton Or texto1.StartsWith(nombreBoton)) Then
                    'If esperar Then
                    'retVal = clsAPI.PostMessage(vhija1, clsAPI.eMensajes.WM_click, 0, 0)
                    'Else
                    'retVal = clsAPI.SendMessage(vhija1, clsAPI.eMensajes.WM_click, 0, 0)
                    'End If
                    'System.Windows.Forms.SendKeys.Send("{ENTER}")
                    'System.Windows.Forms.SendKeys.Send("{ENTER}")
                    'retVal = clsAPI.SendMessage(vhija1, eMensajes.WM_ACTIVATE, eMensajes.WA_ACTIVE, 0)
                    'retVal = clsAPI.SendMessage(vhija1, eMensajes.BM_CLICK, 0, 0)
                    retVal = clsAPI.SendMessage(vhija1, eMensajes.WM_KEYDOWN, eMensajes.VK_SPACE, 0)
                    retVal = clsAPI.SendMessage(vhija1, eMensajes.WM_KEYUP, eMensajes.VK_SPACE, 0)
                End If
            Next
        Next
        ''
        'While oProceso.HasExited = False
        '' No hacemos nada.
        'End While
        Return retVal
    End Function
    ''
    Public Shared Function PulsaBotonFormVentana(
                                nombreVentana As String,
                                nombreBoton As String,
                                Optional esperar As Boolean = False) As Long
        ''
        '' Para esperar, utilizar PostMessage en vez de SendMessage
        Dim retVal As Long = 0
        ''
        'Dim oProceso As Process = Process.GetCurrentProcess
        Dim ventana0 As System.IntPtr = clsAPI.FindWindow("", nombreVentana) 'oProIn.Handle ' clsAPI.FindWindow("#32769 (Escritorio)", "")
        ultimaVentana = ventana0
        ''Debug.Print(clsAPI.GetText(vInforme))   ' & " / " & clsAPI.g)
        If ventana0 = IntPtr.Zero Then
            Return retVal
            Exit Function
        End If
        For Each vhija As IntPtr In clsAPI.GetChildWindows(ventana0)
            Dim texto As String = clsAPI.GetText(vhija)
            If Not (texto Is Nothing) AndAlso texto <> "" AndAlso (texto = nombreBoton Or texto.StartsWith(nombreBoton)) Then
                'retVal = clsAPI.SendMessage(vhija, clsAPI.eMensajes.WM_LBUTTONDOWN, 0, 0)
                'retVal = clsAPI.SendMessage(vhija, clsAPI.eMensajes.WM_LBUTTONDOWN, 0, 0)
                'retVal = clsAPI.SendMessage(vhija, clsAPI.eMensajes.WM_KEYDOWN, clsAPI.eMensajes.VK_SPACE, 0)
                'retVal = clsAPI.SendMessage(vhija, clsAPI.eMensajes.WM_LBUTTONUP, 0, 0)
                'If esperar Then
                'retVal = clsAPI.PostMessage(vhija, clsAPI.eMensajes.WM_click, 0, 0)
                'Else
                'retVal = clsAPI.SendMessage(vhija, clsAPI.eMensajes.WM_click, 0, 0)
                'End If
                '' PostMessage espera a que se complete la tarea.
                ''retVal = clsAPI.SendMessage(ventana0, eMensajes.WM_SYSCOMMAND, clsAPI.eMensajes.CLIK_BUTTON, vhija)
                'System.Windows.Forms.SendKeys.Send("{ENTER}")
                'System.Windows.Forms.SendKeys.Send("{ENTER}")
                'retVal = clsAPI.SendMessage(vhija, eMensajes.WM_ACTIVATE, eMensajes.WA_ACTIVE, 0)
                'retVal = clsAPI.SendMessage(vhija, eMensajes.BM_CLICK, 0, 0)
                retVal = clsAPI.SendMessage(vhija, eMensajes.WM_KEYDOWN, ConsoleKey.Enter, 0)
                retVal = clsAPI.SendMessage(vhija, eMensajes.WM_KEYUP, ConsoleKey.Enter, 0)
                Exit For
            End If
            For Each vhija1 As IntPtr In clsAPI.GetChildWindows(vhija)
                Dim texto1 As String = clsAPI.GetText(vhija1)
                If Not (texto1 Is Nothing) AndAlso texto1 <> "" AndAlso (texto1 = nombreBoton Or texto1.StartsWith(nombreBoton)) Then
                    'If esperar Then
                    'retVal = clsAPI.PostMessage(vhija1, clsAPI.eMensajes.WM_click, 0, 0)
                    'Else
                    'retVal = clsAPI.SendMessage(vhija1, clsAPI.eMensajes.WM_click, 0, 0)
                    'End If
                    'System.Windows.Forms.SendKeys.Send("{ENTER}")
                    'System.Windows.Forms.SendKeys.Send("{ENTER}")
                    'retVal = clsAPI.SendMessage(vhija1, eMensajes.WM_ACTIVATE, eMensajes.WA_ACTIVE, 0)
                    'retVal = clsAPI.SendMessage(vhija1, eMensajes.BM_CLICK, 0, 0)
                    retVal = clsAPI.SendMessage(vhija1, eMensajes.WM_KEYDOWN, eMensajes.VK_SPACE, 0)
                    retVal = clsAPI.SendMessage(vhija1, eMensajes.WM_KEYUP, eMensajes.VK_SPACE, 0)
                End If
            Next
        Next
        ''
        'While oProceso.HasExited = False
        '' No hacemos nada.
        'End While
        Return retVal
    End Function

    ''
    Public Shared Function PulsaEnterEnVentanaActiva(Optional esperar As Boolean = False) As Long
        ''
        '' Para esperar, utilizar PostMessage en vez de SendMessage
        Dim retVal As Long = 0
        ''
        'If oProceso Is Nothing Then oProceso = Process.GetCurrentProcess
        'Dim ventana0 As System.IntPtr = oProceso.MainWindowHandle ' clsAPI.GetForegroundWindow
        Dim ventana0 As System.IntPtr = clsAPI.GetForegroundWindow
        SetForegroundWindow(ventana0)
        ultimaVentana = ventana0
        ''Debug.Print(clsAPI.GetText(vInforme))   ' & " / " & clsAPI.g)
        If ventana0 = IntPtr.Zero Then
            Return retVal
            Exit Function
        End If
        ''
        If esperar = True Then
            retVal = clsAPI.PostMessage(ventana0, eMensajes.WM_KEYDOWN, ConsoleKey.Enter, 0)
            retVal = clsAPI.PostMessage(ventana0, eMensajes.WM_KEYUP, ConsoleKey.Enter, 0)
        Else
            retVal = clsAPI.SendMessage(ventana0, eMensajes.WM_KEYDOWN, ConsoleKey.Enter, 0)
            retVal = clsAPI.SendMessage(ventana0, eMensajes.WM_KEYUP, ConsoleKey.Enter, 0)
        End If
        ''
        Return retVal
    End Function
    ''
    Public Shared Function EscribeEnTextBoxBuscaFinal(textoBusco As String, textoEscribo As String) As Long
        Dim retVal As Long = 0
        ''
        'Dim oProceso As Process = Process.GetCurrentProcess
        Dim ventana0 As System.IntPtr = clsAPI.GetForegroundWindow  '  clsAPI.FindWindow("Afx:000000013F1C0000:8:0000000000010003:0000000000000000:0000000000500D79", "Autodesk Inventor Professional 2015") 'oProIn.Handle ' clsAPI.FindWindow("#32769 (Escritorio)", "")
        ultimaVentana = ventana0
        ''Debug.Print(clsAPI.GetText(vInforme))   ' & " / " & clsAPI.g)
        For Each vhija As IntPtr In clsAPI.GetChildWindows(ventana0)
            Dim texto As String = clsAPI.GetText(vhija)
            'Dim k As Integer
            'Dim L1 As Integer, ltexto As Integer
            '' K retorna el Número de líneas del TextBox  
            'k = SendMessage(vhija, clsAPI.eMensajes.EM_GETLINECOUNT, 0, 0&)
            ' L1 el Primer carácter de la línea actual  
            'L1 = SendMessage(vhija, clsAPI.eMensajes.EM_LINEINDEX, k - 1, 0&) + 1
            ' Longitud de la línea actual (Cantidad de caracteres )  
            'ltexto = SendMessage(vhija, clsAPI.eMensajes.EM_LINELENGTH, L1, 0&)
            ' Mostramos la ultima línea del textbox  
            'MsgBox(Mid$(Text1.Text, L1, L2), vbInformation)
            'EM_LINELENGTH
            If Not (texto Is Nothing) AndAlso texto <> "" AndAlso texto.ToLower.EndsWith(textoBusco.ToLower) Then
                '' Primero borrar el texto actual
                retVal = clsAPI.SendMessage(vhija, clsAPI.eMensajes.WM_SETTEXT, 255, StrDup(255, " "))
                '' Segundo poner el texto definitivo
                retVal = clsAPI.SendMessage(vhija, clsAPI.eMensajes.WM_SETTEXT, textoEscribo.Length, textoEscribo)
                Exit For
            End If
            For Each vhija1 As IntPtr In clsAPI.GetChildWindows(vhija)
                Dim texto1 As String = clsAPI.GetText(vhija1)
                'Dim k1 As Integer
                'Dim L1a As Integer, ltexto1 As Integer
                '' K retorna el Número de líneas del TextBox  
                'k1 = SendMessage(vhija1, clsAPI.eMensajes.EM_GETLINECOUNT, 0, 0&)
                ' L1 el Primer carácter de la línea actual  
                'L1a = SendMessage(vhija1, clsAPI.eMensajes.EM_LINEINDEX, k1 - 1, 0&) + 1
                ' Longitud de la línea actual (Cantidad de caracteres )  
                'ltexto1 = SendMessage(vhija1, clsAPI.eMensajes.EM_LINELENGTH, L1a, 0&)
                If Not (texto1 Is Nothing) AndAlso texto1 <> "" AndAlso texto1.ToLower.EndsWith(textoBusco.ToLower) Then
                    '' Primero borrar el texto actual
                    retVal = clsAPI.SendMessage(vhija1, clsAPI.eMensajes.WM_SETTEXT, 255, StrDup(255, " "))
                    '' Segundo poner el texto definitivo
                    retVal = clsAPI.SendMessage(vhija1, clsAPI.eMensajes.WM_SETTEXT, textoEscribo.Length, textoEscribo)
                    Exit For
                End If
            Next
        Next
        ''
        Return retVal
    End Function
    ''
    Public Shared Function EscribeEnTextBoxBuscaInicio(textoBusco As String, textoEscribo As String) As Long
        Dim retVal As Long = 0
        ''
        Dim oProceso As Process = Process.GetCurrentProcess
        Dim ventana0 As System.IntPtr = clsAPI.GetForegroundWindow  '  clsAPI.FindWindow("Afx:000000013F1C0000:8:0000000000010003:0000000000000000:0000000000500D79", "Autodesk Inventor Professional 2015") 'oProIn.Handle ' clsAPI.FindWindow("#32769 (Escritorio)", "")
        ultimaVentana = ventana0
        ''Debug.Print(clsAPI.GetText(vInforme))   ' & " / " & clsAPI.g)
        For Each vhija As IntPtr In clsAPI.GetChildWindows(ventana0)
            Dim texto As String = clsAPI.GetText(vhija)
            If Not (texto Is Nothing) AndAlso texto <> "" AndAlso texto.StartsWith(textoBusco) Then
                retVal = clsAPI.SendMessage(vhija, clsAPI.eMensajes.WM_SETTEXT, textoEscribo.Length, textoEscribo)
                Exit For
            End If
            For Each vhija1 As IntPtr In clsAPI.GetChildWindows(vhija)
                Dim texto1 As String = clsAPI.GetText(vhija1)
                If Not (texto1 Is Nothing) AndAlso texto1 <> "" AndAlso texto1.StartsWith(textoBusco) Then
                    retVal = clsAPI.SendMessage(vhija1, clsAPI.eMensajes.WM_SETTEXT, textoEscribo.Length, textoEscribo)
                    Exit For
                End If
            Next
        Next
        ''
        Return retVal
    End Function
    ''
    Public Shared Function EscribeEnTextBoxBuscaIgualContine(textoBusco As String, textoEscribo As String) As Long
        Dim retVal As Long = 0
        ''
        Dim oProceso As Process = Process.GetCurrentProcess
        Dim ventana0 As System.IntPtr = clsAPI.GetForegroundWindow  '  clsAPI.FindWindow("Afx:000000013F1C0000:8:0000000000010003:0000000000000000:0000000000500D79", "Autodesk Inventor Professional 2015") 'oProIn.Handle ' clsAPI.FindWindow("#32769 (Escritorio)", "")
        ultimaVentana = ventana0
        ''Debug.Print(clsAPI.GetText(vInforme))   ' & " / " & clsAPI.g)
        For Each vhija As IntPtr In clsAPI.GetChildWindows(ventana0)
            Dim texto As String = clsAPI.GetText(vhija)
            If Not (texto Is Nothing) AndAlso texto <> "" AndAlso (texto = textoBusco Or texto.Contains(textoBusco)) Then
                retVal = clsAPI.SendMessage(vhija, clsAPI.eMensajes.WM_SETTEXT, textoEscribo.Length, textoEscribo)
                Exit For
            End If
            For Each vhija1 As IntPtr In clsAPI.GetChildWindows(vhija)
                Dim texto1 As String = clsAPI.GetText(vhija1)
                If Not (texto1 Is Nothing) AndAlso texto1 <> "" AndAlso (texto1 = textoBusco Or texto1.Contains(textoBusco)) Then
                    retVal = clsAPI.SendMessage(vhija1, clsAPI.eMensajes.WM_SETTEXT, textoEscribo.Length, textoEscribo)
                    Exit For
                End If
            Next
        Next
        ''
        Return retVal
    End Function
    ' 64-bit version
    'Public Shared Declare Sub PtrSafe Sleep Lib "kernel32" (ByVal dwMilliseconds As LongLong)

    ' 32-bit version
    Friend Declare Unicode Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

    Public Shared Sub Retardo(ByVal segundos As Integer)
        Const NSPerSecond As Long = 10000000
        Dim ahora As Long = Date.Now.Ticks
        'Console.WriteLine(Date.Now.Ticks)
        Debug.Print(Date.Now.Ticks)
        Do
            ' No hacemos nada
            'My.Application.DoEvents()
        Loop While Date.Now.Ticks < ahora + (segundos * NSPerSecond)
        'Console.WriteLine(Date.Now.Ticks)
        Debug.Print(Date.Now.Ticks)
    End Sub
End Class
