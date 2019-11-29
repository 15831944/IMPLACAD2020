Module modTeclas
    '************* Este código se pondrá en el formulario Form1...
    '
    'Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
    '   Dim proc As Process
    '   proc = Process.Start("notepad.exe") 'ejecutamos el notepad, pa ejecutar un autocad con un fichero, tan solo hay q pasarle la ruta del mismo como un argumento 
    '   proc.WaitForInputIdle() 'esperamos a que ejecute la aplicacion bien 
    '       con la funcion del api decimos que el handle del notepad(proc.mainwindow...) pertenece al hanbdle del panel1 
    '   SetParent(proc.MainWindowHandle, Me.Panel1.Handle) ' e.g. Me.Handle, Me.RichTextBox1.Handle, etc. 
    '       y pa adaptar el tamaño del notepad al del panel, volvemos a utilizar el api 
    '   SendMessage(proc.MainWindowHandle, WM_SYSCOMMAND, SC_MAXIMIZE, 0)
    'End Sub
    Private Const WM_SYSCOMMAND As Integer = 274
    Private Const SC_MAXIMIZE As Integer = 61488

    Public Declare Function BringWindowToTop Lib "user32" Alias "BringWindowToTop" _
    (ByVal hwnd As IntPtr) As Long
    Public Declare Function SetFocus Lib "user32" Alias "SetFocus" _
    (ByVal hwnd As IntPtr) As Long

    Public Declare Function ReleaseCapture Lib "user32" () As Long              'Soltar si hemos cogido algo
    'Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" _

    Declare Auto Function SetParent Lib "user32.dll" _
    (ByVal hWndChild As IntPtr, ByVal hWndNewParent As IntPtr) As Integer

    Declare Auto Function SendMessage Lib "user32.dll" _
    (ByVal hWnd As IntPtr, ByVal Msg As Integer, ByVal wParam As Integer, ByVal lParam As Integer) As Integer

    Declare Function PostMessage Lib "user32" Alias "PostMessageA" _
            (ByVal hWnd As IntPtr, ByVal wMsg As teclas, _
            Optional ByVal wParam As comandos = 0, Optional ByVal lParam As Long = 0) As Long

    ' (RedrawWindow) Redibuja una ventana
    Public Const RDW_INVALIDATE = &H1
    Public Declare Function RedrawWindow Lib "user32" _
    (ByVal hwnd As Long, ByVal lprcUpdate As Object, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long

    '** Informa de si está activa (1) o no activa (0) una ventana
    Public Declare Function IsWindowEnabled Lib "user32" (ByVal hwnd As Long) As Long
    'Activa (1) o desactiva (0) una ventana
    Public Declare Function EnableWindow Lib "user32" _
    (ByVal hwnd As Long, ByVal fEnable As Long) As Long

    ' Ejemplo: if IsWindowEnabled(hwnd)=0 then call EnableWindows(hwnd, 1)
    '****


    Public Enum comandos As Integer
        NADA = 0
        SC_MOVE = &HF010 'Mueve la ventana.
        MOUSE_MOVE = &HF012
        KEYEVENTF_KEYUP = &H2
        SC_MAXIMIZE = 61488  '(o SC_ZOOM) Maximiza la ventana. 
        SC_CLOSE 'Cierra la ventana. 
        SC_CONTEXTHELP 'Cambia el cursor al signo de interrogación con un puntero. Si el usuario hace click en un control en el cuadro de diálogo, el control recibirá un mensaje WM_HELP. 
        SC_DEFAULT 'Selecciona el elemento por defecto; el usuario hizo doble click en el menú de sistema. 
        SC_HOTKEY 'Activa la ventana asociada con la acción especificada en el "hot key". La palabra de menor peso del parámetro lParam identifica la ventana a activar. 
        SC_HSCROLL 'Desplaza el contenido de la ventana horizontalmente. 
        SC_KEYMENU 'Recupera el menú del sistema como resultado de la pulsación de una tecla. 
        SC_MINIMIZE  '(o SC_ICON) Minimiza la ventana. 
        SC_MONITORPOWER 'Sólo para Windows 95: ajusta el estado del monitor. Este comando soporta dispositivos que tengan propiedades de ahorro de energía, como las baterías de un ordenador portátil. 
        SC_MOUSEMENU 'Recupera el menú del sistema como resultado de un click del ratón. 
        SC_NEXTWINDOW 'Cambia a la siguiente ventana. 
        SC_PREVWINDOW 'Cambia a la ventana anterior. 
        SC_RESTORE 'Recupera el tamaño y posición normales de la ventana. 
        SC_SCREENSAVE 'Ejecuta la aplicación de salva pantallas especificada en la sección [boot] del fichero SYSTEM.INI. 
        SC_SIZE 'Cambia el tamaño de la ventana. 
        SC_TASKLIST 'Ejecuta o activa el gestor de tareas de Windows. 
        SC_VSCROLL 'Desplaza el contenido de la ventana verticalmente. 
    End Enum

    'VK - Teclado   WM - Comandos a utilizar (parametro 2) y hay que poner parámetro 3 una tecla
	Public Enum teclas As Integer
        WM_ACTIVATE = &H6
        WM_SETFOCUS = &H7
		WM_SYSKEYDOWN = &H104
		WM_SYSKEYUP = &H105	'?????
		WM_IME_KEYDOWN = &H290
		WM_IME_KEYUP = &H291
		WM_CLOSE = &H10	'16 EN DECIMAL
		WM_KEYDOWN = &H100
		WM_KEYUP = &H101
		WM_SETTEXT = &HC
		WM_CHAR = &H102
		WM_COMMAND = &H111
		GW_HWNDFIRST = 0&
		GW_HWNDNEXT = 2&
		GW_CHILD = 5&
		WM_MOUSEWHEEL = &H20A	'522 '&H20A
		WM_LBUTTONUP = &H202 '514  '&H202        'Mouse UP
		WM_SYSCOMMAND = &H112	'274 '&H112       'Para indicar que hay que hacer algo (indicar el comando en wParam)
		WM_DESTROY
		LB_FINDSTRING = &H18F
		VK_0 = Asc("0")
		VK_1 = &H31
		VK_2 = &H32
		VK_3 = &H33
		VK_4 = &H34
		VK_5 = &H35
		VK_6 = &H36
		VK_7 = &H37
		VK_8 = &H38
		VK_9 = &H39
		VK_A = &H41
		VK_B = &H42
		VK_C = &H43
		VK_D = &H44
		VK_E = &H45
		VK_F = &H46
		VK_G = &H47
		VK_H = &H48
		VK_I = &H49
		VK_J = &H4A
		VK_K = &H4B
		VK_L = &H4C
		VK_M = &H4D
		VK_N = &H4E
		VK_O = &H4F
		VK_P = &H50
		VK_Q = &H51
		VK_R = &H52
		VK_S = &H53
		VK_T = &H54
		VK_U = &H55
		VK_V = &H56
		VK_W = &H57
		VK_X = &H58
		VK_Y = &H59
		VK_Z = &H5A
		VK_ADD = &H6B
		VK_ATTN = &HF6
		VK_BACK = &H8
		VK_CANCEL = &H3
		VK_CAPITAL = &H14
		VK_CLEAR = &HC
		VK_CONTROL = &H11
		VK_CRSEL = &HF7
		VK_DECIMAL = &H6E
		VK_DELETE = &H2E
		VK_DIVIDE = &H6F
		VK_DOWN = &H28
		VK_END = &H23
		VK_EREOF = &HF9
		VK_ESCAPE = &H1B	'27  'VK_ESCAPE = &H1B
		VK_EXECUTE = &H2B
		VK_EXSEL = &HF8
		VK_F1 = &H70
		VK_F10 = &H79
		VK_F11 = &H7A
		VK_F12 = &H7B
		VK_F13 = &H7C
		VK_F14 = &H7D
		VK_F15 = &H7E
		VK_F16 = &H7F
		VK_F17 = &H80
		VK_F18 = &H81
		VK_F19 = &H82
		VK_F2 = &H71
		VK_F20 = &H83
		VK_F21 = &H84
		VK_F22 = &H85
		VK_F23 = &H86
		VK_F24 = &H87
		VK_F3 = &H72
		VK_F4 = &H73
		VK_F5 = &H74
		VK_F6 = &H75
		VK_F7 = &H76
		VK_F8 = &H77
		VK_F9 = &H78
		VK_HELP = &H2F
		VK_HOME = &H24
		VK_INSERT = &H2D
		VK_LBUTTON = &H1
		VK_LCONTROL = &HA2
		VK_LEFT = &H25
		VK_LMENU = &HA4
		VK_LSHIFT = &HA0
		VK_MBUTTON = &H4
		VK_MENU = &H12
		VK_MULTIPLY = &H6A
		VK_NEXT = &H22
		VK_NONAME = &HFC
		VK_NUMLOCK = &H90
		VK_NUMPAD0 = &H60
		VK_NUMPAD1 = &H61
		VK_NUMPAD2 = &H62
		VK_NUMPAD3 = &H63
		VK_NUMPAD4 = &H64
		VK_NUMPAD5 = &H65
		VK_NUMPAD6 = &H66
		VK_NUMPAD7 = &H67
		VK_NUMPAD8 = &H68
		VK_NUMPAD9 = &H69
		VK_OEM_CLEAR = &HFE
		VK_PA1 = &HFD
		VK_PAUSE = &H13
		VK_PLAY = &HFA
		VK_PRINT = &H2A
		VK_PRIOR = &H21
		VK_PROCESSKEY = &HE5
		VK_RBUTTON = &H2
		VK_RCONTROL = &HA3
		VK_RETURN = &HD
		VK_RIGHT = &H27
		VK_RMENU = &HA5
		VK_RSHIFT = &HA1
		VK_SCROLL = &H91
		VK_SELECT = &H29
		VK_SEPARATOR = &H6C
		VK_SHIFT = &H10
		VK_SNAPSHOT = &H2C
		VK_SPACE = &H20
		VK_SUBTRACT = &H6D
		VK_TAB = &H9
		VK_UP = &H26
		VK_ZOOM = &HFB
	End Enum

End Module

Module teclasVariables
    'integer VK_LBUTTON = 01
    'integer VK_RBUTTON = 02
    'integer VK_CANCEL = 03
    'integer VK_MBUTTON = 04 /* NOT contiguous with L & RBUTTON */

    'integer VK_BACK = 08
    'integer VK_TAB = 09

    'integer VK_CLEAR = 12
    'integer VK_RETURN = 13

    'integer VK_SHIFT = 16
    'integer VK_CONTROL = 17
    'integer VK_MENU = 18
    'integer VK_PAUSE = 19
    'integer VK_CAPITAL = 20

    'integer VK_ESCAPE = 27

    'integer VK_SPACE = 32
    'integer VK_PRIOR = 33
    'integer VK_NEXT = 34
    'integer VK_END = 35
    'integer VK_HOME = 36
    'integer VK_LEFT = 37
    'integer VK_UP = 38
    'integer VK_RIGHT = 39
    'integer VK_DOWN = 40
    'integer VK_SELECT = 21
    'integer VK_PRINT = 42
    'integer VK_EXECUTE = 43
    'integer VK_SNAPSHOT = 44
    'integer VK_INSERT = 45
    'integer VK_DELETE = 46
    'integer VK_HELP = 47


    Public Const VK_0 As Long = Asc("0")
    Public Const VK_1 As Long = &H31
    Public Const VK_2 As Long = &H32
    Public Const VK_3 As Long = &H33
    Public Const VK_4 As Long = &H34
    Public Const VK_5 As Long = &H35
    Public Const VK_6 As Long = &H36
    Public Const VK_7 As Long = &H37
    Public Const VK_8 As Long = &H38
    Public Const VK_9 As Long = &H39
    Public Const VK_A As Long = &H41
    Public Const VK_B As Long = &H42
    Public Const VK_C As Long = &H43
    Public Const VK_D As Long = &H44
    Public Const VK_E As Long = &H45
    Public Const VK_F As Long = &H46
    Public Const VK_G As Long = &H47
    Public Const VK_H As Long = &H48
    Public Const VK_I As Long = &H49
    Public Const VK_J As Long = &H4A
    Public Const VK_K As Long = &H4B
    Public Const VK_L As Long = &H4C
    Public Const VK_M As Long = &H4D
    Public Const VK_N As Long = &H4E
    Public Const VK_O As Long = &H4F
    Public Const VK_P As Long = &H50
    Public Const VK_Q As Long = &H51
    Public Const VK_R As Long = &H52
    Public Const VK_S As Long = &H53
    Public Const VK_T As Long = &H54
    Public Const VK_U As Long = &H55
    Public Const VK_V As Long = &H56
    Public Const VK_W As Long = &H57
    Public Const VK_X As Long = &H58
    Public Const VK_Y As Long = &H59
    Public Const VK_Z As Long = &H5A
    Public Const VK_ADD As Long = &H6B
    Public Const VK_ATTN As Long = &HF6
    Public Const VK_BACK As Long = &H8
    Public Const VK_CANCEL As Long = &H3
    Public Const VK_CAPITAL As Long = &H14
    Public Const VK_CLEAR As Long = &HC
    Public Const VK_CONTROL As Long = &H11
    Public Const VK_CRSEL As Long = &HF7
    Public Const VK_DECIMAL As Long = &H6E
    Public Const VK_DELETE As Long = &H2E
    Public Const VK_DIVIDE As Long = &H6F
    Public Const VK_DOWN As Long = &H28
    Public Const VK_END As Long = &H23
    Public Const VK_EREOF As Long = &HF9
    Public Const VK_ESCAPE As Long = &H1B
    Public Const VK_EXECUTE As Long = &H2B
    Public Const VK_EXSEL As Long = &HF8
    Public Const VK_F1 As Long = &H70
    Public Const VK_F10 As Long = &H79
    Public Const VK_F11 As Long = &H7A
    Public Const VK_F12 As Long = &H7B
    Public Const VK_F13 As Long = &H7C
    Public Const VK_F14 As Long = &H7D
    Public Const VK_F15 As Long = &H7E
    Public Const VK_F16 As Long = &H7F
    Public Const VK_F17 As Long = &H80
    Public Const VK_F18 As Long = &H81
    Public Const VK_F19 As Long = &H82
    Public Const VK_F2 As Long = &H71
    Public Const VK_F20 As Long = &H83
    Public Const VK_F21 As Long = &H84
    Public Const VK_F22 As Long = &H85
    Public Const VK_F23 As Long = &H86
    Public Const VK_F24 As Long = &H87
    Public Const VK_F3 As Long = &H72
    Public Const VK_F4 As Long = &H73
    Public Const VK_F5 As Long = &H74
    Public Const VK_F6 As Long = &H75
    Public Const VK_F7 As Long = &H76
    Public Const VK_F8 As Long = &H77
    Public Const VK_F9 As Long = &H78
    Public Const VK_HELP As Long = &H2F
    Public Const VK_HOME As Long = &H24
    Public Const VK_INSERT As Long = &H2D
    Public Const VK_LBUTTON As Long = &H1
    Public Const VK_LCONTROL As Long = &HA2
    Public Const VK_LEFT As Long = &H25
    Public Const VK_LMENU As Long = &HA4
    Public Const VK_LSHIFT As Long = &HA0
    Public Const VK_MBUTTON As Long = &H4
    Public Const VK_MENU As Long = &H12
    Public Const VK_MULTIPLY As Long = &H6A
    Public Const VK_NEXT As Long = &H22
    Public Const VK_NONAME As Long = &HFC
    Public Const VK_NUMLOCK As Long = &H90
    Public Const VK_NUMPAD0 As Long = &H60
    Public Const VK_NUMPAD1 As Long = &H61
    Public Const VK_NUMPAD2 As Long = &H62
    Public Const VK_NUMPAD3 As Long = &H63
    Public Const VK_NUMPAD4 As Long = &H64
    Public Const VK_NUMPAD5 As Long = &H65
    Public Const VK_NUMPAD6 As Long = &H66
    Public Const VK_NUMPAD7 As Long = &H67
    Public Const VK_NUMPAD8 As Long = &H68
    Public Const VK_NUMPAD9 As Long = &H69
    Public Const VK_OEM_CLEAR As Long = &HFE
    Public Const VK_PA1 As Long = &HFD
    Public Const VK_PAUSE As Long = &H13
    Public Const VK_PLAY As Long = &HFA
    Public Const VK_PRINT As Long = &H2A
    Public Const VK_PRIOR As Long = &H21
    Public Const VK_PROCESSKEY As Long = &HE5
    Public Const VK_RBUTTON As Long = &H2
    Public Const VK_RCONTROL As Long = &HA3
    Public Const VK_RETURN As Long = &HD
    Public Const VK_RIGHT As Long = &H27
    Public Const VK_RMENU As Long = &HA5
    Public Const VK_RSHIFT As Long = &HA1
    Public Const VK_SCROLL As Long = &H91
    Public Const VK_SELECT As Long = &H29
    Public Const VK_SEPARATOR As Long = &H6C
    Public Const VK_SHIFT As Long = &H10
    Public Const VK_SNAPSHOT As Long = &H2C
    Public Const VK_SPACE As Long = &H20
    Public Const VK_SUBTRACT As Long = &H6D
    Public Const VK_TAB As Long = &H9
    Public Const VK_UP As Long = &H26
    Public Const VK_ZOOM As Long = &HFB
End Module

'<System.Runtime.InteropServices.DllImport("user32.DLL")> _
'Public Function SendMessage( _
'ByVal hWnd As Long, ByVal wMsg As teclas, _
'ByVal wParam As Long, ByVal lParam As Long _
') As Long
'End Function
'Para poder enviar lParam como cadenas tambien. En vez de números Long
'<System.Runtime.InteropServices.DllImport("user32.DLL")> _
'Public Function SendMessage( _
'ByVal hWnd As Long, ByVal wMsg As teclas, _
'ByVal wParam As Long, ByVal lParam As String _
') As Long
'End Function

'ByVal lParam As integer

'--------------------------------------------------------------------
'NOTAS:
'Listado a insertar en un módulo (.bas)
'si se quiere poner en un formulario (.frm)
'declarar la función como Private y quitar el Global de las constantes
'--------------------------------------------------------------------
'Constantes y declaración de función:
'
'Constantes para SendMessage
'Public Const WM_LBUTTONUP = &H202
'Public Const WM_SYSCOMMAND = &H112
'Public Const SC_MOVE = &HF010
'Public Const MOUSE_MOVE = &HF012

'#If Win32 Then
'Public Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As teclas, ByVal wParam As Long, ByVal lParam As Long) As Long
'#Else
'Public Declare Function SendMessage Lib "User" (ByVal hWnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, ByVal lParam As Any) As Long
'#End If
'
'
'Este código se pondrá en el Control_MouseDown...
'
'Dim lngRet As Long

'Simular que se mueve la ventana, pulsando en el Control
'If Button = 1 Then
'Envía un MouseUp al Control
'lngRet = SendMessage(Control.hWnd, WM_LBUTTONUP, 0, 0)
'Envía la orden de mover el form
'lngRet = SendMessage(FormX.hWnd, WM_SYSCOMMAND, MOUSE_MOVE, 0)
'End If


'keybd_event VK_RETURN, 0, KEYEVENTF_KEYUP, 0 

'Public Const KEYEVENTF_KEYUP = &H2