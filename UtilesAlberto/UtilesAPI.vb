Imports Microsoft.VisualBasic
Imports System
Imports System.Runtime.InteropServices
Imports System.Collections.Generic
Imports System.ComponentModel
'Imports System.Data
Imports System.Diagnostics
Imports System.Drawing
Imports System.Text
Imports System.Windows.Automation
Imports System.Windows.Forms
Imports System.Windows.Input
Partial Public Class Utiles
    Public Shared ultimaVentana As IntPtr = IntPtr.Zero
    Public Shared ultimoProceso As Integer = 0
    Public Shared comboBoxElement As System.Windows.Automation.AutomationElement = Nothing

    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto)>
    Public Shared Function SetWindowText(ByVal hwnd As IntPtr, ByVal lpString As String) As Boolean
    End Function

    <System.Runtime.InteropServices.DllImport("user32.dll", EntryPoint:="SetWindowLong")>
    Public Shared Function SetWindowLong32(ByVal hWnd As IntPtr, ByVal nIndex As Integer, ByVal dwNewLong As Integer) As Integer
    End Function

    <System.Runtime.InteropServices.DllImport("user32.dll", EntryPoint:="SetWindowLongPtr")>
    Public Shared Function SetWindowLongPtr64(ByVal hWnd As IntPtr, ByVal nIndex As Integer, ByVal dwNewLong As IntPtr) As IntPtr
    End Function

    Public Shared Function SetWindowLongPtr(ByVal hWnd As IntPtr, ByVal nIndex As Integer, ByVal dwNewLong As IntPtr) As IntPtr
        If IntPtr.Size = 8 Then
            Return SetWindowLongPtr64(hWnd, nIndex, dwNewLong)
        Else
            Return New IntPtr(SetWindowLong32(hWnd, nIndex, dwNewLong.ToInt32))
        End If
    End Function

    ' Declaraciones para extraer iconos de los programas
    <System.Runtime.InteropServices.DllImport("shell32.dll")>
    Public Shared Function ExtractIconEx(
      ByVal lpszFile As String, ByVal nIconIndex As Integer,
      ByRef phiconLarge As Integer, ByRef phiconSmall As Integer,
      ByVal nIcons As UInteger) As Integer
    End Function

    <System.Runtime.InteropServices.DllImport("shell32.dll")>
    Public Shared Function ExtractIcon(
      ByVal hInst As Integer, ByVal lpszExeFileName As String,
      ByVal nIconIndex As Integer) As IntPtr
    End Function

    <DllImport("user32.dll")>
    Private Shared Function GetClassWord(
            ByVal hWnd As IntPtr, ByVal nIndex As Integer) As Integer
    End Function

    <DllImport("user32.dll")>
    Private Shared Function GetClassLong(
                ByVal hWnd As IntPtr, ByVal nIndex As Integer) As Integer
    End Function

    Private Const GCW_HMODULE As Integer = (-16)

    Declare Function SHGetImageList Lib "shell32.dll" (ByVal iImageList As Long, ByRef riid As Long, ByRef ppv As Object) As Long
    Declare Function SHGetImageListXP Lib "shell32.dll" Alias "#727" (ByVal iImageList As Long, ByRef riid As Long, ByRef ppv As Object) As Long


    ' Hace que una ventana sea hija (o esté contenida) en otra
    <System.Runtime.InteropServices.DllImport("user32.dll")>
    Public Shared Function SetParent(
        ByVal hWndChild As IntPtr,
        ByVal hWndNewParent As IntPtr) As IntPtr
    End Function

    ' Nos da un array con el texto de la ventana.
    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto)>
    Public Shared Function GetWindowText(
        ByVal hwnd As IntPtr,
        ByVal lpString As Text.StringBuilder,
        ByVal cch As Integer) As Integer
    End Function
    ' Nos da un array con el texto de la ventana.
    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto)>
    Public Shared Function GetWindowText(
        ByVal hwnd As IntPtr,
        ByVal lpString As String,
        ByVal cch As Integer) As Integer
    End Function

    ' Nos da la longitud del texto de la ventana.
    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto)>
    Public Shared Function GetWindowTextLength(ByVal hwnd As IntPtr) As Integer
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

    Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hdc As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As POINTAPI) As Long
    Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long

    'Public Shared Function text_length(Mytext As String) As Long
    '    Dim TextSize As POINTAPI
    '    GetTextExtentPoint32(GetWindowDC(hwnd), Mytext, Len(Mytext), TextSize)
    '    text_length = TextSize.X
    'End Function
    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto)>
    Private Shared Function GetDlgItem(
    ByVal hDlg As IntPtr,
    nIDDlgItem As Integer) As IntPtr
    End Function

    ' Capturar el Handle de la ventana activa.
    <System.Runtime.InteropServices.DllImport("user32.dll", CharSet:=CharSet.Auto, ExactSpelling:=True)>
    Public Shared Function GetActiveWindow() As IntPtr
    End Function

    ' Capturar el Handle de la ventana activa y que está en primer plano.
    <System.Runtime.InteropServices.DllImport("user32.dll", CharSet:=CharSet.Auto, ExactSpelling:=True)>
    Public Shared Function GetForegroundWindow() As IntPtr
    End Function
    '' El Id de proceso de la venta.
    <DllImport("user32.dll", SetLastError:=True)>
    Public Shared Function GetWindowThreadProcessId(ByVal hwnd As IntPtr,
                          ByRef lpdwProcessId As Integer) As Integer
    End Function

    ' Devuelve el Handle (hWnd) de una ventana de la que sabemos el título
    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto)>
    Public Shared Function FindWindow(
     ByVal lpClassName As String,
     ByVal lpWindowName As String) As IntPtr
    End Function

    <DllImport("user32.dll", EntryPoint:="FindWindow", SetLastError:=True, CharSet:=CharSet.Auto)>
    Public Shared Function FindWindowByClass(
     ByVal lpClassName As String,
     ByVal zero As IntPtr) As IntPtr
    End Function

    <DllImport("user32.dll", EntryPoint:="FindWindow", SetLastError:=True, CharSet:=CharSet.Auto)>
    Public Shared Function FindWindowByCaption(
     ByVal zero As IntPtr,
     ByVal lpWindowName As String) As IntPtr
    End Function

    'Busca el Handle o Hwnd del botón hijo de la ventana encontrada con FindWindow 
    <System.Runtime.InteropServices.DllImport("user32.dll")>
    Public Shared Function FindWindowEx(
        ByVal hWnd1 As IntPtr,
        ByVal hWnd2 As IntPtr,
        ByVal className As String,
        ByVal Caption As String) As IntPtr
    End Function
    'ByVal lpsz1 As String, _
    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto)>
    Public Shared Function GetClassName(
        ByVal hWnd As IntPtr,
        ByVal lpClassName As System.Text.StringBuilder,
        ByVal nMaxCount As Integer) As Integer
    End Function

    ' Cambia el tamaño y la posición de una ventana
    <System.Runtime.InteropServices.DllImport("user32.dll")>
    Public Shared Function MoveWindow(ByVal hWnd As IntPtr,
    ByVal x As Integer, ByVal y As Integer,
     ByVal nWidth As Integer, ByVal nHeight As Integer,
     ByVal bRepaint As Integer) As Integer
    End Function

    ' Destruye (cierra y vacia de la memoria) una ventana
    <System.Runtime.InteropServices.DllImport("user32.dll")>
    Public Shared Function DestroyWindow(ByVal hwnd As System.IntPtr) As System.IntPtr
    End Function
    ''
    <DllImport("User32.dll")>
    Private Shared Function EnumChildWindows(ByVal WindowHandle As IntPtr, ByVal Callback As EnumWindowProcess,
        ByVal lParam As IntPtr) As Boolean
    End Function
    Public Delegate Function EnumWindowProcess(ByVal Handle As IntPtr, ByVal Parameter As IntPtr) As Boolean

    Public Shared Function GetChildWindowsArray(ByVal ParentHandle As IntPtr) As IntPtr()
        Dim ChildrenList As New List(Of IntPtr)
        Dim ListHandle As GCHandle = GCHandle.Alloc(ChildrenList)
        Try
            EnumChildWindows(ParentHandle, AddressOf EnumWindow, GCHandle.ToIntPtr(ListHandle))
        Finally
            If ListHandle.IsAllocated Then ListHandle.Free()
        End Try
        Return ChildrenList.ToArray
    End Function
    Public Shared Function GetChildWindowsList(ByVal parent As IntPtr) As List(Of IntPtr)
        Dim result As List(Of IntPtr) = New List(Of IntPtr)()
        Dim listHandle As GCHandle = GCHandle.Alloc(result)
        Try
            Dim childProc As EnumWindowProcess = New EnumWindowProcess(AddressOf EnumWindow)
            EnumChildWindows(parent, childProc, GCHandle.ToIntPtr(listHandle))
        Finally
            If listHandle.IsAllocated Then listHandle.Free()
        End Try
        Return result
    End Function

    Private Shared Function EnumWindow(ByVal handle As IntPtr, ByVal pointer As IntPtr) As Boolean
        Dim gch As GCHandle = GCHandle.FromIntPtr(pointer)
        Dim list As List(Of IntPtr) = TryCast(gch.Target, List(Of IntPtr))

        If list IsNot Nothing Then
            list.Add(handle)
        End If

        Return True
    End Function
    'Private Shared Function EnumWindow(ByVal Handle As IntPtr, ByVal Parameter As IntPtr) As Boolean
    '    Dim ChildrenList As List(Of IntPtr) = GCHandle.FromIntPtr(Parameter).Target
    '    If ChildrenList Is Nothing Then Throw New Exception("GCHandle Target could not be cast as List(Of IntPtr)")
    '    ChildrenList.Add(Handle)
    '    Return True
    'End Function
    'Public Delegate Function EnumWindowProc(ByVal hWnd As IntPtr, ByVal parameter As IntPtr) As Boolean
    Public Delegate Function EnumWindowsDelegate(ByVal hWnd As System.IntPtr, ByVal parametro As Integer) As Boolean
    '
    <System.Runtime.InteropServices.DllImport("user32.dll")>
    Private Shared Function EnumWindows(
            ByVal lpfn As EnumWindowsDelegate,
            ByVal lParam As Integer) As Boolean
    End Function
    '
    <System.Runtime.InteropServices.DllImport("user32.dll")>
    Private Shared Function EnumChildWindows _
            (ByVal hWndParent As System.IntPtr,
            ByVal lpEnumFunc As EnumWindowsDelegate,
            ByVal lParam As Integer) As Integer
    End Function

    '<System.Runtime.InteropServices.DllImport("user32.dll")>
    'Private Shared Function GetWindowText(
    '        ByVal hWnd As System.IntPtr,
    '        ByVal lpString As System.Text.StringBuilder,
    '        ByVal cch As Integer) As Integer
    'End Function

    ' Para EnumWindows
    Public Shared colWin As New System.Collections.Specialized.StringDictionary
    Public Shared Function EnumWindowsProc(ByVal hWnd As System.IntPtr, ByVal parametro As Integer) As Boolean
        ' Esta función "callback" se usará con EnumWindows y EnumChildWindows
        Dim titulo As New System.Text.StringBuilder(New String(" "c, 256))
        Dim ret As Integer
        Dim nombreVentana As String
        '
        ret = GetWindowText(hWnd, titulo, titulo.Length)
        If ret = 0 Then Return True
        '
        nombreVentana = titulo.ToString.Substring(0, ret)
        If nombreVentana <> Nothing AndAlso nombreVentana.Length > 0 Then
            colWin.Add(hWnd.ToString, nombreVentana)
        End If
        '
        Return True
    End Function
    ' Leer colWin después de ejecutar este procedimiento
    Public Shared Sub EnumerarVentanas(Optional lvTop As ListView = Nothing)
        ' Enumera las ventanas principales (TopWindows)
        colWin.Clear()
        EnumWindows(AddressOf EnumWindowsProc, 0)
        '
        If lvTop IsNot Nothing Then
            lvTop.Items.Clear()
            For Each s As String In colWin.Keys
                Dim lvi As ListViewItem = lvTop.Items.Add(s)
                lvi.SubItems.Add(colWin(s))
            Next
        End If
    End Sub
    ' Leer colWin después de ejecutar este procedimiento
    Private Shared Sub enumerarVentanasHijas(ByVal handleParent As System.IntPtr, Optional lvChild As ListView = Nothing)
        ' Enumera las ventanas hijas del handle indicado
        colWin.Clear()
        EnumChildWindows(handleParent, AddressOf EnumWindowsProc, 0)
        '
        If lvChild IsNot Nothing Then
            lvChild.Items.Clear()
            For Each s As String In colWin.Keys
                Dim lvi As ListViewItem = lvChild.Items.Add(s)
                lvi.SubItems.Add(colWin(s))
            Next
        End If
    End Sub


    Public Shared Sub WindowAPI_Click(ByVal hwnd As IntPtr)
        Dim retVal As Long
        retVal = SendMessage(hwnd, eMensajes.WM_LBUTTONDOWN, 0, 0)
        retVal = SendMessage(hwnd, eMensajes.WM_LBUTTONDOWN, 0, 0)
        retVal = SendMessage(hwnd, eMensajes.WM_KEYUP, eMensajes.VK_SPACE, 0)
        retVal = SendMessage(hwnd, eMensajes.WM_LBUTTONUP, 0, 0)
    End Sub
    ''
    '' hWnd = handle of destination window
    '' wMsg = message to send   (eMensaje u Object para enviar un integer)
    '' wParam = first message parameter
    '' lParam = second message parameter
    <System.Runtime.InteropServices.DllImport("user32.DLL")>
    Public Shared Function SendMessage(
        ByVal hWnd As HandleRef,
        ByVal wMsg As eMensajes,
        ByVal wParam As Integer,
        ByVal lParam As Integer) As Integer
    End Function
    <System.Runtime.InteropServices.DllImport("user32.DLL")>
    Public Shared Function SendMessage(
        ByVal hWnd As HandleRef,
        ByVal wMsg As eMensajes,
        ByVal wParam As Integer,
        ByVal lParam As Object) As Integer
    End Function
    <System.Runtime.InteropServices.DllImport("user32.DLL")>
    Public Shared Function SendMessage(
        ByVal hWnd As System.IntPtr,
        ByVal wMsg As eMensajes,
        ByVal wParam As Integer,
        ByVal lParam As Integer) As Integer
    End Function
    '' Si queremos usar un String en el último parámetro y queremos usarla al mismo tiempo que la anterior,
    '' sólo hay que declararla nuevamente con los parámetros diferentes, (sin necesidad de cambiar el nombre)
    <System.Runtime.InteropServices.DllImport("user32.DLL")>
    Public Shared Function SendMessage(
        ByVal hWnd As System.IntPtr,
        ByVal wMsg As eMensajes,
        ByVal wParam As Integer,
        ByVal lParam As Object
        ) As Integer
    End Function
    '' Si queremos usar un String en el último parámetro y queremos usarla al mismo tiempo que la anterior,
    '' sólo hay que declararla nuevamente con los parámetros diferentes, (sin necesidad de cambiar el nombre)
    <System.Runtime.InteropServices.DllImport("user32.DLL", CharSet:=CharSet.Auto)>
    Public Shared Function SendMessage(
        ByVal hWnd As System.IntPtr,
        ByVal wMsg As eMensajes,
        ByVal wParam As Integer,
        ByVal lParam As String
        ) As Integer
    End Function
    <DllImport("user32.dll", CharSet:=CharSet.Auto)>
    Public Shared Function SendMessage(
        ByVal hWnd As IntPtr,
        ByVal wMsg As eMensajes,
        ByVal wParam As Integer,
        ByVal lParam As StringBuilder) _
        As IntPtr
    End Function
    '
    '<System.Runtime.InteropServices.DllImport("user32.DLL", CharSet:=CharSet.Auto)> _
    Private Declare Function SendMessage_Long Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef LParam As Long) As Integer

    <DllImport("user32.DLL")>
    Public Shared Sub keybd_event(bVk As Byte, bScan As Byte, dwFlags As UInteger, dwExtraInfo As UIntPtr)
    End Sub
    '
    '' No espera a que termine. SendMessage si espera.
    <System.Runtime.InteropServices.DllImport("user32.dll")>
    Public Shared Function PostMessage(
        ByVal hwnd As System.IntPtr,
        ByVal wMsg As eMensajes,
        ByVal wParam As Integer,
        ByVal lParam As Integer) As Integer
    End Function

    <DllImport("user32.dll", SetLastError:=True)>
    Public Shared Function PostMessage(
        ByVal hWnd As IntPtr,
        <MarshalAs(UnmanagedType.U4)> ByVal Msg As UInteger,
        ByVal wParam As IntPtr,
        ByVal lParam As IntPtr) As Boolean
    End Function
    <DllImport("user32.dll", SetLastError:=True)>
    Public Shared Function PostMessage(
        ByVal hWnd As IntPtr,
        ByVal Msg As eMensajes,
        ByVal wParam As IntPtr,
        ByVal lParam As IntPtr) As Boolean
    End Function
    <System.Runtime.InteropServices.DllImport("user32.DLL")>
    Public Shared Function PostMessage(
        ByVal hWnd As IntPtr,
        ByVal wMsg As eMensajes,
        ByVal wParam As Integer,
        ByVal lParam As String) As Integer
    End Function
    '' Ejemplo PostMessage
    'Private Sub TestExample(ByVal hWndTextBox as Int32, ByVal TextString as String)
    '####Using the hWnd of a Combo box control and a string to enter in the combo box.####    
    'Dim TextLength As Int32
    'Dim ItemText As String

    'Do Until ItemText = TextString
    '    SendMessage(hWndTextBox, WM_SETTEXT, 0, TextString) > 0
    '    TextLength = SendMessage(hWndTextBox, WM_GETTEXTLENGTH, 0, 0)
    '    ItemText = Space(TextLength)
    '    SendMessage(hWndTextBox, WM_GETTEXT, TextLen + 1, ItemText)      
    'Loop
    'End Sub
    ''
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

    <DllImport("user32.dll", SetLastError:=True)>
    Public Shared Function GetWindowRect(ByVal hWnd As IntPtr, <Out> ByRef lpRect As RECT) As Boolean
    End Function
    <System.Runtime.InteropServices.DllImport("user32.dll")>
    Public Shared Function GetWindowPlacement(
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
    <System.Runtime.InteropServices.DllImport("user32.dll")>
    Public Shared Function SetWindowWord(
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
            If MoverC = True Then System.Windows.Forms.Cursor.Position = New System.Drawing.Point(100, CInt(System.Windows.Forms.Screen.PrimaryScreen.WorkingArea.Height / 2))
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
    Public Shared Function SetForegroundWindow(
        ByVal hWnd As System.IntPtr) As Integer
    End Function
    ''
    <DllImport("user32.dll", SetLastError:=True)>
    Public Shared Function IsWindowVisible(ByVal hWnd As IntPtr) As <MarshalAs(UnmanagedType.Bool)> Boolean
    End Function
    ''
    <DllImport("user32.dll", SetLastError:=True)>
    Public Shared Function GetShellWindow() As IntPtr
    End Function
    ' Para que la ventanta tenga el foco del teclado
    '<System.Runtime.InteropServices.DllImport("user32.dll")> _
    'Public Declare Function SetFocus Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long

    <System.Runtime.InteropServices.DllImport("user32.dll")>
    Public Shared Function SetFocus(
        ByVal hwnd As IntPtr) As Long
    End Function


    Public Shared Function DameIntPtr(ByVal queApp As String) As IntPtr
        Dim intP As IntPtr = Nothing
        intP = FindWindow(Nothing, queApp)
        DameIntPtr = intP
    End Function


    Public Structure COPYDATASTRUCT
        Dim dwData As Long
        Dim cbData As Long
        Dim lpData As Long
    End Structure

    Public Declare Sub CopyMemory _
    Lib "kernel32" _
    Alias "RtlMoveMemory" (ByVal hpvDest As Object,
    ByVal hpvSource As Object,
    ByVal cbCopy As Long)

    Public Shared Function SendCmd(ByVal sCommand As String, ByVal ventana As IntPtr) As Long

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
    Public Shared Function ActivaProceso(nombrePro As String) As IntPtr
        Dim resultado As IntPtr = IntPtr.Zero
        ''
        '' Cerrar firefox, Google Chrome o Internet Explorer.
        Dim oProc As Process() = Process.GetProcesses   '.GetProcessesByName(nombrePro)
        If oProc IsNot Nothing AndAlso oProc.Count > 0 Then
            For Each oPr As Process In oProc
                If oPr.ProcessName.ToUpper <> nombrePro.ToUpper Then Continue For
                Try
                    Call SetForegroundWindow(oPr.MainWindowHandle)
                    resultado = oPr.MainWindowHandle
                    Exit For
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
        ''Debug.Print(GetText(vInforme))   ' & " / " & g)
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
        ''Debug.Print(GetText(vInforme))   ' & " / " & g)
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
        Dim mensaje As String = GetText(ventanaPadre) & vbCrLf
        IO.File.WriteAllText(queFi, mensaje)
        Dim nivel As Integer = 1
        ''
        ''Debug.Print(GetText(vInforme))   ' & " / " & g)
        For Each vhija As IntPtr In GetChildWindowsList(ventanaPadre)
            DameVentanasHijaRecursivo(vhija, nivel, queFi)
            System.Windows.Forms.Application.DoEvents()
        Next
        ''
        Return mensaje
    End Function
    ''
    ''
    Public Shared Sub DameVentanasHijaRecursivo(ventanaPadre As IntPtr, nivel As Integer, queFi As String)
        Dim queTexto As String = GetText(ventanaPadre)
        If queTexto IsNot Nothing AndAlso queTexto <> "" Then
            IO.File.AppendAllText(queFi, StrDup(nivel, vbTab) & queTexto & vbCrLf)
        End If
        ''
        ''Debug.Print(GetText(vInforme))   ' & " / " & g)
        For Each vhija As IntPtr In GetChildWindowsList(ventanaPadre)
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
        Dim ventana0 As System.IntPtr = GetForegroundWindow()  '  FindWindow("Afx:000000013F1C0000:8:0000000000010003:0000000000000000:0000000000500D79", "Autodesk Inventor Professional 2015") 'oProIn.Handle ' FindWindow("#32769 (Escritorio)", "")
        ultimaVentana = ventana0
        If ventana0 = IntPtr.Zero Then
            Return retVal
            Exit Function
        End If
        ''Debug.Print(GetText(vInforme))   ' & " / " & g)
        For Each vhija As IntPtr In GetChildWindowsList(ventana0)
            Dim texto As String = GetText(vhija)
            If Not (texto Is Nothing) AndAlso texto <> "" AndAlso (texto = nombreBoton Or texto.StartsWith(nombreBoton)) Then
                'retVal = SendMessage(vhija, eMensajes.WM_LBUTTONDOWN, 0, 0)
                'retVal = SendMessage(vhija, eMensajes.WM_LBUTTONDOWN, 0, 0)
                'retVal = SendMessage(vhija, eMensajes.WM_KEYDOWN, eMensajes.VK_SPACE, 0)
                'retVal = SendMessage(vhija, eMensajes.WM_LBUTTONUP, 0, 0)
                'If esperar Then
                'retVal = PostMessage(vhija, eMensajes.WM_click, 0, 0)
                'Else
                'retVal = SendMessage(vhija, eMensajes.WM_click, 0, 0)
                'End If
                '' PostMessage espera a que se complete la tarea.
                ''retVal = SendMessage(ventana0, eMensajes.WM_SYSCOMMAND, eMensajes.CLIK_BUTTON, vhija)
                'System.Windows.Forms.SendKeys.Send("{ENTER}")
                'System.Windows.Forms.SendKeys.Send("{ENTER}")
                'retVal = SendMessage(vhija, eMensajes.WM_ACTIVATE, eMensajes.WA_ACTIVE, 0)
                'retVal = SendMessage(vhija, eMensajes.BM_CLICK, 0, 0)
                retVal = SendMessage(vhija, eMensajes.WM_KEYDOWN, eMensajes.VK_SPACE, 0)
                retVal = SendMessage(vhija, eMensajes.WM_KEYUP, eMensajes.VK_SPACE, 0)
                Exit For
            End If
            For Each vhija1 As IntPtr In GetChildWindowsList(vhija)
                Dim texto1 As String = GetText(vhija1)
                If Not (texto1 Is Nothing) AndAlso texto1 <> "" AndAlso (texto1 = nombreBoton Or texto1.StartsWith(nombreBoton)) Then
                    'If esperar Then
                    'retVal = PostMessage(vhija1, eMensajes.WM_click, 0, 0)
                    'Else
                    'retVal = SendMessage(vhija1, eMensajes.WM_click, 0, 0)
                    'End If
                    'System.Windows.Forms.SendKeys.Send("{ENTER}")
                    'System.Windows.Forms.SendKeys.Send("{ENTER}")
                    'retVal = SendMessage(vhija1, eMensajes.WM_ACTIVATE, eMensajes.WA_ACTIVE, 0)
                    'retVal = SendMessage(vhija1, eMensajes.BM_CLICK, 0, 0)
                    retVal = SendMessage(vhija1, eMensajes.WM_KEYDOWN, eMensajes.VK_SPACE, 0)
                    retVal = SendMessage(vhija1, eMensajes.WM_KEYUP, eMensajes.VK_SPACE, 0)
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
        Dim ventana0 As System.IntPtr = FindWindow("", nombreVentana) 'oProIn.Handle ' FindWindow("#32769 (Escritorio)", "")
        ultimaVentana = ventana0
        ''Debug.Print(GetText(vInforme))   ' & " / " & g)
        If ventana0 = IntPtr.Zero Then
            Return retVal
            Exit Function
        End If
        For Each vhija As IntPtr In GetChildWindowsList(ventana0)
            Dim texto As String = GetText(vhija)
            If Not (texto Is Nothing) AndAlso texto <> "" AndAlso (texto = nombreBoton Or texto.StartsWith(nombreBoton)) Then
                'retVal = SendMessage(vhija, eMensajes.WM_LBUTTONDOWN, 0, 0)
                'retVal = SendMessage(vhija, eMensajes.WM_LBUTTONDOWN, 0, 0)
                'retVal = SendMessage(vhija, eMensajes.WM_KEYDOWN, eMensajes.VK_SPACE, 0)
                'retVal = SendMessage(vhija, eMensajes.WM_LBUTTONUP, 0, 0)
                'If esperar Then
                'retVal = PostMessage(vhija, eMensajes.WM_click, 0, 0)
                'Else
                'retVal = SendMessage(vhija, eMensajes.WM_click, 0, 0)
                'End If
                '' PostMessage espera a que se complete la tarea.
                ''retVal = SendMessage(ventana0, eMensajes.WM_SYSCOMMAND, eMensajes.CLIK_BUTTON, vhija)
                'System.Windows.Forms.SendKeys.Send("{ENTER}")
                'System.Windows.Forms.SendKeys.Send("{ENTER}")
                'retVal = SendMessage(vhija, eMensajes.WM_ACTIVATE, eMensajes.WA_ACTIVE, 0)
                'retVal = SendMessage(vhija, eMensajes.BM_CLICK, 0, 0)
                retVal = SendMessage(vhija, eMensajes.WM_KEYDOWN, ConsoleKey.Enter, 0)
                retVal = SendMessage(vhija, eMensajes.WM_KEYUP, ConsoleKey.Enter, 0)
                Exit For
            End If
            For Each vhija1 As IntPtr In GetChildWindowsList(vhija)
                Dim texto1 As String = GetText(vhija1)
                If Not (texto1 Is Nothing) AndAlso texto1 <> "" AndAlso (texto1 = nombreBoton Or texto1.StartsWith(nombreBoton)) Then
                    'If esperar Then
                    'retVal = PostMessage(vhija1, eMensajes.WM_click, 0, 0)
                    'Else
                    'retVal = SendMessage(vhija1, eMensajes.WM_click, 0, 0)
                    'End If
                    'System.Windows.Forms.SendKeys.Send("{ENTER}")
                    'System.Windows.Forms.SendKeys.Send("{ENTER}")
                    'retVal = SendMessage(vhija1, eMensajes.WM_ACTIVATE, eMensajes.WA_ACTIVE, 0)
                    'retVal = SendMessage(vhija1, eMensajes.BM_CLICK, 0, 0)
                    retVal = SendMessage(vhija1, eMensajes.WM_KEYDOWN, eMensajes.VK_SPACE, 0)
                    retVal = SendMessage(vhija1, eMensajes.WM_KEYUP, eMensajes.VK_SPACE, 0)
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
    Public Shared Function PulsaEnterEnVentanaActiva(ventanaactiva As IntPtr, Optional esperar As Boolean = False) As Long
        ''
        '' Para esperar, utilizar PostMessage en vez de SendMessage
        Dim ventana0 As System.IntPtr = ventanaactiva
        Dim retVal As Long = 0
        If ventanaactiva = IntPtr.Zero Then
            ventana0 = GetForegroundWindow()
        End If
        ''
        'If oProceso Is Nothing Then oProceso = Process.GetCurrentProcess
        'Dim ventana0 As System.IntPtr = oProceso.MainWindowHandle ' GetForegroundWindow
        SetForegroundWindow(ventana0)
        If ultimaVentana <> ventana0 Then ultimaVentana = ventana0
        ''Debug.Print(GetText(vInforme))   ' & " / " & g)
        If ventana0 = IntPtr.Zero Then
            Return retVal
            Exit Function
        End If
        ''
        If esperar = True Then
            retVal = PostMessage(ventana0, eMensajes.WM_KEYDOWN, ConsoleKey.Enter, 0)
            retVal = PostMessage(ventana0, eMensajes.WM_KEYUP, ConsoleKey.Enter, 0)
        Else
            retVal = SendMessage(ventana0, eMensajes.WM_KEYDOWN, ConsoleKey.Enter, 0)
            retVal = SendMessage(ventana0, eMensajes.WM_KEYUP, ConsoleKey.Enter, 0)
        End If
        ''
        Return retVal
    End Function
    ''
    Public Shared Function EscribeEnTextBoxBuscaFinal(textoBusco As String, textoEscribo As String) As Long
        Dim retVal As Long = 0
        ''
        'Dim oProceso As Process = Process.GetCurrentProcess
        Dim ventana0 As System.IntPtr = GetForegroundWindow()  '  FindWindow("Afx:000000013F1C0000:8:0000000000010003:0000000000000000:0000000000500D79", "Autodesk Inventor Professional 2015") 'oProIn.Handle ' FindWindow("#32769 (Escritorio)", "")
        ultimaVentana = ventana0
        ''Debug.Print(GetText(vInforme))   ' & " / " & g)
        For Each vhija As IntPtr In GetChildWindowsList(ventana0)
            Dim texto As String = GetText(vhija)
            'Dim k As Integer
            'Dim L1 As Integer, ltexto As Integer
            '' K retorna el Número de líneas del TextBox  
            'k = SendMessage(vhija, eMensajes.EM_GETLINECOUNT, 0, 0&)
            ' L1 el Primer carácter de la línea actual  
            'L1 = SendMessage(vhija, eMensajes.EM_LINEINDEX, k - 1, 0&) + 1
            ' Longitud de la línea actual (Cantidad de caracteres )  
            'ltexto = SendMessage(vhija, eMensajes.EM_LINELENGTH, L1, 0&)
            ' Mostramos la ultima línea del textbox  
            'MsgBox(Mid$(Text1.Text, L1, L2), vbInformation)
            'EM_LINELENGTH
            If Not (texto Is Nothing) AndAlso texto <> "" AndAlso texto.ToLower.EndsWith(textoBusco.ToLower) Then
                '' Primero borrar el texto actual
                retVal = SendMessage(vhija, eMensajes.WM_SETTEXT, 255, StrDup(255, " "))
                '' Segundo poner el texto definitivo
                retVal = SendMessage(vhija, eMensajes.WM_SETTEXT, textoEscribo.Length, textoEscribo)
                Exit For
            End If
            For Each vhija1 As IntPtr In GetChildWindowsList(vhija)
                Dim texto1 As String = GetText(vhija1)
                'Dim k1 As Integer
                'Dim L1a As Integer, ltexto1 As Integer
                '' K retorna el Número de líneas del TextBox  
                'k1 = SendMessage(vhija1, eMensajes.EM_GETLINECOUNT, 0, 0&)
                ' L1 el Primer carácter de la línea actual  
                'L1a = SendMessage(vhija1, eMensajes.EM_LINEINDEX, k1 - 1, 0&) + 1
                ' Longitud de la línea actual (Cantidad de caracteres )  
                'ltexto1 = SendMessage(vhija1, eMensajes.EM_LINELENGTH, L1a, 0&)
                If Not (texto1 Is Nothing) AndAlso texto1 <> "" AndAlso texto1.ToLower.EndsWith(textoBusco.ToLower) Then
                    '' Primero borrar el texto actual
                    retVal = SendMessage(vhija1, eMensajes.WM_SETTEXT, 255, StrDup(255, " "))
                    '' Segundo poner el texto definitivo
                    retVal = SendMessage(vhija1, eMensajes.WM_SETTEXT, textoEscribo.Length, textoEscribo)
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
        Dim ventana0 As System.IntPtr = GetForegroundWindow()  '  FindWindow("Afx:000000013F1C0000:8:0000000000010003:0000000000000000:0000000000500D79", "Autodesk Inventor Professional 2015") 'oProIn.Handle ' FindWindow("#32769 (Escritorio)", "")
        ultimaVentana = ventana0
        ''Debug.Print(GetText(vInforme))   ' & " / " & g)
        For Each vhija As IntPtr In GetChildWindowsList(ventana0)
            Dim texto As String = GetText(vhija)
            If Not (texto Is Nothing) AndAlso texto <> "" AndAlso texto.StartsWith(textoBusco) Then
                retVal = SendMessage(vhija, eMensajes.WM_SETTEXT, textoEscribo.Length, textoEscribo)
                Exit For
            End If
            For Each vhija1 As IntPtr In GetChildWindowsList(vhija)
                Dim texto1 As String = GetText(vhija1)
                If Not (texto1 Is Nothing) AndAlso texto1 <> "" AndAlso texto1.StartsWith(textoBusco) Then
                    retVal = SendMessage(vhija1, eMensajes.WM_SETTEXT, textoEscribo.Length, textoEscribo)
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
        Dim ventana0 As System.IntPtr = GetForegroundWindow()  '  FindWindow("Afx:000000013F1C0000:8:0000000000010003:0000000000000000:0000000000500D79", "Autodesk Inventor Professional 2015") 'oProIn.Handle ' FindWindow("#32769 (Escritorio)", "")
        ultimaVentana = ventana0
        ''Debug.Print(GetText(vInforme))   ' & " / " & g)
        For Each vhija As IntPtr In GetChildWindowsList(ventana0)
            Dim texto As String = GetText(vhija)
            If Not (texto Is Nothing) AndAlso texto <> "" AndAlso (texto = textoBusco Or texto.Contains(textoBusco)) Then
                retVal = SendMessage(vhija, eMensajes.WM_SETTEXT, textoEscribo.Length, textoEscribo)
                Exit For
            End If
            For Each vhija1 As IntPtr In GetChildWindowsList(vhija)
                Dim texto1 As String = GetText(vhija1)
                If Not (texto1 Is Nothing) AndAlso texto1 <> "" AndAlso (texto1 = textoBusco Or texto1.Contains(textoBusco)) Then
                    retVal = SendMessage(vhija1, eMensajes.WM_SETTEXT, textoEscribo.Length, textoEscribo)
                    Exit For
                End If
            Next
        Next
        ''
        Return retVal
    End Function
    '
    Public Shared Function DameTextoBoxBuscaIgualContine(textoBusco As String) As IntPtr
        Dim retVal As IntPtr = IntPtr.Zero
        ''
        Dim oProceso As Process = Process.GetCurrentProcess
        Dim ventana0 As System.IntPtr = GetForegroundWindow()  '  FindWindow("Afx:000000013F1C0000:8:0000000000010003:0000000000000000:0000000000500D79", "Autodesk Inventor Professional 2015") 'oProIn.Handle ' FindWindow("#32769 (Escritorio)", "")
        ultimaVentana = ventana0
        ''Debug.Print(GetText(vInforme))   ' & " / " & g)
        For Each vhija As IntPtr In GetChildWindowsList(ventana0)
            Dim texto As String = GetText(vhija)
            If Not (texto Is Nothing) AndAlso texto <> "" AndAlso (texto = textoBusco Or texto.Contains(textoBusco)) Then
                Dim titulo As New System.Text.StringBuilder(New String(" "c, 256))
                retVal = SendMessage(vhija, eMensajes.WM_GETTEXT, titulo.Length, titulo)
                Exit For
            End If
            For Each vhija1 As IntPtr In GetChildWindowsList(vhija)
                Dim texto1 As String = GetText(vhija1)
                If Not (texto1 Is Nothing) AndAlso texto1 <> "" AndAlso (texto1 = textoBusco Or texto1.Contains(textoBusco)) Then
                    Dim titulo1 As New System.Text.StringBuilder(New String(" "c, 256))
                    retVal = SendMessage(vhija, eMensajes.WM_GETTEXT, titulo1.Length, titulo1)
                    Exit For
                End If
            Next
        Next
        ''
        Return retVal
    End Function
    ' Buscar en Save As--ComboBox-Edit
    Public Shared Function DameTexto_SaveAs() As String
        Dim retVal As String = ""
        ''
        Dim oProceso As Process = Process.GetCurrentProcess
        Dim ventana0 As System.IntPtr = GetForegroundWindow()  '  FindWindow("Afx:000000013F1C0000:8:0000000000010003:0000000000000000:0000000000500D79", "Autodesk Inventor Professional 2015") 'oProIn.Handle ' FindWindow("#32769 (Escritorio)", "")
        ultimaVentana = ventana0
        Dim texto As String = ""
        Dim sizeText As Integer = 0
        ''Debug.Print(GetText(vInforme))   ' & " / " & g)
        For Each vhija As IntPtr In GetChildWindowsList(ventana0)
            Dim className As StringBuilder = New StringBuilder(100)
            Dim classIdName As Integer = GetClassName(vhija, className, 100)
            If className.ToString.Contains("ComboBox") Then
                'Retorna el tamaño del caption ( los caracteres )  
                sizeText = SendMessage(vhija, eMensajes.WM_GETTEXTLENGTH, 0, 0)
                retVal = SendMessage(vhija, eMensajes.WM_GETTEXT, sizeText, texto).ToString.Trim
                If retVal <> "" Then Exit For
                For Each vhija1 As IntPtr In GetChildWindowsList(vhija)
                    Dim className1 As StringBuilder = New StringBuilder(100)
                    Dim classIdName1 As Integer = GetClassName(vhija1, className1, 100)
                    texto = GetText(vhija1)
                    If className1.ToString.Contains("Edit") Then
                        'Retorna el tamaño del caption ( los caracteres )  
                        sizeText = SendMessage(vhija1, eMensajes.WM_GETTEXTLENGTH, 0, 0)
                        retVal = SendMessage(vhija, eMensajes.WM_GETTEXT, sizeText, texto).ToString.Trim
                        Exit For
                    End If
                Next
            End If
            If retVal <> "" Then Exit For
        Next
        ''
        Return retVal
    End Function
    ' Buscar en Save As--ComboBox-Edit
    Public Shared Function Busca_SaveAs_Revit(Optional queBusco As String = ".rvt", Optional conMensaje As Boolean = False) As IntPtr
        Dim resultado As IntPtr = IntPtr.Zero
        ' Buscamos en el escritorio
        Dim mensaje As String = ""
        For Each vhija As IntPtr In GetChildWindowsList(IntPtr.Zero)
            Dim className As StringBuilder = New StringBuilder(100)
            Dim classIdName As Integer = GetClassName(vhija, className, 100)
            If className.ToString.StartsWith("#32770") Then
                mensaje &= className.ToString & vbCrLf
                Dim textoEdit As String = DameTexto_SaveAs(vhija).ToString.Trim
                If textoEdit.Contains(queBusco) Or textoEdit = queBusco Then
                    mensaje &= vbTab & vhija.ToString & " (" & DameTexto_SaveAs(vhija) & ")" & vbCrLf
                    resultado = vhija
                    Exit For
                Else
                    mensaje &= vbTab & DameTexto_SaveAs(vhija) & vbCrLf
                End If
            End If
        Next
        '
        If conMensaje = True And mensaje <> "" Then MsgBox(mensaje)
        Return resultado
    End Function
    Public Shared Function DameTexto_SaveAs(queIntPrt As IntPtr) As String
        Dim retVal As String = ""
        ' Comprobar si es dialogo Save As (Class name empieza por "#32770")
        Dim className0 As StringBuilder = New StringBuilder(100)
        Dim classIdName0 As Integer = GetClassName(queIntPrt, className0, 100)
        If className0.ToString.Contains("#32770") = False Then
            Return retVal
            Exit Function
        End If
        '
        If ultimaVentana <> queIntPrt Then
            ultimaVentana = queIntPrt
        End If
        Dim texto As String = ""
        Dim sizeText As Integer = 0
        ''Debug.Print(GetText(vInforme))   ' & " / " & g)
        For Each vhija As IntPtr In GetChildWindowsList(queIntPrt)
            Dim className As StringBuilder = New StringBuilder(100)
            Dim classIdName As Integer = GetClassName(vhija, className, 100)
            If className.ToString.Contains("ComboBox") Then
                'Retorna el tamaño del caption ( los caracteres )  
                sizeText = SendMessage(vhija, eMensajes.WM_GETTEXTLENGTH, 0, 0)
                'Buffer para almacenar el valor devuelto por SendMessage  
                retVal = Space$(sizeText + 1)
                sizeText = SendMessage(vhija, eMensajes.WM_GETTEXT, sizeText + 1, retVal)
                'If retVal IsNot Nothing AndAlso retVal <> "" Then
                '    Exit For
                'End If
                For Each vhija1 As IntPtr In GetChildWindowsList(vhija)
                    Dim className1 As StringBuilder = New StringBuilder(100)
                    Dim classIdName1 As Integer = GetClassName(vhija1, className1, 100)
                    If className1.ToString.Contains("Edit") Then
                        'Retorna el tamaño del caption ( los caracteres )  
                        sizeText = SendMessage(vhija1, eMensajes.WM_GETTEXTLENGTH, 0, 0)
                        'Buffer para almacenar el valor devuelto por SendMessage  
                        retVal = Space$(sizeText + 1)
                        sizeText = SendMessage(vhija1, eMensajes.WM_GETTEXT, sizeText + 1, retVal)
                        Exit For
                    End If
                Next
            End If
            If retVal IsNot Nothing AndAlso retVal <> "" Then
                Exit For
            End If
        Next
        ''
        If retVal IsNot Nothing Then
            Return retVal.Trim
        Else
            Return ""
        End If
    End Function

    ' Buscar en Save As--ComboBox-Edit
    Public Shared Function EscribeTexto_SaveAsEdit(intPtrDialog As IntPtr, textoEscribo As String, Optional pulsaEnter As Boolean = False, Optional esperar As Boolean = False) As Integer
        Dim retVal As Integer = 0
        If ultimaVentana <> intPtrDialog Then
            ultimaVentana = intPtrDialog
        End If
        Dim texto As String = ""
        Dim sizeText As Integer = 0
        Dim intPtrComboBox As IntPtr = IntPtr.Zero
        Dim intPtrEdit As IntPtr = IntPtr.Zero
        ''Debug.Print(GetText(vInforme))   ' & " / " & g)
        For Each vhija As IntPtr In GetChildWindowsList(intPtrDialog)
            Dim className As StringBuilder = New StringBuilder(100)
            Dim classIdName As Integer = GetClassName(vhija, className, 100)
            If className.ToString.Contains("ComboBox") Then
                intPtrComboBox = vhija
                '' Primero borrar el texto actual
                retVal = SendMessage(vhija, eMensajes.WM_SETTEXT, 255, StrDup(255, " "))
                '' Segundo poner el texto definitivo
                retVal = SendMessage(vhija, eMensajes.WM_SETTEXT, textoEscribo.Length, textoEscribo)
                ' Si se ha escrito en el ComboBox correctamente
                'If retVal > 0 Then
                '    Exit For
                'End If
                For Each vhija1 As IntPtr In GetChildWindowsList(vhija)
                    Dim className1 As StringBuilder = New StringBuilder(100)
                    Dim classIdName1 As Integer = GetClassName(vhija1, className1, 100)
                    If className1.ToString.Contains("Edit") Then
                        '' Primero borrar el texto actual
                        retVal = SendMessage(vhija, eMensajes.WM_SETTEXT, 255, StrDup(255, " "))
                        '' Segundo poner el texto definitivo
                        retVal = SendMessage(vhija, eMensajes.WM_SETTEXT, textoEscribo.Length, textoEscribo)
                        If retVal > 0 Then
                            intPtrEdit = vhija1
                        End If
                        Exit For
                    End If
                Next
            End If
            If retVal > 0 Then
                Exit For
            End If
        Next
        ''
        Dim res As Integer = SendMessage(intPtrDialog, eMensajes.CDN_FOLDERCHANGE, 0, 0)
        If pulsaEnter And intPtrEdit <> IntPtr.Zero Then
            'SetForegroundWindow(intPtrEdit)
            SendMessage(intPtrEdit, eMensajes.WA_ACTIVE, 0, 0)
            SendKeys.SendWait("{ENTER}")
            ''
            'SetForegroundWindow(intPtrComboBox)
            'SendMessage(queIntPrt, eMensajes.BM_CLICK, 0, 0)
            'SendMessage(queIntPrt, eMensajes.WM_ACTIVATE, eMensajes.WA_ACTIVE, 0)
            'SendMessage(queIntPrt, eMensajes.BM_CLICK, 0, 0)
            'Call PulsaEnterEnVentanaActiva(intPtrHija, True)
            'SendCmd("{ENTER}", intPtrHija)
            'PulsaEnterEnVentanaActiva(intPtrEdit, esperar)
            'PulsaEnterEnVentanaActiva(intPtrComboBox, True)
        End If
        Return retVal
    End Function
    '
    ' Buscar en Save As--ComboBox-Edit
    Public Shared Function EscribeTexto_SaveAsComboBoxFolders(intPtrDialog As IntPtr, queDirectorio As String) As Integer
        Dim className0 As StringBuilder = New StringBuilder(100)
        Dim classIdName0 As Integer = GetClassName(intPtrDialog, className0, 100)
        ' Si no es un SaveAs Dialog (#32770) salir sin hacer nada.
        If className0.ToString.Contains("#32770") = False Then
            Return 0
            Exit Function
        End If
        '
        Dim retVal As Integer = 0
        If ultimaVentana <> intPtrDialog Then
            ultimaVentana = intPtrDialog
        End If
        Dim texto As String = ""
        Dim sizeText As Integer = 0
        Dim intPtrComboBox As IntPtr = IntPtr.Zero
        Dim intPtrEdit As IntPtr = IntPtr.Zero
        ''Debug.Print(GetText(vInforme))   ' & " / " & g)
        For Each vhija As IntPtr In GetChildWindowsList(intPtrDialog)
            Dim className As StringBuilder = New StringBuilder(100)
            Dim classIdName As Integer = GetClassName(vhija, className, 100)
            If className.ToString.Contains("ComboBox") Then
                intPtrComboBox = vhija
                If GetChildWindowsList(vhija).Count > 0 Then
                    Continue For
                End If
                '' Primero borrar el texto actual
                'retVal = SendMessage(vhija, eMensajes.WM_SETTEXT, 255, StrDup(255, " "))
                '' Segundo poner el texto definitivo
                retVal = SendMessage(vhija, eMensajes.WM_SETTEXT, queDirectorio.Length + 1, queDirectorio)
                ' Si se ha escrito en el ComboBox correctamente
                'If retVal > 0 Then
                '    Exit For
                'End If
                'For Each vhija1 As IntPtr In GetChildWindowsList(vhija)
                '    Dim className1 As StringBuilder = New StringBuilder(100)
                '    Dim classIdName1 As Integer = GetClassName(vhija1, className1, 100)
                '    If className1.ToString.Contains("Edit") Then
                '        '' Primero borrar el texto actual
                '        retVal = SendMessage(vhija, eMensajes.WM_SETTEXT, 255, StrDup(255, " "))
                '        '' Segundo poner el texto definitivo
                '        retVal = SendMessage(vhija, eMensajes.WM_SETTEXT, queDirectorio.Length, queDirectorio)
                '        If retVal > 0 Then
                '            intPtrEdit = vhija1
                '        End If
                '        Exit For
                '    End If
                'Next
                If retVal > 0 Then
                    Exit For
                End If
            End If
            'If retVal > 0 Then
            '    Exit For
            'End If
        Next
        ''
        Dim res As Integer = SendMessage(intPtrDialog, eMensajes.WM_COMMAND, eMensajes.CDN_FOLDERCHANGE, 0)
        'If pulsaEnter And intPtrEdit <> IntPtr.Zero Then
        '    'SetForegroundWindow(intPtrEdit)
        '    SendMessage(intPtrEdit, eMensajes.WA_ACTIVE, 0, 0)
        '    SendKeys.SendWait("{ENTER}")
        '    ''
        '    'SetForegroundWindow(intPtrComboBox)
        '    'SendMessage(queIntPrt, eMensajes.BM_CLICK, 0, 0)
        '    'SendMessage(queIntPrt, eMensajes.WM_ACTIVATE, eMensajes.WA_ACTIVE, 0)
        '    'SendMessage(queIntPrt, eMensajes.BM_CLICK, 0, 0)
        '    'Call PulsaEnterEnVentanaActiva(intPtrHija, True)
        '    'SendCmd("{ENTER}", intPtrHija)
        '    'PulsaEnterEnVentanaActiva(intPtrEdit, esperar)
        '    'PulsaEnterEnVentanaActiva(intPtrComboBox, True)
        'End If
        Return retVal
    End Function
    '
    Public Shared Sub Windows_Flash(hWnd As Long, totalseconds As Integer)
        For x As Integer = 1 To totalseconds
            ' Flash the window's title bar on and off once.
            Dim retval As Long = FlashWindow(hWnd, 1)
            Sleep(500)  ' pause for half a second
            retval = FlashWindow(hWnd, 0)
            Sleep(500)  ' pause for half a second
        Next
    End Sub

    Public Declare Function FlashWindow Lib "user32.dll" (ByVal hWnd As Long, ByVal bInvert As Long) As Long
    ' 64-bit version
    'Public Shared Declare Sub PtrSafe Sleep Lib "kernel32" (ByVal dwMilliseconds As LongLong)
    ' 32-bit version
    Public Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)

    Public Shared Sub Retardo_Sleep(ByVal segundos As Integer)
        Sleep(segundos * 1000)
    End Sub
    Private Sub GetComboBoxStatusBar()
        Dim maxX As Integer = -1
        Dim revits As Process() = Process.GetProcessesByName("Revit")
        Dim cb As IntPtr = IntPtr.Zero

        If revits.Length > 0 Then
            Dim children As List(Of IntPtr) = GetChildWindowsList(revits(0).MainWindowHandle)

            For Each child As IntPtr In children
                Dim classNameBuffer As StringBuilder = New StringBuilder(100)
                Dim className As Integer = GetClassName(child, classNameBuffer, 100)
                '' Class Statusbar en Revit = "msctls_statusbar32"
                If classNameBuffer.ToString().Contains("msctls_statusbar32") Then
                    Dim grandChildren As List(Of IntPtr) = GetChildWindowsList(child)

                    For Each grandChild As IntPtr In grandChildren
                        Dim classNameBuffer2 As StringBuilder = New StringBuilder(100)
                        Dim className2 As Integer = GetClassName(grandChild, classNameBuffer2, 100)

                        If classNameBuffer2.ToString().Contains("ComboBox") Then
                            Dim r As RECT
                            GetWindowRect(grandChild, r)

                            If r.Left > maxX Then
                                maxX = CInt(r.Left)
                                cb = grandChild
                            End If
                        End If
                    Next
                End If
            Next
        End If

        If cb <> IntPtr.Zero Then
            comboBoxElement = System.Windows.Automation.AutomationElement.FromHandle(cb)
        End If
    End Sub
    ' "Edit" para que nos de la casilla donde se escribe el nombre.
    Public Shared Function Dame_Revit_GuardarComoIntPrt(Optional queClass As String = "") As IntPtr
        Dim resultado As IntPtr = IntPtr.Zero
        Dim revits As Process() = Process.GetProcessesByName("Revit")
        '
        If revits.Length > 0 Then
            Dim children As List(Of IntPtr) = GetChildWindowsList(revits(0).MainWindowHandle)
            For Each child As IntPtr In children
                Dim classNameBuffer As StringBuilder = New StringBuilder(100)
                Dim className As Integer = GetClassName(child, classNameBuffer, 100)
                '' Class Statusbar en Revit = "msctls_statusbar32"
                '' 32770 = Guardar Como
                If classNameBuffer.ToString().Contains("32770") Then
                    If queClass = "" Then
                        resultado = child
                        Exit For
                    Else
                        Dim children1 As List(Of IntPtr) = GetChildWindowsList(child)
                        For Each child1 As IntPtr In children1
                            Dim classNameBuffer1 As StringBuilder = New StringBuilder(100)
                            Dim className1 As Integer = GetClassName(child1, classNameBuffer1, 100)
                            If classNameBuffer1.ToString().Contains(queClass) Then
                                resultado = child1
                                Exit For
                            Else
                                Dim children2 As List(Of IntPtr) = GetChildWindowsList(child1)
                                For Each child2 As IntPtr In children2
                                    Dim classNameBuffer2 As StringBuilder = New StringBuilder(100)
                                    Dim className2 As Integer = GetClassName(child2, classNameBuffer2, 100)
                                    If classNameBuffer2.ToString().Contains(queClass) Then
                                        resultado = child2
                                        Exit For
                                    Else

                                    End If
                                Next
                            End If
                        Next

                    End If
                End If
            Next
        End If
        '
        Return resultado
    End Function
    Public Shared Function Dame_Revit_GuardarComo() As String
        Dim resultado As String = ""
        Dim maxX As Long = -1
        Dim revits As Process() = Process.GetProcessesByName("Revit")
        Dim cb As IntPtr = IntPtr.Zero

        If revits.Length > 0 Then
            Dim children As List(Of IntPtr) = GetChildWindowsList(revits(0).MainWindowHandle)

            For Each child As IntPtr In children
                Dim classNameBuffer As StringBuilder = New StringBuilder(100)
                Dim className As Integer = GetClassName(child, classNameBuffer, 100)
                '' Class Statusbar en Revit = "msctls_statusbar32"
                'Debug.Print(classNameBuffer.ToString & vbTab & className.ToString)
                resultado &= classNameBuffer.ToString & " (" & className.ToString & ")" & vbCrLf
                If classNameBuffer.ToString().Contains("32770") Then
                    Dim grandChildren As List(Of IntPtr) = GetChildWindowsList(child)

                    For Each grandChild As IntPtr In grandChildren
                        Dim classNameBuffer2 As StringBuilder = New StringBuilder(100)
                        Dim className2 As Integer = GetClassName(grandChild, classNameBuffer2, 100)
                        'Debug.Print(vbTab & classNameBuffer2.ToString & vbTab & className2.ToString)
                        resultado &= vbTab & classNameBuffer2.ToString & " (" & className2.ToString & ")" & vbCrLf
                        If classNameBuffer2.ToString().Contains("ComboBox") Then
                            resultado = className2 & vbCrLf & resultado
                            Dim r As RECT
                            GetWindowRect(grandChild, r)

                            If r.Left > maxX Then
                                maxX = r.Left
                                cb = grandChild
                            End If
                        End If
                    Next
                End If
            Next
        End If

        If cb <> IntPtr.Zero Then
            comboBoxElement = System.Windows.Automation.AutomationElement.FromHandle(cb)
        End If
        Return resultado
    End Function
    ' 32770 es la ventana "Guardar como..."
    Public Shared Function Dame_Revit_Ventanas(queClass As String, Optional subPro As Boolean = False) As String
        Dim resultado As String = ""
        If queClass = "" Then queClass = "Revit"
        Dim procesos As Process() = Process.GetProcessesByName(queClass)
        '
        If procesos.Length > 0 Then
            If subPro = False Then
                Dim children As List(Of IntPtr) = GetChildWindowsList(procesos(0).MainWindowHandle)
                For Each child As IntPtr In children
                    Dim classNameBuffer As StringBuilder = New StringBuilder(100)
                    Dim className As Integer = GetClassName(child, classNameBuffer, 100)
                    '' Class Statusbar en Revit = "msctls_statusbar32"
                    'Debug.Print(classNameBuffer.ToString & vbTab & className.ToString)
                    'If queClass <> "" AndAlso classNameBuffer.ToString.Contains(queClass) = False Then
                    '    Continue For
                    'ElseIf queClass <> "" AndAlso classNameBuffer.ToString.Contains(queClass) = True Then
                    '    resultado &= classNameBuffer.ToString & " (" & className.ToString & ") " & GetText(child) & vbCrLf
                    '    GetVentanasRevitRecursivo(child, resultado, 1)
                    'Else
                    resultado &= classNameBuffer.ToString & " (" & className.ToString & ")" & GetText(child) & vbCrLf
                    GetVentanasRevitRecursivo(child, resultado, 1)
                    'End If
                Next
            Else
                For Each oPro As ProcessThread In procesos(0).Threads
                    Dim oProc As Process = Nothing
                    Try
                        oProc = Process.GetProcessById(oPro.Id)
                    Catch ex As Exception
                        Continue For
                    End Try
                    '
                    Dim children As List(Of IntPtr) = GetChildWindowsList(oProc.MainWindowHandle)
                    For Each child As IntPtr In children
                        Dim classNameBuffer As StringBuilder = New StringBuilder(100)
                        Dim className As Integer = GetClassName(child, classNameBuffer, 100)
                        '' Class Statusbar en Revit = "msctls_statusbar32"
                        'Debug.Print(classNameBuffer.ToString & vbTab & className.ToString)
                        'If queClass <> "" AndAlso classNameBuffer.ToString.Contains(queClass) = False Then
                        '    Continue For
                        'ElseIf queClass <> "" AndAlso classNameBuffer.ToString.Contains(queClass) = True Then
                        '    resultado &= classNameBuffer.ToString & " (" & className.ToString & ") " & GetText(child) & vbCrLf
                        '    GetVentanasRevitRecursivo(child, resultado, 1)
                        'Else
                        resultado &= classNameBuffer.ToString & " (" & className.ToString & ")" & GetText(child) & vbCrLf
                        GetVentanasRevitRecursivo(child, resultado, 1)
                        ' End If
                    Next
                Next
            End If
        End If
        Return resultado
    End Function

    Public Shared Sub GetVentanasRevitRecursivo(queIntPtr As IntPtr, ByRef mensaje As String, nivel As Integer)
        Dim grandChildren As List(Of IntPtr) = GetChildWindowsList(queIntPtr)
        If grandChildren IsNot Nothing AndAlso grandChildren.Count > 0 Then
            For Each grandChild As IntPtr In grandChildren
                Dim classNameBuffer As StringBuilder = New StringBuilder(100)
                Dim className As Integer = GetClassName(grandChild, classNameBuffer, 100)
                'Debug.Print(vbTab & classNameBuffer2.ToString & vbTab & className2.ToString)
                mensaje &= StrDup(nivel, vbTab) & classNameBuffer.ToString & " (" & className.ToString & ")" & vbCrLf
                GetVentanasRevitRecursivo(grandChild, mensaje, nivel + 1)
            Next
        End If
    End Sub
    '
    '
    'Dim cbi As COMBOBOXINFO
    'cbi.cbSize = System.Runtime.InteropServices.Marshal.SizeOf(cbi)
    'If GetComboBoxInfo(comboBox1.Handle, cbi) Then
    '...
    'End If
    <DllImport("user32.dll")>
    Public Shared Function GetComboBoxInfo(ByVal hWnd As IntPtr, ByRef pcbi As COMBOBOXINFO) As Boolean
    End Function

    <StructLayout(LayoutKind.Sequential)>
    Public Structure COMBOBOXINFO
        Public cbSize As Int32
        Public rcItem As RECT
        Public rcButton As RECT
        Public buttonState As ComboBoxButtonState
        Public hwndCombo As IntPtr
        Public hwndEdit As IntPtr
        Public hwndList As IntPtr
    End Structure

    Public Enum ComboBoxButtonState
        STATE_SYSTEM_NONE = 0
        STATE_SYSTEM_INVISIBLE = &H8000
        STATE_SYSTEM_PRESSED = &H8
    End Enum
    'Private Sub buttonGet_Click(ByVal sender As Object, ByVal e As EventArgs)
    '    If (comboBoxElement IsNot Nothing) AndAlso (comboBoxElement.Current.IsEnabled) Then
    '        Dim expandPattern As ExpandCollapsePattern = CType(comboBoxElement.GetCurrentPattern(ExpandCollapsePattern.Pattern), ExpandCollapsePattern)
    '        expandPattern.Expand()
    '        listBox1.Items.Clear()
    '        Dim cacheRequest As CacheRequest = New CacheRequest()
    '        cacheRequest.Add(AutomationElement.NameProperty)
    '        cacheRequest.TreeScope = TreeScope.Element Or TreeScope.Children
    '        Dim comboboxItems As AutomationElement = comboBoxElement.GetUpdatedCache(cacheRequest)

    '        For Each item As AutomationElement In comboboxItems.CachedChildren

    '            If item.Current.Name = "" Then
    '                Dim cacheRequest2 As CacheRequest = New CacheRequest()
    '                cacheRequest2.Add(AutomationElement.NameProperty)
    '                cacheRequest2.TreeScope = TreeScope.Element Or TreeScope.Children
    '                Dim comboboxItems2 As AutomationElement = item.GetUpdatedCache(cacheRequest)

    '                For Each item2 As AutomationElement In comboboxItems2.CachedChildren
    '                    listBox1.Items.Add(item2.Current.Name)
    '                Next
    '            End If
    '        Next

    '        expandPattern.Collapse()
    '    End If

    '    GetSelection()
    'End Sub

    'Private Sub GetSelection()
    '    If (comboBoxElement IsNot Nothing) AndAlso (comboBoxElement.Current.IsEnabled) Then
    '        Dim selPattern As SelectionPattern = TryCast(comboBoxElement.GetCurrentPattern(SelectionPattern.Pattern), SelectionPattern)
    '        Dim items As AutomationElement() = selPattern.Current.GetSelection()

    '        For Each item As AutomationElement In items
    '            Dim index As Integer = 0

    '            For Each listItem In listBox1.Items

    '                If CStr(listItem) = item.Current.Name Then
    '                    listBox1.SelectedIndex = index
    '                    Exit For
    '                End If

    '                index += 1
    '            Next
    '        Next
    '    End If

    '    SetForegroundWindow(Me.Handle)
    'End Sub

    'Private Sub SetSelection(ByVal [option] As String)
    '    If (comboBoxElement IsNot Nothing) AndAlso (comboBoxElement.Current.IsEnabled) Then
    '        Dim sel As AutomationElement = Nothing
    '        Dim expandPattern As ExpandCollapsePattern = CType(comboBoxElement.GetCurrentPattern(ExpandCollapsePattern.Pattern), ExpandCollapsePattern)
    '        expandPattern.Expand()
    '        Dim cacheRequest As CacheRequest = New CacheRequest()
    '        cacheRequest.Add(AutomationElement.NameProperty)
    '        cacheRequest.TreeScope = TreeScope.Element Or TreeScope.Children
    '        Dim comboboxItems As AutomationElement = comboBoxElement.GetUpdatedCache(cacheRequest)

    '        For Each item As AutomationElement In comboboxItems.CachedChildren

    '            If item.Current.Name = "" Then
    '                Dim cacheRequest2 As CacheRequest = New CacheRequest()
    '                cacheRequest2.Add(AutomationElement.NameProperty)
    '                cacheRequest2.TreeScope = TreeScope.Element Or TreeScope.Children
    '                Dim comboboxItems2 As AutomationElement = item.GetUpdatedCache(cacheRequest)

    '                For Each item2 As AutomationElement In comboboxItems2.CachedChildren

    '                    If item2.Current.Name = [option] Then
    '                        sel = item2
    '                    End If
    '                Next
    '            End If
    '        Next

    '        If sel IsNot Nothing Then
    '            Dim [select] As SelectionItemPattern = CType(sel.GetCurrentPattern(SelectionItemPattern.Pattern), SelectionItemPattern)
    '            [select].[Select]()
    '        End If

    '        expandPattern.Collapse()
    '    End If

    '    SetForegroundWindow(Me.Handle)
    'End Sub

    'Private Sub buttonSet_Click(ByVal sender As Object, ByVal e As EventArgs)
    '    If listBox1.Items.Count > 0 AndAlso listBox1.SelectedIndex > -1 Then
    '        Dim [option] As String = CStr(listBox1.Items(listBox1.SelectedIndex))
    '        SetSelection([option])
    '    End If
    '    SetForegroundWindow(Me.Handle)
    'End Sub
    'Public Shared Async Function ExportWarinings(ByVal uiapp As UIApplication) As Task(Of String)
    '    Dim tempDir = path.Combine(path.GetTempPath(), "RevitWarnings")
    '    Dim path = path.Combine(tempDir, "warning_" & Guid.NewGuid() & ".html")
    '    If Not Directory.Exists(tempDir) Then Directory.CreateDirectory(tempDir)
    '    Dim id = RevitCommandId.LookupPostableCommandId(PostableCommand.ReviewWarnings)
    '    Dim mainWin As IntPtr = GetForegroundWindow()
    '    uiapp.PostCommand(id)

    '    While GetForegroundWindow() = mainWin
    '        Await DelayWork(100)
    '    End While

    '    Dim warningsWindow As IntPtr = GetForegroundWindow()
    '    Task.Run(Function() TypeSaveFileName(path))
    '    ClickControl(warningsWindow, "&Export...")
    '    ClickControl(warningsWindow, "Close")
    '    Return path
    'End Function
    '
    'Task.Run(Function() SaveAs_PonTexto(path))
    Public Shared Async Sub SaveAs_PonTexto(ByVal text As String)
        Dim mainWin As IntPtr = GetForegroundWindow()

        While GetForegroundWindow() = mainWin
            Await DelayWork(100)
        End While

        Dim saveWindows As IntPtr = GetForegroundWindow()
        Dim comboHandle As IntPtr = GetDlgItem(saveWindows, 13006)
        Dim editHandle As IntPtr = GetDlgItem(comboHandle, 1001)
        SetControlText(comboHandle, 1001, text)
        SetFocus(editHandle)
        PostMessage(editHandle, eMensajes.WM_KEYDOWN, eMensajes.VK_SPACE, 0)
        Await DelayWork(100)
        ClickControl(saveWindows, "&Save")
    End Sub
    Public Shared Sub ClickControl(ByVal windowHandle As IntPtr, ByVal buttonText As String)
        Dim export As IntPtr = FindWindowEx(windowHandle, IntPtr.Zero, Nothing, buttonText)
        SendMessage(export, eMensajes.BM_CLICK, IntPtr.Zero.ToInt32, IntPtr.Zero.ToInt32)
    End Sub
    Public Shared Sub ClickControl(ByVal windowHandle As IntPtr)
        SendMessage(windowHandle, eMensajes.BM_CLICK, IntPtr.Zero.ToInt32, IntPtr.Zero.ToInt32)
    End Sub

    Public Shared Sub SetControlText(ByVal windowHandle As IntPtr, ByVal controlId As Integer, ByVal text As String)
        Dim iptrHWndControl As IntPtr = GetDlgItem(windowHandle, controlId)
        Dim hrefHWndTarget As HandleRef = New HandleRef(Nothing, iptrHWndControl)
        SendMessage(hrefHWndTarget, eMensajes.WM_SETTEXT, IntPtr.Zero.ToInt32, text)
    End Sub

    Public Shared Async Function DelayWork(ByVal i As Integer) As System.Threading.Tasks.Task
        Await System.Threading.Tasks.Task.Delay(i)
    End Function

#Region "LETREROS DE DIALOGO"
    'Estructuras para coordenadas y dimensiones de las ventanas y objetos, en este caso el cd
    '    Public Enum POINTAPI
    '        X As Long
    '  Y As Long
    'End Enum

    '    Private Type RECT
    '  left As Long
    '  top As Long
    '  Right As Long
    '  Bottom As Long
    'End Type
    ''Estructura para las opciones del cuadro de dialogo
    'Type OPENFILENAME
    '    lStructSize As Long
    '    hwndOwner As Long
    '    hInstance As Long
    '    lpstrFilter As String
    '    lpstrCustomFilter As String
    '    nMaxCustFilter As Long
    '    nFilterIndex As Long
    '    lpstrFile As String
    '    nMaxFile As Long
    '    lpstrFileTitle As String
    '    nMaxFileTitle As Long
    '    lpstrInitialDir As String
    '    lpstrTitle As String
    '    Flags As Long
    '    nFileOffset As Integer
    '    nFileExtension As Integer
    '    lpstrDefExt As String
    '    lCustData As Long
    '    lpfnHook As Long
    '    lpTemplateName As String
    'End Type
    '' Estructura usada en la Notificaciones de mensajes
    'Type OPENFILENAME2
    '        lStructSize As Long
    '        hwndOwner As Long
    '        hInstance As Long
    '        lpstrFilter As Long
    '        lpstrCustomFilter As Long
    '        nMaxCustFilter As Long
    '        nFilterIndex As Long
    '        lpstrFile As Long
    '        nMaxFile As Long
    '        lpstrFileTitle As Long
    '        nMaxFileTitle As Long
    '        lpstrInitialDir As Long
    '        lpstrTitle As Long
    '        Flags As Long
    '        nFileOffset As Integer
    '        nFileExtension As Integer
    '        lpstrDefExt As Long
    '        lCustData As Long
    '        lpfnHook As Long
    '        lpTemplateName As Long
    'End Type
    'Type NMHDR
    '    hwndFrom As Long
    '    idFrom As Long
    '    code As Long
    'End Type
    ''estructura para la rutina de Notificacion de mensajes del cuadro de dialogo
    'Type OFNOTIFY
    '        hdr As NMHDR
    '        lpOFN As OPENFILENAME2
    '        pszFile As Long 'String '  May be NULL
    'End Type
#End Region

#Region "ENUMERACIONES"
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
    Public Shared Function ShowWindow(
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
        ' ## Para leer y escribir texto en controles
        WM_SETTEXT = &HC        '0x000C
        WM_GETTEXT = &HD
        WM_GETTEXTLENGTH = &HE
        '
        ' ## Comandos a enviar
        ' Ejemplo: Cerrar ventana
        'SendMessage(iHandle, WM_SYSCOMMAND, SC_CLOSE, 0)
        cero = 0
        WM_ACTIVATE = &H6
        WM_SETHOTKEY = &H32
        WM_SHOWWINDOW = &H18
        WM_COMMAND = &H111
        WM_SYSCOMMAND = &H112   '0x0112
        SC_CLOSE = &HF060&      '0xF060
        ' ## Ejemplo: Activar y hacer clic 
        ' SendMessage(hWnd, WM_ACTIVATE, WA_ACTIVE, 0)
        ' SendMessage(hWnd, BM_CLICK, 0, 0) 
        WA_ACTIVE = 1
        WM_NULL = &H0
        WM_CREATE = &H1
        WM_CHAR = &H102
        WM_DESTROY = &H2
        WM_CLOSE = &H10    '16 EN DECIMAL
        WM_MOVE = &H3
        WM_SIZE = &H5
        WM_COPYDATA = &H4A
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
        'Constantes de SendMessage:

        GW_HWNDFIRST = 0&
        GW_HWNDNEXT = 2&
        GW_CHILD = 5&
        '
        '' ***** TECLADO
        VK_LBUTTON = &H1
        VK_RBUTTON = &H2
        VK_CTRLBREAK = &H3
        VK_MBUTTON = &H4
        VK_BACKSPACE = &H8
        VK_TAB = &H9
        VK_ENTER = &HD
        VK_SHIFT = &H10
        VK_CONTROL = &H11
        VK_ALT = &H12
        VK_PAUSE = &H13
        VK_CAPSLOCK = &H14
        VK_ESCAPE = &H1B    '0x1B
        VK_SPACE = &H20
        VK_PAGEUP = &H21
        VK_PAGEDOWN = &H22
        VK_END = &H23
        VK_HOME = &H24
        VK_LEFT = &H25
        VK_UP = &H26
        VK_RIGHT = &H27
        VK_DOWN = &H28
        VK_PRINTSCREEN = &H2C
        VK_INSERT = &H2D
        VK_DELETE = &H2E
        VK_0 = &H30
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
        VK_n = &H4E
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
        VK_LWINDOWS = &H5B
        VK_RWINDOWS = &H5C
        VK_APPSPOPUP = &H5D
        VK_NUMPAD_0 = &H60
        VK_NUMPAD_1 = &H61
        VK_NUMPAD_2 = &H62
        VK_NUMPAD_3 = &H63
        VK_NUMPAD_4 = &H64
        VK_NUMPAD_5 = &H65
        VK_NUMPAD_6 = &H66
        VK_NUMPAD_7 = &H67
        VK_NUMPAD_8 = &H68
        VK_NUMPAD_9 = &H69
        VK_NUMPAD_MULTIPLY = &H6A
        VK_NUMPAD_ADD = &H6B
        VK_NUMPAD_PLUS = &H6B
        VK_NUMPAD_SUBTRACT = &H6D
        VK_NUMPAD_MINUS = &H6D
        VK_NUMPAD_MOINS = &H6D
        VK_NUMPAD_DECIMAL = &H6E
        '
        '' ***** Constantes de notificaciones de mensajes del CommonDialog
        SWP_SHOWWINDOW = &H40
        CDM_FIRST = (WM_USER + 100)
        CDM_GETFILEPATH = (CDM_FIRST + &H1)

        CDN_FIRST = (-601)
        CDN_LAST = (-699)
        CDN_INITDONE = (CDN_FIRST - &H0)
        CDN_SELCHANGE = (CDN_FIRST - &H1)
        CDN_FOLDERCHANGE = (CDN_FIRST - &H2)
        CDN_SHAREVIOLATION = (CDN_FIRST - &H3)
        CDN_HELP = (CDN_FIRST - &H4)
        CDN_FILEOK = (CDN_FIRST - &H5)
        CDN_TYPECHANGE = (CDN_FIRST - &H6)
        WM_NOTIFY = &H4E

        OFN_ENABLEHOOK = &H20
        OFN_HIDEREADONLY = &H4
        OFN_PATHMUSTEXIST = &H800
        OFN_FILEMUSTEXIST = &H1000
        OFN_EXPLORER = &H80000
        OFN_LONGNAMES = &H200000
        lst1 = &H460
    End Enum
#End Region
End Class
