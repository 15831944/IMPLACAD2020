'------------------------------------------------------------------------------
' formulario de prueba para enumerar ventanas                       (10/Ene/05)
'
' ©Guillermo 'guille' Som, 2005
'------------------------------------------------------------------------------
Imports System
Imports System.Windows.Forms

Public Class clsVentanas
    Public Shared lc As IntPtr      'Aquí tendremos el identificados Ventana de "MountTam" para enviar comandos Linea de comandos.
    '
    Private Delegate Function EnumWindowsDelegate _
            (ByVal hWnd As System.IntPtr, ByVal parametro As Integer) As Boolean
    '
    <System.Runtime.InteropServices.DllImport("user32.dll")>
    Private Shared Function EnumWindows(
            ByVal lpfn As EnumWindowsDelegate,
            ByVal lParam As Integer) As Boolean
    End Function
    '
    <System.Runtime.InteropServices.DllImport("user32.dll")> _
    Private Shared Function EnumChildWindows _
            (ByVal hWndParent As System.IntPtr, _
            ByVal lpEnumFunc As EnumWindowsDelegate, _
            ByVal lParam As Integer) As Integer
    End Function
    '
    '
    <System.Runtime.InteropServices.DllImport("user32.dll")> _
    Private Shared Function GetWindowText( _
            ByVal hWnd As System.IntPtr, _
            ByVal lpString As System.Text.StringBuilder, _
            ByVal cch As Integer) As Integer
    End Function
    '
    ' Para EnumWindows
    Private Shared colWin As New System.Collections.Specialized.StringDictionary
    '
    Private Shared Function EnumWindowsProc(ByVal hWnd As System.IntPtr, ByVal parametro As Integer) As Boolean
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
            If nombreVentana = "MountTam" Then lc = hWnd
        End If
        '
        Return True
    End Function
    '
    Public Shared Sub EnumerarVentanas()
        ' Enumera las ventanas principales (TopWindows)
        colWin.Clear()
        EnumWindows(AddressOf EnumWindowsProc, 0)
        '
        Dim ventanas As String = ""
        For Each s As String In colWin.Keys
            'Dim lvi As ListViewItem = lvTop.Items.Add(s)
            'lvi.SubItems.Add(colWin(s))
            ventanas &= s & vbCrLf
        Next
        'If lvTop.Items.Count > 0 Then
        'btnChild.Enabled = True
        'End If
        MessageBox.Show(ventanas)
    End Sub
    '
    Public Shared Sub enumerarVentanasHijas(ByVal handleParent As System.IntPtr)
        ' Enumera las ventanas hijas del handle indicado
        colWin.Clear()
        EnumChildWindows(handleParent, AddressOf EnumWindowsProc, 0)
        '
        'lvChild.Items.Clear()
        Dim ventanas As String = ""
        For Each s As String In colWin.Keys
            'Dim lvi As ListViewItem = lvChild.Items.Add(s)
            'lvi.SubItems.Add(colWin(s))
            ventanas &= s & vbCrLf
        Next
        MessageBox.Show(ventanas)
    End Sub

    Public Shared Function DameVentanasHijas(ByVal handleparent As System.IntPtr) As System.Collections.Specialized.StringDictionary
        colWin.Clear()
        EnumChildWindows(handleparent, AddressOf EnumWindowsProc, 0)
        '
        'lvChild.Items.Clear()
        'Dim arrayl As New ArrayList
        'For Each s As String In colWin.Keys
        'Dim lvi As ListViewItem = lvChild.Items.Add(s)
        'lvi.SubItems.Add(colWin(s))
        'arrayl.Add(s)
        'Next
        DameVentanasHijas = colWin
	End Function

	Public Shared Function DameNombresVentanasHijas(ByVal handleparent As System.IntPtr) As String
		Dim mensaje As String = ""
		colWin.Clear()
		EnumChildWindows(handleparent, AddressOf EnumWindowsProc, 0)
		'
		For Each s As String In colWin.Keys
			mensaje &= "Id : " & s & vbTab & "Nombre : " & colWin(s) & vbCrLf
		Next
		DameNombresVentanasHijas = mensaje
	End Function
End Class
