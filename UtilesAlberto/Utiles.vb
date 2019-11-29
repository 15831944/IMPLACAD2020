Imports System.Xml
Imports System.Linq
Imports System.Data
Imports System.Reflection
Imports System.Management

Partial Public Class Utiles
    Public WithEvents oTt As Windows.Forms.ToolTip

    Public Enum EAddresses2aCAD
        Ninguna
        ACoruña
        Bilbao
        Barcelona
        Gipuzkoa
        Madrid
        Portugal
        Portugal1
        Sevilla
        Valencia
        Zaragoza
    End Enum
    '' Direcciones de 2aCAD Delegaciones (8 líneas máximo)
    Public Shared Function Direcciones2aCAD(ByVal queDire As EAddresses2aCAD) As String
        Dim resultado As String = ""
        Select Case queDire
            Case EAddresses2aCAD.ACoruña
                resultado &= "A Coruña"
                resultado &= "C/ Rosaía de Castro, 4" & vbCrLf
                resultado &= "15004 - A Coruña" & vbCrLf & vbCrLf
                resultado &= "Teléfono: 902 570 325" & vbCrLf
                resultado &= "Email: galicia@2acad.com" & vbCrLf
            Case EAddresses2aCAD.Bilbao
                resultado &= "Bilbao"
                resultado &= "C/ Gran Vía Don Diego López de Haro, 19-21, 2ª Planta" & vbCrLf
                resultado &= "48009 - Bilbao (Bizkaia)" & vbCrLf & vbCrLf
                resultado &= "Teléfono: 902 570 325" & vbCrLf
                resultado &= "Email: bizkaia@2acad.com" & vbCrLf
            Case EAddresses2aCAD.Barcelona
                resultado &= "Barcelona" & vbCrLf
                resultado &= "C/ Bac de Roda, 63" & vbCrLf
                resultado &= "08005 - Barcelona" & vbCrLf & vbCrLf
                resultado &= "Teléfono: 902 570 325" & vbCrLf
                resultado &= "Email: barcelona@2acad.com" & vbCrLf
            Case EAddresses2aCAD.Gipuzkoa
                resultado &= "Gipuzkoa" & vbCrLf
                resultado &= "Polo de Innovación Garaia, Goiru 1 - Edf. A" & vbCrLf
                resultado &= "20500 - Arrasate-Mondragón (Gipuzkoa)" & vbCrLf & vbCrLf
                resultado &= "Teléfono: 902 570 325" & vbCrLf
                resultado &= "Email: gipuzkoa@2acad.com" & vbCrLf
            Case EAddresses2aCAD.Madrid
                resultado &= "Madrid" & vbCrLf
                resultado &= "Paseo de la Castellana, 135, Planta 7" & vbCrLf
                resultado &= "28046 - Madrid" & vbCrLf & vbCrLf
                resultado &= "Teléfono: 902 570 325" & vbCrLf
                resultado &= "Email: madrid@2acad.com" & vbCrLf
            Case EAddresses2aCAD.Portugal
                resultado &= "Portugal" & vbCrLf
                resultado &= "Av. D. Joâo II, 50" & vbCrLf
                resultado &= "1990-095 - Lisboa (Portugal)" & vbCrLf & vbCrLf
                resultado &= "Teléfonos: +35 121 414 7303 / +35 193 319 1920 /  +35 193 319 1929" & vbCrLf
                resultado &= "Email: lisboa@2acad.com" & vbCrLf
            Case EAddresses2aCAD.Portugal1
                resultado &= "Portugal" & vbCrLf
                resultado &= "Largo da Lagoa 7C. Escritorio 108" & vbCrLf
                resultado &= "2795-116 - Linda-a-Velha (Portugal)" & vbCrLf & vbCrLf
                resultado &= "Teléfonos: +35 121 414 7303 / +35 193 319 1920 /  +35 193 319 1929" & vbCrLf
                resultado &= "Email: lisboa@2acad.com" & vbCrLf
            Case EAddresses2aCAD.Sevilla
                resultado &= "Sevilla" & vbCrLf
                resultado &= "Centro de Empresas Aerópolis" & vbCrLf
                resultado &= "C/ Ingeniero Rafael Rubio Elola, 1. Módulo 2.5 Oeste" & vbCrLf
                resultado &= "41300 - La Rinconada (Sevilla)" & vbCrLf & vbCrLf
                resultado &= "Teléfono: 902 570 325" & vbCrLf
                resultado &= "Email: sevilla@2acad.com" & vbCrLf
            Case EAddresses2aCAD.Valencia
                resultado &= "Valencia" & vbCrLf
                resultado &= "Ronda Narciso Monturiol, 6" & vbCrLf
                resultado &= "Parque Tecnológico de Paterna" & vbCrLf
                resultado &= "46980 - Paterna (Valencia)" & vbCrLf & vbCrLf
                resultado &= "Teléfono: 963 134 035" & vbCrLf
                resultado &= "Email: valencia@2acad.com" & vbCrLf
            Case EAddresses2aCAD.Zaragoza
                resultado &= "Zaragoza" & vbCrLf
                resultado &= "C/Bari, 57" & vbCrLf
                resultado &= "Centro Tecnológico TICXXI PLA-ZA" & vbCrLf
                resultado &= "50197 - Zaragoza" & vbCrLf & vbCrLf
                resultado &= "Teléfono: 976 45 81 45/51 (CAD directo)" & vbCrLf
                resultado &= "Fax: 976 75 20 15" & vbCrLf
                resultado &= "Email: zaragoza@2acad.com" & vbCrLf
        End Select
        Return resultado
    End Function
    '
    Public Shared Function GUID_OnlyNumber(Optional nChar As Integer = 32) As String
        Dim temp As String = System.Guid.NewGuid.ToString.Replace("-", "")
        Return IIf(temp.Length >= nChar, temp.Substring(0, nChar), temp)
    End Function
    Public Shared Function GUID_Full() As String
        Return System.Guid.NewGuid.ToString
    End Function


    Public Shared Function Random_Get(min As Integer, max As Integer) As Integer
        Randomize()
        Return CInt(Math.Floor((max - min + 1) * Rnd())) + min
        'randomValue = CInt(Math.Floor((upperbound - lowerbound + 1) * Rnd())) + lowerbound
    End Function

    Public Shared Function Number_GetListRandom(min As Integer, max As Integer, total As Integer, Optional sort As Boolean = True) As List(Of Integer)
        Dim lNumber As New List(Of Integer)
        For x = 1 To total
            Dim numero As Integer = Random_Get(min, max)
            If lNumber.Contains(numero) Then
                x -= 1
                Continue For
            Else
                lNumber.Add(numero)
            End If
        Next
        If sort = True Then lNumber.Sort()
        Return lNumber
    End Function
    ''
    'Public Shared Function Eval(ByVal Expresion As String) As Object
    '    Dim vbcp As New Microsoft.VisualBasic.VBCodeProvider
    '    Dim vbc As System.CodeDom.Compiler.ICodeCompiler = vbcp.CreateCompiler
    '    Dim cpar As New System.CodeDom.Compiler.CompilerParameters
    '    Dim res As System.CodeDom.Compiler.CompilerResults

    '    cpar.GenerateExecutable = False ' Generar DLL
    '    cpar.GenerateInMemory = True ' Generar en memoria
    '    cpar.IncludeDebugInformation = True

    '    ' Agregar referencias
    '    cpar.ReferencedAssemblies.Add("Microsoft.VisualBasic.dll")

    '    ' Compilar
    '    res = vbc.CompileAssemblyFromSource(cpar,
    '        "Imports Microsoft.VisualBasic" & vbCrLf &
    '        "Namespace MiNamespace" & vbCrLf &
    '        " Public Class MiClase" & vbCrLf &
    '        "  Public Shared Function Eval() As Object " & vbCrLf &
    '        "   Return " & Expresion & vbCrLf &
    '        "  End Function" & vbCrLf &
    '        " End Class" & vbCrLf &
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
    Declare Function WNetGetConnection Lib "mpr.dll" Alias "WNetGetConnectionA" (ByVal lpszLocalName As String,
           ByVal lpszRemoteName As String, ByRef cbRemoteName As Integer) As Integer

    ''' <summary>
    ''' Metodo para comprobar si una impresora esta online o offline
    ''' Imports System.Drawing.Printing (Agregar System.Drawing a Referencias)
    ''' Imports System.Management (Agregar System.Management a Referencias)
    ''' </summary>
    ''' <param name="printerName">nombre de la impresora</param>
    ''' <returns></returns>
    Public Shared Function IsPrinterOnline(printerName As String) As Boolean
        Dim str As String = ""
        Dim online As Boolean = False

        'set the scope of this search to the local machine
        Dim scope As New ManagementScope(ManagementPath.DefaultPath)
        'connect to the machine
        scope.Connect()

        'query for the ManagementObjectSearcher
        Dim query As New SelectQuery("select * from Win32_Printer")

        Dim m As New ManagementClass("Win32_Printer")

        Dim obj As New ManagementObjectSearcher(scope, query)

        'get each instance from the ManagementObjectSearcher object
        Using printers As ManagementObjectCollection = m.GetInstances()
            'now loop through each printer instance returned
            For Each printer As ManagementObject In printers
                'first make sure we got something back
                If printer IsNot Nothing Then
                    'get the current printer name in the loop
                    str = printer("Name").ToString().ToLower()

                    'check if it matches the name provided
                    If str.Equals(printerName.ToLower()) Then
                        'since we found a match check it's status
                        If printer("WorkOffline").ToString().ToLower().Equals("true") OrElse printer("PrinterStatus").Equals(7) Then
                            'it's offline
                            online = False
                        Else
                            'it's online
                            online = True
                        End If
                    End If
                Else
                    Throw New Exception("No printers were found")
                End If
            Next
        End Using
        Return online
    End Function
End Class
