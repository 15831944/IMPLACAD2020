Imports System.Diagnostics
Imports System.Collections
Imports System.IO

Imports Autodesk.AutoCAD.Interop
Imports Autodesk.AutoCAD.Interop.Common
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.EditorInput

Namespace A2acad
    Partial Public Class A2acad
        'oAppA.GetAcadState.IsQuiescent (si devuelve TRUE está libre) (si devuelve FALSE esta ocupado y
        'no podemos hacer nada con él. Darán error las llamadas a documentos, etc.
        Public oAppA As Autodesk.AutoCAD.Interop.AcadApplication = Nothing
        Public oAppS As Autodesk.AutoCAD.ApplicationServices.Application
        'Public oAppS As Autodesk.AutoCAD.ApplicationServices.Application = Nothing
        '
        Public oSel As AcadSelectionSet = Nothing           'Seleccion actual
        'Public oSelTemp As AcadSelectionSet = Nothing       'Seleccion temporal
        Public OC As AcadAcCmColor = Nothing                'Objeto color
        Public oVp As AcadViewport = Nothing                'Viewport actual
        Public oV As AcadView = Nothing                     'Vista actual. Para restituir en algunos casos
        Public oBlult As AcadBlockReference = Nothing       'Ultimo bloque seleccionado
        Public oBlold As AcadBlockReference = Nothing       'Bloque seleccionada para sustituir (Será igual a oBlult)
        Public oBlnew As AcadBlockReference = Nothing       'Bloque nuevo que insertamos (Sutituye a oBlold)
        Public oH As AcadHatch = Nothing                            'Hatch
        '
        ' CONSTANTES
        Private Const keyPre As String = "aiiao2K19"
        Private Const keySuf As String = "2aCAD"
        Private ReadOnly activado As Boolean = False
        Public Const nSel As String = "2acad"
        Public Const estadocapas As String = "estado"
        ' En cada instancia de esta clase, le tendremos que dar regAPPA
        Public regAPPA As String = ""           'Nombre Aplicación para elementos XData [CLIENTE]2acad
        '
        ' Tamaños DIN
        Public A0 As Array = New String() {"1189", "841"}
        Public A1 As Array = New String() {"841", "594"}
        Public A2 As Array = New String() {"594", "420"}
        Public A3 As Array = New String() {"420", "297"}
        Public A4 As Array = New String() {"297", "210"}
        '
        ' Variables para nombres DLL llamantes y AutoCAD2acad.dll
        Public Const _appYoNameExt As String = "AutoCAD2acad.dll"
        Public Const _appYoName As String = "AutoCAD2acad"
        Public Const _appYoConfig As String = _appYoNameExt & ".config"
        Public _appYoFullPath As String = IO.Path.Combine(My.Application.Info.DirectoryPath, _appYoNameExt)
        Public _appYoDir As String = IO.Path.GetDirectoryName(_appYoFullPath)
        Public _appYoConfigFullPath As String = IO.Path.Combine(_appYoDir, _appYoConfig)
        '
        Public _appPath As String = ""                  ' Camino completo a la aplicación que llama a esta clase.
        Public _appDir As String = ""                   ' Directorio de la aplicación que llama a esta clase.
        Public _appNombre As String = ""                ' Nombre de la DLL que llamará a esta clase.
        Public _appLog As String = ""                   ' FulltPath al fichero log. (Tendrá el nombre _appNombre)
        '
        ' Datos equipo
        Public _usuario As String = Environment.UserName    ' My.User.Name
        Public _maquina As String = Environment.MachineName    ' My.Computer.Name
        Public _dominio As String = Environment.UserDomainName
        Public _OSFull As String = My.Computer.Info.OSFullName
        Public _OSBits As String = My.Computer.Info.OSPlatform
        Public _OSVersion As String = My.Computer.Info.OSVersion
        '
        Public PRs() As Process = Nothing                            'Todos los procesos de AutoCAD "acad.exe"
        Public ventanas As System.Collections.Specialized.StringDictionary = Nothing
        Public prApp As Process = Nothing     'Proceso AutoCAD.
        'Public oAppIP As IntPtr
        Public oBd As AcadDatabase = Nothing
        Public oBds As Autodesk.AutoCAD.DatabaseServices.Database = Nothing
        Public oLsm As AcadLayerStateManager = Nothing
        'Public oAppAT As String = ""                     ' Titulo de la Aplicacion
        Public oDocFull As String = ""                      ' Nombre completo (FullName) del último dibujo activo (sobre el que actuamos) FullName completo
        Public oAppAintP As IntPtr = IntPtr.Zero                     ' intPtr de AutoCAD
        Public oDocintP As IntPtr = IntPtr.Zero                      ' intPrt del dibujo
        Public oDochw As Integer = 0                    ' HWND del dibujo
        Public lineacomandos As IntPtr = IntPtr.Zero              ' Linea de comandos Autocad (MountTam) para enviar teclas y cancelar comandos.
        Public atributoNUMERO As String = Nothing             ' Nombre del atributo que usamos para numeración única.
        '
        'End Get
        'End Property
        Public Function oDoc() As Autodesk.AutoCAD.Interop.AcadDocument
            If oAppA IsNot Nothing AndAlso oAppA.Documents.Count > 0 Then
                Return oAppA.ActiveDocument
            Else
                Return Nothing
            End If
        End Function
        ''' <summary>
        ''' Convierte en metros o metros cuadrados un valor
        ''' en función de la configuración de AutoCAD.
        ''' Si ponemos "distancia" no evaluará superficie "DameMetros(300.456)"
        ''' Si ponemos "superfice" poner 0 en distancia "DameMetros(0, 4320.234)"
        ''' </summary>
        ''' <param name="distancia">distancia a convertir en m. No indicar nada en superficie</param>
        ''' <param name="superficie">superficie a convertir en m2. (poner 0 en distancia)</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function DameMetros(ByVal distancia As Double, Optional ByVal superficie As Double = 0) As Double
            Dim queUnid As Integer = oAppA.ActiveDocument.GetVariable("INSUNITS")
            Dim resultado As Double = 0
            If distancia > 0 Then
                Select Case queUnid
                    Case 1  'El dibujo está en pulgadas
                        resultado = (distancia * 2.54) / 100
                    Case 4  'El dibujo está en milimetros
                        resultado = distancia / 1000
                    Case 5  'El dibujo está en centímetros
                        resultado = distancia / 100
                    Case 6  'El dibujo está en metros
                        resultado = distancia   'No haces nada. Ya esta en metros.
                    Case 14 'El dibujo está en decímetros
                        resultado = distancia / 10
                End Select
            ElseIf superficie > 0 Then
                Select Case queUnid
                    Case 1  'El dibujo está en pulgadas
                        resultado = (superficie * 2.54) / 10000
                    Case 4  'El dibujo está en milimetros
                        resultado = superficie / 1000000
                    Case 5  'El dibujo está en centímetros
                        resultado = superficie / 10000
                    Case 6  'El dibujo está en metros
                        resultado = superficie   'No haces nada. Ya esta en metros.
                    Case 14 'El dibujo está en decímetros
                        resultado = superficie / 100
                End Select
            Else
                resultado = 0
            End If
            DameMetros = FormatNumber(resultado, 2, TriState.True)
        End Function


        ''' <summary>
        ''' Convierte en unidades AutoCAD los metros o metros cuadrados
        ''' que le enviemos.
        ''' Si ponemos "distancia" no evaluará superficie "DameMetros(2.75)"
        ''' Si ponemos "superfice" poner 0 en distancia "DameMetros(0, 123.23)"
        ''' </summary>
        ''' <param name="distancia">distancia a convertir de m. a unidades Acad. No indicar nada en superficie</param>
        ''' <param name="superficie">superficie a convertir de m2. a unidades Acad. (poner 0 en distancia)</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function DameUnidAcad(ByVal distancia As Double, Optional ByVal superficie As Double = 0) As Double
            Dim queUnid As Integer = oAppA.ActiveDocument.GetVariable("INSUNITS")
            Dim resultado As Double = 0
            If distancia > 0 Then
                Select Case queUnid
                    Case 1  'El dibujo está en pulgadas
                        resultado = (distancia * 100) / 2.54
                    Case 4  'El dibujo está en milimetros
                        resultado = distancia * 1000
                    Case 5  'El dibujo está en centímetros
                        resultado = distancia * 100
                    Case 6  'El dibujo está en metros
                        resultado = distancia   'No haces nada. Ya esta en metros.
                    Case 14 'El dibujo está en decímetros
                        resultado = distancia * 10
                End Select
            ElseIf superficie > 0 Then
                Select Case queUnid
                    Case 1  'El dibujo está en pulgadas
                        resultado = (superficie * 10000) / 2.54
                    Case 4  'El dibujo está en milimetros
                        resultado = superficie * 1000000
                    Case 5  'El dibujo está en centímetros
                        resultado = superficie * 10000
                    Case 6  'El dibujo está en metros
                        resultado = superficie   'No haces nada. Ya esta en metros.
                    Case 14 'El dibujo está en decímetros
                        resultado = superficie * 100
                End Select
            Else
                resultado = 0
            End If
            DameUnidAcad = FormatNumber(resultado, 2, TriState.True)
        End Function
        Public Function FicheroEstaEnUso(ByVal pPath As String) As Boolean
            Dim stream As IO.FileStream = Nothing
            Try
                Dim file As IO.FileInfo = New FileInfo(pPath)
                stream = file.Open(FileMode.Open, FileAccess.ReadWrite, FileShare.None)
            Catch __unusedIOException1__ As IOException
                Return True
            Finally
                If stream IsNot Nothing Then stream.Close()
            End Try
            '
            Return False
        End Function
#Region "ENUMERACIONES"
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
#End Region
    End Class
End Namespace