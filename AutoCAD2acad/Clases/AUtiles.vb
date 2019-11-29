Imports System.Diagnostics
Imports System.Collections
Imports System.Windows.Forms
Imports System.Drawing
Imports System.Windows.Media.Imaging
Imports System.Reflection


Imports Autodesk.AutoCAD.Interop
Imports Autodesk.AutoCAD.Interop.Common
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.EditorInput

Namespace A2acad
    Partial Public Class A2acad
        Public Sub PonLog(ByVal quetexto As String, Optional queF As String = "", Optional ByVal borrar As Boolean = False)
            If queF = "" Then queF = _appLog
            If queF = "" Then Exit Sub
            If borrar = True Then IO.File.Delete(queF)
            If quetexto.EndsWith(vbCrLf) = False Then quetexto &= vbCrLf
            IO.File.AppendAllText(queF, Date.Now & vbTab & quetexto)
        End Sub
        Public Function GetPathActiveDocument() As String
            Dim doc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim hs As HostApplicationServices = HostApplicationServices.Current
            Dim strPath As String = hs.FindFile(doc.Name, doc.Database, FindFileHint.[Default])
            Return strPath
        End Function
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
            '
            oDInf = Nothing
            '
            Return resultado
        End Function
        '
        Public Function CoordenadasDameTransformadas(quePt As Object, desde As AcCoordinateSystem, hasta As AcCoordinateSystem, esVector As Boolean, Optional objNormal As Object = Nothing) As Double()
            ' quePt = Array de double con las coordenadas  X, Y, Z (Se interpretará como Vector o Point, según valor de esVector)
            ' esVector = True (Vector)  / esVector=False(Point)
            Dim resultado As Double() = Nothing
            resultado = oAppA.ActiveDocument.Utility.TranslateCoordinates(quePt, AcCoordinateSystem.acWorld, AcCoordinateSystem.acOCS, esVector, New Double() {0, 0, 0})
            '
            Return resultado
        End Function

        'Private Sub CommandButton1_Click()
        '        Dim Wcsp(0 To 2) As Double
        '        Dim ocsp
        '        Dim Bref As AcadBlockReference
        '        Dim Myent As AcadEntity
        '        On Error Resume Next
        '        Me.Hide
        '        ThisDrawing.StartUndoMark
        '        Do
        '            ThisDrawing.Utility.GetEntity Bref, pick, "xx"
        '               DoEvents
        '            For Each Myent In ThisDrawing.Blocks(Bref.Name)
        '                If Myent.ObjectName = "AcDbPoint" Then
        '                    ocsp = Myent.Coordinates
        '                    BOcsToWcs Bref, ocsp, Wcsp
        ''                       For i = 0 To 2
        '                           Debug.Print Wcsp(i)
        '                       Next
        '                    Debug.Print "***********"
        '                    ThisDrawing.ModelSpace.AddCircle Wcsp, 500
        '                End If
        '            Next
        '            Debug.Print Err.Number & "--" & Err.Description
        '  Loop While Err.Number <> -2147352567
        '        ThisDrawing.EndUndoMark
        '        ThisDrawing.SendCommand "U" & Chr(13)
        '  Me.Show 0

        'End Sub

        Public Function BloqueDameCoordenadasOCStoWCS(b As AcadBlockReference, Optional ocsp As Object = Nothing) As Double()
            Dim resultado(2) As Double
            '
            Dim X As Double, Y As Double, R As Double, P As Double, Z As Double
            Dim WcsX As Double, WcsY As Double, WcsZ As Double
            Dim Instrpoint As Object
            Dim V1 As Integer = 0 : Dim V2 As Integer = 0
            Const Pi As Double = 3.14159265358979
            X = b.XScaleFactor
            Y = b.YScaleFactor
            Z = b.ZScaleFactor
            R = b.Rotation
            Instrpoint = b.InsertionPoint
            '
            If CDbl(ocsp(0)) > 0 And CDbl(ocsp(1)) = 0 Then
                R = R + 0
            ElseIf CDbl(ocsp(0)) > 0 And CDbl(ocsp(1)) > 0 Then '
                R = R + Math.Atan(CDbl(ocsp(1)) / CDbl(ocsp(0)))
            ElseIf CDbl(ocsp(0)) = 0 And CDbl(ocsp(1)) > 0 Then
                R = R + Pi / 2
            ElseIf CDbl(ocsp(0)) < 0 And CDbl(ocsp(1)) > 0 Then
                R = R + Math.Atan(CDbl(ocsp(1)) / CDbl(ocsp(0))) + Pi
            ElseIf CDbl(ocsp(0)) < 0 And CDbl(ocsp(1)) > 0 Then
                R = R + Pi
            ElseIf CDbl(ocsp(0)) < 0 And CDbl(ocsp(1)) < 0 Then
                R = R + Math.Atan(CDbl(ocsp(1)) / CDbl(ocsp(0))) + Pi
            ElseIf CDbl(ocsp(0)) = 0 And CDbl(ocsp(1)) < 0 Then
                R = R + Pi * 3 / 2
            ElseIf CDbl(ocsp(0)) > 0 And CDbl(ocsp(1)) < 0 Then
                R = R + Math.Atan(CDbl(ocsp(1)) / CDbl(ocsp(0))) + 2 * Pi
            Else
                R = R + 0
            End If

            P = (CDbl(ocsp(0)) ^ 2 + CDbl(ocsp(1)) ^ 2 + CDbl(ocsp(2)) ^ 2) ^ 0.5
            '
            WcsX = P * Math.Cos(R) * X : WcsY = P * Math.Sin(R) * Y : WcsZ = P * CDbl(ocsp(2))
            '
            resultado(0) = WcsX + Instrpoint(0)
            resultado(1) = WcsY + Instrpoint(1)
            resultado(2) = WcsZ + Instrpoint(2)
            '
            Return resultado
        End Function

        Public Function BloqueDameCoordenadasOCStoWCS(b As BlockReference, Optional ocsp As Object = Nothing) As Double()
            Dim resultado(2) As Double
            '
            Dim X As Double, Y As Double, R As Double, P As Double, Z As Double
            Dim WcsX As Double, WcsY As Double, WcsZ As Double
            Dim Instrpoint As Object
            Dim V1 As Integer = 0 : Dim V2 As Integer = 0
            Const Pi As Double = 3.14159265358979
            X = b.ScaleFactors.X
            Y = b.ScaleFactors.Y
            Z = b.ScaleFactors.Z
            R = b.Rotation
            Instrpoint = b.Position
            '
            '
            If CDbl(ocsp(0)) > 0 And CDbl(ocsp(1)) = 0 Then
                R = R + 0
            ElseIf CDbl(ocsp(0)) > 0 And CDbl(ocsp(1)) > 0 Then
                R = R + Math.Atan(CDbl(ocsp(1)) / CDbl(ocsp(0)))
            ElseIf CDbl(ocsp(0)) = 0 And CDbl(ocsp(1)) > 0 Then
                R = R + Pi / 2
            ElseIf CDbl(ocsp(0)) < 0 And CDbl(ocsp(1)) > 0 Then
                R = R + Math.Atan(CDbl(ocsp(1)) / CDbl(ocsp(0))) + Pi
            ElseIf CDbl(ocsp(0)) < 0 And CDbl(ocsp(1)) > 0 Then
                R = R + Pi
            ElseIf CDbl(ocsp(0)) < 0 And CDbl(ocsp(1)) < 0 Then
                R = R + Math.Atan(CDbl(ocsp(1)) / CDbl(ocsp(0))) + Pi
            ElseIf CDbl(ocsp(0)) = 0 And CDbl(ocsp(1)) < 0 Then
                R = R + Pi * 3 / 2
            ElseIf CDbl(ocsp(0)) > 0 And CDbl(ocsp(1)) < 0 Then
                R = R + Math.Atan(CDbl(ocsp(1)) / CDbl(ocsp(0))) + 2 * Pi
            Else
                R = R + 0
            End If

            P = (CDbl(ocsp(0)) ^ 2 + CDbl(ocsp(1)) ^ 2 + CDbl(ocsp(2)) ^ 2) ^ 0.5
            '
            WcsX = P * Math.Cos(R) * X : WcsY = P * Math.Sin(R) * Y : WcsZ = P * CDbl(ocsp(2))
            '
            resultado(0) = WcsX + Instrpoint(0)
            resultado(1) = WcsY + Instrpoint(1)
            resultado(2) = WcsZ + Instrpoint(2)
            '
            Return resultado
        End Function

        Public Function GetAllGridEntries(ByVal grid As System.Windows.Forms.PropertyGrid) As GridItemCollection
            If grid Is Nothing Then Throw New ArgumentNullException("grid")
            Dim view As Object = grid.[GetType]().GetField("gridView", System.Reflection.BindingFlags.NonPublic Or BindingFlags.Instance).GetValue(grid)
            Return CType(view.[GetType]().InvokeMember("GetAllGridEntries", BindingFlags.InvokeMethod Or BindingFlags.NonPublic Or BindingFlags.Instance, Nothing, view, Nothing), GridItemCollection)
        End Function

        '' Permite escalar el dibujo para ponerlo en metros.
        '' Pide 2 puntos para calcular medida y nos pide cuando debe medir esto en metros.
        Public Function EscalaDibujoMetros() As Double
            Dim resultado As Double = 0
            ''
            If oAppA Is Nothing Then _
                oAppA = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
            'oDoc = oAppA.ActiveDocument
            oAppA.ActiveDocument.Activate()
            Dim oPt1 As Object = Nothing
            'Dim oPt2 As Object = Nothing
            Dim medidaorigen As Double = 0
            Dim medidadestino As Double = 0
            Dim queescala As Double = 0
            Try
                'oPunto = PuntoDame_NET("Punto de inserción de la tabla:")
                oPt1 = oAppA.ActiveDocument.Utility.GetPoint(, vbLf & "Primer punto para MEDIR:")
                'oPt2 = oAppA.ActiveDocument.Utility.GetPoint(oPt1, vbLf & "Segundo punto para MEDIR:")
                medidaorigen = oAppA.ActiveDocument.Utility.GetDistance(oPt1, vbLf & "Segundo punto para MEDIR:")
                medidadestino = oAppA.ActiveDocument.Utility.GetDistance(, "Medida final en METROS")
            Catch ex As System.Exception
                'MsgBox(ex.Message)
                resultado = 0
                GoTo FINAL
                Exit Function
            End Try
            ''
            If oPt1 Is Nothing Or medidaorigen = 0 Or medidadestino = 0 Then
                resultado = 0
                GoTo FINAL
                Exit Function
            Else
                queescala = medidadestino / medidaorigen
                resultado = queescala
            End If
            ''
            Try
                Dim cadena As String = "._scale t  0,0 " & queescala & " "
                oAppA.DocumentManager.MdiActiveDocument.SendStringToExecute(cadena, True, False, False)
            Catch ex As System.Exception
                MsgBox(ex)
                resultado = 0
            End Try
            ''
FINAL:
            Return resultado
        End Function

        ''
        '' 
        ''' <summary>
        ''' Le daremos array de ancho, largo (A4, A3, A2, A1 o A0), el largo actual (X) y el ancho actual (Y)
        ''' Nos devolverá qué escala tenemos que aplicar para encajarlo en el papel indicado.
        ''' </summary>
        ''' <param name="queDin"></param>
        ''' <param name="queL"></param>
        ''' <param name="queA"></param>
        ''' <returns></returns>
        Public Function DameEscala(queDin As Array, queL As Double, queA As Double) As Double
            Dim resultado As Double = 1
            Dim escalaX As Double
            Dim escalaY As Double
            Dim largoIni As Double
            Dim anchoIni As Double
            Dim largoFin As Double
            Dim anchoFin As Double
            ''
            If queL > queA Then     '' En Horizontal
                largoIni = queDin(0)
                anchoIni = queDin(1)
            ElseIf queL < queA Then '' En Vertical
                largoIni = queDin(1)
                anchoIni = queDin(0)
            End If
            ''
            escalaX = Math.Abs(largoIni / queL)
            escalaY = Math.Abs(anchoIni / queA)
            largoFin = queL * escalaX
            anchoFin = queA * escalaY
            ''
            If queL > largoIni Then
                '' el valor menor. Porque hay que reducir
                resultado = Math.Max(escalaX, escalaY)
            Else
                '' el valor mayor. Porque hay que ampliar
                resultado = Math.Min(escalaX, escalaY)
            End If
            ''
            Return resultado
        End Function

        Public Function DameAcadPath(queVersion As String) As String
            Dim resultado As String = ""
            ' Buscamos en diferentes ubicaciones, según la versión de AutoCAD 2020 instalada.
            ' AutoCAD 2020 o AutoCAD Mechanical 2020 o AutoCAD Architecture 2020
            Dim rk As Microsoft.Win32.RegistryKey
            Dim version As String() = Nothing
            Select Case queVersion
                Case 2016
                    version = New String() {"SOFTWARE\Autodesk\AutoCAD\R20.1\ACAD-F001",
                    "SOFTWARE\Autodesk\AutoCAD\R20.1\ACAD-F001:40A",
                    "SOFTWARE\Autodesk\AutoCAD\R20.1\ACAD-F004:40A",
                    "SOFTWARE\Autodesk\AutoCAD\R20.1\ACAD-F005:40A"}
                Case 2017
                    version = New String() {"SOFTWARE\Autodesk\AutoCAD\R21.0\ACAD-0001",
                    "SOFTWARE\Autodesk\AutoCAD\R21.0\ACAD-0001:40A",
                    "SOFTWARE\Autodesk\AutoCAD\R21.0\ACAD-0004:40A",
                    "SOFTWARE\Autodesk\AutoCAD\R21.0\ACAD-0005:40A"}
                Case 2018
                    version = New String() {"SOFTWARE\Autodesk\AutoCAD\R22.0\ACAD-1001",
                    "SOFTWARE\Autodesk\AutoCAD\R22.0\ACAD-1001:40A",
                    "SOFTWARE\Autodesk\AutoCAD\R22.0\ACAD-1004:40A",
                    "SOFTWARE\Autodesk\AutoCAD\R22.0\ACAD-1005:40A"}
                Case 2019
                    version = New String() {"SOFTWARE\Autodesk\AutoCAD\R23.0\ACAD-2001",
                    "SOFTWARE\Autodesk\AutoCAD\R23.0\ACAD-2001:40A",
                    "SOFTWARE\Autodesk\AutoCAD\R23.0\ACAD-2004:40A",
                    "SOFTWARE\Autodesk\AutoCAD\R23.0\ACAD-2005:40A"}
                Case 2020
                    version = New String() {"SOFTWARE\Autodesk\AutoCAD\R23.1\ACAD-3001",
                    "SOFTWARE\Autodesk\AutoCAD\R23.1\ACAD-3001:40A",
                    "SOFTWARE\Autodesk\AutoCAD\R23.1\ACAD-3004:40A",
                    "SOFTWARE\Autodesk\AutoCAD\R23.1\ACAD-3005:40A"}
            End Select
            '
            Try
                rk = My.Computer.Registry.LocalMachine.OpenSubKey(version(0), False)
                resultado = rk.GetValue("GlobUPILocation")
            Catch ex As System.Exception
                For x As Integer = 1 To version.Length - 1
                    Dim queRk As String = version(x)
                    Try
                        rk = My.Computer.Registry.LocalMachine.OpenSubKey(queRk)
                        resultado = rk.GetValue("AcadLocation")
                        Exit For
                    Catch ex1 As System.Exception
                        resultado = ""
                        Continue For
                    End Try
                Next
            End Try
            '' 
            Return resultado
        End Function
        '
        Public Sub PonTRUSTEDPATHS()
            Dim str_TR As String = Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable("TRUSTEDPATHS")
            str_TR.Replace(";;", ";")
            str_TR.Replace(";;", ";")
            If str_TR = "." Then str_TR = ""
            ''
            Dim C_Paths As String = LCase(str_TR)

            Dim Old_Path_Ary As List(Of String) = New List(Of String)
            Old_Path_Ary.AddRange(C_Paths.Split(";"))

            Dim New_paths As List(Of String) = New List(Of String)

            New_paths.Add(My.Application.Info.DirectoryPath)
            New_paths.Add(My.Application.Info.DirectoryPath & "\..\..\Resources")

            For Each Str As String In New_paths
                If Not Old_Path_Ary.Contains(LCase(Str)) Then
                    Old_Path_Ary.Add(Str)
                End If
            Next
            '' Quitar los que están vacios en 
            For x As Integer = Old_Path_Ary.Count - 1 To 0 Step -1
                If Old_Path_Ary(x) = "" Then
                    Old_Path_Ary.RemoveAt(x)
                End If
            Next
            ''
            Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("TRUSTEDPATHS", String.Join(";", Old_Path_Ary.ToArray()))
        End Sub

        Public Sub CierraDibujo(queFullname As String)
            If oAppA Is Nothing Then _
    oAppA = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
            '' Activar Capa 0 primero.
            For Each queDoc As AcadDocument In oAppA.Documents
                If queDoc.FullName = queFullname Then
                    queDoc.Close()
                    Exit For
                End If
            Next
        End Sub

        Public Sub CierraDibujoTodos()
            If oAppA Is Nothing Then _
oAppA = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
            'oAppA.Documents.Close()
            '' Activar Capa 0 primero.
            For Each queDoc As AcadDocument In oAppA.Documents
                queDoc.Activate()
                queDoc.Save()
                'DameDependencias()
            Next
            oAppA.Documents.Close()
        End Sub
        Public Function DocumentoEstaYaAbierto(queDwg As String, Optional activar As Boolean = True) As Boolean
            Dim resultado As Boolean = False
            ''
            For Each queDoc As Autodesk.AutoCAD.Interop.AcadDocument In oAppA.Documents
                If queDoc.FullName.ToUpper = queDwg.ToUpper Then
                    If activar = True Then oAppA.ActiveDocument = queDoc
                    resultado = True
                    Exit For
                End If
            Next
            ''
            Return resultado
        End Function
        '
        Public Function DocumentoAbiertoDame(queDwg As String) As Autodesk.AutoCAD.Interop.AcadDocument
            Dim resultado As Autodesk.AutoCAD.Interop.AcadDocument = Nothing
            ''
            For Each queDoc As Autodesk.AutoCAD.Interop.AcadDocument In oAppA.Documents
                If queDoc.FullName.ToUpper = queDwg.ToUpper Then
                    resultado = queDoc
                    Exit For
                End If
            Next
            ''
            Return resultado
        End Function
        '
        Public Function DocumentoAbiertoCierra(queDwg As String) As Boolean
            Dim resultado As Boolean = Nothing
            ''
            For Each queDoc As Autodesk.AutoCAD.Interop.AcadDocument In oAppA.Documents
                If queDoc.FullName.ToUpper = queDwg.ToUpper Then
                    queDoc.Close(False)
                    resultado = True
                    Exit For
                End If
            Next
            ''
            Return resultado
        End Function
        Public Function DocumentoGuardaDWG(carpetainicio As String) As String
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
        Public Function DocumentoAbreDWG(carpetainicio As String) As String
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

        'Public Function DependenciasDame(Optional vermensaje As Boolean = False) As ArrayList
        '    Dim resultado As New ArrayList
        '    If oAppA Is Nothing Then _
        'oAppA = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)

        '    Dim objFDLCol As AcadFileDependencies
        '    Dim objFDL As AcadFileDependency

        '    objFDLCol = oAppA.ActiveDocument.FileDependencies
        '    ' MsgBox("Nº de File Dependency List = " & objFDLCol.Count)   ' & ".")


        '    Dim mensaje As String = ""
        '    Try
        '        If objFDLCol IsNot Nothing AndAlso objFDLCol.Count > 0 Then
        '            For Each objFDL In objFDLCol
        '                mensaje &= objFDL.FullFileName & vbCrLf
        '                If resultado.Contains(objFDL.FullFileName) = False Then resultado.Add(objFDL.FullFileName)
        '            Next
        '            If vermensaje Then _
        '            Autodesk.AutoCAD.ApplicationServices.Application.ShowAlertDialog(mensaje)
        '        Else
        '            If vermensaje Then _
        '            Autodesk.AutoCAD.ApplicationServices.Application.ShowAlertDialog("No hay Dependencias...")
        '        End If
        '    Catch ex As System.Exception
        '        '' no hacemos nada.
        '    End Try
        '    ''
        '    Return resultado
        'End Function
        Public Function XrefComprueba(Optional vermensaje As Boolean = False) As Hashtable
            Dim resultado As Hashtable = New Hashtable
            If oAppA Is Nothing Then _
oAppA = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
            ''
            Dim mensaje As String = ""
            For Each oEref As AcadBlock In oAppA.ActiveDocument.Blocks
                If TypeOf oEref Is RasterImage Then
                    'Debug.Print(oEref.Name)
                End If
                If oEref.IsXRef AndAlso IO.File.Exists(oEref.Path) = False Then
                    If resultado.Contains(oEref.EffectiveName) = False Then
                        resultado.Add(oEref.EffectiveName, oEref.Path)
                        mensaje &= oEref.EffectiveName & " / " & oEref.Path & vbCrLf
                    End If
                End If
            Next
            ''
            If vermensaje = True Then
                Autodesk.AutoCAD.ApplicationServices.Application.ShowAlertDialog(mensaje)
                'MsgBox(mensaje)
            End If
            '
            Return resultado
        End Function

        ''
        Public Function XrefImagenDame(Optional vermensaje As Boolean = False) As Hashtable
            Dim resultado As Hashtable = New Hashtable
            If oAppA Is Nothing Then _
oAppA = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
            ''
            Dim mensaje As String = ""
            Dim oDict As AcadDictionary
            oDict = oAppA.ActiveDocument.Dictionaries.Item("ACAD_IMAGE_DICT")
            Dim oImage As AcadRasterImage = Nothing
            Dim oImageDef As Object
            MsgBox(oDict.Count)
            For Each oImageDef In oDict
                Dim oImg As AcadRasterImage = CType(oImageDef, AcadRasterImage)
                If resultado.ContainsKey(oImg.Name) = False Then
                    resultado.Add(oImg.Name, oImg.ImageFile)
                    mensaje &= oImg.Name & " / " & oImg.ImageFile & vbCrLf
                    'If TypeOf oImageDef Is AcadRasterImage Then
                    'MsgBox(oImageDef.GetType.ToString)
                End If
            Next

            If vermensaje = True Then
                Autodesk.AutoCAD.ApplicationServices.Application.ShowAlertDialog(mensaje)
                'MsgBox(mensaje)
            End If
            '
            Return resultado
        End Function
        ''
        Public Function DameUltimoObjeto() As AcadEntity
            If oAppA Is Nothing Then _
        oAppA = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
            ''
            Return oAppA.ActiveDocument.ModelSpace.Item(oAppA.ActiveDocument.ModelSpace.Count - 1)
        End Function
        ''
        Public Function PuntoEstaDentroPolyline(ByRef p2d As Autodesk.AutoCAD.Geometry.Point3d, ByRef Ligne As Polyline) As Boolean
            Dim OkRetour As Boolean = False
            '===========================================================================
            'Gestión de errores
            If IsNothing(p2d) = True Or IsNothing(Ligne) = True Then Return OkRetour
            'Si ligne non fermée
            If Ligne.Closed = False Then Return False
            '===========================================================================
            '===========================================================================
            'Convertir punto 3D en 2D
            Dim pt As New Autodesk.AutoCAD.Geometry.Point2d(p2d.X, p2d.Y)
            'Si el punto esta sobre la polilinea
            Dim vn As Integer = Ligne.NumberOfVertices
            Dim colpt() As Autodesk.AutoCAD.Geometry.Point2d = Nothing
            ReDim colpt(vn)
            For i As Integer = 0 To vn - 1
                Dim pts As Autodesk.AutoCAD.Geometry.Point2d = Ligne.GetPoint2dAt(i)
                colpt(i) = New Autodesk.AutoCAD.Geometry.Point2d(pts.X, pts.Y)
                'Si el punto está sobre un segmento de la polilinea
                Dim seg As Autodesk.AutoCAD.Geometry.Curve2d = Nothing
                Dim segType As SegmentType = Ligne.GetSegmentType(i)
                If (segType = SegmentType.Arc) Then
                    seg = Ligne.GetArcSegment2dAt(i)
                ElseIf (segType = SegmentType.Line) Then
                    seg = Ligne.GetLineSegment2dAt(i)
                End If
                If IsNothing(seg) = False Then
                    OkRetour = seg.IsOn(pt)
                    If OkRetour = True Then
                        Return True
                    End If
                End If
            Next
            'Punto final = punto inicial de la polilinea
            colpt(vn) = Ligne.GetPoint2dAt(0)
            '===========================================================================
            Dim RetFonction As Double
            RetFonction = wn_PnPoly(pt, colpt, vn)
            If RetFonction = 0 Then
                OkRetour = False
            Else
                OkRetour = True
            End If
            ''
            Return OkRetour
        End Function

        ''' <summary>
        ''' isLeft(): tests if a point is Left|On|Right of an infinite line.
        ''' Input: three points P0, P1, and P2
        ''' Return:
        ''' Es mayor de 0 para P2 a la izquierda de la linea a través de P0 y P1
        ''' Es igual a 0 para P2 sobre la linea
        ''' Es menor de 0 para P2 a la derecha de la linea
        ''' Ver Algorithm 1 Area of Triangles and Polygons
        ''' </summary>
        ''' <param name="P0"></param>
        ''' <param name="P1"></param>
        ''' <param name="P2"></param>
        ''' <returns></returns>
        Private Function isLeft(ByVal P0 As Autodesk.AutoCAD.Geometry.Point2d, ByVal P1 As Autodesk.AutoCAD.Geometry.Point2d, ByVal P2 As Autodesk.AutoCAD.Geometry.Point2d) As Double
            Return ((P1.X - P0.X) * (P2.Y - P0.Y) - (P2.X - P0.X) * (P1.Y - P0.Y))
        End Function
        '**************************************************************************************
        '**************************************************************************************
        '// wn_PnPoly(): winding number test for a point in a polygon
        '// Input: P = a point,
        '// V[] = vertex points of a polygon V[n+1] with V[n]=V[0]
        '// Return: wn = the winding number (=0 only when P is outside)
        'int
        ''' <summary>
        ''' wn_PnPoly: Testear si un punto está dentro de un poligono
        ''' </summary>
        ''' <param name="P">Punto 2D a testear</param>
        ''' <param name="V">Array de puntos 2D que definen el poligono. V[n+1] with V[n]=V[0]</param>
        ''' <param name="n">Valor de retorno</param>
        ''' <returns>Número de coincidencias (=0 el punto P estará fuera)</returns>
        Private Function wn_PnPoly(ByVal P As Autodesk.AutoCAD.Geometry.Point2d, ByVal V() As Autodesk.AutoCAD.Geometry.Point2d, ByVal n As Double) As Double
            Dim wn As Double = 0 '// the winding number counter
            '// loop through all edges of the polygon
            For i As Integer = 0 To n - 1
                If (V(i).Y <= P.Y) Then '// eje desde V[i] a V[i+1]
                    '// inicio y <= P.y
                    If (V(i + 1).Y > P.Y) Then '// an upward crossing
                        If (isLeft(V(i), V(i + 1), P) > 0) Then '// Si P esta a la izquierda del eje
                            wn = wn + 1 ' Tiene una intersección hacia arriba
                        End If
                    End If
                Else '// inicio y > P.y (No se necesita testear)
                    If (V(i + 1).Y <= P.Y) Then '// a downward crossing
                        If (isLeft(V(i), V(i + 1), P) < 0) Then '// Si P esta a la derecha del eje
                            wn = wn - 1 ' Tiene una intersección hacia abajo
                        End If
                    End If
                End If
            Next
            Return wn
        End Function
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

        Public Sub LinklabelRellenaLinks(ByRef quelblL As LinkLabel, fullpath As String, Optional queQuito As String = "")
            Dim dirSin As String = ""
            Dim dirSinSep As String = ""
            Dim dirLink As String = ""
            Dim partes() As String = fullpath.Split("\"c)
            If queQuito = "" Then
                Dim partes1 As String() = Nothing
                partes.CopyTo(partes1, 1)
                dirSin = Join(partes1, "\")
                dirLink = partes(0)
            Else
                If queQuito.EndsWith("\") = False Then queQuito &= "\"
                dirSin = fullpath.Replace(queQuito, "")     'Camino completo, sin el camino inicial pathDirCoB2D
                dirLink = queQuito
            End If
            '
            dirSinSep = dirSin.Replace("\", " ---> ")
            quelblL.Text = dirSinSep
            quelblL.Links.Clear()
            ' Partir en trozos los subdirectorios y crear los links.
            Dim trozos As String() = dirSin.Split("\"c)
            For Each trozo As String In trozos
                Dim oLink As LinkLabel.Link = quelblL.Links.Add(dirSinSep.IndexOf(trozo), trozo.Length, trozo)
                oLink.Description = trozo
                oLink.Name = trozo
                oLink.Enabled = True
                oLink.Tag = IO.Path.Combine(dirLink, trozo)
                dirLink = oLink.Tag
            Next
        End Sub
        Public Function TextoDameTamañoPixels(queTexto As String, queFont As System.Drawing.Font, ancho As Boolean) As Integer
            Dim resultado As Integer = 0
            If ancho Then
                resultado = System.Windows.Forms.TextRenderer.MeasureText(queTexto, queFont).Width
            Else
                resultado = System.Windows.Forms.TextRenderer.MeasureText(queTexto, queFont).Height
            End If
            Return resultado
        End Function
        Public Sub VaciaMemoria()
            Try
                GC.Collect()
                GC.WaitForPendingFinalizers()
                GC.Collect()
            Catch ex As Exception
                '
            End Try
        End Sub
    End Class
End Namespace