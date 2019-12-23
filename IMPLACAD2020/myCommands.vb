' (C) Copyright 2011 by  
'
Imports System
Imports System.Net
Imports System.IO
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Interop
Imports Autodesk.AutoCAD.Interop.Common

' This line is not mandatory, but improves loading performances
<Assembly: CommandClass(GetType(IMPLACAD.MyCommands))> 
Namespace IMPLACAD

    ' This class is instantiated by AutoCAD for each document when
    ' a command is called by the user the first time in the context
    ' of a given document. In other words, non static data in this class
    ' is implicitly per-document!
    Public Class MyCommands
        Public WithEvents oTimer As Timers.Timer
        ' The CommandMethod attribute can be applied to any public  member 
        ' function of any public class.
        ' The function should take no arguments and return nothing.
        ' If the method is an instance member then the enclosing class is 
        ' instantiated for each document. If the member is a static member then
        ' the enclosing class is NOT instantiated.
        '
        ' NOTE: CommandMethod has overloads where you can provide helpid and
        ' context menu.

        ' Modal Command with localized name
        ' AutoCAD will search for a resource string with Id "MyCommandLocal" in the 
        ' same namespace as this command class. 
        ' If a resource string is not found, then the string "MyLocalCommand" is used 
        ' as the localized command name.
        ' To view/edit the resx file defining the resource strings for this command, 
        ' * click the 'Show All Files' button in the Solution Explorer;
        ' * expand the tree node for myCommands.vb;
        ' * and double click on myCommands.resx

        <CommandMethod(regAPP, "IMPLACADMENU", "IMPLACADMENU", CommandFlags.Modal)> _
        Public Sub IMPLACADMENU() ' This method can have any name
            ' Put your command code here
            ''
            If ImplacadActivado() = False Then Exit Sub
            If ImplacadEscalaM() = False Then Exit Sub
            ''
            If EsParaTrabajar() = False Then
                MsgBox("Este fichero DWG es sólo para imprimir" & vbCrLf & vbCrLf & _
                       "Debe abrir el DWG original de trabajo")
                Exit Sub
            End If
            If oApp Is Nothing Then _
        oApp = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
            'oDoc = oApp.ActiveDocument
            '' 0.- Anular cualquier comando
            ''oApp.ActiveDocument.PostCommand(Chr(27) & Chr(27))
            CapaCeroActiva()
            Try
                Dim oMg As AcadMenuGroups = Application.MenuGroups
                Call oMg.Load(IO.Path.Combine(dirApp, "IMPLACAD.cuix"))
            Catch ex As Autodesk.AutoCAD.Runtime.Exception
            End Try
        End Sub

        <CommandMethod(regAPP, "INSERTAREDITAR", "INSERTAREDITAR", CommandFlags.Modal)> _
        Public Sub INSERTAREDITAR() ' This method can have any name
            ' Put your command code here
            ''
            If ImplacadActivado() = False Then Exit Sub
            If ImplacadEscalaM() = False Then Exit Sub
            If ImplacadActualizado() = False Then Exit Sub
            ''
            If EsParaTrabajar() = False Then
                MsgBox("Este fichero DWG es sólo para imprimir" & vbCrLf & vbCrLf & _
                       "Debe abrir el DWG original de trabajo")
                Exit Sub
            End If
            If oApp Is Nothing Then _
        oApp = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
            'oDoc = oApp.ActiveDocument
            '' 0.- Anular cualquier comando
            ''oApp.ActiveDocument.PostCommand(Chr(27) & Chr(27))
            'MsgBox("Comando ejecutado")
            If frmE IsNot Nothing Then frmE.Close()
            ''
            frmE = New frmEtiquetas
            'If Application.ShowModalDialog(frmE) = Windows.Forms.DialogResult.OK Then
            'MsgBox("Insertamos el bloque...")
            'End If
            ConfiguraDibujo()
            Application.ShowModelessDialog(Application.MainWindow.Handle, frmE, True)
        End Sub

        <CommandMethod(regAPP, "ADECUA", "ADECUA", CommandFlags.Modal)> _
        Public Sub ADECUA() ' This method can have any name
            ' Put your command code here
            If EsParaTrabajar() = False Then
                MsgBox("Este fichero DWG es sólo para imprimir" & vbCrLf & vbCrLf & _
                       "Debe abrir el DWG original de trabajo")
                Exit Sub
            End If
            ''
            If ImplacadActivado() = False Then Exit Sub
            If ImplacadEscalaM() = False Then Exit Sub
            ''
            If oApp Is Nothing Then _
        oApp = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
            ConfiguraDibujo()
            CapaCeroActiva()
            ''
            Try

                'Dim BloqueAdecua As Autodesk.AutoCAD.Interop.Common.AcadBlockReference = modImplacad.SeleccionaDameBloque()
                'oDoc.SetVariable("XCLIPFRAM", 2)
                ''
                '' Activamos las Capas de Zona de cobertura
                CapaZonaCoberturaACTDES(CapaEstado.Activar)
                ' Seleccionar objeto
                Dim oEtiqueta As AcadBlockReference = Nothing
                oEtiqueta = modImplacad.SeleccionaDameBloque()
                'While bloqueEditar Is Nothing
                'oEtiqueta = modImplacad.SeleccionaDameBloque()
                'End While
                ''
                '' Si el objeto es una etiqueta
                If (oEtiqueta IsNot Nothing) AndAlso (TypeOf oEtiqueta Is AcadBlockReference) Then
                    ''
                    '' 1.- Borrar zona de recorte, si ya tenia antes.
                    oApp.ActiveDocument.SendCommand("_xclip p  s ")
                    '' 2.- Anular cualquier comando
                    oApp.ActiveDocument.PostCommand(Chr(27) & Chr(27))
                    '' 3.- Desactivar el contorno de recorte (1= visual pero no imprime / 2=visual e imprime)
                    oApp.ActiveDocument.SendCommand("XCLIPFRAME 0 ")
                    '' 4.- Anular cualquier comando
                    oApp.ActiveDocument.PostCommand(Chr(27) & Chr(27))
                    '' 5.- Ejecutar comando _Xclip
                    oApp.ActiveDocument.SendCommand("_xclip p  N p ")
                    'oDoc.PostCommand(Chr(27) & Chr(27))
                End If
            Catch ex As System.Exception
                MsgBox("Falla el comando '_xclip' de las Express Tools")
                Exit Try
            End Try
        End Sub
        ''
        <CommandMethod(regAPP, "BALIZARSUELO", "BALIZARSUELO", CommandFlags.Modal)> _
        Public Sub BALIZARSUELO() ' This method can have any name
            ' Put your command code here
            If EsParaTrabajar() = False Then
                MsgBox("Este fichero DWG es sólo para imprimir" & vbCrLf & vbCrLf & _
                       "Debe abrir el DWG original de trabajo")
                Exit Sub
            End If
            ''
            If ImplacadActivado() = False Then Exit Sub
            If ImplacadEscalaM() = False Then Exit Sub
            ''
            If oApp Is Nothing Then _
        oApp = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
            'oDoc = oApp.ActiveDocument
            '' 0.- Anular cualquier comando
            ''oApp.ActiveDocument.PostCommand(Chr(27) & Chr(27))
            ''
            ConfiguraDibujo()
            ''
            Try

                '' Activamos las Capas de Zona de cobertura
                CapaCreaActivaBalizamientoSuelo()
                Dim pt1 As Object = Nothing
                Dim pt2 As Object = Nothing
                Application.DocumentManager.CurrentDocument.Editor.WriteMessage(vbLf)
                pt1 = oApp.ActiveDocument.Utility.GetPoint(, vbLf & "Primer Punto Baliza SUELO :")
                pt2 = oApp.ActiveDocument.Utility.GetPoint(pt1, vbLf & "Siguiente Punto Baliza SUELO :")
                While pt2 IsNot Nothing
                    If pt1 IsNot Nothing Then
                        oApp.ActiveDocument.ModelSpace.AddLine(pt1, pt2)
                    Else
                        Exit While
                    End If
                    pt1 = pt2
                    Try
                        pt2 = oApp.ActiveDocument.Utility.GetPoint(pt1, vbLf & "Siguiente Punto Baliza SUELO :")
                    Catch ex As System.Exception
                        Exit While
                    End Try
                End While
                'If resultado IsNot Nothing Then
                'Dim oPoint As Point3d = resultado.Value
                'Dim puntoInserta(0 To 2) As Double
                'puntoInserta(0) = oPoint.X
                'puntoInserta(1) = oPoint.Y
                'puntoInserta(2) = oPoint.Z
                'Dim inserta As AcadBlockReference = Nothing
                'Dim ultimoBlo As Object = Nothing
                'inserta = oApp.ActiveDocument.ActiveLayout.Block.InsertBlock(puntoInserta, ultimoCaminoDWG, escalaTotal, escalaTotal, escalaTotal, 0)
                ''
                'Dim comando As String = "(command ""_line"")" & vbCr & "(while(>(getvar ""cmdactive"")0)(command pause))"
                'oApp.ActiveDocument.SendCommand(comando & vbCr)
                'oApp.ActiveDocument.SendCommand("_line pause")
                'oApp.ActiveDocument.SendCommand( _
                '"(Command ""_line"" pause)" & vbCr)
                ''CommandString = "(Command ""-Insert"" """ & Selection & """ pause ""1"" ""1"" ""0"")"
                ''ThisDrawing.SendCommand(CommandString & vbCr)
                '' Ponemos la capa 0 como activa al terminar.
            Catch ex As System.Exception
                Exit Try
            End Try
            ''
            CapaCeroActiva()
        End Sub
        ''
        <CommandMethod(regAPP, "BALIZARPARED", "BALIZARPARED", CommandFlags.Modal)> _
        Public Sub BALIZARPARED() ' This method can have any name
            ' Put your command code here
            If EsParaTrabajar() = False Then
                MsgBox("Este fichero DWG es sólo para imprimir" & vbCrLf & vbCrLf & _
                       "Debe abrir el DWG original de trabajo")
                Exit Sub
            End If
            ''
            If ImplacadActivado() = False Then Exit Sub
            If ImplacadEscalaM() = False Then Exit Sub
            ''
            If oApp Is Nothing Then _
        oApp = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
            'oDoc = oApp.ActiveDocument
            '' 0.- Anular cualquier comando
            ''oApp.ActiveDocument.PostCommand(Chr(27) & Chr(27))
            ''
            ConfiguraDibujo()
            ''
            Try
            '' Activamos las Capas de Zona de cobertura
            CapaCreaActivaBalizamientoPared()
            Dim pt1 As Object = Nothing
            Dim pt2 As Object = Nothing
            Application.DocumentManager.CurrentDocument.Editor.WriteMessage(vbLf)
            pt1 = oApp.ActiveDocument.Utility.GetPoint(, vbLf & "Primer Punto Baliza PARED :")
            pt2 = oApp.ActiveDocument.Utility.GetPoint(pt1, vbLf & "Siguiente Punto Baliza PARED :")
            While pt2 IsNot Nothing
                If pt1 IsNot Nothing Then
                    oApp.ActiveDocument.ModelSpace.AddLine(pt1, pt2)
                Else
                    Exit While
                End If
                pt1 = pt2
                Try
                    pt2 = oApp.ActiveDocument.Utility.GetPoint(pt1, vbLf & "Siguiente Punto Baliza PARED :")
                Catch ex As System.Exception
                    Exit While
                End Try
            End While
            ''
            'Dim comando As String = "(command ""_line"")" & vbCr & "(while(>(getvar ""cmdactive"")0)(command pause))"
            'oApp.ActiveDocument.SendCommand(comando & vbCr)
            'oApp.ActiveDocument.SendCommand("_line \")
            'oApp.ActiveDocument.SendCommand( _
            '"(Command ""_line"" pause)" & vbCr)
            ''CommandString = "(Command ""-Insert"" """ & Selection & """ pause ""1"" ""1"" ""0"")"
            ''ThisDrawing.SendCommand(CommandString & vbCr)
                '' Ponemos la capa 0 como activa al terminar.
            Catch ex As System.Exception
                Exit Try
            End Try
            CapaCeroActiva()
        End Sub
        ''
        <CommandMethod(regAPP, "BALIZARESCALERA", "BALIZARESCALERA", CommandFlags.Modal)> _
        Public Sub BALIZARESCALERA() ' This method can have any name
            ' Put your command code here
            If EsParaTrabajar() = False Then
                MsgBox("Este fichero DWG es sólo para imprimir" & vbCrLf & vbCrLf & _
                       "Debe abrir el DWG original de trabajo")
                Exit Sub
            End If
            ''
            If ImplacadActivado() = False Then Exit Sub
            If ImplacadEscalaM() = False Then Exit Sub
            ''
            If oApp Is Nothing Then _
        oApp = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
            'oDoc = oApp.ActiveDocument
            '' 0.- Anular cualquier comando
            ''oApp.ActiveDocument.PostCommand(Chr(27) & Chr(27))
            ''
            ConfiguraDibujo()
            ''
            Try
                '' Activamos las Capas de Zona de cobertura
                CapaCreaActivaBalizamientoEscalera()
                '' Nombre de la escalera
                Dim nombreescalera As String = oApp.ActiveDocument.Utility.GetString(1, vbLf & "Nombre de la escalera : ")
                '' Número de escalones
                Dim numeroescalores As Integer = oApp.ActiveDocument.Utility.GetInteger(vbLf & "Número de escalones : ")
                '' Opción Ascendente o Descendente
                Dim ascendentedescendente As String = GetOpcionDame_NET()
                '' Punto de inserción y primer punto para medir escalón
                Dim oPt As Object = GetPuntoDame_NET(vbLf & "Ancho escalón--> Primer punto : ")
                Dim oPt1 As Object = Nothing
                '' Segunda punto para medir escalón.
                'Dim oPt1 As Object = GetPuntoDame_NET(vbLf & "Ancho escalón--> Segundo punto : ")
                '' Medida con el ancho del escalón.
                Dim anchoescalon As Double = 0
                If oPt IsNot Nothing Then
                    'anchoescalon = oApp.ActiveDocument.Utility.GetDistance(CType(oPt, Point3d).ToArray, "Ancho escalón--> Segundo punto : ")
                    oPt1 = oApp.ActiveDocument.Utility.GetPoint(CType(oPt, Point3d).ToArray, "Ancho escalón--> Segundo punto : ")
                    ''
                    ' Create a line from the base point and the last point entered
                    Dim lineObj As AcadLine = oApp.ActiveDocument.ModelSpace.AddLine(CType(oPt, Point3d).ToArray, oPt1)
                    anchoescalon = lineObj.Length
                    If anchoescalon = 0 Then
                        CapaCeroActiva()
                        Exit Sub
                    End If
                    ''
                    Dim oBl As AcadBlockReference = oApp.ActiveDocument.ModelSpace.InsertBlock(
                        CType(oPt, Point3d).ToArray,
                        dirApp & "etiquetaescalera.dwg",
                        0.001, 0.001, 0.001, 0)

                    If oBl Is Nothing Then
                        'CapaCeroActiva()
                        Exit Try
                    End If
                    ''
                    '' Poner atributo nombreescalera
                    Dim oAtributos() As Object  ' AttributeReference
                    oAtributos = oBl.GetAttributes
                    For x As Integer = LBound(oAtributos) To UBound(oAtributos)
                        Dim oAtri As AcadAttributeReference = oAtributos(x)
                        Select Case oAtri.TagString
                            Case "NOMBREESCALERA"
                                oAtri.TextString = nombreescalera   ' & " (" & ascendentedescendente & ")"
                            Case "NUMEROESCALONES"
                                oAtri.TextString = numeroescalores
                            Case "CLASE"
                                oAtri.TextString = ascendentedescendente
                            Case "ANCHO"
                                oAtri.TextString = anchoescalon
                        End Select
                        oAtri = Nothing
                    Next
                    'oAtributos(0).TextString = nombreescalera
                    XData.XNuevo(oBl, "Clase=balizaescalera")
                    oBl.Update()
                    oBl = Nothing
                    oAtributos = Nothing
                End If
            Catch ex As System.Exception
                Exit Try
            End Try
            ''
            'Dim comando As String = "(command ""_line"")(while(>(getvar ""cmdactive"")0)(command pause))"
            'oApp.ActiveDocument.SendCommand(comando & " ")
            'oApp.ActiveDocument.SendCommand("_line ")
            '' Ponemos la capa 0 como activa al terminar.
            'Dim mensaje As String = ""
            'For Each oBe As AcadBlockReference In DameBalizasEscaleras()
            '    mensaje &= oBe.GetAttributes(0).TextString & vbCrLf ' NOMBREESCALERA
            '    mensaje &= oBe.GetAttributes(1).TextString & vbCrLf ' NUMEROESCALONES
            '    mensaje &= oBe.GetAttributes(2).TextString & vbCrLf ' CLASE
            '    mensaje &= oBe.GetAttributes(3).TextString & vbCrLf ' ANCHO
            '    mensaje &= vbCrLf & StrDup(15, "*") & vbCrLf & vbCrLf
            'Next
            'MsgBox(mensaje)
            CapaCeroActiva()
        End Sub
        ''
        <CommandMethod(regAPP, "RUTAEVACUACION", "RUTAEVACUACION", CommandFlags.Modal)> _
        Public Sub RUTAEVACUACION() ' This method can have any name
            ' Put your command code here
            ''
            If ImplacadActivado() = False Then Exit Sub
            If ImplacadEscalaM() = False Then Exit Sub
            ''
            If EsParaTrabajar() = False Then
                MsgBox("Este fichero DWG es sólo para imprimir" & vbCrLf & vbCrLf & _
                       "Debe abrir el DWG original de trabajo")
                Exit Sub
            End If
            If oApp Is Nothing Then _
        oApp = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
            'oDoc = oApp.ActiveDocument
            '' 0.- Anular cualquier comando
            ''oApp.ActiveDocument.PostCommand(Chr(27) & Chr(27))
            ''
            ConfiguraDibujo()
            Try
                '' Primaria o secundaria
                Dim primariasecundaria As String = GetOpcionPrimariaSecundariaDame()
                ''
                If primariasecundaria = "Primaria" Then
                    '' Creamos y/o Activamos "Ruta evacuación primaria"
                    CapaCreaActivaRutaEvacuacion(TipoEvacuacion.Primaria)
                    ''
                ElseIf primariasecundaria = "Accesibilidad" Then
                    '' Creamos y/o Activamos "Ruta evacuación secundaria"
                    CapaCreaActivaRutaEvacuacion(TipoEvacuacion.Accesibilidad)
                    ''
                End If
                ''
                '' Pedir puntos para dibujos lineas.
                Dim pt1 As Object = Nothing
                Dim pt2 As Object = Nothing
                Application.DocumentManager.CurrentDocument.Editor.WriteMessage(vbLf)
                pt1 = oApp.ActiveDocument.Utility.GetPoint(, vbLf & "Primer Punto RUTA :")
                pt2 = oApp.ActiveDocument.Utility.GetPoint(pt1, vbLf & "Siguiente Punto RUTA :")
                While pt2 IsNot Nothing
                    If pt1 IsNot Nothing Then
                        Dim oLine As AcadLine = oApp.ActiveDocument.ModelSpace.AddLine(pt1, pt2)
                        oLine.LinetypeScale = 1
                        oLine.Update()
                    Else
                        Exit While
                    End If
                    pt1 = pt2
                    Try
                        pt2 = oApp.ActiveDocument.Utility.GetPoint(pt1, vbLf & "Siguiente Punto RUTA :")
                    Catch ex As System.Exception
                        Exit While
                    End Try
                End While
                ''
                ''
                'Dim comando As String = "(command ""_line"")(while(>(getvar ""cmdactive"")0)(command pause))"
                'oApp.ActiveDocument.SendCommand(comando)
                'oApp.ActiveDocument.SendCommand("_line ")
            Catch ex As System.Exception
                Exit Sub
            End Try
            '' Ponemos la capa 0 como activa al terminar.
            CapaCeroActiva()
        End Sub
        ''
        <CommandMethod(regAPP, "TABLADATOS", "TABLADATOS", CommandFlags.Modal)> _
        Public Sub TABLADATOS() ' This method can have any name
            ''
            If ImplacadActivado() = False Then Exit Sub
            If ImplacadEscalaM() = False Then Exit Sub
            ''
            If EsParaTrabajar() = False Then
                MsgBox("Este fichero DWG es sólo para imprimir" & vbCrLf & vbCrLf & _
                       "Debe abrir el DWG original de trabajo")
                Exit Sub
            End If
            If oApp Is Nothing Then _
        oApp = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
            'oDoc = oApp.ActiveDocument
            ' Put your command code here
            '' 0.- Anular cualquier comando
            '' **********
            '' ********** No poner SendCommand ni PostCommand antes de hacer selecciones.
            'oApp.ActiveDocument.SendCommand(Chr(27) & Chr(27))
            ''
            ' ConfiguraDibujo()
            TablaDatosInserta()
        End Sub
        ''
        <CommandMethod(regAPP, "TABLABAESCALERAS", "TABLAESCALERAS", CommandFlags.Modal)> _
        Public Sub TABLAESCALERAS() ' This method can have any name
            ''
            If ImplacadActivado() = False Then Exit Sub
            If ImplacadEscalaM() = False Then Exit Sub
            ''
            If EsParaTrabajar() = False Then
                MsgBox("Este fichero DWG es sólo para imprimir" & vbCrLf & vbCrLf & _
                       "Debe abrir el DWG original de trabajo")
                Exit Sub
            End If
            If oApp Is Nothing Then _
        oApp = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
            'oDoc = oApp.ActiveDocument
            ' Put your command code here
            '' 0.- Anular cualquier comando
            '' **********
            '' ********** No poner SendCommand ni PostCommand antes de hacer selecciones.
            'oApp.ActiveDocument.SendCommand(Chr(27) & Chr(27))
            ''
            'ConfiguraDibujo()
            TablaEscalerasInserta()
        End Sub
        ''
        <CommandMethod(regAPP, "CAPASCOBERTURA", "CAPASCOBERTURA", CommandFlags.Modal)> _
        Public Sub CAPASCOBERTURA() ' This method can have any name
            ' Put your command code here
            If EsParaTrabajar() = False Then
                MsgBox("Este fichero DWG es sólo para imprimir" & vbCrLf & vbCrLf & _
                       "Debe abrir el DWG original de trabajo")
                Exit Sub
            End If
            ''
            If ImplacadActivado() = False Then Exit Sub
            If ImplacadEscalaM() = False Then Exit Sub
            ''
            If oApp Is Nothing Then _
        oApp = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
            'oDoc = oApp.ActiveDocument
            '' 0.- Anular cualquier comando
            ''oApp.ActiveDocument.PostCommand(Chr(27) & Chr(27))
            ''
            ConfiguraDibujo()
            CapaCeroActiva()
            ''
            '' Activar/Desactivar las Capas de Zona de cobertura
            CapaZonaCoberturaACTDES(CapaEstado.Inverso)
        End Sub

        <CommandMethod(regAPP, "GROSORLINEAS", "GROSORLINEAS", CommandFlags.Modal)> _
        Public Sub GROSORLINEAS() ' This method can have any name
            ' Put your command code here
            ''
            If ImplacadActivado() = False Then Exit Sub
            If ImplacadEscalaM() = False Then Exit Sub
            ''
            If EsParaTrabajar() = False Then
                MsgBox("Este fichero DWG es sólo para imprimir" & vbCrLf & vbCrLf & _
                       "Debe abrir el DWG original de trabajo")
                Exit Sub
            End If
            If oApp Is Nothing Then _
        oApp = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
            'oDoc = oApp.ActiveDocument
            '' 0.- Anular cualquier comando
            ''oApp.ActiveDocument.PostCommand(Chr(27) & Chr(27))
            ''
            ConfiguraDibujo()
            CapaCeroActiva()
            ''
            '' Activar/Desactivar grosores de linea
            GrosorLineasACTDES(CapaEstado.Inverso)
        End Sub
        ''
        <CommandMethod(regAPP, "ESCALAM", "ESCALAM", CommandFlags.Modal)> _
        Public Sub ESCALAM() ' This method can have any name
            If EsParaTrabajar() = False Then
                MsgBox("Este fichero DWG es sólo para imprimir" & vbCrLf & vbCrLf & _
                       "Debe abrir el DWG original de trabajo")
                Exit Sub
            End If
            ''
            If ImplacadActivado() = False Then Exit Sub
            ''
            '' Documento sin guardar
            If ImplacadGuardado() = False Then Exit Sub
            ''
            If oApp Is Nothing Then _
        oApp = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
            'oDoc = oApp.ActiveDocument
            ' Put your command code here
            '' 0.- Anular cualquier comando
            '' **********
            '' ********** No poner SendCommand ni PostCommand antes de hacer selecciones.
            'oApp.ActiveDocument.SendCommand(Chr(27) & Chr(27))
            ''
            ConfiguraDibujo()
            'TablaDatosInserta()
            Dim resultado As Double = EscalaDibujoMetros()
            If resultado <> 0 Then
                PropiedadEscribe("ESCALAM", resultado)
                ''
                '' Si ya hemos puesto equiquetas, balizas o tablas. Ponerlas
                '' en su escala correcta.
                'Dim factor As Double = 0
                '
                Dim arrEntidades As ArrayList = DameTodoImplacad()
                If Not (arrEntidades Is Nothing) AndAlso arrEntidades.Count > 0 Then
                    ''
                    For Each oEnt As AcadEntity In arrEntidades
                        If TypeOf oEnt Is AcadBlockReference Then
                            Dim oBl As AcadBlockReference = oEnt
                            If oBl.XScaleFactor <> 0.001 Then
                                oBl.ScaleEntity(oBl.InsertionPoint, 1)
                                'oBl.ScaleEntity(oBl.InsertionPoint, (0.001 / oBl.XScaleFactor))
                            End If

                            oBl.Update()
                        ElseIf TypeOf oEnt Is AcadTable Then
                            'Dim oTa As AcadTable = oEnt
                            'oTa.ScaleEntity(oTa.InsertionPoint, factor)
                        End If
                    Next
                End If
            End If
            ''
            oApp.ActiveDocument.SendCommand("_zoom" & vbCr & "e" & vbCr)
        End Sub
        ''
        <CommandMethod(regAPP, "IMPRIMIRINS", "IMPRIMIRINS", CommandFlags.Modal)> _
        Public Sub IMPRIMIRINS() ' This method can have any name
            ' Put your command code here
            ''
            If ImplacadActivado() = False Then Exit Sub
            If ImplacadEscalaM() = False Then Exit Sub
            ''
            If EsParaTrabajar() = False Then
                MsgBox("Este fichero DWG es sólo para imprimir" & vbCrLf & vbCrLf & _
                       "Debe abrir el DWG original de trabajo")
                Exit Sub
            End If
            '' Documento sin guardar
            If ImplacadGuardado() = False Then Exit Sub
            ''
            If oApp Is Nothing Then _
        oApp = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
            'oDoc = oApp.ActiveDocument
            '' 0.- Anular cualquier comando
            ''oApp.ActiveDocument.PostCommand(Chr(27) & Chr(27))
            ''
            ConfiguraDibujo()
            textoEstilo = Nothing
            Dim pt1 As Object = Nothing
            Dim pt2 As Object = Nothing

            Dim largo As Double = 0
            Dim ancho As Double = 0
            Dim resultado As String = ""

            Dim valorescala As Double = 1
            Try
                pt1 = oApp.ActiveDocument.Utility.GetPoint(, vbLf & "PRIMER punto de la ventana:")
                pt2 = oApp.ActiveDocument.Utility.GetCorner(pt1, vbLf & "SEGUNDO punto de la ventana:")
                largo = Math.Abs(pt2(0) - pt1(0))
                ancho = Math.Abs(pt2(1) - pt1(1))
                Dim mensaje As String = "Tamaño de salida"
                Dim arrop As String() = New String() {"A4", "A3", "A2", "A1", "A0"}
                resultado = DameOpcionTexto(mensaje, arrop)
            Catch ex As System.Exception
                Exit Sub
            End Try
            ''
            If resultado = "" Then
                Exit Sub
            End If
            ''
            Dim margen As Integer = 30
            Dim A4 As Array = New Double() {297 - margen, 210 - margen} '{297, 210}
            Dim A3 As Array = New Double() {420 - margen, 297 - margen} '{420, 297}
            Dim A2 As Array = New Double() {595 - margen, 420 - margen} '{595, 420}
            Dim A1 As Array = New Double() {840 - margen, 595 - margen} '{840, 595}
            Dim A0 As Array = New Double() {1190 - margen, 840 - margen} '{1190, 840}
            ''
            Select Case resultado.ToUpper
                Case "A4"
                    valorescala = DameEscala(A4, largo, ancho)
                Case "A3"
                    valorescala = DameEscala(A3, largo, ancho)
                Case "A2"
                    valorescala = DameEscala(A2, largo, ancho)
                Case "A1"
                    valorescala = DameEscala(A1, largo, ancho)
                Case "A0"
                    valorescala = DameEscala(A0, largo, ancho)
                Case Else
                    Exit Sub
            End Select
            ''
            If valorescala = 0 Then
                Exit Sub
            End If
            ''
            ''
            '' Guardamos el nuevo documento con el nombre codificado de la escala.
            '' Borrandolo antes, si existe.
            '' O abrimos el base
            Dim nuevoDocBase As String = IO.Path.GetDirectoryName(oApp.ActiveDocument.FullName) & "\" & IO.Path.GetFileNameWithoutExtension(oApp.ActiveDocument.FullName).Split("·")(0)
            Dim nuevoDocSuf As String = "·INS·" & resultado & ".dwg"
            Dim nuevoDoc As String = nuevoDocBase & nuevoDocSuf
            Dim viejoDoc As String = oApp.ActiveDocument.FullName
            ''
            ''
            '' Cerramos el dibujo si estuviera abierto en AutoCAD (Sin guardar. Ya que lo vamos a sobrescribir)
            ImplacadDibujoCierra(nuevoDoc)
            ''
            '' Borramos antes el fichero, si ya existía.
            If IO.File.Exists(nuevoDoc) Then IO.File.Delete(nuevoDoc)
            oApp.ActiveDocument.Save()
            ''
            CapaZonaCoberturaACTDES(CapaEstado.Desactivar)
            oApp.ActiveDocument.SaveAs(nuevoDoc)
            ''
            '' Abrimos también el dibujo anterior (viejoDoc), pero dejamos como activo el nuevo (nuevoDoc)
            Call oApp.Documents.Open(viejoDoc)
            ImplacadActivaDocumento(nuevoDoc, True, True)
            ''
            '' Escala todo el dibujo
            Dim cadenaejecuta As String = "_scale" & vbCr & "t" & vbCr & vbCr & pt1(0) & "," & pt1(1) & vbCr & valorescala & vbCr
            oApp.ActiveDocument.SendCommand(cadenaejecuta)
            ''
            ''
            'Dim escalafinal As Double = 1
            '(command "_.scale" nom-obj "" p-ins (/ 0.05 escala))
            vermensajes = False
            Try
                Me.TABLADATOS()
            Catch ex As System.Exception
                '' No hacemos nada. No existía la tabla
            End Try
            ''
            Try
                Me.TABLAESCALERAS()
            Catch ex As System.Exception
                '' No hacemos nada. No existía la tabla
            End Try
            Dim arrtodo As ArrayList = DameTodoImplacad()
            vermensajes = True
            ''
            'Dim mensaje As String = ""
            If arrtodo IsNot Nothing AndAlso arrtodo.Count > 0 Then
                For Each oEnt As AcadEntity In arrtodo
                    '' Si no es AcadBlockReference, saltar al siguiente
                    If Not (TypeOf oEnt Is AcadBlockReference) Then
                        Continue For
                    End If
                    ''
                    Dim texto As String = XData.XLeeDato(oEnt, xT.TEXTOS)
                    If texto <> "Clase=etiqueta" Then
                        Continue For
                    End If
                    texto = ""
                    ''
                    Dim oBloque As AcadBlockReference = oApp.ActiveDocument.ObjectIdToObject(oEnt.ObjectID)
                    'oBloque.XScaleFactor = escalaTotal
                    'oBloque.YScaleFactor = escalaTotal
                    'oBloque.ZScaleFactor = escalaTotal
                    '' 2*7*X=20
                    '' X=20/((2*7)
                    '' X=20/14
                    'If valorescala > 15 Then
                    '    oBloque.ScaleEntity(oBloque.InsertionPoint, 0.25)
                    'ElseIf valorescala > 10 Then
                    '    oBloque.ScaleEntity(oBloque.InsertionPoint, 0.5)
                    'ElseIf valorescala > 5 Then
                    '    oBloque.ScaleEntity(oBloque.InsertionPoint, 2)
                    'ElseIf valorescala > 1 Then
                    '    oBloque.ScaleEntity(oBloque.InsertionPoint, 4)
                    'Else
                    '    oBloque.ScaleEntity(oBloque.InsertionPoint, 10)
                    'End If
                    oBloque.ScaleEntity(oBloque.InsertionPoint, (60 / (2 * valorescala)))
                    'oBloque.XEffectiveScaleFactor = 0.05
                    'oBloque.YEffectiveScaleFactor = 0.05
                    'oBloque.ZEffectiveScaleFactor = 0.05
                    'oBloque.ScaleEntity(oBloque.InsertionPoint, (0.25 * valorescala))
                    oBloque.Update()
                    oBloque = Nothing
                Next
            End If
            ''
            '' Ponemos la propiedad personalizada ESTADO=PARAIMPRIMIR
            Try
                oApp.ActiveDocument.SummaryInfo.AddCustomInfo("ESTADO", "PARAIMPRIMIR")
            Catch ex As System.Exception
                oApp.ActiveDocument.SummaryInfo.SetCustomByKey("ESTADO", "PARAIMPRIMIR")
            End Try
            oApp.ActiveDocument.Regen(AcRegenType.acAllViewports)
            'oApp.ZoomAll()
            If oApp.ActiveDocument.Saved = False Then
                oApp.ActiveDocument.Save()
            End If
            oApp.ActiveDocument.Regen(AcRegenType.acAllViewports)
            ''
            '' Aviso al usuario
            MsgBox("Recuerde realizar las siguientes acciones:" & vbCrLf & vbCrLf & _
                   vbTab & "- Reposicionar las etiquetas escaladas..." & vbCrLf & _
                   vbTab & "- Al imprimir: Tamaño de papel (" & resultado.ToUpper & ")" & vbCrLf & _
                   vbTab & "- Al imprimir: Area de trazado (Seleccione Ventana)" & vbCrLf & _
                   vbTab & "- Al imprimir: Desfase de trazado (Centrado)" & vbCrLf & _
                   vbTab & "- Al imprimir: Escala de trazado (1:1)" & vbCrLf & _
                   "** Anote las opciones de impresión")
            ''
            ''
            For Each oEnt As AcadEntity In oApp.ActiveDocument.ModelSpace
                ''
                If oEnt.Layer = "Ruta evacuación primaria" Or oEnt.Layer = "Ruta evacuación accesibilidad" Then
                    '' Cambiamos la escala de linea
                    If TypeOf oEnt Is AcadLine Then
                        CType(oEnt, AcadLine).LinetypeScale = 5
                        CType(oEnt, AcadLine).Update()
                    End If
                End If
            Next
            ''
            oApp.ActiveDocument.SendCommand("_zoom" & vbCr & "e" & vbCr)
            oApp.ActiveDocument.Save()
        End Sub
        ''
        ''<CommandMethod(regAPP, "IMPRIMIREVA", "IMPRIMIREVA", CommandFlags.DocExclusiveLock)>
        <CommandMethod(regAPP, "IMPRIMIREVA", "IMPRIMIREVA", CommandFlags.Session)> _
        Public Sub IMPRIMIREVA() ' This method can have any name
            ' Put your command code here
            If oApp Is Nothing Then _
        oApp = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
            ''
            Dim oDocBase As AcadDocument = oApp.ActiveDocument
            Dim oDocEva As AcadDocument = Nothing
            ''
            If ImplacadActivado() = False Then Exit Sub
            If ImplacadEscalaM() = False Then Exit Sub
            If ImplacadActualizado() = False Then Exit Sub
            ''
            If EsParaTrabajar() = False Then
                MsgBox("Este fichero DWG es sólo para imprimir" & vbCrLf & vbCrLf & _
                       "Debe abrir el DWG original de trabajo")
                Exit Sub
            End If
            '' Documento sin guardar
            If ImplacadGuardado() = False Then Exit Sub
            ''
            If oDocBase.Saved = False Then oDocBase.Save()
            'oDoc = oApp.ActiveDocument
            '' 0.- Anular cualquier comando
            ''oApp.ActiveDocument.PostCommand(Chr(27) & Chr(27))
            ''
            'ConfiguraDibujo()
            textoEstilo = Nothing
            Dim pt1 As Object = Nothing
            Dim pt2 As Object = Nothing
            Dim ptmed(2) As Double
            Dim ptplantilla(2) As Double

            Dim largo As Double = 0
            Dim ancho As Double = 0
            Dim resultado As String = ""
            Dim tipo As String = ""         ' Horizontal o Vertical
            Dim prefijo As String = "EVACUACION_"
            Dim plantilla As String = ""        '' Nombre de la plantilla para plano evacuación
            Dim disX As Double
            Dim disY As Double

            Dim valorescala As Double = 1
            Dim hueco As Array = Nothing
            Try
                pt1 = oApp.ActiveDocument.Utility.GetPoint(, vbLf & "PRIMER punto de la ventana:")
                pt2 = oApp.ActiveDocument.Utility.GetCorner(pt1, vbLf & "SEGUNDO punto de la ventana:")
                '' LARGO
                If pt1(0) > pt2(0) Then
                    largo = Math.Abs(pt1(0) - pt2(0))
                Else
                    largo = Math.Abs(pt2(0) - pt1(0))
                End If
                '' ANCHO
                If pt1(1) > pt2(1) Then
                    ancho = Math.Abs(pt1(1) - pt2(1))
                Else
                    ancho = Math.Abs(pt2(1) - pt1(1))
                End If
                '' HORIZONTAL o VERTICAL
                If largo > ancho Then
                    tipo = "H"
                Else
                    tipo = "V"
                End If
                Dim mensaje As String = "Tamaño de salida"
                Dim arrop As String() = New String() {"A4", "A3", "A2"}
                resultado = DameOpcionTexto(mensaje, arrop)
                ''
                '' Punto de inserción de la plantilla
                ptmed(0) = (pt1(0) + pt2(0)) * 0.5
                ptmed(1) = (pt1(1) + pt2(1)) * 0.5
                ptmed(2) = 0
                ''
                Dim margen As Integer = 0
                'Dim A4 As Array = Nothing
                'Dim A3 As Array = Nothing
                'Dim A2 As Array = Nothing
                ''
                '' Hueco X e Y
                If tipo = "H" Then
                    Select Case resultado
                        Case "A4"
                            disX = 202
                            disY = 154
                            'A4 = New Double() {202 - margen, 161 - margen}      '' 297 x 210
                        Case "A3"
                            disX = 283
                            disY = 214
                            'A3 = New Double() {290 - margen, 229 - margen}      '' 420 x 297
                        Case "A2"
                            disX = 452
                            disY = 308
                            'A2 = New Double() {500 - margen, 330 - margen}      '' 595 x 420
                    End Select
                ElseIf tipo = "V" Then
                    Select Case resultado
                        Case "A4"
                            disX = 140
                            disY = 236
                            'A4 = New Double() {253 - margen, 140 - margen}
                        Case "A3"
                            disX = 191
                            disY = 350
                            'A3 = New Double() {357 - margen, 197 - margen}
                        Case "A2"
                            disX = 303
                            disY = 485
                            'A2 = New Double() {500 - margen, 310 - margen}
                    End Select
                End If
                hueco = New Double() {disX - margen, disY - margen}
                'disX = IIf(tipo = "H", IIf(resultado = "A4", 202, 290), IIf(resultado = "A4", 140, 197))
                'disY = IIf(tipo = "H", IIf(resultado = "A4", 161, 229), IIf(resultado = "A4", 253, 357))
                ''
            Catch ex As System.Exception
                Exit Sub
            End Try
            ''
            If resultado = "" Then
                Exit Sub
            End If
            ''
            '' Nombre de la plantilla para imprimir plano de evacuación
            plantilla = IO.Path.Combine(dirApp, prefijo & resultado & tipo & ".dwg")
            ''
            'Select Case resultado.ToUpper
            '    Case "A4"
            '        valorescala = DameEscala(A4, largo, ancho) '* 1.2
            '    Case "A3"
            '        valorescala = DameEscala(A3, largo, ancho) '* 1.2
            '    Case "A2"
            '        valorescala = DameEscala(A2, largo, ancho) '* 1.2
            '    Case Else
            '        Exit Sub
            'End Select
            valorescala = DameEscala(hueco, largo, ancho) '* 1.2
            ''
            If valorescala = 0 Then
                Exit Sub
            End If
            'ptmed(0) += (largo * valorescala) / 2
            'ptmed(1) += (ancho * valorescala) / 2
            ''
            Dim restarX As Double = 0
            Dim restarY As Double = 0
            If tipo = "H" Then
                Select Case resultado
                    Case "A4"
                        restarX = 184
                        restarY = 103
                    Case "A3"
                        restarX = 258
                        restarY = 146
                    Case "A2"
                        restarX = 336
                        restarY = 213
                End Select
            ElseIf tipo = "V" Then
                Select Case resultado
                    Case "A4"
                        restarX = 131
                        restarY = 148
                    Case "A3"
                        restarX = 187
                        restarY = 218
                    Case "A2"
                        restarX = 243
                        restarY = 305
                End Select
            End If
            ''
            ptplantilla(0) = ptmed(0) - restarX
            ptplantilla(1) = ptmed(1) - restarY ' Math.Abs(disY) + (IIf(tipo = "H", IIf(resultado = "A4", 25, 60), IIf(resultado = "A4", 40, 45)))
            ptplantilla(2) = 0
            ''
            ''
            '' Guardamos el nuevo documento con el nombre codificado de la escala.
            '' Borrandolo antes, si existe.
            '' O abrimos el base
            Dim nuevoDocBase As String = IO.Path.GetDirectoryName(oDocBase.FullName) & "\" & IO.Path.GetFileNameWithoutExtension(oDocBase.FullName).Split("·")(0)
            Dim nuevoDocSuf As String = "·EVA·" & resultado & tipo & ".dwg"
            Dim nuevoDoc As String = nuevoDocBase & nuevoDocSuf
            Dim viejoDoc As String = oDocBase.FullName  ' oApp.ActiveDocument.FullName
            ''
            ''
            '' Cerramos el dibujo si estuviera abierto en AutoCAD (Sin guardar. Ya que lo vamos a sobrescribir)
            ImplacadDibujoCierra(nuevoDoc)
            ''
            '' Borramos antes el fichero, si ya existía.
            If IO.File.Exists(nuevoDoc) Then IO.File.Delete(nuevoDoc)
            'CapaZonaCoberturaACTDES(CapaEstado.Desactivar)
            'oApp.ActiveDocument.Save()
            ''
            '' Cerramos el documento actual (viejoDoc) sin guardar.
            '' Copiamos viejoDoc en NuevoDoc y lo abrimos
            oDocBase.Close(False)
            oDocBase = Nothing
            IO.File.Copy(viejoDoc, nuevoDoc, True)
            oDocEva = oApp.Documents.Open(nuevoDoc)

            ''
            GrosorLineasACTDES(CapaEstado.Activar)
            CapaZonaCoberturaACTDES(CapaEstado.Desactivar)
            CapaCreaActivaBalizamientoEscalera()
            CapaCreaActivaBalizamientoPared()
            CapaCreaActivaRutaEvacuacion(TipoEvacuacion.Primaria)
            CapaCreaActivaRutaEvacuacion(TipoEvacuacion.Accesibilidad)
            CapaCreaActivaTablaDatos()
            ''
            '' Escala todo el dibujo
            Dim cadenaejecuta As String
            'cadenaejecuta = "_scale" & vbCr & "t" & vbCr & vbCr & Math.Min(pt1(0), pt2(0)) & "," & Math.Min(pt1(1), pt2(1)) & vbCr & valorescala & vbCr
            cadenaejecuta = "_scale" & vbCr & "t" & vbCr & vbCr & ptmed(0) & "," & ptmed(1) & vbCr & valorescala & vbCr
            ''
            oDocEva.SendCommand(cadenaejecuta)
            ''
            oDocEva.Regen(AcRegenType.acAllViewports)
            'oApp.ZoomAll()
            ''
            ''
            'Dim escalafinal As Double = 1
            '(command "_.scale" nom-obj "" p-ins (/ 0.05 escala))
            vermensajes = False
            'Me.TABLADATOS()
            'Me.TABLAESCALERAS()
            'Dim arrtodo As ArrayList = DameTodoImplacad()
            vermensajes = True
            ''
            'Dim mensaje As String = ""
            For Each oEnt As AcadEntity In oDocEva.ModelSpace
                ''
                '' Borramos todo lo que tengamos puesto en las siguientes capas
                If oEnt.Layer = "capacuadro" Or _
                    oEnt.Layer = "BALIZAMIENTO ESCALERA" Or _
                    oEnt.Layer = "BALIZAMIENTO PARED" Or _
                    oEnt.Layer = "BALIZAMIENTO SUELO" Then
                    oEnt.Delete()
                    'Continue For
                ElseIf oEnt.Layer = "Ruta evacuación primaria" Or oEnt.Layer = "Ruta evacuación accesibilidad" Then
                    '' Cambiamos la escala de linea
                    If TypeOf oEnt Is AcadLine Then
                        CType(oEnt, AcadLine).LinetypeScale = 5
                        CType(oEnt, AcadLine).Update()
                    End If
                    'Continue For
                End If
            Next
            ''
            '' Escalamos los bloques de etiquetas
            Dim arrtodo As ArrayList = DameTodoImplacad()
            If arrtodo IsNot Nothing AndAlso arrtodo.Count > 0 Then
                For Each oEnt As AcadEntity In arrtodo
                    '' Si no es AcadBlockReference, saltar al siguiente
                    If Not (TypeOf oEnt Is AcadBlockReference) Then
                        Continue For
                    End If
                    ''
                    Dim texto As String = XData.XLeeDato(oEnt, xT.TEXTOS)
                    If texto <> "Clase=etiqueta" Then
                        Continue For
                    End If
                    texto = ""
                    ''
                    Dim oBloque As AcadBlockReference = oDocEva.ObjectIdToObject(oEnt.ObjectID)
                    oBloque.ScaleEntity(oBloque.InsertionPoint, (50 / (2 * valorescala)))
                    'oBloque.ScaleEntity(oBloque.InsertionPoint, (60 / (2 * valorescala)))
                    'oBloque.XEffectiveScaleFactor = 0.05
                    'oBloque.YEffectiveScaleFactor = 0.05
                    'oBloque.ZEffectiveScaleFactor = 0.05
                    'oBloque.ScaleEntity(oBloque.InsertionPoint, (0.25 * valorescala))
                    oBloque.Update()
                    oBloque = Nothing
                Next
            End If
            ''
            '' Ponemos la propiedad personalizada ESTADO=PARAIMPRIMIR
            Try
                oDocEva.SummaryInfo.AddCustomInfo("ESTADO", "PARAIMPRIMIR")
            Catch ex As System.Exception
                oDocEva.SummaryInfo.SetCustomByKey("ESTADO", "PARAIMPRIMIR")
            End Try
            ''
            '' Insertamos la plantilla elegida.
            'Dim cadenainserta As String = "-insert" & vbCr & plantilla & vbCr & _
            'ptplantilla(0) + (IIf(tipo = "H", IIf(resultado = "A4", 43, 65), IIf(resultado = "A4", 53, 57))) & "," & _
            'ptplantilla(1) + (IIf(tipo = "H", IIf(resultado = "A4", 35, 30), IIf(resultado = "A4", 50, 100))) & vbCr & _
            '"1" & vbCr & "1" & vbCr & "0" & vbCr
            Dim cadenainserta As String = "-insert" & vbCr & plantilla & vbCr & _
                    ptplantilla(0) & "," & _
                    ptplantilla(1) & vbCr & _
                    "1" & vbCr & "1" & vbCr & "0" & vbCr
            oDocEva.SendCommand(cadenainserta)
            ''
            'Dim oblCaj As AcadBlockReference = oApp.ActiveDocument.ModelSpace.InsertBlock(ptplantilla, plantilla, 1, 1, 1, 0)
            'Dim oblCaj As AcadBlockReference = DameBloque(IO.Path.GetFileNameWithoutExtension(plantilla))
            'If oblCaj Is Nothing Then
            '    Dim oent1 As AcadEntity = UltimoObjeto()
            '    If TypeOf oent1 Is AcadBlockReference Then
            '        If CType(oent1, AcadBlockReference).Name = IO.Path.GetFileNameWithoutExtension(plantilla) Then
            '            oblCaj = UltimoObjeto()
            '        End If
            '    End If
            'End If
            ''
            oDocEva.Regen(AcRegenType.acAllViewports)
            oApp.Update()
            'oApp.ActiveDocument.Save()
            ''
            'While Application.IsQuiescent = False
            ''
            'System.Threading.Thread.Sleep(0)
            'End While
            'oApp.ActiveDocument.Close(True)
            ''
            '' Lo volvemos a abrir
            'oApp.Documents.Open(nuevoDoc)
            ''
            'Dim oblCaj As AcadBlockReference = DameBloque(IO.Path.GetFileNameWithoutExtension(plantilla))
            'If oblCaj IsNot Nothing Then
            'oblCaj.Explode()
            'Else
            'oApp.ActiveDocument.SendCommand("_explode" & vbCr & "l" & vbCr & vbCr)
            'End If
            'oApp.Update()
            ''
            'oApp.ActiveDocument.ModelSpace.XRefDatabase
            ''
            'CType(oApp.ActiveDocument.ModelSpace.Item(oApp.ActiveDocument.ModelSpace.Count - 1), AcadBlockReference).Explode()
            'CType(UltimoObjeto(), AcadBlockReference).Explode()
            'Dim oPlantilla As AcadBlockReference = oApp.ActiveDocument.ModelSpace.InsertBlock(ptplantilla, plantilla, 1, 1, 1, 0)
            'oPlantilla.ScaleEntity(oPlantilla.InsertionPoint, 1)
            'oPlantilla.XScaleFactor = 1
            ' oPlantilla.YScaleFactor = 1
            'oPlantilla.ZScaleFactor = 1
            'oPlantilla.Update()
            ''
            'oApp.ActiveDocument.Regen(AcRegenType.acAllViewports)
            oApp.ZoomAll()
            'If oApp.ActiveDocument.Saved = False Then
            'oApp.ActiveDocument.Save()
            'End If
            'oApp.ActiveDocument.Regen(AcRegenType.acAllViewports)
            'oApp.ActiveDocument.Regen(AcRegenType.acAllViewports)
            ConfiguraDibujo()
            'oApp.ActiveDocument.SendCommand("_zoom" & vbCr & "e" & vbCr)
            'MsgBox(avisofinal)
            ''
            '' Cambiar todas las etiquetas de evacuacion y extincion por los bloques correspondientes
            ' Dim oent As AcadEntity = UltimoObjeto()
            'If TypeOf oent Is AcadBlockReference Then
            If arrtodo IsNot Nothing AndAlso arrtodo.Count > 0 Then
                CambiaBloquePorImagen(oDocEva)
            End If
            'End If
            ''
            'BotonesImplacad(ActivarBotones.Cualquiera, "EXPLOTAEVA", True)
            oApp.Update()
            'oDocBase.Regen(AcRegenType.acAllViewports)
            ''
            oDocEva.SendCommand("_zoom" & vbCr & "e" & vbCr)
            'oApp.ZoomAll()
            'oApp.ActiveDocument.SendCommand("_zoom" & vbCr & "_all" & vbCr)
            'oDocBase.SaveAs(nuevoDoc)
            'ImplacadDibujoCierra(viejoDoc)
            'oDocBase = Nothing
            'ImplacadActivaDocumento(nuevoDoc, True, False)
            'oDocEva = oApp.ActiveDocument
            'oApp.ActiveDocument.Save()
            ''
            '' Abrimos también el dibujo anterior (viejoDoc), pero dejamos como activo el nuevo (nuevoDoc)
            Call oApp.Documents.Open(viejoDoc)
            ImplacadActivaDocumento(nuevoDoc, True, True)
            'ImplacadGuardaDocumentos()
            If Not (oDocBase Is Nothing) Then System.Runtime.InteropServices.Marshal.ReleaseComObject(oDocBase)
            oDocBase = Nothing
            If Not (oDocEva Is Nothing) Then System.Runtime.InteropServices.Marshal.ReleaseComObject(oDocEva)
            oDocEva = Nothing
            GC.Collect()
            GC.WaitForPendingFinalizers()
            GC.Collect()

            '' Aviso al usuario
            MsgBox("Recuerde realizar las siguientes acciones:" & vbCrLf & vbCrLf & _
                   vbTab & "- Reposicionar las etiquetas escaladas..." & vbCrLf & _
                   vbTab & "- Mover bloque marco para centrar plano..." & vbCrLf & _
                   vbTab & "- Insertar bloque USTEDAQUI en 'posición actual' en plano" & vbCrLf & _
                   vbTab & "- Al imprimir: Tamaño de papel (" & resultado.ToUpper & ")" & vbCrLf & _
                   vbTab & "- Al imprimir: Area de trazado (Seleccione Ventana)" & vbCrLf & _
                   vbTab & "- Al imprimir: Desfase de trazado (Centrado)" & vbCrLf & _
                   vbTab & "- Al imprimir: Escala de trazado (1:1)" & vbCrLf & vbCrLf & _
                   "** Anote las opciones de impresión")
            ''
        End Sub
        ''
        <CommandMethod(regAPP, "EXPLOTAEVA", "EXPLOTAEVA", CommandFlags.DocExclusiveLock)> _
        Public Sub EXPLOTAEVA() ' This method can have any name
            ' Put your command code here
            ''
            If ImplacadActivado() = False Then Exit Sub
            If ImplacadEscalaM() = False Then Exit Sub
            ''
            If oApp Is Nothing Then _
oApp = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
            ''
            '' Documento sin guardar
            If ImplacadGuardado() = False Then Exit Sub
            ''
            '' Ver si es un plano de evacuación y termina en: "·EVA·A4H.dwg", "·EVA·A4V.dwg", "·EVA·A3H.dwg", "·EVA·A3V.dwg"
            ''
            'If oApp.ActiveDocument.Name.Contains("·EVA·A4H.dwg") = False Or _
            '    oApp.ActiveDocument.Name.Contains("·EVA·A4H.dwg") = False Or _
            '    oApp.ActiveDocument.Name.Contains("·EVA·A4H.dwg") = False Or _
            '    oApp.ActiveDocument.Name.Contains("·EVA·A4H.dwg") = False Then
            '    MsgBox("Este plano no es de EVACUACIÓN y no se puede usar esta herramienta...")
            '    Exit Sub
            'End If

            Dim blEva As AcadBlockReference = Nothing
            Dim nEva As String() = New String() {"EVACUACION_A4H", "EVACUACION_A4V", "EVACUACION_A3H", "EVACUACION_A3V"}
            Dim mensaje As String = ""
            For Each nombreEva As String In nEva
                blEva = DameBloque(nombreEva)
                If blEva IsNot Nothing Then
                    mensaje &= nombreEva & vbCrLf
                    blEva.Explode()
                    blEva.Delete()
                End If
            Next
            ''
            If mensaje <> "" Then
                MsgBox("Encontrados y explotados los bloques :" & vbCrLf & vbCrLf & mensaje & vbCrLf & vbCrLf & "** Ya puede editarlos...")
            Else
                MsgBox("No se ha encontrado ninguno de estos bloques :" & vbCrLf & vbCrLf & String.Join(vbCrLf, nEva))
            End If
        End Sub
        ''
        <CommandMethod(regAPP, "ACTUALIZARIMPLACAD", "ACTUALIZARIMPLACAD", CommandFlags.Session)> _
        Public Sub ACTUALIZARIMPLACAD() ' This method can have any name
            '' Documento sin guardar
            If oApp.ActiveDocument.FullName = "" Or (oApp.ActiveDocument.FullName <> "" And oApp.ActiveDocument.Saved = False) Then
                MsgBox("Este documento aún no se ha guardado" & vbCrLf & vbCrLf & _
                       "Guarde antes el documento y vuelva a probar...")
                Exit Sub
            End If
            ''
            '' Comprobar si está activada la aplicación
            '' Introducir código para crear el fichero clave IMPLACAD.IMP,
            '' que contendrá el valor de codigoactivacion
            ''
            If ImplacadActivado(False) = False Then
                Dim resultado As String = InputBox("Codigo VIP de activación --> ", "ACTIVAR " & nApp & " · " & My.Application.Info.Version.ToString)
                If resultado.ToUpper = codigoactivacion.ToUpper Then
                    IO.File.WriteAllText(nImp, codigoactivacion)
                    MsgBox(nApp & " · " & My.Application.Info.Version.ToString & _
                           " ha sido activada. Ahora podrá elegir si quiere comprobar actualizaciones", MsgBoxStyle.Exclamation)
                Else
                    MsgBox("Código VIP incorrecto. Contacte con IMPLASER para solicitarlo...", MsgBoxStyle.Exclamation, "AVISOS AL USUARIO")
                    Exit Sub
                End If
            End If
            ''
            Dim respuesta As Microsoft.VisualBasic.MsgBoxResult
            respuesta = MsgBox("¿Desea ahora comprobar si existen actualizaciones?" & vbCrLf & vbCrLf & _
                               "Se guardarán y cerrarán todos los ficheros abiertos.", _
                               MsgBoxStyle.YesNo)
            If respuesta = MsgBoxResult.No Then
                respuesta = Nothing
                Exit Sub
            End If
            respuesta = Nothing
            ''
            If oApp Is Nothing Then _
        oApp = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
            '' Guardar y cerrar todos los dibujos abiertos.
            CierraDibujoTodos()
            ''oApp.Update()     '' Da error si no hay dibujos abiertos.
            ''
            '' Descargar DATOS (Fichero .zip que hay que descomprimir)
            Dim mensaje As String = ""
            Dim origenWeb As String = nApp & "DATOS.zip"
            Dim destinoHD As String = IO.Path.Combine(IMPLACAD_DATA, origenWeb)
            Try
                If IO.Directory.Exists(IMPLACAD_DATA) = False Then
                    IO.Directory.CreateDirectory(IMPLACAD_DATA)
                    Call PermisosTodoCarpeta(IMPLACAD_DATA)
                End If
            Catch ex As System.Exception
                MsgBox("No se puede crear el directorio " & IMPLACAD_DATA & vbCrLf & vbCrLf &
                       "Verifique si dispone de permisos para crearlo en C:\ProgramData" & vbCrLf &
                       "O creelo a mano y vuelva a probar...")
                Exit sub
            End Try
            ''
            Try
                mensaje &= DescargaFicheroZIPDescomprime(origenWeb, destinoHD, "Descargando " & origenWeb) & vbCrLf & vbCrLf
            Catch ex As System.Exception
                mensaje &= "Error actualizando " & origenWeb & vbCrLf & ex.Message & vbCrLf & vbCrLf
            End Try
            ''
            '' Descargar BD (Fichero .sdf que hay que actualizar)
            ''origenWeb = nApp & ".sdf"
            ''
            '' Descargar XLSX (Fichero .xlsx que hay que actualizar)
            origenWeb = nApp & ".xlsx"
            destinoHD = IO.Path.Combine(dirApp, origenWeb)
            Try
                mensaje &= DescargaFicheroZIPDescomprime(origenWeb, destinoHD, "Descargando " & origenWeb)
            Catch ex As System.Exception
                mensaje &= "Error actualizando " & origenWeb & vbCrLf & ex.Message & vbCrLf & vbCrLf
            End Try
            Application.ShowAlertDialog(mensaje)
            ''
        End Sub
        ''
        Public Function DescargaFicheroZIPDescomprime(queFiWeb As String,
                                                      queFiLocal As String, _
                                                      Optional quePrefijo As String = "Descargando ") As String
            If quePrefijo.EndsWith(" ") = False Then quePrefijo &= " "
            ''
            If oApp Is Nothing Then _
        oApp = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
            ''
            '' 1.- Cerrar los formularios abiertos.
            If frmE IsNot Nothing Then
                frmE.Close()
                frmE = Nothing
            End If
            ''
            If My.Computer.Network.IsAvailable = False Then
                'Application.ShowAlertDialog("No dispone de conexión a internet. No se puede actualizar.")
                Return queFiWeb & " = " & "No dispone de conexión a internet. No se puede actualizar." & vbCrLf
                Exit Function
            End If
            ''
            '' 2.- Comprobar si existe el directorio con los recursos
            If IO.Directory.Exists(IMPLACAD_DATA) = False Then
                'Application.ShowAlertDialog("No se puede actualizar..." & vbCrLf & vbCrLf & "No existe--> " & dirBase)
                Return queFiWeb & " = " & "No se puede actualizar... No existe--> " & IMPLACAD_DATA & vbCrLf
                Exit Function
            End If
            ''
            'Dim origenWebbak As String = nApp & "DATOS.bak"
            'Dim origenWeb As String = nApp & "DATOS.zip"
            'Dim destinoHDbak As String = dirBase & origenWebbak
            'Dim destinoHD As String = dirBase & origenWeb
            ''
            Dim docurl = webActualiza & queFiWeb
            Dim queFiLocalBak As String = IO.Path.ChangeExtension(queFiLocal, ".bak")
            ''
            '' Para comprobar información del fichero a descargar (tamaño, fecha y todos los datos)
            'MsgBox(DameTamañoFicheroWEB(docurl))
            'MsgBox(DameFechaFicheroWEB(docurl))
            'MsgBox(DameDatosFicheroWEB(docurl))
            ''
            '' 3.- Comprobar si existe el fichero .zip en la Web.
            '' Si existe cogemos su fecha de ultima modificación. Si no existe salir sin actualizar.
            Dim fechaWeb As Date = DameFechaFicheroWEB(docurl, False)   '' True para ver todos las cabeceras y sus datos
            Dim fechaLocal As Date = New System.DateTime(0)
            Dim tamWeb As Integer = DameTamañoFicheroWEB(docurl)
            Dim tamLocal As Integer
            If fechaWeb = New System.DateTime(0) Then
                Return False
                Exit Function
            End If
            '' Comprobar si existe el fichero .bak a descargar en el disco duro y borrarlo.
            If IO.File.Exists(queFiLocalBak) Then IO.File.Delete(queFiLocalBak)
            ''
            ''
            '' 4.- Comprobar si existe el fichero Local comprimido con los recursos
            If IO.File.Exists(queFiLocal) Then
                '' Si existe 
                fechaLocal = New IO.FileInfo(queFiLocal).LastWriteTime.ToLocalTime
                tamLocal = New IO.FileInfo(queFiLocal).Length
            End If
            ''
            '' 5.- Si fechaWeb > fechaLocal, descargar y actualizar.
            '' Si no salir sin actualizar.
            If fechaWeb <= fechaLocal Then
                'Application.ShowAlertDialog("No existen nuevas actualizaciones para " & queFiWeb)
                Return queFiWeb & " = " & "No existen nuevas actualizaciones" & vbCrLf
                Exit Function
            End If
            ''
            '' 6.- Descargar fichero Web con la extensión .bak primero
            '' Utilizamos un formulario para ver el proceso de descarga.
            '' Poner un ProgressBar para ir rellenándolo
            Dim frmD As New frmDescargar
            frmD.prefijo = quePrefijo
            frmD.StartPosition = Windows.Forms.FormStartPosition.CenterScreen
            Application.ShowModelessDialog(Application.MainWindow.Handle, frmD, True)
            DescargaFicheroWEB(docurl, queFiLocalBak, frmD.pb1)
            frmD.Close() : frmD.Dispose()
            frmD = Nothing
            ''
            '' 7.- Borramos el anterior si existía
            If IO.File.Exists(queFiLocal) Then
                '' Si existe lo borramos antes y después lo descargamos.
                IO.File.Delete(queFiLocal)
            End If
            ''
            '' 8.- Renombramos la extensión de .bak a .zip
            FileSystem.Rename(queFiLocalBak, queFiLocal)  ' IO.Path.ChangeExtension(destinoHDbak, ".zip")
            ''
            '' 9.- Ahora descomprimimos el fichero, si era un fichero .zip
            If queFiLocal.ToLower.EndsWith(".zip") Then DescomprimeZIP(queFiLocal)
            ''
            'Application.ShowAlertDialog("Datos actualizados...")
            Return queFiWeb & " = " & "Datos actualizados..." & vbCrLf
        End Function
        ''
        ' Modal Command with pickfirst selection
        <CommandMethod(regAPP, "TABLAPARCIAL", "TABLAPARCIAL", CommandFlags.Modal + CommandFlags.UsePickSet)>
        Public Sub TABLAPARCIAL() ' This method can have any name
            ''
            If ImplacadActivado() = False Then Exit Sub
            If ImplacadEscalaM() = False Then Exit Sub
            ''
            ''
            If EsParaTrabajar() = False Then
                MsgBox("Este fichero DWG es sólo para imprimir" & vbCrLf & vbCrLf &
                       "Debe abrir el DWG original de trabajo")
                Exit Sub
            End If
            If oApp Is Nothing Then _
       oApp = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
            'oDoc = oApp.ActiveDocument
            ' Put your command code here
            '' 0.- Anular cualquier comando
            ''oApp.ActiveDocument.PostCommand(Chr(27) & Chr(27))
            CapaCeroActiva()
            Dim estadocapa As Boolean = True
            Try
                estadocapa = oApp.ActiveDocument.Layers.Item("Zona Cobertura Evacuación").LayerOn
            Catch ex As System.Exception
                ''
            End Try
            CapaZonaCoberturaACTDES(CapaEstado.Desactivar)      '' Crear y desactivar las 2 zonas de cobertura.
            ''
            Try
                Dim result As PromptSelectionResult = Application.DocumentManager.MdiActiveDocument.Editor.GetSelection()
                If (result.Status = PromptStatus.OK) Then
                    ' There are selected entities
                    ' Put your command using pickfirst set code here
                    'MsgBox("TABLAPARCIAL" & vbCrLf & vbCrLf & "Seleccionados : " & result.Value.Count)
                    TablaDatosParcialInserta()
                Else
                    ' There are no selected entities
                    ' Put your command code here
                    'MsgBox("TABLAPARCIAL" & vbCrLf & vbCrLf & "No hay entidades seleccionadas...")
                End If
            Catch ex As System.Exception

            End Try
            ''
            If estadocapa = True Then
                CapaZonaCoberturaACTDES(CapaEstado.Activar)
            End If
        End Sub

        ' Modal Command with pickfirst selection
        <CommandMethod(regAPP, "RESOLVERREFX", "RESOLVERREFX", CommandFlags.Modal + CommandFlags.Session)>
        Public Sub RESOLVERREFX() ' This method can have any name
            ''
            If ImplacadActivado() = False Then Exit Sub
            If ImplacadEscalaM() = False Then Exit Sub
            ''
            ''
            If EsParaTrabajar() = False Then
                MsgBox("Este fichero DWG es sólo para imprimir" & vbCrLf & vbCrLf &
                       "Debe abrir el DWG original de trabajo")
                Exit Sub
            End If
            If oApp Is Nothing Then _
       oApp = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
            XRef_DWGListar(False)
            XRef_IMGListar(False)
        End Sub

        ' Application Session Command with localized name
        <CommandMethod(regAPP, "MySessionCmd", "MySessionCmdLocal", CommandFlags.Modal + CommandFlags.Session)> _
        Public Sub MySessionCmd() ' This method can have any name
            ' Put your command code here
            ''
            If ImplacadActivado() = False Then Exit Sub
            If ImplacadEscalaM() = False Then Exit Sub
            ''
            If EsParaTrabajar() = False Then
                MsgBox("Este fichero DWG es sólo para imprimir" & vbCrLf & vbCrLf & _
                       "Debe abrir el DWG original de trabajo")
                Exit Sub
            End If
            If oApp Is Nothing Then _
       oApp = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
            'oDoc = oApp.ActiveDocument
            ' Put your command code here
            '' 0.- Anular cualquier comando
            ''oApp.ActiveDocument.PostCommand(Chr(27) & Chr(27))
            ''
        End Sub

        ' LispFunction is similar to CommandMethod but it creates a lisp 
        ' callable function. Many return types are supported not just string
        ' or integer.
        <LispFunction("MyLispFunction", "MyLispFunctionLocal")> _
        Public Function MyLispFunction(ByVal args As ResultBuffer) ' This method can have any name
            ' Put your command code here
            ''
            If ImplacadActivado() = False Then
                Return 0
                Exit Function
            End If
            If ImplacadEscalaM() = False Then
                Return 0
                Exit Function
            End If
            ''
            If EsParaTrabajar() = False Then
                MsgBox("Este fichero DWG es sólo para imprimir" & vbCrLf & vbCrLf & _
                       "Debe abrir el DWG original de trabajo")
                Return 0
                Exit Function
            End If
            'If oApp Is Nothing Then _
            'oApp = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
            ' Put your command code here
            '' 0.- Anular cualquier comando
            'oApp.ActiveDocument.PostCommand(Chr(27) & Chr(27))
            ''
            ' Return a value to the AutoCAD Lisp Interpreter
            Return 1
        End Function

        Private Sub oTimer_Elapsed(sender As Object, e As Timers.ElapsedEventArgs) Handles oTimer.Elapsed
            CierraDibujo(nombreviejo)
            'oTimer.Enabled = False
            'oDoc = oApp.ActiveDocument
        End Sub
    End Class

End Namespace