Imports System
Imports System.Text
Imports System.Linq
Imports System.Xml
Imports System.Reflection
Imports System.ComponentModel
Imports System.Collections
Imports System.Collections.Generic
Imports System.Windows
Imports System.Windows.Media.Imaging
Imports System.Windows.Forms
Imports System.IO

Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Interop
Imports Autodesk.AutoCAD.Interop.Common
Imports Autodesk.AutoCAD.Geometry
Imports AXApp = Autodesk.AutoCAD.ApplicationServices.Application
Imports AXDoc = Autodesk.AutoCAD.ApplicationServices.Document
Imports AXWin = Autodesk.AutoCAD.Windows


'
''' <summary>
''' Inicializar los eventos que necesitemos (Poner esto al inicio de AddIn. Primera linea que se ejecuta)
''' Eventos.Eventos_Inicializa()
''' </summary>
Partial Public Class Eventos

    ' ***** OBJETOS AUTOCAD
    '
    ' ***** Los tipos de objetos que vamos a controlar con eventos
    Public Shared lTypesAXObj As String() = {"BlockReference", "Circle", "MLeader"}
    Public Shared lTypesCOMObj As String() = {"AcDbCircle", "AcDbBlockReference", "AcadMLeader"}
    Public Shared lIds As List(Of ObjectId)     ' Lista de objetos modificados
    '
    ' ***** Variables
    Public Shared ultimoObjectId As ObjectId = Nothing
    Public Shared ultimoAXDoc As String = ""
    Public Shared FicheroLog As String = ""
    ' ***** FLAG
    Public Const coneventos As Boolean = True
    Public Const logeventos As Boolean = False
    Public Shared SYSMON As Object = 0
    '
    Public Shared Sub Eventos_Inicializa()
        FicheroLog = System.IO.Path.Combine(System.IO.Path.GetTempPath, "logeventos.csv")
        If logeventos Then
            File.WriteAllText(FicheroLog, "EVENTOS EN ORDEN;DATOS" & vbCrLf, Encoding.UTF8)
        Else
            If File.Exists(FicheroLog) Then
                Try
                    File.Delete(FicheroLog)
                Catch ex As System.Exception
                End Try
            End If
        End If
        lIds = New List(Of ObjectId)
        'If AXDocM.Count > 0 Then
        '    Subscribe_AXDoc()
        'End If
        'Subscribe_AXEventM()
        'Subscribe_AXApp()
        'Subscribe_COMApp()
        'Subscribe_AXDocM()
    End Sub
    Public Shared Sub Eventos_Vacia()
        Try
            lIds = Nothing
            'If AXDocM.Count > 0 Then
            '    Unsubscribe_AXDoc()
            'End If
            'Unsubscribe_AXEventM()
            'Unsubscribe_AXApp()
            'Unsubscribe_COMApp()
            'Unsubscribe_AXDocM()
        Catch ex As System.Exception
            Debug.Print(ex.ToString)
        End Try
    End Sub

    Public Shared Sub PonLogEv(queTexto As String)
        If queTexto.EndsWith(vbCrLf) = False Then queTexto &= vbCrLf
        File.AppendAllText(FicheroLog, queTexto, Encoding.UTF8)
        'PonLogEvAsync(queTexto)  ' Falla al dejar abierto el fichero
    End Sub
    'Public Shared Async Sub PonLogEvAsync(quetexto As String)
    '    Exit Sub  ' Anulado hasta resolver problema con cerrar fichero.
    '    Dim uniencoding As UnicodeEncoding = New UnicodeEncoding()
    '    Dim filename As String = FicheroLog

    '    Dim result As Byte() = uniencoding.GetBytes(quetexto)

    '    Using SourceStream As FileStream = File.Open(filename, FileMode.Append)
    '        SourceStream.Seek(0, SeekOrigin.End)
    '        Await SourceStream.WriteAsync(result, 0, result.Length)
    '        SourceStream.Close()
    '        SourceStream.Dispose()
    '    End Using
    'End Sub

    Public Shared Sub AcadBlockReference_PonEventosModified()
        'Dim AcadBlockReference As ArrayList = clsA.SeleccionaDameBloquesTODOS(regAPPA)
        'For Each oBl As AcadBlockReference In AcadBlockReference
        '    Dim queTipo As String = clsA.XLeeDato(oBl, "tipo")
        '    If queTipo = "cinta" Then
        '        AddHandler oBl.Modified, AddressOf modTavil.AcadBlockReference_Modified
        '    End If
        'Next
    End Sub
    '
    Public Shared Function AXEventM() As Autodesk.AutoCAD.Internal.Reactors.ApplicationEventManager
        Return Autodesk.AutoCAD.Internal.Reactors.ApplicationEventManager.Instance
    End Function

    Public Shared Function AXDocM() As Autodesk.AutoCAD.ApplicationServices.DocumentCollection
        Return Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager
    End Function

    Public Shared Function AXDoc() As Autodesk.AutoCAD.ApplicationServices.Document
        Return AXDocM.MdiActiveDocument
    End Function
    Public Shared Function AXDb() As Autodesk.AutoCAD.DatabaseServices.Database
        Return AXDoc.Database
    End Function
    Public Shared Function AXEditor() As Autodesk.AutoCAD.EditorInput.Editor
        Return AXDoc.Editor
    End Function
    Public Shared Function COMApp() As Autodesk.AutoCAD.Interop.AcadApplication
        Return CType(AXApp.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication).Application
    End Function
    Public Shared Function COMDoc() As Autodesk.AutoCAD.Interop.AcadDocument
        'Return COMApp.ActiveDocument
        Return AXDocM.MdiActiveDocument.GetAcadDocument
    End Function
    Public Shared Function COMDb() As Autodesk.AutoCAD.Interop.Common.AcadDatabase
        Return COMDoc.Database
    End Function
    '
    Public Shared Sub AcadPopupMenuItem_PonerQuitar(ByRef ShortcutMenu As AcadPopupMenu, comando As String, queLabel As String, poner As Boolean)
        'AXDoc.Editor.WriteMessage("AcadPopupMenuItem_PonerQuitar")
        If logeventos Then PonLogEv("AcadPopupMenuItem_PonerQuitar")
        ' Poner el AcadPopupMenuItem (Tiene que existir el comando TAVILACERCADE)
        Dim acercadeMacro As String = Chr(27) + Chr(27) + Chr(95) + IIf(comando.EndsWith(" "), comando, comando & " ")  ' "TAVILACERCADE "
        'Dim queLabel As String = "Tavil. Acerca de..."
        Try
            Dim encontrado As Boolean = False
            For x As Integer = 0 To ShortcutMenu.Count - 1
                Dim oItem As AcadPopupMenuItem = ShortcutMenu.Item(x)
                If oItem.Label = queLabel Then
                    encontrado = True
                    If poner = False Then
                        oItem.Delete()
                        Exit Sub
                    End If
                    Exit For
                End If
            Next
            If encontrado = False Then
                Dim newItem As AcadPopupMenuItem = ShortcutMenu.AddMenuItem(ShortcutMenu.Count, queLabel, acercadeMacro)
            End If
        Catch ex As System.Exception

        End Try
    End Sub
    Public Shared Sub SYSMONVAR(Optional guardar As Boolean = False)
        If guardar Then
            SYSMON = AXApp.GetSystemVariable("SYSMON")
            AXApp.SetSystemVariable("SYSMON", 0)
        Else
            AXApp.SetSystemVariable("SYSMON", SYSMON)
        End If
    End Sub
    Public Shared Sub SYSMONVARCOM(Optional guardar As Boolean = False)
        If guardar Then
            SYSMON = Convert.ToInt32(Eventos.COMDoc.GetVariable("SYSMON"))
            Eventos.COMDoc.SetVariable("SYSMON", 0)
        Else
            Eventos.COMDoc.SetVariable("SYSMON", SYSMON)
        End If
    End Sub
End Class