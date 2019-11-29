Imports System.Diagnostics
Imports System.Collections
Imports System.Windows.Forms
Imports System.Drawing
Imports System.Windows.FrameworkCompatibilityPreferences

Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.Interop
Imports Autodesk.AutoCAD.Interop.Common
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.EditorInput
Imports AXApp = Autodesk.AutoCAD.ApplicationServices.Application
Imports AXDoc = Autodesk.AutoCAD.ApplicationServices.Document
Imports AXWin = Autodesk.AutoCAD.Windows

Namespace A2acad
    Partial Public Class A2acad
#Region ".NET"
        Public Function GetPointDelUsuario(ByVal strMensaje As String) As Object
            Dim doc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim db As Database = doc.Database
            Dim ed As Editor = doc.Editor
            Try
                Dim ops As PromptPointOptions = New PromptPointOptions(strMensaje)
                Dim selRes As PromptPointResult = ed.GetPoint(ops)
                If selRes.Status <> PromptStatus.OK Then
                    ed.WriteMessage(vbLf & "Selected Point FAILED.")
                    Return Nothing
                Else
                    ed.WriteMessage(vbLf & "Selected Point CORRECT.")
                    Return selRes.Value
                End If
            Catch ex As Exception
                ed.WriteMessage(ex.ToString())
                Return Nothing
            End Try
        End Function

        Public Function PuntoDameGet_NET(mensaje As String) As Object
            Dim resultado As Object = Nothing
            '' Get the current database and start the Transaction Manager
            Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim ed As Editor = acDoc.Editor
            Dim acCurDb As Database = acDoc.Database
            Dim pPtRes As PromptPointResult
            Dim pPtOpts As PromptPointOptions = New PromptPointOptions("")
            '' Prompt for the start point
            pPtOpts.Message = vbLf & mensaje
            pPtOpts.AllowNone = True    ' Permite terminar con boton derecho.
            Using oUi As EditorUserInteraction = ed.StartUserInteraction(oAppAintP) ' oAppS.Application.MainWindow.Handle)
                ' Using (acDoc.LockDocument) '' Necesario para dialogos modales.
                '' Poner el foco en la zona de dibujo
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                pPtRes = acDoc.Editor.GetPoint(pPtOpts)
                Dim ptStart As Autodesk.AutoCAD.Geometry.Point3d = pPtRes.Value

                '' Exit if the user presses ESC or cancels the command
                If pPtRes.Status = PromptStatus.Cancel Then
                    resultado = Nothing
                ElseIf pPtRes.Status = PromptStatus.OK Then
                    resultado = ptStart
                End If
                ''
                oUi.End()
            End Using
            acDoc = Nothing
            acCurDb = Nothing
            Return resultado
        End Function

        Public Function OpcionesTextoDameGet_NET(listaopciones As String, frase As String) As String
            Dim resultado As String = ""
            ''
            If oAppA Is Nothing Then _
        oAppA = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
            'oDoc = oAppA.ActiveDocument
            oAppA.ActiveDocument.Activate()
            ''

            ' Define the list of valid keywords
            'Dim listaopciones As String
            'listaopciones = "Ascendente Descendente"
            oAppA.ActiveDocument.Utility.InitializeUserInput(1, listaopciones)

            ' Prompt para que el usuario introduzca una palabra. Return "Ascendente", "Descendente" en
            ' resultado variable o puede introducir "A", "D".
            'Dim returnString As String
            resultado = oAppA.ActiveDocument.Utility.GetKeyword(vbLf & frase)   ' "Tipo de Escalera [Ascendente/Descendente] <Ascendente>: ")
            'MsgBox("You entered " & resultado, , "GetKeyword Example")
            If resultado = "" Then resultado = listaopciones.Split(" "c)(0)
            Return resultado
        End Function

        Public Function OpcionesTextoDameGet_NET(queopciones As String(), fraseSinOpciones As String) As String
            Dim resultado As String = ""
            ''
            If oAppA Is Nothing Then _
        oAppA = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
            'oDoc = oAppA.ActiveDocument
            oAppA.ActiveDocument.Activate()
            ''

            ' Define the list of valid keywords
            Dim listaopciones As String = String.Join(" ", queopciones)
            'listaopciones = "Primaria Secundaria"
            oAppA.ActiveDocument.Utility.InitializeUserInput(0, listaopciones)   '' 0 para poder contestar con return.

            ' Prompt para que el usuario introduzca una palabra del array. queopciones(0) sera el valor por defecto.
            Dim cadenaopciones As String = String.Join("/", queopciones)
            Dim cadenadefecto As String = queopciones(0)
            resultado = oAppA.ActiveDocument.Utility.GetKeyword(vbLf & fraseSinOpciones & " [" & cadenaopciones & "] <" & cadenadefecto & ">: ")
            'MsgBox("You entered " & resultado, , "GetKeyword Example")
            If resultado = "" Then
                resultado = cadenadefecto
            ElseIf InStr(listaopciones, resultado) = 0 Then
                resultado = ""
            End If
            ''
            Return resultado
        End Function
        Public Function OpcionesTextoDINDame_NET() As String
            Dim resultado As String = ""
            ''
            If oAppA Is Nothing Then _
        oAppA = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
            'oDoc = oAppA.ActiveDocument
            oAppA.ActiveDocument.Activate()
            '' Define the list of valid keywords
            Dim listaopciones, pordefecto As String
            listaopciones = "A4 A3 A2 A1 A0"
            pordefecto = "A4"
            'oDoc.Utility.InitializeUserInput(1, listaopciones)
            oAppA.ActiveDocument.Utility.InitializeUserInput(7, listaopciones)
            ' Prompt para que el usuario introduzca una palabra. Return "Ascendente", "Descendente" en
            ' resultado variable o puede introducir "A", "D".
            resultado = oAppA.ActiveDocument.Utility.GetKeyword(vbLf & "Formato DIN [A4/A3/A2/A1/A0] <A4>: ")
            'MsgBox("You entered " & resultado, , "GetKeyword Example")
            If resultado = "" Then resultado = pordefecto
            Return resultado
        End Function
#End Region
    End Class
End Namespace