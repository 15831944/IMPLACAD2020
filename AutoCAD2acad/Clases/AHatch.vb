Imports System.Diagnostics
Imports System.Collections
Imports System.Windows.Forms
Imports System.Drawing


Imports Autodesk.AutoCAD.Interop
Imports Autodesk.AutoCAD.Interop.Common
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.EditorInput

Namespace A2acad
    Partial Public Class A2acad

        Public Sub Sombrea(objoH As AcadHatch,
                       Optional ByVal obj As AcadObject = Nothing,
                       Optional ByVal coleccion As ArrayList = Nothing,
                       Optional ByVal quitar As Boolean = False)
            'If oAppA.GetAcadState.IsQuiescent = False Then
            Me.LiberaApp(False)
            SombreadoBorra(objoH)
            '
            If quitar = True Then
                Me.oDoc.Regen(AcRegenType.acActiveViewport)
                Exit Sub
            End If

            If quitar = False And (Not (obj Is Nothing) Or Not (coleccion Is Nothing)) Then
                'oDoc.Regen(AcRegenType.acActiveViewport)
                'Else
                oH = oDoc.ModelSpace.AddHatch(1, "SOLID", False)    '1, "SOLID", True
                'If ultZona <> "" Then oH.Layer = modVariables.preCapa & ultZona  'OJO la capa aún no existe
                'oH = oDoc.ModelSpace.AddHatch(1, "SOLID", True)
                oH.PatternDouble = False
                oH.PatternAngle = 45
                oH.PatternSpace = 100
                oH.HatchStyle = AcHatchStyle.acHatchStyleNormal
                oH.AssociativeHatch = False


                OC.ColorIndex = AcColor.acYellow
                oH.TrueColor = OC
                Dim OE(0) As AcadEntity
                Try
                    If coleccion Is Nothing Then
                        OE(0) = obj
                        oH.AppendOuterLoop(OE)
                        oH.Evaluate()
                    Else
                        'oH.Delete()
                        For Each pol As AcadLWPolyline In coleccion
                            'Dim pol As AcadLWPolyline = CType(coleccion(x), AcadLWPolyline)
                            Debug.Print(pol.Handle)
                            OE(0) = pol 'coleccion(x)
                            oH.AppendOuterLoop(OE)
                        Next
                        oH.Evaluate()
                        Me.oDoc.SendCommand(Chr(27))
                    End If
                    'oH.Update
                    'Me.ActivaApp(, False)
                Catch ex As Exception
                    If Not (oH Is Nothing) Then
                        oH.Delete()
                        oH = Nothing
                    End If
                End Try
            End If
            'Me.oDoc.ActiveViewport.GridOn = Not (Me.oDoc.ActiveViewport.GridOn)
            'Me.oDoc.ActiveViewport.GridOn = Not (Me.oDoc.ActiveViewport.GridOn)
            Me.oDoc.Regen(AcRegenType.acActiveViewport)
        End Sub
        Public Sub SombreadoBorra(Optional queoH As AcadHatch = Nothing)
            If queoH IsNot Nothing Then
                Try
                    queoH.Delete()
                    queoH = Nothing
                Catch ex As Exception
                    Console.WriteLine(ex.Message.ToString)
                End Try
            End If
        End Sub

        '' El primer objeto del ArrayList es el sombreado.
        Public Function SombreadoDameContornos() As ArrayList
            Dim resultado As New ArrayList
            ''
repetir:
            Dim objCadEnt As AcadEntity = Nothing
            Dim vrRetPnt As Object = Nothing
            Try
                oAppA.ActiveDocument.Utility.GetEntity(objCadEnt, vrRetPnt, "Seleccione Sombreado")
                '' Si no seleccionamos nada, salimos
                If objCadEnt Is Nothing Then
                    Return resultado
                    Exit Function
                ElseIf Not (TypeOf objCadEnt Is AcadHatch) Then
                    '' Volver a solicitar entidad
                    GoTo repetir
                End If
            Catch ex As System.Exception
                Return resultado
                Exit Function
            End Try
            '' Si el objeto seleccionado es un sombreado
            If TypeOf (objCadEnt) Is AcadHatch Then   '|-- Chequear si es un sombreado (Hatch) --|
                Dim objHatch As AcadHatch = objCadEnt
                '' Añadimos el sombreado (0) El resto serán contornos.
                resultado.Add(objHatch)
                ''
                '' Iteramos con los Loops que tenga
                For x As Integer = 0 To objHatch.NumberOfLoops - 1
                    Dim loopObjs As Object = Nothing
                    objHatch.GetLoopAt(x, loopObjs)
                    ' Iteramos con los objetos dentre de cada Loop
                    If loopObjs Is Nothing Then Continue For
                    For I = LBound(loopObjs) To UBound(loopObjs)
                        resultado.Add(loopObjs(I))
                    Next
                Next
                ''
            End If
            ''
            oAppA.ActiveDocument.SetVariable("pickadd", 2)   '' Restaurar varias selecciones.
            Return resultado
        End Function

        Public Function SombreadoDameContornos(queSom As AcadHatch) As ArrayList
            If oAppA Is Nothing Then _
        oAppA = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
            '' ArrayList de AcadEntity con todos los contornos.
            Dim resultado As New ArrayList
            For x As Integer = 0 To queSom.NumberOfLoops - 1
                ' Find the objects that make up the first loop
                Dim loopObjs As Object = Nothing
                queSom.GetLoopAt(x, loopObjs)

                ' Find the types of the objects in the loop
                Dim I As Integer
                For I = LBound(loopObjs) To UBound(loopObjs)
                    resultado.Add(CType(loopObjs(I), AcadEntity))
                Next
            Next
            ''
            Return resultado
        End Function
    End Class
End Namespace