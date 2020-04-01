Imports Autodesk.AutoCAD.Interop.Common

Module modBloques

    ' Nos devuelve el valor de un atributo (Busca en Atributos y en Constant Atributos)
    Public Function Bloque_AtributoDame(oBl As AcadBlockReference, nombreAtri As String) As String
        Dim resultado As String = ""
        ''
        If oApp Is Nothing Then _
            oApp = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
        ' Cargamos los atributos
        Dim varAttributes As Object
        varAttributes = oBl.GetAttributes
        Dim consAttributes As Object
        consAttributes = oBl.GetConstantAttributes

        ' Localizar los atributos y añadir valores a la colección (ORDEN y NUMERO). O todos
        '' Atributos editables.
        Dim oAtri As Object ' Autodesk.AutoCAD.Interop.Common.AcadAttribute
        For I As Integer = LBound(varAttributes) To UBound(varAttributes)
            oAtri = varAttributes(I)
            If oAtri.TagString.ToString.ToUpper = nombreAtri.ToUpper Then
                resultado = oAtri.TextString
                Exit For
            End If
        Next
        ' Si no lo hemos encontrado. Busca en Constantes.
        If resultado = "" Then
            '' Atributos constantes.
            For I As Integer = LBound(consAttributes) To UBound(consAttributes)
                oAtri = consAttributes(I)
                If oAtri.TagString.ToString.ToUpper = nombreAtri.ToUpper Then
                    resultado = oAtri.TextString
                    Exit For
                End If
            Next
        End If
        ''
        VaciaMemoria()
        oAtri = Nothing
        oBl = Nothing
        ''
        Return resultado
    End Function

    '
    Public Function Bloque_AtributoDame(queId As Long, nombreAtri As String) As String
        Dim resultado As String = ""
        ''
        If oApp Is Nothing Then _
            oApp = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
        ''
        Dim oBl As Autodesk.AutoCAD.Interop.Common.AcadBlockReference = oApp.ActiveDocument.ObjectIdToObject(queId)
        resultado = Bloque_AtributoDame(oBl, nombreAtri)
        oBl = Nothing
        Return resultado
    End Function

    '
    Public Function Bloque_AtributoDame(queHandle As String, nombreAtri As String) As String
        Dim resultado As String = ""
        ''
        If oApp Is Nothing Then _
            oApp = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
        ''
        Dim oBl As Autodesk.AutoCAD.Interop.Common.AcadBlockReference = oApp.ActiveDocument.HandleToObject(queHandle)
        resultado = Bloque_AtributoDame(oBl, nombreAtri)
        oBl = Nothing
        Return resultado
    End Function

    Public Function Bloque_AtributoDame(arrAtributos As Object, nombreAtri As String, Optional arrAtributosConstantes As Object = Nothing) As String
        Dim resultado As String = ""
        '
        ' Localizar los atributos y añadir valores a la colección (ORDEN y NUMERO). O todos
        '' Atributos editables.
        Dim oAtri As Object ' Autodesk.AutoCAD.Interop.Common.AcadAttribute
        For I As Integer = LBound(arrAtributos) To UBound(arrAtributos)
            oAtri = arrAtributos(I)
            If oAtri.TagString.ToString.ToUpper = nombreAtri.ToUpper Then
                resultado = oAtri.TextString
                Exit For
            End If
        Next
        If resultado = "" AndAlso arrAtributosConstantes IsNot Nothing Then
            ' Atributos constantes.
            For I As Integer = LBound(arrAtributosConstantes) To UBound(arrAtributosConstantes)
                oAtri = arrAtributosConstantes(I)
                If oAtri.TagString.ToString.ToUpper = nombreAtri.ToUpper Then
                    resultado = oAtri.TextString
                    Exit For
                End If
            Next
        End If
        ''
        VaciaMemoria()
        oAtri = Nothing
        ''
        Return resultado
    End Function

    Public Sub Bloque_AtributoEscribe(OBl As AcadBlockReference, nombreAtri As String, valorAtri As String)
        ''
        If oApp Is Nothing Then _
        oApp = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
        ' Cargamos los atributos
        Dim varAttributes As Object
        varAttributes = OBl.GetAttributes
        ''
        '' Atributos editables.
        Dim oAtri As Object ' Object ' Autodesk.AutoCAD.Interop.Common.AcadAttribute
        For I As Integer = LBound(varAttributes) To UBound(varAttributes)
            oAtri = varAttributes(I)
            If oAtri.TagString.ToString.ToUpper = nombreAtri.ToUpper Then
                oAtri.TextString = valorAtri
                Exit For
            End If
        Next
        OBl.Update()
        ''
        OBl = Nothing
        varAttributes = Nothing
        oAtri = Nothing
        ''
        VaciaMemoria()
    End Sub

    Public Sub Bloque_AtributoEscribe(queId As Long, nombreAtri As String, valorAtri As String)
        If oApp Is Nothing Then _
        oApp = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
        Dim oBl As Autodesk.AutoCAD.Interop.Common.AcadBlockReference = oApp.ActiveDocument.ObjectIdToObject(queId)
        Bloque_AtributoEscribe(oBl, nombreAtri, valorAtri)
        oBl = Nothing
    End Sub

    Public Sub Bloque_AtributoEscribe(queHandle As String, nombreAtri As String, valorAtri As String)
        If oApp Is Nothing Then _
        oApp = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
        Dim oBl As Autodesk.AutoCAD.Interop.Common.AcadBlockReference = oApp.ActiveDocument.HandleToObject(queHandle)
        Bloque_AtributoEscribe(oBl, nombreAtri, valorAtri)
        oBl = Nothing
    End Sub

    Public Sub Bloque_AtributoEscribe(oBl As AcadBlockReference, dicNombreValor As Dictionary(Of String, String))
        ''
        If oApp Is Nothing Then _
        oApp = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
        ' Cargamos los atributos
        Dim varAttributes As Object
        varAttributes = oBl.GetAttributes
        ''
        '' Atributos editables.
        Dim oAtri As Object ' Object ' Autodesk.AutoCAD.Interop.Common.AcadAttribute
        For I As Integer = LBound(varAttributes) To UBound(varAttributes)
            oAtri = varAttributes(I)
            If dicNombreValor.ContainsKey(oAtri.TagString.ToString.ToUpper) Then
                If oAtri.TextString.ToString <> dicNombreValor(oAtri.TagString.ToString.ToUpper) Then
                    oAtri.TextString = dicNombreValor(oAtri.TagString.ToString.ToUpper)
                End If
            End If
        Next
        oBl.Update()
        ''
        oBl = Nothing
        varAttributes = Nothing
        oAtri = Nothing
        ''
        VaciaMemoria()
    End Sub

    Public Sub Bloque_AtributoEscribe(queId As Long, dicNombreValor As Dictionary(Of String, String))
        ''
        If oApp Is Nothing Then _
        oApp = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
        ''
        Dim oBl As Autodesk.AutoCAD.Interop.Common.AcadBlockReference = oApp.ActiveDocument.ObjectIdToObject(queId)
        Bloque_AtributoEscribe(oBl, dicNombreValor)
        oBl = Nothing
    End Sub

    Public Sub Bloque_AtributoEscribe(queHandle As String, dicNombreValor As Dictionary(Of String, String))
        ''
        If oApp Is Nothing Then _
        oApp = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
        ''
        Dim oBl As Autodesk.AutoCAD.Interop.Common.AcadBlockReference = oApp.ActiveDocument.HandleToObject(queHandle)
        Bloque_AtributoEscribe(oBl, dicNombreValor)
        oBl = Nothing
    End Sub

    Public Function Bloque_SeleccionaDame(Optional conconfirmacion As Boolean = False) As AcadBlockReference
        ' Begin the selection
        Dim bloque As AcadBlockReference = Nothing
        Dim obj As AcadObject = Nothing
        Dim basePnt As Object = Nothing

        ' Dim doc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        'Using acLckDoc As DocumentLock = doc.LockDocument()
        ' Foco en AutoCAD
        Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()

        ' The following example waits for a selection from the user
        On Error Resume Next
RETRY:
        oApp.ActiveDocument.Utility.GetEntity(obj, basePnt, vbCrLf & "Seleccione Bloque : ")
        'MessageBox.Show(obj.ObjectName)

        If Err.Number <> 0 Then
            Err.Clear()
            'MsgBox("Program ended.", , "GetEntity Example")
            VaciaMemoria()
            Return bloque
            Exit Function
        End If
        '
        If obj.ObjectName = "AcDbBlockReference" Then
            If conconfirmacion = True Then
                Dim resultado As String = ""
                resultado = oApp.ActiveDocument.Utility.GetString(False, vbLf & "[Intro] o 'S' Acepta Selección / 'N' anula selección --> : ")
                If resultado.ToUpper = "N" Then
                    obj = Nothing
                    GoTo RETRY
                End If
            End If
            bloque = CType(oApp.ActiveDocument.HandleToObject(obj.Handle), AcadBlockReference)
            bloque.Update()
        Else
            obj = Nothing
            basePnt = Nothing
            GoTo RETRY
        End If
        'End Using
        'doc = Nothing
        VaciaMemoria()
        Return bloque
    End Function

    Public Function Bloques_DameNombreContiene(quenombre As String, Optional exacto As Boolean = False) As List(Of AcadBlockReference)
        Dim resultado As New List(Of AcadBlockReference)
        If oApp Is Nothing Then _
        oApp = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
        ''
        For Each oEnt As AcadEntity In oApp.ActiveDocument.ModelSpace
            If TypeOf oEnt Is AcadBlockReference Then
                Dim oBl As AcadBlockReference = oApp.ActiveDocument.ObjectIdToObject(oEnt.ObjectID)
                If exacto = True Then
                    If oBl.EffectiveName.ToUpper = quenombre.ToUpper OrElse oBl.Name.ToUpper = quenombre.ToUpper Then
                        resultado.Add(oBl)
                    End If
                Else
                    If oBl.EffectiveName.ToUpper.Contains(quenombre.ToUpper) OrElse oBl.Name.ToUpper.Contains(quenombre.ToUpper) Then
                        resultado.Add(oBl)
                    End If
                End If
                oBl = Nothing
            End If
        Next
        VaciaMemoria()
        Return resultado
    End Function

    Public Function Bloques_DameNombreContiene(quenombres() As String) As List(Of AcadBlockReference)
        Dim resultado As New List(Of AcadBlockReference)
        If oApp Is Nothing Then _
        oApp = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
        '
        For Each oEnt As AcadEntity In oApp.ActiveDocument.ModelSpace
            If TypeOf oEnt Is AcadBlockReference Then
                Dim oBl As AcadBlockReference = oApp.ActiveDocument.ObjectIdToObject(oEnt.ObjectID)
                For Each nombre As String In quenombres
                    If oBl.EffectiveName.ToUpper.Contains(nombre.ToUpper) OrElse oBl.Name.ToUpper.Contains(nombre.ToUpper) Then
                        resultado.Add(oBl)
                    End If
                    oBl = Nothing
                Next
            End If
            System.Windows.Forms.Application.DoEvents()
        Next
        VaciaMemoria()
        Return resultado
    End Function

End Module