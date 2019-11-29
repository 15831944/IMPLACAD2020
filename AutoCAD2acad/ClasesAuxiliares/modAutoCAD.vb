Imports System
Imports System.Drawing
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.ApplicationServices
Imports acApp = Autodesk.AutoCAD.ApplicationServices.Application
Imports Autodesk.AutoCAD.Interop
Imports Autodesk.AutoCAD.Interop.Common
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Ribbon
Imports Autodesk.Windows
Imports Autodesk.AutoCAD.Customization


Module modAutoCAD
    ''
    Public WithEvents oTimer As System.Timers.Timer
    'Public Const atributoNUMERO As String = "NUMERO"
    'Public arrPreBloques As ArrayList   '' Prefijo de los bloques usados en el desarrollo (Para escalarlos)
    ''
    '' Configurar Dibujo Actual
    ''
    '    ''

    Public Function SeleccionaDameBloqueXData(Optional ByVal nombre As Object = "", Optional ByVal capa As Object = "", Optional ByVal DatosX As Boolean = True, Optional ByVal SelTemp As Boolean = False) As AcadBlockReference
        Dim resultado As AcadBlockReference = Nothing
        'Dim cSeleccion As AcadSelectionSets
        Dim F1(0) As Short
        Dim F2(0) As Object
        Dim vF1 As Object = Nothing
        Dim vF2 As Object = Nothing

        '' Las 2 maneras valen igual. AcDbBlckReference es mejor (Solo coge bloques) INSERT coge sombreados también.
        F1(0) = 100 : F2(0) = "AcDbBlockReference"
        'F1(0) = 0 : F2(0) = "INSERT"
        'DatosX Siempre tiene que estar despues de entidad. Si no no funciona
        If DatosX = True Then
            ReDim Preserve F1(F1.Length)
            ReDim Preserve F2(F2.Length)
            'Filtro de bloque, nombre y capa
            F1(F1.Length - 1) = 1001 : F2(F2.Length - 1) = regAPP
        End If
        If nombre <> "" Then
            ReDim Preserve F1(F1.Length)
            ReDim Preserve F2(F2.Length)
            'Filtro de bloque y nombre
            F1(F1.Length - 1) = 2 : F2(F2.Length - 1) = nombre
        End If
        If capa <> "" Then
            ReDim Preserve F1(F1.Length)
            ReDim Preserve F2(F2.Length)
            'Filtro de bloque, nombre y capa
            F1(F1.Length - 1) = 8 : F2(F1.Length - 1) = capa
        End If

        vF1 = F1
        vF2 = F2
        ''
        If oApp Is Nothing Then _
        oApp = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
        ''
        ''
        Try
            oSel = oApp.ActiveDocument.SelectionSets.Add(regAPP)
        Catch ex As System.Exception
            oSel = oApp.ActiveDocument.SelectionSets.Item(regAPP)
        End Try
        ''
        Try
            oSelTemp = oApp.ActiveDocument.SelectionSets.Add("TEMPORAL")
        Catch ex As System.Exception
            oSelTemp = oApp.ActiveDocument.SelectionSets.Item("TEMPORAL")
        End Try
        ''
        ''
        oApp.ActiveDocument.SetVariable("pickadd", 0)   '' Solo una selección. Se quita lo que hubiera
        If SelTemp = False Then
            oSel.Clear()
            Try
                oSel.SelectOnScreen(vF1, vF2)
            Catch ex As System.Exception
                'Debug.Print(ex.Message)
            End Try
            'MsgBox(oSel.Count)
            If oSel.Count > 0 Then
                If TypeOf oSel.Item(oSel.Count - 1) Is AcadBlockReference Then
                    resultado = oSel.Item(oSel.Count - 1)
                Else
                    resultado = Nothing
                End If
            Else
                resultado = Nothing
            End If
        Else
            oSelTemp.Clear()
            Try
                oSelTemp.Select(vF1, vF2)
            Catch ex As System.Exception
                'Debug.Print(ex.Message)
            End Try
            'MsgBox(oSel.Count)
            If oSelTemp.Count > 0 Then
                'Dim texto As String = XData.XLeeDato(oEnt, xT.TEXTOS)
                'If TypeOf oEnt Is AcadTable Or texto = "Clase=tabla" Then
                If TypeOf oSelTemp.Item(oSelTemp.Count - 1) Is AcadTable Then
                    resultado = Nothing
                Else
                    resultado = oSelTemp.Item(oSelTemp.Count - 1)
                End If
            Else
                resultado = Nothing
            End If
        End If
        ''
        oApp.ActiveDocument.SetVariable("pickadd", 2)   '' La seleccion actual se suma a la que hubiera.
        Return resultado
        Exit Function
    End Function
    ''
    Public Function SeleccionaDameBloqueUno(Optional ByVal nombre As Object = "", Optional ByVal capa As Object = "") As AcadBlockReference
        Dim resultado As AcadBlockReference = Nothing
        'Dim cSeleccion As AcadSelectionSets
        Dim F1(1) As Short
        Dim F2(1) As Object
        Dim vF1 As Object = Nothing
        Dim vF2 As Object = Nothing

        '' Las 2 maneras valen igual. AcDbBlckReference es mejor (Solo coge bloques) INSERT coge sombreados también.
        F1(0) = 100 : F2(0) = "AcDbBlockReference"
        F1(1) = 0 : F2(1) = "INSERT"
        'DatosX Siempre tiene que estar despues de entidad. Si no no funciona
        'If DatosX = True Then
        '    ReDim Preserve F1(F1.Length)
        '    ReDim Preserve F2(F2.Length)
        '    'Filtro de bloque, nombre y capa
        '    F1(F1.Length - 1) = 1001 : F2(F2.Length - 1) = regAPP
        'End If
        If nombre <> "" Then
            ReDim Preserve F1(F1.Length)
            ReDim Preserve F2(F2.Length)
            'Filtro de bloque y nombre
            F1(F1.Length - 1) = 2 : F2(F2.Length - 1) = nombre
        End If
        If capa <> "" Then
            ReDim Preserve F1(F1.Length)
            ReDim Preserve F2(F2.Length)
            'Filtro de bloque, nombre y capa
            F1(F1.Length - 1) = 8 : F2(F1.Length - 1) = capa
        End If

        vF1 = F1
        vF2 = F2
        ''
        If oApp Is Nothing Then _
        oApp = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
        ''
        ''
        Try
            oSel = oApp.ActiveDocument.SelectionSets.Add(regAPP)
        Catch ex As System.Exception
            oSel = oApp.ActiveDocument.SelectionSets.Item(regAPP)
        End Try
        ''
        oApp.ActiveDocument.SetVariable("pickadd", 0)   '' Solo una selección. Se quita lo que hubiera
        oSel.Clear()
        Try
            oSel.Select(vF1, vF2)
        Catch ex As System.Exception
            'Debug.Print(ex.Message)
        End Try
        'MsgBox(oSel.Count)
        If oSel.Count > 0 Then
            'Dim texto As String = XData.XLeeDato(oEnt, xT.TEXTOS)
            'If TypeOf oEnt Is AcadTable Or texto = "Clase=tabla" Then
            If TypeOf oSel.Item(oSel.Count - 1) Is AcadTable Then
                resultado = Nothing
            Else
                resultado = oSel.Item(oSel.Count - 1)
            End If
        Else
            resultado = Nothing
        End If
        ''
        oApp.ActiveDocument.SetVariable("pickadd", 2)   '' La seleccion actual se suma a la que hubiera.
        Return resultado
        Exit Function
    End Function
    ''
    Public Function SeleccionaDameBloquesVarios() As ArrayList
        Dim resultado As New ArrayList
        'Dim cSeleccion As AcadSelectionSets
        Dim F1(1) As Short
        Dim F2(1) As Object
        Dim vF1 As Object = Nothing
        Dim vF2 As Object = Nothing

        '' Las 2 maneras valen igual. AcDbBlckReference es mejor (Solo coge bloques) INSERT coge sombreados también.
        F1(0) = 100 : F2(0) = "AcDbBlockReference"
        F1(1) = 0 : F2(1) = "INSERT"
        vF1 = F1
        vF2 = F2
        ''
        If oApp Is Nothing Then _
        oApp = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
        ''
        ''
        Try
            oSel = oApp.ActiveDocument.SelectionSets.Add(regAPP)
        Catch ex As System.Exception
            oSel = oApp.ActiveDocument.SelectionSets.Item(regAPP)
        End Try
        ''
        oSel.Clear()
        ''
        oApp.ActiveDocument.SetVariable("pickadd", 2)   '' Solo una selección. Se quita lo que hubiera
        ''
        Try
            oSel.SelectOnScreen(vF1, vF2)
        Catch ex As System.Exception
            'Debug.Print(ex.Message)
        End Try
        'MsgBox(oSel.Count)
        If oSel.Count > 0 Then
            For Each oEnt As AcadEntity In oSel
                If TypeOf oEnt Is AcadBlockReference Then
                    resultado.Add(CType(oEnt, AcadBlockReference))
                End If
            Next
        Else
            resultado = Nothing
        End If
        ''
        oApp.ActiveDocument.SetVariable("pickadd", 2)   '' La seleccion actual se suma a la que hubiera.
        Return resultado
        Exit Function
    End Function
    ''
    Public Function SeleccionaDamePolilineasVarias() As ArrayList
        Dim resultado As New ArrayList
        'Dim cSeleccion As AcadSelectionSets
        Dim F1(1) As Short
        Dim F2(1) As Object
        Dim vF1 As Object = Nothing
        Dim vF2 As Object = Nothing

        '' Las 2 maneras valen igual. AcDbBlckReference es mejor (Solo coge bloques) INSERT coge sombreados también.
        F1(0) = 100 : F2(0) = "AcDbPolyline"
        F1(1) = 0 : F2(1) = "LWPOLYLINE"
        vF1 = F1
        vF2 = F2
        ''
        If oApp Is Nothing Then _
        oApp = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
        ''
        ''
        Try
            oSel = oApp.ActiveDocument.SelectionSets.Add(regAPP)
        Catch ex As System.Exception
            oSel = oApp.ActiveDocument.SelectionSets.Item(regAPP)
        End Try
        ''
        oSel.Clear()
        ''
        oApp.ActiveDocument.SetVariable("pickadd", 2)   '' Solo una selección. Se quita lo que hubiera
        ''
        Try
            oSel.SelectOnScreen(vF1, vF2)
        Catch ex As System.Exception
            'Debug.Print(ex.Message)
        End Try
        'MsgBox(oSel.Count)
        If oSel.Count > 0 Then
            For Each oEnt As AcadEntity In oSel
                If TypeOf oEnt Is AcadLWPolyline Then
                    resultado.Add(CType(oEnt, AcadLWPolyline))
                End If
            Next
        Else
            resultado = Nothing
        End If
        ''
        oApp.ActiveDocument.SetVariable("pickadd", 2)   '' La seleccion actual se suma a la que hubiera.
        Return resultado
        Exit Function
    End Function
    ''
    Public Function SeleccionaDamePolilineasTODAS(Optional queCapas As String = "*") As ArrayList
        Dim resultado As New ArrayList
        'Dim cSeleccion As AcadSelectionSets
        Dim F1(2) As Short
        Dim F2(2) As Object
        Dim vF1 As Object = Nothing
        Dim vF2 As Object = Nothing

        '' Las 2 maneras valen igual. AcDbBlckReference es mejor (Solo coge bloques) INSERT coge sombreados también.
        F1(0) = 100 : F2(0) = "AcDbPolyline"
        F1(1) = 0 : F2(1) = "*POLYLINE*"    '"LWPOLYLINE"
        F1(2) = 8 : F2(2) = queCapas
        vF1 = F1
        vF2 = F2
        ''
        If oApp Is Nothing Then _
        oApp = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
        ''
        ''
        Try
            oSel = oApp.ActiveDocument.SelectionSets.Add(regAPP)
        Catch ex As System.Exception
            oSel = oApp.ActiveDocument.SelectionSets.Item(regAPP)
        End Try
        ''
        oSel.Clear()
        ''
        oApp.ActiveDocument.SetVariable("pickadd", 2)   '' Solo una selección. Se quita lo que hubiera
        ''
        Try
            oSel.Select(AcSelect.acSelectionSetAll,,, vF1, vF2)
        Catch ex As System.Exception
            'Debug.Print(ex.Message)
        End Try
        'MsgBox(oSel.Count)
        If oSel.Count > 0 Then
            For Each oEnt As AcadEntity In oSel
                If TypeOf oEnt Is AcadLWPolyline Then
                    resultado.Add(CType(oEnt, AcadLWPolyline))
                    'resultado.Add(CType(oEnt, AcadLWPolyline))
                End If
            Next
        Else
            resultado = Nothing
        End If
        ''
        oApp.ActiveDocument.SetVariable("pickadd", 2)   '' La seleccion actual se suma a la que hubiera.
        Return resultado
        Exit Function
    End Function
    ''
    Public Function SeleccionaDameObjetosEnCapa(queCapa As String) As ArrayList
        Dim resultado As New ArrayList
        ''
        'Add bit about counting text on layer
        Dim myEd As Editor = acApp.DocumentManager.MdiActiveDocument.Editor
        Dim myTVs(0) As TypedValue
        myTVs.SetValue(New TypedValue(DxfCode.LayerName, queCapa), 0)
        Dim myFilter As New SelectionFilter(myTVs)
        Dim myPSR As PromptSelectionResult = myEd.SelectAll(myFilter)
        Dim oSel As SelectionSet = Nothing
        If myPSR.Status = PromptStatus.OK Then
            oSel = myPSR.Value
        End If
        '    Dim myTVs(3) As TypedValue
        'myTVs.SetValue(New TypedValue(DxfCode.Operator, "<AND"), 0)
        'myTVs.SetValue(New TypedValue(DxfCode.Start, "TEXT"), 1)
        'myTVs.SetValue(New TypedValue(DxfCode.LayerName, "0"), 2)
        'myTVs.SetValue(New TypedValue(DxfCode.Operator, "AND>"), 3)
        ''myTVs(0) = New TypedValue(DxfCode.l .Start, "MTEXT")
        'Dim myFilter As New SelectionFilter(myTVs)
        '    Dim myPSR As PromptSelectionResult = myEd.SelectAll(myFilter)
        '    If myPSR.Status = PromptStatus.OK Then
        '    Dim mySS As SelectionSet = myPSR.Value
        'myForm.Label2.Text = mySS.Count
        'End If
        ''
        If oApp Is Nothing Then _
            oApp = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
        ''
        ''
        'Try
        'oSel = oApp.ActiveDocument.SelectionSets.Add(regAPP)
        'Catch ex As System.Exception
        'oSel = oApp.ActiveDocument.SelectionSets.Item(regAPP)
        'End Try
        ''
        ''
        'oApp.ActiveDocument.SetVariable("pickadd", 0)   '' Solo una selección. Se quita lo que hubiera
        'oSel.Clear()
        'Try
        'oSel.Select(vF1, vF2)
        'Catch ex As System.Exception
        'Debug.Print(ex.Message)
        'End Try
        'MsgBox(oSel.Count)
        If oSel IsNot Nothing AndAlso oSel.Count > 0 Then
            For x As Integer = 0 To oSel.Count - 1
                resultado.Add(oApp.ActiveDocument.ObjectIdToObject(oSel.Item(x).ObjectId.OldIdPtr))
            Next
        Else
            resultado = Nothing
        End If
        ''
        oApp.ActiveDocument.SetVariable("pickadd", 2)   '' La seleccion actual se suma a la que hubiera.
        Return resultado
        Exit Function
    End Function
    ''
    Public Function SeleccionaDameBloquesEnCapa(queCapa As String) As ArrayList
        Dim resultado As New ArrayList
        'Dim cSeleccion As AcadSelectionSets
        Dim F1(0) As Short
        Dim F2(0) As Object
        Dim vF1 As Object = Nothing
        Dim vF2 As Object = Nothing

        '' Las 2 maneras valen igual. AcDbBlckReference es mejor (Solo coge bloques) INSERT coge sombreados también.
        F1(0) = 100 : F2(0) = "AcDbBlockReference"
        F1(0) = 8 : F2(0) = queCapa
        vF1 = F1
        vF2 = F2
        ''
        If oApp Is Nothing Then _
        oApp = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
        ''
        ''
        Try
            oSel = oApp.ActiveDocument.SelectionSets.Add(regAPP)
        Catch ex As System.Exception
            oSel = oApp.ActiveDocument.SelectionSets.Item(regAPP)
        End Try
        ''
        ''
        'oApp.ActiveDocument.SetVariable("pickadd", 0)   '' Solo una selección. Se quita lo que hubiera
        oSel.Clear()
        Try
            oSel.Select(vF1, vF2)
        Catch ex As System.Exception
            'Debug.Print(ex.Message)
        End Try
        'MsgBox(oSel.Count)
        If oSel.Count > 0 Then
            For x As Integer = 0 To oSel.Count - 1
                resultado.Add(oSel.Item(x))
            Next
        Else
            resultado = Nothing
        End If
        ''
        oApp.ActiveDocument.SetVariable("pickadd", 2)   '' La seleccion actual se suma a la que hubiera.
        Return resultado
        Exit Function
    End Function
    ''
    Public Function DameNombresCapasVacias() As ArrayList
        Dim resultado As New ArrayList
        ''
        'Add bit about counting text on layer
        oApp.ActiveDocument.SetVariable("pickadd", 0)   '' Se quitan las selecciones anteriores
        For Each oLa As AcadLayer In oApp.ActiveDocument.Layers
            Dim queCapa As String = oLa.Name
            Dim myTVs(0) As TypedValue
            myTVs.SetValue(New TypedValue(DxfCode.LayerName, queCapa), 0)
            Dim myFilter As New SelectionFilter(myTVs)
            Dim myEd As Editor = acApp.DocumentManager.MdiActiveDocument.Editor
            Dim myPSR As PromptSelectionResult = myEd.SelectAll(myFilter)
            Dim oSel As SelectionSet = Nothing
            If myPSR.Value.Count = 0 Then
                resultado.Add(queCapa)
            End If
        Next
        ''
        oApp.ActiveDocument.SetVariable("pickadd", 2)   '' La seleccion actual se suma a la que hubiera.
        Return resultado
    End Function
    ''
    Public Sub CapaCreaActiva(nombreCapa As String,
                              queColor As Integer,
                              Optional activa As Boolean = True,
                              Optional reutilizar As Boolean = True,
                              Optional escapaactiva As Boolean = False)
        ''
        '' Crear una capa y poner sus características.
        ''
        If oApp Is Nothing Then _
        oApp = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
        '' Activar Capa 0 primero.
        'CapaCeroActiva()
        ''
        '' Coger la capa BALIZAMIENTO SUELO o crearla
        Dim existia As Boolean = False
        Dim oLayer As AcadLayer = Nothing
        Try
            oLayer = oApp.ActiveDocument.Layers.Item(nombreCapa)
            existia = True
            ''
            '' Ya existía, salimos sin hacer nada mas.
            ''Exit Sub
        Catch ex As System.Exception
            oLayer = oApp.ActiveDocument.Layers.Add(nombreCapa)
            existia = False
        End Try
        ''
        '' Poner la capa como visible (LayerOn=True) y Reutilizada (Freeze=False)
        If activa = True And oLayer.LayerOn = False Then
            oLayer.LayerOn = True
        ElseIf activa = False And oLayer.LayerOn = True Then
            oLayer.LayerOn = False
        End If
        ''
        If reutilizar = True And oLayer.Freeze = True Then
            oLayer.Freeze = False
        ElseIf reutilizar = False And oLayer.Freeze = False Then
            oLayer.Freeze = True
        End If
        ''
        '' Poner alguna de sus características
        Dim oColor As New AcadAcCmColor
        If queColor >= 0 And queColor <= 256 Then
            oColor.ColorIndex = queColor
            If oLayer.TrueColor IsNot oColor Then oLayer.TrueColor = oColor
        Else
            oColor.ColorIndex = 7
            If oLayer.TrueColor IsNot oColor Then oLayer.TrueColor = oColor
        End If
        'If oLayer.Lineweight <> ACAD_LWEIGHT.acLnWt100 Then oLayer.Lineweight = ACAD_LWEIGHT.acLnWt100

        ''
        If escapaactiva = True And oApp.ActiveDocument.ActiveLayer.Name <> oLayer.Name Then
            oApp.ActiveDocument.ActiveLayer = oLayer
        End If
        ''
        oLayer = Nothing
        oColor = Nothing
        ''oApp.ActiveDocument.SendCommand("_line ")
        oApp.ActiveDocument.Regen(AcRegenType.acAllViewports)
    End Sub
    ''
    Public Sub CapaCreaActivaTablaDatos()
        '(SETQ capacuadro "CUADRO")
        '(IF   (NULL (TBLOBJNAME "LAYER" capacuadro)) 
        '(ENTMAKE (LIST '(0 . "LAYER")'(100 . "AcDbSymbolTableRecord")'(100 . "AcDbLayerTableRecord") 
        '                     (CONS 2 capacuadro)'(70 . 0)(CONS 62 7) (CONS 370 25))) 
        ' ) 
        '' 70 = reutilizar 0
        '' 62 = Color 7
        '' 370 = ACAD_LWEIGHT.acLnWt25 'grosor 2 mm
        ''
        '' Crear una capa y poner sus características.
        ''
        If oApp Is Nothing Then _
            oApp = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
        '' Activar Capa 0 primero.
        oApp.ActiveDocument.ActiveLayer = oApp.ActiveDocument.Layers.Item("0")
        ''
        '' Coger la capa BALIZAMIENTO SUELO o crearla
        Dim oLayer As AcadLayer = Nothing
        Try
            oLayer = oApp.ActiveDocument.Layers.Item("capacuadro")
        Catch ex As System.Exception
            oLayer = oApp.ActiveDocument.Layers.Add("capacuadro")
        End Try
        ''
        '' Poner la capa como visible (LayerOn=True) y Reutilizada (Freeze=False)
        oLayer.LayerOn = True
        oLayer.Freeze = False
        ''
        '' Poner alguna de sus características
        Dim oColor As New AcadAcCmColor
        oColor.ColorIndex = 7
        oLayer.TrueColor = oColor
        oLayer.Lineweight = ACAD_LWEIGHT.acLnWt025
        ''
        oApp.ActiveDocument.ActiveLayer = oLayer    ' oApp.ActiveDocument.Layers.Item("BALIZAMIENTO SUELO")
        oLayer = Nothing
        oColor = Nothing
        ''
        oApp.ActiveDocument.Regen(AcRegenType.acAllViewports)
    End Sub
    ''
    ''
    Public Sub CapaCreaActivaBalizamientoPared()
        ' (DEFUN c:BalizarPared ( / )
        '  (SETQ capabalizar1 "BALIZAMIENTO PARED")
        '   (IF   (NULL (TBLOBJNAME "LAYER" capabalizar1)) 
        '   (ENTMAKE (LIST '(0 . "LAYER")'(100 . "AcDbSymbolTableRecord")'(100 . "AcDbLayerTableRecord") 
        '                        (CONS 2 capabalizar1)'(70 . 0)(CONS 62 52) (CONS 370 100))) 
        '     );If 
        '(SETVAR "clayer" capabalizar1)
        '  (COMMAND "_line")
        '  (PRINC)
        ')
        '' 70 = reutilizar 0
        '' 62 = Color 52
        '' 370 = ACAD_LWEIGHT.acLnWt100 'grosor 2 mm
        ''
        '' Crear una capa y poner sus características.
        ''
        If oApp Is Nothing Then _
        oApp = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
        '' Activar Capa 0 primero.
        oApp.ActiveDocument.ActiveLayer = oApp.ActiveDocument.Layers.Item("0")
        ''
        '' Coger la capa BALIZAMIENTO SUELO o crearla
        Dim oLayer As AcadLayer = Nothing
        Try
            oLayer = oApp.ActiveDocument.Layers.Item("BALIZAMIENTO PARED")
        Catch ex As System.Exception
            oLayer = oApp.ActiveDocument.Layers.Add("BALIZAMIENTO PARED")
        End Try
        ''
        '' Poner la capa como visible (LayerOn=True) y Reutilizada (Freeze=False)
        oLayer.LayerOn = True
        oLayer.Freeze = False
        ''
        '' Poner alguna de sus características
        Dim oColor As New AcadAcCmColor
        oColor.ColorIndex = 52
        oLayer.TrueColor = oColor
        oLayer.Lineweight = ACAD_LWEIGHT.acLnWt100
        ''
        oApp.ActiveDocument.ActiveLayer = oLayer    ' oApp.ActiveDocument.Layers.Item("BALIZAMIENTO PARED")
        oLayer = Nothing
        oColor = Nothing
        ''oApp.ActiveDocument.SendCommand("_line ")
        oApp.ActiveDocument.Regen(AcRegenType.acAllViewports)
    End Sub
    ''
    ''
    Public Sub CapaCreaActivaBalizamientoEscalera()
        ' (SETQ capabalizarescalera "BALIZAMIENTO ESCALERA")
        '   (IF   (NULL (TBLOBJNAME "LAYER" capabalizarescalera)) 
        '   (ENTMAKE (LIST '(0 . "LAYER")'(100 . "AcDbSymbolTableRecord")'(100 . "AcDbLayerTableRecord") 
        '                        (CONS 2 capabalizarescalera)'(70 . 0)(CONS 62 68) (CONS 370 100))) 
        '     );If
        '(SETQ capacuadro "CUADRO")
        '   (IF   (NULL (TBLOBJNAME "LAYER" capacuadro)) 
        '   (ENTMAKE (LIST '(0 . "LAYER")'(100 . "AcDbSymbolTableRecord")'(100 . "AcDbLayerTableRecord") 
        '                        (CONS 2 capacuadro)'(70 . 0)(CONS 62 7) (CONS 370 25))) 
        '    )
        '' 70 = reutilizar 0
        '' 62 = Color 68
        '' 370 = ACAD_LWEIGHT.acLnWt100
        ''
        '' Crear una capa y poner sus características.
        ''
        If oApp Is Nothing Then _
        oApp = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
        '' Activar Capa 0 primero.
        CapaCeroActiva()
        ''
        '' Coger la capa BALIZAMIENTO SUELO o crearla
        Dim oLayer As AcadLayer = Nothing
        Try
            oLayer = oApp.ActiveDocument.Layers.Item("BALIZAMIENTO ESCALERA")
        Catch ex As System.Exception
            oLayer = oApp.ActiveDocument.Layers.Add("BALIZAMIENTO ESCALERA")
        End Try
        ''
        '' Poner la capa como visible (LayerOn=True) y Reutilizada (Freeze=False)
        oLayer.LayerOn = True
        oLayer.Freeze = False
        ''
        '' Poner alguna de sus características
        Dim oColor As New AcadAcCmColor
        oColor.ColorIndex = 7       ' Antes era 68, pero no se veían las lineas que dibujamos para anchura.
        oLayer.TrueColor = oColor
        oLayer.Lineweight = ACAD_LWEIGHT.acLnWt100
        ''
        oApp.ActiveDocument.ActiveLayer = oLayer    ' oApp.ActiveDocument.Layers.Item("BALIZAMIENTO PARED")
        oLayer = Nothing
        oColor = Nothing
        ''oApp.ActiveDocument.SendCommand("_line ")
        oApp.ActiveDocument.Regen(AcRegenType.acAllViewports)
    End Sub

    Public Sub CapaCeroActiva()
        '' Pondremos la capa 0 como activa.
        If oApp Is Nothing Then _
        oApp = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
        '' Activar Capa 0 primero.
        oApp.ActiveDocument.ActiveLayer = oApp.ActiveDocument.Layers.Item("0")
        ''
    End Sub
    ''
    Public Function CapaBorraVacias(arrPreCapas As ArrayList) As Integer
        Dim resultado As Integer = 0
        '' Podemos indicar un ArrayList con los prefijos de capas
        '' a tener en cuenta para borrar.
        '' Pondremos la capa 0 como activa.
        CapaCeroActiva()
        For Each oLay As AcadLayer In oApp.ActiveDocument.Layers
            Try
                For Each preC As String In arrPreCapas
                    If oLay.Name.ToUpper.StartsWith(preC.ToUpper) Then  ' AndAlso oLay.Used = False Then
                        Dim arrEnt As ArrayList = SeleccionaDameObjetosEnCapa(oLay.Name)
                        If arrEnt Is Nothing OrElse arrEnt.Count = 0 Then
                            oLay.Delete()
                            resultado += 1
                            Exit For
                        End If
                    End If
                Next
            Catch ex As System.Exception
                '' No hacemos nada, la capa no estaba vacia.
            End Try
        Next
        ''
        Return resultado
    End Function
    ''
    Public Function CapaBorraVaciasTodo() As Integer
        Dim resultado As Integer = 0
        '' Tiene en cuenta todas las capas que estén vacías.
        '' Pondremos la capa 0 como activa.
        CapaCeroActiva()
        'Dim arrCapasBorrar As New ArrayList     '' Arraylist de las capas a borrar
        Dim esCapaAnemed As Boolean = False
        For Each oLay As AcadLayer In oApp.ActiveDocument.Layers
            Try
                Dim arrEnt As ArrayList = SeleccionaDameObjetosEnCapa(oLay.Name)
                If arrEnt Is Nothing OrElse arrEnt.Count = 0 Then
                    'arrCapasBorrar.Add(oLay.Name)
                    'oApp.ActiveDocument.Layers.Item(oLay.Name).Delete()
                    oLay.Delete()
                    'Exit For
                    resultado += 1
                End If

            Catch ex As System.Exception
                '' No hacemos nada, la capa no estaba vacia.
            End Try
        Next
        ''
        Return resultado
    End Function
    ''
    Public Function CapaCambiaNivelBloques(queArrBlo As ArrayList, queNivel As String, queFrecuencia As String) As Integer()
        Dim resultado(1) As Integer
        resultado(0) = 0 : resultado(1) = 0     '' (0) Nº cambiados (1) Nº no cambiados
        If modVar.oApp Is Nothing Then _
        modVar.oApp = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
        ''
        '' Iteramos con el ArrayList de bloques insertados
        For Each oBlr As AcadBlockReference In queArrBlo
            Dim cambio As Boolean = False
            Dim oBl As AcadBlockReference = oApp.ActiveDocument.ObjectIdToObject(oBlr.ObjectID)
            ''
            Dim capaVieja As String = oBl.Layer
            Dim partes() As String = capaVieja.Split("·"c)
            Dim nombreBase As String = oBl.EffectiveName
            Dim nombreDinamico As String = oBl.Name
            Dim nombre As String = BloqueAtributoDame(oBl.ObjectID, "NOMBRE")
            If nombre = "" Then nombre = nombreBase
            Dim nivel As String = BloqueAtributoDame(oBl.ObjectID, "NIVEL")
            Dim frecuencia As String = BloqueAtributoDame(oBl.ObjectID, "FRECUENCIA")
            Dim capaNueva As String = nombre & "·" & queNivel
            ''
            If partes.Length > 2 Then
                capaNueva &= "·" & partes(2)
            End If
            '' Crear la capa, si no existe.
            CapaCreaActiva(capaNueva, -1)
            ''
            '' Ponemos atributos y capa.
            If capaVieja <> capaNueva Then
                oBl.Layer = capaNueva
                cambio = True
            End If
            If nivel <> queNivel Then
                BloqueAtributoEscribe(oBl.ObjectID, "NIVEL", queNivel)
                cambio = True
            End If
            If frecuencia <> queFrecuencia Then
                BloqueAtributoEscribe(oBl.ObjectID, "FRECUENCIA", queFrecuencia)
                cambio = True
            End If
            ''
            If cambio = True Then
                oBl.Update()
                resultado(0) += 1
            Else
                resultado(1) += 1
            End If
        Next
        ''
        oApp.ActiveDocument.Regen(AcRegenType.acActiveViewport)
        Return resultado
    End Function
    ''
    Public Function CapaCambiaZonaBloques(queArrBlo As ArrayList, queZona As String) As Integer()
        Dim resultado(1) As Integer
        resultado(0) = 0 : resultado(1) = 0     '' (0) Nº cambiados (1) Nº no cambiados
        ''
        If modVar.oApp Is Nothing Then _
        modVar.oApp = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
        ''
        '' Iteramos con el ArrayList de bloques insertados
        For Each oBlr As AcadBlockReference In queArrBlo
            Dim cambio As Boolean = False
            Dim oBl As AcadBlockReference = oApp.ActiveDocument.ObjectIdToObject(oBlr.ObjectID)
            ''
            Dim capaVieja As String = oBl.Layer
            Dim partes() As String = capaVieja.Split("·"c)
            Dim nombreBase As String = oBl.EffectiveName
            Dim nombreDinamico As String = oBl.Name
            Dim nombre As String = BloqueAtributoDame(oBl.ObjectID, "NOMBRE")
            If nombre = "" Then nombre = nombreBase
            Dim nivel As String = BloqueAtributoDame(oBl.ObjectID, "NIVEL")
            Dim frecuencia As String = BloqueAtributoDame(oBl.ObjectID, "FRECUENCIA")
            Dim zona As String = BloqueAtributoDame(oBl.ObjectID, "ZONA")
            ''
            Dim capaNueva As String = nombre & "·" & nivel & "·" & queZona
            ''
            '' Crear la capa, si no existe.
            CapaCreaActiva(capaNueva, -1)
            ''
            '' Ponemos atributos y capa.
            If capaVieja <> capaNueva Then
                oBl.Layer = capaNueva
            End If
            If zona <> queZona Then
                BloqueAtributoEscribe(oBl.ObjectID, "ZONA", queZona)
                cambio = True
            End If
            ''
            If cambio = True Then
                oBl.Update()
                resultado(0) += 1
            Else
                resultado(1) += 1
            End If
        Next
        ''
        oApp.ActiveDocument.Regen(AcRegenType.acActiveViewport)
        Return resultado
    End Function
    ''
    Public Sub CapaFreezePViewport(layers As List(Of String))
        '' Agregue nombres de capa para congelar/descongelar
        '' en las ventanas de espacio papel separados por comas
        ''Dim layers As List(Of String) = {"Wall", "ANNO-TEXT"}.ToList() ''<-- layers for test
        Dim doc As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = doc.Editor
        Dim db As Database = doc.Database
        Dim idLay As ObjectId
        Dim idLayTblRcd As ObjectId
        Dim lt As LayerTableRecord
        Dim layOut As Layout
        Dim tm As Autodesk.AutoCAD.ApplicationServices.TransactionManager = db.TransactionManager
        Dim ta As Transaction = tm.StartTransaction()
        Try
            Dim acLayoutMgr As LayoutManager
            acLayoutMgr = LayoutManager.Current
            Dim layDict As DBDictionary = DirectCast(ta.GetObject(db.LayoutDictionaryId, OpenMode.ForRead, False), DBDictionary)
            For Each itmdict As DBDictionaryEntry In layDict
                layOut = DirectCast(ta.GetObject(itmdict.Value, OpenMode.ForRead, False), Layout)
                ed.WriteMessage(vbLf + "Layout: {0}" + vbLf, layOut.LayoutName)
                If layOut.LayoutName <> "Model" Then
                    acLayoutMgr.CurrentLayout = layOut.LayoutName
                    Dim ltt As LayerTable = DirectCast(ta.GetObject(db.LayerTableId, OpenMode.ForRead, False), LayerTable)
                    For Each lname As String In layers
                        idLay = ltt(lname)
                        lt = ta.GetObject(idLay, OpenMode.ForRead)
                        If ltt.Has(lname) Then
                            idLayTblRcd = ltt.Item(lname)
                        Else
                            ed.WriteMessage("Layer: """ + lname + """ not available")
                            Return
                        End If
                        Dim idCol As ObjectIdCollection = New ObjectIdCollection
                        idCol.Add(idLayTblRcd)
                        ' Check that we are in paper space 
                        Dim vpid As ObjectId = ed.CurrentViewportObjectId
                        If vpid.IsNull() Then
                            ed.WriteMessage("No Viewport current.")
                            Return
                        End If
                        'VP need to be open for write 
                        Dim oViewport As Viewport = DirectCast(tm.GetObject(vpid, OpenMode.ForWrite, False), Viewport)
                        If Not oViewport.IsLayerFrozenInViewport(idLayTblRcd) Then
                            oViewport.FreezeLayersInViewport(idCol.GetEnumerator())
                        Else
                            oViewport.ThawLayersInViewport(idCol.GetEnumerator())
                        End If
                    Next
                End If
            Next
            ta.Commit()
        Finally
            ta.Dispose()
        End Try
    End Sub
    Sub VpLayerOff(oDoc As AcadDocument, VPlayer As String)

        '****************************************
        '*** Code from VisibleVisual.com ********
        '****************************************
        ' freeze the layer in the CURRENT viewport
        Dim objEntity As AcadObject = Nothing
        Dim objPViewport As AcadObject = Nothing
        Dim objPViewport2 As AcadObject = Nothing
        Dim XdataType As Object = Nothing
        Dim XdataValue As Object = Nothing
        Dim I As Integer
        Dim Counter As Integer
        Dim PT1 As Object = Nothing
        ' Get the active ViewPort
        objPViewport = oDoc.ActivePViewport
        ' Get the Xdata from the Viewport
        objPViewport.GetXData("ACAD", XdataType, XdataValue)
        For I = LBound(XdataType) To UBound(XdataType)
            ' Look for frozen Layers in this viewport
            If XdataType(I) = 1003 Then
                ' Set the counter AFTER the position of the Layer frozen layer(s)
                Counter = I + 1
                ' If the layer is already in the frozen layers xdata of this viewport the
                ' exit this sub program
                If XdataValue(I) = VPlayer Then Exit Sub
            End If
        Next
        ' If no frozen layers exist in this viewport then
        ' find the Xdata location 1002 and place the frozen layer infront of the "}"
        ' found at Xdata location 1002
        If Counter = 0 Then
            For I = LBound(XdataType) To UBound(XdataType)
                If XdataType(I) = 1002 Then Counter = I - 1
            Next
        End If
        ' set the Xdata for the layer that is beeing frozen
        XdataType(Counter) = 1003
        XdataValue(Counter) = VPlayer
        ReDim Preserve XdataType(Counter + 1)
        ReDim Preserve XdataValue(Counter + 1)
        ' put the first "}" back into the xdata array
        XdataType(Counter + 1) = 1002
        XdataValue(Counter + 1) = "}"
        ' Keep the xdata Array and add one more to the array
        ReDim Preserve XdataType(Counter + 2)
        ReDim Preserve XdataValue(Counter + 2)
        ' put the second "}" back into the xdata array
        XdataType(Counter + 2) = 1002
        XdataValue(Counter + 2) = "}"
        ' Reset the Xdata on to the viewport
        objPViewport.SetXData(XdataType, XdataValue)
        'If no change is visible run VPupdate after this code.
        VPupdateUpdate(oDoc)
    End Sub
    ''
    Sub VPupdateUpdate(oDoc As AcadDocument)
        ' Update the viewport...
        Dim objPViewport As AcadObject = oDoc.ActiveViewport
        oDoc.MSpace = False
        'objPViewport.Display (False)
        objPViewport.Display(True)
        oDoc.MSpace = True
        oDoc.Utility.Prompt("Viewport Updated!" & vbCr)
    End Sub
    ''
    ' Freezes the selected layers in all other existing viewport layouts
    Public Sub freezeOtherLayouts(ByVal pageNumber As Integer, ByVal layersToFreezeLayerIds As ObjectIdCollection)
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor
        Dim vp As Viewport = Nothing
        Dim viewPortFound As Boolean
        Dim freezeVPtrans As Transaction = Nothing

        Try
            freezeVPtrans = db.TransactionManager.StartTransaction()
            Dim myBT As BlockTable = db.BlockTableId.GetObject(OpenMode.ForRead)
            For Each btrID As ObjectId In myBT
                Dim myBTR As BlockTableRecord = btrID.GetObject(OpenMode.ForRead)
                ' If the block table record is a layout
                If myBTR.IsLayout Then
                    viewPortFound = False
                    If Not myBTR.Name = "*Model_Space" Then
                        Dim layOut As Layout = myBTR.LayoutId.GetObject(OpenMode.ForRead)
                        'Dim initId As ObjectId = Nothing
                        'initId = layOut.Initialize()
                        ' If the layout is not the new layout
                        If layOut.TabOrder <> pageNumber Then
                            SetCurrentLayoutTab(layOut.LayoutName)
                            For Each id As ObjectId In myBTR
                                Dim obj As DBObject = id.GetObject(OpenMode.ForWrite)
                                ' If the object is a viewport (there is a model viewport which is found first, we want the second one)
                                If TypeOf obj Is Viewport And viewPortFound = True Then
                                    Dim vpref As Viewport = DirectCast(obj, Viewport)
                                    ' Selected Viewport for write.
                                    vp = freezeVPtrans.GetObject(vpref.ObjectId, OpenMode.ForWrite)
                                    Dim layerTable As LayerTable = freezeVPtrans.GetObject(db.LayerTableId, OpenMode.ForRead)
                                    ' Freeze the selected layers in the viewports
                                    vp.FreezeLayersInViewport(layersToFreezeLayerIds.GetEnumerator())
                                    'vp.UpdateDisplay()
                                    ed.Regen()
                                    layerTable.Dispose()
                                End If
                                If TypeOf obj Is Viewport Then
                                    viewPortFound = True
                                End If
                            Next
                        End If
                    End If
                End If
            Next
        Catch
            ed.WriteMessage("Error!")
        Finally
            freezeVPtrans.Commit()
            ed.Regen()
            freezeVPtrans.Dispose()
        End Try
    End Sub
    ''
    Public Sub SetCurrentLayoutTab(ByVal LayoutName As String)
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Using doclock As DocumentLock = acDoc.LockDocument
            Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
                Try
                    Dim acLayoutMgr As LayoutManager
                    acLayoutMgr = LayoutManager.Current
                    acLayoutMgr.CurrentLayout = LayoutName
                    acTrans.Commit()
                Catch ex As Autodesk.AutoCAD.Runtime.Exception
                    Application.ShowAlertDialog(ex.Message & vbCr & ex.StackTrace)
                    acTrans.Abort()
                End Try
            End Using
        End Using

    End Sub
    Private Sub oTimer_Elapsed(sender As Object, e As Timers.ElapsedEventArgs) Handles oTimer.Elapsed
        '' Pondremos la capa 0 como activa.
        If oApp Is Nothing Then _
        oApp = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
        '' Activar Capa 0 primero.
        oApp.ActiveDocument.ActiveLayer = oApp.ActiveDocument.Layers.Item("0")
        ''
    End Sub
    ''
    Public Function GetPuntoDame_NET(mensaje As String) As Object
        Dim resultado As Object = Nothing
        '' Get the current database and start the Transaction Manager
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = acDoc.Editor
        Dim acCurDb As Database = acDoc.Database
        Dim pPtRes As PromptPointResult
        Dim pPtOpts As PromptPointOptions = New PromptPointOptions("")
        '' Prompt for the start point
        pPtOpts.Message = vbLf & mensaje
        pPtOpts.AllowNone = True    ' Permite terminar con boton derecho.
        Using oUi As EditorUserInteraction = ed.StartUserInteraction(Application.MainWindow.Handle)
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
    ''
    Public Function GetOpcionDame_NET() As String
        Dim resultado As String = ""
        ''
        If oApp Is Nothing Then _
        oApp = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
        'oDoc = oApp.ActiveDocument
        oApp.ActiveDocument.Activate()
        ''

        ' Define the list of valid keywords
        Dim listaopciones As String
        listaopciones = "Ascendente Descendente"
        oApp.ActiveDocument.Utility.InitializeUserInput(1, listaopciones)

        ' Prompt para que el usuario introduzca una palabra. Return "Ascendente", "Descendente" en
        ' resultado variable o puede introducir "A", "D".
        'Dim returnString As String
        resultado = oApp.ActiveDocument.Utility.GetKeyword(vbLf & "Tipo de Escalera [Ascendente/Descendente] : ")
        'MsgBox("You entered " & resultado, , "GetKeyword Example")
        Return resultado
    End Function
    ''
    Public Function GetOpcionPrimariaSecundariaDame() As String
        Dim resultado As String = ""
        ''
        If oApp Is Nothing Then _
        oApp = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
        'oDoc = oApp.ActiveDocument
        oApp.ActiveDocument.Activate()
        ''

        ' Define the list of valid keywords
        Dim listaopciones As String
        listaopciones = "Primaria Accesibilidad"
        oApp.ActiveDocument.Utility.InitializeUserInput(0, listaopciones)   '' 0 para poder contestar con return.

        ' Prompt para que el usuario introduzca una palabra. Return "Primaria", "Secundaria" en
        ' resultado variable o puede introducir "P", "S".
        'Dim returnString As String
        resultado = oApp.ActiveDocument.Utility.GetKeyword(vbLf & "Tipo de Ruta evacuación [Primaria/Accesibilidad] <Primaria>: ")
        'MsgBox("You entered " & resultado, , "GetKeyword Example")
        If resultado = "" Then resultado = "Primaria"
        Return resultado
    End Function
    ''
    Public Function DameOpcionTexto(mensaje As String, queopciones As String()) As String
        Dim resultado As String = ""
        ''
        If oApp Is Nothing Then _
        oApp = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
        'oDoc = oApp.ActiveDocument
        oApp.ActiveDocument.Activate()
        ''

        ' Define the list of valid keywords
        Dim listaopciones As String = String.Join(" ", queopciones)
        'listaopciones = "Primaria Secundaria"
        oApp.ActiveDocument.Utility.InitializeUserInput(0, listaopciones)   '' 0 para poder contestar con return.

        ' Prompt para que el usuario introduzca una palabra del array. queopciones(0) sera el valor por defecto.
        Dim cadenaopciones As String = String.Join("/", queopciones)
        Dim cadenadefecto As String = queopciones(0)
        resultado = oApp.ActiveDocument.Utility.GetKeyword(vbLf & mensaje & " [" & cadenaopciones & "] <" & cadenadefecto & ">: ")
        'MsgBox("You entered " & resultado, , "GetKeyword Example")
        If resultado = "" Then
            resultado = cadenadefecto
        ElseIf InStr(listaopciones, resultado) = 0 Then
            resultado = ""
        End If
        ''
        Return resultado
    End Function
    ''
    ''
    Public Function GetDINDame_NET() As String
        Dim resultado As String = ""
        ''
        If oApp Is Nothing Then _
        oApp = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
        'oDoc = oApp.ActiveDocument
        oApp.ActiveDocument.Activate()
        '' Define the list of valid keywords
        Dim listaopciones As String
        listaopciones = "A4 A3 A2 A1 A0"
        'oDoc.Utility.InitializeUserInput(1, listaopciones)
        oApp.ActiveDocument.Utility.InitializeUserInput(7, listaopciones)

        ' Prompt para que el usuario introduzca una palabra. Return "Ascendente", "Descendente" en
        ' resultado variable o puede introducir "A", "D".
        'Dim returnString As String
        resultado = oApp.ActiveDocument.Utility.GetKeyword(vbLf & "Formato DIN [A4/A3/A2/A1/A0] : ")
        'MsgBox("You entered " & resultado, , "GetKeyword Example")
        Return resultado
    End Function
    ''
    '' Le daremos array de ancho, largo (A4, A3, A2, A1 o A0), el largo actual (X) y el ancho actual (Y)
    '' Nos devolverá qué escala tenemos que aplicar para encajarlo en el papel indicado.
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
    ''
    Public Sub CierraDibujo(queFullname As String)
        If oApp Is Nothing Then _
    oApp = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
        '' Activar Capa 0 primero.
        For Each queDoc As AcadDocument In oApp.Documents
            If queDoc.FullName = queFullname Then
                queDoc.Close()
                Exit For
            End If
        Next
    End Sub

    Public Sub CierraDibujoTodos()
        If oApp Is Nothing Then _
oApp = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
        'oApp.Documents.Close()
        '' Activar Capa 0 primero.
        For Each queDoc As AcadDocument In oApp.Documents
            queDoc.Activate()
            queDoc.Save()
            'DameDependencias()
        Next
        oApp.Documents.Close()
    End Sub
    ''
    Public Function DameDependencias() As ArrayList
        Dim resultado As New ArrayList
        If oApp Is Nothing Then _
oApp = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)


        Dim objFDLCol As AcadFileDependencies
        Dim objFDL As AcadFileDependency

        objFDLCol = oApp.ActiveDocument.FileDependencies
        ' MsgBox("Nº de File Dependency List = " & objFDLCol.Count)   ' & ".")

        'Dim FDLIndex As Long
        'FDLIndex = objFDLCol.CreateEntry("acad:xref", "c:\referenced.dwg", True, True)
        'MsgBox("The number of entries in the File Dependency List is " & objFDLCol.Count & ".")

        'Dim IndexNumber As Long
        'IndexNumber = objFDLCol.IndexOf("acad:xref", "c:\referenced.dwg")
        'Dim IndexString As String
        'IndexString = CStr(IndexNumber)
        'MsgBox("The index of the new entry is " & IndexString & ".")

        'objFDLCol.UpdateEntry(FDLIndex)

        'objFDLCol.RemoveEntry(FDLIndex, True)
        'MsgBox("The number of entries in the File Dependency List is " & objFDLCol.Count & ".")
        Dim mensaje As String = ""
        Try
            If objFDLCol IsNot Nothing AndAlso objFDLCol.Count > 0 Then
                For Each objFDL In objFDLCol
                    mensaje &= objFDL.FullFileName & vbCrLf
                    If resultado.Contains(objFDL.FullFileName) = False Then resultado.Add(objFDL.FullFileName)
                Next
                Application.ShowAlertDialog(mensaje)
            Else
                Application.ShowAlertDialog("No hay Dependencias...")
            End If
        Catch ex As System.Exception
            '' no hacemos nada.
        End Try
        If mensaje <> "" Then MsgBox(mensaje)
        Return resultado
    End Function
    ''
    Public Function XrefComprueba() As Hashtable
        Dim resultado As Hashtable = New Hashtable
        If oApp Is Nothing Then _
oApp = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
        ''
        Dim mensaje As String = ""
        For Each oEref As AcadBlock In oApp.ActiveDocument.Blocks
            If TypeOf oEref Is RasterImage Then
                Debug.Print(oEref.Name)
            End If
            If oEref.IsXRef AndAlso IO.File.Exists(oEref.Path) = False Then
                If resultado.Contains(oEref.EffectiveName) = False Then
                    resultado.Add(oEref.EffectiveName, oEref.Path)
                    mensaje &= oEref.EffectiveName & " / " & oEref.Path & vbCrLf
                End If
            End If
        Next
        ''
        If mensaje <> "" Then MsgBox(mensaje)
        Return resultado
    End Function

    ''
    Public Function XrefImagenDame() As Hashtable
        Dim resultado As Hashtable = New Hashtable
        If oApp Is Nothing Then _
oApp = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
        ''
        Dim mensaje As String = ""
        Dim oDict As AcadDictionary
        oDict = oApp.ActiveDocument.Dictionaries.Item("ACAD_IMAGE_DICT")
        Dim oImage As AcadRasterImage = Nothing
        Dim oImageDef As Object
        MsgBox(oDict.Count)
        For Each oImageDef In oDict
            'If TypeOf oImageDef Is AcadRasterImage Then
            MsgBox(oImageDef.GetType.ToString)
            'oImage = oImageDef
            'If resultado.Contains(oImage.Name) = False Then
            '    resultado.Add(oImage.Name, oImage)
            '    mensaje &= oImage.Name & " / " & oImage.ImageFile & vbCrLf
            'End If
            'End If
        Next
        ''
        If mensaje <> "" Then MsgBox(mensaje)
        Return resultado
    End Function
    ''
    Public Function UltimoObjeto() As AcadEntity
        If oApp Is Nothing Then _
        oApp = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
        ''
        Return oApp.ActiveDocument.ModelSpace.Item(oApp.ActiveDocument.ModelSpace.Count - 1)
    End Function
    ''
    Public Function DameSombreadoContornos(queSom As AcadHatch) As ArrayList
        If oApp Is Nothing Then _
        oApp = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
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
    ''
    Public Function DameBloque(quenombre As String) As AcadBlockReference
        Dim oBl As AcadBlockReference = Nothing
        If oApp Is Nothing Then _
        oApp = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
        ''
        For Each oEnt As AcadEntity In oApp.ActiveDocument.ModelSpace
            If TypeOf oEnt Is AcadBlockReference Then
                oBl = oApp.ActiveDocument.ObjectIdToObject(oEnt.ObjectID)
                If oBl.Name = quenombre Or oBl.EffectiveName = quenombre Then
                    Exit For
                Else
                    oBl = Nothing
                End If
            End If
        Next
        Return oBl
    End Function
    ''
    Public Function DameBloquesNombre(quenombre As String) As ArrayList
        Dim resultado As New ArrayList
        If oApp Is Nothing Then _
        oApp = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
        ''
        For Each oEnt As AcadEntity In oApp.ActiveDocument.ModelSpace
            If TypeOf oEnt Is AcadBlockReference Then
                Dim oBl As AcadBlockReference = Nothing
                oBl = oApp.ActiveDocument.ObjectIdToObject(oEnt.ObjectID)
                If oBl.Name = quenombre Or oBl.EffectiveName = quenombre Then
                    resultado.Add(oBl)
                End If
                oBl = Nothing
            End If
        Next
        Return resultado
    End Function
    ''
    Public Function DameBloquesNombreEmpiezaPor(quenombre As String) As ArrayList
        Dim resultado As New ArrayList
        If oApp Is Nothing Then _
        oApp = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
        ''
        For Each oEnt As AcadEntity In oApp.ActiveDocument.ModelSpace
            If TypeOf oEnt Is AcadBlockReference Then
                Dim oBl As AcadBlockReference = Nothing
                oBl = oApp.ActiveDocument.ObjectIdToObject(oEnt.ObjectID)
                If oBl.Name = quenombre Or oBl.EffectiveName = quenombre Or oBl.Name.StartsWith(quenombre) Or oBl.EffectiveName.StartsWith(quenombre) Then
                    resultado.Add(oBl)
                End If
                oBl = Nothing
            End If
        Next
        Return resultado
    End Function
    ''

    '***********************************************************************
    '** procedure de détection si un point est dans une région(polyline) **
    '***********************************************************************
    '** ENTREE : **
    '** - Ligne : la polyline **
    '** Attention vérifier avant que la polyligne ne se **
    '** croise pas elle-même (IsPolylineSelfIntercept) **
    '** - Bloc : Le bloc a tester **
    '** SORTIE : **
    '** - True si dedans sinon false **
    '***********************************************************************
    Public Function IsBlocInZoneFromPolyline(ByRef Bloc As BlockReference, ByRef Ligne As Polyline) As Boolean
        Dim OkRetour As Boolean = False
        '===========================================================================
        'Gestion des erreurs d'entrée :
        'Si pas définit
        If IsNothing(Bloc) = True Or IsNothing(Ligne) = True Then Return OkRetour
        'Si ligne non fermée
        If Ligne.Closed = False Then Return False
        '===========================================================================
        '===========================================================================
        'génération des points d'entrée de la fonction de calcul + vérification simple
        'du point de test
        Dim pt As New Autodesk.AutoCAD.Geometry.Point2d(Bloc.Position.X, Bloc.Position.Y)
        ' si le point est sur la polyligne
        Dim vn As Integer = Ligne.NumberOfVertices
        Dim colpt() As Autodesk.AutoCAD.Geometry.Point2d = Nothing
        ReDim colpt(vn)
        For i As Integer = 0 To vn - 1
            '// Could also get the 3D point here
            Dim pts As Autodesk.AutoCAD.Geometry.Point2d = Ligne.GetPoint2dAt(i)
            colpt(i) = New Autodesk.AutoCAD.Geometry.Point2d(pts.X, pts.Y)
            'Test si le point est sur un segment de la polyline si oui exit !
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
        'ajout du point [0] pour finir la boucle !
        colpt(vn) = Ligne.GetPoint2dAt(0)
        '===========================================================================
        Dim RetFonction As Double
        RetFonction = wn_PnPoly(pt, colpt, vn)
        If RetFonction = 0 Then
            OkRetour = False
        Else
            OkRetour = True
        End If
        Return OkRetour
    End Function
    ''
    Public Function IsPoint2DInZoneFromPolyline(ByRef p2d As Autodesk.AutoCAD.Geometry.Point3d, ByRef Ligne As Polyline) As Boolean
        Dim OkRetour As Boolean = False
        '===========================================================================
        'Gestion des erreurs d'entrée :
        'Si pas définit
        If IsNothing(p2d) = True Or IsNothing(Ligne) = True Then Return OkRetour
        'Si ligne non fermée
        If Ligne.Closed = False Then Return False
        '===========================================================================
        '===========================================================================
        'génération des points d'entrée de la fonction de calcul + vérification simple
        'du point de test
        Dim pt As New Autodesk.AutoCAD.Geometry.Point2d(p2d.X, p2d.Y)
        ' si le point est sur la polyligne
        Dim vn As Integer = Ligne.NumberOfVertices
        Dim colpt() As Autodesk.AutoCAD.Geometry.Point2d = Nothing
        ReDim colpt(vn)
        For i As Integer = 0 To vn - 1
            '// Could also get the 3D point here
            Dim pts As Autodesk.AutoCAD.Geometry.Point2d = Ligne.GetPoint2dAt(i)
            colpt(i) = New Autodesk.AutoCAD.Geometry.Point2d(pts.X, pts.Y)
            'Test si le point est sur un segment de la polyline si oui exit !
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
        'ajout du point [0] pour finir la boucle !
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
    '**************************************************************************************
    '// Copyright 2000 softSurfer, 2012 Dan Sunday
    '// This code may be freely used and modified for any purpose
    '// providing that this copyright notice is included with it.
    '// SoftSurfer makes no warranty for this code, and cannot be held
    '// liable for any real or imagined damage resulting from its use.
    '// Users of this code must verify correctness for their application.
    '**************************************************************************************
    'isLeft(): tests if a point is Left|On|Right of an infinite line.
    '// Input: three points P0, P1, and P2
    '// Return: >0 for P2 left of the line through P0 and P1
    '// =0 for P2 on the line
    '// <0 for P2 right of the line
    '// See: Algorithm 1 "Area of Triangles and Polygons"
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
    Private Function wn_PnPoly(ByVal P As Autodesk.AutoCAD.Geometry.Point2d, ByVal V() As Autodesk.AutoCAD.Geometry.Point2d, ByVal n As Double) As Double
        Dim wn As Double = 0 '// the winding number counter
        '// loop through all edges of the polygon
        For i As Integer = 0 To n - 1
            'for (int i=0; i<n; i++) {
            If (V(i).Y <= P.Y) Then '// edge from V[i] to V[i+1]
                '// start y <= P.y
                If (V(i + 1).Y > P.Y) Then '// an upward crossing
                    If (isLeft(V(i), V(i + 1), P) > 0) Then '// P left of edge
                        wn = wn + 1 '// have a valid up intersect
                    End If
                End If
            Else '// start y > P.y (no test needed)
                If (V(i + 1).Y <= P.Y) Then '// a downward crossing
                    If (isLeft(V(i), V(i + 1), P) < 0) Then '// P right of edge
                        wn = wn - 1 '// have a valid down intersect
                    End If
                End If
            End If
        Next
        Return wn
    End Function
    '**************************************************************************************
#Region "RIBBON"
    '' Con AdWindows (Autodesk.Windows)
    '' menugroup = IMO2017
    '' Alias id del Ribbon IMO son : ID_TABIMO, IDTAMIMO, TABIMO
    ''
    '' Grupos IMO2017:
    ''      IMOUTILIDADES, Alias = IMOUTILIDADES, IMOUTIL, IMOU
    ''      IMOPROYECTOS, Alias = IMOPROYECTOS, IMOPROJECT, IMOP
    ''      IMOBLOQUES, Alias = IMOBLOQUES, IMOBLOCKS, IMOB
    Public Function RibbonTabDameAdWindows(menugroup As String, idRibbonTab As String) As Autodesk.Windows.RibbonTab
        Dim resultado As Autodesk.Windows.RibbonTab = Nothing
        ''
        Dim rc As Autodesk.Windows.RibbonControl = Autodesk.Windows.ComponentManager.Ribbon
        ''
        Dim rTab As Autodesk.Windows.RibbonTab = rc.FindTab(menugroup & "." & idRibbonTab) ''  AutoCA Home Tab = rc.FindTab("ACAD.ID_TabHome3D")
        ''
        If rTab IsNot Nothing Then
            resultado = rTab
        End If
        ''
        Return resultado
    End Function
    '' Con AcCui (Autodesk.AutoCAD.Customization)
    '' Menugroup = IMO2017
    '' Alias del Ribbon IMO son : ID_TABIMO, IDTAMIMO, TABIMO
    Public Function RibbonTabDameAcCui(menugroup As String, idRbbonTab As String) As Autodesk.AutoCAD.Customization.WSRibbonTabSourceReference
        Dim resultado As Autodesk.AutoCAD.Customization.WSRibbonTabSourceReference = Nothing
        ''
        Dim menuName As String = acApp.GetSystemVariable("MENUNAME")
        Dim curWorkspace As String = acApp.GetSystemVariable("WSCURRENT")
        ''
        Dim cs As Autodesk.AutoCAD.Customization.CustomizationSection = New Autodesk.AutoCAD.Customization.CustomizationSection(menuName & ".cuix")
        Dim rr As Autodesk.AutoCAD.Customization.WSRibbonRoot = cs.getWorkspace(curWorkspace).WorkspaceRibbonRoot
        Dim tab As Autodesk.AutoCAD.Customization.WSRibbonTabSourceReference = rr.FindTabReference(menugroup, idRbbonTab)  '' AutoCAD Home tab = rr.FindTabReference("ACAD", "ID_TabHome3D")
        ''
        If tab IsNot Nothing Then
            resultado = tab
        End If
        ''
        If cs.IsModified Then
            cs.Save()
            acApp.ReloadAllMenus()
        End If

        ''
        Return resultado
    End Function
#End Region
End Module