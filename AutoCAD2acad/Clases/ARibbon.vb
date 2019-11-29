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
            Dim menuName As String = Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable("MENUNAME")
            Dim curWorkspace As String = Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable("WSCURRENT")
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
                Autodesk.AutoCAD.ApplicationServices.Application.ReloadAllMenus()
            End If

            ''
            Return resultado
        End Function
        Public Sub RibbonActiva(activar As Boolean)
            '' rellenar rps con el RibbonPaletteSet actual
            Dim rps As Autodesk.AutoCAD.Ribbon.RibbonPaletteSet
            rps = Autodesk.AutoCAD.Ribbon.RibbonServices.CreateRibbonPaletteSet()
            ''
            rps.RibbonControl.IsEnabled = activar
            ''
            Dim _showTipsOnDisabled As Boolean
            If activar = False Then
                ' Store the current setting for "Show tooltips when the ribbon is disabled"
                ' And then modify the setting
                _showTipsOnDisabled = Autodesk.Windows.ComponentManager.ToolTipSettings.ShowOnDisabled
                Autodesk.Windows.ComponentManager.ToolTipSettings.ShowOnDisabled = activar
            Else
                ' Restore the setting for "Show tooltips when the ribbon is disabled"
                Autodesk.Windows.ComponentManager.ToolTipSettings.ShowOnDisabled = _showTipsOnDisabled
            End If
            ''      
            '' Enable Or disable background tab rendering
            rps.RibbonControl.IsBackgroundTabRenderingEnabled = activar
        End Sub
        ''
        Public Sub ToolbarActiva(activar As Boolean)
            Dim _visibleToolbars = New List(Of String)
            ''
            ''Clear the list Of toolbars that were previously visible, If we're disabling
            If activar = False Then _visibleToolbars.Clear()
            ''
            '' Use dynamic .NET to avoid the COM typelib reference
            '' (all the implicit "vars" below will also be resolved to dynamics)
            Dim app As Object = Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication
            ''
            '' Iterate through all the menu groups
            Dim mgs As Object = app.MenuGroups
            ''
            For Each mg As AcadMenuGroup In CType(mgs, AcadMenuGroups)
                '' Iterate through a menu group's toolbars
                Dim tbs As AcadToolbars = mg.Toolbars
                ''
                For Each tb As AcadToolbar In tbs
                    '' If we're disabling, check whether the toolbar is visible
                    '' And, if so, add its ID to the list before turning it off
                    If activar = False Then
                        If tb.Visible = True Then
                            _visibleToolbars.Add(tb.TagString)
                            tb.Visible = activar
                        End If
                    Else
                        '' If we're enabling, check whether the toolbar is in our list
                        '' before turning it on
                        If _visibleToolbars.Contains(tb.TagString) Then
                            tb.Visible = activar
                        End If
                    End If
                Next
            Next
        End Sub
    End Class
End Namespace