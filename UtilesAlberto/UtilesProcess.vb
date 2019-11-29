Imports System.Runtime.InteropServices

Partial Public Class Utiles
    'Private Declare Function SetFocus Lib "user32.dll" (ByVal hwnd As IntPtr) As IntPtr
    'Dim p As Process = Process.GetProcessById("ID Here")
    'Dim p As Process = Process.GetProcessByName("Name Here")
    'SetFocus(p.MainWindowHandle)

    <StructLayout(LayoutKind.Sequential)>
    Structure RM_UNIQUE_PROCESS
        Public dwProcessId As Integer
        Public ProcessStartTime As System.Runtime.InteropServices.ComTypes.FILETIME
    End Structure

    Const RmRebootReasonNone As Integer = 0
    Const CCH_RM_MAX_APP_NAME As Integer = 255
    Const CCH_RM_MAX_SVC_NAME As Integer = 63

    Enum RM_APP_TYPE
        RmUnknownApp = 0
        RmMainWindow = 1
        RmOtherWindow = 2
        RmService = 3
        RmExplorer = 4
        RmConsole = 5
        RmCritical = 1000
    End Enum

    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode)>
    Structure RM_PROCESS_INFO
        Public Process As RM_UNIQUE_PROCESS
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=CCH_RM_MAX_APP_NAME + 1)>
        Public strAppName As String
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=CCH_RM_MAX_SVC_NAME + 1)>
        Public strServiceShortName As String
        Public ApplicationType As RM_APP_TYPE
        Public AppStatus As UInteger
        Public TSSessionId As UInteger
        <MarshalAs(UnmanagedType.Bool)>
        Public bRestartable As Boolean
    End Structure

    <DllImport("rstrtmgr.dll", CharSet:=CharSet.Unicode)>
    Private Shared Function RmRegisterResources(ByVal pSessionHandle As UInteger, ByVal nFiles As UInt32, ByVal rgsFilenames As String(), ByVal nApplications As UInt32,
    <[In]> ByVal rgApplications As RM_UNIQUE_PROCESS(), ByVal nServices As UInt32, ByVal rgsServiceNames As String()) As Integer
    End Function

    <DllImport("rstrtmgr.dll", CharSet:=CharSet.Auto)>
    Private Shared Function RmStartSession(<Out> ByRef pSessionHandle As UInteger, ByVal dwSessionFlags As Integer, ByVal strSessionKey As String) As Integer
    End Function

    <DllImport("rstrtmgr.dll")>
    Private Shared Function RmEndSession(ByVal pSessionHandle As UInteger) As Integer
    End Function

    <DllImport("rstrtmgr.dll")>
    Private Shared Function RmGetList(ByVal dwSessionHandle As UInteger, <Out> ByRef pnProcInfoNeeded As UInteger, ByRef pnProcInfo As UInteger,
    <[In], Out> ByVal rgAffectedApps As RM_PROCESS_INFO(), ByRef lpdwRebootReasons As UInteger) As Integer
    End Function

    Public Shared Function Process_GetProcessBlockFile(ByVal path As String) As System.Collections.Generic.List(Of Process)
        Dim handle As UInteger
        Dim key As String = Guid.NewGuid().ToString()
        Dim processes As System.Collections.Generic.List(Of Process) = New System.Collections.Generic.List(Of Process)()
        Dim res As Integer = RmStartSession(handle, 0, key)
        If res <> 0 Then Throw New Exception("Could not begin restart session.  Unable to determine file locker.")

        Try
            Const ERROR_MORE_DATA As Integer = 234
            Dim pnProcInfoNeeded As UInteger = 0, pnProcInfo As UInteger = 0, lpdwRebootReasons As UInteger = RmRebootReasonNone
            Dim resources As String() = New String() {path}
            res = RmRegisterResources(handle, CUInt(resources.Length), resources, 0, Nothing, 0, Nothing)
            If res <> 0 Then Throw New Exception("Could not register resource.")
            res = RmGetList(handle, pnProcInfoNeeded, pnProcInfo, Nothing, lpdwRebootReasons)

            If res = ERROR_MORE_DATA Then
                Dim processInfo As RM_PROCESS_INFO() = New RM_PROCESS_INFO(pnProcInfoNeeded - 1) {}
                pnProcInfo = pnProcInfoNeeded
                res = RmGetList(handle, pnProcInfoNeeded, pnProcInfo, processInfo, lpdwRebootReasons)

                If res = 0 Then
                    processes = New System.Collections.Generic.List(Of Process)(CInt(pnProcInfo))

                    For i As Integer = 0 To pnProcInfo - 1

                        Try
                            processes.Add(Process.GetProcessById(processInfo(i).Process.dwProcessId))
                        Catch __unusedArgumentException1__ As ArgumentException
                        End Try
                    Next
                Else
                    Throw New Exception("Could not list processes locking resource.")
                End If
            ElseIf res <> 0 Then
                Throw New Exception("Could not list processes locking resource. Failed to get size of result.")
            End If

        Finally
            RmEndSession(handle)
        End Try
        Return processes
    End Function

    Public Shared Sub Process_Activate(nProcess As String)
        Dim processID As Integer = Process_GetID(nProcess)
        Try
            If processID <> 0 Then AppActivate(processID)
        Catch ex As Exception
            '    'MessageBox.Show("No hay instancias de " & queProceso & " en ejecución.")
        End Try
    End Sub

    Public Shared Sub Process_SetFocus(nProcess As String)
        Dim p() As Process = Nothing
        Try
            p = Process.GetProcessesByName(nProcess)
            If p IsNot Nothing AndAlso p.Length > 0 Then
                SetFocus(p(0).MainWindowHandle)
            End If
        Catch ex As Exception

        End Try
    End Sub
    '
    Public Shared Sub Process_SetFocus(idProcess As Integer)
        Dim p As Process = Nothing
        Try
            p = Process.GetProcessById(idProcess)
            If p IsNot Nothing Then
                SetFocus(p.MainWindowHandle)
            End If
        Catch ex As Exception

        End Try
    End Sub

    Public Shared Function Process_GetID(nProcess As String) As Integer
        Dim processID As Integer = 0
        Try
            Dim proc As System.Diagnostics.Process
            ''
            For Each proc In System.Diagnostics.Process.GetProcessesByName(nProcess)
                Try
                    processID = proc.Id
                    'proc.MainWindowHandle
                    Exit For
                Catch ex As Exception
                    ' No hacemos nada
                End Try
            Next
        Catch ex As Exception
            '    'MessageBox.Show("No hay instancias de " & queProceso & " en ejecución.")
        End Try
        Return processID
    End Function
    Public Shared Sub Process_Close(nProcess As String)   'Optional queProceso As String = "EXCEL"
        Try
            Dim proc As System.Diagnostics.Process
            ''
            For Each proc In System.Diagnostics.Process.GetProcessesByName(nProcess)
                Try
                    proc.Kill()
                Catch ex As Exception
                    ' No hacemos nada
                End Try
            Next
        Catch ex As Exception
            '    'MessageBox.Show("No hay instancias de " & queProceso & " en ejecución.")
        End Try
    End Sub
    '
    Public Shared Sub Process_Close(oProcess As Process)   'Optional queProceso As String = "EXCEL"
        Try
            Dim proc As System.Diagnostics.Process
            ''
            For Each proc In System.Diagnostics.Process.GetProcessesByName(oProcess.ProcessName)
                Try
                    proc.Kill()
                Catch ex As Exception
                    ' No hacemos nada
                End Try
            Next
        Catch ex As Exception
            '    'MessageBox.Show("No hay instancias de " & queProceso & " en ejecución.")
        End Try
    End Sub
    '
    Public Shared Sub Process_CloseProcessBlockFile(queFi As String)
        Dim procesos As System.Collections.Generic.List(Of Process) = Process_GetProcessBlockFile(queFi)
        If procesos IsNot Nothing AndAlso procesos.Count > 0 Then
            For Each queP As Process In procesos
                Process_Close(queP)
            Next
            Threading.Thread.Sleep(1000)
        End If
    End Sub

    Public Shared Sub Process_CloseProcessBlockFiles(arrFi As String())
        For Each queFi As String In arrFi
            If IO.File.Exists(queFi) Then Process_CloseProcessBlockFile(queFi)
        Next
    End Sub
End Class
