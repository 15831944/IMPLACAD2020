Imports System.Management
Imports System.Net
Imports System.Text
Imports System.Text.RegularExpressions
Partial Public Class Utiles
    ' Devuelve la IP privada (Intranet, por defecto). La que tiene este equipo en la Red interna de la empresa
    ' O todas, si solointranet = False
    Public Shared Function IPPrivada_DameLista(Optional solointranet As Boolean = True) As String
        Dim valorIp As String = ""
        Dim listAddres = Dns.GetHostEntry(My.Computer.Name).AddressList
        If solointranet Then
            valorIp = listAddres.FirstOrDefault(
                Function(i) i.AddressFamily = Sockets.AddressFamily.InterNetwork).ToString()
        Else
            ' ** Solo InterNetwork
            'Dim varios = From i In listAddres
            '             Where i.AddressFamily = Sockets.AddressFamily.InterNetwork
            '             Select i

            'Dim ips As New List(Of String)
            'For Each ip As IPAddress In varios
            '    If ip.ToString.StartsWith("192") Then ips.Add(ip.ToString)
            'Next
            ' ** Todas
            Dim ips As New List(Of String)
            For Each ip As IPAddress In listAddres
                If ip.ToString.StartsWith("192") Then ips.Add(ip.ToString)
            Next

            If ips.Count > 0 Then valorIp = String.Join("|", ips.ToArray)
        End If
        Return valorIp
    End Function

    Public Shared Function IPPrivada_DameCorto() As String
        Dim ip As System.Net.IPHostEntry
        ip = Dns.GetHostEntry(My.Computer.Name)
        Return ip.AddressList.Where(Function(i) i.ToString.StartsWith("192"))(0).ToString
    End Function
    Public Shared Function IPPublica_Dame() As String  ' IPAddress
        Dim resultado As String = ""
        Dim lol As WebClient = New WebClient()
        Try
            Dim str As String = lol.DownloadString("http://checkip.dyndns.org/")    ' El html con el resultado
            resultado = str.Split(":"c)(1).Split("<"c)(0).Trim
        Catch ex As Exception
            ' Error de conexión, no hay internet o la página no ha dado la IP
            resultado = "¿?"
        End Try

        Return resultado
    End Function

    Public Shared Function DameMac(Optional soloactivas As Boolean = False, Optional TambienNombreEquipo As Boolean = True) As String
        Dim resultado As String = ""
        Dim oMang As New System.Management.ManagementClass("Win32_NetworkAdapterConfiguration")
        Dim adaptadores As System.Management.ManagementObjectCollection = oMang.GetInstances
        ''
        Dim adapterEnumerator As ManagementObjectCollection.ManagementObjectEnumerator = adaptadores.GetEnumerator()
        While adapterEnumerator.MoveNext()
            Dim adapter As ManagementObject = CType(adapterEnumerator.Current, ManagementObject)
            Dim temp As String = CStr(adapter.Item("MacAddress"))
            If temp = "" Then Continue While
            If CBool(adapter.Item("IPEnabled")) = True And soloactivas = True Then
                resultado &= CStr(adapter.Item("MacAddress")) & ","
            ElseIf soloactivas = False Then
                resultado &= CStr(adapter.Item("MacAddress")) & ","
            End If
        End While
        ''
        resultado = resultado.Remove(resultado.Length - 1, 1)
        resultado = resultado.Replace(",", vbCrLf)

        If TambienNombreEquipo Then
            Return resultado & vbCrLf & System.Environment.MachineName
        Else
            Return resultado '& vbCrLf & vbCrLf & resultado.Split(vbCrLf).GetLength(0)
        End If
    End Function
End Class
