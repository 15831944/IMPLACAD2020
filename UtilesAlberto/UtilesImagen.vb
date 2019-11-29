Imports Microsoft.VisualBasic
Imports System
' Para el PictureBox
Imports System.Windows.Forms
' Para Icon y Bitmap
Imports System.Drawing
' Para DllImport
Imports System.Runtime.InteropServices


''' <summary>
''' Clase para extraer iconos de ficheros de recursos. O interactuar con imagenes
''' </summary>
''' <remarks></remarks>
Partial Public Class Utiles
    ' Cargar imagen desde un fichero en disco
    Public Shared Function Image_FromFile(ByVal fileName As String) As Image
        Return System.Drawing.Image.FromFile(fileName)
    End Function
    '
    ' Cargar icono desde un fichero en disco. Tamaño de 48x48
    Public Shared Function Icon_FromFile(ByVal fileName As String,
                                         Optional w As Integer = 48,
                                         Optional h As Integer = 48) As Icon
        Return New Icon(fileName, New Size(w, h))
    End Function
    Public Shared Function Bitmap_FromIcon(ByVal icon As Icon) As Bitmap
        Dim bmp As New Bitmap(icon.Width, icon.Height)
        Dim g As Graphics = Graphics.FromImage(bmp)
        g.DrawIcon(icon, 0, 0)
        g.Dispose()
        Return bmp
    End Function

    ''' <summary>
    ''' Asigna a un objeto Picture el icono indicado
    ''' </summary>
    ''' <param name="picBox">
    ''' El objeto Picture que tendrá el icono</param>
    ''' <param name="sPath">
    ''' El nombre del ejecutable o dll que contiene el icono</param>
    ''' <param name="indice">
    ''' El índice en base cero del icono</param>
    ''' <returns>
    ''' Devuelve un valor verdadero o falso según se haya asignado o no
    ''' </returns>
    ''' <remarks></remarks>
    Public Shared Function Icon_ToPictureBox(
                    ByVal picBox As PictureBox,
                    ByVal sPath As String,
                    ByVal indice As Integer) As Boolean

        Dim icono As Icon = Icon_PorIndice(sPath, indice)
        If icono IsNot Nothing Then
            picBox.Image = icono.ToBitmap
            Return True
        End If

        Return False
    End Function

    ''' <summary>
    ''' Devuelve un objeto Icon con el icono indicado
    ''' </summary>
    ''' <param name="sPath">
    ''' El nombre del ejecutable o dll que contiene el icono</param>
    ''' <param name="indice">
    ''' El índice en base cero del icono</param>
    ''' <returns>Un objeto Icon con el icono o Nothing si no existe</returns>
    ''' <remarks></remarks>
    Public Shared Function Icon_PorIndice(
                    ByVal sPath As String,
                    ByVal indice As Integer) As Icon
        Dim hInst As Integer
        Dim hIcon As IntPtr

        'hInst = GetClassWord(IntPtr.Zero, GCW_HMODULE)
        hInst = GetClassLong(IntPtr.Zero, GCW_HMODULE)
        hIcon = ExtractIcon(hInst, sPath, indice)
        If hIcon <> IntPtr.Zero Then
            Return Icon.FromHandle(hIcon)
        End If

        Return Nothing
    End Function

    ''' <summary>
    ''' Saber el número de iconos que tiene un exe o dll
    ''' </summary>
    ''' <param name="sPath">
    ''' El nombre del ejecutable o dll que contiene el icono</param>
    ''' <returns>Un valor entero con el número de iconos</returns>
    ''' <remarks></remarks>
    Public Shared Function Icon_DameTotalIconos(ByVal sPath As String) As Integer
        Return ExtractIconEx(sPath, -1, 0, 0, 0)
    End Function

    ''' <summary>
    ''' Devuelve el icono asociado al fichero indicado
    ''' </summary>
    ''' <param name="sFic">
    ''' Un fichero existente del que se quiere obtener el icono asociado
    ''' </param>
    ''' <returns>
    ''' Un objeto Icon con el icono o un valor nulo
    ''' si se produjo una excepción (o no tiene icono)
    ''' </returns>
    ''' <remarks></remarks>
    Public Shared Function Icon_DameDelFichero(ByVal sFic As String) As Icon
        Dim icono As Icon = Nothing
        Try
            icono = Icon.ExtractAssociatedIcon(sFic)
        Catch 'ex As Exception
        End Try

        Return icono
    End Function

    ''' <summary>
    ''' Devuelve un Bitmap de un icono.
    ''' Si el icono tiene un valor nulo, se devuelve un valor nulo.
    ''' </summary>
    ''' <param name="icono">El icono a convertir en Bitmap</param>
    ''' <returns>El Bitmap o un valor nulo</returns>
    ''' <remarks>
    ''' Esta función es para evitar una excepción si se usa directamente
    ''' y el icono es un valor nulo.
    ''' </remarks>
    Public Shared Function Icon_ToBitmap(ByVal icono As Icon) As Bitmap
        If icono Is Nothing Then
            Return Nothing
        Else
            Return icono.ToBitmap
        End If
    End Function
    '
    '' code for loading the icon into a PictureBox '
    'Dim theIcon As Icon = Icon_FromFile("C:\path\file.ico")
    'pbIcon.Image = theIcon.ToBitmap()
    'theIcon.Dispose()

    '' Método para dibujas icon en un Form. En coordenadas x=20, y=20
    Public Shared Sub Form_DibujaIcono(ByRef f As Form, pathIco As String)
        If f Is Nothing OrElse IO.File.Exists(pathIco) = False Then
            Exit Sub
        End If
        '
        Dim g As Graphics = Nothing
        Dim theIcon As Icon = Nothing
        Try
            g = f.CreateGraphics()
            theIcon = Icon_FromFile(pathIco)
            g.DrawIcon(theIcon, 20, 20)
        Catch ex As Exception
        Finally
            g.Dispose()
            g = Nothing
            theIcon.Dispose()
            theIcon = Nothing
        End Try
    End Sub

    Public Shared Function Image_ToByteArray(ByVal imageIn As System.Drawing.Image) As Byte()
        Dim ms As New IO.MemoryStream()
        imageIn.Save(ms, System.Drawing.Imaging.ImageFormat.Png)
        Return ms.ToArray()
    End Function


    Public Shared Function Image_ByteArrayToImage(ByVal byteArrayIn As Byte()) As Image
        Dim ms As New IO.MemoryStream(byteArrayIn)
        Dim returnImage As Image = Image.FromStream(ms)
        Return returnImage
    End Function

    Public Shared Sub Image_ByteArrayToFile(ByVal bytes() As Byte, queFi As String)
        Dim fs As System.IO.FileStream = New System.IO.FileStream(queFi, System.IO.FileMode.Create)
        ''Y escribimos en disco el array de bytes que conforman
        ''el fichero que sea (DWG, Word, Excel, etc.)
        fs.Write(bytes, 0, Convert.ToInt32(bytes.Length))
        fs.Close()
    End Sub
End Class
