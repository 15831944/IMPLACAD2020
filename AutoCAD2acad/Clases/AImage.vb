Imports System.Diagnostics
Imports System.Collections
Imports System.Windows.Forms
Imports System.Windows.Media.Imaging
Imports System.Drawing
Imports System.IO
Imports System.Runtime.InteropServices


Imports Autodesk.AutoCAD.Interop
Imports Autodesk.AutoCAD.Interop.Common
Imports Autodesk.AutoCAD.Runtime
Imports appS = Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.EditorInput

Namespace A2acad
    Partial Public Class A2acad
        Public Function DameBitmapDWG_Thumbnail(ByVal archivoDWG As String) As System.Drawing.Bitmap
            Dim fiPNG As String = IO.Path.ChangeExtension(archivoDWG, ".png")
            Dim fiJPG As String = IO.Path.ChangeExtension(archivoDWG, ".jpg")
            If IO.File.Exists(fiPNG) Then
                Return Drawing.Image.FromFile(fiPNG)
                Exit Function
            ElseIf IO.File.Exists(fiJPG) Then
                Return Drawing.Image.FromFile(fiJPG)
                Exit Function
            End If
            '
            Dim imagen As System.Drawing.Bitmap
            imagen = My.Resources.imgparabloquessinimage
            Try
                If IO.File.Exists(archivoDWG) Then
                    Using tempDB As New Autodesk.AutoCAD.DatabaseServices.Database(buildDefaultDrawing:=False, noDocument:=True)
                        tempDB.ReadDwgFile(archivoDWG, FileOpenMode.OpenForReadAndReadShare,
                                   allowCPConversion:=False, password:="")  ' System.IO.FileShare.Read
                        imagen = tempDB.ThumbnailBitmap
                    End Using
                End If
            Catch ex As System.Exception
                '//Por si el archivo .dwg esta corrupto y no se puede leer.
                imagen = My.Resources.imgparabloquessinimage
            End Try
            Return (ImagenCambiaFondo(imagen, System.Drawing.Color.White))
        End Function
        ''
        Function DameBitmapDWG_IO(fileName As String) As System.Drawing.Bitmap
            Dim fiPNG As String = IO.Path.ChangeExtension(fileName, ".png")
            Dim fiJPG As String = IO.Path.ChangeExtension(fileName, ".jpg")
            If IO.File.Exists(fiPNG) Then
                Return Drawing.Image.FromFile(fiPNG)
                Exit Function
            ElseIf IO.File.Exists(fiJPG) Then
                Return Drawing.Image.FromFile(fiJPG)
                Exit Function
            End If
            '
            Using fs As New FileStream(fileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
                Using br As New BinaryReader(fs)
                    fs.Seek(&HD, SeekOrigin.Begin)
                    fs.Seek(&H14 + br.ReadInt32(), SeekOrigin.Begin)
                    Dim bytCnt As Byte = br.ReadByte()
                    If bytCnt <= 1 Then
                        Return Nothing
                    End If
                    Dim imageHeaderStart As Integer
                    Dim imageHeaderSize As Integer
                    Dim imageCode As Byte
                    For i As Short = 1 To bytCnt
                        imageCode = br.ReadByte()
                        imageHeaderStart = br.ReadInt32()
                        imageHeaderSize = br.ReadInt32()
                        If imageCode = 2 Then ' BMP Preview (2012 file format)
                            ' BITMAPINFOHEADER (40 bytes)
                            br.ReadBytes(&HE)
                            'biSize, biWidth, biHeight, biPlanes
                            Dim biBitCount As UShort = br.ReadUInt16()
                            br.ReadBytes(4)
                            'biCompression
                            Dim biSizeImage As UInteger = br.ReadUInt32()
                            'br.ReadBytes(0x10); //biXPelsPerMeter, biYPelsPerMeter, biClrUsed, biClrImportant
                            '-----------------------------------------------------
                            fs.Seek(imageHeaderStart, SeekOrigin.Begin)
                            Dim bitmapBuffer As Byte() = br.ReadBytes(imageHeaderSize)
                            Dim colorTableSize As UInteger = CUInt(Math.Truncate(If((biBitCount < 9), 4 * Math.Pow(2, biBitCount), 0)))
                            Using ms As New MemoryStream()
                                Using bw As New BinaryWriter(ms)
                                    bw.Write(CUShort(&H4D42))
                                    bw.Write(54UI + colorTableSize + biSizeImage)
                                    bw.Write(New UShort())
                                    bw.Write(New UShort())
                                    bw.Write(54UI + colorTableSize)
                                    bw.Write(bitmapBuffer)
                                    Return New Bitmap(ms)
                                End Using
                            End Using
                        ElseIf imageCode = 6 Then ' PNG Preview (2013 file format)
                            fs.Seek(imageHeaderStart, SeekOrigin.Begin)
                            Using ms As New MemoryStream
                                fs.CopyTo(ms, imageHeaderStart)
                                Dim img = System.Drawing.Image.FromStream(ms)
                                Return img
                            End Using
                        ElseIf imageCode = 3 Then
                            Return Nothing
                        End If
                    Next
                End Using
            End Using
            Return Nothing
        End Function
        Private Function ImagenCambiaFondo(ByVal imagen As Bitmap, colordefondo As System.Drawing.Color) As Bitmap
            If imagen Is Nothing Then
                Return My.Resources.imgparabloquessinimage
                Exit Function
            End If
            Dim imagenConFondoOscuro As Bitmap
            imagenConFondoOscuro = New Bitmap(imagen)
            Dim graph As Drawing.Graphics
            graph = Drawing.Graphics.FromImage(imagenConFondoOscuro)
            graph.Clear(colordefondo)
            imagen.MakeTransparent()
            graph.DrawImage(imagen, New Point(0, 0))
            Return (imagenConFondoOscuro)
        End Function
        Public Function ConvertToBitmapImage(ByVal image As System.Drawing.Bitmap) As System.Windows.Media.Imaging.BitmapImage
            Dim png As System.Windows.Media.Imaging.BitmapImage = New System.Windows.Media.Imaging.BitmapImage()
            If image IsNot Nothing Then
                Dim stream As System.IO.MemoryStream = New System.IO.MemoryStream()
                image.Save(stream, System.Drawing.Imaging.ImageFormat.Png)
                png.BeginInit()
                png.StreamSource = stream
                png.EndInit()
            End If
            '
            Return png
        End Function

        Public Function ConvertToBitmapImage(ByVal image As System.Drawing.Image) As System.Windows.Media.Imaging.BitmapImage
            Dim png As BitmapImage = New BitmapImage()
            If image IsNot Nothing Then
                Dim stream As System.IO.MemoryStream = New System.IO.MemoryStream()
                image.Save(stream, System.Drawing.Imaging.ImageFormat.Png)
                png.BeginInit()
                png.StreamSource = stream
                png.EndInit()
            End If
            Return png
        End Function
        '
        <System.Security.SuppressUnmanagedCodeSecurity(), DllImport("acdb18.dll", CallingConvention:=CallingConvention.Cdecl, EntryPoint:="?acdbGetPreviewBitmap@@YAPEAUtagBITMAPINFO@@PEB_W@Z")>
        Private Shared Function acdbGetPreviewBitmap(<MarshalAs(UnmanagedType.LPWStr)> ByVal sFilename As String) As IntPtr
        End Function

        Private Function GetBitmapFromDwg(ByVal filename As String) As System.Drawing.Bitmap
            Dim fiPNG As String = IO.Path.ChangeExtension(filename, ".png")
            Dim fiJPG As String = IO.Path.ChangeExtension(filename, ".jpg")
            If IO.File.Exists(fiPNG) Then
                Return Drawing.Image.FromFile(fiPNG)
                Exit Function
            ElseIf IO.File.Exists(fiJPG) Then
                Return Drawing.Image.FromFile(fiJPG)
                Exit Function
            End If
            '
            Dim bmpInfo As IntPtr = acdbGetPreviewBitmap(filename)
            Return Autodesk.AutoCAD.Runtime.Marshaler.BitmapInfoToBitmap(bmpInfo)
        End Function
        '
        '***** Ejemplo de imagen previa de un DWG desde un fichero
        'Dim viewDwg As New ViewDWG()
        'PictureBox1.Image = viewDwg.GetDwgImage("d:\s.dwg")
        '*****
        Public Function DameImageDWGFichero(queFiDWG As String) As System.Drawing.Image
            Dim fiPNG As String = IO.Path.ChangeExtension(queFiDWG, ".png")
            Dim fiJPG As String = IO.Path.ChangeExtension(queFiDWG, ".jpg")
            If IO.File.Exists(fiPNG) Then
                Return Drawing.Image.FromFile(fiPNG)
                Exit Function
            ElseIf IO.File.Exists(fiJPG) Then
                Return Drawing.Image.FromFile(fiJPG)
                Exit Function
            End If
            '
            Dim viewDwg As New ViewDWG()
            Return viewDwg.GetDwgImage(queFiDWG)
        End Function
        Private Class ViewDWG
            Private Structure BITMAPFILEHEADER
                Public bfType As Short
                Public bfSize As Integer
                Public bfReserved1 As Short
                Public bfReserved2 As Short
                Public bfOffBits As Integer
            End Structure
            Public Function GetDwgImage(ByVal FileName As String) As System.Drawing.Image
                If Not (File.Exists(FileName)) Then
                    Throw New FileNotFoundException("the file didn't find")
                End If
                'declare a filestream to read the DWG file
                Dim DwgF As FileStream
                'the position of file description block
                Dim PosSentinel As Integer
                'Binary Reader
                Dim br As BinaryReader
                'the thumbnail farmat
                Dim TypePreview As Integer
                'the thumbnail postion
                Dim PosBMP As Integer
                'the thumbnail size
                Dim LenBMP As Integer
                'the thumbnail depth
                Dim biBitCount As Short
                'BMP file header，because DWG doesn't include the header，we need add it by ourself
                Dim biH As BITMAPFILEHEADER
                'BMP file body in the DWG file
                Dim BMPInfo As Byte()
                'to store the memory stream
                Dim BMPF As New MemoryStream()
                'Binary Writer
                Dim bmpr As New BinaryWriter(BMPF)
                Dim myImg As System.Drawing.Image = Nothing

                Try
                    'DWG file stream
                    DwgF = New FileStream(FileName, FileMode.Open, FileAccess.Read)
                    br = New BinaryReader(DwgF)
                    'read from the thirteenth byte
                    DwgF.Seek(13, SeekOrigin.Begin)
                    'read the thumbnail description block position
                    PosSentinel = br.ReadInt32()
                    'read from the thirty-first byte of the thumbnail description block
                    DwgF.Seek(PosSentinel + 30, SeekOrigin.Begin)
                    'the thirty-first byte represent the thumbnail picture format, 2 means BMP, 3 means WMF
                    TypePreview = br.ReadByte()

                    If TypePreview = 1 Then
                    ElseIf TypePreview = 2 OrElse TypePreview = 3 Then
                        'the BMP block postion saved by DWG file
                        PosBMP = br.ReadInt32()
                        'the BMP block size
                        LenBMP = br.ReadInt32()
                        'move to the BMP block
                        DwgF.Seek(PosBMP + 14, SeekOrigin.Begin)
                        'read the btye depth
                        biBitCount = br.ReadInt16()
                        'read all BMP content
                        DwgF.Seek(PosBMP, SeekOrigin.Begin)
                        'the BMP with no header
                        BMPInfo = br.ReadBytes(LenBMP)
                        br.Close()
                        DwgF.Close()
                        biH.bfType = 19778
                        'build the header
                        If biBitCount < 9 Then
                            biH.bfSize = 54 + 4 * CInt(Math.Truncate(Math.Pow(2, biBitCount))) + LenBMP
                        Else
                            biH.bfSize = 54 + LenBMP
                        End If
                        'reserved byte
                        biH.bfReserved1 = 0
                        'reserved byte
                        biH.bfReserved2 = 0
                        'BMP data offset
                        'write BMP header
                        biH.bfOffBits = 14 + 40 + 1024
                        'file type
                        bmpr.Write(biH.bfType)
                        'file size
                        bmpr.Write(biH.bfSize)
                        '0
                        bmpr.Write(biH.bfReserved1)
                        '0
                        bmpr.Write(biH.bfReserved2)
                        'offset
                        bmpr.Write(biH.bfOffBits)
                        'write BMP file
                        bmpr.Write(BMPInfo)
                        'move the pointer to the beginning
                        BMPF.Seek(0, SeekOrigin.Begin)
                        'create the BMP file
                        myImg = System.Drawing.Image.FromStream(BMPF)
                        bmpr.Close()
                        BMPF.Close()
                    End If
                    Return myImg
                Catch ex As Exception
                    Throw New Exception(ex.Message)
                End Try
            End Function
        End Class
        '
        Public Function Imagen_PreviaDeBloque(name As String) As System.Drawing.Image
            Dim resultado As System.Drawing.Image = Nothing
            Dim doc As appS.Document = appS.Application.DocumentManager.MdiActiveDocument
            Dim ed As Editor = doc.Editor
            Dim db As Database = doc.Database
            Using tr As Transaction = doc.TransactionManager.StartTransaction()
                Dim bt As BlockTable = CType(tr.GetObject(db.BlockTableId, OpenMode.ForRead), BlockTable)
                For Each blkId As ObjectId In bt
                    Dim blk As BlockTableRecord = CType(tr.GetObject(blkId, OpenMode.ForRead), BlockTableRecord)
                    If blk.IsLayout OrElse blk.IsAnonymous Then Continue For

                    If blk.PreviewIcon IsNot Nothing And (blk.Name = name Or blk.Name.ToUpper = name.ToUpper) Then
                        resultado = blk.PreviewIcon
                        Exit For
                    End If
                Next
                tr.Abort()
            End Using
            doc = Nothing
            ed = Nothing
            db = Nothing
            Return resultado
        End Function
        '
        Public Function Imagen_PreviaDeBloque(id As IntPtr) As System.Drawing.Image
            Dim resultado As System.Drawing.Image = Nothing
            Dim oId As New ObjectId(id)
            Dim doc As appS.Document = appS.Application.DocumentManager.MdiActiveDocument
            Dim ed As Editor = doc.Editor
            Dim db As Database = doc.Database
            Using tr As Transaction = doc.TransactionManager.StartTransaction()
                Dim bt As BlockTable = CType(tr.GetObject(db.BlockTableId, OpenMode.ForRead), BlockTable)
                For Each blkId As ObjectId In bt
                    Dim blk As BlockTableRecord = CType(tr.GetObject(blkId, OpenMode.ForRead), BlockTableRecord)
                    If blk.IsLayout OrElse blk.IsAnonymous Then Continue For

                    If blk.PreviewIcon IsNot Nothing And blk.ObjectId.Equals(oId) Then
                        resultado = blk.PreviewIcon
                        Exit For
                    End If
                Next
                tr.Abort()
            End Using
            oId = Nothing
            doc = Nothing
            ed = Nothing
            db = Nothing
            Return resultado
        End Function
        Public Function Imagen_PreviaDeBloque(blkId As ObjectId) As System.Drawing.Image
            Return Imagen_PreviaDeBloque(blkId.OldIdPtr)
        End Function
    End Class
End Namespace