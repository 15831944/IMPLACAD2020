' Importar referencia System.IO.Compression y System.IO.Compression.FileSystem
Imports System.Windows.Forms

Public Module Module1

    Private directoryPath As String = My.Computer.FileSystem.SpecialDirectories.Temp

    'Public Sub Main()
    '    Dim directorySelected As New DirectoryInfo(directoryPath)
    '    Compress(directorySelected)

    '    For Each fileToDecompress As FileInfo In directorySelected.GetFiles("*.gz")
    '        Decompress(fileToDecompress)
    '    Next
    'End Sub

#Region "FICHEROS_GZ"

    Public Sub Compress(directorySelected As DirectoryInfo)
        For Each fileToCompress As FileInfo In directorySelected.GetFiles()
            Using originalFileStream As FileStream = fileToCompress.OpenRead()
                If (File.GetAttributes(fileToCompress.FullName) And FileAttributes.Hidden) <> FileAttributes.Hidden And fileToCompress.Extension <> ".gz" Then
                    Using compressedFileStream As FileStream = File.Create(fileToCompress.FullName & ".gz")
                        Using compressionStream As New GZipStream(compressedFileStream, CompressionMode.Compress)

                            originalFileStream.CopyTo(compressionStream)
                        End Using
                    End Using
                    Dim info As New FileInfo(directoryPath & "\" & fileToCompress.Name & ".gz")
                    Console.WriteLine("Compressed {0} from {1} to {2} bytes.", fileToCompress.Name,
                                      fileToCompress.Length.ToString(), info.Length.ToString())

                End If
            End Using
        Next
    End Sub

    Private Sub Decompress(ByVal fileToDecompress As FileInfo)
        Using originalFileStream As FileStream = fileToDecompress.OpenRead()
            Dim currentFileName As String = fileToDecompress.FullName
            Dim newFileName = currentFileName.Remove(currentFileName.Length - fileToDecompress.Extension.Length)

            Using decompressedFileStream As FileStream = File.Create(newFileName)
                Using decompressionStream As GZipStream = New GZipStream(originalFileStream, CompressionMode.Decompress)
                    decompressionStream.CopyTo(decompressedFileStream)
                    Console.WriteLine("Decompressed: {0}", fileToDecompress.Name)
                End Using
            End Using
        End Using
    End Sub

#End Region

#Region "FICHEROS_ZIP"

    'Sub Main()
    '    Dim startPath As String = "c:\example\start"
    '    Dim zipPath As String = "c:\example\result.zip"
    '    Dim extractPath As String = "c:\example\extract"

    '    ZipFile.CreateFromDirectory(startPath, zipPath)

    '    ZipFile.ExtractToDirectory(zipPath, extractPath)
    'End Sub

    Private Sub Ficheros_ComprimeZIP(lFicheros As List(Of String), zFile As String,
                                     Optional quePb As ProgressBar = Nothing,
                                     Optional abrircarpeta As Boolean = True)
        If quePb IsNot Nothing Then
            quePb.Value = 0 : quePb.Maximum = lFicheros.Count
        End If
        '
        Dim zFolder As String = IO.Path.GetDirectoryName(zFile)
        '
        Using zipToOpen As FileStream = New FileStream(zFile, FileMode.Create)
            Using archive As System.IO.Compression.ZipArchive = New System.IO.Compression.ZipArchive(zipToOpen, ZipArchiveMode.Create)
                For Each queFi As String In lFicheros
                    If quePb IsNot Nothing Then
                        If quePb.Value <= quePb.Maximum Then quePb.Value += 1
                    End If
                    ' FileInfo
                    Dim fInfo As New FileInfo(queFi)
                    ' FileStream
                    Dim fStream As FileStream = fInfo.OpenRead
                    ' Crear entrada para el fichero
                    Dim readmeEntry As ZipArchiveEntry = archive.CreateEntry(fStream.Name, CompressionLevel.Optimal)
                    ' Abri Flujo de datos
                    Dim zipStream As Stream = readmeEntry.Open
                    fStream.CopyTo(zipStream)
                    zipStream.Close()
                    fStream.Close()
                Next
            End Using
            zipToOpen.Close()
        End Using
        If abrircarpeta Then Process.Start(zFolder)
    End Sub

    Public Sub Directorio_ComprimeZIP(queOrigenDIR As String, queDestinoZIP As String)
        System.IO.Compression.ZipFile.CreateFromDirectory(queOrigenDIR, queDestinoZIP)
    End Sub

    Public Sub DescomprimeZIP(queOrigenZIP As String, Optional queDestinoDIR As String = "")
        If queDestinoDIR = "" Then
            queDestinoDIR = IO.Path.Combine(IO.Path.GetDirectoryName(queOrigenZIP), IO.Path.GetFileNameWithoutExtension(queOrigenZIP))
        End If
        '
        If IO.Directory.Exists(queDestinoDIR) Then
            ' Antes borraremos todos los directorios y los ficheros que contengan.
            For Each queD As String In IO.Directory.GetDirectories(queDestinoDIR)
                IO.Directory.Delete(queD, True)
            Next
        End If
        System.IO.Compression.ZipFile.ExtractToDirectory(queOrigenZIP, queDestinoDIR)
    End Sub

    ''
    Public Function FicheroZIP_DameCantidadFicheros(quefi As String, tipo As QueTipo) As Integer
        Dim cantidad As Integer = 0
        ''
        Using archivo As System.IO.Compression.ZipArchive = System.IO.Compression.ZipFile.OpenRead(quefi)
            For Each entry As ZipArchiveEntry In archivo.Entries
                'If entry.FullName.EndsWith(".txt", StringComparison.OrdinalIgnoreCase) Then
                'entry.ExtractToFile(Path.Combine(extractPath, entry.FullName))
                If tipo = QueTipo.FicherosTodos Then
                    cantidad += 1
                ElseIf tipo = QueTipo.FicherosDWG And entry.FullName.EndsWith(".dwg", StringComparison.OrdinalIgnoreCase) Then
                    cantidad += 1
                ElseIf tipo = QueTipo.FicherosPNG And entry.FullName.EndsWith(".png", StringComparison.OrdinalIgnoreCase) Then
                    cantidad += 1
                ElseIf tipo = QueTipo.Carpetas And entry.FullName.EndsWith("/", StringComparison.OrdinalIgnoreCase) Then
                    cantidad += 1
                End If
                'End If
            Next
        End Using
        ''
        Return cantidad
    End Function

    ''
    Public Enum QueTipo As Integer
        FicherosTodos = 0
        FicherosDWG = 1
        FicherosPNG = 2
        Carpetas = 3
    End Enum

#End Region

End Module