Imports System
Imports System.IO
Imports System.IO.Compression

Partial Public Class Utiles
    Private Shared fiInv As String() = {""}
    Private Shared fiOff As String() = {""}

    Public Shared Async Sub Compress_FileNewZipAsync(FileIn As String, Optional openfolder As Boolean = False)
        If IO.File.Exists(FileIn) = False Then
            MsgBox("Not exist file:" & vbCrLf & FileIn)
            Exit Sub
        End If
        Dim FileOut As String = FileIn & ".zip"
        If IO.File.Exists(FileOut) Then
            Try
                IO.File.Delete(FileOut)
            Catch ex As Exception
                Exit Sub
            End Try
        End If
        '
        Using zipToOpen As FileStream = New FileStream(FileOut, FileMode.Create)
            Using archive As System.IO.Compression.ZipArchive = New System.IO.Compression.ZipArchive(zipToOpen, ZipArchiveMode.Update)
                Try
                    Dim fInfo As New FileInfo(FileIn)
                    ' FileStream
                    Using fStream As FileStream = fInfo.OpenRead
                        ' Crear entrada para el fichero
                        Dim queName As String = IO.Path.GetFileName(FileIn)
                        Dim readmeEntry As ZipArchiveEntry = archive.CreateEntry(queName, CompressionLevel.Optimal)
                        ' Abri Flujo de datos
                        Using zipStream As Stream = readmeEntry.Open
                            'fStream.CopyTo(zipStream)
                            Await fStream.CopyToAsync(zipStream)
                            'zipStream.Close()
                            'zipStream.Flush()
                        End Using
                        'fStream.Close()
                    End Using
                Catch ex As Exception
                    Debug.Print(ex.ToString)
                End Try
            End Using
            zipToOpen.Close()
        End Using
        If openfolder Then
            Process.Start(IO.Path.GetDirectoryName(FileIn))
        End If
    End Sub

    Public Shared Sub Compress_FileNewZipSync(FileIn As String, Optional openfolder As Boolean = False)
        If IO.File.Exists(FileIn) = False Then
            MsgBox("Not exist file:" & vbCrLf & FileIn)
            Exit Sub
        End If
        Dim FileOut As String = FileIn & ".zip"
        If IO.File.Exists(FileOut) Then
            Try
                IO.File.Delete(FileOut)
            Catch ex As Exception
                Exit Sub
            End Try
        End If
        '
        Using zipToOpen As FileStream = New FileStream(FileOut, FileMode.Create)
            Using archive As System.IO.Compression.ZipArchive = New System.IO.Compression.ZipArchive(zipToOpen, ZipArchiveMode.Update)
                Try
                    Dim fInfo As New FileInfo(FileIn)
                    ' FileStream
                    Using fStream As FileStream = fInfo.OpenRead
                        ' Crear entrada para el fichero
                        Dim queName As String = IO.Path.GetFileName(FileIn)
                        Dim readmeEntry As ZipArchiveEntry = archive.CreateEntry(queName, CompressionLevel.Optimal)
                        ' Abri Flujo de datos
                        Using zipStream As Stream = readmeEntry.Open
                            'fStream.CopyTo(zipStream)
                            Dim t As Task = fStream.CopyToAsync(zipStream)
                            t.Wait()
                            'zipStream.Close()
                        End Using
                        'fStream.Close()
                    End Using
                Catch ex As Exception
                    Debug.Print(ex.ToString)
                End Try
            End Using
            zipToOpen.Close()
        End Using
        If openfolder Then
            Process.Start(IO.Path.GetDirectoryName(FileIn))
        End If
    End Sub
    Public Shared Async Sub Compress_FilesExistZipAsync(FilesIn As Object(), ZipAppend As String, Optional OpenFolder As Boolean = False)
        Using zipToOpen As FileStream = New FileStream(ZipAppend, FileMode.Append)
            Dim ZipFolder As String = IO.Path.GetDirectoryName(ZipAppend)
            Using archive As System.IO.Compression.ZipArchive = New System.IO.Compression.ZipArchive(zipToOpen, ZipArchiveMode.Create)
                For Each f As Object In FilesIn
                    Dim fileIn As String = f.ToString
                    Try
                        Dim fInfo As New FileInfo(fileIn)
                        ' FileStream
                        Using fStream As FileStream = fInfo.OpenRead
                            ' Crear entrada para el fichero (Nombre de fichero solo o path relativo)
                            Dim queName As String = IO.Path.GetFileName(fileIn)
                            Dim readmeEntry As ZipArchiveEntry = archive.CreateEntry(queName, CompressionLevel.Optimal)
                            ' Abri Flujo de datos
                            Using zipStream As Stream = readmeEntry.Open
                                'fStream.CopyTo(zipStream)
                                Await fStream.CopyToAsync(zipStream)
                                'zipStream.Close()
                            End Using
                            'fStream.Close()
                        End Using
                        fInfo = Nothing
                    Catch ex As Exception
                        Debug.Print(ex.ToString)
                    End Try
                Next
            End Using
            zipToOpen.Close()
        End Using
        If OpenFolder Then
            Process.Start(IO.Path.GetDirectoryName(ZipAppend))
        End If
    End Sub
    Public Shared Sub Compress_FilesExistZipSync(FilesIn As Object(), ZipAppend As String, Optional OpenFolder As Boolean = False)
        Using zipToOpen As FileStream = New FileStream(ZipAppend, FileMode.Append)
            Dim ZipFolder As String = IO.Path.GetDirectoryName(ZipAppend)
            Using archive As System.IO.Compression.ZipArchive = New System.IO.Compression.ZipArchive(zipToOpen, ZipArchiveMode.Create)
                For Each fileIn As String In FilesIn
                    Try
                        Dim fInfo As New FileInfo(fileIn)
                        ' FileStream
                        Using fStream As FileStream = fInfo.OpenRead
                            ' Crear entrada para el fichero (Nombre de fichero solo o path relativo)
                            Dim queName As String = IO.Path.GetFileName(fileIn)
                            Dim readmeEntry As ZipArchiveEntry = archive.CreateEntry(queName, CompressionLevel.Optimal)
                            ' Abri Flujo de datos
                            Using zipStream As Stream = readmeEntry.Open
                                fStream.CopyTo(zipStream)
                                'Dim t As Task = fStream.CopyToAsync(zipStream)
                                't.Wait()
                            End Using
                        End Using
                        fInfo = Nothing
                    Catch ex As Exception
                        Debug.Print(ex.ToString)
                    End Try
                Next
            End Using
            zipToOpen.Close()
        End Using
        If OpenFolder Then
            Process.Start(IO.Path.GetDirectoryName(ZipAppend))
        End If
    End Sub
    Public Shared Async Sub Compress_FolderFilesExistZipAsync(FolderIn As String, Optional ZipAppend As String = "", Optional openfolder As Boolean = False)
        If IO.Directory.Exists(FolderIn) = False Then
            MsgBox("Not exist folder:" & vbCrLf & FolderIn)
        End If
        If ZipAppend = "" Then
            ZipAppend = IO.Path.Combine(IO.Path.GetDirectoryName(FolderIn), IO.Path.GetFileName(FolderIn) & ".zip")
        End If
        Using zipToOpen As FileStream = New FileStream(ZipAppend, FileMode.Append)
            Dim ZipFolder As String = IO.Path.GetDirectoryName(ZipAppend)
            Using archive As System.IO.Compression.ZipArchive = New System.IO.Compression.ZipArchive(zipToOpen, ZipArchiveMode.Create)
                For Each fileIn As String In IO.Directory.GetFiles(FolderIn, "*.*", SearchOption.AllDirectories)
                    Try
                        Dim fInfo As New FileInfo(fileIn)
                        ' FileStream
                        Using fStream As FileStream = fInfo.OpenRead
                            ' Crear entrada para el fichero (Nombre de fichero solo o path relativo)
                            Dim queName As String = fileIn.Replace(FolderIn & "\", "")
                            Dim readmeEntry As ZipArchiveEntry = archive.CreateEntry(queName, CompressionLevel.Optimal)
                            ' Abri Flujo de datos
                            Using zipStream As Stream = readmeEntry.Open
                                'fStream.CopyTo(zipStream)
                                Await fStream.CopyToAsync(zipStream)
                                'zipStream.Close()
                            End Using
                            'fStream.Close()
                        End Using
                        fInfo = Nothing
                    Catch ex As Exception
                        Debug.Print(ex.ToString)
                        Continue For
                    End Try
                Next
            End Using
            zipToOpen.Close()
        End Using
        If openfolder Then
            Process.Start(IO.Path.GetDirectoryName(ZipAppend))
        End If
    End Sub

    Public Shared Sub Compress_FolderFilesExistZipSync(FolderIn As String, Optional ZipAppend As String = "", Optional openfolder As Boolean = False)
        If IO.Directory.Exists(FolderIn) = False Then
            MsgBox("Not exist folder:" & vbCrLf & FolderIn)
        End If
        If ZipAppend = "" Then
            ZipAppend = IO.Path.Combine(IO.Path.GetDirectoryName(FolderIn), IO.Path.GetFileName(FolderIn) & ".zip")
        End If
        Using zipToOpen As FileStream = New FileStream(ZipAppend, FileMode.Append)
            Dim ZipFolder As String = IO.Path.GetDirectoryName(ZipAppend)
            Using archive As System.IO.Compression.ZipArchive = New System.IO.Compression.ZipArchive(zipToOpen, ZipArchiveMode.Create)
                For Each fileIn As String In IO.Directory.GetFiles(FolderIn, "*.*", SearchOption.AllDirectories)
                    If fileIn.Contains("OldVersions") Then Continue For
                    If fileIn.ToLower.EndsWith("bak") Then Continue For
                    If fileIn.ToLower.EndsWith("sv$") Then Continue For
                    Try
                        Dim fInfo As New FileInfo(fileIn)
                        ' FileStream
                        Using fStream As FileStream = fInfo.OpenRead
                            ' Crear entrada para el fichero (Nombre de fichero solo o path relativo)
                            Dim queName As String = fileIn.Replace(FolderIn & "\", "")
                            Dim readmeEntry As ZipArchiveEntry = archive.CreateEntry(queName, CompressionLevel.Optimal)
                            ' Abri Flujo de datos
                            Using zipStream As Stream = readmeEntry.Open
                                'fStream.CopyTo(zipStream)
                                Dim t As Task = fStream.CopyToAsync(zipStream)
                                t.Wait()
                                'zipStream.Close()
                            End Using
                            'fStream.Close()
                        End Using
                        fInfo = Nothing
                    Catch ex As Exception
                        Debug.Print(ex.ToString)
                        Continue For
                    End Try
                Next
            End Using
            zipToOpen.Close()
        End Using
        If openfolder Then
            Process.Start(IO.Path.GetDirectoryName(ZipAppend))
        End If
    End Sub
End Class
