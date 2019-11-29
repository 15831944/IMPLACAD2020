Imports Microsoft.VisualBasic
Imports System
Imports System.Runtime.InteropServices
Imports System.Collections.Generic
Imports System.ComponentModel
'Imports System.Data
Imports System.Diagnostics
Imports System.Drawing
Imports System.Text
Imports System.Windows.Automation
Imports System.Windows.Forms
Imports System.Windows.Input

Partial Public Class Utiles
    Public Shared Function ImageToPictureDisp(ByVal image As Image) As stdole.IPictureDisp
        Return AxHostConverter.ImageToPictureDisp(image)
    End Function

    Public Shared Function PictureDispToImage(ByVal pictureDisp As stdole.IPictureDisp) As Image
        Return AxHostConverter.PictureDispToImage(pictureDisp)
    End Function
End Class

Friend Class AxHostConverter
    Inherits System.Windows.Forms.AxHost

    Public Sub New()
        MyBase.New("")
    End Sub
    Public Shared Function ImageToPictureDisp(ByVal image As Image) As stdole.IPictureDisp
        Return CType(GetIPictureDispFromPicture(image), stdole.IPictureDisp)
    End Function

    Public Shared Function PictureDispToImage(ByVal pictureDisp As stdole.IPictureDisp) As Image
        Return GetPictureFromIPicture(pictureDisp)
    End Function
End Class
