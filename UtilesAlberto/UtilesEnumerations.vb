Partial Public Class Utiles

    Public Shared Function Enumeration_Read(queEnum As [Enum], Optional message As Boolean = False) As String
        Dim resultado As String = ""
        Dim valores As String() = [Enum].GetValues(queEnum.[GetType])
        resultado = String.Join(vbCrLf, valores)
        'For Each valor As String In valores

        'Next
        If message = True Then
            MsgBox(resultado)
        End If
        Return resultado
    End Function
    'Public Enum EUnits
    '    su      ' Without unit (ES)
    '    ul      ' Without unit  (EN)
    '    mm      ' millimeters
    '    cm      ' centimeters
    '    m       ' meters
    '    [in]      ' inches
    '    gr      ' degrees
    '    rad     ' radians
    '    none    ' For texts. They do not carry units
    '    Texto   ' For texts. They do not carry units
    '    [Object]    ' Cualquier objeto
    '    kg      ' Kilos
    'End Enum

    'Public Enum EEstructureFolder
    '    DirPrincipal = FileIO.SearchOption.SearchTopLevelOnly
    '    DirTodos = FileIO.SearchOption.SearchAllSubDirectories
    'End Enum
End Class
