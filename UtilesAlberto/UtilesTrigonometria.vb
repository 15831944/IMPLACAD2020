Partial Public Class Utiles


    Public Shared Function Dist_GetBetweenPoints(
                                 ByVal x1 As Double, ByVal y1 As Double, ByVal z1 As Double,
                                 ByVal x2 As Double, ByVal y2 As Double, ByVal z2 As Double) As Double
        Return (((x2 - x1) ^ 2) + ((y2 - y1) ^ 2) + ((z2 - z1) ^ 2)) ^ 0.5
    End Function

    Public Shared Function DegreesToRadians(ByVal gra As Double) As Double
        Return (gra * Math.PI) / 180
    End Function

    Public Shared Function RadiansToDegrees(ByVal rad As Double) As Double
        Return (rad * 180) / Math.PI
    End Function
End Class
