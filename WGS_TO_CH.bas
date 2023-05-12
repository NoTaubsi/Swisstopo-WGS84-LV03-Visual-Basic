Attribute VB_Name = "WGS_TO_CH"
Function Convert_WGS84_DD_to_Sexagesimal_Seconds(latitude As Double) As Double

    'Berechnung der sexagesimalen Sekunden
    Dim degrees As Double: degrees = Int(latitude)
    Dim minutes As Double: minutes = Int((latitude - degrees) * 60)
    Dim seconds As Double: seconds = ((latitude - degrees) * 3600) - (minutes * 60)
    
    'Ergebnis zurueckgeben
    Convert_WGS84_DD_to_Sexagesimal_Seconds = (degrees * 3600) + (minutes * 60) + seconds

End Function

Function WGStoCHy(X As Double, Y As Double) As Double

    'Berechnung der sexagesimalen Sekunden
    Dim XasSEX As Double: XasSEX = Convert_WGS84_DD_to_Sexagesimal_Seconds(X)
    Dim YasSEX As Double: YasSEX = Convert_WGS84_DD_to_Sexagesimal_Seconds(Y)
    
    Dim X_aux As Double: X_aux = (XasSEX - 169028.66) / 10000
    Dim Y_aux As Double: Y_aux = (YasSEX - 26782.5) / 10000
    
    'Process X
    Dim resultY As Double: resultY = 200147.07 + 308807.95 * X_aux + 3745.25 * (Y_aux ^ 2) + 76.63 * (X_aux ^ 2) - 194.56 * (Y_aux ^ 2) * X_aux + 119.79 * (X_aux ^ 3)
    
    'Ergebnis zurueckgeben
    WGStoCHy = resultY
    
End Function

Function WGStoCHx(X As Double, Y As Double) As Double

    'Berechnung der sexagesimalen Sekunden
    Dim XasSEX As Double: XasSEX = Convert_WGS84_DD_to_Sexagesimal_Seconds(X)
    Dim YasSEX As Double: YasSEX = Convert_WGS84_DD_to_Sexagesimal_Seconds(Y)
    
    Dim X_aux As Double: X_aux = (XasSEX - 169028.66) / 10000
    Dim Y_aux As Double: Y_aux = (YasSEX - 26782.5) / 10000
    
    'Process X
    Dim resultX As Double: resultX = 600072.37 + 211455.93 * Y_aux - 10938.51 * Y_aux * X_aux - 0.36 * Y_aux * (X_aux ^ 2) - 44.54 * (Y_aux ^ 3)
    
    'Ergebnis zurueckgeben
    WGStoCHx = resultX
    
End Function




