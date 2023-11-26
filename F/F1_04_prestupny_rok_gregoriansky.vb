Module F1_04_prestupny_rok_gregoriansky
    '4) Napište funkci (a pak předělejte na subrutinu), které se předá rok (čtyřciferný) a která vrací logickou hodnotu "True",
    'pokud je rok přestupný (přepracujte z Juliánského kalendáře z BasicCv.doc na Gregoriánský).  Řešte dvěma způsoby:
    'dvěma podmínkami (či if-else) i jednou složenou podmínkou.
    'Potřebná teorie: Podle gregoriánského kalendáře jsou přestupné roky ty, které jsou dělitelné čtyřmi.
    'Ale roky dělitelné stem jsou přestupné jenom tehdy, jsou-li dělitelné také 400. Přestupnými roky jsou proto
    'například roky 1600, 2000, 2400 apod., zatímco roky 1700, 1800, 1900, 2100 atd. přestupné nejsou. (zdroj: wikipedia)
    Sub Mainx()
        MsgBox("Jsem F1-04")

        Dim zadany_rok As Integer
        Dim vystup_rok As Boolean

        zadany_rok = InputBox("Zadej rok (čtyřciferný): ")

        If F_prestupny_rok(zadany_rok) Then
            MsgBox("Rok" + Str(zadany_rok) + " je přestupný.")
        Else
            MsgBox("Rok" + Str(zadany_rok) + " není přestupný.")
        End If

        S_prestupny_rok(zadany_rok, vystup_rok)
        If vystup_rok Then
            MsgBox("Rok" + Str(zadany_rok) + " je přestupný.")
        Else
            MsgBox("Rok" + Str(zadany_rok) + " není přestupný.")
        End If
    End Sub
    Sub S_prestupny_rok(rok As Integer, ByRef vystup As Boolean)
        If (rok Mod 4 = 0 And rok Mod 100 <> 0) Or rok Mod 400 = 0 Then
            vystup = True
        End If
    End Sub
    Function F_prestupny_rok(rok As Integer) As Boolean
        If rok Mod 4 = 0 And rok Mod 400 = 0 Then
            F_prestupny_rok = True
        End If
        Return F_prestupny_rok
    End Function
End Module
