Module C15_dve_stejna_nahodna_vedle_sebe_atd
    Sub Mainx()
        '15) Vygenerujte 100 náhodných celých čísel od 1 do 5. Zjistěte,
        'zda jsou někdy dvě stejná vedle sebe, která to jsou a jaké je jejich pořadové číslo.
        MsgBox("Jsem C15")
        Dim generovana_cisla As Integer
        Dim i As Byte, prvni_porovnavane_cislo As Byte, druhe_porovnavane_cislo As Byte
        Dim vystup As String, vystup_generovanych As String

        Randomize()
        prvni_porovnavane_cislo = 0
        druhe_porovnavane_cislo = 0

        For i = 1 To 100
            generovana_cisla = Fix(Rnd() * (5 - 1 + 1) + 1)
            prvni_porovnavane_cislo = generovana_cisla
            If prvni_porovnavane_cislo = druhe_porovnavane_cislo Then
                vystup = vystup + Str(generovana_cisla) + " jsou vedle sebe na pozicích" + Str(i - 1) + Str(i) + Chr(10)
            Else
                druhe_porovnavane_cislo = prvni_porovnavane_cislo
            End If

            vystup_generovanych = vystup_generovanych + Str(generovana_cisla) + " / "

        Next
        MsgBox(Str(i))
        MsgBox(vystup_generovanych + Chr(10) + vystup)
    End Sub
End Module
