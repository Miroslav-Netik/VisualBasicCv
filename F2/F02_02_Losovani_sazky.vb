Module F02_02_Losovani_sazky
    '2) Stroj na losování Sazky se porouchal. Tahal sice z osudí čísla 1 až 49, ale některá opakovaně. Vytvořte funkci,
    'které se předá pole takovýchto čísel, funkce pole vrátí očištěné od druhých a dalších výskytů opakujících se čísel
    '(pro jednoduchost stačí, aby na jejich místě dosadila nulu). Nevadí, že tedy může být ve výsledku méně čísel než požadovaných pět.
    Sub Mainx()
        Dim pole_tazenych() As Byte = {}
        Dim pole_upravene() As Byte = {}
        Dim tazene_cislo As Byte, rozsah_max As Byte, rozsah_min As Byte
        Dim pocet_opakovani As Byte
        Dim i As Byte, i_2 As Byte
        Dim txt_vystup As String, txt_upraveny_vystup As String

        txt_upraveny_vystup = ""
        txt_vystup = ""
        rozsah_max = 49
        rozsah_min = 1
        pocet_opakovani = 5

        Randomize()

        For i = 1 To pocet_opakovani + 1
            tazene_cislo = Int(Rnd() * (rozsah_max - rozsah_min + 1) + rozsah_min)
            txt_vystup += Str(tazene_cislo) + ","
            ReDim Preserve pole_tazenych(i - 1)
            pole_tazenych(i - 1) = tazene_cislo
        Next

        pole_upravene = Ocisti(pole_tazenych)
        For i_2 = 0 To pocet_opakovani
            txt_upraveny_vystup += Str(pole_upravene(i_2)) + ","
        Next

        MsgBox(txt_vystup + Chr(10) + txt_upraveny_vystup)
    End Sub

    Function Ocisti(vstupni_pole() As Byte) As Byte()
        Dim vystupni_pole() As Byte = {}
        Dim i As Byte, i_2 As Byte

        ReDim Preserve vystupni_pole(UBound(vstupni_pole))

        'Naplním vystupní pole vstupním, aby se se vstupním nic nestalo
        For i = 0 To UBound(vstupni_pole)
            vystupni_pole(i) = vstupni_pole(i)
        Next

        'Pak postupně projíždím výstupní
        For i = 0 To UBound(vystupni_pole)
            For i_2 = i + 1 To UBound(vystupni_pole)
                If vystupni_pole(i) = vystupni_pole(i_2) Then   'A nahrazuju opakující se čísla nulou
                    vystupni_pole(i_2) = 0
                End If
            Next
        Next

        Return vystupni_pole
    End Function
End Module
