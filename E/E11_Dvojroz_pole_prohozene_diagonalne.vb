Module E11_Dvojroz_pole_prohozene_diagonalne
    '11) Program Doc. Homoly: Dvojrozměrné pole: Naplňte zadáním pole 4x4. Poté ji zobrazí překlopenou kolem hlavní diagonály
    '(tedy prohodí indexy řádků a sloupců). Takže například místo pole:
    '1  2  3  4 
    '5  6  7  8 
    '9 10 11 12
    '13 14 15 16 
    '   bude pole
    '1  5  9  13
    '2  6  10 14
    '3  7  11 15
    '4  8  12 16
    Sub Mainx()
        MsgBox("Jsem E11")
        Dim pole_1(3, 3) As Byte
        Dim i As Byte, j As Byte, pocitadlo As Byte
        Dim txt_vystup As String, txt_prohozene As Byte

        i = 0
        j = 0
        pocitadlo = 1
        txt_vystup = ""
        txt_prohozene = ""

        'Naplní pole
        Do While i < 4
            Do While j < 4
                pole_1(i, j) = pocitadlo
                pocitadlo += 1
                txt_vystup += Str(pole_1(i, j)) + Chr(9)
                j += 1
            Loop
            j = 0
            txt_vystup += Chr(10)
            i += 1
        Loop

        'Výpis obsahu pole
        MsgBox(txt_vystup)
    End Sub
End Module
