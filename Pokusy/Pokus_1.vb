Module Pokus_1
    Sub Mainx()
        Dim vstup As Byte, soucet As Byte
        Dim vystup As Byte, pocitadlo As String
        Dim prumer As Single
        Dim vysledek As Double
        Dim vystup_txt As String

        soucet = 0
        pocitadlo = 0

        vstup = InputBox("Napiš číslo: ")
        Do While vstup <> 0
            soucet = soucet + vstup
            pocitadlo = pocitadlo + 1
            vstup = InputBox("Napiš číslo: ")
        Loop

        prumer = soucet / pocitadlo
        'vysledek = Math.Round(prumer)
        vysledek = Fix(prumer + 0.5)
        MsgBox(vysledek)
    End Sub

End Module
