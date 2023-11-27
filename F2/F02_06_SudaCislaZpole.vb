Module F02_06_SudaCislaZpole
    '6) Sestavte funkci SudaCislaZpole, která má jediný parametr: pole celých čísel. Výsledkem funkce je jednorozměrné celočíselné pole,
    've kterém jsou jen sudá čísla obsažená v poli parametru. Funkci ověřte na zadání: sestavte program, který načte pole celých čísel
    'ukončených nulou, která už k číslům nepatří, a pak zobrazí všechna ze zadaných čísel, která jsou sudá. 
    Sub Mainx()
        Dim pole_cisel() As Integer = {}
        Dim pole_sudych() As Integer = {}
        Dim vstup_cislo As Integer
        Dim i As Integer, j As Integer
        Dim txt_vystup As String = ""
        Dim txt_vystup_suda As String = ""

        vstup_cislo = InputBox("Zadej celé číslo (ukonči 0) : ")

        'Načte vstup do pole a do testoveho vystupu (pro kontrolu)
        Do While vstup_cislo <> 0
            ReDim Preserve pole_cisel(i)
            pole_cisel(i) = vstup_cislo
            txt_vystup += Str(pole_cisel(i)) + ","
            i += 1
            vstup_cislo = InputBox("Zadej číslo (ukonči 0) : ")
        Loop

        'Načte výsledek z funkce do pole_sudych a uloží hodnoty do txt vystupu
        pole_sudych = SudaCislaZpole(pole_cisel)
        For j = 0 To UBound(pole_sudych)
            txt_vystup_suda += Str(pole_sudych(j)) + ","
        Next

        'Výpis na obrazovku
        MsgBox(txt_vystup + Chr(10) + txt_vystup_suda)
    End Sub

    Function SudaCislaZpole(pole() As Integer) As Integer()
        'Ze zadaného pole celých čísel vybere jen sudá
        Dim pole_sudych() As Integer = {}
        Dim pocitadlo As Integer = 0
        Dim i As Integer, j As Integer

        For i = 0 To UBound(pole)
            If pole(i) Mod 2 = 0 Then
                ReDim Preserve pole_sudych(pocitadlo)
                pole_sudych(pocitadlo) = pole(i)
                pocitadlo += 1
            End If
        Next
        Return pole_sudych
    End Function
End Module
