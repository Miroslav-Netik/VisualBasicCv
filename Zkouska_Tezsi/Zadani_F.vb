Module Zadani_F
    'f) Napište funkci, která z pole deseti čísel vrací nejmenší. Přitom čísla budou v hlavním programu generována například v rozsahu -10,10
    '(ale funkce musí být univerzální, tedy musí si poradit i s čísly v jiném rozsahu). Potom vytvořte subrutinu se stejnou funkčností.
    'Z hlavního programu zavolejte funkci i subrutinu.
    'Pozn.: V hlavním programu musí být použity jiné názvy proměnných než v subrutině či funkci. Vstupy a výstupy budou v hlavním programu
    '(tedy ve funkci a subrutině by neměl být příkaz MsgBox). Funkce musí fungovat dobře beze změny i v případě, že se v hlavním programu změní rozměr pole. 
    Sub Mainx()
        'MsgBox("Jsem zkouška - zadání F!")    'Jen pro ověření při spuštění

        Dim pole_cisel() As Integer = {}
        Dim min As Integer
        Dim nejmensi As Integer
        Dim rozsah_min As Integer = -10
        Dim rozsah_max As Integer = 10
        Dim pocet_opakovani As Integer = 10
        Dim txt_generovana As String = ""

        ' Generování čísel v rozsahu -10 až 10
        Randomize()
        For i = 0 To pocet_opakovani - 1    'Počítá od 0 proto odečtu 1
            ReDim Preserve pole_cisel(i)
            pole_cisel(i) = Int(Rnd() * (rozsah_max - (rozsah_min) + 1) + (rozsah_min))
            txt_generovana += Str(pole_cisel(i)) + ","
        Next

        ' Volání funkce
        min = NajdiMin(pole_cisel)
        MsgBox(txt_generovana + Chr(10) + "Nejmenší číslo - funkce: " + Str(min))

        ' Volání subrutiny
        NajdiMinS(pole_cisel, nejmensi)
        MsgBox(txt_generovana + Chr(10) + "Nejmenší číslo - subrutina: " + Str(nejmensi))
    End Sub
    Function NajdiMin(cisla() As Integer) As Integer
        'Najde nejmenší v zadaném poli
        Dim min As Integer = cisla(0)
        For i As Integer = 0 To cisla.Length - 1
            If cisla(i) < min Then
                min = cisla(i)
            End If
        Next
        Return min
    End Function
    Sub NajdiMinS(cisla() As Integer, ByRef minimum As Integer)
        'Najde nejmenší v zadaném poli
        minimum = cisla(0)
        For i As Integer = 0 To UBound(cisla)
            If cisla(i) < minimum Then
                minimum = cisla(i)
            End If
        Next
    End Sub


End Module
