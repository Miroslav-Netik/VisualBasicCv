Module F02_04_PocetVyskytu
    '4) Sestavte funkci PocetVyskytu, která má dva parametry: Prvním je jednorozměrné pole čísel A, druhým
    'je číslo C.
    'Výsledkem funkce je celé číslo udávající, kolikrát se číslo C vyskytuje v poli A. Funkci ověřte
    'na zadání:
    'sestavte program, který postupně načte pole A a hodnotu C, a poté zobrazí počet výskytů
    'hodnoty C v poli A. 
    Sub Mainx()
        Dim pole_hodnot() As Integer = {}
        Dim hledane_cislo As Integer, pocet_opakovani As Integer
        Dim rozsah_max As Integer, rozsah_min As Integer
        Dim i As Integer
        Dim txt_vystup As String, txt_generovana_cisla As String

        rozsah_max = 100
        rozsah_min = 1
        pocet_opakovani = 1000
        txt_generovana_cisla = ""


        hledane_cislo = InputBox("Zadej hledané celé číslo: ")
        For i = 0 To pocet_opakovani
            ReDim Preserve pole_hodnot(i)
            pole_hodnot(i) = Int(Rnd() * (rozsah_max - rozsah_min + 1) + rozsah_min)
            txt_generovana_cisla += Str(pole_hodnot(i)) + ","
        Next
        txt_vystup = Str(PocetVyskytu(pole_hodnot, hledane_cislo))

        MsgBox(txt_generovana_cisla + Chr(10) + txt_vystup + " krát")
    End Sub

    Function PocetVyskytu(A() As Integer, C As Integer) As Integer
        Dim i As Integer
        Dim pocitadlo As Integer

        For i = 0 To UBound(A)
            If A(i) = C Then
                pocitadlo += 1
            End If
        Next

        Return pocitadlo
    End Function
End Module
