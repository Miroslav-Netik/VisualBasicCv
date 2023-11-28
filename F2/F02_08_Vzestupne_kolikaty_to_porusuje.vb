Module F02_08_Vzestupne_kolikaty_to_porusuje
    '8) Sestavte program, který nejprve čte řadu čísel ukončených nulou, která už do řady nepatří. Pole se předá funkci,
    'která sdělí, kolikátý prvek v řadě porušuje vzestupné řazení čísel. Pokud je pole celé vzestupné, pak vrátí -1
    Sub Mainx()
        Dim pole_cisel() As Single = {}
        Dim vstup_zadani_cisla As Single
        Dim vysledek_funkcce As Single
        Dim i As Byte = 0
        Dim txt_vystup As String = ""
        Dim txt_vystup_plati As String

        vstup_zadani_cisla = InputBox("Zadej číslo (0 pro ukončení):")

        Do
            ReDim Preserve pole_cisel(i)
            pole_cisel(i) = vstup_zadani_cisla
            txt_vystup += Str(pole_cisel(i)) + ","
            i += 1
            vstup_zadani_cisla = InputBox("Zadej číslo (0 pro ukončení):")
        Loop While vstup_zadani_cisla <> 0

        vysledek_funkcce = KolikatyNeniVzestupny(pole_cisel)
        If vysledek_funkcce = -1 Then
            txt_vystup_plati = "Čísla jsou řazena vzestupně."
        Else
            txt_vystup_plati = "Čísla nejsou vzestupně. Porušuje prvek číslo" + Str(vysledek_funkcce)
        End If

        MsgBox(txt_vystup + Chr(10) + txt_vystup_plati)
    End Sub

    Function KolikatyNeniVzestupny(pole_vstup() As Single) As Integer
        'Jakmile je cislo nasledujici coslo v poli menší nez predchozi, prepne vetsi_2 na True
        Dim vetsi As Single
        Dim vetsi_2 As Boolean
        Dim kolikaty_porusuje As Integer
        Dim fi As Byte = 0

        'vetsi = pole_vstup(fi)
        Do
            vetsi = pole_vstup(fi)
            fi += 1
            If pole_vstup(fi) < vetsi Then
                If Not vetsi_2 Then 'Aby to napočítalo jenom první porušující bez toho napočítá všechny a zobrazí poslední - to by šlo na pole
                    vetsi_2 = True
                    kolikaty_porusuje = fi + 1
                End If
            End If
        Loop While fi < UBound(pole_vstup)

        If vetsi_2 Then
            Return kolikaty_porusuje
        Else
            Return -1
        End If
    End Function
End Module
