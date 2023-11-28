Module F02_07_Je_serazeno_vzestupne_
    '7) Sestavte program, který nejprve čte řadu čísel ukončených nulou, která už do řady nepatří. Pole se předá funkci,
    'která sdělí (boolean), zda jsou čísla seřazena vzestupně. Tedy každé další číslo je větší než předchozí.
    ''Čili stačí jeden jediný pokles a už to není pravda.
    Sub Mainx()
        Dim pole_cisel() As Single = {}
        Dim vstup_zadani_cisla As Single
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

        If Vzestupne(pole_cisel) Then
            txt_vystup_plati = "Čísla jsou řazena vzestupně."
        Else
            txt_vystup_plati = "Čísla nejsou vzestupně."
        End If

        MsgBox(txt_vystup + Chr(10) + txt_vystup_plati)
    End Sub

    Function Vzestupne(pole_vstup() As Single) As Boolean
        'Jakmile je cislo nasledujici coslo v poli nez predchozi, prepne vetsi_2 na true
        Dim vetsi As Single
        Dim vetsi_2 As Boolean
        Dim fi As Byte = 0

        'vetsi = pole_vstup(fi)
        Do
            vetsi = pole_vstup(fi)
            fi += 1
            If pole_vstup(fi) < vetsi Then
                vetsi_2 = True
            End If
        Loop While fi < UBound(pole_vstup)

        If vetsi_2 Then
            Return False
        Else
            Return True
        End If
    End Function
End Module
