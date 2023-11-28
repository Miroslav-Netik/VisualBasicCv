Module F02_09_Linearni_posloupnost
    '9) Sestavte program, který nejprve čte řadu čísel ukončených nulou, která už do řady nepatří.
    'Pole se předá funkci, která sdělí (boolean), zda čísla tvoří lineární aritmetickou (tj. lineární) posloupnost
    '(rozdíl dvou sousedních je vždy tentýž). 
    Sub Mainx()
        MsgBox("Jsem F02 - 09 - Lineární posloupnost")
        Dim pole_cisel() As Single = {}
        Dim zadane_cislo As Single
        Dim i As Integer = 0
        Dim txt_vystup As String = ""
        Dim txt_Je_neni As String

        'Zadání čísel - ukončeno nulou
        zadane_cislo = InputBox("Zadej číslo posloupnosti (0 pro ukončení: ")
        Do
            ReDim Preserve pole_cisel(i)
            pole_cisel(i) = zadane_cislo
            txt_vystup += Str(pole_cisel(i)) + Chr(10)
            i += 1
            zadane_cislo = InputBox("Zadej číslo posloupnosti (0 pro ukončení: ")
        Loop While zadane_cislo <> 0

        'Naplní textovou proměnnou, podle výsledku funkce
        If ZjistiPosloupnost(pole_cisel) Then
            txt_Je_neni = "Je to posloupnost."
        Else
            txt_Je_neni = "Není to posloupnost."
        End If

        'Výpis hodnot z pole i výsledků
        MsgBox(txt_vystup + Chr(10) + txt_Je_neni)
    End Sub

    Function ZjistiPosloupnost(vstupni_pole() As Single) As Boolean
        'Zjistí, jestli je zadané pole lineární aritmetická posloupnost
        Dim pole_funkce() As Single = {}
        Dim je As Boolean = True
        Dim fi As Integer
        Dim diference As Integer

        If UBound(vstupni_pole) > 1 Then
            diference = vstupni_pole(fi + 1) - vstupni_pole(fi)
            For fi = 1 To UBound(vstupni_pole)
                If ((vstupni_pole(fi) - vstupni_pole(fi - 1)) <> diference) Then
                    je = False
                End If
            Next
        End If

        If je = False Then
            Return False
        Else
            Return True
        End If
    End Function
End Module
