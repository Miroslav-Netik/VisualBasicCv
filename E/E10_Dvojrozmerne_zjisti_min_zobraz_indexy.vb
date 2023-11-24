Module E10_Dvojrozmerne_zjisti_min_zobraz_indexy
    '10) Program Doc. Homoly: Dvojrozměrné pole: Naplňte zadáním po řádcích pole 4x4.
    'Potom zjistěte minimum a zobrazte řádkové a sloupcové indexy buněk, na kterých se minimum nachází
    Sub Mainx()
        MsgBox("Jsem E15")

        Dim pole_1(3, 3) As Long
        Dim i As Byte, j As Byte, nahodne As Byte, min As Byte
        Dim txt_vypis As String, txt_souradnice_minima As String

        i = 0
        min = 10
        txt_vypis = ""
        txt_souradnice_minima = ""
        Randomize()

        'Naplní pole náhodnými čísly od 1 do 10
        Do While i < 4
            For j = 0 To 3
                nahodne = Fix(Rnd() * (10 - 1 + 1) + 1)
                pole_1(i, j) = nahodne
                txt_vypis = txt_vypis + Chr(9) + Str(nahodne)
            Next
            txt_vypis += Chr(10)
            i += 1
        Loop

        'Najde v poli minimalni hodnotu
        For i = 0 To 3
            For j = 0 To 3
                If pole_1(i, j) < min Then
                    min = pole_1(i, j)
                End If
            Next
        Next
        txt_vypis += Chr(10) + "Minimum je: " + Str(min)

        'Najde v poli vyskyt minimalni hodnoty a uloží indexy
        For i = 0 To 3
            For j = 0 To 3
                If pole_1(i, j) = min Then
                    txt_souradnice_minima += Str(i) + "," + Str(j) + Chr(10)
                End If
            Next
        Next

        'Výsledný výpis
        MsgBox(txt_vypis + Chr(10) + "Souřadnice minima jsou: " + Chr(10) + txt_souradnice_minima)


    End Sub
End Module
