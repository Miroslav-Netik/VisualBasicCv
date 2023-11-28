Module E14_bublinkove_trideni
    '(14) Vygenerujte deset celých čísel a setřiďte je.  Použijte např. algoritmus Bublinkového třídění (Bubble sort)
    Sub Mainx()
        MsgBox("Jsem E14")
        Dim pole_cisel() As Long
        Dim nahodna As Long, porovnavaci_i As Long
        Dim i As Integer, i_2 As Integer, i_3 As Integer, velikost As Integer
        Dim txt_nahodna As String, txt_serazena As String

        txt_nahodna = ""
        txt_serazena = ""
        velikost = 50
        ReDim pole_cisel(velikost)

        Randomize()

        For i = 0 To velikost
            nahodna = Int(Rnd() * (10 - (-10) + 1) + (-10))
            pole_cisel(i) = nahodna
            txt_nahodna += Str(nahodna) + ", "
        Next

        For i = velikost To 0 Step -1
            For i_2 = 0 To i - 1
                If pole_cisel(i_2) > pole_cisel(i_2 + 1) Then   'Pokud "index 0" je větší než "index 1"
                    porovnavaci_i = pole_cisel(i_2 + 1)         'Tak vlož "index 1" do porovnávací proměnné 
                    pole_cisel(i_2 + 1) = pole_cisel(i_2)       'Pak vlož na hodnotu "indexu 1" hodnotu "indexu 0"
                    pole_cisel(i_2) = porovnavaci_i             'Do hodnoty "indexu 0" vlož hodnotu "indexu 1"
                End If
            Next
        Next

        For i_3 = 0 To velikost
            txt_serazena += Str(pole_cisel(i_3)) + ", "
        Next
        MsgBox(txt_nahodna + Chr(10) + txt_serazena)
    End Sub

    Sub Mainx_2()
        MsgBox("Jsem E 14 - verze ze skript")

        Dim A(100) As Long, n As Long, cislo As Long
        Dim i As Long, j As Long, p As Long
        Dim s As String

        n = 0
        i = 0
        Do
            cislo = InputBox("Zadej číslo (0 končí): ")
            If cislo <> 0 Then
                n = n + 1
                A(n) = cislo
            End If
        Loop Until cislo = 0

        For j = 0 To n - 1
            For i = 0 To n - 1
                If A(i) > A(i + 1) Then
                    p = A(i)
                    A(i) = A(i + 1)
                    A(i + 1) = p
                End If
            Next
        Next

        For i = 0 To n
            s += Chr(10) + Str(A(i))
        Next

        MsgBox(s)
    End Sub
End Module
