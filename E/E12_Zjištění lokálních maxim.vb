Module E12_Zjištění_lokálních_maxim
    '12) Program zjistí indexy všech lokálních maxim (tedy prvků pole, které sousedí s nižšími hodnotami,
    'čili které mají vlevo i vpravo nižší hodnotu). Např. mějme pole:
    '1 2 3 1 5 6 7 4 3 2 6 5
    'Lokální maxima jsou 3, 7, 6
    Sub Mainx()
        MsgBox("Jsem E12")

        Dim pole_1() As Long
        Dim nahodne_cislo As Byte, n As Byte, max As Byte
        Dim i As Byte, pocet_opakovani As Byte
        Dim txt_vystup As String, txt_maxima As String

        txt_vystup = ""
        txt_maxima = ""
        n = 0
        pocet_opakovani = 10
        ReDim Preserve pole_1(pocet_opakovani)
        Randomize()

        'Naplní pole
        For i = 0 To pocet_opakovani
            nahodne_cislo = Fix(Rnd() * (10 - 1 + 1) + 1)
            pole_1(i) = nahodne_cislo
            txt_vystup += Str(nahodne_cislo) + "   "
        Next

        'Zjistí lokální maximum a zapíše ho do txt_maxima i s indexem
        Do While n < (pocet_opakovani + 1)
            max = pole_1(n)
            If n > 0 And n < pocet_opakovani Then
                max = pole_1(n)
                If max > pole_1(n - 1) And max > pole_1(n + 1) Then
                    txt_maxima += Str(pole_1(n)) + "- index:" + Str(n) + Chr(10)
                End If
            End If
            n += 1
        Loop
        'Console.Write(txt_vystup + Chr(10) + txt_maxima)
        'Console.ReadLine()
        MsgBox(txt_vystup + Chr(10) + txt_maxima)
    End Sub
End Module
