Module C13_tabulka_male_nasobilky
    '13) Vypište tabulku malé násobilky.Bude 10 sloupců (+záhlaví) a 10 řádků (+záhlaví). Tedy první sloupec a první řádek jsou záhlaví:
    '    1   2   3   4     5    6   7   8   9  10      toto je záhlaví
    '1   1   2   3   4     5    6   7   8   9  10
    '2   2   4   6   8   10 12 14 16 18 20
    '3   3   6   9  12  15 18 21 24 27 30
    '4   atd.
    Sub Mainx()
        MsgBox("Jsem C13")
        Dim i_radek As Byte, i_opakovani_radku As Byte
        Dim radek As String, prvni_radek As String, tabulka As String

        radek = ""
        prvni_radek = ""
        tabulka = ""

        For i_opakovani_radku = 1 To 10
            For i_radek = 1 To 7
                radek = radek + Str(i_radek * i_opakovani_radku) + Chr(9)
            Next
            If i_opakovani_radku = 1 Then
                prvni_radek = Chr(9) + radek
            End If
            radek = Str(i_opakovani_radku) + Chr(9) + radek
            tabulka = tabulka + radek + Chr(10)
            radek = ""

        Next
        MsgBox(prvni_radek + Chr(10) + tabulka)
    End Sub



End Module
