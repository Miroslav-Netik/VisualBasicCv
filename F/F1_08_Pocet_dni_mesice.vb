Module F1_08_Pocet_dni_mesice
    '8) Sestavte funkci PocetDniMesice, která má jako jediný parametr číslo měsíce. Výsledkem volání funkce
    'je počet dní tohoto měsíce (nebo -1, není-li hodnota měsíce správná).
    'Přestupné roky pro únor neuvažujte. Funkci ověřte na zadání: sestavte program,
    'který načte číslo měsíce a poté zobrazí, kolik dní má tento měsíc.
    'Vyzkoušejte i pro měsíc = 13. Návod: řešte pomocí složené podmínky nebo konstrukcí Select Case.
    '(program Doc. Homoly)
    Sub Mainx()
        MsgBox("Jsem F1-08")

        Dim zadane_cislo_mesice As Byte
        Dim cislo_mesice As Integer

        zadane_cislo_mesice = InputBox("Zadej číslo měsíce (1 až 12)" + Chr(10) _
                                + "a já ti napíšu počet jeho dní: ")

        cislo_mesice = PocetDniMesice(zadane_cislo_mesice)

        MsgBox("Měsíc číslo" + Str(zadane_cislo_mesice) + " má" + Str(cislo_mesice) + " dní.")
    End Sub

    Function PocetDniMesice(cislo As Byte) As Integer
        Select Case cislo
            Case 1
                Return 31
            Case 2
                Return 28
            Case 3
                Return 31
            Case 4
                Return 30
            Case 5
                Return 31
            Case 6
                Return 30
            Case 7
                Return 31
            Case 8
                Return 31
            Case 9
                Return 30
            Case 10
                Return 31
            Case 11
                Return 30
            Case 12
                Return 31
            Case Else
                Return -1
        End Select
        Return 0
    End Function
End Module
