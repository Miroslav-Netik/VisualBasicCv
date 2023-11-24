Module B16_mini_kalkulacka
    Sub Mainx()
        '16) Uživatel zadá dvě čísla. Poté se objeví nabídka, zda chce sčítat, odčítat, násobit, dělit nebo končit.
        'Poté se objeví výsledek vybrané operace se zadanými čísly nebo program skončí
        MsgBox("Jsem B16")

        Dim volba As Byte
        Dim operand_1 As Double, operand_2 As Double, vysledek As Double
        Dim znamenko As String

        operand_1 = InputBox("Zadej první operand: ")
        operand_2 = InputBox("Zadej druhý operand: ")
        volba = InputBox("Vyber typ operace:" + Chr(10) _
                         + "Sčítání  - 1" + Chr(10) _
                         + "Odčítání - 2" + Chr(10) _
                         + "Násobení - 3" + Chr(10) _
                         + "Dělení   - 4")
        Select Case volba
            Case 1
                vysledek = operand_1 + operand_2
                znamenko = " +"
            Case 2
                vysledek = operand_1 - operand_2
                znamenko = " -"
            Case 3
                vysledek = operand_1 * operand_2
                znamenko = " x"
            Case 4
                If operand_2 = 0 Then
                    MsgBox("Nemůžu dělit nulou. Zhroutil bych se.")
                End If
                vysledek = operand_1 / operand_2
                znamenko = " :"
            Case Else
                MsgBox("Ukončuji program.")
        End Select

        If volba >= 1 And volba <= 4 And operand_2 <> 0 Then
            MsgBox(Str(operand_1) + znamenko + Str(operand_2) + "=" + Str(vysledek))
        End If
    End Sub

End Module
