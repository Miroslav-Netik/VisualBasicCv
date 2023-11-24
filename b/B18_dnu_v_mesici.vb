Module B18_dnu_v_mesici
    '18) Zadejte číslo měsíce (1 až 12). Program vypíše, kolik má měsíc dní (únor má zjednodušeně 28 dní). Proveďte pomocí if i case.
    '01-31/ 02-28/ 03-31/ 04-30/ 05-31/ 06-30/ 07-31/ 08-31/ 09-30/ 10-31/ 11-30/ 12-31
    Sub Mainx()
        MsgBox("Jsem B18")
        Dim volba As Byte
        Dim txt_mesic As String

        volba = InputBox("Zadej číslo měsíce: ")

        Select Case volba
            Case 1
                txt_mesic = "Leden má 31 dnů."
            Case 2
                txt_mesic = "Únor má 28 dnů."
            Case 3
                txt_mesic = "Březen má 31 dnů."
            Case 4
                txt_mesic = "Duben má 30 dnů."
            Case 5
                txt_mesic = "Květen má 31 dnů."
            Case 6
                txt_mesic = "Červen má 30 dnů."
            Case 7
                txt_mesic = "Červenec má 31 dnů."
            Case 8
                txt_mesic = "Srpen má 31 dnů."
            Case 9
                txt_mesic = "Září má 30 dnů."
            Case 10
                txt_mesic = "Říjen má 31 dnů."
            Case 11
                txt_mesic = "Listopad má 31 dnů."
            Case 12
                txt_mesic = "Prosinec má 30 dnů."
            Case Else
                txt_mesic = "Fakt nevím jaký je to měsíc..."
        End Select

        MsgBox(txt_mesic)
    End Sub
End Module
