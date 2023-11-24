Module B17_prevod_men
    '17) Zadejte částku v Kč a pak si z nabídky vyberte měnu, na kterou chcete směnit (např. DEM, USD, FRF atd.). Počítač přepočítá částku na tuto měnu.
    Sub Mainx()
        MsgBox("Jsem B17")

        Dim zadana_castka As Single, vystupni_castka As Single
        Dim volba As Byte
        Dim txt_prevod As String

        txt_prevod = ""

        zadana_castka = InputBox("Zadej částku v korunách: ")
        volba = InputBox("Na jakou měnu chceš převádět:" + Chr(10) _
                         + "1 - EUR" + Chr(10) _
                         + "2 - USD" + Chr(10) _
                         + "3 - FRF")
        Select Case volba
            Case 1
                vystupni_castka = zadana_castka / 24
                txt_prevod = Str(vystupni_castka) + " EUR."
            Case 2
                vystupni_castka = zadana_castka / 26
                txt_prevod = Str(vystupni_castka) + " USD."

            Case 3
                vystupni_castka = zadana_castka / 10
                txt_prevod = Str(vystupni_castka) + " FRF."
        End Select

        MsgBox(Str(zadana_castka) + " Kč je" + txt_prevod)
    End Sub

End Module
