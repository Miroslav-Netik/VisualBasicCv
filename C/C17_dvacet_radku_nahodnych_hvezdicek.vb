Module C17_dvacet_radku_nahodnych_hvezdicek
    '17) Počítač popíše 20 řádků náhodným počtem hvězdiček.(použijte vnořenou smyčku)
    Sub Mainx()
        MsgBox("Jsem C17")

        Dim i As Byte, j As Byte, nahodny As Byte
        Dim txt_vystup As String

        txt_vystup = ""
        Randomize()

        For i = 1 To 50
            nahodny = Fix((Rnd() * (10 - 1 + 1)) + 1)
            For j = 1 To nahodny
                txt_vystup += "*"
            Next
            txt_vystup += Chr(10)
        Next

        MsgBox(txt_vystup)
    End Sub
End Module
