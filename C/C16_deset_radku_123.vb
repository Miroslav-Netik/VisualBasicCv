Module C16_deset_radku_123
    '16) Vypište 10 řádků. Na prvním řádku číslo 1, na druhém řádku 12, na třetím 123 atd. (použijte vnořenou smyčku)
    Sub Mainx()
        MsgBox("Jsem C16")

        Dim i As Byte, j As Byte
        Dim txt_vystup As String

        txt_vystup = ""

        For i = 1 To 10
            For j = 1 To i
                txt_vystup += Str(j)
            Next
            txt_vystup += Chr(10)
        Next
        MsgBox(txt_vystup)
    End Sub

End Module
