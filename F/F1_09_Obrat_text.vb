Module F1_09_Obrat_text
    '9) Sestavte funkci ObratText s jedním parametrem - textovým řetězcem. Výsledkem volání funkce
    'je tentýž textový řetězec, ale má znaky v obráceném pořadí.
    'Funkci ověřte na zadání: sestavte program, který načte textový řetězec a zobrazí ho v
    'obráceném pořadí znaků.
    Sub Mainx()
        MsgBox("Jsem F1-09")

        Dim zadany_text As String, obraceny_text As String
        Dim x As String
        zadany_text = InputBox("Zadej text a já ho celý otočím: ")
        obraceny_text = ObratText_F(zadany_text)

        MsgBox("Funkce: " + Chr(10) + zadany_text + Chr(10) + obraceny_text)

        ObratText_S(zadany_text, x)
        MsgBox("Subrutina: " + Chr(10) + zadany_text + Chr(10) + x)
    End Sub

    Function ObratText_F(text As String) As String
        'Obrátí zadyný text
        Dim obraceny As String, vystup As String
        Dim i As Double
        obraceny = ""
        For i = Len(text) To 1 Step -1
            obraceny = Mid(text, i, 1)
            vystup += obraceny
        Next
        Return vystup
    End Function

    Sub ObratText_S(text As String, ByRef vyst As String)
        'Obrátí zadyný text
        Dim obraceny As String
        Dim vystup As String = ""
        Dim i As Double
        obraceny = ""
        For i = Len(text) To 1 Step -1
            obraceny = Mid(text, i, 1)
            vystup += obraceny
        Next
        vyst = vystup
    End Sub
End Module
