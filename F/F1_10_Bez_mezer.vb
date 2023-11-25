Module F1_10_Bez_mezer
    '10) Sestavte funkci BezMezer s jedním parametrem - textovým řetězcem. Výsledkem volání funkce
    'je tentýž textový řetězec, ve kterém však jsou vypuštěny všechny mezery. Funkci ověřte
    'na zadání: sestavte program, který načte textový řetězec a zobrazí ho tak,
    'že v něm jsou vypuštěny všechny mezery.
    Sub Mainx()
        MsgBox("Jsem F10")

        Dim zadany_text As String, vystupni_text As String

        zadany_text = InputBox("Zadej text a já tě zbavím mezer: ")
        vystupni_text = BezMezer(zadany_text)
        MsgBox(vystupni_text)
    End Sub

    Function BezMezer(text As String) As String
        Dim i As Long
        Dim upraveny As String, vystup As String
        vystup = ""
        For i = 1 To Len(text)
            upraveny = Mid(text, i, 1)
            If upraveny = " " Then
                upraveny = ""
                vystup += upraveny
            Else
                vystup += upraveny
            End If
        Next
        Return vystup
    End Function
End Module
