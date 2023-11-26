Module F1_11_Vyskyty_retezce
    '11) Sestavte funkci Vyskyty se dvěma parametry - textovými řetězci. Druhý parametr je subřetězec,
    '(případně jen jednoznakový řetězec. Výsledkem volání funkce je počet opakování druhého parametru v parametru prvním.
    'Funkci ověřte na zadání: sestavte program, který načte textový řetězec a zobrazí například počet písmen "A" v tomto řetězci.
    Sub Mainx()
        MsgBox("Jsem F11")

        Dim txt_vstup As String, txt_kontrolovany_vstup As String, txt_vystupni As String
        Dim pocet_nalezenych As Integer

        txt_vstup = InputBox("Zadej kontrolovaný text: ")
        txt_kontrolovany_vstup = InputBox("Zadej jeden znak, jehož počet chceš v textu najít:")

        pocet_nalezenych = Vyskyty(txt_vstup, txt_kontrolovany_vstup)

        MsgBox("Počet znaků " + txt_kontrolovany_vstup + " je" + Str(pocet_nalezenych))
    End Sub

    Function Vyskyty(prvni As String, druhy As String) As String
        'Spočítá počet výskytů zadaného znaku v zadeném řetězci
        Dim i As Integer
        Dim rozdeleny As String
        Dim pocitadlo As Integer

        rozdeleny = ""
        For i = 1 To Len(prvni)
            rozdeleny = Mid(prvni, i, 1)
            If druhy = rozdeleny Then
                pocitadlo += 1
                rozdeleny = ""
            End If
        Next
        Return pocitadlo
    End Function
End Module
