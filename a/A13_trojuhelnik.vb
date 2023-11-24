Module A13_trojuhelnik
    Sub mainx()
        '13) Zadejte délku odvěsen pravoúhlého trojúhelníka. Program pomocí Pythagorovy věty vypočte délku přepony. Pro kontrolu:
        'zadáte-li odvěsny o délce 3 a 4, pak výsledek je 5 (přesněji 5, 0, protože výsledek odmocniny se musí ukládat do
        'desetinného typu )
        MsgBox("Jsem A13.")
        Dim prepona As Single, odvesna_1 As Single, odvesna_2 As Single
        Dim txt_vystup As String

        odvesna_1 = InputBox("Zadej délku první odvěsny: ")
        odvesna_2 = InputBox("Zadej délku druhé odvěsny: ")
        prepona = Math.Sqrt(odvesna_1 ^ 2 + odvesna_2 ^ 2)
        txt_vystup = String.Format("Přepona trojúhelníku je {0} dlouhá.", prepona)
        MsgBox(txt_vystup)

    End Sub
End Module
