Module B12_pojd_si_zaplavat
    Sub Mainx()
        '12) Program se zeptá, zda je den a zda je horko (odpověď a/n). Pouze v případě obou odpovědí kladných vám navrhne jít si zaplavat.
        'Použijte jen jeden příkaz IF (ale se složenou podmínkou). Na rozdíl od obdobného příkladu v BasicCv.doc uvažujte tentokrát i případ,
        'že by uživatel omylem zadal odpověď na první či druhý dotaz velkým písmenem.
        'Toto není moc elegantní, ani nepokrývá všechny varianty
        'If den = "a" And horko = "a" Then
        'MsgBox("jít si zaplavat")
        'ElseIf den = "A" And horko = "A" Then
        'MsgBox("jít si zaplavat")
        'Zkuste použít složenou podmínku obsahující And i Or. Tedy zkuste výše uvedené dvě podmínky spojit do jedné.

        Dim dotaz_horko As String, dotaz_den As String
        Dim horko As Boolean, den As Boolean
        MsgBox("Jsem B12")

        dotaz_den = InputBox("Je den? a/n")
        dotaz_horko = InputBox("Je horko? a/n")

        If dotaz_den = "a" Or dotaz_den = "A" Then
            den = True
        End If

        If dotaz_horko = "a" Or dotaz_horko = "A" Then
            horko = True
        End If

        If den And horko Then
            MsgBox("Jdi si zaplavat.")
        ElseIf den And Not horko Then
            MsgBox("Na plavání bude zima.")
        ElseIf Not den And horko Then
            MsgBox("V noci ja plavání nabezpečné...")
        Else
            MsgBox("Zůstaň doma.")
        End If
    End Sub

End Module
