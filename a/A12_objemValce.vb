Module A12_objemValce
    Sub mainx()
        '12) Sestavte program, který se zeptá na průměr vašeho kruhového bazénu a poté
        'na jeho výšku (obojí v metrech). Poté zobrazí,
        'kolik tun vody se do něj vejde (1 m3 vody = 1 tuna). (program Doc. Homoly)
        'V = pi * r^2 * v

        MsgBox("Jsem A12.")

        Dim zadany_prumer As Single, zadana_vyska As Single
        Dim hmotnost As Single, polomer As Single

        zadany_prumer = InputBox("Jaký má tvůj kruhový bazén průměr (v metrech): ")
        zadana_vyska = InputBox("Jakou má tvůj bazén výšku (v metrech): ")

        polomer = zadany_prumer / 2
        hmotnost = Math.PI * (polomer ^ 2) * zadana_vyska
        MsgBox(Str(Math.Round(hmotnost, 2)) + " tun  vody.")
    End Sub
End Module
