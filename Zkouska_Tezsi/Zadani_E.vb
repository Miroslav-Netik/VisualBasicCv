Module Zadani_E
    'e) Zadejte z klávesnice řadu čísel ukončenou nulou (tu již do pole neukládejte). Poté program zobrazí průměr čísel na lichých POZICÍCH.
    '(Řešte ve dvou cyklech, jeden bude sloužit pro naplnění pole a druhý pro zpracovávání pole.) 
    Sub Mainx()
        'MsgBox("Jsem zkouška - zadání E") 'Jen pro ověření při spuštění

        Dim cisla_p(100) As Integer
        Dim i As Integer = 0
        Dim vst_cislo As Integer
        Dim soucet As Integer = 0
        Dim pocitadlo As Integer = 0
        Dim prumer As Double
        Dim txt_zadani As String = ""

        ' Naplnění pole
        Do
            vst_cislo = InputBox("Zadejte číslo (nebo 0 pro ukončení):", "Vstup")
            If vst_cislo = 0 Then
                Exit Do
            End If
            cisla_p(i) = vst_cislo
            txt_zadani += Str(cisla_p(i)) + ","
            i += 1
        Loop

        ' Zpracování pole
        For j As Integer = 0 To i - 1 Step 2
            soucet += cisla_p(j)
            pocitadlo += 1
        Next

        ' Spočte průměr a pak ho vypíše
        prumer = soucet / pocitadlo
        MsgBox(txt_zadani + Chr(10) + "Průměr čísel na lichých pozicích je: " + Str(prumer))

    End Sub
End Module
