Imports System.Runtime.CompilerServices

Module E13_Posloupnost_celych_opakovani_ktera_kolikrat
    '13) Je dána posloupnost celých čísel (ne lineární, tedy čísla budou nějak na přeskáčku).
    ''Zjistěte, zda se v dané posloupnosti nějaké číslo opakuje. 13 b) až to vyřešíte, vytvořte další program, kde zjistíte, která čísla se opakují.
    '13c) přidejte ještě informaci, kolikrát se opakují.
    'Pozn. 13b a 13c jsou těžké, je možno přeskočit
    Sub Mainx()
        MsgBox("Jsem E13")
        Dim pole_cisel() As Long, pole_opakovanych() As Long
        Dim i As Byte, j As Byte, m As Byte, velikost As Byte
        Dim nahodne As Long, rozsah As Long, k As Long
        Dim pocitadlo As Long, max As Long, min As Long
        Dim txt_generovana_cisla As String, txt_opakovana_cisla As String, txt_min_a_max As String

        txt_generovana_cisla = ""
        txt_opakovana_cisla = ""
        txt_min_a_max = ""
        velikost = 30
        rozsah = 10
        pocitadlo = 0
        max = 0
        min = 0

        ReDim pole_cisel(velikost)
        Randomize()

        'Naplní pole náhodnými celými čísĺy a nastaví min a max
        For i = 0 To velikost
            nahodne = Int(Rnd() * (rozsah - (-10) + 1) + (-10))
            pole_cisel(i) = nahodne
            If nahodne < min Then
                min = nahodne
            End If
            If nahodne > max Then
                max = nahodne
            End If
            txt_generovana_cisla += Str(nahodne) + " "
            txt_min_a_max = "min = " + Str(min) + Chr(10) + "max= " + Str(max)
        Next

        'Pokud jsou dve stejna, vse je OK, kdyz jich je víc, tak je napocíté znova, protoze jeste neopustil smycku for k - min atd...
        For k = min To max
            pocitadlo = 0
            For m = 0 To velikost
                If k = pole_cisel(m) Then
                    pocitadlo += 1
                End If
                If pocitadlo > 1 Then
                    txt_opakovana_cisla += Str(k) + ","
                    pocitadlo = 0
                End If
            Next
        Next

        MsgBox(txt_generovana_cisla + Chr(10) + txt_min_a_max + Chr(10) + txt_opakovana_cisla)
    End Sub

End Module
