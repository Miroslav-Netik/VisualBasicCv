Module F02_10_SerazenePole_vzestupne
    '10) Obtížné, třídění: Sestavte funkci SerazenePole, která má jediný parametr: pole celých čísel. Výsledkem funkce
    'je jednorozměrné celočíselné pole, ve kterém jsou stejná čísla jako v poli parametru, ale jsou seřazena vzestupně.
    'Funkci ověřte na zadání: sestavte program, který načte pole celých čísel ukončených nulou, která už k číslům nepatří,
    'a pak zobrazí všechna zadaná čísla ve vzestupném pořadí. (obdoba E14, ale třídění proběhne ve funkci)
    Sub Mainx()
        MsgBox("Jsem F02 - 10 - SerazenePole vzestupne")
        Dim pole_cisel() As Long = {}
        Dim setridene_pole() As Long = {}
        Dim pocet_opakovani As Integer = 10 ' 0 včetně až zadaný počet
        Dim max As Integer = 10
        Dim min As Integer = 1
        Dim i As Integer, j As Integer
        Dim txt_vystup As String = ""
        Dim txt_setridene As String = ""

        For i = 0 To pocet_opakovani
            ReDim Preserve pole_cisel(i)
            pole_cisel(i) = Int(Rnd() * (max - min + 1) + min)
            txt_vystup += Str(pole_cisel(i)) + ","
        Next

        setridene_pole = SerazenePole(pole_cisel)
        For j = 0 To UBound(setridene_pole)
            txt_setridene += Str(setridene_pole(j)) + ","
        Next

        MsgBox(txt_vystup + Chr(10) + txt_setridene)
    End Sub

    Function SerazenePole(vstupni_pole() As Long) As Long()
        Dim setridene_pole(UBound(vstupni_pole)) As Long
        Dim odkladaci As Long
        Dim fi As Integer, fj As Integer

        Array.Copy(vstupni_pole, setridene_pole, vstupni_pole.Length)

        For fi = UBound(setridene_pole) To 1 Step -1
            For fj = 0 To fi - 1
                If setridene_pole(fj) > setridene_pole(fj + 1) Then
                    odkladaci = setridene_pole(fj + 1)
                    setridene_pole(fj + 1) = setridene_pole(fj)
                    setridene_pole(fj) = odkladaci
                End If
            Next
        Next

        Return setridene_pole
    End Function
End Module
