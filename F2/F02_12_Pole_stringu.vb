Module F02_12_Pole_stringu
    '12) ještě zbývá procvičit funkci, která má na vstupu pole stringů. Takže naprogramujte variantu na f2-11,
    'ale funkci se bude předávat několik stringů (tedy několik slov). Funkce slova vrátí zašifrovaná.
    Sub Mainx()
        Dim pole_slov_nezasifrovane() As String = {}
        Dim pole_slov_sifrovanych() As String = {}
        Dim sifrovaci_krok As String
        Dim i As Integer = 0
        Dim j As Integer, k As Byte
        Dim vstup_text As String
        Dim txt_zadana_slova As String = ""
        Dim txt_vystup As String = ""

        sifrovaci_krok = InputBox("Zadej krok posunu šifry (1 - 25): ")
        vstup_text = InputBox("Zadej slova, bez háčků a čárek, která chceš zašifrovat (0 ukončuje zadávání):")

        'Naplnění pole + textova promenná pro ověření
        Do
            ReDim Preserve pole_slov_nezasifrovane(i)
            pole_slov_nezasifrovane(i) = vstup_text
            txt_zadana_slova += pole_slov_nezasifrovane(i) + Chr(10)
            i += 1
            vstup_text = InputBox("Zadej slova, bez háčků a čárek, která chceš zašifrovat:")
        Loop While vstup_text <> "0"

        'Zašifrování slov pole funkcí, naplnění textové proměnné pr zobrazení
        pole_slov_sifrovanych = Sifruj(pole_slov_nezasifrovane, sifrovaci_krok)
        For k = 0 To UBound(pole_slov_sifrovanych)
            txt_vystup += pole_slov_sifrovanych(k) + Chr(10)
        Next

        MsgBox(txt_zadana_slova + Chr(10) + txt_vystup)
    End Sub

    Function Sifruj(pole_vstupni() As String, posun As Byte) As String()
        'Caesarova sifra
        'Posune zadana slova o zadany kladny krok
        'Sifruje jen písmena, ostatní znaky nechá beze změn
        'Maxinalni krok 25
        'V ASCII - velka pismena 65 - 90, mala 97 - 122
        Dim pole_sifrovane() As String = {}
        Dim fi As Integer
        Dim fj As Integer
        Dim kod_pismena As Byte, korekce As Byte, posunuti As Byte
        Dim odkladaci_vstupni As String
        Dim odkladaci_vystupni As String = ""

        ReDim pole_sifrovane(UBound(pole_vstupni))


        For fj = 0 To UBound(pole_vstupni)
            odkladaci_vstupni = pole_vstupni(fj)

            For fi = 0 To (Len(odkladaci_vstupni) - 1)
                kod_pismena = Asc(Mid(odkladaci_vstupni, fi + 1, 1))
                posunuti = kod_pismena + posun

                'pro velka pismena
                If kod_pismena >= 65 And kod_pismena <= 90 Then
                    '   posunuti = kod_pismena + posun
                    If posunuti > 90 Then
                        korekce = posunuti - 90
                        posunuti = 64 + korekce
                    End If

                    'pro mala pismena
                ElseIf kod_pismena >= 97 And kod_pismena <= 122 Then
                    '   posunuti = kod_pismena + posun
                    If posunuti > 122 Then
                        korekce = posunuti - 122
                        posunuti = 96 + korekce
                    End If
                Else
                    posunuti = kod_pismena
                End If
                odkladaci_vystupni += Chr(posunuti)
            Next
            pole_sifrovane(fj) = odkladaci_vystupni
            odkladaci_vystupni = ""

        Next

        'Sifruj = Chr(posunuti)
        Return pole_sifrovane
    End Function

End Module
