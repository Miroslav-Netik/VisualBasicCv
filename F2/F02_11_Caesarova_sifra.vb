Module F02_11_Caesarova_sifra
    '11) Obdoba f1-3 ale pro pole znaků. Tedy funkce má na vstupu pole znaků, které vrátí zašifrované
    Sub Mainx()
        Dim pole_znaku_nezasifrovane() As String = {}
        Dim pole_znaku_sifrovanych() As String = {}
        Dim sifrovaci_krok As String
        Dim i As Integer, j As Integer, k As Byte
        Dim vstup_text As String
        Dim txt_zadane_znaky As String = ""
        Dim txt_vystup As String = ""

        vstup_text = InputBox("Zadej zprávu, bez háčků a čárek, kterou chceš zašifrovat:")
        sifrovaci_krok = InputBox("Zadej krok posunu šifry (1 - 25): ")
        j = 1

        ReDim pole_znaku_nezasifrovane(Len(vstup_text) - 1)
        For i = 0 To UBound(pole_znaku_nezasifrovane)
            pole_znaku_nezasifrovane(i) = Mid(vstup_text, j, 1)
            j += 1
            txt_zadane_znaky += pole_znaku_nezasifrovane(i)
        Next
        'MsgBox(txt_zadane_znaky) 'Kontrolní výpis

        pole_znaku_sifrovanych = Sifruj(pole_znaku_nezasifrovane, sifrovaci_krok)
        For k = 0 To UBound(pole_znaku_sifrovanych)
            txt_vystup += pole_znaku_sifrovanych(k)
        Next

        MsgBox(txt_zadane_znaky + Chr(10) + txt_vystup)

    End Sub

    Function Sifruj(pole_vstupni() As String, posun As Byte) As String()
        'Caesarova sifra
        'Posune zadana velka nebo mala pismena o zadany kladny krok
        'Sifruje jen písmena, ostatní znaky nechá beze změn
        'Maxinalni krok 25
        'V ASCII - velka pismena 65 - 90, mala 97 - 122
        Dim pole_sifrovane() As String = {}
        Dim fi As Integer
        Dim kod_pismena As Byte, korekce As Byte, posunuti As Byte

        ReDim pole_sifrovane(UBound(pole_vstupni))



        For fi = 0 To UBound(pole_vstupni)
            kod_pismena = Asc(pole_vstupni(fi))
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
            Else posunuti = kod_pismena
            End If
            pole_sifrovane(fi) = Chr(posunuti)
        Next

        'Sifruj = Chr(posunuti)
        Return pole_sifrovane
    End Function
End Module
