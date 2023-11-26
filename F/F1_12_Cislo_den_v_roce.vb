Imports System.Security.Cryptography

Module F1_12_Cislo_den_v_roce
    '12) Sestavte funkci CisloDneVRoce, která má tři parametry: prvním je číslo dne, druhým číslo měsíce, třetím číslo roku včetně století.
    'Výsledkem volání funkce je pořadové číslo zadaného dne v daném roce (nebo -1, jestliže takové datum neexistuje).
    'Uvažujte i přestupné roky. Funkci ověřte na zadání: sestavte program, který postupně načte den, měsíc a rok, a poté zobrazí,
    'kolikátý den v roce to je. Vyzkoušejte i pro den = 32 nebo pro měsíc = 13. Dále vyzkoušejte pro [31/12/1900] a pro [31/12/2000]. 
    'V programu využijete datový typ Date a funkce Weekday, WeekdayName a CDate
    'Pozor, Excel považuje chybně rok 1900 za přestupný. 
    'Tento program je dost náročný, můžete si jej nechat na konec
    Sub Mainx()
        MsgBox("Jsem F12")
        Dim cislo_dne As Byte, cislo_mesice As Byte
        Dim cislo_roku As Integer, pocet_dnu As String

        cislo_dne = InputBox("Zadej číslo dne: ")
        cislo_mesice = InputBox("Zadej číslo měsíce: ")
        cislo_roku = InputBox("Zadej číslo roku: ")

        pocet_dnu = CisloDneVRoce(cislo_dne, cislo_mesice, cislo_roku)
        MsgBox(Str(pocet_dnu))
    End Sub

    Function CisloDneVRoce(den As Byte, mesic As Byte, rok As Integer) As Integer
        Dim soucet_dnu As Integer
        Dim i As Byte
        soucet_dnu = 0
        'prestupny_rok = F_prestupny_rok(rok)
        If (den > 0 And den <= PocetDniMesice_2(mesic, rok)) _
            And (mesic > 0 And mesic <= 12) _
            And (rok > 999) Then

            For i = 1 To mesic  'Sečte dny v měsících
                soucet_dnu += Str(PocetDniMesice_2(i, rok))
            Next

            soucet_dnu = soucet_dnu - PocetDniMesice_2(mesic, rok) + den   'Přepočítá poslední měsíc z plného počtu na aktuální
            Return soucet_dnu
        Else
            Return -1
        End If
    End Function

    Function F_prestupny_rok_2(rok As Integer) As Boolean
        'Zjistí, jestli je rok přestupný
        Dim prestupny_rok As Boolean
        If (rok Mod 4 = 0 And rok Mod 100 <> 0) Or rok Mod 400 = 0 Then
            prestupny_rok = True
        Else
            prestupny_rok = False
        End If
        Return prestupny_rok
    End Function

    Function PocetDniMesice_2(cislo As Byte, rok As Integer) As Integer
        'Vrátí počet dní v daném měsíci
        'Volá funkci F_prestupny_rok_2, kvůli sélce února v přestupném roce.
        Dim prestupny_rok As Boolean
        prestupny_rok = F_prestupny_rok_2(rok)
        Select Case cislo
            Case 1
                Return 31
            Case 2
                If prestupny_rok Then
                    Return 29
                Else
                    Return 28
                End If
            Case 3
                Return 31
            Case 4
                Return 30
            Case 5
                Return 31
            Case 6
                Return 30
            Case 7
                Return 31
            Case 8
                Return 31
            Case 9
                Return 30
            Case 10
                Return 31
            Case 11
                Return 30
            Case 12
                Return 31
            Case Else
                Return -1
        End Select
    End Function
End Module
