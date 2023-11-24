Module B15_SC_jak_je_venku
    Sub Mainx()
        '15) Program se zeptá, jak je venku, nabídne vám: úmorné vedro, zima, déšť, mlha, tma, chumelenice, atd.
        'Podle toho vám navrhne, co máte dělat, např. dobře se obléci, jít na plovárnu atd. Nepoužívejte příkaz IF.
        MsgBox("Jsem B15.")

        Dim volba As Byte

        volba = InputBox("Jak je venku? " + Chr(10) _
                         + " 1 - krásně" + Chr(10) _
                         + " 2 - prší" + Chr(10) _
                         + " 3 - sněží " + Chr(10) _
                         + " Tvoje volba: ")

        Select Case volba
            Case 1
                MsgBox("Je krásně. Vem si plavky.")
            Case 2
                MsgBox("Prší. Vem si deštník.")
            Case 3
                MsgBox("Humelí. Vem si kožich.")
            Case Else
                MsgBox("Tak to nevím, vem si co chceš...")
        End Select
    End Sub
End Module
