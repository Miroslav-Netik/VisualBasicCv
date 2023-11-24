Module C14_hvezdicky
    Sub Mainx()
        '14) Vytisknete na prvním řádku jednu hvězdu, na druhém dvě až do desátého.
        'Potom se začnou hvězdy ubírat.
        'Varianta: Totéž, ale na prvním řádku se vytisknou dvě hvězdy, na druhém
        'čtyři atd.
        'Varianta 2: Totéž, ale na prvnim i druhem řádku se vytiskne hvězda, na
        'třetím a čtvrtém dvě hvězdy atd.
        MsgBox("Jsem C14")

        Dim hvezda As String, hvezdy As String, vystup As String
        Dim i As Integer

        hvezda = "*"
        hvezdy = ""
        vystup = ""

        For i = 1 To 10
            hvezdy = hvezdy + hvezda
            vystup = vystup + hvezdy + Chr(10)
        Next

        For i = 1 To 10

        Next
        MsgBox(vystup)

        '***********************
        'Následující řešení Bing
        '***********************

        'Dim hvezda As String, hvezdy As String
        Dim j As Byte

        hvezda = "*"
        hvezdy = ""

        ' Přidávání hvězd
        For i = 1 To 10
            For j = 1 To i
                hvezdy = hvezdy + hvezda
            Next
            hvezdy = hvezdy + vbCrLf
        Next

        ' Odebírání hvězd
        For i = 9 To 1 Step -1
            For j = 1 To i
                hvezdy = hvezdy + hvezda
            Next
            hvezdy = hvezdy + vbCrLf
        Next

        MsgBox(hvezdy)

        'Varianta: Totéž, ale na prvním řádku se vytisknou dvě hvězdy, na druhém čtyři atd.
        Dim k As Byte

        hvezda = "*"
        hvezdy = ""

        ' Přidávání hvězd
        For i = 1 To 10
            For k = 1 To 2
                For j = 1 To i
                    hvezdy = hvezdy + hvezda
                Next
            Next
            hvezdy = hvezdy + vbCrLf
        Next

        ' Odebírání hvězd
        For i = 9 To 1 Step -1
            For k = 1 To 2
                For j = 1 To i
                    hvezdy = hvezdy + hvezda
                Next
            Next
            hvezdy = hvezdy + vbCrLf
        Next

        MsgBox(hvezdy)

        'Varianta 2: Totéž, ale na prvnim i druhem řádku se vytiskne hvězda, na
        'třetím a čtvrtém dvě hvězdy atd.

        hvezda = "*"
        hvezdy = ""

        ' Přidávání hvězd
        For i = 1 To 10
            For k = 1 To 2
                For j = 1 To i
                    hvezdy = hvezdy + hvezda
                Next
                hvezdy = hvezdy + vbCrLf
            Next
        Next

        ' Odebírání hvězd
        For i = 9 To 1 Step -1
            For k = 1 To 2
                For j = 1 To i
                    hvezdy = hvezdy + hvezda
                Next
                hvezdy = hvezdy + vbCrLf
            Next
        Next

        MsgBox(hvezdy)
    End Sub
End Module
