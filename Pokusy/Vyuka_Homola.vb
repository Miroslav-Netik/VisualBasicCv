Module Vyuka_Homola
    Sub Mainx()
        MsgBox("Jsem pokusný z výuky docenta Homoly.")
        Dim A() As Long
        Dim B(,,) As Long   'Musím mu říct hned, kolika rozměrné to pole bude - trojrozměrné
        Dim n As Long

        n = 10
        ReDim A(0 To n)
        A(5) = 9

        n = 20

        'Dim unused As Long()
        ReDim Preserve A(0 To n)
        A(20) = 10

        ReDim B(0 To 3, 0 To 4, 0 To 5)

    End Sub
    Sub Soucet_cisel_pole()
        Dim A() As Long
        Dim soucet As Long
        Dim n As Long, i As Long, cislo As Long

        n = 0
        soucet = 0

        Do
            cislo = Val(InputBox("Další číslo: "))
            If cislo <> 0 Then
                n += 1
                ReDim Preserve A(0 To n)
                A(n) = cislo
            End If
        Loop While cislo <> 0

        For i = 1 To n
            soucet += A(i)
        Next

        MsgBox(soucet)
        MsgBox(Trojnasobek(soucet))
        MsgBox(Pojistne_hrubeho(10))
    End Sub
    Sub TestPojistneho()
        Dim P() As Double, h As Double

        h = InputBox("Hrubého? :")
        P = Pojistne_hrubeho(h)
        MsgBox(Str(P(0)) + "....." + Str(P(1)))
    End Sub

    Function Trojnasobek(cislo As Long) As Long
        'Vrátí trojnásobek zadaného čísla
        Return cislo * 3
    End Function

    Function Pojistne_hrubeho(hrubeho As Double) As Double()
        Dim V(0 To 1) As Double
        'Dim V(1 To 2, 1 To 1) As Double
        V(0) = hrubeho / 100 * 4.5
        V(1) = hrubeho / 100 * 6.5
        Return V
    End Function
End Module
