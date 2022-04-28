Option Explicit
Sub APIKEYS()
    Dim i As Long, j As Long, m As Long, s As String, pool As String, number As Integer, epochs As Integer
    Dim w As Integer, k As Integer, epoch As Integer, index As Integer, capacity As Integer, current As Integer
    w = 0
    k = 1
    epoch = 81
    capacity = InputBox("Ile wygenerować kluczy w jednym epochu?")
    number = InputBox("Dla ilu epochów wygenerować te " & capacity & " kluczy?")
    current = InputBox("Który epoch jest teraz?")
    For epochs = 1 To number
        For index = 1 To capacity
        pool = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ"
        m = Len(pool)
        For i = 1 To 16
            j = 1 + Int(m * Rnd())
            s = s & Mid(pool, j, 1)
        Next i
        Range("A" & k) = s & "," & current
        If k = 1 Then
            Range("E" & k) = "[" & Chr(34) & s & Chr(34) & ","
        Else
            Range("E" & k) = Chr(34) & s & Chr(34) & "," '
            End If
        If k = capacity * number Then
            Range("E" & k) = Chr(34) & s & Chr(34) & "]"
        End If
        k = k + 1
        s = ""
        Next index
        current = current + 1
    Next epochs
End Sub

Function Combine(WorkRng As Range, Optional Sign As String = " AND ") As String
Dim Rng As Range
Dim OutStr As String

For Each Rng In WorkRng
    If Rng.Text <> "," Then
        OutStr = OutStr & Rng.Text & Sign
    End If
Next

Combine = Left(OutStr, Len(OutStr) - 5)

End Function
