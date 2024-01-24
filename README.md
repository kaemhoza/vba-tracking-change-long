# vba-tracking-change-longDim OldValue As Variant

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    ' Śledzenie wybranej wartości w kolumnach C, D i F
    If Not Application.Intersect(Target, Me.Range("B:B,E:E,I:I,J:J,K:K,L:L,P:P,U:U")) Is Nothing Then
        OldValue = Target.Value
    End If
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    Dim KeyCells As Range
    ' Ustawienie zakresu KeyCells na kolumny B, E i I,J,K,P,U
    Set KeyCells = Me.Range("B:B,E:E,I:I,J:J,K:K,L:L,P:P,U:U")

    If Not Application.Intersect(KeyCells, Target) Is Nothing Then
        Dim Answer As VbMsgBoxResult
        Answer = MsgBox("Czy chcesz zapisać zmianę w komórce " & Target.Address & "?", vbYesNo + vbQuestion, "Potwierdzenie zmiany")

        If Answer = vbYes Then
            Dim ArchSheet As Worksheet
            Set ArchSheet = ThisWorkbook.Sheets("A2")
            Dim NextRow As Long
            NextRow = ArchSheet.Cells(ArchSheet.Rows.Count, "A").End(xlUp).Row + 1

            Dim SourceSheet As Worksheet
            Set SourceSheet = ThisWorkbook.Sheets("A1")
            Dim SourceValue As Variant

            ' Przypisanie wartości z kolumny 5 arkusza A1 bez względu na to, w której kolumnie (B, E i I,J,K,P,U) nastąpiła zmiana
            SourceValue = SourceSheet.Cells(Target.Row, 5).Value

            ArchSheet.Cells(NextRow, 1).Value = NextRow - 1 ' Numer porządkowy
            ArchSheet.Cells(NextRow, 2).Value = Now ' Data zmiany
            ArchSheet.Cells(NextRow, 3).Value = SourceValue ' Wartość z kolumny 2 arkusza A1
            ArchSheet.Cells(NextRow, 4).Value = OldValue ' Poprzednia wartość komórki, w której nastąpiła zmiana
            
            MsgBox "Zmiana zapisana.", vbInformation, "Informacja"
        End If
    End If
End Sub

