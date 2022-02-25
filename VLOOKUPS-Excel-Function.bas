Function VLOOKUPS( _
    DataRange As Range, _
    LookupValue1 As String, _
    LookupValue1_ColIndex As Integer, _
    LookupValue2 As String, _
    LookupValue2_ColIndex As Integer, _
    FindValue_ColIndex As Integer)
    
    Dim CurrentRow, MatchingRow As Integer
    
    VLOOKUPS = CVErr(xlErrNA)
    MatchingRow = 0
    CurrentRow = 1
       
    If FindValue_ColIndex > DataRange.Columns.Count Then
        VLOOKUPS = CVErr(xlErrRef)
    Else
        Do
            If ((DataRange.Cells(CurrentRow, LookupValue1_ColIndex).Value = LookupValue1) And _
                (DataRange.Cells(CurrentRow, LookupValue2_ColIndex).Value = LookupValue2)) Then
                MatchingRow = CurrentRow
            End If
            CurrentRow = CurrentRow + 1
        Loop Until ((CurrentRow > DataRange.Rows.Count) Or (MatchingRow <> 0))
    
        If MatchingRow <> 0 Then VLOOKUPS = DataRange.Cells(MatchingRow, FindValue_ColIndex)
    
    End If
    
End Function