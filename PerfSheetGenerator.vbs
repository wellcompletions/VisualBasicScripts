Sub Rebuild_Sheet()
'Turn off screen updating to speed up macro. i & "
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual


If MsgBox("This will rebuild the sheet formulas!! Are you sure ?", vbYesNo) = vbYes Then
    'code to send message
Else
    Exit Sub 'terminate macro
End If
If MsgBox("There are a lot of cells to rebuild!! Are you sure you want to commit changes?", vbYesNo) = vbYes Then
    'code to send message
Else
    Exit Sub 'terminate macro
End If


'start with the bottom cells needed
Dim btmRow As Integer
btmRow = (2 * Range("W5").Value) + 9

Range("A" & btmRow).Formula = "Test Sub"
Range("A" & btmRow + 1).Formula = "Setback (GL) +30"
Range("A" & btmRow + 2).Formula = "# Intervals"
Range("E" & btmRow).Formula = "Toe Setback @"
Range("E" & btmRow + 1).Formula = "Heel Setback @"
Range("J" & btmRow).Formula = "Interval Length"
Range("C" & btmRow).Formula = "='Well Info'!$AF$10"
Range("G" & btmRow).Formula = "='Well Info'!$AF$22"
Range("G" & btmRow + 1).Formula = "='Well Info'!$AF$26"
Range("C" & btmRow + 1).Formula = "=$G$" & (2 * Range("W5").Value + 10) & " +30"
Range("C" & btmRow + 2).Formula = "=W5"
Range("L" & btmRow).Formula = "=(C8-C" & btmRow + 1 & ")/C" & btmRow + 2
Range("C9").Formula = "=IF(C" & btmRow & "<G" & btmRow & ",C" & btmRow & "-20,G" & btmRow & "-20)"
Range("D" & btmRow).Formula = "ft MD (GL)"
Range("D" & btmRow + 1).Formula = "ft MD (GL)"
Range("H" & btmRow).Formula = "ft MD (GL)"
Range("H" & btmRow + 1).Formula = "ft MD (GL)"

'for later to clear cells
'Range("Z1:Z200").ClearContents

 'a different loop - didn't need
' Dim i As Integer
' For i = 9 To ((2*Range("W5").value)) + 7 Step 2
'     Range("z" & i).Formula = i
' Next i

' even rows, populate rows A, B, C, D
Dim j As Integer
j = 1
For i = 8 To (2 * Range("W5").Value) + 6 Step 2
    Range("A" & i).Formula = "=$L$" & btmRow
    Range("B" & i).Formula = j
    If i = 8 Then
        Range("C" & i + 1).Formula = "=IF(C" & btmRow & "<G" & btmRow & ",C" & btmRow & "-20,G" & btmRow & "-20)"
        Range("C" & i).Formula = "=IF(C" & btmRow & "<G" & btmRow & ",C" & btmRow & "-20,G" & btmRow & "-20)"
    Else
        Range("C" & i).Formula = "=IF(COUNTIF(collars,ROUND($C$8-($L$" & btmRow & "*(B" & i & "-1)),0)),$C$8-($L$" & btmRow & "*(B" & i & "-1))+4,IF(COUNTIF(collars,ROUND($C$8-($L$" & btmRow & "*(B" & i & "-1)),0)+1),$C$8-($L$" & btmRow & "*(B" & i & "-1))-3,IF(COUNTIF(collars,ROUND($C$8-($L$" & btmRow & "*(B" & i & "-1)),0)-1),$C$8-($L$" & btmRow & "*(B" & i & "-1))+3,IF(COUNTIF(collars,ROUND($C$8-($L$" & btmRow & "*(B" & i & "-1)),0)+2),$C$8-($L$" & btmRow & "*(B" & i & "-1))-2,IF(COUNTIF(collars,ROUND($C$8-($L$" & btmRow & "*(B" & i & "-1)),0)-2),$C$8-($L$" & btmRow & "*(B" & i & "-1))+2,$C$8-($L$" & btmRow & "*(B" & i & "-1)))))))"
        Range("C" & i + 1).Formula = "=C" & i - 1 & "-$L$" & btmRow

    End If
    ' this line does the fix for circular references with the -17 adjustment
    Range("D" & i).Formula = "=IF(COUNTIF(collars,ROUND(C" & i & "-15,0)),C" & i & "-17,IF(COUNTIF(collars,ROUND(C" & i & "-15,0)+1),C" & i & "-15-1,IF(COUNTIF(collars,ROUND(C" & i & "-15,0)-1),C" & i & "-15+1,C" & i & "-15)))"
    Range("D" & i + 1).Formula = "=C" & i + 1 & "-15"
    Range("E" & i + 1).Formula = "=F" & i + 1 & "+ $S$" & i
    Range("F" & i + 1).Formula = "=G" & i + 1 & "+ $S$" & i
    Range("G" & i + 1).Formula = "=H" & i + 1 & "+ $S$" & i
    Range("H" & i + 1).Formula = "=I" & i + 1 & "+ $S$" & i
    Range("I" & i + 1).Formula = "=J" & i + 1 & "+ $S$" & i
    Range("J" & i + 1).Formula = "=K" & i + 1 & "+ $S$" & i
    Range("K" & i + 1).Formula = "=L" & i + 1 & "+ $S$" & i
    Range("L" & i + 1).Formula = "=M" & i + 1 & "+ $S$" & i
    Range("M" & i + 1).Formula = "=N" & i + 1 & "+ $S$" & i
    Range("N" & i + 1).Formula = "=O" & i + 1 & "+ $S$" & i
    Range("O" & i + 1).Formula = "=P" & i + 1 & "+ $S$" & i
    Range("P" & i + 1).Formula = "=Q" & i + 1 & "+ $S$" & i
        
    ' need to have sepatate values in the bottom right corner to pick up the Heel Setback
    If i = (2 * Range("W5").Value) + 6 Then
        ' all values for the last row
        Range("Q" & i + 1).Formula = "=C" & i + 4
        Range("E" & i).Formula = "=IF(COUNTIF(collars,ROUND($C$8 -($L$" & btmRow & "* $B" & i & ")+$S" & i & "*12,0)), $C$8 -($L$" & btmRow & "*$B" & i & ")+$S" & i & "*12+2, IF(COUNTIF(collars,ROUND($C$8-($L$" & btmRow & " * $B" & i & ")+$S" & i & "*12, 0)+1), $C$8 - ($L$" & btmRow & " * $B" & i & ")+$S" & i & " *12-1, IF(COUNTIF(collars,ROUND($C$8-($L$" & btmRow & " * $B" & i & ")+$S" & i & " *12,0)-1), $C$8-($L$" & btmRow & " * $B" & i & ")+$S" & i & " * 12+1, $C$8-($L$" & btmRow & " * $B" & i & ")+$S" & i & "*12)))"
        Range("F" & i).Formula = "=IF(COUNTIF(collars,ROUND($C$8 -($L$" & btmRow & "* $B" & i & ")+$S" & i & "*11,0)), $C$8 -($L$" & btmRow & "*$B" & i & ")+$S" & i & "*11+2, IF(COUNTIF(collars,ROUND($C$8-($L$" & btmRow & " * $B" & i & ")+$S" & i & "*11, 0)+1), $C$8 - ($L$" & btmRow & " * $B" & i & ")+$S" & i & " *11-1, IF(COUNTIF(collars,ROUND($C$8-($L$" & btmRow & " * $B" & i & ")+$S" & i & " *11,0)-1), $C$8-($L$" & btmRow & " * $B" & i & ")+$S" & i & " * 11+1, $C$8-($L$" & btmRow & " * $B" & i & ")+$S" & i & "*11)))"
        Range("G" & i).Formula = "=IF(COUNTIF(collars,ROUND($C$8 -($L$" & btmRow & "* $B" & i & ")+$S" & i & "*10,0)), $C$8 -($L$" & btmRow & "*$B" & i & ")+$S" & i & "*10+2, IF(COUNTIF(collars,ROUND($C$8-($L$" & btmRow & " * $B" & i & ")+$S" & i & "*10, 0)+1), $C$8 - ($L$" & btmRow & " * $B" & i & ")+$S" & i & " *10-1, IF(COUNTIF(collars,ROUND($C$8-($L$" & btmRow & " * $B" & i & ")+$S" & i & " *10,0)-1), $C$8-($L$" & btmRow & " * $B" & i & ")+$S" & i & " * 10+1, $C$8-($L$" & btmRow & " * $B" & i & ")+$S" & i & "*10)))"
        Range("H" & i).Formula = "=IF(COUNTIF(collars,ROUND($C$8 -($L$" & btmRow & "* $B" & i & ")+$S" & i & "*9,0)), $C$8 -($L$" & btmRow & "*$B" & i & ")+$S" & i & "*9+2, IF(COUNTIF(collars,ROUND($C$8-($L$" & btmRow & " * $B" & i & ")+$S" & i & "*9, 0)+1), $C$8 - ($L$" & btmRow & " * $B" & i & ")+$S" & i & " *9-1, IF(COUNTIF(collars,ROUND($C$8-($L$" & btmRow & " * $B" & i & ")+$S" & i & " *9,0)-1), $C$8-($L$" & btmRow & " * $B" & i & ")+$S" & i & " * 9+1, $C$8-($L$" & btmRow & " * $B" & i & ")+$S" & i & "*9)))"
        Range("I" & i).Formula = "=IF(COUNTIF(collars,ROUND($C$8 -($L$" & btmRow & "* $B" & i & ")+$S" & i & "*8,0)), $C$8 -($L$" & btmRow & "*$B" & i & ")+$S" & i & "*8+2, IF(COUNTIF(collars,ROUND($C$8-($L$" & btmRow & " * $B" & i & ")+$S" & i & "*8, 0)+1), $C$8 - ($L$" & btmRow & " * $B" & i & ")+$S" & i & " *8-1, IF(COUNTIF(collars,ROUND($C$8-($L$" & btmRow & " * $B" & i & ")+$S" & i & " *8,0)-1), $C$8-($L$" & btmRow & " * $B" & i & ")+$S" & i & " * 8+1, $C$8-($L$" & btmRow & " * $B" & i & ")+$S" & i & "*8)))"
        Range("J" & i).Formula = "=IF(COUNTIF(collars,ROUND($C$8 -($L$" & btmRow & "* $B" & i & ")+$S" & i & "*7,0)), $C$8 -($L$" & btmRow & "*$B" & i & ")+$S" & i & "*7+2, IF(COUNTIF(collars,ROUND($C$8-($L$" & btmRow & " * $B" & i & ")+$S" & i & "*7, 0)+1), $C$8 - ($L$" & btmRow & " * $B" & i & ")+$S" & i & " *7-1, IF(COUNTIF(collars,ROUND($C$8-($L$" & btmRow & " * $B" & i & ")+$S" & i & " *7,0)-1), $C$8-($L$" & btmRow & " * $B" & i & ")+$S" & i & " * 7+1, $C$8-($L$" & btmRow & " * $B" & i & ")+$S" & i & "*7)))"
        Range("K" & i).Formula = "=IF(COUNTIF(collars,ROUND($C$8 -($L$" & btmRow & "* $B" & i & ")+$S" & i & "*6,0)), $C$8 -($L$" & btmRow & "*$B" & i & ")+$S" & i & "*6+2, IF(COUNTIF(collars,ROUND($C$8-($L$" & btmRow & " * $B" & i & ")+$S" & i & "*6, 0)+1), $C$8 - ($L$" & btmRow & " * $B" & i & ")+$S" & i & " *6-1, IF(COUNTIF(collars,ROUND($C$8-($L$" & btmRow & " * $B" & i & ")+$S" & i & " *6,0)-1), $C$8-($L$" & btmRow & " * $B" & i & ")+$S" & i & " * 6+1, $C$8-($L$" & btmRow & " * $B" & i & ")+$S" & i & "*6)))"
        Range("L" & i).Formula = "=IF(COUNTIF(collars,ROUND($C$8 -($L$" & btmRow & "* $B" & i & ")+$S" & i & "*5,0)), $C$8 -($L$" & btmRow & "*$B" & i & ")+$S" & i & "*5+2, IF(COUNTIF(collars,ROUND($C$8-($L$" & btmRow & " * $B" & i & ")+$S" & i & "*5, 0)+1), $C$8 - ($L$" & btmRow & " * $B" & i & ")+$S" & i & " *5-1, IF(COUNTIF(collars,ROUND($C$8-($L$" & btmRow & " * $B" & i & ")+$S" & i & " *5,0)-1), $C$8-($L$" & btmRow & " * $B" & i & ")+$S" & i & " * 5+1, $C$8-($L$" & btmRow & " * $B" & i & ")+$S" & i & "*5)))"
        Range("M" & i).Formula = "=IF(COUNTIF(collars,ROUND($C$8 -($L$" & btmRow & "* $B" & i & ")+$S" & i & "*4,0)), $C$8 -($L$" & btmRow & "*$B" & i & ")+$S" & i & "*4+2, IF(COUNTIF(collars,ROUND($C$8-($L$" & btmRow & " * $B" & i & ")+$S" & i & "*4, 0)+1), $C$8 - ($L$" & btmRow & " * $B" & i & ")+$S" & i & " *4-1, IF(COUNTIF(collars,ROUND($C$8-($L$" & btmRow & " * $B" & i & ")+$S" & i & " *4,0)-1), $C$8-($L$" & btmRow & " * $B" & i & ")+$S" & i & " * 4+1, $C$8-($L$" & btmRow & " * $B" & i & ")+$S" & i & "*4)))"
        Range("N" & i).Formula = "=IF(COUNTIF(collars,ROUND($C$8 -($L$" & btmRow & "* $B" & i & ")+$S" & i & "*3,0)), $C$8 -($L$" & btmRow & "*$B" & i & ")+$S" & i & "*3+2, IF(COUNTIF(collars,ROUND($C$8-($L$" & btmRow & " * $B" & i & ")+$S" & i & "*3, 0)+1), $C$8 - ($L$" & btmRow & " * $B" & i & ")+$S" & i & " *3-1, IF(COUNTIF(collars,ROUND($C$8-($L$" & btmRow & " * $B" & i & ")+$S" & i & " *3,0)-1), $C$8-($L$" & btmRow & " * $B" & i & ")+$S" & i & " * 3+1, $C$8-($L$" & btmRow & " * $B" & i & ")+$S" & i & "*3)))"
        Range("O" & i).Formula = "=IF(COUNTIF(collars,ROUND($C$8 -($L$" & btmRow & "* $B" & i & ")+$S" & i & "*2,0)), $C$8 -($L$" & btmRow & "*$B" & i & ")+$S" & i & "*2+2, IF(COUNTIF(collars,ROUND($C$8-($L$" & btmRow & " * $B" & i & ")+$S" & i & "*2, 0)+1), $C$8 - ($L$" & btmRow & " * $B" & i & ")+$S" & i & " *2-1, IF(COUNTIF(collars,ROUND($C$8-($L$" & btmRow & " * $B" & i & ")+$S" & i & " *2,0)-1), $C$8-($L$" & btmRow & " * $B" & i & ")+$S" & i & " * 2+1, $C$8-($L$" & btmRow & " * $B" & i & ")+$S" & i & "*2)))"
        Range("P" & i).Formula = "=IF(COUNTIF(collars,ROUND($C$8 -($L$" & btmRow & "* $B" & i & ")+$S" & i & "*1,0)), $C$8 -($L$" & btmRow & "*$B" & i & ")+$S" & i & "*1+2, IF(COUNTIF(collars,ROUND($C$8-($L$" & btmRow & " * $B" & i & ")+$S" & i & "*1, 0)+1), $C$8 - ($L$" & btmRow & " * $B" & i & ")+$S" & i & " *1-1, IF(COUNTIF(collars,ROUND($C$8-($L$" & btmRow & " * $B" & i & ")+$S" & i & " *1,0)-1), $C$8-($L$" & btmRow & " * $B" & i & ")+$S" & i & " * 1+1, $C$8-($L$" & btmRow & " * $B" & i & ")+$S" & i & "*1)))"
        Range("Q" & i).Formula = "=IF(COUNTIF(collars,ROUND($C$8 -($L$" & btmRow & "* $B" & i & "),0)), $C$8 -($L$" & btmRow & "*$B" & i & ")+2, IF(COUNTIF(collars,ROUND($C$8-($L$" & btmRow & " * $B" & i & "), 0)+1), $C$8 - ($L$" & btmRow & " * $B" & i & ")-1, IF(COUNTIF(collars,ROUND($C$8-($L$" & btmRow & " * $B" & i & "), 0)-1), $C$8-($L$" & btmRow & " * $B" & i & ")+1, $C$8-($L$" & btmRow & " * $B" & i & "))))"
        Range("Q" & i + 1).Formula = "=C" & i + 4
    Else
        'the other rows
        Range("Q" & i).Formula = "=IF(COUNTIF(collars,ROUND($C$8 -($L$" & btmRow & "* $B" & i & ")+15,0)), $C$8 -($L$" & btmRow & "*$B" & i & ")+15+2, IF(COUNTIF(collars,ROUND($C$8-($L$" & btmRow & " * $B" & i & ")+15, 0)+1), $C$8 - ($L$" & btmRow & " * $B" & i & ")+15-1, IF(COUNTIF(collars,ROUND($C$8-($L$" & btmRow & " * $B" & i & ")+15,0)-1), $C$8-($L$" & btmRow & " * $B" & i & ")+15+1, $C$8-($L$" & btmRow & " * $B" & i & ")+15)))"
        Range("Q" & i + 1).Formula = "=C" & i + 3 & "+ 15"
        Range("E" & i).Formula = "=IF(COUNTIF(collars,ROUND($C$8 -($L$" & btmRow & "* $B" & i & ")+15+$S" & i & "*12,0)), $C$8 -($L$" & btmRow & "*$B" & i & ")+15+$S" & i & "*12+2, IF(COUNTIF(collars,ROUND($C$8-($L$" & btmRow & " * $B" & i & ")+15+$S" & i & "*12, 0)+1), $C$8 - ($L$" & btmRow & " * $B" & i & ")+15+$S" & i & " *12-1, IF(COUNTIF(collars,ROUND($C$8-($L$" & btmRow & " * $B" & i & ")+15+$S" & i & " *12,0)-1), $C$8-($L$" & btmRow & " * $B" & i & ")+15+$S" & i & " * 12+1, $C$8-($L$" & btmRow & " * $B" & i & ")+15+$S" & i & "*12)))"
        Range("F" & i).Formula = "=IF(COUNTIF(collars,ROUND($C$8 -($L$" & btmRow & "* $B" & i & ")+15+$S" & i & "*11,0)), $C$8 -($L$" & btmRow & "*$B" & i & ")+15+$S" & i & "*11+2, IF(COUNTIF(collars,ROUND($C$8-($L$" & btmRow & " * $B" & i & ")+15+$S" & i & "*11, 0)+1), $C$8 - ($L$" & btmRow & " * $B" & i & ")+15+$S" & i & " *11-1, IF(COUNTIF(collars,ROUND($C$8-($L$" & btmRow & " * $B" & i & ")+15+$S" & i & " *11,0)-1), $C$8-($L$" & btmRow & " * $B" & i & ")+15+$S" & i & " * 11+1, $C$8-($L$" & btmRow & " * $B" & i & ")+15+$S" & i & "*11)))"
        Range("G" & i).Formula = "=IF(COUNTIF(collars,ROUND($C$8 -($L$" & btmRow & "* $B" & i & ")+15+$S" & i & "*10,0)), $C$8 -($L$" & btmRow & "*$B" & i & ")+15+$S" & i & "*10+2, IF(COUNTIF(collars,ROUND($C$8-($L$" & btmRow & " * $B" & i & ")+15+$S" & i & "*10, 0)+1), $C$8 - ($L$" & btmRow & " * $B" & i & ")+15+$S" & i & " *10-1, IF(COUNTIF(collars,ROUND($C$8-($L$" & btmRow & " * $B" & i & ")+15+$S" & i & " *10,0)-1), $C$8-($L$" & btmRow & " * $B" & i & ")+15+$S" & i & " * 10+1, $C$8-($L$" & btmRow & " * $B" & i & ")+15+$S" & i & "*10)))"
        Range("H" & i).Formula = "=IF(COUNTIF(collars,ROUND($C$8 -($L$" & btmRow & "* $B" & i & ")+15+$S" & i & "*9,0)), $C$8 -($L$" & btmRow & "*$B" & i & ")+15+$S" & i & "*9+2, IF(COUNTIF(collars,ROUND($C$8-($L$" & btmRow & " * $B" & i & ")+15+$S" & i & "*9, 0)+1), $C$8 - ($L$" & btmRow & " * $B" & i & ")+15+$S" & i & " *9-1, IF(COUNTIF(collars,ROUND($C$8-($L$" & btmRow & " * $B" & i & ")+15+$S" & i & " *9,0)-1), $C$8-($L$" & btmRow & " * $B" & i & ")+15+$S" & i & " * 9+1, $C$8-($L$" & btmRow & " * $B" & i & ")+15+$S" & i & "*9)))"
        Range("I" & i).Formula = "=IF(COUNTIF(collars,ROUND($C$8 -($L$" & btmRow & "* $B" & i & ")+15+$S" & i & "*8,0)), $C$8 -($L$" & btmRow & "*$B" & i & ")+15+$S" & i & "*8+2, IF(COUNTIF(collars,ROUND($C$8-($L$" & btmRow & " * $B" & i & ")+15+$S" & i & "*8, 0)+1), $C$8 - ($L$" & btmRow & " * $B" & i & ")+15+$S" & i & " *8-1, IF(COUNTIF(collars,ROUND($C$8-($L$" & btmRow & " * $B" & i & ")+15+$S" & i & " *8,0)-1), $C$8-($L$" & btmRow & " * $B" & i & ")+15+$S" & i & " * 8+1, $C$8-($L$" & btmRow & " * $B" & i & ")+15+$S" & i & "*8)))"
        Range("J" & i).Formula = "=IF(COUNTIF(collars,ROUND($C$8 -($L$" & btmRow & "* $B" & i & ")+15+$S" & i & "*7,0)), $C$8 -($L$" & btmRow & "*$B" & i & ")+15+$S" & i & "*7+2, IF(COUNTIF(collars,ROUND($C$8-($L$" & btmRow & " * $B" & i & ")+15+$S" & i & "*7, 0)+1), $C$8 - ($L$" & btmRow & " * $B" & i & ")+15+$S" & i & " *7-1, IF(COUNTIF(collars,ROUND($C$8-($L$" & btmRow & " * $B" & i & ")+15+$S" & i & " *7,0)-1), $C$8-($L$" & btmRow & " * $B" & i & ")+15+$S" & i & " * 7+1, $C$8-($L$" & btmRow & " * $B" & i & ")+15+$S" & i & "*7)))"
        Range("K" & i).Formula = "=IF(COUNTIF(collars,ROUND($C$8 -($L$" & btmRow & "* $B" & i & ")+15+$S" & i & "*6,0)), $C$8 -($L$" & btmRow & "*$B" & i & ")+15+$S" & i & "*6+2, IF(COUNTIF(collars,ROUND($C$8-($L$" & btmRow & " * $B" & i & ")+15+$S" & i & "*6, 0)+1), $C$8 - ($L$" & btmRow & " * $B" & i & ")+15+$S" & i & " *6-1, IF(COUNTIF(collars,ROUND($C$8-($L$" & btmRow & " * $B" & i & ")+15+$S" & i & " *6,0)-1), $C$8-($L$" & btmRow & " * $B" & i & ")+15+$S" & i & " * 6+1, $C$8-($L$" & btmRow & " * $B" & i & ")+15+$S" & i & "*6)))"
        Range("L" & i).Formula = "=IF(COUNTIF(collars,ROUND($C$8 -($L$" & btmRow & "* $B" & i & ")+15+$S" & i & "*5,0)), $C$8 -($L$" & btmRow & "*$B" & i & ")+15+$S" & i & "*5+2, IF(COUNTIF(collars,ROUND($C$8-($L$" & btmRow & " * $B" & i & ")+15+$S" & i & "*5, 0)+1), $C$8 - ($L$" & btmRow & " * $B" & i & ")+15+$S" & i & " *5-1, IF(COUNTIF(collars,ROUND($C$8-($L$" & btmRow & " * $B" & i & ")+15+$S" & i & " *5,0)-1), $C$8-($L$" & btmRow & " * $B" & i & ")+15+$S" & i & " * 5+1, $C$8-($L$" & btmRow & " * $B" & i & ")+15+$S" & i & "*5)))"
        Range("M" & i).Formula = "=IF(COUNTIF(collars,ROUND($C$8 -($L$" & btmRow & "* $B" & i & ")+15+$S" & i & "*4,0)), $C$8 -($L$" & btmRow & "*$B" & i & ")+15+$S" & i & "*4+2, IF(COUNTIF(collars,ROUND($C$8-($L$" & btmRow & " * $B" & i & ")+15+$S" & i & "*4, 0)+1), $C$8 - ($L$" & btmRow & " * $B" & i & ")+15+$S" & i & " *4-1, IF(COUNTIF(collars,ROUND($C$8-($L$" & btmRow & " * $B" & i & ")+15+$S" & i & " *4,0)-1), $C$8-($L$" & btmRow & " * $B" & i & ")+15+$S" & i & " * 4+1, $C$8-($L$" & btmRow & " * $B" & i & ")+15+$S" & i & "*4)))"
        Range("N" & i).Formula = "=IF(COUNTIF(collars,ROUND($C$8 -($L$" & btmRow & "* $B" & i & ")+15+$S" & i & "*3,0)), $C$8 -($L$" & btmRow & "*$B" & i & ")+15+$S" & i & "*3+2, IF(COUNTIF(collars,ROUND($C$8-($L$" & btmRow & " * $B" & i & ")+15+$S" & i & "*3, 0)+1), $C$8 - ($L$" & btmRow & " * $B" & i & ")+15+$S" & i & " *3-1, IF(COUNTIF(collars,ROUND($C$8-($L$" & btmRow & " * $B" & i & ")+15+$S" & i & " *3,0)-1), $C$8-($L$" & btmRow & " * $B" & i & ")+15+$S" & i & " * 3+1, $C$8-($L$" & btmRow & " * $B" & i & ")+15+$S" & i & "*3)))"
        Range("O" & i).Formula = "=IF(COUNTIF(collars,ROUND($C$8 -($L$" & btmRow & "* $B" & i & ")+15+$S" & i & "*2,0)), $C$8 -($L$" & btmRow & "*$B" & i & ")+15+$S" & i & "*2+2, IF(COUNTIF(collars,ROUND($C$8-($L$" & btmRow & " * $B" & i & ")+15+$S" & i & "*2, 0)+1), $C$8 - ($L$" & btmRow & " * $B" & i & ")+15+$S" & i & " *2-1, IF(COUNTIF(collars,ROUND($C$8-($L$" & btmRow & " * $B" & i & ")+15+$S" & i & " *2,0)-1), $C$8-($L$" & btmRow & " * $B" & i & ")+15+$S" & i & " * 2+1, $C$8-($L$" & btmRow & " * $B" & i & ")+15+$S" & i & "*2)))"
        Range("P" & i).Formula = "=IF(COUNTIF(collars,ROUND($C$8 -($L$" & btmRow & "* $B" & i & ")+15+$S" & i & "*1,0)), $C$8 -($L$" & btmRow & "*$B" & i & ")+15+$S" & i & "*1+2, IF(COUNTIF(collars,ROUND($C$8-($L$" & btmRow & " * $B" & i & ")+15+$S" & i & "*1, 0)+1), $C$8 - ($L$" & btmRow & " * $B" & i & ")+15+$S" & i & " *1-1, IF(COUNTIF(collars,ROUND($C$8-($L$" & btmRow & " * $B" & i & ")+15+$S" & i & " *1,0)-1), $C$8-($L$" & btmRow & " * $B" & i & ")+15+$S" & i & " * 1+1, $C$8-($L$" & btmRow & " * $B" & i & ")+15+$S" & i & "*1)))"
        Range("Q" & i).Formula = "=IF(COUNTIF(collars,ROUND($C$8 -($L$" & btmRow & "* $B" & i & ")+15,0)), $C$8 -($L$" & btmRow & "*$B" & i & ")+15+2, IF(COUNTIF(collars,ROUND($C$8-($L$" & btmRow & " * $B" & i & ")+15, 0)+1), $C$8 - ($L$" & btmRow & " * $B" & i & ")+15-1, IF(COUNTIF(collars,ROUND($C$8-($L$" & btmRow & " * $B" & i & ")+15,0)-1), $C$8-($L$" & btmRow & " * $B" & i & ")+15+1, $C$8-($L$" & btmRow & " * $B" & i & ")+15)))"
        Range("Q" & i + 1).Formula = "=C" & i + 3 & "+ 15"
    End If
    Range("R" & i).Formula = "=SUM($D$7:$Q$7)"
    Range("S" & i).Formula = "=(D" & i & "-Q" & i & ")/(COUNT($D$7:$Q$7)-1)"
    j = j + 1
Next i
' deal with inconsistant formula errors
Dim xCell As Range, xTarget As Range
    Set xTarget = Range("C8:Q" & btmRow)
    For Each xCell In xTarget
        xCell.Errors(xlInconsistentFormula).Ignore = True
    Next



MsgBox ActiveSheet.Name
Application.EnableEvents = True
Application.ScreenUpdating = True
Application.DisplayStatusBar = True
Application.Calculation = xlCalculationAutomatic

End Sub