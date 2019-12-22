----

 Sub ExcelCompare()

     Dim i As Long
     Dim j As Long
     Sheets("table").Select

     For j = 1 To 10
     For i = 1 To 30
        If Worksheets("table").Cells(i, j) = Worksheets("form").Cells(i, j) Then
           Worksheets("comp").Cells(i, j) = ""
               Else
           Worksheets("comp").Cells(i, j) = 1
        End If
     Next
 Next
Sheets("comp").Select

end sub

sample

Sub kaiten()

    'kaitensuru
    
    MsgBox "hajimaru"
    
    Range("D2").Activate
    Application.Wait Now() + TimeValue("00:00:01")
    Range("e2").Activate
    Application.Wait Now() + TimeValue("00:00:01")
    Range("f2").Activate
    Application.Wait Now() + TimeValue("00:00:01")
    Range("g2").Activate
    Application.Wait Now() + TimeValue("00:00:01")
    Range("g3").Activate
    Application.Wait Now() + TimeValue("00:00:01")
    Range("g4").Activate
    Application.Wait Now() + TimeValue("00:00:01")
    Range("g5").Activate
    Application.Wait Now() + TimeValue("00:00:01")
    Range("g6").Activate
    Application.Wait Now() + TimeValue("00:00:01")
    Range("g7").Activate
    
    Range("g7").Value = 6
        
    Application.Wait Now() + TimeValue("00:00:01")
    Range("f7").Activate
    Application.Wait Now() + TimeValue("00:00:01")
    Range("e7").Activate
    Application.Wait Now() + TimeValue("00:00:01")
    Range("d7").Activate
    Application.Wait Now() + TimeValue("00:00:01")
    Range("d6").Activate
    Application.Wait Now() + TimeValue("00:00:01")
    Range("d5").Activate
    Application.Wait Now() + TimeValue("00:00:01")
    Range("d4").Activate
    Application.Wait Now() + TimeValue("00:00:01")
    Range("d3").Activate
    Application.Wait Now() + TimeValue("00:00:01")
    Range("d2").Activate
          
    MsgBox "owaru"
        
        
        
End Sub


Sub OpenFile()
'
    Dim buf As String, cnt As Long
    Const Path As String = "C:\Users\tm2\Desktop\hogehogecsv\"
    
    buf = Dir(Path & "*.csv")
    
    Do While buf <> ""
        cnt = cnt + 1
        Cells(cnt, 1) = buf
        buf = Dir()
    Loop

  Dim i As Long
    
   For i = 1 To 10
   
    Cells(i, 2) = Replace(Cells(i, 1), "aa", "")
    Cells(i, 3) = Replace(Cells(i, 2), ".csv", "")
    Cells(i, 4) = Mid(Cells(i, 3), 2, 4) & Mid(Cells(i, 3), 6, 2) & Mid(Cells(i, 3), 8, 2)
      
  Next

    Range("G1").Formula = Year(Range("F1")) * 10000 + Month(Range("F1")) * 100 + Day(Range("F1"))
     
    Range("e1").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-1]=R1C7,1,0)"
    Range("E1").Select
    Selection.Copy
    Range("E1:E4").Select
    ActiveSheet.Paste
  
  
  Range ("C3")
     
  Dim j As Long
  Dim file1 As String
   
   For j = 1 To 10
   
    If Cells(j, 5) = 1 Then
    
       file1 = Cells(j, 1).Value
       Workbooks.Open "C:\Users\tm2\Desktop\hogehogecsv\" & file1

    Else
    End If
    
  Next
    
End Sub



Sub kaiten2()

    'kaiten-suru
    
    MsgBox "hajimaruyo"
    
        
    'right
    Range("D2").Select
    Application.Wait Now() + TimeValue("00:00:01")
    Range("E2").Select
    Application.Wait Now() + TimeValue("00:00:01")
    Range("F2").Select
    Application.Wait Now() + TimeValue("00:00:01")
    Range("G2").Select
    Application.Wait Now() + TimeValue("00:00:01")
    Range("H2").Select
    Application.Wait Now() + TimeValue("00:00:01")
    
    'down
    Range("H2").Select
    Application.Wait Now() + TimeValue("00:00:01")
    Range("H3").Select
    Application.Wait Now() + TimeValue("00:00:01")
    Range("H4").Select
    Application.Wait Now() + TimeValue("00:00:01")
    Range("H5").Select
    Application.Wait Now() + TimeValue("00:00:01")
    Range("H6").Select
    Application.Wait Now() + TimeValue("00:00:01")
    Range("H7").Select
    Application.Wait Now() + TimeValue("00:00:01")
    
    'bunki-syori
    Call Bunki
    
    'left
    Range("G7").Select
    Application.Wait Now() + TimeValue("00:00:01")
    Range("F7").Select
    Application.Wait Now() + TimeValue("00:00:01")
    Range("E7").Select
    Application.Wait Now() + TimeValue("00:00:01")
    Range("D7").Select
    Application.Wait Now() + TimeValue("00:00:01")
    
    
    'up
    Range("D6").Select
    Application.Wait Now() + TimeValue("00:00:01")
    Range("D5").Select
    Application.Wait Now() + TimeValue("00:00:01")
    Range("D4").Select
    Application.Wait Now() + TimeValue("00:00:01")
    Range("D3").Select
    Application.Wait Now() + TimeValue("00:00:01")
    Range("D2").Select
    Application.Wait Now() + TimeValue("00:00:01")
   
            
    MsgBox "owattayo"
        
          
End Sub


Sub Bunki()

    If Range("H7") = 1 Then
    
   'right
    Range("H7").Select
    Application.Wait Now() + TimeValue("00:00:01")
    Range("I7").Select
    Application.Wait Now() + TimeValue("00:00:01")
    Range("J7").Select
    Application.Wait Now() + TimeValue("00:00:01")
    Range("K7").Select
    Application.Wait Now() + TimeValue("00:00:01")
    Range("L7").Select
    Application.Wait Now() + TimeValue("00:00:01")
    
    'down
    Range("L8").Select
    Application.Wait Now() + TimeValue("00:00:01")
    Range("L9").Select
    Application.Wait Now() + TimeValue("00:00:01")
    Range("L10").Select
    Application.Wait Now() + TimeValue("00:00:01")
    Range("L11").Select
    Application.Wait Now() + TimeValue("00:00:01")
    Range("L12").Select
    Application.Wait Now() + TimeValue("00:00:01")
        
    'left
    Range("K12").Select
    Application.Wait Now() + TimeValue("00:00:01")
    Range("J12").Select
    Application.Wait Now() + TimeValue("00:00:01")
    Range("I12").Select
    Application.Wait Now() + TimeValue("00:00:01")
    Range("H12").Select
    Application.Wait Now() + TimeValue("00:00:01")
    
    
    'up
    Range("H11").Select
    Application.Wait Now() + TimeValue("00:00:01")
    Range("H10").Select
    Application.Wait Now() + TimeValue("00:00:01")
    Range("H9").Select
    Application.Wait Now() + TimeValue("00:00:01")
    Range("H8").Select
    Application.Wait Now() + TimeValue("00:00:01")
    Range("H7").Select
    Application.Wait Now() + TimeValue("00:00:01")
   
    Else
    End If
    
    Call bunki2
    
End Sub

Sub bunki2()


    'left
    Range("G7").Select
    Application.Wait Now() + TimeValue("00:00:01")
    Range("F7").Select
    Application.Wait Now() + TimeValue("00:00:01")
    Range("E7").Select
    Application.Wait Now() + TimeValue("00:00:01")
    Range("D7").Select
    Application.Wait Now() + TimeValue("00:00:01")
    
    
    'up
    Range("D6").Select
    Application.Wait Now() + TimeValue("00:00:01")
    Range("D5").Select
    Application.Wait Now() + TimeValue("00:00:01")
    Range("D4").Select
    Application.Wait Now() + TimeValue("00:00:01")
    Range("D3").Select
    Application.Wait Now() + TimeValue("00:00:01")
    Range("D2").Select
    Application.Wait Now() + TimeValue("00:00:01")
   
            
    MsgBox "owatta"
       
End Sub

