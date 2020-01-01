Dim path1 As String
Dim path2 As String

Dim book1 As String
Dim book2 As String

Dim status1 As String
Dim status2 As String


Sub BalanceCheck1()

path1 = Range("C3").Value
book1 = Range("C4").Value
path2 = Range("C7").Value
book2 = Range("C8").Value
    
    Call OpenBalance1(path1, book1, status1)
    Call OpenBalance2(path2, book2, status2)
            
    Dim done As String
    done = "No.1 " & status1 & vbCr & _
           "No.2 " & status2 & vbCr & _
           "No.3 " & status7 & vbCr
           
    MsgBox done & vbCr & "OK_execution", vbYes
    
End Sub


Function OpenBalance1(ByRef path1 As String, ByRef book1 As String, status1)

If Dir(path1 & book1) <> "" Then
    Range("C5").Value = "OK"
    Else
    Range("C5").Value = "none"
   
End If

    status1 = Range("C5").Value

End Function

Function OpenBalance2(ByRef path2 As String, ByRef book2 As String, status2)

If Dir(path2 & book2) <> "" Then
    Range("C9").Value = "OK"
    Else
    Range("C9").Value = "none"
   
End If

    status2 = Range("C9").Value
    

End Function

Sub BalanceExe1()
    
path1 = Range("C3").Value
book1 = Range("C4").Value
path2 = Range("C7").Value
book2 = Range("C8").Value
   
status1 = Range("C5").Value
status2 = Range("C9").Value
    
exe1 = Range("a5").Value
exe2 = Range("a9").Value
    
    Dim rc As String
    Dim done As String
    done = "No.1 " & exe1 & vbCr & _
           "No.2 " & exe2 & vbCr & _
           "No.3 " & exe7 & vbCr
           
    rc = MsgBox(done & vbCr & " TRUE_execution", vbYesNo)
    
    If rc = vbNo Then
       GoTo finish
    Else
    End If
    
If status1 = "OK" And Range("A5") = True Then
    Workbooks.Open path1 & book1
        Call BalanceCopy1(book1)
End If

If status2 = "OK" And Range("A9") = True Then
    Workbooks.Open path2 & book2
        Call BalanceCopy2(book2)
End If

Exit Sub

finish:


End Sub

Public Sub BalanceCopy1(ByRef book1 As String)
  
    Workbooks(book1).Activate
    Worksheets("N").Select
    Range("B2:B13").Select
    Selection.Copy
    
    Worksheets("J").Select
    Range("C2").Select
    Selection.PasteSpecial Paste:=xlPasteValues
    
    Workbooks(book1).Save
    Workbooks(book1).Close
    
    
End Sub

Public Sub BalanceCopy2(ByRef book2 As String)
  
    Workbooks(book2).Activate
    Worksheets("N").Select
    Range("B2:B13").Select
    Selection.Copy
    
    Worksheets("J").Select
    Range("C2").Select
    Selection.PasteSpecial Paste:=xlPasteValues
    
    Workbooks(book2).Save
    Workbooks(book2).Close
    
    
End Sub


Sub done_1()
    
path1 = Range("C3").Value
book1 = Range("C4").Value
status1 = Range("C5").Value
exe1 = Range("a5").Value
    
    Call BalanceCheck
    
If status1 = "OK" And Range("A5") = True Then
    Workbooks.Open path1 & book1
        Call BalanceCopy1(book1)
End If

End Sub


------------------------------------------------------------





Sub balance()

Dim path As String
Dim book As String
Dim status As String

path = Range("C10").Value
book = Range("C11").Value
    
    Call OpenBalance(path, book, status)
    
If status = "OK" Then
    Workbooks.Open path & book
    Else
    MsgBox "stop"
End If

End Sub


Function OpenBalance(ByRef path As String, ByRef book As String, status)

If Dir(path & book) <> "" Then
    Range("C12").Value = "OK"
    Else
    Range("C12").Value = "none"
   
End If

    status = Range("C12").Value

End Function






Sub consumption()

     Dim str As String
     Dim price As Long
   
    price = Range("c4").Value
    str = tax(price)

 Range("C6").Value = str

 End Sub

Function tax(price)

     'Dim tax As Long     
     tax = price * 0.1

End Function








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

    'kaiten-suru  5x5
    
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
    If Range(H7).Value = 1 Then
    Exit Sub
    End If
   
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

Sub sendmail_sample1()

'---code1｜outlook
    Dim toaddress, ccaddress, bccaddress As String  
    Dim subject, mailBody, credit As String 
    Dim outlookObj As Outlook.Application    
    Dim mailItemObj As Outlook.MailItem      
    
'---code2｜
    toaddress = Range("B2").Value   
    ccaddress = Range("B3").Value   
    bccaddress = Range("B4").Value  
    subject = Range("B5").Value    
    mailBody = Range("B6").Value    
    credit = Range("B7").Value      

'---code3｜
    Set outlookObj = CreateObject("Outlook.Application")
    Set mailItemObj = outlookObj.CreateItem(olMailItem)
    mailItemObj.BodyFormat = 3      
    mailItemObj.To = toaddress      
    mailItemObj.CC = ccaddress      
    mailItemObj.BCC = bccaddress    
    mailItemObj.subject = subject   
    
'---code4｜
    mailItemObj.Body = mailBody & vbCrLf & vbCrLf & credit   
    
'---code5｜
    Dim attached As String
    Dim myattachments As Outlook.Attachments 
    Set myattachments = mailItemObj.Attachments
    attached = Range("B9").Value     
    myattachments.Add attached

'---code6｜
    'mailItemObj.Save   
    mailItemObj.Display  

'---code7｜
    Set outlookObj = Nothing
    et mailItemObj = Nothing

End Sub

http://www.fingeneersblog.com/1778/

Public Sub CreateMailWithTable()
    
    '---  ---'
    Dim objOutlook As Outlook.Application
    Set objOutlook = New Outlook.Application
    Dim objMail As Outlook.MailItem
    Set objMail = objOutlook.CreateItem(olMailItem)
        
    '---  ---'
'    Dim objOutlook As Object
'    Set objOutlook = CreateObject("Outlook.Application")
'    Dim objMail As Object
'    Set objMail = objOutlook.CreateItem(0)
        
    '---  ---'
    Dim toStr As String
    Dim ccStr As String
    Dim bccStr As String
    Dim subjectStr As String
    Dim bodyStr As String
    
    '---  ---'
    toStr = "[to mail address]"
    ccStr = "[cc mail address]"
    bccStr = "[vbb mail address]"
    
    '---  ---'
    subjectStr = "[subject]"
    
    '---  ---'
    bodyStr = "[body]"
        
    '---  ---'
    objMail.To = toStr
    objMail.CC = ccStr
    objMail.BCC = bccStr
    objMail.Subject = subjectStr
    objMail.Body = bodyStr
    
    '---  ---'
    objMail.Display
    
    
    '--- Excel worksheet ---'
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("[worksheet]")
    
    '--- attachment range（A1:H10） ---'
    Dim tableAddress As String
    tableAddress = "[table address]"
    
    '--- paste mail ---'
    Call ws.Range(tableAddress).Copy
    objMail.GetInspector().WordEditor.Windows(1).Selection.Paste
    
    '--- send mail ---'
    objMail.Send
    
End Sub



Option Explicit

Public Sub Sample()
  Dim app As Object
  Dim doc As Object 'Documentオブジェクト(Word)
  
  Const olMailItem = 0
  Const olImportanceHigh = 2
  Const olFormatRichText = 3
  Const wdUnderlineSingle = 1
  Const wdColorAutomatic = -16777216
  
  Set app = CreateObject("Outlook.Application")
  With app.CreateItem(olMailItem)
    .Display
    .BodyFormat = olFormatRichText
    .To = "aaa@com"
    .CC = "bbb@com"
    .Importance = olImportanceHigh
    .Subject = "test"
    Set doc = .GetInspector.WordEditor
  End With
  
  'コピー&ペースト
  ActiveWorkbook.Worksheets("Sheet1").Range("A1:A10").Copy
  doc.Characters.Last.Paste
  
  '文字列挿入
  With doc.Characters.Last
    'フォント設定
    .Font.NameFarEast = "メイリオ"
    .Font.NameAscii = "メイリオ"
    .Font.NameOther = "メイリオ"
    .Font.Name = "メイリオ"
    .Font.Size = 14
    .Font.Color = vbRed
    .Font.Bold = False
    .Font.Italic = False
    .Font.Underline = wdUnderlineSingle
    .Font.UnderlineColor = wdColorAutomatic
    .InsertBefore "あいうえお" & vbCr '文字列挿入
  End With
  
  'コピー&ペースト
  ActiveWorkbook.Worksheets("Sheet2").Range("A1:A10").Copy
  doc.Characters.Last.Paste
End Sub

