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
    toStr = "[宛先のメールアドレス]"
    ccStr = "[CCのメールアドレス]"
    bccStr = "[BCCのメールアドレス]"
    
    '---  ---'
    subjectStr = "[件名]"
    
    '---  ---'
    bodyStr = "[本文]"
        
    '---  ---'
    objMail.To = toStr
    objMail.CC = ccStr
    objMail.BCC = bccStr
    objMail.Subject = subjectStr
    objMail.Body = bodyStr
    
    '---  ---'
    objMail.Display
    
    
    '--- Excelワークシート ---'
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("[ワークシート名]")
    
    '--- 貼り付ける表の範囲（A1:H10のような文字列） ---'
    Dim tableAddress As String
    tableAddress = "[表のアドレス]"
    
    '--- メール本文に表を貼り付け ---'
    Call ws.Range(tableAddress).Copy
    objMail.GetInspector().WordEditor.Windows(1).Selection.Paste
    
    '--- メールを送付 ---'
    objMail.Send
    
End Sub
