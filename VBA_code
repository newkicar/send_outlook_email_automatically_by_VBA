Option Explicit

Sub SendEmails()

    Dim ok As String
    
    ok = MsgBox("Are you sure you want to send all emails in the table?", vbYesNo, "Bulk Email Assistant")
    If ok <> "6" Then Exit Sub
    
    'Check quantity, maximum 100 emails per batch to avoid server overload
    If Cells(Rows.Count, 2).End(xlUp).Row > 100 Then MsgBox "To avoid excessive server load, the maximum number of emails per batch is 100!": Exit Sub
    
    'Check if emails with content have recipients
    Dim standard As Integer, prompt As String
    prompt = "Please note:" & Chr(10)
    For standard = 2 To WorksheetFunction.Max(Cells(Rows.Count, 2).End(xlUp).Row, _
    Cells(Rows.Count, 3).End(xlUp).Row, Cells(Rows.Count, 4).End(xlUp).Row, _
    Cells(Rows.Count, 5).End(xlUp).Row, Cells(Rows.Count, 6).End(xlUp).Row)
        'Recipient is empty
        If Cells(standard, 2) = "" Then
            prompt = prompt & "Row " & standard & " has empty recipient field!" & Chr(10)
        End If
        'Subject is empty
        If Cells(standard, 5) = "" Then
            prompt = prompt & "Row " & standard & " has empty subject field!" & Chr(10)
        End If
        'Body is empty
        If Cells(standard, 6) = "" And Cells(standard, 7) = "" Then
            prompt = prompt & "Row " & standard & " has empty body and body image fields!" & Chr(10)
        End If
    Next standard
    If Len(prompt) > 5 Then MsgBox prompt & "Email sending aborted at error rows, please correct!", , "Error": Exit Sub
    

    Dim olApp As Object
    Dim olMail As Object
    Dim ws As Worksheet
    Dim rng As Range
    Dim recipient As String
    Dim ccRecipient As String
    Dim attachmentPaths As String
    Dim subject As String
    Dim body As String
    Dim cell
    Dim i As Integer
    Dim SigString As String
    Dim sig As String
    Dim s As String, s0 As String
    Dim pic As String
    
    ' Set Outlook application object
    Set olApp = CreateObject("Outlook.Application")
    
    ' Set worksheet and data range
    Set ws = ThisWorkbook.ActiveSheet ' Replace with your worksheet name
    Set rng = ws.Range("A2:F" & ws.Cells(ws.Rows.Count, 2).End(xlUp).Row) ' Replace with your data range

    sig = ws.Cells(1, 9).Value      'Reference signature
    
    ' Loop through data range to send emails
    For Each cell In rng.Rows
        recipient = cell.Cells(1, 2).Value
        ccRecipient = cell.Cells(1, 3).Value
        attachmentPaths = cell.Cells(1, 4).Value
        subject = cell.Cells(1, 5).Value
        body = cell.Cells(1, 6).Value
        pic = cell.Cells(1, 7).Value
        s = ""
        s0 = ""
        
        ' Create new email
        Set olMail = olApp.CreateItem(0)
        
        ' Set recipient, CC, subject and body
        olMail.To = recipient
        If ccRecipient <> "" Then
            olMail.CC = ccRecipient
        End If
        olMail.subject = subject
        
        'Add body
        s = "<P STYLE='font-family:Microsoft YaHei;font-size:14;color: rgb(0, 0, 0);line-height=0.1'>"
        For i = 0 To UBound(Split(body, Chr(10)))
            If Split(body, Chr(10))(i) <> "" Then
                s = s & Split(body, Chr(10))(i) & "<br/>"
            End If
        Next i

        'Insert body image (if any)
        For i = 0 To UBound(Split(pic, Chr(10)))
            If Split(pic, Chr(10))(i) <> "" Then
                s0 = s0 & "<img src='image path placeholder' width=1000>" & "<p></p>"
                s0 = Replace(s0, "image path placeholder", Split(pic, Chr(10))(i))
            End If
        Next i

        s = s & s0
        
        s = s & "</p>" & "<p></p>" & "<p></p>" & "<P STYLE='font-family:Microsoft YaHei;font-size:10;color: rgb(0, 0, 0);line-height=0.1'>"
        
        'Add signature
        For i = 0 To UBound(Split(sig, Chr(10)))
            If Split(sig, Chr(10))(i) <> "" Then
                s = s & Split(sig, Chr(10))(i) & "<br/>"
            End If
        Next i
        
        s = s & "</p>"
        
        olMail.HTMLBody = s

        ' Add attachments (if any)
        For i = 0 To UBound(Split(attachmentPaths, Chr(10)))
            If Split(attachmentPaths, Chr(10))(i) <> "" Then
                olMail.Attachments.Add Split(attachmentPaths, Chr(10))(i)
            End If
        Next i

        ' Send email
        olMail.Send
        
        ' Release mail object
        Set olMail = Nothing
    Next cell
    
    ' Release Outlook application object
    Set olApp = Nothing
    
    MsgBox "Email sending completed!"
End Sub
