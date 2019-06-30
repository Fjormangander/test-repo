Attribute VB_Name = "Module1"
Option Explicit

Sub GetData()

    'Константы по созданию файла

    Dim strFile As String
    Dim wbkX As Workbook, wbkY As Workbook
    Dim shtX As Worksheet, shtY As Worksheet, shtZ As Worksheet
    Dim rngPaste As Range, rngCopy As Range, rngAll As Range
    Dim N1 As Long, N2 As Long
    Dim fd As FileDialog
    Dim strX As String
    Dim Question As Integer
    
    'Константы по созданию письма
    
    Dim objOutlookApp As Object, objMail As Object
    Dim sTo As String, sCC As String, sSubject As String, sBody As String, sAttachment As String, strAdress As String
    
    With Application
    
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
        .DisplayAlerts = False
    
    End With
    
    Set wbkX = ThisWorkbook
    Set shtX = wbkX.Sheets(">>DATA")
    Set shtY = wbkX.Sheets(">>SET")
    Set shtZ = wbkX.Sheets("СВОД")
        
    If shtX.FilterMode Then
        shtX.ShowAllData
    End If
    
    Set rngAll = shtX.Range("A1").CurrentRegion
       
    shtX.Range("A3:Q1048576").Clear
      
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
    With fd
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Файлы Excel", "*.xls;*.xlsx"
        '.InitialFileName = "C:\Users..."
        .Show
   End With
   
   If fd.SelectedItems.Count = 0 Then
      Exit Sub
   End If
   
   strFile = fd.SelectedItems(1)
      
   Set wbkY = Application.Workbooks.Open(strFile)
   
   Set rngCopy = wbkY.Sheets(1).Range("A2").CurrentRegion
   N1 = rngCopy.Rows.Count
   Set rngCopy = rngCopy.Range("A2").Resize(N1 - 1, 17)
   rngCopy.Copy
   
   Set rngPaste = shtX.Range("A3")
   
   rngPaste.PasteSpecial
   shtX.Rows("2:2").Delete Shift:=xlUp
   wbkY.Close False
   
   Set rngAll = shtX.Range("A1").CurrentRegion
   
   With shtZ
   
        .Activate
        .Calculate
        .Range("A2").Select
        
   End With
   
   shtY.Calculate
   wbkX.Save
   
   Question = MsgBox("Данные были успешно обновлены. Сформировать письмо на отправку?", vbYesNo + vbQuestion)
   
   Select Case Question
   
        Case vbYes
            
            Set objOutlookApp = GetObject(, "Outlook.Application")
            Err.Clear
    
            If objOutlookApp Is Nothing Then
                Set objOutlookApp = CreateObject("Outlook.Application")
            End If
    
            objOutlookApp.Session.Logon
   
            Set objMail = objOutlookApp.CreateItem(0)
   
            If Err.Number <> 0 Then Set objOutlookApp = Nothing: Set objMail = Nothing: Exit Sub
            
            'If Timer() > 59200 Then
                strX = " сегодня."
            'Else
            '    strX = " вчера."
           ' End If
   
            sTo = shtY.Range("J12").Value
            sCC = shtY.Range("J13").Value
            sSubject = shtZ.Range("A1").Value
            sBody = "Добрый день, Коллеги!" & vbCrLf & "Направляю актуальный статус по исполнению обращений ЦБ РФ с крайним сроком -" & strX
   
            'Формируем копию файла:
    
            shtY.Calculate
   
            Set rngAll = shtY.Range("J10")
            strX = rngAll.Value
   
            strAdress = "\\hq.rgs.ru\users\RGS\Департамент взаимодействия с регуляторами\05_Методология и анализ\Отчетность\Отчеты 2018\ЦБ\Статусы исполнения заявок\" & strX & "Статус исполнения обращений.xlsx"
    
            ChDir "\\hq.rgs.ru\users\RGS\Департамент взаимодействия с регуляторами\05_Методология и анализ\Отчетность\Отчеты 2018\ЦБ\Статусы исполнения заявок"
            
            wbkX.SaveAs Filename:=strAdress, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
       
            sAttachment = strAdress
   
            With objMail
    
                .To = sTo 'aa?an iieo?aoaey
                .CC = sCC 'aa?an aey eiiee
                '.BCC = "vladislav_belov@rgs.ru"
                .Subject = sSubject 'oaia niiauaiey
                .Body = sBody 'oaeno niiauaiey
                '.HTMLBody = sBody 'anee iaiaoiaei oi?iaoe?iaaiiua oaeno niiauaiey(?acee?iua o?eoou, oaao o?eooa e o.i.)
                .Attachments.Add sAttachment '?oiau ioi?aaeou aeoeaio? eieao aianoi sAttachment oeacaou ActiveWorkbook.FullName
                '.Send
                .Display ', anee iaiaoiaeii i?iniio?aou niiauaiea, a ia ioi?aaeyou aac i?iniio?a
    
            End With
 
            Set objOutlookApp = Nothing: Set objMail = Nothing
       
       
            With Application
    
                .ScreenUpdating = True
                .DisplayAlerts = True
    
            End With
   
            'MsgBox "Обработка данных завершена, проверьте ошибки!", vbExclamation
            
            Set wbkX = ThisWorkbook
            
            wbkX.Close True
            
    
        Case vbNo
            MsgBox "Отчет обновлен, формирование письма отменено.", vbExclamation
            Exit Sub
        
        End Select
    

End Sub

Private Sub SaveAs_and_MailTo()

    'Константы по созданию файла
    
    Dim wbkX As Workbook
    Dim strX As String, strAdress As String
    Dim rngX As Range
     
    'Константы по созданию письма
    
    Dim objOutlookApp As Object, objMail As Object
    Dim sTo As String, sCC As String, sSubject As String, sBody As String, sAttachment As String
            
    Set wbkX = ThisWorkbook
        
    Set objOutlookApp = GetObject(, "Outlook.Application")
    Err.Clear
    
    If objOutlookApp Is Nothing Then
    
        Set objOutlookApp = CreateObject("Outlook.Application")
    
    End If
    
    objOutlookApp.Session.Logon
    
    Set objMail = objOutlookApp.CreateItem(0)
    
    If Err.Number <> 0 Then Set objOutlookApp = Nothing: Set objMail = Nothing: Exit Sub
    
    sTo = wbkX.Worksheets("настройки").Range("J12").Value
    sCC = wbkX.Worksheets("настройки").Range("J13").Value
    sSubject = wbkX.Worksheets("СВОД").Range("A1").Value
    
    sBody = "Отчет сформирован автоматически, копия сохранена в базовой директории."
    
    wbkX.Worksheets("настройки").Calculate
    
    Set rngX = wbkX.Worksheets("настройки").Range("J9")
    
    strX = rngX.Value
    
    strAdress = "\\hq.rgs.ru\users\RGS\Департамент взаимодействия с регуляторами\Методология\Отчетность\Отчеты 2018\ЦБ\Статусы исполнения заявок\Статус исполнения обращений_" & strX & ".xlsx"
    
    ChDir "\\hq.rgs.ru\users\RGS\Департамент взаимодействия с регуляторами\Методология\Отчетность\Отчеты 2018\ЦБ\Статусы исполнения заявок"
    
    wbkX.SaveAs Filename:=strAdress, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    
    MsgBox "Копия отчета сохранена!", vbExclamation
    
    sAttachment = strAdress
    
    With objMail
    
        .To = sTo 'aa?an iieo?aoaey
        .CC = sCC 'aa?an aey eiiee
        .BCC = "vladislav_belov@rgs.ru" 'aa?an aey ne?uoie eiiee
        .Subject = sSubject 'oaia niiauaiey
        .Body = sBody 'oaeno niiauaiey
        '.HTMLBody = sBody 'anee iaiaoiaei oi?iaoe?iaaiiua oaeno niiauaiey(?acee?iua o?eoou, oaao o?eooa e o.i.)
        .Attachments.Add sAttachment '?oiau ioi?aaeou aeoeaio? eieao aianoi sAttachment oeacaou ActiveWorkbook.FullName
       ' .Send
       .Display ', anee iaiaoiaeii i?iniio?aou niiauaiea, a ia ioi?aaeyou aac i?iniio?a
    
    End With
 
    Set objOutlookApp = Nothing: Set objMail = Nothing
    Application.ScreenUpdating = True
    
    
    
    
End Sub

