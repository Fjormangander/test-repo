Attribute VB_Name = "Module3"
Option Explicit

Sub GetData()

    Dim strFile As String
    Dim wbkX As Workbook, wbkY As Workbook
    Dim shtX As Worksheet
    Dim rngPaste As Range, rngCopy As Range, rngAll As Range
    Dim N1 As Long, N2 As Long
    Dim fd As FileDialog
    
    With Application
    
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
        .DisplayAlerts = False
    
    End With
    
    Set wbkX = ThisWorkbook
    Set shtX = wbkX.Sheets(">>DATA")
        
    If shtX.FilterMode Then
            shtX.ShowAllData
    End If
        
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
    With fd
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "����� Excel", "*.xls;*.xlsx"
        '.InitialFileName = "C:\Users..."
        .Show
    End With
   
    If fd.SelectedItems.Count = 0 Then
        Exit Sub
    End If
   
    Set rngAll = shtX.Range("A1").CurrentRegion
       
    shtX.Range("A3:AF1048576").Clear
      
    strFile = fd.SelectedItems(1)
      
    Set wbkY = Application.Workbooks.Open(strFile)
   
    Set rngCopy = wbkY.Sheets(1).Range("A2").CurrentRegion
    N1 = rngCopy.Rows.Count
    Set rngCopy = rngCopy.Range("A2").Resize(N1 - 1, 24)
    rngCopy.Copy
   
    Set rngPaste = shtX.Range("A3")
   
    rngPaste.PasteSpecial
    wbkY.Close False
   
    Set rngAll = shtX.Range("A1").CurrentRegion
    N2 = rngAll.Rows.Count
    shtX.Range("Y2:AF" & N2).FillDown
      
    shtX.Rows("2:2").Delete Shift:=xlUp
    shtX.Calculate
    
    With Application
    
        .ScreenUpdating = True
        .DisplayAlerts = True
    
    End With
   
    MsgBox "��������� ������ ���������, ��������� ������!", vbExclamation

End Sub

Sub Main()

    'NtY: �������� ��������� ����� ����� ����������!
    
    Dim wbkX As Workbook
    
    Dim shtX As Worksheet
    Dim rngAll As Range
    Dim X1 As Long, Y1 As Long
    
    Dim strX As String, strAdress As String, strFile As String
    
    Dim Names_Array
    Dim g As Integer, i As Integer, Result As Integer
        
    With Application
    
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
        .DisplayAlerts = False
    
    End With
    
    Set wbkX = ThisWorkbook
    Dim wks As Worksheet
    
    For Each wks In wbkX.Worksheets
        wks.Calculate
    Next wks
       
   Result = MsgBox("�������� ������� ���������. ������������ pdf �����?", vbYesNo + vbQuestion)
    
   Select Case Result
        
        Case vbYes
                
        strX = wbkX.Sheets(">>CALC").Range("M2").Value
        strX = strX & "_������������ ����� ����� (���� + ��)"
        
        strAdress = _
"\\hq.rgs.ru\users\RGS\����������� �������������� � ������������\05_����������� � ������\����������\������ 2019\��\����� �� �������� � ������������ (���-��)\" & strX & ".pdf"
        
       
    wbkX.ExportAsFixedFormat Type:=xlTypePDF, Filename:=strAdress _
        , Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
        :=False, From:=1, To:=22, OpenAfterPublish:=True 'False True
    Application.WindowState = xlNormal
        

     Case vbNo
     MsgBox "�������� ���������!", vbExclamation
        
   End Select
    
    With Application
    
        .ScreenUpdating = True
        .DisplayAlerts = True
    
    End With
       
    MsgBox "�������� ���������!", vbExclamation
   
            
End Sub

Sub Send_email()

'
   
   Dim Question As Integer
   Dim wbkX As Workbook
   Dim shtX As Worksheet
   Dim strX As String
   
   Dim objOutlookApp As Object, objMail As Object
   Dim sTo As String, sCC As String, sSubject As String, sBody As String, sAttachment As String, strAdress As String
   
   
   Set wbkX = ThisWorkbook
   Set shtX = wbkX.Sheets(">>CALC")
   strX = shtX.Range("M3").Value
      
   Question = MsgBox("������ ���� ������� ���������. ������������ ������ �� ��������?", vbYesNo + vbQuestion)
   
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
   
            sTo = wbkX.Sheets(">>SET").Range("V3").Value
            sCC = wbkX.Sheets(">>SET").Range("V4").Value
            sSubject = "����� �� �������� � ������������ �� ������ " & strX
            
            sBody = "<p>������ ����, �������!</p>" & _
                    "��������� �������� ������ ��� ������ �� ������ � ��������� � ������������ �� �� �� ��������� ������ (�� ��������� �� 10-30). " & _
                    "� ���������� ��������� ����� �� ������������ � ������ 2019 �. � ������� ����� �������� � ����������� �� ��.<br>"
            
          strX = wbkX.Sheets(">>CALC").Range("M2").Value
        strX = strX & "_������������ ����� ����� (���� + ��)"
          
          strAdress = _
"\\hq.rgs.ru\users\RGS\����������� �������������� � ������������\05_����������� � ������\����������\������ 2019\��\����� �� �������� � ������������ (���-��)\" & strX & ".pdf"
            
            sAttachment = strAdress
            
            
                                               
            With objMail
    
                .To = sTo
                .CC = sCC
                .BCC = "vladislav_belov@rgs.ru"
                .Subject = sSubject
                '.Body = sBody
                .HTMLBody = sBody '& "<br><br>" & Signature
                .Attachments.Add sAttachment
                '.Send
                .Display
    
            End With
 
            Set objOutlookApp = Nothing: Set objMail = Nothing
        
            wbkX.Close True
             
        
        Case vbNo
        
            MsgBox "����� ��������, ������������ ������ ��������.", vbExclamation
            Exit Sub
        
   End Select
   
   With Application
   
        .ScreenUpdating = True
        .DisplayAlerts = True
    
   End With

End Sub
