Attribute VB_Name = "Module1"
Option Explicit

Sub GetStart()

'��������� �� �������� �����

    Dim strFile As String
    Dim wbkX As Workbook, wbkY As Workbook
    Dim shtX As Worksheet, shtY As Worksheet, shtZ As Worksheet
    Dim rngPaste As Range, rngCopy As Range, rngAll As Range
    Dim N1 As Long, N2 As Long
    Dim fd As FileDialog
    Dim strX As String, SigString As String, Signature As String
    Dim Question As Integer
    
    Dim Names_Array
    Dim i As Integer
    
    '��������� �� �������� ������
    
    Dim objOutlookApp As Object, objMail As Object
    Dim sTo As String, sCC As String, sSubject As String, sBody As String, sAttachment As String, strAdress As String
        
    With Application
    
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
        .DisplayAlerts = False
    
    End With
    
    Set wbkX = ThisWorkbook
    
    Names_Array = Array(">>SET", "����", "���������� �� ������������", "� ������_������ ���������", _
    "� ������_�������")
        
    Set shtX = wbkX.Sheets(">>DATA")
            
    If shtX.FilterMode Then
        shtX.ShowAllData
    End If
    
    Set rngAll = shtX.Range("A1").CurrentRegion
       
         
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
   
   shtX.Range("A3:T1048576").Clear
   
   strFile = fd.SelectedItems(1)
      
   Set wbkY = Application.Workbooks.Open(strFile)
   
   Set rngCopy = wbkY.Sheets(1).Range("A2").CurrentRegion
   N1 = rngCopy.Rows.Count
   N2 = rngCopy.Columns.Count
   Set rngCopy = rngCopy.Range("A2").Resize(N1 - 1, 16)
   rngCopy.Copy
      
   Set rngPaste = shtX.Range("A3")
   
   rngPaste.PasteSpecial
   wbkY.Close False
   
   Set rngAll = shtX.Range("A1").CurrentRegion
   N2 = rngAll.Rows.Count
   shtX.Range("Q2:T" & N2).FillDown
   shtX.Calculate
   
   shtX.Rows("2:2").Delete Shift:=xlUp
   
   For i = LBound(Names_Array) To UBound(Names_Array)
   
        Set shtX = wbkX.Sheets(Names_Array(i))
        shtX.Calculate
   Next i
     
   With wbkX.Sheets("����")
   
        .Activate
        .Calculate
        .Range("A2").Select
        
   End With
   
   wbkX.Save

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
   
            sTo = wbkX.Sheets(">>SET").Range("F22").Value
            sCC = wbkX.Sheets(">>SET").Range("F23").Value
            sSubject = wbkX.Sheets("����").Range("B1").Value
            'sBody = "������ ����, �������!" & vbCrLf & "��������� ���������� ������ �� ������� � ������ �� ��������� �� �������."
            sBody = "<p>������ ����, �������!</p>" & _
                    "��������� ���������� ������ �� ������� � ������ �� ��������� �� �������.<br>"
            
            Names_Array = Array(">>DATA", ">>SET", "����", "���������� �� ������������", "� ������_������ ���������", _
            "� ������_�������")
   
            For i = LBound(Names_Array) To UBound(Names_Array)
            
                    wbkX.Sheets(Names_Array(i)).Activate
                    Cells.Select
                    Selection.Copy
                    Range("A1").Select
                    Selection.PasteSpecial Paste:=xlPasteAllUsingSourceTheme, Operation:=xlNone _
                    , SkipBlanks:=False, Transpose:=False
                    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                    :=False, Transpose:=False
                    Range("A2").Select
            
            Next i
       
            wbkX.Sheets("����").Activate
                       
            Set rngAll = wbkX.Sheets(">>SET").Range("F25")
            
            strX = rngAll.Value
   
            strAdress = "\\hq.rgs.ru\users\RGS\����������� �������������� � ������������\05_����������� � ������\����������\������ 2019\������\����������� �������\" & strX & "����������� ������ �� �������.xlsx"
   
            ChDir "\\hq.rgs.ru\users\RGS\����������� �������������� � ������������\05_����������� � ������\����������\������ 2019\������\����������� �������"
             
            wbkX.SaveAs Filename:=strAdress, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
                  
            sAttachment = strAdress
                                               
            With objMail
    
                .To = sTo
                .CC = sCC
                '.BCC = "vladislav_belov@rgs.ru"
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
