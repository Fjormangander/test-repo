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
    
    'N1 = rngAll.Columns.Count '�������� ��� ���� ����������
    'N2 = rngAll.Rows.Count
       
    shtX.Range("A3:AL1048576").Clear
    wbkX.Sheets(">>SQL").Range("A2:I1048576").Clear
      
    strFile = fd.SelectedItems(1)
      
    Set wbkY = Application.Workbooks.Open(strFile)
   
    Set rngCopy = wbkY.Sheets(1).Range("A2").CurrentRegion
    N1 = rngCopy.Rows.Count
    Set rngCopy = rngCopy.Range("A2").Resize(N1 - 1, 30)
    rngCopy.Copy
   
    Set rngPaste = shtX.Range("A3")
   
    rngPaste.PasteSpecial
    wbkY.Close False
   
    Set rngAll = shtX.Range("A1").CurrentRegion
    N2 = rngAll.Rows.Count
    shtX.Range("AE2:AL" & N2).FillDown
      
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
    
    Dim shtX As Worksheet, shtSourse As Worksheet
    Dim rngPaste As Range, rngCopy As Range, rngAll As Range, rng_b As Range
    Dim X1 As Long, Y1 As Long
    
    Dim str_list As String, str_tit As String, strFile As String
    Dim strX As String
    Dim strAdress As String
    Dim str_Crit_1 As String, str_Crit_2 As String, str_Crit_3 As String, Field_1 As String, Field_2 As String
    Dim Names_Array
    Dim g As Integer, i As Integer, Result As Integer, Calc As Integer
        
    With Application
    
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
        .DisplayAlerts = False
    
    End With
    
    'On Error Resume Next
   
    Set wbkX = ThisWorkbook
    
    Names_Array = Array("����_2", "����_3", "����_4", "����_5", "����_6", "����_7", "����_8")
    
    For i = LBound(Names_Array) To UBound(Names_Array)
                       
            Set shtX = wbkX.Sheets(Names_Array(i))
           
            With shtX.Range("A6:I300")
                    .Clear
                    .Borders(xlEdgeLeft).LineStyle = xlNone
                    .Borders(xlEdgeTop).LineStyle = xlNone
                    .Borders(xlEdgeBottom).LineStyle = xlNone
                    .Borders(xlEdgeRight).LineStyle = xlNone
                    .Borders(xlInsideVertical).LineStyle = xlNone
                    .Borders(xlInsideHorizontal).LineStyle = xlNone
                    
            End With
            
    Next i
    
    Set shtSourse = wbkX.Sheets(">>DATA")
    
    Dim wks As Worksheet
    
    For Each wks In wbkX.Worksheets
        wks.Calculate
    Next wks
    
    str_Crit_1 = wbkX.Sheets(">>SET").Range("M17").Value
        
    Names_Array = Array("����_2", "����_3", "����_4")
    
    For i = LBound(Names_Array) To UBound(Names_Array)
                       
            Set shtX = wbkX.Sheets(Names_Array(i))
            
            Select Case shtX.Name
    
                Case "����_2"
                
                str_Crit_2 = ">1" '������� � �������
                Field_1 = 8
                Field_2 = 34
                str_Crit_3 = "��������� ����� �����������"
    
                Case "����_3"
                
                str_Crit_2 = ">0"
                Field_1 = 30
                Field_2 = 35
                str_Crit_3 = "��������� ����� ��������"
                
                Case "����_4"
                str_Crit_1 = "� ������"
                str_Crit_2 = ">0"
                Field_1 = 24
                Field_2 = 35
    
            End Select
    
            shtSourse.Activate
            Set rngAll = shtSourse.Range("A1").CurrentRegion
    
            rngAll.AutoFilter Field:=Field_1, Criteria1:=str_Crit_1
            rngAll.AutoFilter Field:=Field_2, Criteria1:=str_Crit_2
    
            Calc = shtSourse.AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count
            
            Select Case Calc
                
                Case 1
                shtX.Visible = False
                
                Case Else
                
                Range("A1").Select
                Range(Selection, Selection.End(xlDown)).Copy
                Set rngPaste = shtX.Range("A6")
                rngPaste.PasteSpecial
                shtX.Rows("6:6").Delete Shift:=xlUp
                
                With shtX
    
                    .Range("B6").FormulaLocal = "=���($A6;'>>DATA'!$A:$AL;2;0)"
                    .Range("C6").FormulaLocal = "=���($A6;'>>DATA'!$A:$AL;9;0)"
                    .Range("D6").FormulaLocal = "=���($A6;'>>DATA'!$A:$AL;10;0)"
                    .Range("E6").FormulaLocal = "=���($A6;'>>DATA'!$A:$AL;25;0)"
                    .Range("F6").FormulaLocal = _
                    "=����(���($A6;'>>DATA'!$A:$AL;26;0)="""";""�� ����������"";���($A6;'>>DATA'!$A:$AL;26;0))"
                    .Range("G6").FormulaLocal = "=���($A6;'>>DATA'!$A:$AL;14;0)"
                    .Range("H6").FormulaLocal = "=���($A6;'>>DATA'!$A:$AL;33;0)"
  
                End With
                
                If shtX.Name <> "����_4" Then
                    shtX.Range("I6").Value = str_Crit_3
                    Else: shtX.Range("I6").FormulaLocal = "=���($A6;'>>DATA'!$A:$AL;27;0)"
                End If
                
                    Select Case Calc
                    
                        Case 2
                        
                        Cells.Replace What:="=", Replacement:="=", LookAt:=xlPart, SearchOrder _
                        :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
                        
                        shtX.Calculate
                        
                        With shtX.Range("A5:I6")
                        
                            .Font.Name = "Arial"
                            .Columns.AutoFit
                            .HorizontalAlignment = xlCenter
                            .VerticalAlignment = xlCenter
                            .Borders(xlEdgeLeft).LineStyle = xlContinuous
                            .Borders(xlEdgeTop).LineStyle = xlContinuous
                            .Borders(xlEdgeBottom).LineStyle = xlContinuous
                            .Borders(xlEdgeRight).LineStyle = xlContinuous
                            .Borders(xlInsideVertical).LineStyle = xlContinuous
                            .Borders(xlInsideHorizontal).LineStyle = xlContinuous
                        
                        End With
                                            
                        Case Else
                        shtX.Range("B6:I" & Calc + 4).FillDown
                        
                        Cells.Replace What:="=", Replacement:="=", LookAt:=xlPart, SearchOrder _
                        :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
                        
                        shtX.Calculate
                        
                        With shtX.Range("A5:I" & Calc + 4)
    
                            .Font.Name = "Arial"
                            .Columns.AutoFit
                            .HorizontalAlignment = xlCenter
                            .VerticalAlignment = xlCenter
                            .Borders(xlEdgeLeft).LineStyle = xlContinuous
                            .Borders(xlEdgeTop).LineStyle = xlContinuous
                            .Borders(xlEdgeBottom).LineStyle = xlContinuous
                            .Borders(xlEdgeRight).LineStyle = xlContinuous
                            .Borders(xlInsideVertical).LineStyle = xlContinuous
                            .Borders(xlInsideHorizontal).LineStyle = xlContinuous
            
                        End With
                
                    End Select
                    
            End Select
    
            shtSourse.ShowAllData

    Next i
    
    Set shtSourse = wbkX.Sheets(">>SQL")

    Set rngAll = shtSourse.Range("A1").CurrentRegion
    
    shtSourse.Range("A1").AutoFilter
    Calc = rngAll.Rows.Count - 1
    
    shtSourse.AutoFilter.Sort.SortFields.Add Key:=Range( _
        "G1:G" & Calc), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    
    With shtSourse.AutoFilter.Sort
    
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
        
    End With
    
    Names_Array = Array("����_5", "����_6", "����_7", "����_8")
    
    Field_1 = 9
    
    For i = LBound(Names_Array) To UBound(Names_Array)
                       
            Set shtX = wbkX.Sheets(Names_Array(i))
            
            Select Case shtX.Name
    
                Case "����_5"
                
                str_Crit_1 = "1 ���� �� ��������� ����� ���������"
               
                Case "����_6"
                
                str_Crit_1 = "��������� ���. �����"
                
                Case "����_7"
                
                str_Crit_1 = "��������� ������������� �� ��"
                
                Case "����_8"
                
                str_Crit_1 = "����������� �� �� � ������"
    
            End Select
            
            shtSourse.Activate
            
            Set rngAll = shtSourse.Range("A1").CurrentRegion
    
            rngAll.AutoFilter Field:=Field_1, Criteria1:=str_Crit_1
    
            Calc = shtSourse.AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count
            
            Select Case Calc
                
                Case 1
                shtX.Visible = False
                
                Case Else
                
                Range("A1").Select
                Range(Selection, Selection.End(xlDown)).Select
                Range(Selection, Selection.End(xlToRight)).Copy
                'Range(Selection, Selection.End(xlDown)).Copy
                Set rngPaste = shtX.Range("A6")
                rngPaste.PasteSpecial
                shtX.Rows("6:6").Delete Shift:=xlUp
                        
                With shtX.Range("A5:I" & Calc + 4)
    
                    .Font.Name = "Arial"
                    .Columns.AutoFit
                    .HorizontalAlignment = xlCenter
                    .VerticalAlignment = xlCenter
                    .Borders(xlEdgeLeft).LineStyle = xlContinuous
                    .Borders(xlEdgeTop).LineStyle = xlContinuous
                    .Borders(xlEdgeBottom).LineStyle = xlContinuous
                    .Borders(xlEdgeRight).LineStyle = xlContinuous
                    .Borders(xlInsideVertical).LineStyle = xlContinuous
                    .Borders(xlInsideHorizontal).LineStyle = xlContinuous
            
                End With
                
                    End Select
    
            shtSourse.ShowAllData

    Next i
    
    Dim xSht As Variant

    For Each xSht In ActiveWorkbook.Sheets
        If xSht.Visible Then g = g + 1
    Next
    
    g = g - 4
    
    str_Crit_1 = wbkX.Sheets(">>SET").Range("M17").Value
    
    strX = "!" & str_Crit_1 & "_"
        
    strAdress = _
"\\hq.rgs.ru\users\RGS\����������� �������������� � ������������\05_����������� � ������\����������\������ 2019\��\�������\" & strX & "������ ������ ����� (�� ��).pdf"
        
    wbkX.ExportAsFixedFormat Type:=xlTypePDF, Filename:=strAdress _
        , Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
        :=False, From:=1, To:=g, OpenAfterPublish:=True 'False True
    Application.WindowState = xlNormal
    
    
    With Application
    
        .ScreenUpdating = True
        .DisplayAlerts = True
    
    End With
       
    MsgBox "�������� ���������!", vbExclamation
   
            
End Sub
    
   Sub Send_email()
   
   Dim Question As Integer
   Dim wbkX As Workbook
   Dim shtX As Worksheet
   Dim strX As String
   
   Dim objOutlookApp As Object, objMail As Object
   Dim sTo As String, sCC As String, sSubject As String, sBody As String, sAttachment As String, strAdress As String
   
   
   Set wbkX = ThisWorkbook
   Set shtX = wbkX.Sheets(">>SET")
   strX = wbkX.Sheets(">>SET").Range("M17").Value
      
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
   
            sTo = shtX.Range("V3").Value
            sCC = shtX.Range("V4").Value
            sSubject = "������ ������ ����� �� " & strX
            
            sBody = "<p>������ ����, �������!</p>" & _
                    "��������� ����� ""������ ������ �����"" �� ��������� �� ������� (�� ���������� ������� ����).<br>"
              
            strAdress = _
            "\\hq.rgs.ru\users\RGS\����������� �������������� � ������������\05_����������� � ������\����������\������ 2019\��\�������\" & "!" & strX & "_" & "������ ������ ����� (�� ��).pdf"
            
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
        
            wbkX.Close False
             
        
        Case vbNo
        
            MsgBox "����� ��������, ������������ ������ ��������.", vbExclamation
            Exit Sub
        
   End Select
   
   With Application
   
        .ScreenUpdating = True
        .DisplayAlerts = True
    
   End With

End Sub

