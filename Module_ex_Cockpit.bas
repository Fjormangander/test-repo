Attribute VB_Name = "Module1"
Option Explicit

Sub Formula()

Dim wbkX As Workbook

Dim shtX As Worksheet

Dim rngX As Range, rngX1 As Range, rngX2 As Range
Dim N1 As Long, N2 As Long


Set wbkX = ThisWorkbook

Set shtX = wbkX.Worksheets("C-1")

Set rngX1 = shtX.Range("D11")
Set rngX2 = shtX.Range("E11")

shtX.Calculate

N1 = rngX1.Value

N2 = rngX2.Value

'----------------------------------------------------------
'������� �� 2017 - I ��������

Set rngX = shtX.Range("C8")

rngX.FormulaR1C1 = "=AVERAGE(R27C2:R" & N2 & "C2)"

Set rngX = shtX.Range("C9")

rngX.FormulaR1C1 = "=AVERAGE(R27C3:R" & N2 & "C3)"

Set rngX = shtX.Range("C10")

rngX.FormulaR1C1 = "=AVERAGE(R27C6:R" & N2 & "C6)"

'----------------------------------------------------------
'������� �� ������ ������ - I ��������

Set rngX = shtX.Range("D8")

rngX.FormulaR1C1 = "=AVERAGE(R" & N1 & "C2:R" & N2 & "C2)"

Set rngX = shtX.Range("D9")

rngX.FormulaR1C1 = "=AVERAGE(R" & N1 & "C3:R" & N2 & "C3)"

Set rngX = shtX.Range("D10")

rngX.FormulaR1C1 = "=AVERAGE(R" & N1 & "C6:R" & N2 & "C6)"

'----------------------------------------------------------
'������� �� 2017 - II ��������

Set rngX = shtX.Range("C22")

rngX.FormulaR1C1 = "=AVERAGE(R27C4:R" & N2 & "C4)"

Set rngX = shtX.Range("C23")

rngX.FormulaR1C1 = "=AVERAGE(R42C5:R" & N2 & "C5)"

'----------------------------------------------------------
'������� �� ������ ������ - II ��������

Set rngX = shtX.Range("D22")

rngX.FormulaR1C1 = "=AVERAGE(R" & N1 & "C4:R" & N2 & "C4)"

Set rngX = shtX.Range("D23")

rngX.FormulaR1C1 = "=AVERAGE(R" & N1 & "C5:R" & N2 & "C5)"


'----------------------------------------------------------
'������� �� 2017 - III ��������

Set rngX = shtX.Range("C88")

rngX.FormulaR1C1 = "=AVERAGE(R27C7:R" & N2 & "C7)"

Set rngX = shtX.Range("C89")

rngX.FormulaR1C1 = "=AVERAGE(R27C9:R" & N2 & "C9)"

Set rngX = shtX.Range("C91")

rngX.FormulaR1C1 = "=AVERAGE(R27C10:R" & N2 & "C10)"

'----------------------------------------------------------
'������� �� ������ ������ - III ��������

Set rngX = shtX.Range("D88")

rngX.FormulaR1C1 = "=AVERAGE(R" & N1 & "C7:R" & N2 & "C7)"

Set rngX = shtX.Range("D89")

rngX.FormulaR1C1 = "=AVERAGE(R" & N1 & "C9:R" & N2 & "C9)"

Set rngX = shtX.Range("D91")

rngX.FormulaR1C1 = "=AVERAGE(R" & N1 & "C10:R" & N2 & "C10)"

'----------------------------------------------------------
'������� �� 2017 - IV ��������

Set rngX1 = shtX.Range("D92")

N1 = rngX1.Value

Set rngX2 = shtX.Range("E92")

N2 = rngX2.Value

'----------------------------------------------------------

Set rngX = shtX.Range("C102")

rngX.FormulaR1C1 = "=AVERAGE(R109C2:R" & N2 & "C2)"

Set rngX = shtX.Range("C103")

rngX.FormulaR1C1 = "=AVERAGE(R109C3:R" & N2 & "C3)"

Set rngX = shtX.Range("C104")

rngX.FormulaR1C1 = "=AVERAGE(R109C4:R" & N2 & "C4)"

Set rngX = shtX.Range("C105")

rngX.FormulaR1C1 = "=AVERAGE(R109C5:R" & N2 & "C5)"

Set rngX = shtX.Range("C106")

rngX.FormulaR1C1 = "=AVERAGE(R109C6:R" & N2 & "C6)"

'----------------------------------------------------------
'������� �� ������ ������ - IV ��������

Set rngX = shtX.Range("D102")

rngX.FormulaR1C1 = "=AVERAGE(R" & N1 & "C2:R" & N2 & "C2)"

Set rngX = shtX.Range("D103")

rngX.FormulaR1C1 = "=AVERAGE(R" & N1 & "C3:R" & N2 & "C3)"

Set rngX = shtX.Range("D104")

rngX.FormulaR1C1 = "=AVERAGE(R" & N1 & "C4:R" & N2 & "C4)"

Set rngX = shtX.Range("D105")

rngX.FormulaR1C1 = "=AVERAGE(R" & N1 & "C5:R" & N2 & "C5)"

Set rngX = shtX.Range("D106")

rngX.FormulaR1C1 = "=AVERAGE(R" & N1 & "C5:R" & N2 & "C6)"


shtX.Calculate
wbkX.Worksheets("Cockpit").Activate
wbkX.Worksheets("Cockpit").Calculate



MsgBox "������� ���������!", vbExclamation


End Sub

Sub toWord()

    Dim wbkX As Workbook
    Dim strX As String, strAdress As String
    Dim rngX As Range
    Dim N As Long
        
    With Application
    
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
        .DisplayAlerts = False
    
    End With
    
    Set wbkX = ThisWorkbook
   
    
   'On Error Resume Next
    
    
    
    
    Dim objWrdApp As Object, objWrdDoc As Object
           
    If objWrdApp Is Nothing Then
        
        Set objWrdApp = CreateObject("Word.Application")
        
        objWrdApp.Visible = True
    Else
        
        MsgBox "���������� Word ��� �������!", vbInformation, "Check_OpenWord"
    End If
    
    
     Set objWrdDoc = objWrdApp.Documents.Open("C:\Users\VVBelov\Desktop\����\������ � ���������\������� OSAGO SALES Cockpit.docx")
    
    wbkX.Worksheets("�������").Calculate
    
    'wbkX.Worksheets("�������").Range("B3").Copy
    
    objWrdDoc.Bookmarks("������").Range.Text = wbkX.Worksheets("�������").Range("B2")
    objWrdDoc.Bookmarks("�������").Range.Text = wbkX.Worksheets("�������").Range("C2")
    
    objWrdDoc.Bookmarks("���������_���������").Range.Text = wbkX.Worksheets("�������").Range("B3")
    
    'wbkX.Worksheets("�������").Range("B4").Copy
    
    objWrdDoc.Bookmarks("���������_��").Range.Text = wbkX.Worksheets("�������").Range("B4")
    
    'wbkX.Worksheets("�������").Range("B5").Copy
    
    objWrdDoc.Bookmarks("���������_���").Range.Text = wbkX.Worksheets("�������").Range("B5")
    
    
    objWrdDoc.Bookmarks("��������_���������").Range.Text = wbkX.Worksheets("�������").Range("B7")
    
    objWrdDoc.Bookmarks("��������_��_���").Range.Text = wbkX.Worksheets("�������").Range("B6")
    
    objWrdDoc.Bookmarks("�_�������_���").Range.Text = wbkX.Worksheets("�������").Range("B8")
    
    '
    objWrdDoc.Bookmarks("���_��").Range.Text = wbkX.Worksheets("�������").Range("B9")
    
    objWrdDoc.Bookmarks("���_��").Range.Text = wbkX.Worksheets("�������").Range("C9")
    
    objWrdDoc.Bookmarks("���_��������").Range.Text = wbkX.Worksheets("�������").Range("E9")
    
    objWrdDoc.Bookmarks("���_���").Range.Text = wbkX.Worksheets("�������").Range("D9")
    
    '
    objWrdDoc.Bookmarks("������_��").Range.Text = wbkX.Worksheets("�������").Range("B10")
    
    objWrdDoc.Bookmarks("������_��").Range.Text = wbkX.Worksheets("�������").Range("C10")
    
    objWrdDoc.Bookmarks("������_��������").Range.Text = wbkX.Worksheets("�������").Range("E10")
    
    objWrdDoc.Bookmarks("������_���").Range.Text = wbkX.Worksheets("�������").Range("D10")
    
    objWrdDoc.Bookmarks("������_8���").Range.Text = wbkX.Worksheets("�������").Range("C11")
    
    objWrdDoc.Bookmarks("������_��_��0").Range.Text = wbkX.Worksheets("�������").Range("B11")
    
    objWrdDoc.Bookmarks("������_��2").Range.Text = wbkX.Worksheets("�������").Range("D11")
    
    '
    objWrdDoc.Bookmarks("��_��").Range.Text = wbkX.Worksheets("�������").Range("B12")
    
    objWrdDoc.Bookmarks("��_��").Range.Text = wbkX.Worksheets("�������").Range("C12")
    
    objWrdDoc.Bookmarks("��_��������").Range.Text = wbkX.Worksheets("�������").Range("E12")
    
    objWrdDoc.Bookmarks("��_���").Range.Text = wbkX.Worksheets("�������").Range("D12")
    
    '
    
    objWrdDoc.Bookmarks("����_��").Range.Text = wbkX.Worksheets("�������").Range("B13")
    
    objWrdDoc.Bookmarks("����_��").Range.Text = wbkX.Worksheets("�������").Range("C13")
    
    objWrdDoc.Bookmarks("����_��������").Range.Text = wbkX.Worksheets("�������").Range("E13")
    
    objWrdDoc.Bookmarks("����_���").Range.Text = wbkX.Worksheets("�������").Range("D13")
    
    '
    
    objWrdDoc.Bookmarks("��_��_1").Range.Text = wbkX.Worksheets("�������").Range("J5")
    
    objWrdDoc.Bookmarks("��_��_2").Range.Text = wbkX.Worksheets("�������").Range("J6")
    
    objWrdDoc.Bookmarks("���������").Range.Text = wbkX.Worksheets("�������").Range("B15")
    
    objWrdDoc.Bookmarks("�_������").Range.Text = wbkX.Worksheets("�������").Range("B16")
      
    objWrdDoc.Close True
    
    objWrdApp.Quit
   
    Set objWrdDoc = Nothing: Set objWrdApp = Nothing
    
    
    

End Sub
