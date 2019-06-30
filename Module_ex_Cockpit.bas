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
'Среднее за 2017 - I квадрант

Set rngX = shtX.Range("C8")

rngX.FormulaR1C1 = "=AVERAGE(R27C2:R" & N2 & "C2)"

Set rngX = shtX.Range("C9")

rngX.FormulaR1C1 = "=AVERAGE(R27C3:R" & N2 & "C3)"

Set rngX = shtX.Range("C10")

rngX.FormulaR1C1 = "=AVERAGE(R27C6:R" & N2 & "C6)"

'----------------------------------------------------------
'Среднее за восемь недель - I квадрант

Set rngX = shtX.Range("D8")

rngX.FormulaR1C1 = "=AVERAGE(R" & N1 & "C2:R" & N2 & "C2)"

Set rngX = shtX.Range("D9")

rngX.FormulaR1C1 = "=AVERAGE(R" & N1 & "C3:R" & N2 & "C3)"

Set rngX = shtX.Range("D10")

rngX.FormulaR1C1 = "=AVERAGE(R" & N1 & "C6:R" & N2 & "C6)"

'----------------------------------------------------------
'Среднее за 2017 - II квадрант

Set rngX = shtX.Range("C22")

rngX.FormulaR1C1 = "=AVERAGE(R27C4:R" & N2 & "C4)"

Set rngX = shtX.Range("C23")

rngX.FormulaR1C1 = "=AVERAGE(R42C5:R" & N2 & "C5)"

'----------------------------------------------------------
'Среднее за восемь недель - II квадрант

Set rngX = shtX.Range("D22")

rngX.FormulaR1C1 = "=AVERAGE(R" & N1 & "C4:R" & N2 & "C4)"

Set rngX = shtX.Range("D23")

rngX.FormulaR1C1 = "=AVERAGE(R" & N1 & "C5:R" & N2 & "C5)"


'----------------------------------------------------------
'Среднее за 2017 - III квадрант

Set rngX = shtX.Range("C88")

rngX.FormulaR1C1 = "=AVERAGE(R27C7:R" & N2 & "C7)"

Set rngX = shtX.Range("C89")

rngX.FormulaR1C1 = "=AVERAGE(R27C9:R" & N2 & "C9)"

Set rngX = shtX.Range("C91")

rngX.FormulaR1C1 = "=AVERAGE(R27C10:R" & N2 & "C10)"

'----------------------------------------------------------
'Среднее за восемь недель - III квадрант

Set rngX = shtX.Range("D88")

rngX.FormulaR1C1 = "=AVERAGE(R" & N1 & "C7:R" & N2 & "C7)"

Set rngX = shtX.Range("D89")

rngX.FormulaR1C1 = "=AVERAGE(R" & N1 & "C9:R" & N2 & "C9)"

Set rngX = shtX.Range("D91")

rngX.FormulaR1C1 = "=AVERAGE(R" & N1 & "C10:R" & N2 & "C10)"

'----------------------------------------------------------
'Среднее за 2017 - IV квадрант

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
'Среднее за восемь недель - IV квадрант

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



MsgBox "Формулы обновлены!", vbExclamation


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
        
        MsgBox "Приложение Word уже открыто!", vbInformation, "Check_OpenWord"
    End If
    
    
     Set objWrdDoc = objWrdApp.Documents.Open("C:\Users\VVBelov\Desktop\АВТО\Отчеты и алгоритмы\Справка OSAGO SALES Cockpit.docx")
    
    wbkX.Worksheets("СПРАВКА").Calculate
    
    'wbkX.Worksheets("СПРАВКА").Range("B3").Copy
    
    objWrdDoc.Bookmarks("Неделя").Range.Text = wbkX.Worksheets("СПРАВКА").Range("B2")
    objWrdDoc.Bookmarks("Сегодня").Range.Text = wbkX.Worksheets("СПРАВКА").Range("C2")
    
    objWrdDoc.Bookmarks("Поступило_обращений").Range.Text = wbkX.Worksheets("СПРАВКА").Range("B3")
    
    'wbkX.Worksheets("СПРАВКА").Range("B4").Copy
    
    objWrdDoc.Bookmarks("Поступило_ЦБ").Range.Text = wbkX.Worksheets("СПРАВКА").Range("B4")
    
    'wbkX.Worksheets("СПРАВКА").Range("B5").Copy
    
    objWrdDoc.Bookmarks("Поступило_РСА").Range.Text = wbkX.Worksheets("СПРАВКА").Range("B5")
    
    
    objWrdDoc.Bookmarks("Динамика_обращений").Range.Text = wbkX.Worksheets("СПРАВКА").Range("B7")
    
    objWrdDoc.Bookmarks("Динамика_шт_обр").Range.Text = wbkX.Worksheets("СПРАВКА").Range("B6")
    
    objWrdDoc.Bookmarks("В_среднем_обр").Range.Text = wbkX.Worksheets("СПРАВКА").Range("B8")
    
    '
    objWrdDoc.Bookmarks("КБМ_шт").Range.Text = wbkX.Worksheets("СПРАВКА").Range("B9")
    
    objWrdDoc.Bookmarks("КБМ_пр").Range.Text = wbkX.Worksheets("СПРАВКА").Range("C9")
    
    objWrdDoc.Bookmarks("КБМ_динамика").Range.Text = wbkX.Worksheets("СПРАВКА").Range("E9")
    
    objWrdDoc.Bookmarks("КБМ_изм").Range.Text = wbkX.Worksheets("СПРАВКА").Range("D9")
    
    '
    objWrdDoc.Bookmarks("еОСАГО_шт").Range.Text = wbkX.Worksheets("СПРАВКА").Range("B10")
    
    objWrdDoc.Bookmarks("еОСАГО_пр").Range.Text = wbkX.Worksheets("СПРАВКА").Range("C10")
    
    objWrdDoc.Bookmarks("еОСАГО_динамика").Range.Text = wbkX.Worksheets("СПРАВКА").Range("E10")
    
    objWrdDoc.Bookmarks("еОСАГО_изм").Range.Text = wbkX.Worksheets("СПРАВКА").Range("D10")
    
    objWrdDoc.Bookmarks("еОСАГО_8нед").Range.Text = wbkX.Worksheets("СПРАВКА").Range("C11")
    
    objWrdDoc.Bookmarks("еОСАГО_пр_пр0").Range.Text = wbkX.Worksheets("СПРАВКА").Range("B11")
    
    objWrdDoc.Bookmarks("еОСАГО_СК2").Range.Text = wbkX.Worksheets("СПРАВКА").Range("D11")
    
    '
    objWrdDoc.Bookmarks("ОК_шт").Range.Text = wbkX.Worksheets("СПРАВКА").Range("B12")
    
    objWrdDoc.Bookmarks("ОК_пр").Range.Text = wbkX.Worksheets("СПРАВКА").Range("C12")
    
    objWrdDoc.Bookmarks("ОК_динамика").Range.Text = wbkX.Worksheets("СПРАВКА").Range("E12")
    
    objWrdDoc.Bookmarks("ОК_изм").Range.Text = wbkX.Worksheets("СПРАВКА").Range("D12")
    
    '
    
    objWrdDoc.Bookmarks("Иные_шт").Range.Text = wbkX.Worksheets("СПРАВКА").Range("B13")
    
    objWrdDoc.Bookmarks("Иные_пр").Range.Text = wbkX.Worksheets("СПРАВКА").Range("C13")
    
    objWrdDoc.Bookmarks("Иные_динамика").Range.Text = wbkX.Worksheets("СПРАВКА").Range("E13")
    
    objWrdDoc.Bookmarks("Иные_изм").Range.Text = wbkX.Worksheets("СПРАВКА").Range("D13")
    
    '
    
    objWrdDoc.Bookmarks("пр_пр_1").Range.Text = wbkX.Worksheets("СПРАВКА").Range("J5")
    
    objWrdDoc.Bookmarks("пр_пр_2").Range.Text = wbkX.Worksheets("СПРАВКА").Range("J6")
    
    objWrdDoc.Bookmarks("Выполнено").Range.Text = wbkX.Worksheets("СПРАВКА").Range("B15")
    
    objWrdDoc.Bookmarks("В_работе").Range.Text = wbkX.Worksheets("СПРАВКА").Range("B16")
      
    objWrdDoc.Close True
    
    objWrdApp.Quit
   
    Set objWrdDoc = Nothing: Set objWrdApp = Nothing
    
    
    

End Sub
