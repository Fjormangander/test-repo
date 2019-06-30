Attribute VB_Name = "Module1"
Option Explicit

Sub SaveAs()

    Dim wbkX As Workbook
    Dim strX1 As String, strX2 As String, strAdress As String
    Dim rngX1 As Range, rngX2 As Range
    Dim Names_Array
    Dim g As Integer
    Dim objApp As Object
        
    With Application
    
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
        .DisplayAlerts = False
    
    End With
    
    On Error Resume Next

    Set wbkX = ThisWorkbook
    Set rngX1 = wbkX.Worksheets("настройки").Range("J9")
    Set rngX2 = wbkX.Worksheets("настройки").Range("J10")
    
    wbkX.Worksheets("настройки").Calculate
        
    strX1 = rngX1.Value
    strX2 = rngX2.Value
        
    strAdress = "\\hq.rgs.ru\users\RGS\Департамент взаимодействия с регуляторами\05_Методология и анализ\Отчетность\Отчеты 2019\ЦБ\Cockpit\" & strX2 & "_Отчет Cockpit ОСАГО 8 недель_" & strX1 & ".xlsx"
         
    ChDir "\\hq.rgs.ru\users\RGS\Департамент взаимодействия с регуляторами\05_Методология и анализ\Отчетность\Отчеты 2019\ЦБ\Cockpit"
      
    wbkX.SaveAs Filename:=strAdress, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    
    
    Names_Array = Array("Cockpit", "Cockpit (ОСАГО)", "C-1")
           
    For g = LBound(Names_Array) To UBound(Names_Array)
    
        wbkX.Sheets(Names_Array(g)).Activate
        Cells.Select
        Selection.Copy
        Range("A1").Select
        Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=False, Transpose:=False
        Range("A2").Select
    
    Next g
    
    Names_Array = Array("opt", "еОСАГО", "УУ", "Справка", "настройки", "Отчет БЦБ", ">>DATA", ">>SET", "Cockpit (ОСАГО)")
    
    For g = LBound(Names_Array) To UBound(Names_Array)
    
        wbkX.Sheets(Names_Array(g)).Delete
    
    Next g
    
    Names_Array = Array("C-1")
           
    For g = LBound(Names_Array) To UBound(Names_Array)
    
        wbkX.Sheets(Names_Array(g)).Visible = False
        
    Next g
        
    wbkX.Worksheets("Cockpit").Activate
        
    MsgBox "Копия отчета сохранена!", vbExclamation

    wbkX.Close (True)
       
    Set objApp = GetObject(, "Excel.Application")
    objApp.Quit
    
    With Application
    
        .ScreenUpdating = True
        .DisplayAlerts = True
    
    End With
       
End Sub

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
        .Filters.Add "Файлы Excel", "*.xls;*.xlsx"
        '.InitialFileName = "C:\Users..."
        .Show
    End With
   
    If fd.SelectedItems.Count = 0 Then
        Exit Sub
    End If
   
    Set rngAll = shtX.Range("A1").CurrentRegion
       
    shtX.Range("A3:M1048576").Clear
      
    strFile = fd.SelectedItems(1)
      
    Set wbkY = Application.Workbooks.Open(strFile)
   
    Set rngCopy = wbkY.Sheets(1).Range("A2").CurrentRegion
    N1 = rngCopy.Rows.Count
    Set rngCopy = rngCopy.Range("A2").Resize(N1 - 1, 15)
    rngCopy.Copy
   
    Set rngPaste = shtX.Range("A3")
   
    rngPaste.PasteSpecial
    wbkY.Close False
   
    Set rngAll = shtX.Range("A1").CurrentRegion
    N2 = rngAll.Rows.Count
    shtX.Range("P2:P" & N2).FillDown
      
    shtX.Rows("2:2").Delete Shift:=xlUp
    
    shtX.Calculate
    
    Dim wks As Worksheet
    
    For Each wks In wbkX.Worksheets
        wks.Calculate
    Next wks
    
    With wbkX
        
        .Sheets("Cockpit").Activate
        .Save
    
    End With
    
    With Application
    
        .ScreenUpdating = True
        .DisplayAlerts = True
    
    End With
   
    MsgBox "Обработка данных завершена, проверьте ошибки!", vbExclamation

End Sub
