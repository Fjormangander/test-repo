Attribute VB_Name = "Module1"
Option Explicit

Sub FileList()
    Dim V As String
    Dim BrowseFolder As String
     
    'открываем диалоговое окно выбора папки
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Выберите папку или диск"
        .Show
        On Error Resume Next
        Err.Clear
        V = .SelectedItems(1)
        If Err.Number <> 0 Then
            MsgBox "Вы ничего не выбрали!"
            Exit Sub
        End If
    End With
    BrowseFolder = CStr(V)
   
    Sheets("Данные ЛК").Select
    'Worksheets("Данные_ЛК").Range("A3:E" & Range("A65536").End(xlUp).Row).ClearContents
    Worksheets("Данные_ЛК").Range("A3:L65536").ClearContents

        
    'вызываем процедуру вывода списка файлов
    'измените True на False, если не нужно выводить файлы из вложенных папок
    ListFilesInFolder BrowseFolder, True
End Sub
 
 
Private Sub ListFilesInFolder(ByVal SourceFolderName As String, ByVal IncludeSubfolders As Boolean)
 
    Dim FSO As Object
    Dim SourceFolder As Object
    Dim SubFolder As Object
    Dim FileItem As Object
    Dim r As Long
    Dim X As String
 
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set SourceFolder = FSO.getfolder(SourceFolderName)
 
    r = Range("A65536").End(xlUp).Row + 1   'находим первую пустую строку
    'выводим данные по файлу
    For Each FileItem In SourceFolder.Files
        Cells(r, 1).Formula = FileItem.Name
        Cells(r, 2).Formula = FileItem.Path
        Cells(r, 3).Formula = FileItem.Size
        Cells(r, 4).Formula = FileItem.DateCreated
        Cells(r, 5).Formula = FileItem.DateLastModified
     
        r = r + 1
        X = SourceFolder.Path
    Next FileItem
     
    'вызываем процедуру повторно для каждой вложенной папки
    If IncludeSubfolders Then
        For Each SubFolder In SourceFolder.SubFolders
            ListFilesInFolder SubFolder.Path, True
        Next SubFolder
    End If
 
    Columns("A:L").AutoFit
 
    Set FileItem = Nothing
    Set SourceFolder = Nothing
    Set FSO = Nothing
 
End Sub

Sub GetData()

    
    Dim wbkX As Workbook, wbkY As Workbook
    Dim sht_lk As Worksheet, sht_nsd As Worksheet
    Dim fd As FileDialog
    Dim strFile As String, strX As String, strAdress As String
    Dim rngPaste As Range, rngCopy As Range, rngFormula As Range, rngDate As Range
    Dim N1 As Long, N2 As Long
 
    With Application
    
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
        .DisplayAlerts = False
    
    End With
    
    Set wbkY = ThisWorkbook
    Set sht_lk = wbkY.Sheets("Данные_ЛК")
    Set sht_nsd = wbkY.Sheets("Данные_NSD")
    
    MsgBox "Сейчас потребуется указать путь к файлу с данными из Naumen SD. Нажмите 'ОК', чтобы продолжить.", vbExclamation
      
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
    
    wbkY.Sheets("Данные_NSD").Range("A1:L1048576").Clear
      
    Set wbkX = Application.Workbooks.Open(strFile)
   
    Set rngCopy = wbkX.Sheets(1).Range("A2").CurrentRegion
    N1 = rngCopy.Rows.Count
    Set rngCopy = rngCopy.Range("A2").Resize(N1 - 1, 13)
    rngCopy.Copy
   
    Set rngPaste = wbkY.Sheets("Данные_NSD").Range("A1")
   
    rngPaste.PasteSpecial
    wbkX.Close False
    
    MsgBox "Сейчас потребуется указать путь к файлу с данными из ЛК. Нажмите 'ОК', чтобы продолжить.", vbExclamation
    
    sht_lk.Activate
    
    Call FileList
    
    Set rngFormula = sht_lk.Range("A1").CurrentRegion
    N1 = rngFormula.Rows.Count
    N2 = rngFormula.Columns.Count
    
    sht_lk.Range("F2:L" & N1).FillDown
    sht_lk.Calculate
    
    wbkY.Sheets("ИТОГ").Activate
    wbkY.Sheets("ИТОГ").Calculate
    
    wbkY.Sheets("ИТОГ").PivotTables(1).SourceData = "Данные_ЛК!R1C1:R" & N1 & "C" & N2
    wbkY.Sheets("ИТОГ").PivotTables(1).PivotCache.Refresh
    
    '-------------------------------
    Set rngDate = sht_lk.Range("S1")
    
    strX = rngDate.Value
    
    strAdress = "\\hq.rgs.ru\users\RGS\Департамент взаимодействия с регуляторами\Методология\Отчетность\Отчеты 2018\ЦБ\Контроль ЛК\Контроль поступлений из Личного Кабинета за_" & strX & ".xlsx"
    
    ChDir "\\hq.rgs.ru\users\RGS\Департамент взаимодействия с регуляторами\Методология\Отчетность\Отчеты 2018\ЦБ\Контроль ЛК"
   
    wbkY.SaveAs Filename:=strAdress, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
  
  
   
    
    '-------------------------------
    
      
    With Application
    
        .ScreenUpdating = True
        .DisplayAlerts = True
    
    End With
   
    MsgBox "Обработка данных завершена, проверьте ошибки!", vbExclamation

End Sub

