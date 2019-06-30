Attribute VB_Name = "Module1"
Option Explicit




Public Function FuzzyVLOOKUP(Искомое_Значение As Range, Словарь As Range, Значние_или_индекс As Boolean, Optional Процент_совпадения As Long = 50, Optional Минимальная_длина_слова As Long = 0, Optional Слова_исключения As String = "")

    Dim x As Long
    Dim y As Double
    Dim Score As Double
    Dim min As Double
    Dim Max As Long
    Dim max2 As Long
    Dim d As Long
    For x = 1 To Словарь.Count
        Score = SringCompare(Искомое_Значение.Value, Словарь(x).Value, 50, Минимальная_длина_слова, Слова_исключения, Max)
        Score = Score + SringCompare(Словарь(x).Value, Искомое_Значение.Value, 50, Минимальная_длина_слова, Слова_исключения, max2)
        'If min > Score Then min = Score
        'If Max < Score Then Max = Score
        If Score > y Then
            y = Score

            If Значние_или_индекс = False Then
                d = max2
'Debug.Print Score / ((Max + d) * 100)
                If Score / ((Max + d) * 100) >= (Процент_совпадения / 100) Then

                    FuzzyVLOOKUP = Словарь(x).Value

                Else: FuzzyVLOOKUP = "MISS"
                End If
            End If
            If Значние_или_индекс = True Then
                d = max2
                'Debug.Print rngWith(x).Value
                FuzzyVLOOKUP = Score / ((Max + d) * 100)
            End If
        End If

    Next x

End Function






Function SringCompare(strFrom As String, strTo As String, SamePercent As Long, Optional exclLen As Long = 0, Optional ExclMask As String = "", Optional ByRef maximum As Long) As Double

    Dim ArrWordFrom() As String
    Dim ArrWordTo() As String
    Dim ArrExclMask() As String
    Dim x As Long
    Dim y As Long
    Dim z As Long
    Dim k As Long
    Dim wordCounter As Long
    Dim wordCounterComapared As Long
    Dim tmpPosition As Long
    Dim flag As Boolean
    Dim result1 As Double
    Dim result2 As Double
    Dim avResult As Double
    Dim resultIndex As Double

    ' Удаляем знаки препинания
    strFrom = UCase(replacepuctuation(strFrom))
    strTo = UCase(replacepuctuation(strTo))

    'maximum = wordscounter(strFrom, exclLen, ExclMask)


    'Строки в маасив слов
    ArrWordFrom() = Split(strFrom, " ", -1, vbTextCompare)
    ArrWordTo() = Split(strTo, " ", -1, vbTextCompare)
    'Если маска не пустая создать массив исключений
    If ExclMask <> "" Then ArrExclMask() = Split(ExclMask, ";", -1, vbTextCompare)
    ' перебор первых слов
    For x = 0 To UBound(ArrWordFrom())
        If ArrWordFrom(x) = "" Then GoTo nextwoord1:
        'Проверка на соответствие маске исключений
        If ExclMask <> "" Then
            For z = 0 To UBound(ArrExclMask())
                If InStr(1, ArrWordFrom(x), ArrExclMask(z), vbTextCompare) > 0 Then GoTo nextwoord1    ' если такое есть в маске то след слово из первой строки
            Next z
        End If
        'Проверка на соответствие маске длиннны
        If exclLen > 0 Then

            If Len(ArrWordFrom(x)) <= exclLen Then GoTo nextwoord1    ' если слово короткое то след слово из первой строки
        End If
        ' перебор слов из второй строки
        wordCounter = wordCounter + 1
        For y = 0 To UBound(ArrWordTo())
            If ArrWordTo(y) = "" Then GoTo nextwoord2:
            If ArrWordTo(y) <> "" Then
                If ExclMask <> "" Then
                    For z = 1 To UBound(ArrExclMask())
                        If InStr(1, ArrWordTo(y), ArrExclMask(z), vbTextCompare) > 0 Then GoTo nextwoord2    ' если такое есть в маске то след слово из первой строки
                    Next z
                End If
                'Проверка на соответствие маске длиннны
                If exclLen > 0 Then
                    If Len(ArrWordTo(y)) <= exclLen Then GoTo nextwoord2    ' если слово короткое то след слово из первой строки
                End If
                'Вызов процедуры сравнения слов по знакам
                ' считать слова
                result1 = WordsCompareByChar(ArrWordFrom(x), ArrWordTo(y))    ' сравнение первого со вторым словом
                result2 = WordsCompareByChar(ArrWordTo(y), ArrWordFrom(x))    ' сравнение второго с первым
                'Debug.Print ArrWordFrom(x) & "   " & ArrWordTo(y) & " " & result1 & "-" & result2

                If ((result1 + result2) / 2) * 100 >= SamePercent Then    'счиать слово совпавшим
                    If ((result1 + result2) / 2) > avResult Then    ' поиск максимально совпавших слов
                        avResult = ((result1 + result2) / 2)
                        tmpPosition = y
                        flag = True
                    End If
                End If
            End If

nextwoord2:

        Next y
        If flag = True Then
            wordCounterComapared = wordCounterComapared + 1
            resultIndex = resultIndex + ((100 * avResult))
            ArrWordTo(tmpPosition) = ""
            avResult = 0
            flag = False
        End If
nextwoord1:
    Next x
    'Debug.Print WordsCompareByChar(replace(strFrom, " ", "", 1, -1, vbTextCompare), replace(strTo, " ", "", 1, -1, vbTextCompare))

    maximum = wordCounter
    SringCompare = resultIndex    ' + (200 * WordsCompareByChar(Replace(strFrom, " ", "", 1, -1, vbTextCompare), Replace(strTo, " ", "", 1, -1, vbTextCompare)))

End Function

Function WordsCompareByChar(ByVal word1 As String, ByVal word2 As String, Optional excludelen As Long = 0, Optional excludemask As String = "", Optional replacepunctuation As Boolean = True) As Double
    Dim x As Long
    Dim y As Long
    Dim z As Long
    Dim comIndex As Long
    Dim pos1 As Long
    Dim pos2 As Long
    Dim lenWrld2 As Long

    If word1 = "" Or word2 = "" Then Exit Function

    lenWrld2 = Len(word2)

    For x = 1 To Len(word1)

        pos1 = 1
        pos2 = 1
        'Если символ есть в строке и позиция совпадает
        If InStr(1, word2, Mid(word1, x, 1)) > 0 Then

            If lenWrld2 > 2 And Len(word2) > 2 Then
                'Позиция символа в слове 1
                If (InStr(1, word2, Mid(word1, x, 1)) / lenWrld2) < 0.34 Then pos1 = 1
                If (InStr(1, word2, Mid(word1, x, 1)) / lenWrld2) > 0.339 And (InStr(1, word2, Mid(word1, x, 1)) / lenWrld2) < 0.67 Then pos1 = 2
                If (InStr(1, word2, Mid(word1, x, 1)) / lenWrld2) >= 0.67 Then pos1 = 3
                'Позиция символа в слове 1
                If (x / Len(word1)) < 0.34 Then pos2 = 1
                If (x / Len(word1)) >= 0.34 And (x / lenWrld2) < 0.67 Then pos2 = 2
                If (x / lenWrld2) >= 0.67 Then pos2 = 3
            End If

        End If

        If Abs(pos1 - pos2) <= 1 And InStr(1, word2, Mid(word1, x, 1)) > 0 Then
            comIndex = comIndex + 1
            'Debug.Print Mid(word1, x, 1)
            word2 = Replace(word2, Mid(word1, x, 1), "%", 1, 1, vbTextCompare)
        End If

    Next x
    WordsCompareByChar = (comIndex / Len(word1))
    ' WordsCompareByChar = (comIndex)
End Function
Function replacepuctuation(str As String)
    str = Replace(str, "!", "")
    str = Replace(str, "@", "")
    str = Replace(str, "#", "")
    str = Replace(str, "$", "")
    str = Replace(str, "%", "")
    str = Replace(str, "^", "")
    str = Replace(str, "&", "")
    str = Replace(str, "*", "")
    str = Replace(str, "(", "")
    str = Replace(str, ")", "")
    str = Replace(str, "_", " ")
    str = Replace(str, "-", " ")
    str = Replace(str, "+", "")
    str = Replace(str, "=", "")
    str = Replace(str, "{", "")
    str = Replace(str, "[", "")
    str = Replace(str, "}", "")
    str = Replace(str, "]", "")
    str = Replace(str, "|", "")
    str = Replace(str, "\", "")
    str = Replace(str, ";", "")
    str = Replace(str, ":", "")
    str = Replace(str, "'", "")
    str = Replace(str, """", "")
    str = Replace(str, "<", "")
    str = Replace(str, ".", "")
    str = Replace(str, ",", "")
    str = Replace(str, "/", " ")
    str = Replace(str, "?", "")
    str = Replace(str, "`", "")
    str = Replace(str, "~", "")
    replacepuctuation = str
End Function


