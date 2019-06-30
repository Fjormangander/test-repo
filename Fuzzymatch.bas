Attribute VB_Name = "Module1"
Option Explicit




Public Function FuzzyVLOOKUP(�������_�������� As Range, ������� As Range, �������_���_������ As Boolean, Optional �������_���������� As Long = 50, Optional �����������_�����_����� As Long = 0, Optional �����_���������� As String = "")

    Dim x As Long
    Dim y As Double
    Dim Score As Double
    Dim min As Double
    Dim Max As Long
    Dim max2 As Long
    Dim d As Long
    For x = 1 To �������.Count
        Score = SringCompare(�������_��������.Value, �������(x).Value, 50, �����������_�����_�����, �����_����������, Max)
        Score = Score + SringCompare(�������(x).Value, �������_��������.Value, 50, �����������_�����_�����, �����_����������, max2)
        'If min > Score Then min = Score
        'If Max < Score Then Max = Score
        If Score > y Then
            y = Score

            If �������_���_������ = False Then
                d = max2
'Debug.Print Score / ((Max + d) * 100)
                If Score / ((Max + d) * 100) >= (�������_���������� / 100) Then

                    FuzzyVLOOKUP = �������(x).Value

                Else: FuzzyVLOOKUP = "MISS"
                End If
            End If
            If �������_���_������ = True Then
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

    ' ������� ����� ����������
    strFrom = UCase(replacepuctuation(strFrom))
    strTo = UCase(replacepuctuation(strTo))

    'maximum = wordscounter(strFrom, exclLen, ExclMask)


    '������ � ������ ����
    ArrWordFrom() = Split(strFrom, " ", -1, vbTextCompare)
    ArrWordTo() = Split(strTo, " ", -1, vbTextCompare)
    '���� ����� �� ������ ������� ������ ����������
    If ExclMask <> "" Then ArrExclMask() = Split(ExclMask, ";", -1, vbTextCompare)
    ' ������� ������ ����
    For x = 0 To UBound(ArrWordFrom())
        If ArrWordFrom(x) = "" Then GoTo nextwoord1:
        '�������� �� ������������ ����� ����������
        If ExclMask <> "" Then
            For z = 0 To UBound(ArrExclMask())
                If InStr(1, ArrWordFrom(x), ArrExclMask(z), vbTextCompare) > 0 Then GoTo nextwoord1    ' ���� ����� ���� � ����� �� ���� ����� �� ������ ������
            Next z
        End If
        '�������� �� ������������ ����� �������
        If exclLen > 0 Then

            If Len(ArrWordFrom(x)) <= exclLen Then GoTo nextwoord1    ' ���� ����� �������� �� ���� ����� �� ������ ������
        End If
        ' ������� ���� �� ������ ������
        wordCounter = wordCounter + 1
        For y = 0 To UBound(ArrWordTo())
            If ArrWordTo(y) = "" Then GoTo nextwoord2:
            If ArrWordTo(y) <> "" Then
                If ExclMask <> "" Then
                    For z = 1 To UBound(ArrExclMask())
                        If InStr(1, ArrWordTo(y), ArrExclMask(z), vbTextCompare) > 0 Then GoTo nextwoord2    ' ���� ����� ���� � ����� �� ���� ����� �� ������ ������
                    Next z
                End If
                '�������� �� ������������ ����� �������
                If exclLen > 0 Then
                    If Len(ArrWordTo(y)) <= exclLen Then GoTo nextwoord2    ' ���� ����� �������� �� ���� ����� �� ������ ������
                End If
                '����� ��������� ��������� ���� �� ������
                ' ������� �����
                result1 = WordsCompareByChar(ArrWordFrom(x), ArrWordTo(y))    ' ��������� ������� �� ������ ������
                result2 = WordsCompareByChar(ArrWordTo(y), ArrWordFrom(x))    ' ��������� ������� � ������
                'Debug.Print ArrWordFrom(x) & "   " & ArrWordTo(y) & " " & result1 & "-" & result2

                If ((result1 + result2) / 2) * 100 >= SamePercent Then    '������ ����� ���������
                    If ((result1 + result2) / 2) > avResult Then    ' ����� ����������� ��������� ����
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
        '���� ������ ���� � ������ � ������� ���������
        If InStr(1, word2, Mid(word1, x, 1)) > 0 Then

            If lenWrld2 > 2 And Len(word2) > 2 Then
                '������� ������� � ����� 1
                If (InStr(1, word2, Mid(word1, x, 1)) / lenWrld2) < 0.34 Then pos1 = 1
                If (InStr(1, word2, Mid(word1, x, 1)) / lenWrld2) > 0.339 And (InStr(1, word2, Mid(word1, x, 1)) / lenWrld2) < 0.67 Then pos1 = 2
                If (InStr(1, word2, Mid(word1, x, 1)) / lenWrld2) >= 0.67 Then pos1 = 3
                '������� ������� � ����� 1
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


