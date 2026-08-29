
Sub USPTU()
    Application.ScreenUpdating = False
    ActiveDocument.Repaginate
    CreateLogFolder
    BOT_Cleanup_And_Highlight1
    BOT_Headers2
    BOT_ParaCombined3
    BOT_CheckDocumentSetup4
    BOT_RemoveAllItalic5
    Application.ScreenUpdating = True
'msgBox "Проверка завершена"
End Sub
Sub CreateLogFolder()
    Dim logFolder As String
    logFolder = "C:\tmp"  ' Папка для логов

    ' Проверяем и создаём папку, если её нет
    If Dir(logFolder, vbDirectory) = "" Then
        MkDir logFolder
        Debug.Print "Папка для логов создана: " & logFolder, vbInformation
    End If
End Sub
Sub LogMessage()
    Dim fnum As Integer
    Dim logPath As String
    Static counter As Integer

    logPath = "C:\tmp\vba_log.txt"  ' Полный путь к файлу

    ' Проверяем существует ли папка (на всякий случай)
    If Dir("C:\tmp", vbDirectory) = "" Then
        Debug.Print "Папка для логов не существует! Сначала выполните CreateLogFolder", vbCritical
        Exit Sub
    End If

    ' Записываем сообщение в лог
    fnum = FreeFile
    Open logPath For Append As #fnum
    counter = counter + 1
    Print #fnum, Format(Now, "yyyy-mm-dd hh:nn:ss") & " - " & "Завершилась процедура № " & counter
    Close #fnum
End Sub
Function GetMainRange() As Range

    Dim doc As Document
    Dim totalPages As Long
    Dim startRange As Range
    Dim endRange As Range
    Dim workRange As Range

    Set doc = ActiveDocument

    totalPages = doc.ComputeStatistics(wdStatisticPages)


    If totalPages < 4 Then
        Set GetMainRange = Nothing
        Exit Function
    End If
    Set startRange = doc.GoTo(What:=wdGoToPage, _
        Which:=wdGoToAbsolute, _
        count:=3)

    Set endRange = doc.GoTo(What:=wdGoToPage, _
        Which:=wdGoToAbsolute, _
        count:=totalPages)

    Set workRange = doc.Range(startRange.Start, endRange.Start)

    Set GetMainRange = workRange

End Function
Sub BOT_Cleanup_And_Highlight1()

    Dim rng As Range
    Dim baseRng As Range
    Dim highlightPatterns As Variant
    Dim dashTypes As Variant
    Dim i As Long
    Dim dash As Variant

    Set baseRng = GetMainRange
    If baseRng Is Nothing Then Exit Sub

    ' --- ШАБЛОНЫ ДЛЯ ПОДСВЕТКИ ---
    highlightPatterns = Array( _
        " ^s", _
        "^w^p", _
        "^l", _
        "^p ", _
        "^s^s", _
        " ^p", _
        "^p^t", _
        "^t^p")

    ' --- ПОДСВЕТКА ПРОБЛЕМНЫХ КОМБИНАЦИЙ ---
    For i = 0 To UBound(highlightPatterns)

        Set rng = baseRng.Duplicate

        With rng.Find
            .ClearFormatting
            .Text = highlightPatterns(i)
            .Wrap = wdFindStop

            Do While .Execute
                rng.HighlightColorIndex = wdRed
                rng.Collapse wdCollapseEnd
            Loop
        End With

    Next i

    ' --- ДВОЙНЫЕ ПРОБЕЛЫ ---
    Set rng = baseRng.Duplicate
    With rng.Find
        .Text = "  "
        .Wrap = wdFindStop

        Do While .Execute
            rng.HighlightColorIndex = wdRed

            If rng.Paragraphs(1).Range.Comments.count = 0 Then
                rng.Comments.Add Range:=rng.Paragraphs(1).Range, _
                    Text:="В тексте присутствуют лишние пробелы."
            End If

            rng.Collapse wdCollapseEnd
        Loop
    End With

    ' --- ЛИШНИЕ ПЕРЕНОСЫ (3 подряд) ---
    Set rng = baseRng.Duplicate
    With rng.Find
        .Text = "^p^p^p"
        .Wrap = wdFindStop

        Do While .Execute
            rng.HighlightColorIndex = wdRed
            rng.Comments.Add Range:=rng, _
                Text:="В тексте присутствуют лишние переносы строки."
            rng.Collapse wdCollapseEnd
        Loop
    End With

    ' --- ДЕФИСЫ И ТИРЕ ---
    dashTypes = Array("- ", " -", "^$^=", "^=^$", "^+")

    For Each dash In dashTypes

        Set rng = baseRng.Duplicate

        With rng.Find
            .ClearFormatting
            .Text = dash
            .Wrap = wdFindStop

            Do While .Execute
                rng.HighlightColorIndex = wdRed
                rng.Comments.Add Range:=rng, _
                    Text:="Некорректное использование дефиса или тире."
                rng.Collapse wdCollapseEnd
            Loop
        End With

    Next dash

    LogMessage

End Sub

Sub BOT_Headers2()
    Dim para As Paragraph
    Dim workRng As Range
    Set workRng = GetMainRange
    If workRng Is Nothing Then Exit Sub

    Dim targetWords As Variant
    targetWords = Array( _
        "Введение", _
        "Содержание", _
        "Заключение", _
        "Список использованной литературы", _
        "Список использованных источников", _
        "Список литературы", _
        "Список источников", _
        "Реферат")

    Dim txt As String
    Dim rng As Range
    Dim word As Variant

    For Each word In targetWords

        Set rng = workRng

        With rng.Find
            .Text = word
            .MatchCase = False
            .MatchWholeWord = True
            .Wrap = wdFindStop

            Do While .Execute

'                If rng.Paragraphs(1).Range.Characters.count <= (Len(word) + 3) Then
                Set para = rng.Paragraphs(1)

                If rng.Paragraphs(1).Range.Text <> UCase(rng.Paragraphs(1).Range.Text) Then
'                        rng.Paragraphs(1).Range.Text = UCase(rng.Paragraphs(1).Range.Text)
                    rng.Comments.Add Range:=rng, _
                        Text:="Заголовки «РЕФЕРАТ», «ВВЕДЕНИЕ» , «СОДЕРЖАНИЕ», «ЗАКЛЮЧЕНИЕ», «СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ» должны быть написаны большими буквами."
                End If

                If Right(rng.Paragraphs(1).Range.Text, 2) Like "[.;:,]" & Chr(13) Then
'                        With rng.Paragraphs(1).Range.Find
'                            .ClearFormatting
'                            .Text = "[.;:,]"
'                            .MatchWildcards = True
'                            .Forward = False
'                            .Wrap = wdFindStop
'                            .Replacement.Text = ""
'                            .Execute Replace:=wdReplaceOne
'                        End With
                    rng.Comments.Add Range:=rng, _
                        Text:="В конце заголовка не должен стоять знак препинания."
                End If

                With rng.ParagraphFormat
                    If .Alignment <> wdAlignParagraphCenter _
                        Or .FirstLineIndent <> 0 _
                        Or .LeftIndent <> 0 Then

'                            .Alignment = wdAlignParagraphCenter
'                            .FirstLineIndent = 0
'                            .LeftIndent = 0

                        rng.Comments.Add Range:=rng, _
                            Text:="Заголовки должны быть выровнены по центру без абзацного отступа."
                    End If
                End With
                If Not para.Previous.Range Is Nothing Then
                    If Right(para.Previous.Range, 2) <> Chr(12) & Chr(13) And Left(rng.Paragraphs(1).Range, 1) <> Chr(12) Then
'                            para.Previous.Range.InsertAfter Chr(12) & Chr(13)
                        rng.Comments.Add Range:=rng, _
                            Text:="Перед заголовком рекомендуется ставить разры страницы."
                    End If
                End If
                If Not para.Next.Range Is Nothing Then
                    If Left(para.Next.Range, 1) <> Chr(13) Then
'                            para.Range.InsertAfter Chr(13)
                        rng.Comments.Add Range:=rng, _
                            Text:="После заголовка необходимо ставить перенос строки."
                    End If
                End If
'                End If
                rng.Collapse wdCollapseEnd

            Loop
        End With

    Next word
End Sub


Sub BOT_ParaCombined3()
    'Application.ScreenUpdating = False
    Dim para        As Paragraph
    Dim foundImage  As Boolean
    Dim markerText  As String
    Dim isFontCorrect As Boolean
    Dim workRng     As Range
    Set workRng = GetMainRange
    If workRng Is Nothing Then Exit Sub
    foundImage = False
    Dim txtRng      As Range
    Dim level       As Long
    Dim requiredIndent As Double
    Dim commentText As String

    commentText = "Допускается в качестве маркера первого уровня (с абзацным отступом 1,25 см) тире («" & Chr(150) & "»)," & Chr(13) & _
        "в качестве маркера второго уровня (с абзацным отступом 2,5 см) буква со скобкой («а)»)," & Chr(13) & _
        "в качестве маркера третьего уровня (с абзацным отступом 3,75 см) цифра со скобкой: («1)»)."


    For Each para In workRng.Paragraphs

        ' =====================================================
        ' >>> BOT_CombinedListCheck5 <<<
        ' =====================================================

        If para.Range.listFormat.listType <> wdListNoNumbering Then
            markerText = para.Range.listFormat.ListString
            level = para.Range.listFormat.ListLevelNumber
            If (Not (para.Previous Is Nothing)) And para.Previous.Range.listFormat.listType = wdListNoNumbering Then

                If Trim(Right(para.Previous.Range.Text, 2)) <> ":" & Chr(13) Then
'                    para.Previous.Range.Characters.Last.InsertBefore ":"
'
'                    para.Previous.Range.ParagraphFormat.FirstLineIndent = CentimetersToPoints(1.25)
'                    para.Previous.Range.ParagraphFormat.LeftIndent = CentimetersToPoints(0)
                    para.Previous.Range.HighlightColorIndex = wdRed
                    para.Previous.Range.Comments.Add Range:=para.Range, Text:="Перед началом списка должно стоять двоеточие («:»)."
                End If
            End If

            Select Case level

                Case 1

                    requiredIndent = CentimetersToPoints(1.25)

                    If markerText <> ChrW(61485) Then
                        para.Range.HighlightColorIndex = wdRed
                        para.Range.Comments.Add Range:=para.Range, Text:=commentText
                    End If

                    If Abs(para.Range.ParagraphFormat.FirstLineIndent - requiredIndent) > 0.1 Then
                        para.Range.ParagraphFormat.FirstLineIndent = requiredIndent
                    End If

                Case 2

                    requiredIndent = CentimetersToPoints(2.5)

                    If Not (markerText Like "[а-я])" Or markerText Like "[a-z])") Then
                        para.Range.HighlightColorIndex = wdRed
                        para.Range.Comments.Add Range:=para.Range, Text:=commentText
                    End If

                    If Abs(para.Range.ParagraphFormat.FirstLineIndent - requiredIndent) > 0.1 Then
                        para.Range.ParagraphFormat.FirstLineIndent = requiredIndent
                    End If

                Case 3

                    requiredIndent = CentimetersToPoints(3.75)

                    If Not (markerText Like "#)" Or markerText Like "##)") Then
                        para.Range.HighlightColorIndex = wdRed
                        para.Range.Comments.Add Range:=para.Range, Text:=commentText
                    End If

                    If Abs(para.Range.ParagraphFormat.FirstLineIndent - requiredIndent) > 0.1 Then
                        para.Range.ParagraphFormat.FirstLineIndent = requiredIndent
                    End If

            End Select

            If Not (Trim(Left(para.Range.Text, 1)) Like "[A-Z]" Or Trim(Left(para.Range.Text, 1)) Like "[А-Я]") Then
'                para.Range.Characters(1).Text = UCase(para.Range.Characters(1).Text)
                para.Range.HighlightColorIndex = wdRed
                para.Range.Comments.Add Range:=para.Range, Text:="Элементы списка должны начинаться с прописной буквы."
            End If

            If Not para.Next Is Nothing Then
'                If (Trim(Right(para.Range.Text, 2)) = "." & Chr(13) Or Trim(Right(para.Range.Text, 2)) = ";" & Chr(13) Or Trim(Right(para.Range.Text, 2)) = "," & Chr(13)) _
'                   And para.Next.Range.listFormat.ListLevelNumber > para.Range.listFormat.ListLevelNumber _
'                   And para.Next.Range.listFormat.listType <> wdListNoNumbering Then
'
'                With para.Range.Find
'                    .ClearFormatting
'                    .Text = "[.,;]"
'                    .MatchWildcards = True
'                    .Forward = False
'                    .Wrap = wdFindStop
'                    .Replacement.Text = ":"
'                    .Execute Replace:=wdReplaceOne
'                End With

                If Trim(Right(para.Range.Text, 2)) <> ":" & Chr(13) _
                    And para.Next.Range.listFormat.ListLevelNumber > para.Range.listFormat.ListLevelNumber _
                    And para.Next.Range.listFormat.listType <> wdListNoNumbering Then

'                para.Range.Characters.Last.InsertBefore ":"

                    para.Range.HighlightColorIndex = wdRed
                    para.Range.Comments.Add Range:=para.Range, Text:="Промежуточные элементы списка должны заканчиваться точкой с запятой («;»), последний заканчивается точкой («.»). В случае наличия вложенных элементов (список внутри списка) соответствующий вложенный элемент заканчивается двоеточием («:»)."

'                 ElseIf (Trim(Right(para.Range.Text, 2)) = ":" & Chr(13) Or Trim(Right(para.Range.Text, 2)) = ";" & Chr(13) Or Trim(Right(para.Range.Text, 2)) = "," & Chr(13)) _
'                        And para.Next.Range.listFormat.listType = wdListNoNumbering Then
'
'                 With para.Range.Find
'                     .ClearFormatting
'                     .Text = "[:,;]"
'                     .MatchWildcards = True
'                     .Forward = False
'                     .Wrap = wdFindStop
'                     .Replacement.Text = "."
'                     .Execute Replace:=wdReplaceOne
'                 End With

                ElseIf (Trim(Right(para.Range.Text, 2)) <> "." & Chr(13) And para.Next.Range.listFormat.listType = wdListNoNumbering) Then

'                    para.Range.Characters.Last.InsertBefore "."

                    para.Range.HighlightColorIndex = wdRed
                    para.Range.Comments.Add Range:=para.Range, Text:="Промежуточные элементы списка должны заканчиваться точкой с запятой («;»), последний заканчивается точкой («.»). В случае наличия вложенных элементов (список внутри списка) соответствующий вложенный элемент заканчивается двоеточием («:»)."

'                ElseIf (Trim(Right(para.Range.Text, 2)) = ":" & Chr(13) Or Trim(Right(para.Range.Text, 2)) = "." & Chr(13) Or Trim(Right(para.Range.Text, 2)) = "," & Chr(13)) _
'                       And (para.Next.Range.listFormat.ListLevelNumber = para.Range.listFormat.ListLevelNumber Or para.Next.Range.listFormat.ListLevelNumber < para.Range.listFormat.ListLevelNumber) _
'                       And para.Next.Range.listFormat.listType <> wdListNoNumbering Then
'
'                With para.Range.Find
'                    .ClearFormatting
'                    .Text = "[.,:]"
'                    .MatchWildcards = True
'                    .Forward = False
'                    .Wrap = wdFindStop
'                    .Replacement.Text = ";"
'                    .Execute Replace:=wdReplaceOne
'                End With

                ElseIf Trim(Right(para.Range.Text, 2)) <> ";" & Chr(13) _
                    And (para.Next.Range.listFormat.ListLevelNumber = para.Range.listFormat.ListLevelNumber Or para.Next.Range.listFormat.ListLevelNumber < para.Range.listFormat.ListLevelNumber) _
                    And para.Next.Range.listFormat.listType <> wdListNoNumbering Then

'                para.Range.Characters.Last.InsertBefore ";"

                    para.Range.HighlightColorIndex = wdRed
                    para.Range.Comments.Add Range:=para.Range, Text:="Промежуточные элементы списка должны заканчиваться точкой с запятой («;»), последний заканчивается точкой («.»). В случае наличия вложенных элементов (список внутри списка) соответствующий вложенный элемент заканчивается двоеточием («:»)."

                End If
            End If

        ElseIf Left(Trim(para.Range.Text), 1) = "-" Or _
            Left(Trim(para.Range.Text), 1) = "^=" Or _
            IsNumeric(Left(Trim(para.Range.Text), 1)) Then

            para.Range.HighlightColorIndex = wdRed
            para.Range.Comments.Add Range:=para.Range, Text:="Возможно, вы хотели создать список. В нумерации заголовков не должно быть точки. Заголовки должны быть выделены жирным шрифтом." & Chr(13) & "Пример:" & Chr(13) & "1    Заголовок" & Chr(13) & "1.1     Подзаголовок"

        End If

       ' =====================================================
       ' >>> RedactedHighlightParagraphsFont8 <<<
       ' =====================================================
        Set txtRng = para.Range
        isFontCorrect = True

       ' Проверяем шрифт и размер
        With txtRng
            If .InlineShapes.count = 0 And .ShapeRange.count = 0 And .Information(wdWithInTable) = False And txtRng.Characters.count > 1 Then
                txtRng.End = txtRng.End - 1
                If .Font.Name <> "Times New Roman" Or .Font.Size <> 14 Or .ParagraphFormat.LineSpacingRule <> wdLineSpace1pt5 _
                    Or .ParagraphFormat.LeftIndent <> CentimetersToPoints(0) Or .ParagraphFormat.RightIndent <> CentimetersToPoints(0) _
                    Or .ParagraphFormat.SpaceBefore <> 0 Or .ParagraphFormat.SpaceAfter <> 0 Then
'               .Font.Name = "Times New Roman"
'               .Font.Size = 14
'               .ParagraphFormat.LineSpacingRule = wdLineSpace1pt5
'               .ParagraphFormat.LeftIndent = CentimetersToPoints(0)
'               .ParagraphFormat.RightIndent = CentimetersToPoints(0)
'               .ParagraphFormat.SpaceBefore = 0
'               .ParagraphFormat.SpaceAfter = 0
                    .HighlightColorIndex = wdRed
                    .Comments.Add Range:=para.Range, Text:="Требуется шрифт Times New Roman, размер шрифта 14 и межстрочный интервал 1,5."
                End If

            ElseIf para.Range.Information(wdWithInTable) = True And .InlineShapes.count = 0 And para.Range.ShapeRange.count = 0 Then
                If .Font.Name <> "Times New Roman" Or (.Font.Size <> 14 And .Font.Size <> 12 And .Font.Size <> 10) Or (.ParagraphFormat.LineSpacingRule <> wdLineSpace1pt5 And .ParagraphFormat.LineSpacingRule <> wdLineSpaceSingle) Then
                    para.Range.HighlightColorIndex = wdRed
                    On Error Resume Next
                    para.Range.Comments.Add Range:=para.Range, Text:="В таблице требуется шрифт Times New Roman, размер шрифта 14, 12 или 10 и межстрочный интервал 1 или 1,5. (Предпочтение следует отдавать шрифту с размером 12 и межстрочному интервалу 1)."
                End If
            End If
        End With
        If para.Range.InlineShapes.count = 0 And para.Range.ShapeRange.count = 0 And para.Range.Information(wdWithInTable) = False And para.Range.listFormat.listType = wdListNoNumbering _
            And Left(para.Range.Text, 8) <> "ВВЕДЕНИЕ" _
            And Left((para.Range.Text), 10) <> "ЗАКЛЮЧЕНИЕ" _
            And Left((para.Range.Text), 10) <> "СОДЕРЖАНИЕ" _
            And Left((para.Range.Text), 6) <> "СПИСОК" _
            And Left((para.Range.Text), 7) <> "РЕФЕРАТ" _
            And Left((para.Range.Text), 7) <> "Рисунок" _
            And Left((para.Range.Text), 7) <> "Таблица" _
            And Left((para.Range.Text), 1) <> Chr(13) _
            And Left((para.Range.Text), 2) <> vbFormFeed & Chr(13) Then

            If (para.Range.ParagraphFormat.FirstLineIndent >= CentimetersToPoints(1.251) Or para.Range.ParagraphFormat.FirstLineIndent <= CentimetersToPoints(1.249)) Or para.Range.ParagraphFormat.Alignment <> wdAlignParagraphJustify Then
    '           para.Range.ParagraphFormat.FirstLineIndent = CentimetersToPoints(1.25)
    '           para.Range.ParagraphFormat.Alignment = wdAlignParagraphJustify
                para.Range.HighlightColorIndex = wdRed
                para.Range.Comments.Add Range:=para.Range, Text:="Текст должен быть с абзацным отступом 1,25 см и выровнен по ширине."
            End If
            If Left(Trim(para.Range.Text), 1) = "-" Or _
                Left(Trim(para.Range.Text), 1) = "^=" Or _
                (IsNumeric(Left(Trim(para.Range.Text), 1)) And (Mid(para.Range.Text, 2, 1) = ".")) Or _
                (IsNumeric(Left(Trim(para.Range.Text), 1)) And (Mid(para.Range.Text, 2, 1) = ")")) Then

                para.Range.HighlightColorIndex = wdRed
                para.Range.Comments.Add Range:=para.Range, Text:="Возможно, вы хотели создать список. В нумерации заголовков не должно быть точки. Заголовки должны быть выделены жирным шрифтом." & Chr(13) & "Пример:" & Chr(13) & "1    Заголовок" & Chr(13) & "1.1     Подзаголовок"

            End If

            If InStr(para.Range.Text, "=") > 0 Then

               'para.Range.HighlightColorIndex = wdRed
                para.Range.Comments.Add Range:=para.Range, Text:="Возможно здесь должна быть формула" & Chr(13) & "A+B=C              (1.1)" & Chr(13) & "(формула прописывается через «Вставка» -> «Уравнение»; формула выравнивается по левому краю, с абзацным отступом 1,25 см, номер формулы проставляется в скобках с выравниванием по правому краю)."

            End If
        End If

       ' =====================================================
       ' >>>BOT_AlignPictureTableName10 <<<
       ' =====================================================

        If Not (para Is Nothing Or para.Previous Is Nothing Or para.Next Is Nothing) Then

            If para.Range.InlineShapes.count > 0 Then

                If Left((para.Previous.Range.Text), 1) <> Chr(13) Then
'                   para.Previous.Range.InsertAfter Chr(13)
                    para.Previous.Range.HighlightColorIndex = wdRed
                    para.Previous.Range.Comments.Add Range:=para.Range, Text:="Неправильно оформлен рисунок. Перед самим рисунком и после подписи к рисунку должны быть пустые строки."

                End If

                If para.Range.ParagraphFormat.Alignment <> wdAlignParagraphCenter _
                    Or para.Range.ParagraphFormat.FirstLineIndent <> 0 Then

'               para.Range.ParagraphFormat.FirstLineIndent = CentimetersToPoints(0)
'               para.Range.ParagraphFormat.Alignment = wdAlignParagraphCenter

                    para.Range.HighlightColorIndex = wdRed
                    para.Range.Comments.Add Range:=para.Range, Text:="Неправильно оформлен рисунок. Рисунок и подпись к рисунку должны быть выравнены по центру и не должны иметь абзацный отступ."
         '                Else: para.Previous.Range.Font.Color = wdBlack
                End If

                foundImage = True
            ElseIf foundImage Then

                If Not ( _
                    Left(para.Range.Text, 8) = "Рисунок " And _
                    IsNumeric(Mid(para.Range.Text, 9, 1)) And _
                    Mid(para.Range.Text, 10, 1) = "." And _
                    IsNumeric(Mid(para.Range.Text, 11, 1)) And _
                    Mid(para.Range.Text, 12, 3) = Chr(32) & Chr(150) & Chr(32) And _
                    (Mid(para.Range.Text, 15, 1) Like "[А-Я]" Or Mid(para.Range.Text, 15, 1) Like "[A-Z]") And _
                    Right(para.Range.Text, 2) <> ";" & Chr(13) And _
                    Right(para.Range.Text, 2) <> "." & Chr(13) And _
                    Right(para.Range.Text, 2) <> "," & Chr(13)) And Not ( _
                    Left(para.Range.Text, 8) = "Рисунок " And _
                    IsNumeric(Mid(para.Range.Text, 9, 1)) And _
                    Mid(para.Range.Text, 10, 3) = Chr(32) & Chr(150) & Chr(32) And _
                    (Mid(para.Range.Text, 13, 1) Like "[А-Я]" Or Mid(para.Range.Text, 13, 1) Like "[A-Z]") And _
                    Right(para.Range.Text, 2) <> ";" & Chr(13) And _
                    Right(para.Range.Text, 2) <> "." & Chr(13) And _
                    Right(para.Range.Text, 2) <> "," & Chr(13)) Then

                    para.Range.HighlightColorIndex = wdRed
                    para.Range.Comments.Add Range:=para.Range, Text:="Неправильно оформлен рисунок. Пример подписи: «Рисунок 1.1 " & Chr(150) & " Название» ИЛИ «Рисунок 1 " & Chr(150) & " Название» (используется тире). Перед самим рисунком и после подписи к рисунку должны быть пустые строки. Рисунок и подпись к рисунку должны быть выравнены по центру и не должны иметь абзацный отступ."

                ElseIf para.Range.ParagraphFormat.Alignment <> wdAlignParagraphCenter _
                    Or para.FirstLineIndent <> 0 Then
'       para.Range.ParagraphFormat.FirstLineIndent = CentimetersToPoints(0)
'       para.Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
                    para.Previous.Range.HighlightColorIndex = wdRed
                    para.Previous.Range.Comments.Add Range:=para.Range, Text:="Неправильно оформлен рисунок. Рисунок и подпись к рисунку должны быть выравнены по центру и не должны иметь абзацный отступ."

       'Else: para.Next.Range.Font.Color = wdBlack
                End If

                If Left((para.Next.Range.Text), 1) <> Chr(13) Then
                    para.Next.Range.HighlightColorIndex = wdRed
                    para.Next.Range.Comments.Add Range:=para.Range, Text:="Неправильно оформлен рисунок. Перед самим рисунком и после подписи к рисунку должны быть пустые строки."

                End If

                foundImage = False
            End If

            If para.Next.Range.Tables.count > 0 And para.Range.Tables.count = 0 Then

                If Left((para.Previous.Range.Text), 1) <> Chr(13) Then
'               para.Range.InsertBefore Chr(13)
                    para.Range.HighlightColorIndex = wdRed
                    para.Range.Comments.Add Range:=para.Range, Text:="Неправильно оформлена таблица. После самой таблицы и до подписи к таблице должны быть пустые строки."

                End If

                If Not (Left((para.Range.Text), 8) = "Таблица " And IsNumeric(Mid(para.Range.Text, 9, 1)) _
                    And Mid(para.Range.Text, 10, 3) = Chr(32) & Chr(150) & Chr(32) _
                    And (Mid(para.Range.Text, 13, 1) Like "[А-Я]" Or Mid(para.Range.Text, 13, 1) Like "[A-Z]") _
                    And (Right((para.Range.Text), 2) <> ";" & Chr(13) And Right((para.Range.Text), 2) <> "." & Chr(13))) _
                    _
                    And Not (Left((para.Range.Text), 8) = "Таблица " _
                    And IsNumeric(Mid(para.Range.Text, 9, 1)) And (Mid(para.Range.Text, 10, 1) = ".") And IsNumeric(Mid(para.Range.Text, 11, 1)) _
                    And Mid(para.Range.Text, 12, 3) = Chr(32) & Chr(150) & Chr(32) _
                    And (Mid(para.Range.Text, 15, 1) Like "[А-Я]" Or Mid(para.Range.Text, 15, 1) Like "[A-Z]") _
                    And (Right((para.Range.Text), 2) <> ";" & Chr(13) And Right((para.Range.Text), 2) <> "." & Chr(13))) Then

                    para.Range.HighlightColorIndex = wdRed
                    para.Range.Comments.Add Range:=para.Range, Text:="Неправильно оформлена таблица. Пример подписи: «Таблица 1.1 " & Chr(150) & " Название» ИЛИ «Таблица 1 " & Chr(150) & " Название» (используется тире). После самой таблицы и до подписи к таблице должны быть пустые строки. Таблица и подпись к таблице не должны иметь абзацный отступ, подпись к таблице должна быть выравнена по левому краю."

           'Else: para.Previous.Range.Font.Color = wdBlack

                ElseIf para.Range.ParagraphFormat.Alignment <> wdAlignParagraphLeft _
                    Or para.FirstLineIndent <> 0 Then

'               para.Range.ParagraphFormat.FirstLineIndent = CentimetersToPoints(0)
'               para.Range.ParagraphFormat.Alignment = wdAlignParagraphLeft
                    para.Range.HighlightColorIndex = wdRed
                    para.Range.Comments.Add Range:=para.Range, Text:="Неправильно оформлена таблица. Таблица и подпись к таблице не должны иметь абзацный отступ, подпись к таблице должна быть выравнена по левому краю."

                End If

            ElseIf para.Previous.Range.Tables.count > 0 And para.Range.Tables.count = 0 Then
                If Left((para.Range.Text), 1) <> Chr(13) Then

'               para.Range.InsertBefore Chr(13)
                    para.Previous.Range.HighlightColorIndex = wdRed
                    para.Previous.Range.Comments.Add Range:=para.Range, Text:="Неправильно оформлена таблица. После самой таблицы и до подписи к таблице должны быть пустые строки."

               'Else: para.Range.Font.Color = wdBlack
                End If
            End If
        End If

    Next para

    LogMessage
       'Application.ScreenUpdating = True
End Sub

Sub BOT_CheckDocumentSetup4()

    Dim isMarginsCorrect As Boolean
    Dim hasNumbers  As Boolean

    isMarginsCorrect = True

    ' --- ПРОВЕРКА ПОЛЕЙ ---

    With ActiveDocument.PageSetup

        Const TOLERANCE As Double = 0.1

        If Abs(.TopMargin - CentimetersToPoints(2)) > TOLERANCE Then isMarginsCorrect = False
        If Abs(.BottomMargin - CentimetersToPoints(2)) > TOLERANCE Then isMarginsCorrect = False
        If Abs(.LeftMargin - CentimetersToPoints(2)) > TOLERANCE Then isMarginsCorrect = False
        If Abs(.RightMargin - CentimetersToPoints(1)) > TOLERANCE Then isMarginsCorrect = False
    End With

    If Not isMarginsCorrect Then
'        On Error Resume Next
'        With ActiveDocument.PageSetup
'            .TopMargin = CentimetersToPoints(2)
'            .BottomMargin = CentimetersToPoints(2)
'            .LeftMargin = CentimetersToPoints(2)
'            .RightMargin = CentimetersToPoints(1)
'        End With
'        On Error GoTo 0
        ActiveDocument.Content.Select
        selection.Comments.Add Range:=selection.Range, _
            Text:="Поля документа не соответствуют требованиям: " & _
            "верхнее: 2 см; нижнее: 2 см; левое: 2 см; правое: 1 см."
    End If

    ' --- ПРОВЕРКА НУМЕРАЦИИ СТРАНИЦ ---
    On Error Resume Next
    hasNumbers = (ActiveDocument.Sections(1). _
        Footers(wdHeaderFooterPrimary).Range.Fields.count > 0)
    On Error GoTo 0

    If Not hasNumbers Then
        With ActiveDocument.Content
            .Collapse Direction:=wdCollapseStart
            .Comments.Add Range:=.Duplicate, _
                Text:="В документе отсутствует нумерация страниц. " & _
                "Страницы должны быть пронумерованы внизу, по центру."
        End With
    End If

    LogMessage

End Sub
Sub BOT_RemoveAllItalic5()
    Dim rng         As Range
    Set rng = GetMainRange
    If rng Is Nothing Then Exit Sub

    With rng.Find
        .ClearFormatting
        .Font.Italic = True
        .Text = ""
        .Forward = True
        .Wrap = wdFindStop
        .Format = True
        .MatchWholeWord = True
        Do While .Execute
            rng.HighlightColorIndex = wdRed
            rng.Collapse Direction:=wdCollapseEnd
            rng.Comments.Add Range:=rng, _
                Text:="Использование курсива запрещено."
        Loop
    End With

    Application.ScreenUpdating = True
    LogMessage
End Sub



