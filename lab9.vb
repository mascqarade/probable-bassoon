Sub qwerty()
'
' qwerty Макрос
'

'
    Range("A1:C4").Select
    With Selection.Font
        .ThemeColor = xlThemeColorAccent4
        .TintAndShade = 0.399975585192419
    End With
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 255
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection.Font
        .Name = "Arial"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorAccent4
        .TintAndShade = 0.399975585192419
        .ThemeFont = xlThemeFontNone
    End With
    With Selection.Font
        .Name = "Arial"
        .Size = 16
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorAccent4
        .TintAndShade = 0.399975585192419
        .ThemeFont = xlThemeFontNone
    End With
End Sub
Sub qwerty2()
'
' qwerty2 Макрос
'

'
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "Январь"
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "Февраль"
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "Март"
    Range("D1").Select
    ActiveCell.FormulaR1C1 = "Апрель"
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "Май"
    Range("F1").Select
    ActiveCell.FormulaR1C1 = "Июнь"
    Range("G1").Select
    ActiveCell.FormulaR1C1 = "Июль"
    Range("H1").Select
    ActiveCell.FormulaR1C1 = "Август"
    Range("I1").Select
    ActiveCell.FormulaR1C1 = "Сентябрь"
    Range("J1").Select
    ActiveCell.FormulaR1C1 = "Октябрь"
    Range("K1").Select
    ActiveCell.FormulaR1C1 = "Ноябрь"
    Range("L1").Select
    ActiveCell.FormulaR1C1 = "Декабрь"
    Range("K3").Select
End Sub

Sub qwerty3()
'
' qwerty3 Макрос
'

'
    Range("D5:J12").Select
    With Selection.Font
        .ThemeColor = xlThemeColorAccent4
        .TintAndShade = 0.399975585192419
    End With
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 100
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection.Font
        .Name = "Arial"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorAccent4
        .TintAndShade = 0.399975585192419
        .ThemeFont = xlThemeFontNone
    End With
    With Selection.Font
        .Name = "Arial"
        .Size = 16
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorAccent4
        .TintAndShade = 0.399975585192419
        .ThemeFont = xlThemeFontNone
    End With
End Sub
Sub qwerty4()
'
' qwerty4 Макрос
'

'
    ActiveCell.FormulaR1C1 = "Товар"
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "Цена"
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "Яблоки"
    Range("B2").Select
    ActiveCell.FormulaR1C1 = "60"
    Range("A3").Select
    ActiveCell.FormulaR1C1 = "Хлеб"
    Range("B3").Select
    ActiveCell.FormulaR1C1 = "30"
    Range("A4").Select
    ActiveCell.FormulaR1C1 = "Мука"
    Range("B4").Select
    ActiveCell.FormulaR1C1 = "30"
    Range("A5").Select
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    ActiveCell.FormulaR1C1 = "Сахар"
    Range("B5").Select
    With Selection.Font
        .Name = "Calibri"
        .FontStyle = "обычный"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    ActiveCell.FormulaR1C1 = "40"
    Range("A6").Select
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    ActiveCell.FormulaR1C1 = "Апельсины"
    Range("B6").Select
    With Selection.Font
        .Name = "Calibri"
        .FontStyle = "обычный"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    ActiveCell.FormulaR1C1 = "80"
    Range("A7").Select
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    ActiveCell.FormulaR1C1 = "Вода"
    Range("B7").Select
    With Selection.Font
        .Name = "Calibri"
        .FontStyle = "обычный"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    ActiveCell.FormulaR1C1 = "20"
    Range("A8").Select
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    ActiveCell.FormulaR1C1 = "Молоко"
    Range("B8").Select
    With Selection.Font
        .Name = "Calibri"
        .FontStyle = "обычный"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    ActiveCell.FormulaR1C1 = "50"
    Range("A9").Select
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    ActiveCell.FormulaR1C1 = "Сметана"
    Range("B9").Select
    With Selection.Font
        .Name = "Calibri"
        .FontStyle = "обычный"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    ActiveCell.FormulaR1C1 = "40"
    Range("A10").Select
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    ActiveCell.FormulaR1C1 = "Чай"
    Range("B10").Select
    With Selection.Font
        .Name = "Calibri"
        .FontStyle = "обычный"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    ActiveCell.FormulaR1C1 = "150"
    Range("A11").Select
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    ActiveCell.FormulaR1C1 = "Творог"
    Range("B11").Select
    With Selection.Font
        .Name = "Calibri"
        .FontStyle = "обычный"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    ActiveCell.FormulaR1C1 = "30"
    Range("E3").Select
    Range("A1:B11").Select
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
    ActiveChart.SetSourceData Source:=Range("Лист1!$A$1:$B$11")
End Sub
