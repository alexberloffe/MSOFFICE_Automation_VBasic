Attribute VB_Name = "Módulo11"
Sub AJUSTA_PLANILHAS_PAISES()
'
' Macro1 Macro
'

' Variáveis
    Dim PastaPais As String
    
'
    Application.Calculation = xlAutomatic
    Sheets(Array("FROM TEMPLATE", "OBLIGATORY_TCODE", "OBLIGATORY_SE38")).Select
    Sheets("FROM TEMPLATE").Activate
    Cells.Select
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlCenter
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
    End With
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlCenter
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
    End With
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlCenter
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
    End With
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlCenter
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
    End With
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlCenter
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
    End With
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
    End With
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
    End With
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
    End With
    Selection.UnMerge
    Columns("A:A").Select
    Selection.Copy
    Selection.Insert Shift:=xlToRight
    Columns("A:A").Select
    Application.CutCopyMode = False
    Selection.Copy
    Selection.Insert Shift:=xlToRight
    Range("A1").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "SOURCE"
        
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "COUNTRY"
    Range("B3").Select
    PastaPais = InputBox("Pais de Origem e Modulo", "Qual o país de origem e módulo", "BR_FI")
    ActiveCell.FormulaR1C1 = PastaPais
    Range("B3").Select
    Selection.Copy
    Range("B4").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    
    
    Sheets("FROM TEMPLATE").Activate
    Range("A3").Select
    ActiveCell.FormulaR1C1 = ActiveSheet.Name
    Range("A3").Select
    Selection.Copy
    Range("A4").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Range("I3").Select
    Application.CutCopyMode = False
    Columns("F:F").Select
    Application.CutCopyMode = False
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("F1").Select
    ActiveCell.FormulaR1C1 = "MODULE"
    Sheets("FROM TEMPLATE").Select
    Columns("I:I").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("I1").Select
    ActiveCell.FormulaR1C1 = "PROGRAM"
    Range("I2").Select
    
    Sheets("OBLIGATORY_SE38").Activate
    Range("A3").Select
    ActiveCell.FormulaR1C1 = ActiveSheet.Name
    Range("A3").Select
    Selection.Copy
    Range("A4").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Range("I3").Select
    Application.CutCopyMode = False
    
    Sheets("OBLIGATORY_TCODE").Activate
    Range("A3").Select
    ActiveCell.FormulaR1C1 = ActiveSheet.Name
    Range("A3").Select
    Selection.Copy
    Range("A4").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Range("I3").Select
    Application.CutCopyMode = False
    
    MsgBox "Verifique a coluna de categorias - tem que ser L"

    
End Sub

