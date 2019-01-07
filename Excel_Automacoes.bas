Attribute VB_Name = "Módulo2"
Sub SKY_UNIC_Atualizar_Dashboard()
Attribute SKY_UNIC_Atualizar_Dashboard.VB_ProcData.VB_Invoke_Func = " \n14"
'
' SKY_UNIC_Atualizar_Dashboard Macro
'

'
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Windows("0_SSU_Old Dashboard.MHTML").Activate
    Sheets("Sheet1").Select
    Range("A1").Select
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    Selection.Copy
    Workbooks.Open Filename:= _
        "C:\Users\GLB00156\Documents\SKY\SKY SAP Unicode\006_UAT\4_UAT_Gestao\DASHBOARD_SAP.xlsx" _
        , UpdateLinks:=3
    Sheets("PASTE_SAP_HERE").Select
    Range("B1").Select
    ActiveSheet.Paste
    Sheets("FROM SOLMAN").Select
    MsgBox "DashBoard atualizado"
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
End Sub

Sub SKY_UNIC_Clean_Pais()
    
    Cells.Select
        Selection.ClearContents
    
    ActiveSheet.DrawingObjects.Select
        Selection.Delete

End Sub
