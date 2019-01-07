Attribute VB_Name = "Módulo1"
Sub CRIA_PACOTE_NAMERANGES()
Attribute CRIA_PACOTE_NAMERANGES.VB_ProcData.VB_Invoke_Func = " \n14"
'
' CRIA_PACOTE_NAMERANGES Macro
'

'
    Dim PastaPais As String
    Dim NomePlan As String
    
    
        PastaPais = InputBox("Pais de Origem e Modulo", "Qual o país de origem e módulo", "BR_FI")
        NomePlanTP = PastaPais & "_FROM_TEMPLATE"
        NomePlanOT = PastaPais & "_OBLIGATORY_TCODE"
        NomePlanOS = PastaPais & "_OBLIGATORY_SE38"
        RefrPlanTP = "='C:\UAT_SolMan\UAT_Cenarios por Pais\" & PastaPais & "\[UAT TCP_" & PastaPais & "_V3.xlsx]FROM TEMPLATE'!R1C3:R5000C12"
        RefrPlanOT = "='C:\UAT_SolMan\UAT_Cenarios por Pais\" & PastaPais & "\[UAT TCP_" & PastaPais & "_V3.xlsx]OBLIGATORY_TCODE'!R1C3:R5000C12"
        RefrPlanOS = "='C:\UAT_SolMan\UAT_Cenarios por Pais\" & PastaPais & "\[UAT TCP_" & PastaPais & "_V3.xlsx]OBLIGATORY_SE38'!R1C3:R5000C12"
    
        ActiveWorkbook.Names.Add Name:=NomePlanTP, RefersToR1C1:=RefrPlanTP
        ActiveWorkbook.Names.Add Name:=NomePlanOT, RefersToR1C1:=RefrPlanOT
        ActiveWorkbook.Names.Add Name:=NomePlanOS, RefersToR1C1:=RefrPlanOS
    
End Sub
