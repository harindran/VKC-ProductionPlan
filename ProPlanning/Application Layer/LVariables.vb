Imports SDKLib881

Module LVariables

#Region " ... General Purpose ..."

    Public v_RetVal, v_ErrCode As Long
    Public v_ErrMsg As String = ""
    Public addonName As String = "Production Planning"
    Public oCompany As SAPbobsCOM.Company
    'Attachment Option
    Public ShowFolderBrowserThread As Threading.Thread
    Public BankFileName As String
    Public boolModelForm As Boolean = False
    Public boolModelFormID As String = ""
    Public oGFun As New GFun(addonName)
    'Public HWKEY() As String = New String() {"Q0021813522", "A1027874471", "X1211807750", "Y1334940735", "N2092941383", "F0123559701", "Y0578472451", "A0335651095"}
#End Region

#Region " ... Common For Module ..."

#End Region

#Region " ... Common For Forms ..."
    'Daily Planning....
    Public ProductionPlanningFormId As String = "PPLN"
    Public ProductionPlanningXML As String = "ProductionPlanning.xml"
    Public oProductionPlan As New ProductionPlanning

    'Monthly Planning....
    Public MonthPlanningFormId As String = "MPLN"
    Public MonthPlanningXML As String = "MonthlyPlanning.xml"
    Public oProductionMthPlan As New MonthlyPlanning

    'Production Order....
    Public PROMenuId As String = "4369"
    Public PROTypeEx As String = "65211"
    Public oProductionOrder As New ProductionOrder
    'Receipt From Production 
    Public PRORecptMenuId As String = "4370"
    Public PRORecptTypeEx As String = "65214"
    Public oReceiptFromPro As New ReceiptFromProduction
#End Region

End Module