Public Class TableCreation
    Dim ValidValueYesORNo = New String(,) {{"N", "No"}, {"Y", "Yes"}}
    Dim Replan = New String(,) {{" ", " "}, {"Replan", "Replan"}}
    Dim PROQty = New String(,) {{"N", "No"}, {"Y", "Yes"}}
    Dim Department = New String(,) {{"T", "Technical"}, {"M", "MIXING"}, {"G", "General"}, {"A", "ALL"}, {"E", "Estimation"}, {"S", "Subcontract"}}
    Dim ValidValueYesORNo1 = New String(,) {{"0", "No"}, {"1", "Yes"}}
    Dim WorkOrdersType = New String(,) {{"1", "Stock Order"}, {"2", "Sales Order"}, {"3", "Sales Return"}, {"4", "WIP Update"}}
    Dim AddOnSelectionYesORNo = New String(,) {{"Y", "Y"}, {"N", "N"}}
    Dim TimeMeasurement = New String(,) {{"S", "seconds"}, {"M", "Minute"}, {"H", "Hour"}, {"D", "Day"}}
    Dim PType = New String(,) {{"I", "Inhouse"}, {"S", "Subcontract"}, {"B", "Both"}}
    Dim Depend = New String(,) {{"MC", "Machine"}, {"MN", "Man"}, {"MM", "Machine\Man"}}
    Dim ToolsType = New String(,) {{"P", "Resharper"}, {"C", "Unresharper"}}
    Dim ProType = New String(,) {{"RG", "Reqular"}, {"RW", "Rework"}}
    Dim Status = New String(,) {{"O", "Open"}, {"C", "Close"}, {"H", "Hold"}}
    Dim ModeofTransport = New String(,) {{"S", "Sea"}, {"R", "Road"}, {"A", "Air"}}
    Dim Type = New String(,) {{"Pd", "Product"}, {"Pc", "Process"}}
    Dim PlantType = New String(,) {{"Prod", "Production"}, {"Prot", "Prototype"}, {"PreP", "Pre-Production"}}
    Dim NewType = New String(,) {{"IN", "In-house"}, {"GRN", "GRN"}, {"SC", "SubContract"}, {"FG", "FinishedGoods"}}
    Dim Inspectiontype = New String(,) {{"Acc", "Accepted"}, {"Rej", "Rejected"}, {"Rew", "Rework"}}
    Dim Status1 = New String(,) {{"O", "Open"}, {"C", "Close"}}
    Dim Priority = New String(,) {{"L", "Low"}, {"M", "Medium"}, {"C", "Close"}}
    Dim View = New String(,) {{"D", "Daily"}, {"W", "Weekily"}, {"M", "Monthily"}}
    Dim Statu = New String(,) {{"O", "Open"}, {"C", "Close"}}
    Dim Carriage = New String(,) {{"R", "Road"}, {"A", "Air"}, {"S", "Sea"}}
    Dim PRStatus = New String(,) {{"W", "Waiting"}, {"U", "UnApproved"}, {"A", "Approve"}}
    Dim InvType = New String(,) {{"Invoice", "Invoice"}, {"Sample Invoice", "Sample Invoice"}}
    Dim RowTypes = New String(,) {{"1", "Material"}, {"2", "Operations"}}
    Dim Day = New String(,) {{"1", "Day1"}, {"2", "Day2"}, {"3", "Day3"}, {"4", "Day4"}, {"5", "Day5"}, {"6", "Day6"}, {"7", "Day7"}}
    Dim Inspectiontypes = New String(,) {{"In-house", "In-house"}, {"SubContract", "SubContract"}, {"GRN Type 1", "GRN Type 1"}, {"GRN Type 2", "GRN Type 2"}, {"GRN Type 3", "GRN Type 3"}, {"GRN type 4", "GRN Type 4"}, {"GRN Type 5", "GRN Type 5"}, {"Customer Return", "Customer Return"}, {"FinishedGoods", "FinishedGoods"}, {"Compounds", "Compounds"}}
    Dim Ncofirmpriority = New String(,) {{"M", "Medium"}, {"L", "Low"}, {"H", "High"}}
    Dim WasType = New String(,) {{"W", "Wastage"}, {"R", "Rejection"}}
    Dim Process = New String(,) {{"SUB", "SUB"}}
    Dim Year = New String(,) {{"2017", "2017"}}
    Sub New()
        Try
            oGFun.CreateUserFields("OITT", "ProcessType", "Process Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            oGFun.CreateUserFields("OWOR", "PPLNDocEntry", "Planning Docentry", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            oGFun.CreateUserFields("OWOR", "PPLNDocNum", "Planning DocNum", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            oGFun.CreateUserFields("OWOR", "PPLNLineNum", "Planning LineNumber", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            oGFun.CreateUserFields("WTR1", "DailyPlnNum", "Daily Planning Number", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            oGFun.CreateUserFields("WTR1", "RecptEntry", "Receipt DocEntry", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            oGFun.CreateUserFields("WTR1", "ProQty", "Transfered Qty", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Quantity)
            Me.POPlanning()
            Me.ProductionType()
            Me.MonthPlanning()
            Me.ResourceCharge()
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

#Region "      Production Planning   "
    Sub POPlanning()
        Try
            Me.PROPlanningHeader()
            Me.PROPlanningDetail()
            If Not oGFun.UDOExists("PPLN") Then
                Dim findAliasNDescription = New String(,) {{"DocNum", "Document Number"}}
                oGFun.RegisterUDO("PPLN", "Production Planning", SAPbobsCOM.BoUDOObjType.boud_Document, findAliasNDescription, "MIPL_OPPN", "MIPL_PPN1")
                findAliasNDescription = Nothing
            End If
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub PROPlanningHeader()
        Try
            oGFun.CreateTable("MIPL_OPPN", "Production Planning Header", SAPbobsCOM.BoUTBTableType.bott_Document)
            oGFun.CreateUserFields("@MIPL_OPPN", "StDate", "Start Date", SAPbobsCOM.BoFieldTypes.db_Date)
            oGFun.CreateUserFields("@MIPL_OPPN", "EndDate", "End Date", SAPbobsCOM.BoFieldTypes.db_Date)
            oGFun.CreateUserFields("@MIPL_OPPN", "PostDate", "Post Date", SAPbobsCOM.BoFieldTypes.db_Date)
            oGFun.CreateUserFields("@MIPL_OPPN", "Month", "Month", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            oGFun.CreateUserFields("@MIPL_OPPN", "Location", "Location", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            oGFun.CreateUserFieldsComboBox("@MIPL_OPPN", "CostCenter", "Cost Center", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "", WasType, "W")
            oGFun.CreateUserFields("@MIPL_OPPN", "POSeries", "Production Order Series", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            oGFun.CreateUserFields("@MIPL_OPPN", "PlanBy", "Planned By", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            oGFun.CreateUserFields("@MIPL_OPPN", "CusPONo", "Customer PO No", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            oGFun.CreateUserFields("@MIPL_OPPN", "ProType", "Process Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 20)
            oGFun.CreateUserFields("@MIPL_OPPN", "Remark", "Remark", SAPbobsCOM.BoFieldTypes.db_Alpha, 20)
            oGFun.CreateUserFields("@MIPL_OPPN", "CopyTo", "Copy To", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            oGFun.CreateUserFields("@MIPL_OPPN", "JobNo", "Job Number", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub PROPlanningDetail()
        Try
            oGFun.CreateTable("MIPL_PPN1", "Production Planning Details", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
            oGFun.CreateUserFields("@MIPL_PPN1", "PoNum", "Production Order", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            oGFun.CreateUserFields("@MIPL_PPN1", "ItemCode", "ItemCode", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            oGFun.CreateUserFields("@MIPL_PPN1", "ItemDesc", "ItemDesc", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            oGFun.CreateUserFields("@MIPL_PPN1", "HWhs", "Header WareHouse", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            oGFun.CreateUserFields("@MIPL_PPN1", "CWhs", "Child Warehouse", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            oGFun.CreateUserFields("@MIPL_PPN1", "MplQty", "Monthly PlaneQty", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Quantity)
            oGFun.CreateUserFields("@MIPL_PPN1", "PoPlaQty", "Planned Qty", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Quantity)
            oGFun.CreateUserFields("@MIPL_PPN1", "CPlaQty", "Completed Qty", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Quantity)
            oGFun.CreateUserFields("@MIPL_PPN1", "DCPlaQty", "Daily Planned Qty", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Quantity)
            oGFun.CreateUserFields("@MIPL_PPN1", "DCompQty", "Daily Completed Qty", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Quantity)
            oGFun.CreateUserFields("@MIPL_PPN1", "RPlaQty", "Rejection Qty", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Quantity)
            oGFun.CreateUserFields("@MIPL_PPN1", "PePlaQty", "Pending Qty", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Quantity)
            oGFun.CreateUserFields("@MIPL_PPN1", "PoStatus", "Po Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            oGFun.CreateUserFields("@MIPL_PPN1", "PoDocEntry", "Po DocEntry", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            oGFun.CreateUserFields("@MIPL_PPN1", "Remarks", "Remarks", SAPbobsCOM.BoFieldTypes.db_Alpha, 150)
            oGFun.CreateUserFieldsComboBox("@MIPL_PPN1", "Select", "Selected", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "", AddOnSelectionYesORNo)
            oGFun.CreateUserFields("@MIPL_PPN1", "Phantom", "Phantom", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
#End Region

#Region "      Monthly Production Planning   "
    Sub MonthPlanning()
        Try
            Me.MPlanningHeader()
            Me.MPlanningDetail()
            If Not oGFun.UDOExists("MPLN") Then
                Dim findAliasNDescription = New String(,) {{"DocNum", "Document Number"}}
                oGFun.RegisterUDO("MPLN", "Monthly Production Planning", SAPbobsCOM.BoUDOObjType.boud_Document, findAliasNDescription, "MIPL_OMPN", "MIPL_MPN1")
                findAliasNDescription = Nothing
            End If
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub MPlanningHeader()
        Try
            oGFun.CreateTable("MIPL_OMPN", "Monthly Planning Header", SAPbobsCOM.BoUTBTableType.bott_Document)
            oGFun.CreateUserFieldsComboBox("@MIPL_OMPN", "Year", "Year", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "", WasType, "W")
            oGFun.CreateUserFields("@MIPL_OMPN", "PostDate", "Post Date", SAPbobsCOM.BoFieldTypes.db_Date)
            oGFun.CreateUserFields("@MIPL_OMPN", "Month", "Month", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            oGFun.CreateUserFieldsComboBox("@MIPL_OMPN", "CostCenter", "Cost Center", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "", WasType, "W")
            oGFun.CreateUserFields("@MIPL_OMPN", "PlanBy", "Planned By", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            oGFun.CreateUserFields("@MIPL_OMPN", "ProType", "Process Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 20)
            oGFun.CreateUserFields("@MIPL_OMPN", "Remark", "Remark", SAPbobsCOM.BoFieldTypes.db_Alpha, 200)
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub MPlanningDetail()
        Try
            oGFun.CreateTable("MIPL_MPN1", "Monthly Planning Details", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
            oGFun.CreateUserFields("@MIPL_MPN1", "PoNum", "Production Order", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            oGFun.CreateUserFields("@MIPL_MPN1", "ItemCode", "ItemCode", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            oGFun.CreateUserFields("@MIPL_MPN1", "ItemDesc", "ItemDesc", SAPbobsCOM.BoFieldTypes.db_Alpha, 150)
            oGFun.CreateUserFields("@MIPL_MPN1", "HWhs", "Header WareHouse", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            oGFun.CreateUserFields("@MIPL_MPN1", "CWhs", "Child Warehouse", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            oGFun.CreateUserFields("@MIPL_MPN1", "MplQty", "Monthly PlaneQty", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Quantity)
            oGFun.CreateUserFields("@MIPL_MPN1", "PoPlaQty", "Planned Qty", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Quantity)
            oGFun.CreateUserFields("@MIPL_MPN1", "CPlaQty", "Completed Qty", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Quantity)
            oGFun.CreateUserFields("@MIPL_MPN1", "DCPlaQty", "Daily Planned Qty", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Quantity)
            oGFun.CreateUserFields("@MIPL_MPN1", "DCompQty", "Daily Completed Qty", SAPbobsCOM.BoFieldTypes.db_Float, 0, SAPbobsCOM.BoFldSubTypes.st_Quantity)
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
#End Region

#Region "Production Type"
    Sub ProductionType()
        Try
            oGFun.CreateTable("MIPL_PROTYPE", "Production Process Type", SAPbobsCOM.BoUTBTableType.bott_NoObject)
            oGFun.CreateUserFields("@MIPL_PROTYPE", "Series", "Series", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
        Catch ex As Exception

        End Try

    End Sub
#End Region

#Region "Resource Charges"
    Sub ResourceCharge()
        Try
            oGFun.CreateTable("RESOURCECHARGE", "Resource Charge", SAPbobsCOM.BoUTBTableType.bott_NoObject)
            oGFun.CreateUserFields("@RESOURCECHARGE", "ITEMCODE", "ITEM CODE", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            oGFun.CreateUserFields("@RESOURCECHARGE", "ITEMNAME", "ITEM NAME", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            oGFun.CreateUserFields("@RESOURCECHARGE", "PROPROCESS", "PRO PROCESS", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            oGFun.CreateUserFields("@RESOURCECHARGE", "WHSECODE", "Warehouse Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            oGFun.CreateUserFields("@RESOURCECHARGE", "WHSENAME", "Warehouse Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            oGFun.CreateUserFields("@RESOURCECHARGE", "RESGROUP", "Resource Group Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            oGFun.CreateUserFields("@RESOURCECHARGE", "RESGROUPCODE", "Resource Group Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            oGFun.CreateUserFields("@RESOURCECHARGE", "RESCHARGE", "Resource Charge", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
        Catch ex As Exception
        End Try
    End Sub
#End Region

End Class
