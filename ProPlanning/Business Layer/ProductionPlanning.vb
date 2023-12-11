Imports System.Drawing.Drawing2D
Imports System.IO
Imports SAPbouiCOM

Public Class ProductionPlanning
    Public frmProductionPlanning As SAPbouiCOM.Form
    Dim oDBDSHeader As SAPbouiCOM.DBDataSource
    Dim oDBDSDetail As SAPbouiCOM.DBDataSource
    Public oMatrix1 As SAPbouiCOM.Matrix
    Public UDOID As String = "PPLN"
    Dim oedit As SAPbouiCOM.EditText
    Dim FlageOp = False, FlageLb = False, FlageMc = False
    Dim Itemstrsql As String = ""
    Dim OperationType = ""
    Dim strsql As String = ""
    Dim ProgValue As Integer = 0.0
    Dim oDetailsHeaderInv As SAPbouiCOM.DBDataSource
    Dim oFilters As SAPbouiCOM.EventFilters
    Dim oFilter As SAPbouiCOM.EventFilter

    Sub LoadForm()
        Try
            frmProductionPlanning = oGFun.LoadScreenXML(ProductionPlanningXML, oGFun.enuResourceType.Embeded, UDOID)
            'oGFun.LoadXML(frmProductionPlanning, ProductionPlanningFormId, ProductionPlanningXML)
            ''frmProductionPlanning = oGFun.LoadScreenXML(ProductionPlanningXML, oGFun.enuResourceType.Embeded, ProductionPlanningFormId)
            'frmProductionPlanning = oApplication.Forms.Item(ProductionPlanningFormId)
            setReport(frmProductionPlanning.UniqueID) 'ProductionPlanningFormId
            oDBDSHeader = frmProductionPlanning.DataSources.DBDataSources.Item("@MIPL_OPPN")
            oDBDSDetail = frmProductionPlanning.DataSources.DBDataSources.Item("@MIPL_PPN1")
            oMatrix1 = frmProductionPlanning.Items.Item("Matrix1").Specific
            Me.DefineModesForFields()
            Me.InitForm()
        Catch ex As Exception
            oGFun.Msg("Load Form Method Failed:" & ex.Message)
        End Try
    End Sub
    Sub InitForm()
        Try
            frmProductionPlanning.Freeze(True)
            oGFun.LoadComboBoxSeries(frmProductionPlanning.Items.Item("c_Series").Specific, UDOID)
            oGFun.LoadDocumentDate(frmProductionPlanning.Items.Item("t_PoDate").Specific)
            oGFun.LoadDocumentDate(frmProductionPlanning.Items.Item("t_SDate").Specific)
            oGFun.LoadDocumentDate(frmProductionPlanning.Items.Item("t_EDate").Specific)
            oGFun.setComboBoxValue(frmProductionPlanning.Items.Item("lt_Type").Specific, "Select code,Name from [@MIPL_PROTYPE] ")
            'oGFun.setComboBoxValue(frmProductionPlanning.Items.Item("t_PSeries").Specific, "Select Series,SeriesName from NNM1 where objectCode='202' and Locked='N'")
            strsql = ""
            oedit = frmProductionPlanning.Items.Item("t_PoDate").Specific
            Dim PosDate As Date = Date.ParseExact(oedit.Value, "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo)
            'strsql = " Select Series,SeriesName from NNM1 where objectCode='202' and Locked='N' and indicator =(select top 1 indicator from OFPR where F_RefDate<= convert(date,getdate()) and T_RefDate>= convert (date,getdate())) "
            strsql = " Select Series,SeriesName from NNM1 where objectCode='202' and Locked='N' and indicator =(select top 1 indicator from OFPR where F_RefDate<= convert(date,'" & PosDate.ToString("yyyyMMdd") & "') and T_RefDate>= convert (date,'" & PosDate.ToString("yyyyMMdd") & "')) "
            oGFun.setComboBoxValue(frmProductionPlanning.Items.Item("t_PSeries").Specific, strsql)
            frmProductionPlanning.Items.Item("31").Specific.Caption = ""
            oMatrix1.Clear()
            oDBDSDetail.Clear()
            oGFun.SetNewLine(oMatrix1, oDBDSDetail)
            Dim oMat As SAPbouiCOM.IMatrix = oMatrix1
            oMat.CommonSetting.EnableArrowKey = True
            Me.AddMode()
        Catch ex As Exception
            oApplication.StatusBar.SetText("InitForm Method Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
            frmProductionPlanning.Freeze(False)
        End Try
    End Sub
    Sub DefineModesForFields()
        Try
            frmProductionPlanning.Items.Item("t_DocNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            frmProductionPlanning.Items.Item("lt_CostCe").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Find, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            frmProductionPlanning.Items.Item("lt_CostCe").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Add, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            frmProductionPlanning.Items.Item("lt_CostCe").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            frmProductionPlanning.Items.Item("lt_Type").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Find, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            frmProductionPlanning.Items.Item("lt_Type").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Add, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            frmProductionPlanning.Items.Item("lt_Type").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            frmProductionPlanning.Items.Item("t_Month").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Find, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            frmProductionPlanning.Items.Item("t_Month").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Add, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            frmProductionPlanning.Items.Item("t_Month").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            frmProductionPlanning.Items.Item("t_PlnBy").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Find, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            frmProductionPlanning.Items.Item("t_PlnBy").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Add, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            frmProductionPlanning.Items.Item("t_PlnBy").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            frmProductionPlanning.Items.Item("c_Series").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Find, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            frmProductionPlanning.Items.Item("c_Series").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Add, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            frmProductionPlanning.Items.Item("c_Series").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            frmProductionPlanning.Items.Item("t_PoDate").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Find, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            frmProductionPlanning.Items.Item("t_PoDate").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Add, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            frmProductionPlanning.Items.Item("t_PoDate").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            frmProductionPlanning.Items.Item("t_SDate").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Find, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            frmProductionPlanning.Items.Item("t_SDate").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Add, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            frmProductionPlanning.Items.Item("t_SDate").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            frmProductionPlanning.Items.Item("t_PSeries").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Find, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            frmProductionPlanning.Items.Item("t_PSeries").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Add, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            frmProductionPlanning.Items.Item("t_PSeries").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            frmProductionPlanning.Items.Item("Matrix1").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Find, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            frmProductionPlanning.Items.Item("Matrix1").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Add, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            frmProductionPlanning.Items.Item("Matrix1").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            frmProductionPlanning.Items.Item("3").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Find, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            frmProductionPlanning.Items.Item("3").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Add, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            frmProductionPlanning.Items.Item("3").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            frmProductionPlanning.Items.Item("CmbSt").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Find, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            frmProductionPlanning.Items.Item("CmbSt").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Add, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            frmProductionPlanning.Items.Item("CmbSt").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            frmProductionPlanning.Items.Item("12A").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Visible, SAPbouiCOM.BoAutoFormMode.afm_Add, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            frmProductionPlanning.Items.Item("12A").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            Me.AddMode()
        Catch ex As Exception
            oApplication.StatusBar.SetText("DefineModesForFields Method Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
        End Try
    End Sub
    Function ValidateAll() As Boolean
        Try
            Dim RowCount = 0
            Dim ResCode(20) As String
            Dim rset As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            'rset.DoQuery("Select Distinct U_RESGROUPCODE From [@RESOURCECHARGE] with(nolock) where isnull(Code,0) <> '0'")
            oApplication.StatusBar.SetText("Validating Daily Production Planning. Please wait...", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            rset.DoQuery("Select Distinct U_RESGROUPCODE From [@RESOURCECHARGE] where U_RESGROUPCODE is not null ")
            For i = 0 To rset.RecordCount - 1
                ResCode(i) = rset.Fields.Item("U_RESGROUPCODE").Value
                rset.MoveNext()
            Next
            'For i As Integer = 1 To oMatrix1.VisualRowCount - 1
            '    If oMatrix1.Columns.Item("ItemC").Cells.Item(i).Specific.value <> "" Then
            '        RowCount = RowCount + 1
            '        Exit For
            '    End If
            'Next
            For iRow As Integer = 0 To oDBDSDetail.Size - 1
                If oDBDSDetail.GetValue("U_ItemCode", iRow) <> "" Then
                    RowCount = RowCount + 1
                    Exit For
                End If
            Next
            If RowCount = 0 Then
                oGFun.MsgBox("Minimum one Line Item is Required.. ", "E")
                Return False
            End If
            If frmProductionPlanning.Items.Item("t_PSeries").Specific.value.ToString.Trim.Equals("") = True Then
                oGFun.MsgBox("Production Order Series Should Not Be Left Empty")
                Return False
            End If
            If frmProductionPlanning.Items.Item("lt_CostCe").Specific.value.ToString.Trim.Equals("") = True Then
                oGFun.MsgBox("Cost Center Should Not Be Left Empty")
                Return False
            End If
            strsql = ""
            strsql = " Select Count(Series) from NNM1 where objectCode='202' and Locked='N' and indicator =(select top 1 indicator from OFPR where F_RefDate<= convert(date,'" & frmProductionPlanning.Items.Item("t_PoDate").Specific.value & "') and T_RefDate>= convert (date,'" & frmProductionPlanning.Items.Item("t_PoDate").Specific.value & "')) and Series='" & frmProductionPlanning.Items.Item("t_PSeries").Specific.Value & "' "
            If CDbl(oGFun.getSingleValue(strsql)) = 0 Then
                oGFun.MsgBox("Poduction Order Series Mismatch with Order Date ")
                Return False
            End If
            For iRow As Integer = 0 To oDBDSDetail.Size - 1
                If oDBDSDetail.GetValue("U_ItemCode", iRow) <> "" Then
                    If CDbl(oDBDSDetail.GetValue("U_DCPlaQty", iRow)) <= 0 Then
                        oGFun.MsgBox("Daily Planned Qty Should be greater than 0 in LineNo : [" & iRow + 1 & "] ")
                        Return False
                    End If
                    If oGFun.getSingleValue("Select count (*) from OITT where U_ProcessType='" & oDBDSHeader.GetValue("U_ProType", 0) & "' and code ='" & oDBDSDetail.GetValue("U_ItemCode", iRow) & "'") = "0" Then
                        oGFun.MsgBox("[" & oDBDSDetail.GetValue("U_ItemCode", iRow) & "] This item is not belong to Process Type  in LineNo : [" & iRow + 1 & "] ")
                        Return False
                    End If
                    ''Validate Resource Charges
                    For j As Integer = 0 To ResCode.Count - 1
                        If oGFun.getSingleValue("Select Count(*) from OITT a inner join ITT1 b on a.code=b.Father inner join ORSC c on b.Code=c.ResCode where a.Code='" & oDBDSDetail.GetValue("U_ItemCode", iRow) & "' and c.ResGrpCod='" & ResCode(j) & "'") = 0 Then Continue For
                        If oGFun.getSingleValue("select Count(*) from [@RESOURCECHARGE] where U_itemCode='" & oDBDSDetail.GetValue("U_ItemCode", iRow) & "' and U_WHSECODE ='" & oDBDSDetail.GetValue("U_CWhs", iRow) & "' and U_RESCHARGE >'0'") = 0 Then
                            oGFun.MsgBox("Define a Resource Charges for ItemCode :[" & oDBDSDetail.GetValue("U_ItemCode", iRow + 1) & "] in WareHouse :[" & oDBDSDetail.GetValue("U_CWhs", iRow) & "] LineNo : [" & iRow + 1 & "] ")
                            Return False
                        End If
                    Next
                    frmProductionPlanning.Items.Item("31").Specific.caption = "Validating line No: " & CStr(iRow + 1)
                End If
            Next
            frmProductionPlanning.Items.Item("31").Specific.caption = ""
            oApplication.StatusBar.SetText("Validated Daily Production Planning Successfully...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            'For i As Integer = 1 To oMatrix1.VisualRowCount
            '    If oMatrix1.Columns.Item("ItemC").Cells.Item(i).Specific.value <> "" Then
            '        If oMatrix1.Columns.Item("DPlQty").Cells.Item(i).Specific.value <= 0 Then
            '            oGFun.MsgBox("Daily Planned Qty Should be greater than 0 in LineNo : [" & i & "] ")
            '            Return False
            '        End If
            '        If oGFun.getSingleValue("Select count (*) from OITT where U_ProcessType='" & oDBDSHeader.GetValue("U_ProType", 0) & "' and code ='" & oMatrix1.Columns.Item("ItemC").Cells.Item(i).Specific.value & "'") = "0" Then
            '            oGFun.MsgBox("[" & oMatrix1.Columns.Item("ItemC").Cells.Item(i).Specific.value & "] This item is not belong to Process Type  in LineNo : [" & i & "] ")
            '            ' oApplication.MessageBox("[" & oMatrix1.Columns.Item("ItemC").Cells.Item(i).Specific.value & "] This item is not belong to Process Type  in LineNo : [" & i & "] ")
            '            Return False
            '        End If
            '        ''Validate Resource Charges
            '        For j As Integer = 0 To ResCode.Count - 1
            '            If oGFun.getSingleValue("Select Count(*) from OITT a inner join ITT1 b on a.code=b.Father inner join ORSC c on b.Code=c.ResCode where a.Code='" & oMatrix1.Columns.Item("ItemC").Cells.Item(i).Specific.value & "' and c.ResGrpCod='" & ResCode(j) & "'") = 0 Then Continue For
            '            If oGFun.getSingleValue("select Count(*) from [@RESOURCECHARGE] where U_itemCode='" & oMatrix1.Columns.Item("ItemC").Cells.Item(i).Specific.value & "' and U_WHSECODE ='" & oMatrix1.Columns.Item("CWhs").Cells.Item(i).Specific.value & "' and U_RESCHARGE >'0'") = 0 Then
            '                oGFun.MsgBox("Define a Resource Charges for ItemCode :[" & oMatrix1.Columns.Item("ItemC").Cells.Item(i).Specific.value & "] in WareHouse :[" & oMatrix1.Columns.Item("CWhs").Cells.Item(i).Specific.value & "] LineNo : [" & i & "] ")
            '                Return False
            '            End If
            '        Next
            '        frmProductionPlanning.Items.Item("31").Specific.caption = "Validating line No: " & CStr(i)
            '    End If
            'Next
            'frmProductionPlanning.Items.Item("31").Specific.caption = ""
            'frmProductionPlanning.Items.Item("31").Specific.caption = ""
            'If frmProductionPlanning.Items.Item("t_PlnBy").Specific.value.ToString.Trim.Equals("") = True Then
            '    oGFun.MsgBox("Planned by Should Not Be Left Empty")
            '    Return False
            'End If

            ValidateAll = True
        Catch ex As Exception
            oApplication.StatusBar.SetText("Validate Function Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            ValidateAll = False
        Finally

        End Try
    End Function
    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            'If pVal.ItemUID = "" Then Exit Sub
            'If pVal.EventType <> BoEventTypes.et_VALIDATE Then Exit Sub
            frmProductionPlanning = oApplication.Forms.Item(FormUID)
            'If frmProductionPlanning.Visible = False Then Exit Sub
            Dim oDataTable As SAPbouiCOM.DataTable
            Dim oCFLE As SAPbouiCOM.ChooseFromListEvent

            'oMatrix1 = frmProductionPlanning.Items.Item("Matrix1").Specific
            'oDBDSHeader = frmProductionPlanning.DataSources.DBDataSources.Item("@MIPL_OPPN")
            'oDBDSDetail = frmProductionPlanning.DataSources.DBDataSources.Item("@MIPL_PPN1")

            'oApplication.StatusBar.SetText("Item Event:" & pVal.EventType.ToString & "-Inner Event: " & pVal.InnerEvent.ToString & "- Action Success " & pVal.ActionSuccess.ToString & "- ItemChanged " & pVal.ItemChanged.ToString & "-POP " & pVal.PopUpIndicator.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            If pVal.BeforeAction = True Then
                Select Case pVal.EventType
                    Case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE
                        If frmProductionPlanning.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                            oApplication.MessageBox("Click on update button")
                            BubbleEvent = False
                        End If
                    'Case BoEventTypes.et_KEY_DOWN, BoEventTypes.et_PICKER_CLICKED
                        'If frmProductionPlanning.Mode = BoFormMode.fm_FIND_MODE Or pVal.Row < 1 Then Exit Sub
                        'If oMatrix1.Columns.Item("ProEntry").Cells.Item(pVal.Row).Specific.value <> "" Then
                        '    'BubbleEvent = False
                        'End If
                    Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                        If pVal.ActionSuccess = True Then Exit Sub
                        'If pVal.InnerEvent = True Then Exit Sub
                        oCFLE = pVal
                        oDataTable = oCFLE.SelectedObjects
                        Dim rset As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Select Case oCFLE.ChooseFromListUID
                                'Case "cfl_ItmCode"
                                    'strsql = ""
                                    'oGFun.ChooseFromListFilteration(frmProductionPlanning, "cfl_ItmCode", "ItemCode", strsq         'strsql = "Select ItemCode from OITM T0 inner join OITT T1 on T0.ItemCode=T1.Code where T1.TreeType='P' and T1.U_ProcessType='" & oDBDSHeader.GetValue("U_ProType", 0) & "'"l)
                            Case "cfl_PhCode"
                                strsql = ""
                                oGFun.ChooseFromListFilteration(frmProductionPlanning, "cfl_PhCode", "Code", "Select ItemCode from OITM T0 inner join OITT T1 on T0.ItemCode=T1.Code where T1.TreeType='P' and T0.Phantom='Y'")
                        End Select
                    Case SAPbouiCOM.BoEventTypes.et_CLICK
                        If pVal.ItemUID = "1" And frmProductionPlanning.Mode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                            'RemoveLastrow(oMatrix1, "ItemC")
                        End If
                        If pVal.ItemUID = "3" And frmProductionPlanning.Mode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                            If oApplication.MessageBox("Do you want to create the Inventory Transfer?", 2, "Yes", "No") <> 1 Then BubbleEvent = False : Exit Sub
                        End If
                    Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED
                        Try
                            Select Case pVal.ColUID
                                Case "ProNum"
                                    Try
                                        BubbleEvent = False
                                        frmProductionPlanning.Freeze(True)
                                        oApplication.Menus.Item("4369").Activate()
                                        Dim frm As SAPbouiCOM.Form = oApplication.Forms.ActiveForm
                                        frm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                                        Dim PostDate As String = oGFun.getSingleValue("Select Format(PostDate,'yyyyMMdd') from OWOR where DocEntry='" & oMatrix1.Columns.Item("ProEntry").Cells.Item(pVal.Row).Specific.value & "'")
                                        frm.Items.Item("18").Specific.value = oMatrix1.Columns.Item("ProNum").Cells.Item(pVal.Row).Specific.value
                                        frm.Items.Item("24").Specific.value = PostDate
                                        frm.Items.Item("1").Click()
                                        frmProductionPlanning.Freeze(False)
                                    Catch ex As Exception
                                        oApplication.StatusBar.SetText("Link Pressed Event Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                    End Try
                            End Select

                        Catch ex As Exception

                        End Try
                    Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                        Select Case pVal.ItemUID
                            Case "2A"
                                FlageOp = True
                                If frmProductionPlanning.Items.Item("CmbSt").Specific.selected.value = "R" Then
                                    oApplication.MessageBox("Production orders created already")
                                    Exit Sub
                                End If
                                Dim Flag As Boolean = False
                                If frmProductionPlanning.Mode = BoFormMode.fm_ADD_MODE Then Exit Sub
                                If frmProductionPlanning.Items.Item("CmbSt").Specific.selected.value = "O" Then
                                    ' If oCompany.InTransaction = False Then oCompany.StartTransaction()
                                    'Flag = Me.StockPostingNew()
                                    'Dim oEdit As SAPbouiCOM.EditText
                                    oedit = frmProductionPlanning.Items.Item("t_PoDate").Specific
                                    Dim PostingDate As Date = Date.ParseExact(oedit.Value, "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo)

                                    Dim Status As String = oGFun.getSingleValue("Select 1 as Status from NNM1 where objectCode='202' and Locked='N' and indicator =(select top 1 indicator from OFPR where F_RefDate<= convert(date,'" & PostingDate & "') and T_RefDate>= convert (date,'" & PostingDate & "')) and Series='" & frmProductionPlanning.Items.Item("t_PSeries").Specific.Value & "'")
                                    If Status = "" Then BubbleEvent = False : oApplication.MessageBox("Selected Series is not Matching with the Posting Date of Daily Plan screen...") : Exit Sub
                                    'RemoveLastrow(oMatrix1, "ItemC")
                                    Flag = StockPostingNew_20230307() 'Me.StockPostingNew_22092022()

                                    If Flag = False Then
                                        ' If oCompany.InTransaction Then oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                                        ' RollbackPONumber(FormUID) 'added by Tamizh 16-Nov-2019
                                        oApplication.MessageBox("Some Production Orders are Not Created. Please Check")
                                    Else
                                        ' If oCompany.InTransaction Then oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                                        oApplication.MessageBox("Production Orders Created Successfully")
                                        strsql = "Update [@MIPL_OPPN] set Status='R' where DocEntry =" & oDBDSHeader.GetValue("DocEntry", 0) & ""
                                        Dim rs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                        rs.DoQuery(strsql)
                                        rs = Nothing
                                        frmProductionPlanning.Items.Item("CmbSt").Specific.select("R", SAPbouiCOM.BoSearchKey.psk_ByValue)
                                        'oApplication.Menus.Item("1304").Activate()
                                    End If

                                End If
                        End Select
                End Select
            Else
                Select Case pVal.EventType
                    Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                        If pVal.ActionSuccess = False Then Exit Sub
                        'If pVal.InnerEvent = True Then Exit Sub
                        oCFLE = pVal
                        oDataTable = oCFLE.SelectedObjects
                        If Not oDataTable Is Nothing And pVal.BeforeAction = False Then
                            frmProductionPlanning.Freeze(True)
                            Select Case oCFLE.ChooseFromListUID
                                Case "cfl_ItmCode"
                                    oMatrix1.FlushToDataSource()
                                    oDBDSDetail.SetValue("U_ItemCode", pVal.Row - 1, oDataTable.GetValue("Code", 0))
                                    oDBDSDetail.SetValue("U_ItemDesc", pVal.Row - 1, oDataTable.GetValue("Name", 0)) 'oGFun.getSingleValue("Select IsNull(ItemName,'') as ItemName  From OITM with(nolock) where ItemCode='" & oDataTable.GetValue("Code", 0) & "'")
                                    oMatrix1.LoadFromDataSourceEx()
                                    'oGFun.SetNewLine(oMatrix1, oDBDSDetail, pVal.Row, "ItemC")
                                Case "cfl_Emp"
                                    oDBDSHeader.SetValue("U_PlanBy", 0, oDataTable.GetValue("lastName", 0) + ", " + oDataTable.GetValue("firstName", 0))
                                Case "cfl_WhsCod"
                                    oMatrix1.FlushToDataSource()
                                    oDBDSDetail.SetValue("U_HWhs", pVal.Row - 1, oDataTable.GetValue("WhsCode", 0))
                                    oMatrix1.LoadFromDataSourceEx()
                                Case "cfl_CWhsCod"
                                    oMatrix1.FlushToDataSource()
                                    oDBDSDetail.SetValue("U_CWhs", pVal.Row - 1, oDataTable.GetValue("WhsCode", 0))
                                    oMatrix1.LoadFromDataSourceEx()
                                Case "cfl_PhCode"
                                    oMatrix1.FlushToDataSource()
                                    oDBDSDetail.SetValue("U_Phantom", pVal.Row - 1, oDataTable.GetValue("Code", 0))
                                    oMatrix1.LoadFromDataSourceEx()
                            End Select
                            frmProductionPlanning.Freeze(False)
                        End If

                    Case SAPbouiCOM.BoEventTypes.et_VALIDATE
                        Try
                            Select Case pVal.ColUID
                                Case "ItemC"
                                    oGFun.Matrix_Addrow(oMatrix1, "ItemC", "LineId")
                            End Select
                        Catch ex As Exception
                        End Try
                    Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                        Select Case pVal.ItemUID
                            Case "t_PoDate"
                                strsql = ""
                                frmProductionPlanning.Items.Item("t_SDate").Specific.String = frmProductionPlanning.Items.Item("t_PoDate").Specific.String
                                strsql = " Select Series,SeriesName from NNM1 where objectCode='PPLN' and Locked='N' and indicator =(select top 1 indicator from OFPR where F_RefDate<= convert(date,'" & frmProductionPlanning.Items.Item("t_PoDate").Specific.value & "') and T_RefDate>= convert (date,'" & frmProductionPlanning.Items.Item("t_PoDate").Specific.value & "')) "
                                oGFun.SetComboBoxValueRefresh(frmProductionPlanning.Items.Item("c_Series").Specific, strsql)
                                'oGFun.setComboBoxValue(frmProductionPlanning.Items.Item("c_Series").Specific, strsql)
                                frmProductionPlanning.Items.Item("c_Series").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
                                strsql = " Select Series,SeriesName from NNM1 where objectCode='202' and Locked='N' and indicator =(select top 1 indicator from OFPR where F_RefDate<= convert(date,'" & frmProductionPlanning.Items.Item("t_PoDate").Specific.value & "') and T_RefDate>= convert (date,'" & frmProductionPlanning.Items.Item("t_PoDate").Specific.value & "')) "
                                oGFun.SetComboBoxValueRefresh(frmProductionPlanning.Items.Item("t_PSeries").Specific, strsql)
                                oGFun.setComboBoxValue(frmProductionPlanning.Items.Item("t_PSeries").Specific, strsql)
                            Case "Matrix1"
                                If pVal.ColUID = "ItemC" And pVal.Row = 1 Then
                                    oMatrix1.AutoResizeColumns()
                                End If
                        End Select
                    Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                        Try
                            Select Case pVal.ItemUID
                                Case "c_Series"
                                    If frmProductionPlanning.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                        oGFun.setDocNum(frmProductionPlanning)
                                    End If
                                Case "Matrix1"
                                    Select Case pVal.ColUID
                                    End Select
                            End Select
                        Catch ex As Exception
                            oApplication.StatusBar.SetText("Combo Select Event Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        Finally
                        End Try

                    Case SAPbouiCOM.BoEventTypes.et_CLICK
                        Select Case pVal.ItemUID
                            Case "3"
                                If frmProductionPlanning.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE And oDBDSHeader.GetValue("Status", 0) <> "C" Then
                                    Try
                                        Dim rset As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                        Dim rset1 As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                        Dim Docentry As String = ""
                                        oGFun.Msg("Loading Data.....", "S", "W")
                                        For i = 1 To oMatrix1.RowCount
                                            If oMatrix1.Columns.Item("Chck").Cells.Item(i).Specific.checked = True And oMatrix1.Columns.Item("ProSt").Cells.Item(i).Specific.Value <> "C" And oMatrix1.Columns.Item("ProSt").Cells.Item(i).Specific.Value <> "L" Then
                                                Docentry = Docentry + "'" + oMatrix1.Columns.Item("ProEntry").Cells.Item(i).Specific.value + "',"
                                            End If
                                        Next
                                        Docentry = Docentry + "''"
                                        rset1.DoQuery("Select ItemCode,sum(PlannedQty) 'Quantity',wareHouse from WOR1 where DocEntry in (" & Docentry & ") and ItemType=4 and isnull(DocEntry,'') <> '' group by  ItemCode,wareHouse")
                                        frmProductionPlanning.Freeze(True)
                                        oApplication.ActivateMenuItem("3080")
                                        GC.Collect()
                                        Dim FrmInvTrns As SAPbouiCOM.Form = oApplication.Forms.ActiveForm
                                        Dim oMatrix As SAPbouiCOM.Matrix = FrmInvTrns.Items.Item("23").Specific
                                        oDetailsHeaderInv = FrmInvTrns.DataSources.DBDataSources.Item(0)
                                        Dim PlnDocentry = oDBDSHeader.GetValue("DocEntry", 0)
                                        Dim TowareHouse As String = ""
                                        For i = 0 To rset1.RecordCount - 1
                                            TowareHouse = rset1.Fields.Item("wareHouse").Value
                                            oMatrix.Columns.Item("1").Cells.Item(oMatrix.VisualRowCount).Specific.value = rset1.Fields.Item("ItemCode").Value
                                            oMatrix.Columns.Item("10").Cells.Item(oMatrix.VisualRowCount - 1).Specific.value = rset1.Fields.Item("Quantity").Value
                                            oMatrix.Columns.Item("5").Cells.Item(oMatrix.VisualRowCount - 1).Specific.value = rset1.Fields.Item("wareHouse").Value
                                            oMatrix.Columns.Item("U_DailyPlnNum").Cells.Item(oMatrix.VisualRowCount - 1).Specific.value = PlnDocentry
                                            oMatrix.Columns.Item("U_ProQty").Cells.Item(oMatrix.VisualRowCount - 1).Specific.value = oGFun.getSingleValue("select isnull(sum(Quantity),0) 'Quantity' from WTR1  inner join [@MIPL_OPPN] a on a.DocEntry=U_DailyPlnNum where  ItemCode='" & rset1.Fields.Item("ItemCode").Value & "' and a.DocEntry='" & PlnDocentry & "' group by U_DailyPlnNum,ItemCode")
                                            rset1.MoveNext()
                                        Next
                                        If TowareHouse <> "" Then FrmInvTrns.Items.Item("1470000101").Specific.value = TowareHouse
                                        frmProductionPlanning.Freeze(False)
                                    Catch ex As Exception
                                        frmProductionPlanning.Freeze(False)
                                    End Try
                                End If
                            Case "Matrix1"
                                Try
                                    frmProductionPlanning.Items.Item("Matrix1").Specific.Columns.Item("Chck").Editable = True
                                    If pVal.Row <> 0 And pVal.InnerEvent = False Then
                                        If oMatrix1.IsRowSelected(pVal.Row) = True Then
                                            oMatrix1.SelectRow(pVal.Row, False, False)
                                        Else
                                            oMatrix1.SelectRow(pVal.Row, True, False)
                                        End If
                                    End If
                                Catch ex As Exception
                                End Try
                        End Select
                    Case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK
                        Select Case pVal.ColUID
                            Case "Chck"
                                If pVal.Row = 0 Then
                                    'Dim Checked As Boolean = True
                                    'For i As Integer = 1 To oMatrix1.RowCount
                                    '    If oMatrix1.Columns.Item("Chck").Cells.Item(i).Specific.checked = True Then
                                    '        Checked = False
                                    '        Exit For
                                    '    End If
                                    'Next
                                    'For i As Integer = 1 To oMatrix1.RowCount
                                    '    If oMatrix1.Columns.Item("ItemC").Cells.Item(i).Specific.Value <> "" Then
                                    '        oMatrix1.Columns.Item("Chck").Cells.Item(i).Specific.checked = Checked
                                    '    End If
                                    'Next

                                    oDBDSDetail = frmProductionPlanning.DataSources.DBDataSources.Item("@MIPL_PPN1")
                                    Dim sel As String = IIf(oMatrix1.Columns.Item("Chck").Cells.Item(1).Specific.checked = True, "N", "Y")
                                    oMatrix1.FlushToDataSource()
                                    For rowNum As Integer = 0 To oDBDSDetail.Size - 1
                                        If oDBDSDetail.GetValue("U_ItemCode", rowNum) <> "" Then
                                            oDBDSDetail.SetValue("U_select", rowNum, sel)
                                        End If
                                    Next
                                    frmProductionPlanning.Freeze(True)
                                    oMatrix1.LoadFromDataSource()
                                    frmProductionPlanning.Freeze(False)
                                End If
                        End Select
                    Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                        Select Case pVal.ItemUID
                            Case "2A"
                                frmProductionPlanning.Update()
                                ''If oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                                ''    oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                                ''End If
                                'strsql = oDBDSHeader.GetValue("DocEntry", 0)
                                frmProductionPlanning.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                                'frmProductionPlanning.Items.Item("12A").Enabled = True
                                'frmProductionPlanning.Items.Item("12A").Specific.String = strsql ' oDBDSHeader.GetValue("DocEntry", 0)
                                'frmProductionPlanning.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                'frmProductionPlanning.Items.Item("12A").Enabled = False
                        End Select

                End Select
            End If
#Region "Old"

            'Select Case pVal.EventType
            '    Case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE
            '        If frmProductionPlanning.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
            '            oApplication.MessageBox("Click on update button")
            '            BubbleEvent = False
            '        End If
            '    Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
            '        Try
            '            Dim oDataTable As SAPbouiCOM.DataTable
            '            Dim oCFLE As SAPbouiCOM.ChooseFromListEvent = pVal
            '            oDataTable = oCFLE.SelectedObjects
            '            Dim rset As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            '            If pVal.BeforeAction Then
            '                Select Case oCFLE.ChooseFromListUID
            '                    'Case "cfl_ItmCode"
            '                        'strsql = ""
            '                        'oGFun.ChooseFromListFilteration(frmProductionPlanning, "cfl_ItmCode", "ItemCode", strsq         'strsql = "Select ItemCode from OITM T0 inner join OITT T1 on T0.ItemCode=T1.Code where T1.TreeType='P' and T1.U_ProcessType='" & oDBDSHeader.GetValue("U_ProType", 0) & "'"l)
            '                    Case "cfl_PhCode"
            '                        strsql = ""
            '                        oGFun.ChooseFromListFilteration(frmProductionPlanning, "cfl_PhCode", "Code", "Select ItemCode from OITM T0  inner join OITT T1  on T0.ItemCode=T1.Code where T1.TreeType='P' and T0.Phantom='Y'")
            '                End Select
            '            End If
            '            If Not oDataTable Is Nothing And pVal.BeforeAction = False Then
            '                Select Case oCFLE.ChooseFromListUID
            '                    Case "cfl_ItmCode"
            '                        oMatrix1.FlushToDataSource()
            '                        oDBDSDetail.SetValue("U_ItemCode", pVal.Row - 1, oDataTable.GetValue("Code", 0))
            '                        oDBDSDetail.SetValue("U_ItemDesc", pVal.Row - 1, oGFun.getSingleValue("Select IsNull(ItemName,'') as ItemName  From OITM with(nolock) where ItemCode='" & oDataTable.GetValue("Code", 0) & "'"))
            '                        oMatrix1.LoadFromDataSourceEx()
            '                        'oGFun.SetNewLine(oMatrix1, oDBDSDetail, pVal.Row, "ItemC")
            '                    Case "cfl_Emp"
            '                        oDBDSHeader.SetValue("U_PlanBy", 0, oDataTable.GetValue("lastName", 0) + ", " + oDataTable.GetValue("firstName", 0))
            '                    Case "cfl_WhsCod"
            '                        oMatrix1.FlushToDataSource()
            '                        oDBDSDetail.SetValue("U_HWhs", pVal.Row - 1, oDataTable.GetValue("WhsCode", 0))
            '                        oMatrix1.LoadFromDataSourceEx()
            '                    Case "cfl_CWhsCod"
            '                        oMatrix1.FlushToDataSource()
            '                        oDBDSDetail.SetValue("U_CWhs", pVal.Row - 1, oDataTable.GetValue("WhsCode", 0))
            '                        oMatrix1.LoadFromDataSourceEx()
            '                    Case "cfl_PhCode"
            '                        oMatrix1.FlushToDataSource()
            '                        oDBDSDetail.SetValue("U_Phantom", pVal.Row - 1, oDataTable.GetValue("Code", 0))
            '                        oMatrix1.LoadFromDataSourceEx()
            '                End Select
            '            End If
            '        Catch ex As Exception
            '            oApplication.StatusBar.SetText("Choose From List Event Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            '        Finally
            '        End Try
            '    Case SAPbouiCOM.BoEventTypes.et_VALIDATE
            '        Try
            '            If pVal.BeforeAction = True Then Exit Sub
            '            Select Case pVal.ColUID
            '                Case "ItemC"
            '                    oGFun.Matrix_Addrow(oMatrix1, "ItemC", "LineId")
            '            End Select

            '        Catch ex As Exception

            '        End Try
            '    '    Try
            '    '        Select Case pVal.ItemUID
            '    '        End Select
            '    '    Catch ex As Exception
            '    '        oApplication.StatusBar.SetText("Validate Event Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            '    '    Finally
            '    '    End Try

            '    'Case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS
            '    '    Select Case pVal.ItemUID
            '    '        Case "Matrix1"
            '    '            Select Case pVal.ColUID

            '    '            End Select
            '    '    End Select
            '    Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
            '        Try
            '            Select Case pVal.ItemUID
            '                'Case "Matrix1"
            '                'Select Case pVal.ColUID
            '                'Case "ItemC" ' As per imran words coding commented by chitra on 21st june 2022 11.50 AM
            '                'If pVal.BeforeAction = False And oMatrix1.Columns.Item("ItemC").Cells.Item(pVal.Row).Specific.value <> "" Then
            '                '    Dim MnthPQty As String = oGFun.getSingleValue("Select isnull(sum(U_DCplaQty),'0') 'MonthlyPlanQty' from [@MIPL_OMPN] t1 inner join [@MIPL_MPN1] t2 on t1.DocEntry=t2.DocEntry where t1.U_Month='" & oDBDSHeader.GetValue("U_Month", 0) & "' and t1.U_year= Year('" & oDBDSHeader.GetValue("U_PostDate", 0) & "') and t2.U_ItemCode ='" & oMatrix1.Columns.Item("ItemC").Cells.Item(pVal.Row).Specific.value & "' and t1.U_ProType='" & oDBDSHeader.GetValue("U_ProType", 0) & "' and t1.U_CostCenter ='" & oDBDSHeader.GetValue("U_CostCenter", 0) & "' and Status <> 'C'  ")
            '                '    MnthPQty = IIf(MnthPQty = "", 0, MnthPQty)
            '                '    Dim MnthCQty As String = oGFun.getSingleValue(" Select isnull(Sum(t2.U_DCompQty),0) 'MonthlyCompQty' from [@MIPL_OPPN] t1 inner join [@MIPL_PPN1] t2 on t1.DocEntry =t2.DocEntry  where t1.U_Month='" & oDBDSHeader.GetValue("U_Month", 0) & "' and year(t1.U_PostDate) =Year('" & oDBDSHeader.GetValue("U_PostDate", 0) & "') and t2.U_ItemCode ='" & oMatrix1.Columns.Item("ItemC").Cells.Item(pVal.Row).Specific.value & "' and t1.U_ProType ='" & oDBDSHeader.GetValue("U_ProType", 0) & "' and t1.U_CostCenter ='" & oDBDSHeader.GetValue("U_CostCenter", 0) & "' ")
            '                '    MnthCQty = IIf(MnthCQty = "", 0, MnthCQty)
            '                '    oMatrix1.Columns.Item("MPlQty").Cells.Item(pVal.Row).Specific.Value = CDbl(MnthPQty)
            '                '    oMatrix1.Columns.Item("MCplQty").Cells.Item(pVal.Row).Specific.Value = CDbl(MnthCQty)
            '                'End If
            '                'Case "DPlQty"
            '                'If pVal.BeforeAction = False And frmProductionPlanning.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
            '                'oMatrix1.FlushToDataSource()
            '                'oDBDSDetail.SetValue("U_PePlaQty", pVal.Row - 1, oDBDSDetail.GetValue("U_DCPlaQty", pVal.Row - 1))
            '                'oMatrix1.LoadFromDataSource()
            '                'End If
            '                    'End Select
            '                    'Case "lt_Type"
            '                    '    Try
            '                    '        If pVal.BeforeAction = False And frmProductionPlanning.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
            '                    '            Dim combo As SAPbouiCOM.ComboBox
            '                    '            combo = frmProductionPlanning.Items.Item("t_PSeries").Specific
            '                    '            combo.Select(oGFun.getSingleValue("select top 1 a.seriesName from NNM1 a where SeriesName like '" & oDBDSHeader.GetValue("U_COSTCENTER", 0) & "" & "%" & "'  and a.remark='" & oDBDSHeader.GetValue("U_ProType", 0) & "' and objectCode='202' and Locked='N' "), SAPbouiCOM.BoSearchKey.psk_ByDescription)
            '                    '        End If
            '                    '    Catch ex As Exception
            '                    '    End Try
            '                Case "t_PoDate"
            '                    If pVal.BeforeAction = False Then
            '                        strsql = ""
            '                        strsql = " Select Series,SeriesName from NNM1 where objectCode='202' and Locked='N' and indicator =(select top 1 indicator from OFPR where F_RefDate<= convert(date,'" & frmProductionPlanning.Items.Item("t_PoDate").Specific.value & "') and T_RefDate>= convert (date,'" & frmProductionPlanning.Items.Item("t_PoDate").Specific.value & "')) "
            '                        oGFun.SetComboBoxValueRefresh(frmProductionPlanning.Items.Item("t_PSeries").Specific, strsql)
            '                        oGFun.setComboBoxValue(frmProductionPlanning.Items.Item("t_PSeries").Specific, strsql)
            '                    End If
            '            End Select
            '        Catch ex As Exception
            '            oApplication.StatusBar.SetText("Lost Focus Event Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            '        Finally
            '        End Try
            '    Case (SAPbouiCOM.BoEventTypes.et_COMBO_SELECT)
            '        Try
            '            Select Case pVal.ItemUID
            '                Case "c_Series"
            '                    If frmProductionPlanning.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE And pVal.BeforeAction = False Then
            '                        oGFun.setDocNum(frmProductionPlanning)
            '                    End If
            '                Case "Matrix1"
            '                    Select Case pVal.ColUID
            '                    End Select
            '            End Select
            '        Catch ex As Exception
            '            oApplication.StatusBar.SetText("Combo Select Event Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            '        Finally
            '        End Try

            '    Case SAPbouiCOM.BoEventTypes.et_CLICK
            '        Try
            '            Select Case pVal.ItemUID
            '                '----------------------------Uncommented for displaying error msg in box-----------
            '                'Case "1"
            '                '    If pVal.BeforeAction = True And (frmProductionPlanning.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE) Then
            '                '        If Me.ValidateAll() = False Then
            '                '            ' System.Media.SystemSounds.Asterisk.Play()
            '                '            oApplication.StatusBar.SetText("Validate Failed.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            '                '            BubbleEvent = False
            '                '            Exit Sub
            '                '        Else
            '                '        End If
            '                '    End If
            '                    '    '-----------------------------------------------------
            '                Case "3"
            '                    If pVal.BeforeAction = False And frmProductionPlanning.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE And oDBDSHeader.GetValue("Status", 0) <> "C" Then
            '                        Try
            '                            Dim rset As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            '                            Dim rset1 As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            '                            Dim Docentry As String = ""
            '                            oGFun.Msg("Loading Data.....", "S", "W")
            '                            For i = 1 To oMatrix1.RowCount
            '                                If oMatrix1.Columns.Item("Chck").Cells.Item(i).Specific.checked = True And oMatrix1.Columns.Item("ProSt").Cells.Item(i).Specific.Value <> "C" And oMatrix1.Columns.Item("ProSt").Cells.Item(i).Specific.Value <> "L" Then
            '                                    Docentry = Docentry + "'" + oMatrix1.Columns.Item("ProEntry").Cells.Item(i).Specific.value + "',"
            '                                End If
            '                            Next
            '                            Docentry = Docentry + "''"
            '                            rset1.DoQuery("Select ItemCode,sum(PlannedQty) 'Quantity',wareHouse from WOR1 where DocEntry in (" & Docentry & ") and ItemType=4 and isnull(DocEntry,'') <> '' group by  ItemCode,wareHouse")
            '                            frmProductionPlanning.Freeze(True)
            '                            oApplication.ActivateMenuItem("3080")
            '                            GC.Collect()
            '                            Dim FrmInvTrns As SAPbouiCOM.Form = oApplication.Forms.ActiveForm
            '                            Dim oMatrix As SAPbouiCOM.Matrix = FrmInvTrns.Items.Item("23").Specific
            '                            oDetailsHeaderInv = FrmInvTrns.DataSources.DBDataSources.Item(0)
            '                            Dim PlnDocentry = oDBDSHeader.GetValue("DocEntry", 0)
            '                            Dim TowareHouse As String = ""
            '                            For i = 0 To rset1.RecordCount - 1
            '                                TowareHouse = rset1.Fields.Item("wareHouse").Value
            '                                oMatrix.Columns.Item("1").Cells.Item(oMatrix.VisualRowCount).Specific.value = rset1.Fields.Item("ItemCode").Value
            '                                oMatrix.Columns.Item("10").Cells.Item(oMatrix.VisualRowCount - 1).Specific.value = rset1.Fields.Item("Quantity").Value
            '                                oMatrix.Columns.Item("5").Cells.Item(oMatrix.VisualRowCount - 1).Specific.value = rset1.Fields.Item("wareHouse").Value
            '                                oMatrix.Columns.Item("U_DailyPlnNum").Cells.Item(oMatrix.VisualRowCount - 1).Specific.value = PlnDocentry
            '                                oMatrix.Columns.Item("U_ProQty").Cells.Item(oMatrix.VisualRowCount - 1).Specific.value = oGFun.getSingleValue("select isnull(sum(Quantity),0) 'Quantity' from WTR1  inner join [@MIPL_OPPN] a on a.DocEntry=U_DailyPlnNum where  ItemCode='" & rset1.Fields.Item("ItemCode").Value & "' and a.DocEntry='" & PlnDocentry & "' group by U_DailyPlnNum,ItemCode")
            '                                rset1.MoveNext()
            '                            Next
            '                            If TowareHouse <> "" Then FrmInvTrns.Items.Item("1470000101").Specific.value = TowareHouse
            '                            frmProductionPlanning.Freeze(False)
            '                        Catch ex As Exception
            '                            frmProductionPlanning.Freeze(False)
            '                        End Try
            '                    End If
            '                Case "Matrix1"
            '                    Try
            '                        If pVal.BeforeAction = False Then
            '                            frmProductionPlanning.Items.Item("Matrix1").Specific.Columns.Item("Chck").Editable = True
            '                        End If
            '                    Catch ex As Exception
            '                    End Try
            '            End Select
            '        Catch ex As Exception
            '            oApplication.StatusBar.SetText("Click Event Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            '            'If oCompany.InTransaction Then oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            '            BubbleEvent = False
            '        Finally
            '        End Try
            '    Case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK
            '        Select Case pVal.ColUID
            '            Case "Chck"
            '                If pVal.BeforeAction = False And pVal.Row = 0 Then
            '                    Dim Checked As Boolean = True
            '                    For i As Integer = 1 To oMatrix1.RowCount
            '                        If oMatrix1.Columns.Item("Chck").Cells.Item(i).Specific.checked = True Then
            '                            Checked = False
            '                        End If
            '                    Next
            '                    For i As Integer = 1 To oMatrix1.RowCount
            '                        If oMatrix1.Columns.Item("ItemC").Cells.Item(i).Specific.Value <> "" Then
            '                            oMatrix1.Columns.Item("Chck").Cells.Item(i).Specific.checked = Checked
            '                        End If
            '                    Next
            '                End If
            '        End Select
            '    Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
            '        Try
            '            Select Case pVal.ItemUID
            '                'Case "1"
            '                '    If pVal.BeforeAction = False And frmProductionPlanning.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
            '                '        Me.InitForm()
            '                '        'ElseIf pVal.BeforeAction = True And frmProductionPlanning.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
            '                '        '    If oCompany.InTransaction = False Then oCompany.StartTransaction()
            '                '        '    FlageOp = False
            '                '        '    If Me.StockPosting = False Then
            '                '        '        If oCompany.InTransaction Then oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            '                '        '        BubbleEvent = False
            '                '        '    Else
            '                '        '        If oCompany.InTransaction Then oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            '                '        '    End If
            '                '    ElseIf pVal.BeforeAction = True And frmProductionPlanning.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
            '                '    End If
            '                Case "2A"

            '                    If pVal.BeforeAction = True Then ' changed to true 08-Jan-2021
            '                        FlageOp = True
            '                        If frmProductionPlanning.Items.Item("CmbSt").Specific.selected.value = "R" Then
            '                            oApplication.MessageBox("Production orders created already")
            '                            Exit Sub
            '                        End If
            '                        Dim Flag As Boolean = False
            '                        If frmProductionPlanning.Items.Item("CmbSt").Specific.selected.value = "O" Then
            '                            ' If oCompany.InTransaction = False Then oCompany.StartTransaction()
            '                            Flag = Me.StockPostingNew()
            '                            If Flag = False Then
            '                                ' If oCompany.InTransaction Then oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            '                                ' RollbackPONumber(FormUID) 'added by Tamizh 16-Nov-2019
            '                                oApplication.MessageBox("Some Production Orders are Not Created. Please Check")
            '                            Else
            '                                ' If oCompany.InTransaction Then oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
            '                                oApplication.MessageBox("Production Orders Created Successfully")
            '                                frmProductionPlanning.Items.Item("CmbSt").Specific.select("R", SAPbouiCOM.BoSearchKey.psk_ByValue)
            '                            End If

            '                        End If
            '                    Else
            '                        frmProductionPlanning.Update()
            '                        'If oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
            '                        '    oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
            '                        'End If
            '                        frmProductionPlanning.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
            '                        frmProductionPlanning.Items.Item("12A").Specific.String = oDBDSHeader.GetValue("DocEntry", 0)
            '                        frmProductionPlanning.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)

            '                    End If



            '            End Select
            '        Catch ex As Exception
            '            oApplication.StatusBar.SetText("Item Pressed Event Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            '        Finally
            '        End Try
            '    Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED
            '        Try
            '            If pVal.BeforeAction = True Then
            '                Select Case pVal.ColUID
            '                    Case "ProNum"
            '                        Try
            '                            BubbleEvent = False
            '                            frmProductionPlanning.Freeze(True)
            '                            oApplication.Menus.Item("4369").Activate()
            '                            Dim frm As SAPbouiCOM.Form = oApplication.Forms.ActiveForm
            '                            frm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
            '                            frm.Items.Item("18").Specific.value = oMatrix1.Columns.Item("ProNum").Cells.Item(pVal.Row).Specific.value
            '                            frm.Items.Item("1").Click()
            '                            frmProductionPlanning.Freeze(False)
            '                        Catch ex As Exception
            '                            oApplication.StatusBar.SetText("Link Pressed Event Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            '                        End Try
            '                End Select
            '            End If
            '        Catch ex As Exception

            '        End Try
            'End Select
#End Region
        Catch ex As Exception
            frmProductionPlanning.Freeze(False)
            'oApplication.StatusBar.SetText("Item Event Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
            GC.Collect()
            GC.WaitForPendingFinalizers()
        End Try
    End Sub
    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            frmProductionPlanning = oApplication.Forms.ActiveForm
            Select Case pVal.MenuUID
                Case "1284" 'Cancel
                    If pVal.BeforeAction = True Then If oApplication.MessageBox("Cancelling a document is irreversible. Document status will be changed to ""Canceled"". Do you want to continue?", 2, "Yes", "No") <> 1 Then BubbleEvent = False : Exit Sub
                Case "1286" 'Close
                    If pVal.BeforeAction = True Then If oApplication.MessageBox("Closing a document is irreversible. Document status will be changed to ""Closed"". Do you want to continue?", 2, "Yes", "No") <> 1 Then BubbleEvent = False : Exit Sub
                Case "1282"
                    If pVal.BeforeAction = False Then Me.InitForm()
                'Case "1288" ' Next Record
                '    If pVal.BeforeAction = False Then
                '        Me.UpdateMode()
                '    Else
                '        If frmProductionPlanning.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE And frmProductionPlanning.Items.Item("CmbSt").Specific.selected.value = "R" Then
                '            BubbleEvent = False
                '        End If
                '    End If
                'Case "1289" ' Previous Record
                '    If pVal.BeforeAction = False Then Me.UpdateMode()
                'Case "1290" ' First Record
                '    If pVal.BeforeAction = False Then Me.UpdateMode()
                'Case "1291" 'Last Record
                '    If pVal.BeforeAction = False Then Me.UpdateMode()
                Case "1293"
                    If pVal.BeforeAction Then
                        If frmProductionPlanning.Items.Item("CmbSt").Specific.selected.value = "R" Then
                            BubbleEvent = False
                        End If
                    Else
                        If frmProductionPlanning.Items.Item("CmbSt").Specific.selected.value = "R" Then
                            BubbleEvent = False
                        Else
                            If Trim(CStr(oDBDSDetail.GetValue("U_PoNum", 1))) = "" Then

                            End If
                            DeleteRow(oMatrix1, "@MIPL_PPN1")
                        End If
                    End If
                Case "1281"
                    If pVal.BeforeAction = False Then
                        If frmProductionPlanning.Items.Item("CmbSt").Specific.selected.value <> "O" Then
                            frmProductionPlanning.Items.Item("2A").Enabled = False
                        Else
                            frmProductionPlanning.Items.Item("2A").Enabled = True
                        End If
                    End If
            End Select
        Catch ex As Exception
            oApplication.StatusBar.SetText("Menu Event Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
        End Try
    End Sub
    Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            frmProductionPlanning = oApplication.Forms.Item(BusinessObjectInfo.FormUID)
            oMatrix1 = frmProductionPlanning.Items.Item("Matrix1").Specific
            Select Case BusinessObjectInfo.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                    Dim rsetUpadteIndent As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Try
                        If BusinessObjectInfo.BeforeAction Then
                            If Me.ValidateAll() = False Then ' commented for displaying the error in msg box 05-Jan-2021
                                'System.Media.SystemSounds.Asterisk.Play()
                                BubbleEvent = False
                                Exit Sub
                            End If
                        Else
                            Try
                                If frmProductionPlanning.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                    oGFun.setDocNum(frmProductionPlanning)
                                    ' commented by Tamizh to check with a Button

                                    '    Dim Flag As Boolean = False
                                    '    FlageOp = True
                                    '    If oCompany.InTransaction = False Then oCompany.StartTransaction()
                                    '    Flag = Me.StockPosting()
                                    '    If Flag = False Then
                                    '        If oCompany.InTransaction Then oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                                    '        BubbleEvent = False
                                    '    Else
                                    '        If oCompany.InTransaction Then oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                                    '    End If
                                End If
                            Catch ex As Exception
                                'BubbleEvent = False
                                'If oCompany.InTransaction Then oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                                'oApplication.StatusBar.SetText("Production Order Transaction is Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                'Return
                            End Try
                            oGFun.DeleteEmptyRowInFormDataEvent(oMatrix1, "ItemC", oDBDSDetail)
                        End If
                        'End If
                    Catch ex As Exception
                        'BubbleEvent = False
                    Finally
                    End Try
                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                    If BusinessObjectInfo.BeforeAction = True Then
                        If frmProductionPlanning.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE And frmProductionPlanning.Items.Item("CmbSt").Specific.selected.value = "R" Then
                            BubbleEvent = False
                        End If
                    End If
                    If BusinessObjectInfo.ActionSuccess = True Then
                        frmProductionPlanning.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                        UpdateMode()
                        oMatrix1.AutoResizeColumns()
                    End If
            End Select
        Catch ex As Exception
            'If oCompany.InTransaction Then oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            oApplication.StatusBar.SetText("Form Data Event Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
        End Try
    End Sub
    Private Sub RollbackPONumber(ByVal FormUID As String)
        frmProductionPlanning = oApplication.Forms.Item(FormUID)
        oMatrix1 = frmProductionPlanning.Items.Item("Matrix1").Specific
        For i As Integer = 1 To oMatrix1.RowCount
            If oMatrix1.Columns.Item("ProNum").Cells.Item(i).Specific.value <> "" Then
                oMatrix1.Columns.Item("ProNum").Cells.Item(i).Specific.value = ""
                oMatrix1.Columns.Item("ProEntry").Cells.Item(i).Specific.value = ""
            Else
                Exit For
            End If
        Next


    End Sub
    Private Function StockPostingNew() As Boolean
        'Dim oProgBar As SAPbouiCOM.ProgressBar
        Try
            'oProgBar = oApplication.StatusBar.CreateProgressBar("Progress Bar", oMatrix1.VisualRowCount, True)
            'oProgBar.Value = 0
            Dim ItemCode As String = ""
            'Dim oEdit As SAPbouiCOM.EditText
            oedit = frmProductionPlanning.Items.Item("t_SDate").Specific
            Dim OrderDate As Date = Date.ParseExact(oEdit.Value, "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo)
            oEdit = frmProductionPlanning.Items.Item("t_PoDate").Specific
            Dim PostingDate As Date = Date.ParseExact(oEdit.Value, "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo)
            For i As Integer = 1 To oMatrix1.VisualRowCount
                If oMatrix1.Columns.Item("ItemC").Cells.Item(i).Specific.value <> "" And Trim(CStr(oDBDSDetail.GetValue("U_PoNum", i - 1))) = "" Then
                    frmProductionPlanning.Items.Item("31").Specific.caption = "Creating production order for Line No: " & CStr(i)
                    Dim oProductionorder As SAPbobsCOM.ProductionOrders = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oProductionOrders)
                    Dim ErrCode
                    'Commented by Chitra on 21st june 2022 12.31 PM for fetching date from screen itself
                    'Dim PostingDate = oGFun.getSingleValue(" Select Convert(DateTime,'" & frmProductionPlanning.Items.Item("t_PoDate").Specific.Value & "') Dt ")
                    'Dim OrderDate = oGFun.getSingleValue(" Select Convert(DateTime,'" & frmProductionPlanning.Items.Item("t_SDate").Specific.Value & "') Dt ")
                    oProductionorder.PostingDate = OrderDate 'CDate(OrderDate)
                    oProductionorder.DueDate = PostingDate 'CDate(PostingDate)
                    oProductionorder.Series = frmProductionPlanning.Items.Item("t_PSeries").Specific.Value
                    oProductionorder.ProductionOrderType = SAPbobsCOM.BoProductionOrderTypeEnum.bopotStandard
                    oProductionorder.UserFields.Fields.Item("U_PPLNDocEntry").Value = oDBDSHeader.GetValue("DocEntry", 0)
                    oProductionorder.UserFields.Fields.Item("U_PPLNDocNum").Value = oDBDSHeader.GetValue("DocNum", 0)
                    oProductionorder.UserFields.Fields.Item("U_PPLNLineNum").Value = oDBDSDetail.GetValue("LineId", i - 1) 'oMatrix1.Columns.Item("LineId").Cells.Item(i).Specific.value
                    ItemCode = oDBDSDetail.GetValue("U_ItemCode", i - 1) ' oMatrix1.Columns.Item("ItemC").Cells.Item(i).Specific.value
                    oProductionorder.ItemNo = ItemCode 'oMatrix1.Columns.Item("ItemC").Cells.Item(i).Specific.value
                    oProductionorder.PlannedQuantity = oDBDSDetail.GetValue("U_DCPlaQty", i - 1) 'CDbl(oMatrix1.Columns.Item("DPlQty").Cells.Item(i).Specific.value)


                    oProductionorder.Warehouse = Trim(CStr(oDBDSDetail.GetValue("U_HWhs", i - 1)))  'oMatrix1.Columns.Item("BWhs").Cells.Item(i).Specific.value
                    Dim rs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    'rs.DoQuery(" Select Type,a.Code 'Code',Quantity,isnull(LineText,'') 'LineText',b.Qauntity,isnull(a.IssueMthd,'') 'IssueMthd',isnull(c.Phantom,'N') 'Phantom' from ITT1 a with(nolock) inner join OITT b with(nolock) on a.Father=b.Code Left Join OITM c with(nolock) on a.Code=c.ItemCode where Father ='" & oMatrix1.Columns.Item("ItemC").Cells.Item(i).Specific.value & "' ")
                    rs.DoQuery(" Select Type,a.Code 'Code',Quantity,isnull(LineText,'') 'LineText',b.Qauntity,isnull(a.IssueMthd,'') 'IssueMthd',isnull(c.Phantom,'N') 'Phantom' from ITT1 a with(nolock) inner join OITT b  on a.Father=b.Code Left Join OITM c  on a.Code=c.ItemCode where Father ='" & ItemCode & "' ")

                    For j As Integer = 0 To rs.RecordCount - 1
                        Dim ItemNo As String = rs.Fields.Item("Code").Value
                        Dim BaseQty As Double = CDbl(CDbl(rs.Fields.Item("Quantity").Value) / CDbl(rs.Fields.Item("Qauntity").Value))
                        Dim PlannedQty As Double = BaseQty * (CDbl(oDBDSDetail.GetValue("U_DCPlaQty", i - 1)))
                        If rs.Fields.Item("Code").Value <> "" And rs.Fields.Item("Phantom").Value = "N" Then
                            Select Case (rs.Fields.Item("Type").Value)
                                Case "4"
                                    With oProductionorder.Lines
                                        .ItemType = SAPbobsCOM.ProductionItemType.pit_Item
                                        .ItemNo = ItemNo
                                        .BaseQuantity = BaseQty
                                        .PlannedQuantity = PlannedQty
                                        .Warehouse = Trim(CStr(oDBDSDetail.GetValue("U_CWhs", i - 1))) ' oMatrix1.Columns.Item("CWhs").Cells.Item(i).Specific.value
                                        If rs.Fields.Item("IssueMthd").Value = "B" Then
                                            .ProductionOrderIssueType = SAPbobsCOM.BoIssueMethod.im_Backflush
                                        Else
                                            .ProductionOrderIssueType = SAPbobsCOM.BoIssueMethod.im_Manual
                                        End If
                                        .Add()
                                    End With
                                Case "290"
                                    With oProductionorder.Lines
                                        .ItemType = SAPbobsCOM.ProductionItemType.pit_Resource
                                        .ItemNo = ItemNo
                                        .BaseQuantity = BaseQty
                                        .PlannedQuantity = PlannedQty
                                        .Warehouse = Trim(CStr(oDBDSDetail.GetValue("U_CWhs", i - 1)))
                                        If rs.Fields.Item("IssueMthd").Value = "B" Then
                                            .ProductionOrderIssueType = SAPbobsCOM.BoIssueMethod.im_Backflush
                                        Else
                                            .ProductionOrderIssueType = SAPbobsCOM.BoIssueMethod.im_Manual
                                        End If
                                        .Add()
                                    End With
                                Case "-18"
                                    oProductionorder.Lines.LineText = rs.Fields.Item("LineText").Value
                                    oProductionorder.Lines.Add()
                            End Select
                        ElseIf rs.Fields.Item("Code").Value <> "" And rs.Fields.Item("Phantom").Value = "Y" And oMatrix1.Columns.Item("PhanItm").Cells.Item(i).Specific.value <> "" Then
                            Dim Phrs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            Phrs.DoQuery("Select Type,a.Code 'Code',Quantity,isnull(LineText,'') 'LineText',b.Qauntity,isnull(a.IssueMthd,'') 'IssueMthd',c.Phantom from ITT1 a with(nolock) inner join OITT b  on a.Father=b.Code Inner Join OITM c  on a.Code=c.ItemCode where Father ='" & oMatrix1.Columns.Item("PhanItm").Cells.Item(i).Specific.value & "' and c.Phantom='N'")
                            For Ph = 0 To Phrs.RecordCount - 1

                                ItemNo = Phrs.Fields.Item("Code").Value
                                BaseQty = CDbl(CDbl(rs.Fields.Item("Quantity").Value) * CDbl(Phrs.Fields.Item("Quantity").Value))
                                PlannedQty = BaseQty * (CDbl(oDBDSDetail.GetValue("U_DCPlaQty", i - 1)))
                                Select Case (Phrs.Fields.Item("Type").Value)
                                    Case "4"
                                        With oProductionorder.Lines
                                            .ItemType = SAPbobsCOM.ProductionItemType.pit_Item
                                            .ItemNo = ItemNo
                                            .BaseQuantity = BaseQty
                                            .PlannedQuantity = PlannedQty
                                            .Warehouse = Trim(CStr(oDBDSDetail.GetValue("U_CWhs", i - 1))) ' oMatrix1.Columns.Item("CWhs").Cells.Item(i).Specific.value
                                            If Phrs.Fields.Item("IssueMthd").Value = "B" Then
                                                .ProductionOrderIssueType = SAPbobsCOM.BoIssueMethod.im_Backflush
                                            Else
                                                .ProductionOrderIssueType = SAPbobsCOM.BoIssueMethod.im_Manual
                                            End If
                                            .Add()
                                        End With
                                    Case "290"
                                        With oProductionorder.Lines
                                            .ItemType = SAPbobsCOM.ProductionItemType.pit_Resource
                                            .ItemNo = ItemNo
                                            .BaseQuantity = BaseQty
                                            .PlannedQuantity = PlannedQty
                                            .Warehouse = Trim(CStr(oDBDSDetail.GetValue("U_CWhs", i - 1)))
                                            If Phrs.Fields.Item("IssueMthd").Value = "B" Then
                                                .ProductionOrderIssueType = SAPbobsCOM.BoIssueMethod.im_Backflush
                                            Else
                                                .ProductionOrderIssueType = SAPbobsCOM.BoIssueMethod.im_Manual
                                            End If
                                            .Add()
                                        End With
                                    Case "-18"
                                        oProductionorder.Lines.LineText = rs.Fields.Item("LineText").Value
                                        oProductionorder.Lines.Add()
                                End Select
                                Phrs.MoveNext()
                            Next
                        ElseIf rs.Fields.Item("Code").Value <> "" And rs.Fields.Item("Phantom").Value = "Y" And oMatrix1.Columns.Item("PhanItm").Cells.Item(i).Specific.value = "" Then
                            Dim Phrs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            Phrs.DoQuery("Select Type,a.Code 'Code',Quantity,isnull(LineText,'') 'LineText',b.Qauntity,isnull(a.IssueMthd,'') 'IssueMthd',c.Phantom from ITT1 a with(nolock) inner join OITT b on a.Father=b.Code Inner Join OITM c  on a.Code=c.ItemCode where Father ='" & rs.Fields.Item("Code").Value & "' and c.Phantom='N'")
                            For Ph = 0 To Phrs.RecordCount - 1

                                ItemNo = Phrs.Fields.Item("Code").Value
                                BaseQty = CDbl(CDbl(rs.Fields.Item("Quantity").Value) * CDbl(Phrs.Fields.Item("Quantity").Value))
                                PlannedQty = BaseQty * (CDbl(oDBDSDetail.GetValue("U_DCPlaQty", i - 1)))
                                Select Case (Phrs.Fields.Item("Type").Value)
                                    Case "4"
                                        With oProductionorder.Lines


                                            .ItemType = SAPbobsCOM.ProductionItemType.pit_Item
                                            .ItemNo = ItemNo
                                            .BaseQuantity = BaseQty
                                            .PlannedQuantity = PlannedQty
                                            .Warehouse = Trim(CStr(oDBDSDetail.GetValue("U_CWhs", i - 1)))
                                            If Phrs.Fields.Item("IssueMthd").Value = "B" Then
                                                .ProductionOrderIssueType = SAPbobsCOM.BoIssueMethod.im_Backflush
                                            Else
                                                .ProductionOrderIssueType = SAPbobsCOM.BoIssueMethod.im_Manual
                                            End If
                                            .Add()
                                        End With
                                    Case "290"
                                        With oProductionorder.Lines
                                            .ItemType = SAPbobsCOM.ProductionItemType.pit_Resource
                                            .ItemNo = ItemNo
                                            .BaseQuantity = BaseQty
                                            .PlannedQuantity = PlannedQty
                                            .Warehouse = Trim(CStr(oDBDSDetail.GetValue("U_CWhs", i - 1)))
                                            If Phrs.Fields.Item("IssueMthd").Value = "B" Then
                                                .ProductionOrderIssueType = SAPbobsCOM.BoIssueMethod.im_Backflush
                                            Else
                                                .ProductionOrderIssueType = SAPbobsCOM.BoIssueMethod.im_Manual
                                            End If
                                            .Add()
                                        End With


                                    Case "-18"
                                        oProductionorder.Lines.LineText = rs.Fields.Item("LineText").Value
                                        oProductionorder.Lines.Add()
                                End Select
                                Phrs.MoveNext()
                            Next
                        End If
                        rs.MoveNext()
                    Next
                    oProductionorder.Remarks = oGFun.getSingleValue("Select isnull(WhsName,'') as 'WhsName' From OWHS  where WhsCode ='" & oMatrix1.Columns.Item("CWhs").Cells.Item(i).Specific.value & "'") + "," + oMatrix1.Columns.Item("CWhs").Cells.Item(i).Specific.value + "," + oDBDSHeader.GetValue("DocNum", 0) + "," + oDBDSHeader.GetValue("U_JobNo", 0)
                    ErrCode = oProductionorder.Add()
                    If ErrCode <> 0 Then
                        '   oProgBar.Stop()
                        '  System.Runtime.InteropServices.Marshal.ReleaseComObject(oProgBar)
                        ' oProgBar = Nothing
                        '  GC.Collect()
                        oApplication.MessageBox(oCompany.GetLastErrorDescription & " Line No " & CStr(i)) ' Tamizh  05-Jan-2021
                        oApplication.SetStatusBarMessage(oCompany.GetLastErrorDescription)
                        'oForm.Items.Item("31").Specific.caption = "Creating production order for Line No: " & CStr(i)
                        Return False
                    Else
                        ''If FlageOp = True Then
                        ''    oProgBar.Text = "Production Order Created for [" & ItemCode & "]"
                        ''Else
                        ''    oProgBar.Text = "Validating Production Order : [" & ItemCode & "]...."
                        ''End If
                        'Dim sNewObjCode As String = ""
                        'oCompany.GetNewObjectCode(sNewObjCode)
                        'Dim str = CLng(sNewObjCode)
                        ''  oMatrix1.GetLineData(i)
                        '' oDBDSDetail.SetValue("U_PoStatus", i - 1, "P")
                        ''oDBDSDetail.SetValue("U_PoDocEntry", i - 1, CStr(str))
                        ''oDBDSDetail.SetValue("U_PoNum", i - 1, oGFun.getSingleValue("Select DocNum from OWOR where Docentry='" & str & "'"))
                        ''oMatrix1.SetLineData(i)
                        'U_PPLNDocEntry").Value = oDBDSHeader.GetValue("DocEntry", 0)
                        ''"U_PPLNLineNum").Value = oDBDSDetail.GetValue("LineId", i - 1)
                        Dim strSQL As String = "Select top 1 DocNum, DocEntry from OWOR with(nolock) where U_PPLNDocEntry=" & oDBDSHeader.GetValue("DocEntry", 0) & " and U_PPLNLineNum = '" & oDBDSDetail.GetValue("LineId", i - 1) & "' order by Docentry desc"
                        Dim objRS As SAPbobsCOM.Recordset
                        objRS = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        objRS.DoQuery(strSQL)
                        Dim ProdNum, ProdEntry As String
                        If Not objRS.EoF Then
                            ProdNum = objRS.Fields.Item("DocNum").Value
                            ProdEntry = objRS.Fields.Item("DocEntry").Value
                        End If
                        ' oMatrix1.Columns.Item("ProNum").Cells.Item(i).Specific.value = ProdNum
                        'oMatrix1.Columns.Item("ProEntry").Cells.Item(i).Specific.value = ProdEntry
                        strSQL = "Update [@MIPL_PPN1] set U_PoNum='" & ProdNum & "', U_PoDocEntry='" & ProdEntry & "' where Docentry =" & oDBDSHeader.GetValue("DocEntry", 0) & " and  LineId=" & oDBDSDetail.GetValue("LineId", i - 1) & ""
                        objRS.DoQuery(strSQL)
                        objRS = Nothing
                        frmProductionPlanning.Items.Item("31").Specific.caption = "Creating production order for Line No: " & CStr(i) & " - Created"

                        Dim ocombo As SAPbouiCOM.ComboBox
                        ocombo = oMatrix1.Columns.Item("ProSt").Cells.Item(i).Specific
                        ocombo.Select("P", SAPbouiCOM.BoSearchKey.psk_ByValue)
                    End If
                End If
                '  oProgBar.Value = ial
            Next

            'oProgBar.Stop()
            'System.Runtime.InteropServices.Marshal.ReleaseComObject(oProgBar)
            'oProgBar = Nothing
            'GC.Collect()
            Return True
        Catch ex As Exception
            oApplication.StatusBar.SetText("Posting Event Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            ' oApplication.MessageBox(ex.Message)
            'oProgBar.Stop()
            'System.Runtime.InteropServices.Marshal.ReleaseComObject(oProgBar)
            'oProgBar = Nothing
            'GC.Collect()
            Return False
        Finally
            GC.Collect()
            GC.WaitForPendingFinalizers()
        End Try
    End Function
    Sub DeleteRow(ByVal oMatrix As SAPbouiCOM.Matrix, ByVal TableName As String)
        Try
            Dim DBSource As SAPbouiCOM.DBDataSource
            frmProductionPlanning.Freeze(True)
            oMatrix.FlushToDataSource()
            DBSource = frmProductionPlanning.DataSources.DBDataSources.Item(TableName) '"@MIREJDET1"
            For i As Integer = 1 To oMatrix.VisualRowCount
                oMatrix.GetLineData(i)
                DBSource.Offset = i - 1
                DBSource.SetValue("LineId", DBSource.Offset, i)
                oMatrix.SetLineData(i)
                oMatrix.FlushToDataSource()
            Next
            DBSource.RemoveRecord(DBSource.Size - 1)
            oMatrix.LoadFromDataSource()
            frmProductionPlanning.Freeze(False)

        Catch ex As Exception
            frmProductionPlanning.Freeze(False)
            oApplication.StatusBar.SetText("DeleteRow  Method Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
        End Try
    End Sub
    Private Function StockPostingNew_20230307() As Boolean
        'Dim oProgBar As SAPbouiCOM.ProgressBar
        Try
            'oProgBar = oApplication.StatusBar.CreateProgressBar("Progress Bar", oMatrix1.VisualRowCount, True)
            'oProgBar.Value = 0
            Dim ItemCode As String = ""
            Dim ValidStatus As String = ""
            'Dim oEdit As SAPbouiCOM.EditText
            oDBDSHeader = frmProductionPlanning.DataSources.DBDataSources.Item("@MIPL_OPPN")
            oDBDSDetail = frmProductionPlanning.DataSources.DBDataSources.Item("@MIPL_PPN1")
            oEdit = frmProductionPlanning.Items.Item("t_SDate").Specific
            Dim OrderDate As Date = Date.ParseExact(oEdit.Value, "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo)
            oEdit = frmProductionPlanning.Items.Item("t_PoDate").Specific
            Dim PostingDate As Date = Date.ParseExact(oEdit.Value, "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo)

            For iRow As Integer = 0 To oDBDSDetail.Size - 1
                If Trim(oDBDSDetail.GetValue("U_ItemCode", iRow)) <> "" And Trim(CStr(oDBDSDetail.GetValue("U_PoNum", iRow))) = "" Then
                    frmProductionPlanning.Items.Item("31").Specific.caption = "Creating production order for Line No: " & CStr(iRow + 1)
                    Dim oProductionorder As SAPbobsCOM.ProductionOrders = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oProductionOrders)
                    Dim ErrCode
                    ItemCode = Trim(oDBDSDetail.GetValue("U_ItemCode", iRow)) ' oMatrix1.Columns.Item("ItemC").Cells.Item(i).Specific.value
                    oProductionorder.PostingDate = OrderDate 'CDate(OrderDate)
                    oProductionorder.DueDate = PostingDate 'CDate(PostingDate)
                    oProductionorder.Series = frmProductionPlanning.Items.Item("t_PSeries").Specific.Value
                    oProductionorder.ProductionOrderType = SAPbobsCOM.BoProductionOrderTypeEnum.bopotStandard
                    oProductionorder.UserFields.Fields.Item("U_PPLNDocEntry").Value = oDBDSHeader.GetValue("DocEntry", 0)
                    oProductionorder.UserFields.Fields.Item("U_PPLNDocNum").Value = oDBDSHeader.GetValue("DocNum", 0)
                    oProductionorder.UserFields.Fields.Item("U_PPLNLineNum").Value = oDBDSDetail.GetValue("LineId", iRow) 'oMatrix1.Columns.Item("LineId").Cells.Item(i).Specific.value
                    oProductionorder.ItemNo = ItemCode 'oMatrix1.Columns.Item("ItemC").Cells.Item(i).Specific.value
                    oProductionorder.PlannedQuantity = oDBDSDetail.GetValue("U_DCPlaQty", iRow) 'CDbl(oMatrix1.Columns.Item("DPlQty").Cells.Item(i).Specific.value)
                    oProductionorder.Warehouse = Trim(CStr(oDBDSDetail.GetValue("U_HWhs", iRow)))  'oMatrix1.Columns.Item("BWhs").Cells.Item(i).Specific.value
                    Dim rs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    'rs.DoQuery(" Select Type,a.Code 'Code',Quantity,isnull(LineText,'') 'LineText',b.Qauntity,isnull(a.IssueMthd,'') 'IssueMthd',isnull(c.Phantom,'N') 'Phantom' from ITT1 a with(nolock) inner join OITT b with(nolock) on a.Father=b.Code Left Join OITM c with(nolock) on a.Code=c.ItemCode where Father ='" & oMatrix1.Columns.Item("ItemC").Cells.Item(i).Specific.value & "' ")
                    strsql = " Select Type,a.Code 'Code',a.Father,Quantity,isnull(LineText,'') 'LineText',b.Qauntity,isnull(a.IssueMthd,'') 'IssueMthd',isnull(c.Phantom,'N') 'Phantom' from ITT1 a with(nolock) inner join OITT b  on a.Father=b.Code Left Join OITM c  on a.Code=c.ItemCode where a.Father ='" & ItemCode & "' "
                    rs.DoQuery(strsql)
                    If rs.RecordCount = 0 Then Continue For
                    'strsql = "Select distinct top 1 Case when a.Code in ((Select Code from ITT1 Where Father='" & ItemCode & "')) Then 'TRUE' else 'FALSE' end [Valid Status] from ITT1 a where a.Father ='" & ItemCode & "' Order by [Valid Status] "
                    'ValidStatus = oGFun.getSingleValue(strsql)
                    'If ValidStatus = "FALSE" Then oApplication.MessageBox("Mismatch in Line Items: " & " Line No " & CStr(iRow + 1)) : Return False
                    For j As Integer = 0 To rs.RecordCount - 1
                        Dim ItemNo As String = Trim(rs.Fields.Item("Code").Value)
                        If Trim(rs.Fields.Item("Father").Value) <> ItemCode Then
                            oApplication.MessageBox("Mismatch On Line No: " & CStr(iRow + 1) & vbCrLf & " Daily Plan ItemCode: " & ItemCode & vbCrLf & " BOM ItemCode: " & Trim(rs.Fields.Item("Father").Value))
                            oApplication.StatusBar.SetText("Mismatch On Line No:" & CStr(iRow + 1) & " Daily Plan ItemCode: " & ItemCode & " BOM ItemCode: " & Trim(rs.Fields.Item("Father").Value), SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            Return False
                        End If
                        Dim BaseQty As Double = CDbl(CDbl(rs.Fields.Item("Quantity").Value) / CDbl(rs.Fields.Item("Qauntity").Value))
                        Dim PlannedQty As Double = BaseQty * (CDbl(oDBDSDetail.GetValue("U_DCPlaQty", iRow)))
                        'oApplication.StatusBar.SetText("ItemNo: " & ItemNo, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        If Trim(rs.Fields.Item("Code").Value) <> "" And Trim(rs.Fields.Item("Phantom").Value) = "N" Then
                            Select Case (rs.Fields.Item("Type").Value)
                                Case "4"
                                    With oProductionorder.Lines
                                        .ItemType = SAPbobsCOM.ProductionItemType.pit_Item
                                        .ItemNo = ItemNo
                                        .BaseQuantity = BaseQty
                                        .PlannedQuantity = PlannedQty
                                        .Warehouse = Trim(CStr(oDBDSDetail.GetValue("U_CWhs", iRow))) ' oMatrix1.Columns.Item("CWhs").Cells.Item(i).Specific.value
                                        If rs.Fields.Item("IssueMthd").Value = "B" Then
                                            .ProductionOrderIssueType = SAPbobsCOM.BoIssueMethod.im_Backflush
                                        Else
                                            .ProductionOrderIssueType = SAPbobsCOM.BoIssueMethod.im_Manual
                                        End If
                                        .Add()
                                    End With
                                Case "290"
                                    With oProductionorder.Lines
                                        .ItemType = SAPbobsCOM.ProductionItemType.pit_Resource
                                        .ItemNo = ItemNo
                                        .BaseQuantity = BaseQty
                                        .PlannedQuantity = PlannedQty
                                        .Warehouse = Trim(CStr(oDBDSDetail.GetValue("U_CWhs", iRow)))
                                        If rs.Fields.Item("IssueMthd").Value = "B" Then
                                            .ProductionOrderIssueType = SAPbobsCOM.BoIssueMethod.im_Backflush
                                        Else
                                            .ProductionOrderIssueType = SAPbobsCOM.BoIssueMethod.im_Manual
                                        End If
                                        .Add()
                                    End With
                                Case "-18"
                                    oProductionorder.Lines.LineText = rs.Fields.Item("LineText").Value
                                    oProductionorder.Lines.Add()
                            End Select
                        ElseIf Trim(rs.Fields.Item("Code").Value) <> "" And Trim(rs.Fields.Item("Phantom").Value) = "Y" And Trim(oDBDSDetail.GetValue("U_Phantom", iRow)) <> "" Then
                            Dim Phrs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            Phrs.DoQuery("Select Type,a.Code 'Code',Quantity,isnull(LineText,'') 'LineText',b.Qauntity,isnull(a.IssueMthd,'') 'IssueMthd',c.Phantom from ITT1 a with(nolock) inner join OITT b  on a.Father=b.Code Inner Join OITM c  on a.Code=c.ItemCode where Father ='" & Trim(oDBDSDetail.GetValue("U_Phantom", iRow)) & "' and c.Phantom='N'")
                            For Ph = 0 To Phrs.RecordCount - 1
                                ItemNo = Trim(Phrs.Fields.Item("Code").Value)
                                BaseQty = CDbl(CDbl(rs.Fields.Item("Quantity").Value) * CDbl(Phrs.Fields.Item("Quantity").Value))
                                PlannedQty = BaseQty * (CDbl(oDBDSDetail.GetValue("U_DCPlaQty", iRow)))
                                Select Case (Phrs.Fields.Item("Type").Value)
                                    Case "4"
                                        With oProductionorder.Lines
                                            .ItemType = SAPbobsCOM.ProductionItemType.pit_Item
                                            .ItemNo = ItemNo
                                            .BaseQuantity = BaseQty
                                            .PlannedQuantity = PlannedQty
                                            .Warehouse = Trim(CStr(oDBDSDetail.GetValue("U_CWhs", iRow))) ' oMatrix1.Columns.Item("CWhs").Cells.Item(i).Specific.value
                                            If Phrs.Fields.Item("IssueMthd").Value = "B" Then
                                                .ProductionOrderIssueType = SAPbobsCOM.BoIssueMethod.im_Backflush
                                            Else
                                                .ProductionOrderIssueType = SAPbobsCOM.BoIssueMethod.im_Manual
                                            End If
                                            .Add()
                                        End With
                                    Case "290"
                                        With oProductionorder.Lines
                                            .ItemType = SAPbobsCOM.ProductionItemType.pit_Resource
                                            .ItemNo = ItemNo
                                            .BaseQuantity = BaseQty
                                            .PlannedQuantity = PlannedQty
                                            .Warehouse = Trim(CStr(oDBDSDetail.GetValue("U_CWhs", iRow)))
                                            If Phrs.Fields.Item("IssueMthd").Value = "B" Then
                                                .ProductionOrderIssueType = SAPbobsCOM.BoIssueMethod.im_Backflush
                                            Else
                                                .ProductionOrderIssueType = SAPbobsCOM.BoIssueMethod.im_Manual
                                            End If
                                            .Add()
                                        End With
                                    Case "-18"
                                        oProductionorder.Lines.LineText = Trim(rs.Fields.Item("LineText").Value)
                                        oProductionorder.Lines.Add()
                                End Select
                                Phrs.MoveNext()
                            Next
                        ElseIf Trim(rs.Fields.Item("Code").Value) <> "" And rs.Fields.Item("Phantom").Value = "Y" And oDBDSDetail.GetValue("U_Phantom", iRow) = "" Then
                            Dim Phrs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            Phrs.DoQuery("Select Type,a.Code 'Code',Quantity,isnull(LineText,'') 'LineText',b.Qauntity,isnull(a.IssueMthd,'') 'IssueMthd',c.Phantom from ITT1 a with(nolock) inner join OITT b on a.Father=b.Code Inner Join OITM c  on a.Code=c.ItemCode where Father ='" & Trim(rs.Fields.Item("Code").Value) & "' and c.Phantom='N'")
                            For Ph = 0 To Phrs.RecordCount - 1
                                ItemNo = Trim(Phrs.Fields.Item("Code").Value)
                                BaseQty = CDbl(CDbl(rs.Fields.Item("Quantity").Value) * CDbl(Phrs.Fields.Item("Quantity").Value))
                                PlannedQty = BaseQty * (CDbl(oDBDSDetail.GetValue("U_DCPlaQty", iRow)))
                                Select Case (Phrs.Fields.Item("Type").Value)
                                    Case "4"
                                        With oProductionorder.Lines
                                            .ItemType = SAPbobsCOM.ProductionItemType.pit_Item
                                            .ItemNo = ItemNo
                                            .BaseQuantity = BaseQty
                                            .PlannedQuantity = PlannedQty
                                            .Warehouse = Trim(CStr(oDBDSDetail.GetValue("U_CWhs", iRow)))
                                            If Phrs.Fields.Item("IssueMthd").Value = "B" Then
                                                .ProductionOrderIssueType = SAPbobsCOM.BoIssueMethod.im_Backflush
                                            Else
                                                .ProductionOrderIssueType = SAPbobsCOM.BoIssueMethod.im_Manual
                                            End If
                                            .Add()
                                        End With
                                    Case "290"
                                        With oProductionorder.Lines
                                            .ItemType = SAPbobsCOM.ProductionItemType.pit_Resource
                                            .ItemNo = ItemNo
                                            .BaseQuantity = BaseQty
                                            .PlannedQuantity = PlannedQty
                                            .Warehouse = Trim(CStr(oDBDSDetail.GetValue("U_CWhs", iRow)))
                                            If Phrs.Fields.Item("IssueMthd").Value = "B" Then
                                                .ProductionOrderIssueType = SAPbobsCOM.BoIssueMethod.im_Backflush
                                            Else
                                                .ProductionOrderIssueType = SAPbobsCOM.BoIssueMethod.im_Manual
                                            End If
                                            .Add()
                                        End With
                                    Case "-18"
                                        oProductionorder.Lines.LineText = Trim(rs.Fields.Item("LineText").Value)
                                        oProductionorder.Lines.Add()
                                End Select
                                Phrs.MoveNext()
                            Next
                        End If
                        rs.MoveNext()
                    Next

                    oProductionorder.Remarks = oGFun.getSingleValue("Select isnull(WhsName,'') as 'WhsName' From OWHS  where WhsCode ='" & oDBDSDetail.GetValue("U_CWhs", iRow) & "'") + "," + oDBDSDetail.GetValue("U_CWhs", iRow) + "," + oDBDSHeader.GetValue("DocNum", 0) + "," + oDBDSHeader.GetValue("U_JobNo", 0)
                    ErrCode = oProductionorder.Add()
                    If ErrCode <> 0 Then
                        '   oProgBar.Stop()
                        '  System.Runtime.InteropServices.Marshal.ReleaseComObject(oProgBar)
                        ' oProgBar = Nothing
                        '  GC.Collect()
                        oApplication.MessageBox("Error in Production Order: " & oCompany.GetLastErrorDescription() & " Line No " & CStr(iRow + 1)) ' Tamizh  05-Jan-2021
                        oApplication.StatusBar.SetText("Error in Production Order: " & oCompany.GetLastErrorDescription(), SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        'oApplication.SetStatusBarMessage("Error in Production Order: " & oCompany.GetLastErrorDescription())
                        'oForm.Items.Item("31").Specific.caption = "Creating production order for Line No: " & CStr(i)
                        Return False
                    Else

                        Dim strSQL As String = "Select top 1 DocNum, DocEntry from OWOR with(nolock) where U_PPLNDocEntry=" & oDBDSHeader.GetValue("DocEntry", 0) & " and U_PPLNLineNum = '" & oDBDSDetail.GetValue("LineId", iRow) & "' order by DocEntry desc"
                        Dim objRS As SAPbobsCOM.Recordset
                        objRS = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        objRS.DoQuery(strSQL)
                        Dim ProdNum, ProdEntry As String
                        If Not objRS.EoF Then
                            ProdNum = objRS.Fields.Item("DocNum").Value
                            ProdEntry = objRS.Fields.Item("DocEntry").Value
                        End If
                        ' oMatrix1.Columns.Item("ProNum").Cells.Item(i).Specific.value = ProdNum
                        'oMatrix1.Columns.Item("ProEntry").Cells.Item(i).Specific.value = ProdEntry
                        strSQL = "Update [@MIPL_PPN1] set U_PoNum='" & ProdNum & "', U_PoDocEntry='" & ProdEntry & "',U_PoStatus='P' where DocEntry =" & oDBDSHeader.GetValue("DocEntry", 0) & " and  LineId=" & oDBDSDetail.GetValue("LineId", iRow) & ""
                        objRS.DoQuery(strSQL)
                        objRS = Nothing
                        frmProductionPlanning.Items.Item("31").Specific.caption = "Production order Created for Line No: " & CStr(iRow + 1)
                        frmProductionPlanning.Items.Item("31").Specific.caption = ""
                        oApplication.StatusBar.SetText("Production order Created for Line No: " & CStr(iRow + 1), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        'Dim ocombo As SAPbouiCOM.ComboBox
                        'ocombo = oMatrix1.Columns.Item("ProSt").Cells.Item(iRow + 1).Specific
                        'ocombo.Select("P", SAPbouiCOM.BoSearchKey.psk_ByValue)
                    End If
                End If

            Next


            'oProgBar.Stop()
            'System.Runtime.InteropServices.Marshal.ReleaseComObject(oProgBar)
            'oProgBar = Nothing
            'GC.Collect()
            Return True
        Catch ex As Exception
            oApplication.StatusBar.SetText("Posting Event Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            ' oApplication.MessageBox(ex.Message)
            'oProgBar.Stop()
            'System.Runtime.InteropServices.Marshal.ReleaseComObject(oProgBar)
            'oProgBar = Nothing
            'GC.Collect()
            Return False
        Finally
            GC.Collect()
            GC.WaitForPendingFinalizers()
        End Try
    End Function
    Private Function StockPostingNew_22092022() As Boolean
        'Dim oProgBar As SAPbouiCOM.ProgressBar
        Try
            'oProgBar = oApplication.StatusBar.CreateProgressBar("Progress Bar", oMatrix1.VisualRowCount, True)
            'oProgBar.Value = 0
            Dim ItemCode As String = ""
            'Dim oEdit As SAPbouiCOM.EditText
            oDBDSHeader = frmProductionPlanning.DataSources.DBDataSources.Item("@MIPL_OPPN")
            oDBDSDetail = frmProductionPlanning.DataSources.DBDataSources.Item("@MIPL_PPN1")
            oEdit = frmProductionPlanning.Items.Item("t_SDate").Specific
            Dim OrderDate As Date = Date.ParseExact(oEdit.Value, "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo)
            oEdit = frmProductionPlanning.Items.Item("t_PoDate").Specific
            Dim PostingDate As Date = Date.ParseExact(oEdit.Value, "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo)

            For iRow As Integer = 0 To oDBDSDetail.Size - 1
                If oDBDSDetail.GetValue("U_ItemCode", iRow) <> "" And Trim(CStr(oDBDSDetail.GetValue("U_PoNum", iRow))) = "" Then
                    frmProductionPlanning.Items.Item("31").Specific.caption = "Creating production order for Line No: " & CStr(iRow + 1)
                    Dim oProductionorder As SAPbobsCOM.ProductionOrders = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oProductionOrders)
                    Dim ErrCode
                    oProductionorder.PostingDate = OrderDate 'CDate(OrderDate)
                    oProductionorder.DueDate = PostingDate 'CDate(PostingDate)
                    oProductionorder.Series = frmProductionPlanning.Items.Item("t_PSeries").Specific.Value
                    oProductionorder.ProductionOrderType = SAPbobsCOM.BoProductionOrderTypeEnum.bopotStandard
                    oProductionorder.UserFields.Fields.Item("U_PPLNDocEntry").Value = oDBDSHeader.GetValue("DocEntry", 0)
                    oProductionorder.UserFields.Fields.Item("U_PPLNDocNum").Value = oDBDSHeader.GetValue("DocNum", 0)
                    oProductionorder.UserFields.Fields.Item("U_PPLNLineNum").Value = oDBDSDetail.GetValue("LineId", iRow) 'oMatrix1.Columns.Item("LineId").Cells.Item(i).Specific.value
                    ItemCode = oDBDSDetail.GetValue("U_ItemCode", iRow) ' oMatrix1.Columns.Item("ItemC").Cells.Item(i).Specific.value
                    oProductionorder.ItemNo = ItemCode 'oMatrix1.Columns.Item("ItemC").Cells.Item(i).Specific.value
                    oProductionorder.PlannedQuantity = oDBDSDetail.GetValue("U_DCPlaQty", iRow) 'CDbl(oMatrix1.Columns.Item("DPlQty").Cells.Item(i).Specific.value)
                    oProductionorder.Warehouse = Trim(CStr(oDBDSDetail.GetValue("U_HWhs", iRow)))  'oMatrix1.Columns.Item("BWhs").Cells.Item(i).Specific.value
                    Dim rs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    'rs.DoQuery(" Select Type,a.Code 'Code',Quantity,isnull(LineText,'') 'LineText',b.Qauntity,isnull(a.IssueMthd,'') 'IssueMthd',isnull(c.Phantom,'N') 'Phantom' from ITT1 a with(nolock) inner join OITT b with(nolock) on a.Father=b.Code Left Join OITM c with(nolock) on a.Code=c.ItemCode where Father ='" & oMatrix1.Columns.Item("ItemC").Cells.Item(i).Specific.value & "' ")
                    rs.DoQuery(" Select Type,a.Code 'Code',Quantity,isnull(LineText,'') 'LineText',b.Qauntity,isnull(a.IssueMthd,'') 'IssueMthd',isnull(c.Phantom,'N') 'Phantom' from ITT1 a with(nolock) inner join OITT b  on a.Father=b.Code Left Join OITM c  on a.Code=c.ItemCode where Father ='" & ItemCode & "' ")

                    For j As Integer = 0 To rs.RecordCount - 1
                        Dim ItemNo As String = rs.Fields.Item("Code").Value
                        Dim BaseQty As Double = CDbl(CDbl(rs.Fields.Item("Quantity").Value) / CDbl(rs.Fields.Item("Qauntity").Value))
                        Dim PlannedQty As Double = BaseQty * (CDbl(oDBDSDetail.GetValue("U_DCPlaQty", iRow)))
                        If rs.Fields.Item("Code").Value <> "" And rs.Fields.Item("Phantom").Value = "N" Then
                            Select Case (rs.Fields.Item("Type").Value)
                                Case "4"
                                    With oProductionorder.Lines
                                        .ItemType = SAPbobsCOM.ProductionItemType.pit_Item
                                        .ItemNo = ItemNo
                                        .BaseQuantity = BaseQty
                                        .PlannedQuantity = PlannedQty
                                        .Warehouse = Trim(CStr(oDBDSDetail.GetValue("U_CWhs", iRow))) ' oMatrix1.Columns.Item("CWhs").Cells.Item(i).Specific.value
                                        If rs.Fields.Item("IssueMthd").Value = "B" Then
                                            .ProductionOrderIssueType = SAPbobsCOM.BoIssueMethod.im_Backflush
                                        Else
                                            .ProductionOrderIssueType = SAPbobsCOM.BoIssueMethod.im_Manual
                                        End If
                                        .Add()
                                    End With
                                Case "290"
                                    With oProductionorder.Lines
                                        .ItemType = SAPbobsCOM.ProductionItemType.pit_Resource
                                        .ItemNo = ItemNo
                                        .BaseQuantity = BaseQty
                                        .PlannedQuantity = PlannedQty
                                        .Warehouse = Trim(CStr(oDBDSDetail.GetValue("U_CWhs", iRow)))
                                        If rs.Fields.Item("IssueMthd").Value = "B" Then
                                            .ProductionOrderIssueType = SAPbobsCOM.BoIssueMethod.im_Backflush
                                        Else
                                            .ProductionOrderIssueType = SAPbobsCOM.BoIssueMethod.im_Manual
                                        End If
                                        .Add()
                                    End With
                                Case "-18"
                                    oProductionorder.Lines.LineText = rs.Fields.Item("LineText").Value
                                    oProductionorder.Lines.Add()
                            End Select
                        ElseIf rs.Fields.Item("Code").Value <> "" And rs.Fields.Item("Phantom").Value = "Y" And oDBDSDetail.GetValue("U_Phantom", iRow) <> "" Then
                            Dim Phrs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            Phrs.DoQuery("Select Type,a.Code 'Code',Quantity,isnull(LineText,'') 'LineText',b.Qauntity,isnull(a.IssueMthd,'') 'IssueMthd',c.Phantom from ITT1 a with(nolock) inner join OITT b  on a.Father=b.Code Inner Join OITM c  on a.Code=c.ItemCode where Father ='" & oDBDSDetail.GetValue("U_Phantom", iRow) & "' and c.Phantom='N'")
                            For Ph = 0 To Phrs.RecordCount - 1
                                ItemNo = Phrs.Fields.Item("Code").Value
                                BaseQty = CDbl(CDbl(rs.Fields.Item("Quantity").Value) * CDbl(Phrs.Fields.Item("Quantity").Value))
                                PlannedQty = BaseQty * (CDbl(oDBDSDetail.GetValue("U_DCPlaQty", iRow)))
                                Select Case (Phrs.Fields.Item("Type").Value)
                                    Case "4"
                                        With oProductionorder.Lines
                                            .ItemType = SAPbobsCOM.ProductionItemType.pit_Item
                                            .ItemNo = ItemNo
                                            .BaseQuantity = BaseQty
                                            .PlannedQuantity = PlannedQty
                                            .Warehouse = Trim(CStr(oDBDSDetail.GetValue("U_CWhs", iRow))) ' oMatrix1.Columns.Item("CWhs").Cells.Item(i).Specific.value
                                            If Phrs.Fields.Item("IssueMthd").Value = "B" Then
                                                .ProductionOrderIssueType = SAPbobsCOM.BoIssueMethod.im_Backflush
                                            Else
                                                .ProductionOrderIssueType = SAPbobsCOM.BoIssueMethod.im_Manual
                                            End If
                                            .Add()
                                        End With
                                    Case "290"
                                        With oProductionorder.Lines
                                            .ItemType = SAPbobsCOM.ProductionItemType.pit_Resource
                                            .ItemNo = ItemNo
                                            .BaseQuantity = BaseQty
                                            .PlannedQuantity = PlannedQty
                                            .Warehouse = Trim(CStr(oDBDSDetail.GetValue("U_CWhs", iRow)))
                                            If Phrs.Fields.Item("IssueMthd").Value = "B" Then
                                                .ProductionOrderIssueType = SAPbobsCOM.BoIssueMethod.im_Backflush
                                            Else
                                                .ProductionOrderIssueType = SAPbobsCOM.BoIssueMethod.im_Manual
                                            End If
                                            .Add()
                                        End With
                                    Case "-18"
                                        oProductionorder.Lines.LineText = rs.Fields.Item("LineText").Value
                                        oProductionorder.Lines.Add()
                                End Select
                                Phrs.MoveNext()
                            Next
                        ElseIf rs.Fields.Item("Code").Value <> "" And rs.Fields.Item("Phantom").Value = "Y" And oDBDSDetail.GetValue("U_Phantom", iRow) = "" Then
                            Dim Phrs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            Phrs.DoQuery("Select Type,a.Code 'Code',Quantity,isnull(LineText,'') 'LineText',b.Qauntity,isnull(a.IssueMthd,'') 'IssueMthd',c.Phantom from ITT1 a with(nolock) inner join OITT b on a.Father=b.Code Inner Join OITM c  on a.Code=c.ItemCode where Father ='" & rs.Fields.Item("Code").Value & "' and c.Phantom='N'")
                            For Ph = 0 To Phrs.RecordCount - 1
                                ItemNo = Phrs.Fields.Item("Code").Value
                                BaseQty = CDbl(CDbl(rs.Fields.Item("Quantity").Value) * CDbl(Phrs.Fields.Item("Quantity").Value))
                                PlannedQty = BaseQty * (CDbl(oDBDSDetail.GetValue("U_DCPlaQty", iRow)))
                                Select Case (Phrs.Fields.Item("Type").Value)
                                    Case "4"
                                        With oProductionorder.Lines
                                            .ItemType = SAPbobsCOM.ProductionItemType.pit_Item
                                            .ItemNo = ItemNo
                                            .BaseQuantity = BaseQty
                                            .PlannedQuantity = PlannedQty
                                            .Warehouse = Trim(CStr(oDBDSDetail.GetValue("U_CWhs", iRow)))
                                            If Phrs.Fields.Item("IssueMthd").Value = "B" Then
                                                .ProductionOrderIssueType = SAPbobsCOM.BoIssueMethod.im_Backflush
                                            Else
                                                .ProductionOrderIssueType = SAPbobsCOM.BoIssueMethod.im_Manual
                                            End If
                                            .Add()
                                        End With
                                    Case "290"
                                        With oProductionorder.Lines
                                            .ItemType = SAPbobsCOM.ProductionItemType.pit_Resource
                                            .ItemNo = ItemNo
                                            .BaseQuantity = BaseQty
                                            .PlannedQuantity = PlannedQty
                                            .Warehouse = Trim(CStr(oDBDSDetail.GetValue("U_CWhs", iRow)))
                                            If Phrs.Fields.Item("IssueMthd").Value = "B" Then
                                                .ProductionOrderIssueType = SAPbobsCOM.BoIssueMethod.im_Backflush
                                            Else
                                                .ProductionOrderIssueType = SAPbobsCOM.BoIssueMethod.im_Manual
                                            End If
                                            .Add()
                                        End With


                                    Case "-18"
                                        oProductionorder.Lines.LineText = rs.Fields.Item("LineText").Value
                                        oProductionorder.Lines.Add()
                                End Select
                                Phrs.MoveNext()
                            Next
                        End If
                        rs.MoveNext()
                    Next

                    oProductionorder.Remarks = oGFun.getSingleValue("Select isnull(WhsName,'') as 'WhsName' From OWHS  where WhsCode ='" & oDBDSDetail.GetValue("U_CWhs", iRow) & "'") + "," + oDBDSDetail.GetValue("U_CWhs", iRow) + "," + oDBDSHeader.GetValue("DocNum", 0) + "," + oDBDSHeader.GetValue("U_JobNo", 0)
                    ErrCode = oProductionorder.Add()
                    If ErrCode <> 0 Then
                        '   oProgBar.Stop()
                        '  System.Runtime.InteropServices.Marshal.ReleaseComObject(oProgBar)
                        ' oProgBar = Nothing
                        '  GC.Collect()
                        oApplication.MessageBox(oCompany.GetLastErrorDescription & " Line No " & CStr(iRow)) ' Tamizh  05-Jan-2021
                        oApplication.SetStatusBarMessage(oCompany.GetLastErrorDescription)
                        'oForm.Items.Item("31").Specific.caption = "Creating production order for Line No: " & CStr(i)
                        Return False
                    Else

                        Dim strSQL As String = "Select top 1 DocNum, DocEntry from OWOR with(nolock) where U_PPLNDocEntry=" & oDBDSHeader.GetValue("DocEntry", 0) & " and U_PPLNLineNum = '" & oDBDSDetail.GetValue("LineId", iRow) & "' order by Docentry desc"
                        Dim objRS As SAPbobsCOM.Recordset
                        objRS = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        objRS.DoQuery(strSQL)
                        Dim ProdNum, ProdEntry As String
                        If Not objRS.EoF Then
                            ProdNum = objRS.Fields.Item("DocNum").Value
                            ProdEntry = objRS.Fields.Item("DocEntry").Value
                        End If
                        ' oMatrix1.Columns.Item("ProNum").Cells.Item(i).Specific.value = ProdNum
                        'oMatrix1.Columns.Item("ProEntry").Cells.Item(i).Specific.value = ProdEntry
                        strSQL = "Update [@MIPL_PPN1] set U_PoNum='" & ProdNum & "', U_PoDocEntry='" & ProdEntry & "',U_PoStatus='P' where DocEntry =" & oDBDSHeader.GetValue("DocEntry", 0) & " and  LineId=" & oDBDSDetail.GetValue("LineId", iRow) & ""
                        objRS.DoQuery(strSQL)
                        objRS = Nothing
                        frmProductionPlanning.Items.Item("31").Specific.caption = "Creating production order for Line No: " & CStr(iRow + 1) & " - Created"
                        frmProductionPlanning.Items.Item("31").Specific.caption = ""
                        'Dim ocombo As SAPbouiCOM.ComboBox
                        'ocombo = oMatrix1.Columns.Item("ProSt").Cells.Item(iRow + 1).Specific
                        'ocombo.Select("P", SAPbouiCOM.BoSearchKey.psk_ByValue)
                    End If
                End If

            Next


            'oProgBar.Stop()
            'System.Runtime.InteropServices.Marshal.ReleaseComObject(oProgBar)
            'oProgBar = Nothing
            'GC.Collect()
            Return True
        Catch ex As Exception
            oApplication.StatusBar.SetText("Posting Event Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            ' oApplication.MessageBox(ex.Message)
            'oProgBar.Stop()
            'System.Runtime.InteropServices.Marshal.ReleaseComObject(oProgBar)
            'oProgBar = Nothing
            'GC.Collect()
            Return False
        Finally
            GC.Collect()
            GC.WaitForPendingFinalizers()
        End Try
    End Function
    Private Function StockPosting() As Boolean
        Dim oProgBar As SAPbouiCOM.ProgressBar
        Try
            'oProgBar = oApplication.StatusBar.CreateProgressBar("Progress Bar", oMatrix1.VisualRowCount, True)
            'oProgBar.Value = 0
            Dim ItemCode As String = ""
            For i As Integer = 1 To oMatrix1.VisualRowCount
                If oMatrix1.Columns.Item("ItemC").Cells.Item(i).Specific.value <> "" And Trim(CStr(oDBDSDetail.GetValue("U_PoNum", i - 1))) = "" Then
                    Dim oProductionorder As SAPbobsCOM.ProductionOrders = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oProductionOrders)
                    Dim ErrCode
                    Dim PostingDate = oGFun.getSingleValue(" Select Convert(DateTime,'" & frmProductionPlanning.Items.Item("t_PoDate").Specific.Value & "') Dt ")
                    Dim OrderDate = oGFun.getSingleValue(" Select Convert(DateTime,'" & frmProductionPlanning.Items.Item("t_SDate").Specific.Value & "') Dt ")
                    oProductionorder.PostingDate = CDate(OrderDate)
                    oProductionorder.DueDate = CDate(PostingDate)
                    oProductionorder.Series = frmProductionPlanning.Items.Item("t_PSeries").Specific.Value
                    oProductionorder.ProductionOrderType = SAPbobsCOM.BoProductionOrderTypeEnum.bopotStandard
                    oProductionorder.UserFields.Fields.Item("U_PPLNDocEntry").Value = oDBDSHeader.GetValue("DocEntry", 0)
                    oProductionorder.UserFields.Fields.Item("U_PPLNDocNum").Value = oDBDSHeader.GetValue("DocNum", 0)
                    oProductionorder.UserFields.Fields.Item("U_PPLNLineNum").Value = oDBDSDetail.GetValue("LineId", i - 1) 'oMatrix1.Columns.Item("LineId").Cells.Item(i).Specific.value
                    ItemCode = oDBDSDetail.GetValue("U_ItemCode", i - 1) ' oMatrix1.Columns.Item("ItemC").Cells.Item(i).Specific.value
                    oProductionorder.ItemNo = ItemCode 'oMatrix1.Columns.Item("ItemC").Cells.Item(i).Specific.value
                    oProductionorder.PlannedQuantity = oDBDSDetail.GetValue("U_DCPlaQty", i - 1) 'CDbl(oMatrix1.Columns.Item("DPlQty").Cells.Item(i).Specific.value)


                    oProductionorder.Warehouse = Trim(CStr(oDBDSDetail.GetValue("U_HWhs", i - 1)))  'oMatrix1.Columns.Item("BWhs").Cells.Item(i).Specific.value
                    Dim rs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    rs.DoQuery(" Select Type,a.Code 'Code',Quantity,isnull(LineText,'') 'LineText',b.Qauntity,isnull(a.IssueMthd,'') 'IssueMthd',isnull(c.Phantom,'N') 'Phantom' from ITT1(nolock) a inner join OITT(nolock) b on a.Father=b.Code Left Join OITM(nolock) c on a.Code=c.ItemCode where Father ='" & oMatrix1.Columns.Item("ItemC").Cells.Item(i).Specific.value & "' ")

                    For j As Integer = 0 To rs.RecordCount - 1
                        Dim ItemNo As String = rs.Fields.Item("Code").Value
                        Dim BaseQty As Double = CDbl(CDbl(rs.Fields.Item("Quantity").Value) / CDbl(rs.Fields.Item("Qauntity").Value))
                        Dim PlannedQty As Double = BaseQty * (CDbl(oDBDSDetail.GetValue("U_DCPlaQty", i - 1)))
                        If rs.Fields.Item("Code").Value <> "" And rs.Fields.Item("Phantom").Value = "N" Then

                            Select Case (rs.Fields.Item("Type").Value)
                                Case "4"
                                    With oProductionorder.Lines
                                        .ItemType = SAPbobsCOM.ProductionItemType.pit_Item
                                        .ItemNo = ItemNo
                                        .BaseQuantity = BaseQty
                                        .PlannedQuantity = PlannedQty
                                        .Warehouse = Trim(CStr(oDBDSDetail.GetValue("U_CWhs", i - 1))) ' oMatrix1.Columns.Item("CWhs").Cells.Item(i).Specific.value
                                        If rs.Fields.Item("IssueMthd").Value = "B" Then
                                            .ProductionOrderIssueType = SAPbobsCOM.BoIssueMethod.im_Backflush
                                        Else
                                            .ProductionOrderIssueType = SAPbobsCOM.BoIssueMethod.im_Manual
                                        End If
                                        .Add()
                                    End With
                                Case "290"
                                    With oProductionorder.Lines
                                        .ItemType = SAPbobsCOM.ProductionItemType.pit_Resource
                                        .ItemNo = ItemNo
                                        .BaseQuantity = BaseQty
                                        .PlannedQuantity = PlannedQty
                                        .Warehouse = Trim(CStr(oDBDSDetail.GetValue("U_CWhs", i - 1)))
                                        If rs.Fields.Item("IssueMthd").Value = "B" Then
                                            .ProductionOrderIssueType = SAPbobsCOM.BoIssueMethod.im_Backflush
                                        Else
                                            .ProductionOrderIssueType = SAPbobsCOM.BoIssueMethod.im_Manual
                                        End If
                                        .Add()
                                    End With
                                Case "-18"
                                    oProductionorder.Lines.LineText = rs.Fields.Item("LineText").Value
                                    oProductionorder.Lines.Add()
                            End Select
                        ElseIf rs.Fields.Item("Code").Value <> "" And rs.Fields.Item("Phantom").Value = "Y" And oMatrix1.Columns.Item("PhanItm").Cells.Item(i).Specific.value <> "" Then
                            Dim Phrs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            Phrs.DoQuery("Select Type,a.Code 'Code',Quantity,isnull(LineText,'') 'LineText',b.Qauntity,isnull(a.IssueMthd,'') 'IssueMthd',c.Phantom from ITT1(nolock) a inner join OITT(nolock) b on a.Father=b.Code Inner Join OITM(nolock) c on a.Code=c.ItemCode where Father ='" & oMatrix1.Columns.Item("PhanItm").Cells.Item(i).Specific.value & "' and c.Phantom='N'")
                            For Ph = 0 To Phrs.RecordCount - 1

                                ItemNo = Phrs.Fields.Item("Code").Value
                                BaseQty = CDbl(CDbl(rs.Fields.Item("Quantity").Value) * CDbl(Phrs.Fields.Item("Quantity").Value))
                                PlannedQty = BaseQty * (CDbl(oDBDSDetail.GetValue("U_DCPlaQty", i - 1)))
                                Select Case (Phrs.Fields.Item("Type").Value)
                                    Case "4"
                                        With oProductionorder.Lines
                                            .ItemType = SAPbobsCOM.ProductionItemType.pit_Item
                                            .ItemNo = ItemNo
                                            .BaseQuantity = BaseQty
                                            .PlannedQuantity = PlannedQty
                                            .Warehouse = Trim(CStr(oDBDSDetail.GetValue("U_CWhs", i - 1))) ' oMatrix1.Columns.Item("CWhs").Cells.Item(i).Specific.value
                                            If Phrs.Fields.Item("IssueMthd").Value = "B" Then
                                                .ProductionOrderIssueType = SAPbobsCOM.BoIssueMethod.im_Backflush
                                            Else
                                                .ProductionOrderIssueType = SAPbobsCOM.BoIssueMethod.im_Manual
                                            End If
                                            .Add()
                                        End With
                                    Case "290"
                                        With oProductionorder.Lines
                                            .ItemType = SAPbobsCOM.ProductionItemType.pit_Resource
                                            .ItemNo = ItemNo
                                            .BaseQuantity = BaseQty
                                            .PlannedQuantity = PlannedQty
                                            .Warehouse = Trim(CStr(oDBDSDetail.GetValue("U_CWhs", i - 1)))
                                            If Phrs.Fields.Item("IssueMthd").Value = "B" Then
                                                .ProductionOrderIssueType = SAPbobsCOM.BoIssueMethod.im_Backflush
                                            Else
                                                .ProductionOrderIssueType = SAPbobsCOM.BoIssueMethod.im_Manual
                                            End If
                                            .Add()
                                        End With
                                    Case "-18"
                                        oProductionorder.Lines.LineText = rs.Fields.Item("LineText").Value
                                        oProductionorder.Lines.Add()
                                End Select
                                Phrs.MoveNext()
                            Next
                        ElseIf rs.Fields.Item("Code").Value <> "" And rs.Fields.Item("Phantom").Value = "Y" And oMatrix1.Columns.Item("PhanItm").Cells.Item(i).Specific.value = "" Then
                            Dim Phrs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            Phrs.DoQuery("Select Type,a.Code 'Code',Quantity,isnull(LineText,'') 'LineText',b.Qauntity,isnull(a.IssueMthd,'') 'IssueMthd',c.Phantom from ITT1(nolock) a inner join OITT(nolock) b on a.Father=b.Code Inner Join OITM(nolock) c on a.Code=c.ItemCode where Father ='" & rs.Fields.Item("Code").Value & "' and c.Phantom='N'")
                            For Ph = 0 To Phrs.RecordCount - 1

                                ItemNo = Phrs.Fields.Item("Code").Value
                                BaseQty = CDbl(CDbl(rs.Fields.Item("Quantity").Value) * CDbl(Phrs.Fields.Item("Quantity").Value))
                                PlannedQty = BaseQty * (CDbl(oDBDSDetail.GetValue("U_DCPlaQty", i - 1)))
                                Select Case (Phrs.Fields.Item("Type").Value)
                                    Case "4"
                                        With oProductionorder.Lines


                                            .ItemType = SAPbobsCOM.ProductionItemType.pit_Item
                                            .ItemNo = ItemNo
                                            .BaseQuantity = BaseQty
                                            .PlannedQuantity = PlannedQty
                                            .Warehouse = Trim(CStr(oDBDSDetail.GetValue("U_CWhs", i - 1)))
                                            If Phrs.Fields.Item("IssueMthd").Value = "B" Then
                                                .ProductionOrderIssueType = SAPbobsCOM.BoIssueMethod.im_Backflush
                                            Else
                                                .ProductionOrderIssueType = SAPbobsCOM.BoIssueMethod.im_Manual
                                            End If
                                            .Add()
                                        End With
                                    Case "290"
                                        With oProductionorder.Lines
                                            .ItemType = SAPbobsCOM.ProductionItemType.pit_Resource
                                            .ItemNo = ItemNo
                                            .BaseQuantity = BaseQty
                                            .PlannedQuantity = PlannedQty
                                            .Warehouse = Trim(CStr(oDBDSDetail.GetValue("U_CWhs", i - 1)))
                                            If Phrs.Fields.Item("IssueMthd").Value = "B" Then
                                                .ProductionOrderIssueType = SAPbobsCOM.BoIssueMethod.im_Backflush
                                            Else
                                                .ProductionOrderIssueType = SAPbobsCOM.BoIssueMethod.im_Manual
                                            End If
                                            .Add()
                                        End With


                                    Case "-18"
                                        oProductionorder.Lines.LineText = rs.Fields.Item("LineText").Value
                                        oProductionorder.Lines.Add()
                                End Select
                                Phrs.MoveNext()
                            Next
                        End If
                        rs.MoveNext()
                    Next
                    oProductionorder.Remarks = oGFun.getSingleValue("Select isnull(WhsName,'') as 'WhsName' From OWHS(nolock) where WhsCode ='" & oMatrix1.Columns.Item("CWhs").Cells.Item(i).Specific.value & "'") + "," + oMatrix1.Columns.Item("CWhs").Cells.Item(i).Specific.value + "," + oDBDSHeader.GetValue("DocNum", 0) + "," + oDBDSHeader.GetValue("U_JobNo", 0)
                    ErrCode = oProductionorder.Add()
                    If ErrCode <> 0 Then
                        '   oProgBar.Stop()
                        '  System.Runtime.InteropServices.Marshal.ReleaseComObject(oProgBar)
                        ' oProgBar = Nothing
                        '  GC.Collect()
                        oApplication.MessageBox(oCompany.GetLastErrorDescription & " Line No " & CStr(i)) ' Tamizh  05-Jan-2021
                        oApplication.SetStatusBarMessage(oCompany.GetLastErrorDescription)

                        Return False
                    Else
                        'If FlageOp = True Then
                        '    oProgBar.Text = "Production Order Created for [" & ItemCode & "]"
                        'Else
                        '    oProgBar.Text = "Validating Production Order : [" & ItemCode & "]...."
                        'End If
                        Dim sNewObjCode As String = ""
                        oCompany.GetNewObjectCode(sNewObjCode)
                        Dim str = CLng(sNewObjCode)
                        '  oMatrix1.GetLineData(i)
                        ' oDBDSDetail.SetValue("U_PoStatus", i - 1, "P")
                        'oDBDSDetail.SetValue("U_PoDocEntry", i - 1, CStr(str))
                        'oDBDSDetail.SetValue("U_PoNum", i - 1, oGFun.getSingleValue("Select DocNum from OWOR where Docentry='" & str & "'"))
                        'oMatrix1.SetLineData(i)

                        oMatrix1.Columns.Item("ProNum").Cells.Item(i).Specific.value = oGFun.getSingleValue("Select top 1 DocNum from OWOR where Docentry='" & str & "' order by Docentry")
                        oMatrix1.Columns.Item("ProEntry").Cells.Item(i).Specific.value = str
                        Dim ocombo As SAPbouiCOM.ComboBox
                        ocombo = oMatrix1.Columns.Item("ProSt").Cells.Item(i).Specific
                        ocombo.Select("P", SAPbouiCOM.BoSearchKey.psk_ByValue)
                    End If
                End If
                '  oProgBar.Value = ial
            Next

            'oProgBar.Stop()
            'System.Runtime.InteropServices.Marshal.ReleaseComObject(oProgBar)
            'oProgBar = Nothing
            'GC.Collect()
            Return True
        Catch ex As Exception
            oApplication.StatusBar.SetText("Posting Event Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            ' oApplication.MessageBox(ex.Message)
            'oProgBar.Stop()
            'System.Runtime.InteropServices.Marshal.ReleaseComObject(oProgBar)
            'oProgBar = Nothing
            'GC.Collect()
            Return False
        Finally
        End Try
    End Function
    Private Function StockPostingOld() As Boolean
        'created by Tamizh on 27-Aug-2019 18:35pm
        Dim oProgBar As SAPbouiCOM.ProgressBar
        Try
            oProgBar = oApplication.StatusBar.CreateProgressBar("Progress Bar", oMatrix1.VisualRowCount, True)
            oProgBar.Value = 0
            Dim ItemCode As String = ""
            For i As Integer = 1 To oMatrix1.VisualRowCount
                If oMatrix1.Columns.Item("ItemC").Cells.Item(i).Specific.value <> "" Then
                    Dim oProductionorder As SAPbobsCOM.ProductionOrders = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oProductionOrders)
                    Dim ErrCode
                    Dim PostingDate = oGFun.getSingleValue(" Select Convert(DateTime,'" & frmProductionPlanning.Items.Item("t_PoDate").Specific.Value & "') Dt ")
                    Dim OrderDate = oGFun.getSingleValue(" Select Convert(DateTime,'" & frmProductionPlanning.Items.Item("t_SDate").Specific.Value & "') Dt ")
                    oProductionorder.PostingDate = CDate(OrderDate)
                    oProductionorder.DueDate = CDate(PostingDate)
                    oProductionorder.Series = frmProductionPlanning.Items.Item("t_PSeries").Specific.Value
                    oProductionorder.ProductionOrderType = SAPbobsCOM.BoProductionOrderTypeEnum.bopotStandard
                    oProductionorder.UserFields.Fields.Item("U_PPLNDocEntry").Value = oDBDSHeader.GetValue("DocEntry", 0)
                    oProductionorder.UserFields.Fields.Item("U_PPLNDocNum").Value = oDBDSHeader.GetValue("DocNum", 0)
                    oProductionorder.UserFields.Fields.Item("U_PPLNLineNum").Value = oMatrix1.Columns.Item("LineId").Cells.Item(i).Specific.value
                    ItemCode = oMatrix1.Columns.Item("ItemC").Cells.Item(i).Specific.value
                    oProductionorder.ItemNo = oMatrix1.Columns.Item("ItemC").Cells.Item(i).Specific.value
                    oProductionorder.PlannedQuantity = CDbl(oMatrix1.Columns.Item("DPlQty").Cells.Item(i).Specific.value)
                    oProductionorder.Warehouse = oMatrix1.Columns.Item("BWhs").Cells.Item(i).Specific.value
                    Dim rs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    rs.DoQuery(" Select Type,a.Code 'Code',Quantity,isnull(LineText,'') 'LineText',b.Qauntity,isnull(a.IssueMthd,'') 'IssueMthd',isnull(c.Phantom,'N') 'Phantom' from ITT1 a inner join OITT b on a.Father=b.Code Left Join OITM c on a.Code=c.ItemCode where Father ='" & oMatrix1.Columns.Item("ItemC").Cells.Item(i).Specific.value & "' ")
                    For j As Integer = 0 To rs.RecordCount - 1
                        If rs.Fields.Item("Code").Value <> "" And rs.Fields.Item("Phantom").Value = "N" Then
                            Select Case (rs.Fields.Item("Type").Value)
                                Case "4"
                                    oProductionorder.Lines.ItemType = SAPbobsCOM.ProductionItemType.pit_Item
                                    oProductionorder.Lines.ItemNo = rs.Fields.Item("Code").Value
                                    oProductionorder.Lines.BaseQuantity = CDbl(CDbl(rs.Fields.Item("Quantity").Value) / CDbl(rs.Fields.Item("Qauntity").Value))
                                    oProductionorder.Lines.PlannedQuantity = CDbl(CDbl(rs.Fields.Item("Quantity").Value) / CDbl(rs.Fields.Item("Qauntity").Value) * (CDbl(oMatrix1.Columns.Item("DPlQty").Cells.Item(i).Specific.value)))
                                    oProductionorder.Lines.Warehouse = oMatrix1.Columns.Item("CWhs").Cells.Item(i).Specific.value
                                    If rs.Fields.Item("IssueMthd").Value = "B" Then
                                        oProductionorder.Lines.ProductionOrderIssueType = SAPbobsCOM.BoIssueMethod.im_Backflush
                                    Else
                                        oProductionorder.Lines.ProductionOrderIssueType = SAPbobsCOM.BoIssueMethod.im_Manual
                                    End If
                                    oProductionorder.Lines.Add()
                                Case "290"
                                    oProductionorder.Lines.ItemType = SAPbobsCOM.ProductionItemType.pit_Resource
                                    oProductionorder.Lines.ItemNo = rs.Fields.Item("Code").Value
                                    oProductionorder.Lines.BaseQuantity = CDbl(rs.Fields.Item("Quantity").Value) / CDbl(rs.Fields.Item("Qauntity").Value)
                                    oProductionorder.Lines.PlannedQuantity = CDbl(CDbl(rs.Fields.Item("Quantity").Value) / CDbl(rs.Fields.Item("Qauntity").Value) * (CDbl(oMatrix1.Columns.Item("DPlQty").Cells.Item(i).Specific.value)))
                                    oProductionorder.Lines.Warehouse = oMatrix1.Columns.Item("CWhs").Cells.Item(i).Specific.value
                                    If rs.Fields.Item("IssueMthd").Value = "B" Then
                                        oProductionorder.Lines.ProductionOrderIssueType = SAPbobsCOM.BoIssueMethod.im_Backflush
                                    Else
                                        oProductionorder.Lines.ProductionOrderIssueType = SAPbobsCOM.BoIssueMethod.im_Manual
                                    End If
                                    oProductionorder.Lines.Add()
                                Case "-18"
                                    oProductionorder.Lines.LineText = rs.Fields.Item("LineText").Value
                                    oProductionorder.Lines.Add()
                            End Select
                        ElseIf rs.Fields.Item("Code").Value <> "" And rs.Fields.Item("Phantom").Value = "Y" And oMatrix1.Columns.Item("PhanItm").Cells.Item(i).Specific.value <> "" Then
                            Dim Phrs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            Phrs.DoQuery("Select Type,a.Code 'Code',Quantity,isnull(LineText,'') 'LineText',b.Qauntity,isnull(a.IssueMthd,'') 'IssueMthd',c.Phantom from ITT1 a inner join OITT b on a.Father=b.Code Inner Join OITM c on a.Code=c.ItemCode where Father ='" & oMatrix1.Columns.Item("PhanItm").Cells.Item(i).Specific.value & "' and c.Phantom='N'")
                            For Ph = 0 To Phrs.RecordCount - 1
                                Select Case (Phrs.Fields.Item("Type").Value)
                                    Case "4"
                                        oProductionorder.Lines.ItemType = SAPbobsCOM.ProductionItemType.pit_Item
                                        oProductionorder.Lines.ItemNo = Phrs.Fields.Item("Code").Value
                                        oProductionorder.Lines.BaseQuantity = CDbl(CDbl(rs.Fields.Item("Quantity").Value) * CDbl(Phrs.Fields.Item("Quantity").Value))
                                        oProductionorder.Lines.PlannedQuantity = CDbl(CDbl(rs.Fields.Item("Quantity").Value) * CDbl(Phrs.Fields.Item("Quantity").Value) * (CDbl(oMatrix1.Columns.Item("DPlQty").Cells.Item(i).Specific.value)))
                                        oProductionorder.Lines.Warehouse = oMatrix1.Columns.Item("CWhs").Cells.Item(i).Specific.value
                                        If Phrs.Fields.Item("IssueMthd").Value = "B" Then
                                            oProductionorder.Lines.ProductionOrderIssueType = SAPbobsCOM.BoIssueMethod.im_Backflush
                                        Else
                                            oProductionorder.Lines.ProductionOrderIssueType = SAPbobsCOM.BoIssueMethod.im_Manual
                                        End If
                                        oProductionorder.Lines.Add()
                                    Case "290"
                                        oProductionorder.Lines.ItemType = SAPbobsCOM.ProductionItemType.pit_Resource
                                        oProductionorder.Lines.ItemNo = Phrs.Fields.Item("Code").Value
                                        oProductionorder.Lines.BaseQuantity = CDbl(rs.Fields.Item("Quantity").Value) * CDbl(Phrs.Fields.Item("Quantity").Value)
                                        oProductionorder.Lines.PlannedQuantity = CDbl(CDbl(rs.Fields.Item("Quantity").Value) * CDbl(Phrs.Fields.Item("Quantity").Value) * (CDbl(oMatrix1.Columns.Item("DPlQty").Cells.Item(i).Specific.value)))
                                        oProductionorder.Lines.Warehouse = oMatrix1.Columns.Item("CWhs").Cells.Item(i).Specific.value
                                        If Phrs.Fields.Item("IssueMthd").Value = "B" Then
                                            oProductionorder.Lines.ProductionOrderIssueType = SAPbobsCOM.BoIssueMethod.im_Backflush
                                        Else
                                            oProductionorder.Lines.ProductionOrderIssueType = SAPbobsCOM.BoIssueMethod.im_Manual
                                        End If
                                        oProductionorder.Lines.Add()
                                    Case "-18"
                                        oProductionorder.Lines.LineText = rs.Fields.Item("LineText").Value
                                        oProductionorder.Lines.Add()
                                End Select
                                Phrs.MoveNext()
                            Next
                        ElseIf rs.Fields.Item("Code").Value <> "" And rs.Fields.Item("Phantom").Value = "Y" And oMatrix1.Columns.Item("PhanItm").Cells.Item(i).Specific.value = "" Then
                            Dim Phrs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            Phrs.DoQuery("Select Type,a.Code 'Code',Quantity,isnull(LineText,'') 'LineText',b.Qauntity,isnull(a.IssueMthd,'') 'IssueMthd',c.Phantom from ITT1 a inner join OITT b on a.Father=b.Code Inner Join OITM c on a.Code=c.ItemCode where Father ='" & rs.Fields.Item("Code").Value & "' and c.Phantom='N'")
                            For Ph = 0 To Phrs.RecordCount - 1
                                Select Case (Phrs.Fields.Item("Type").Value)
                                    Case "4"
                                        oProductionorder.Lines.ItemType = SAPbobsCOM.ProductionItemType.pit_Item
                                        oProductionorder.Lines.ItemNo = Phrs.Fields.Item("Code").Value
                                        oProductionorder.Lines.BaseQuantity = CDbl(CDbl(rs.Fields.Item("Quantity").Value) * CDbl(Phrs.Fields.Item("Quantity").Value))
                                        oProductionorder.Lines.PlannedQuantity = CDbl(CDbl(rs.Fields.Item("Quantity").Value) * CDbl(Phrs.Fields.Item("Quantity").Value) * (CDbl(oMatrix1.Columns.Item("DPlQty").Cells.Item(i).Specific.value)))
                                        oProductionorder.Lines.Warehouse = oMatrix1.Columns.Item("CWhs").Cells.Item(i).Specific.value
                                        If Phrs.Fields.Item("IssueMthd").Value = "B" Then
                                            oProductionorder.Lines.ProductionOrderIssueType = SAPbobsCOM.BoIssueMethod.im_Backflush
                                        Else
                                            oProductionorder.Lines.ProductionOrderIssueType = SAPbobsCOM.BoIssueMethod.im_Manual
                                        End If
                                        oProductionorder.Lines.Add()
                                    Case "290"
                                        oProductionorder.Lines.ItemType = SAPbobsCOM.ProductionItemType.pit_Resource
                                        oProductionorder.Lines.ItemNo = Phrs.Fields.Item("Code").Value
                                        oProductionorder.Lines.BaseQuantity = CDbl(rs.Fields.Item("Quantity").Value) * CDbl(Phrs.Fields.Item("Quantity").Value)
                                        oProductionorder.Lines.PlannedQuantity = CDbl(CDbl(rs.Fields.Item("Quantity").Value) * CDbl(Phrs.Fields.Item("Quantity").Value) * (CDbl(oMatrix1.Columns.Item("DPlQty").Cells.Item(i).Specific.value)))
                                        oProductionorder.Lines.Warehouse = oMatrix1.Columns.Item("CWhs").Cells.Item(i).Specific.value
                                        If Phrs.Fields.Item("IssueMthd").Value = "B" Then
                                            oProductionorder.Lines.ProductionOrderIssueType = SAPbobsCOM.BoIssueMethod.im_Backflush
                                        Else
                                            oProductionorder.Lines.ProductionOrderIssueType = SAPbobsCOM.BoIssueMethod.im_Manual
                                        End If
                                        oProductionorder.Lines.Add()
                                    Case "-18"
                                        oProductionorder.Lines.LineText = rs.Fields.Item("LineText").Value
                                        oProductionorder.Lines.Add()
                                End Select
                                Phrs.MoveNext()
                            Next
                        End If
                        rs.MoveNext()
                    Next
                    oProductionorder.Remarks = oGFun.getSingleValue("Select isnull(WhsName,'') as 'WhsName' From OWHS where WhsCode ='" & oMatrix1.Columns.Item("CWhs").Cells.Item(i).Specific.value & "'") + "," + oMatrix1.Columns.Item("CWhs").Cells.Item(i).Specific.value + "," + oDBDSHeader.GetValue("DocNum", 0) + "," + oDBDSHeader.GetValue("U_JobNo", 0)
                    ErrCode = oProductionorder.Add()
                    If ErrCode <> 0 Then
                        oProgBar.Stop()
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oProgBar)
                        oProgBar = Nothing
                        GC.Collect()
                        oGFun.Msg("Child Production Posting Error : " & oCompany.GetLastErrorDescription)
                        Return False
                    Else
                        If FlageOp = True Then
                            oProgBar.Text = "Production Order Created for [" & ItemCode & "]"
                        Else
                            oProgBar.Text = "Validating Production Order : [" & ItemCode & "]...."
                        End If
                        Dim sNewObjCode As String = ""
                        oCompany.GetNewObjectCode(sNewObjCode)
                        Dim str = CLng(sNewObjCode)
                        oMatrix1.Columns.Item("ProNum").Cells.Item(i).Specific.value = oGFun.getSingleValue("Select DocNum from OWOR where Docentry='" & str & "'")
                        oMatrix1.Columns.Item("ProEntry").Cells.Item(i).Specific.value = str
                        Dim ocombo As SAPbouiCOM.ComboBox
                        ocombo = oMatrix1.Columns.Item("ProSt").Cells.Item(i).Specific
                        ocombo.Select("P", SAPbouiCOM.BoSearchKey.psk_ByValue)
                    End If
                End If
                oProgBar.Value = i
            Next
            oProgBar.Stop()
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oProgBar)
            oProgBar = Nothing
            GC.Collect()
            Return True
        Catch ex As Exception
            oApplication.StatusBar.SetText("Posting Event Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            oProgBar.Stop()
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oProgBar)
            oProgBar = Nothing
            GC.Collect()
            Return False
        Finally
        End Try
    End Function
    Private Sub UpdateMode()
        Try
            frmProductionPlanning.Items.Item("Matrix1").Specific.Columns.Item("Chck").Editable = True
            frmProductionPlanning.Items.Item("Matrix1").Specific.Columns.Item("ItemC").Editable = True
            'frmProductionPlanning.Items.Item("Matrix1").Specific.Columns.Item("ItemD").Editable = False
            frmProductionPlanning.Items.Item("Matrix1").Specific.Columns.Item("DPlQty").Editable = True
            'frmProductionPlanning.Items.Item("Matrix1").Specific.Columns.Item("BWhs").Editable = False
            'frmProductionPlanning.Items.Item("Matrix1").Specific.Columns.Item("CWhs").Editable = False
            frmProductionPlanning.Items.Item("Matrix1").Specific.Columns.Item("MPlQty").Editable = False
            frmProductionPlanning.Items.Item("Matrix1").Specific.Columns.Item("MCplQty").Editable = False
            'For i As Integer = 1 To oMatrix1.RowCount
            '    If oMatrix1.Columns.Item("ItemC").Cells.Item(i).Specific.Value <> "" And oMatrix1.Columns.Item("Chck").Cells.Item(i).Specific.checked = False Then
            '        oMatrix1.Columns.Item("Chck").Cells.Item(i).Specific.checked = True
            '    End If
            'Next
            Dim Flag As Boolean = False
            oDBDSDetail = frmProductionPlanning.DataSources.DBDataSources.Item("@MIPL_PPN1")
            oMatrix1.FlushToDataSource()
            For rowNum As Integer = 0 To oDBDSDetail.Size - 1
                If oDBDSDetail.GetValue("U_ItemCode", rowNum) <> "" And (oDBDSDetail.GetValue("U_select", rowNum) = "" Or oDBDSDetail.GetValue("U_select", rowNum) = "N") Then
                    oDBDSDetail.SetValue("U_select", rowNum, "Y") : Flag = True
                End If
            Next
            frmProductionPlanning.Freeze(True)
            oMatrix1.LoadFromDataSource()
            frmProductionPlanning.Freeze(False)
            frmProductionPlanning.Refresh()
            'If Flag = True Then
            If frmProductionPlanning.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then frmProductionPlanning.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
            If frmProductionPlanning.Items.Item("CmbSt").Specific.selected.value = "R" Or frmProductionPlanning.Items.Item("CmbSt").Specific.selected.value = "C" Then
                frmProductionPlanning.Items.Item("2A").Enabled = False
            Else
                frmProductionPlanning.Items.Item("2A").Enabled = True
                '  frmProductionPlanning.Items.Item("Matrix1").Specific.Columns.Item("CWhs").Editable = True ' To change the warehouse
                ' frmProductionPlanning.Items.Item("Matrix1").Specific.Columns.Item("BWhs").Editable = True
            End If
            frmProductionPlanning.Items.Item("t_Month").Enabled = False
            frmProductionPlanning.Items.Item("31").Specific.Caption = ""
            If frmProductionPlanning.Items.Item("CmbSt").Specific.selected.value = "O" Then oGFun.Matrix_Addrow(oMatrix1, "ItemC", "LineId")
        Catch ex As Exception
        End Try
    End Sub
    Private Sub AddMode()
        Try
            frmProductionPlanning.Items.Item("Matrix1").Specific.Columns.Item("Chck").Editable = False
            frmProductionPlanning.Items.Item("Matrix1").Specific.Columns.Item("ItemC").Editable = True
            frmProductionPlanning.Items.Item("Matrix1").Specific.Columns.Item("ItemD").Editable = True
            frmProductionPlanning.Items.Item("Matrix1").Specific.Columns.Item("DPlQty").Editable = True
            frmProductionPlanning.Items.Item("Matrix1").Specific.Columns.Item("BWhs").Editable = True
            frmProductionPlanning.Items.Item("Matrix1").Specific.Columns.Item("CWhs").Editable = True
            frmProductionPlanning.Items.Item("Matrix1").Specific.Columns.Item("MPlQty").Editable = True
            frmProductionPlanning.Items.Item("Matrix1").Specific.Columns.Item("MCplQty").Editable = True
            frmProductionPlanning.Items.Item("2A").Enabled = False
        Catch ex As Exception
        End Try
    End Sub
    Private Sub releasebar()

    End Sub
    Public Sub addReporttype()
        Dim rptTypeService As SAPbobsCOM.ReportTypesService
        Dim newType As SAPbobsCOM.ReportType
        Dim newtypeParam As SAPbobsCOM.ReportTypeParams
        Dim newReportParam As SAPbobsCOM.ReportLayoutParams
        Dim ReportExists As Boolean = False
        Try
            If oCompany Is Nothing Then Exit Sub
            Dim newtypesParam As SAPbobsCOM.ReportTypesParams
            rptTypeService = oCompany.GetCompanyService.GetBusinessService(SAPbobsCOM.ServiceTypes.ReportTypesService)
            newtypesParam = rptTypeService.GetReportTypeList
            strsql = oGFun.getSingleValue("Select 1 as Status from RTYP Where NAME='" & ProductionPlanningFormId & "'")
            If strsql <> "" Then Exit Sub
            ReportExists = True
            'Dim i As Integer
            'For i = 0 To newtypesParam.Count - 1
            '    If newtypesParam.Item(i).TypeName = ProductionPlanningFormId And newtypesParam.Item(i).MenuID = ProductionPlanningFormId Then
            '        ReportExists = True
            '        Exit For
            '    End If
            'Next i

            If ReportExists Then
                rptTypeService = oCompany.GetCompanyService.GetBusinessService(SAPbobsCOM.ServiceTypes.ReportTypesService)
                newType = rptTypeService.GetDataInterface(SAPbobsCOM.ReportTypesServiceDataInterfaces.rtsReportType)


                newType.TypeName = ProductionPlanningFormId
                newType.AddonName = "ProPlanAddon"
                newType.AddonFormType = ProductionPlanningFormId
                newType.MenuID = ProductionPlanningFormId
                newtypeParam = rptTypeService.AddReportType(newType)

                Dim rptService As SAPbobsCOM.ReportLayoutsService
                Dim newReport As SAPbobsCOM.ReportLayout
                rptService = oCompany.GetCompanyService.GetBusinessService(SAPbobsCOM.ServiceTypes.ReportLayoutsService)
                newReport = rptService.GetDataInterface(SAPbobsCOM.ReportLayoutsServiceDataInterfaces.rlsdiReportLayout)
                newReport.Author = oCompany.UserName
                newReport.Category = SAPbobsCOM.ReportLayoutCategoryEnum.rlcCrystal
                newReport.Name = ProductionPlanningFormId
                newReport.TypeCode = newtypeParam.TypeCode

                newReportParam = rptService.AddReportLayout(newReport)

                newType = rptTypeService.GetReportType(newtypeParam)
                newType.DefaultReportLayout = newReportParam.LayoutCode
                rptTypeService.UpdateReportType(newType)

                Dim oBlobParams As SAPbobsCOM.BlobParams
                oBlobParams = oCompany.GetCompanyService.GetDataInterface(SAPbobsCOM.CompanyServiceDataInterfaces.csdiBlobParams)
                oBlobParams.Table = "RDOC"
                oBlobParams.Field = "Template"
                Dim oKeySegment As SAPbobsCOM.BlobTableKeySegment
                oKeySegment = oBlobParams.BlobTableKeySegments.Add
                oKeySegment.Name = "DocCode"
                oKeySegment.Value = newReportParam.LayoutCode

                Dim oFile As FileStream
                oFile = New FileStream(System.Windows.Forms.Application.StartupPath + "\ProPlan.rpt", FileMode.Open)
                Dim fileSize As Integer
                fileSize = oFile.Length
                Dim buf(fileSize) As Byte
                oFile.Read(buf, 0, fileSize)
                oFile.Dispose()

                Dim oBlob As SAPbobsCOM.Blob
                oBlob = oCompany.GetCompanyService.GetDataInterface(SAPbobsCOM.CompanyServiceDataInterfaces.csdiBlob)
                oBlob.Content = Convert.ToBase64String(buf, 0, fileSize)
                oCompany.GetCompanyService.SetBlob(oBlobParams, oBlob)
            End If
        Catch ex As Exception
            oApplication.MessageBox(ex.Message)
        End Try

    End Sub
    Private Sub setReport(ByVal FormUID As String)
        Try
            frmProductionPlanning = oApplication.Forms.Item(FormUID)
            'Dim rptTypeService As SAPbobsCOM.ReportTypesService
            'Dim newType As SAPbobsCOM.ReportType
            'Dim newtypesParam As SAPbobsCOM.ReportTypesParams
            'rptTypeService = oCompany.GetCompanyService.GetBusinessService(SAPbobsCOM.ServiceTypes.ReportTypesService)
            'newtypesParam = rptTypeService.GetReportTypeList

            'Dim i As Integer
            'For i = 0 To newtypesParam.Count - 1

            '    If newtypesParam.Item(i).TypeName = ProductionPlanningFormId And newtypesParam.Item(i).MenuID = ProductionPlanningFormId Then
            '        frmProductionPlanning.ReportType = newtypesParam.Item(i).TypeCode

            '        Exit For
            '    End If
            'Next i

            Dim TypeCode As String
            TypeCode = oGFun.getSingleValue("Select CODE from RTYP where NAME='" & ProductionPlanningFormId & "'")
            frmProductionPlanning.ReportType = TypeCode
        Catch ex As Exception

        End Try

    End Sub
    Private Sub RemoveLastrow(ByVal omatrix As SAPbouiCOM.Matrix, ByVal Columname_check As String)
        Try
            If omatrix.VisualRowCount = 0 Or omatrix.VisualRowCount = 1 Then Exit Sub
            If Columname_check.ToString = "" Then Exit Sub
            If omatrix.Columns.Item(Columname_check).Cells.Item(omatrix.VisualRowCount).Specific.string = "" Then
                omatrix.DeleteRow(omatrix.VisualRowCount)
            End If
        Catch ex As Exception

        End Try
    End Sub
    Public Sub LayoutKeyEvent(ByRef eventInfo As SAPbouiCOM.LayoutKeyInfo, ByRef BubbleEvent As Boolean)
        Try
            frmProductionPlanning = oApplication.Forms.Item(eventInfo.FormUID)
            eventInfo.LayoutKey = frmProductionPlanning.DataSources.DBDataSources.Item("@MIPL_OPPN").GetValue("DocEntry", 0) 'frmProductionPlanning.Items.Item("12A").Specific.string
        Catch ex As Exception
        End Try

    End Sub

End Class
