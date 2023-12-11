Public Class MonthlyPlanning
    Public frmProductionMthPlan As SAPbouiCOM.Form
    Dim oDBDSHeader As SAPbouiCOM.DBDataSource
    Dim oDBDSDetail As SAPbouiCOM.DBDataSource
    Public oMatrix1 As SAPbouiCOM.Matrix
    Dim oMatrixFun As SAPbouiCOM.IMatrix
    Dim UDOID As String = "MPLN"
    Dim oedit As SAPbouiCOM.EditText
    Dim FlageOp = False, FlageLb = False, FlageMc = False
    Dim Itemstrsql As String = ""
    Dim OperationType = ""
    Dim strsql As String = ""
    Dim ProgValue As Integer = 0.0
    Sub LoadForm()
        Try
            oGFun.LoadXML(frmProductionMthPlan, MonthPlanningFormId, MonthPlanningXML)
            frmProductionMthPlan = oApplication.Forms.Item(MonthPlanningFormId)
            oDBDSHeader = frmProductionMthPlan.DataSources.DBDataSources.Item(0)
            oDBDSDetail = frmProductionMthPlan.DataSources.DBDataSources.Item(1)
            oMatrix1 = frmProductionMthPlan.Items.Item("Matrix1").Specific
            Me.DefineModesForFields()
            Me.InitForm()
        Catch ex As Exception
            oGFun.Msg("Load Form Method Failed:" & ex.Message)
        End Try
    End Sub
    Sub InitForm()
        Try
            frmProductionMthPlan.Freeze(True)
            oGFun.LoadComboBoxSeries(frmProductionMthPlan.Items.Item("c_Series").Specific, UDOID)
            oGFun.LoadDocumentDate(frmProductionMthPlan.Items.Item("t_PoDate").Specific)
            oGFun.setComboBoxValue(frmProductionMthPlan.Items.Item("lt_Type").Specific, "Select code,Name from [@MIPL_PROTYPE]")
            oMatrix1.Clear()
            oDBDSDetail.Clear()
            oGFun.SetNewLine(oMatrix1, oDBDSDetail)
            oMatrixFun = oMatrix1
            oMatrixFun.CommonSetting.EnableArrowKey = True
        Catch ex As Exception
            oApplication.StatusBar.SetText("InitForm Method Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
            frmProductionMthPlan.Freeze(False)
        End Try
    End Sub
    Sub DefineModesForFields()
        Try

            frmProductionMthPlan.Items.Item("t_DocNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            frmProductionMthPlan.Items.Item("lt_CostCe").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Find, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            frmProductionMthPlan.Items.Item("lt_CostCe").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Add, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            frmProductionMthPlan.Items.Item("lt_CostCe").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            frmProductionMthPlan.Items.Item("lt_Type").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Find, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            frmProductionMthPlan.Items.Item("lt_Type").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Add, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            frmProductionMthPlan.Items.Item("lt_Type").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            frmProductionMthPlan.Items.Item("t_Month").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Find, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            frmProductionMthPlan.Items.Item("t_Month").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Add, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            frmProductionMthPlan.Items.Item("t_Month").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            frmProductionMthPlan.Items.Item("t_PlnBy").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Find, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            frmProductionMthPlan.Items.Item("t_PlnBy").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Add, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            frmProductionMthPlan.Items.Item("t_PlnBy").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            frmProductionMthPlan.Items.Item("c_Series").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Find, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            frmProductionMthPlan.Items.Item("c_Series").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Add, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            frmProductionMthPlan.Items.Item("c_Series").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            frmProductionMthPlan.Items.Item("t_PoDate").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Find, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            frmProductionMthPlan.Items.Item("t_PoDate").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Add, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            frmProductionMthPlan.Items.Item("t_PoDate").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            frmProductionMthPlan.Items.Item("t_SDate").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Find, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            frmProductionMthPlan.Items.Item("t_SDate").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Add, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            frmProductionMthPlan.Items.Item("t_SDate").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            frmProductionMthPlan.Items.Item("C_status").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Find, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            frmProductionMthPlan.Items.Item("C_status").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Add, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            frmProductionMthPlan.Items.Item("C_status").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            frmProductionMthPlan.Items.Item("Matrix1").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Find, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            frmProductionMthPlan.Items.Item("Matrix1").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Add, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            frmProductionMthPlan.Items.Item("Matrix1").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

        Catch ex As Exception
            oApplication.StatusBar.SetText("DefineModesForFields Method Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
        End Try
    End Sub
    Function ValidateAll() As Boolean
        Try
            Dim RowCount = 0
            'If frmProductionMthPlan.Items.Item("t_PlnBy").Specific.value.ToString.Trim.Equals("") = True Then
            '    oGFun.Msg("Planned by Should Not Be Left Empty")
            '    Return False
            'End If
            If frmProductionMthPlan.Items.Item("lt_CostCe").Specific.value.ToString.Trim.Equals("") = True Then
                oGFun.Msg("Cost Center Should Not Be Left Empty")
                Return False
            End If
            If frmProductionMthPlan.Items.Item("t_SDate").Specific.value.ToString.Trim.Equals("") = True Then
                oGFun.Msg("Year Should Not Be Left Empty")
                Return False
            End If
            For i As Integer = 1 To oMatrix1.VisualRowCount - 1
                If oMatrix1.Columns.Item("ItemC").Cells.Item(i).Specific.value <> "" Then
                    RowCount = RowCount + 1
                End If
            Next
            If RowCount = 0 Then
                oGFun.Msg("Minimum one Line Item is Required.. ", "E")
                Return False
            End If
            For i As Integer = 1 To oMatrix1.VisualRowCount
                If oMatrix1.Columns.Item("ItemC").Cells.Item(i).Specific.value <> "" Then
                    If oGFun.getSingleValue("Select count (*) from OITT where U_ProcessType='" & oDBDSHeader.GetValue("U_ProType", 0) & "' and code ='" & oMatrix1.Columns.Item("ItemC").Cells.Item(i).Specific.value & "'") = "0" Then
                        oGFun.Msg("[" & oMatrix1.Columns.Item("ItemC").Cells.Item(i).Specific.value & "] This item is not belong to Selected Process Type  in LineNo : [" & i & "] ")
                        Return False
                    End If
                    If oMatrix1.Columns.Item("DPlQty").Cells.Item(i).Specific.value <= 0 Then
                        oGFun.Msg("Monthly Planned Qty Should be greater that 0  in LineNo : [" & i & "] ")
                        Return False
                    End If
                End If
            Next
            ValidateAll = True
        Catch ex As Exception
            oApplication.StatusBar.SetText("Validate Function Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            ValidateAll = False
        Finally
        End Try
    End Function
    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                    Try
                        Dim oDataTable As SAPbouiCOM.DataTable
                        Dim oCFLE As SAPbouiCOM.ChooseFromListEvent = pVal
                        oDataTable = oCFLE.SelectedObjects
                        Dim rset As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        If pVal.BeforeAction Then
                            Select Case oCFLE.ChooseFromListUID
                                Case "cfl_ItmCode"
                                    'strsql = ""
                                    'strsql = "Select ItemCode from OITM T0 inner join OITT T1 on T0.ItemCode=T1.Code where T1.TreeType='P' and T1.U_ProcessType='" & oDBDSHeader.GetValue("U_ProType", 0) & "'"
                                    'oGFun.ChooseFromListFilteration(frmProductionMthPlan, "cfl_ItmCode", "ItemCode", strsql)
                            End Select
                        End If
                        If Not oDataTable Is Nothing And pVal.BeforeAction = False Then
                            Select Case oCFLE.ChooseFromListUID
                                Case "cfl_ItmCode"
                                    oMatrix1.FlushToDataSource()
                                    oDBDSDetail.SetValue("U_ItemCode", pVal.Row - 1, oDataTable.GetValue("ItemCode", 0))
                                    oDBDSDetail.SetValue("U_ItemDesc", pVal.Row - 1, oDataTable.GetValue("ItemName", 0))
                                    oMatrix1.LoadFromDataSourceEx()
                                    oGFun.SetNewLine(oMatrix1, oDBDSDetail, pVal.Row, "ItemC")
                                Case "cfl_Emp"
                                    oDBDSHeader.SetValue("U_PlanBy", 0, oDataTable.GetValue("lastName", 0) + ", " + oDataTable.GetValue("firstName", 0))
                            End Select
                        End If
                    Catch ex As Exception
                        oApplication.StatusBar.SetText("Choose From List Event Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    Finally
                    End Try
                Case SAPbouiCOM.BoEventTypes.et_VALIDATE
                    Try
                        Select Case pVal.ItemUID


                        End Select

                    Catch ex As Exception
                        oApplication.StatusBar.SetText("Validate Event Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    Finally
                    End Try

                Case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS
                    Select Case pVal.ColUID

                    End Select
                Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                    Try
                        Select Case pVal.ItemUID

                            Case "Matrix1"
                                Select Case pVal.ColUID

                                End Select
                        End Select
                    Catch ex As Exception
                        oApplication.StatusBar.SetText("Lost Focus Event Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    Finally
                    End Try
                Case (SAPbouiCOM.BoEventTypes.et_COMBO_SELECT)
                    Try
                        Select Case pVal.ItemUID
                            Case "c_Series"
                                If frmProductionMthPlan.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE And pVal.BeforeAction = False Then
                                    oGFun.setDocNum(frmProductionMthPlan)
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
                    Try
                        Select Case pVal.ItemUID
                            Case "1"
                                If pVal.BeforeAction = True And (frmProductionMthPlan.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE) Or (frmProductionMthPlan.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                                    If Me.ValidateAll() = False Then
                                        System.Media.SystemSounds.Asterisk.Play()
                                        BubbleEvent = False
                                        Exit Sub
                                    Else
                                    End If
                                End If
                        End Select
                    Catch ex As Exception
                        oApplication.StatusBar.SetText("Click Event Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                        BubbleEvent = False
                    Finally
                    End Try
                Case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK
                    Try
                        Select Case pVal.ItemUID

                        End Select
                    Catch ex As Exception
                        oApplication.StatusBar.SetText("Double Click Event Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    End Try
                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    Try
                        Select Case pVal.ItemUID
                            Case "1"
                                If pVal.BeforeAction = False And frmProductionMthPlan.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                    Me.InitForm()
                                End If
                        End Select
                    Catch ex As Exception
                        oApplication.StatusBar.SetText("Item Pressed Event Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    Finally
                    End Try
                Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED
                    Try
                        If pVal.BeforeAction = True Then
                            Select Case pVal.ColUID
                                
                            End Select
                        End If
                    Catch ex As Exception

                    End Try
            End Select
        Catch ex As Exception
            oApplication.StatusBar.SetText("Item Event Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
        End Try
    End Sub
    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.MenuUID
                Case "1282"
                    Me.InitForm()
                Case "1284"
                    Try
                        If frmProductionMthPlan.Mode <> SAPbouiCOM.BoFormMode.fm_ADD_MODE And pVal.BeforeAction = True Then
                            Dim check = oApplication.MessageBox("Monthly Production Planning will be Canceled. Continue ? ", 2, "Continue", "Cancel")
                            If check <> 1 Then BubbleEvent = False
                        End If
                    Catch ex As Exception
                        oApplication.StatusBar.SetText("Menu Cancel Event Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    End Try
            End Select
        Catch ex As Exception
            oApplication.StatusBar.SetText("Menu Event Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
        End Try
    End Sub
    Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            Select Case BusinessObjectInfo.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                    Dim rsetUpadteIndent As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Try
                        If BusinessObjectInfo.BeforeAction Then
                            If Me.ValidateAll() = False Then
                                System.Media.SystemSounds.Asterisk.Play()
                                BubbleEvent = False
                                Exit Sub
                            Else
                                If frmProductionMthPlan.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                    oGFun.setDocNum(frmProductionMthPlan)
                                    oGFun.DeleteEmptyRowInFormDataEvent(oMatrix1, "ItemC", oDBDSDetail)
                                End If
                            End If
                        End If
                    Catch ex As Exception
                        BubbleEvent = False
                    Finally
                    End Try
                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                    If BusinessObjectInfo.ActionSuccess = True Then
                        frmProductionMthPlan.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                    End If
            End Select
        Catch ex As Exception
            If oCompany.InTransaction Then oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            oApplication.StatusBar.SetText("Form Data Event Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
        End Try
    End Sub
End Class
