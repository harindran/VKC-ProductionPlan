Public Class ReceiptFromProduction
    Public frmRecFromProduction As SAPbouiCOM.Form
    Dim oDBDSHeader, oDBDSDetail, oDetailsHeaderInv As SAPbouiCOM.DBDataSource
    Dim oMatrix As SAPbouiCOM.Matrix
    Dim ParentCode As String = ""
    Dim CurrentRow As Integer = 0
    Dim FormExist = False
    Dim Instruction As String = ""
    Dim ParCode
    Dim docnum = ""
    Dim FGItemCode = ""
    Dim InQty As Double = 0
    Dim PreWODocEntry = ""
    Sub LoadForm()
        Try
            frmRecFromProduction = oApplication.Forms.ActiveForm
            oDBDSHeader = frmRecFromProduction.DataSources.DBDataSources.Add("OIGN")
            oDBDSDetail = frmRecFromProduction.DataSources.DBDataSources.Add("IGN1")
            oMatrix = frmRecFromProduction.Items.Item("13").Specific
            ' Button creation
            Dim oItem As SAPbouiCOM.Item
            Dim oButton As SAPbouiCOM.Button
            Try
                oItem = frmRecFromProduction.Items.Add("B_InvTrns", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            Catch ex As Exception
            End Try
            oItem = frmRecFromProduction.Items.Item("B_InvTrns")
            oItem.Left = frmRecFromProduction.Items.Item("2").Left + 80
            oItem.Top = frmRecFromProduction.Items.Item("2").Top
            oItem.Height = frmRecFromProduction.Items.Item("1").Height
            oItem.Width = frmRecFromProduction.Items.Item("1").Width + 10
            oItem.Enabled = True
            oButton = oItem.Specific
            oButton.Caption = "Inv.Transfer"
            Me.InitForm()
            Me.DefineModesForFields()
        Catch ex As Exception
            oGFun.Msg("Load Form Method Failed:" & ex.Message)
        End Try
    End Sub
    Sub InitForm()
        Try

        Catch ex As Exception
            oApplication.StatusBar.SetText("InitForm Method Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
        End Try
    End Sub
    Sub DefineModesForFields()
        Try
            frmRecFromProduction.Items.Item("B_InvTrns").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Find, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            frmRecFromProduction.Items.Item("B_InvTrns").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Add, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            frmRecFromProduction.Items.Item("B_InvTrns").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
        Catch ex As Exception
            oApplication.StatusBar.SetText("DefineModesForFields Method Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
        End Try
    End Sub
    Function ValidateAll() As Boolean
        Try
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
                            End Select
                        End If
                        If pVal.BeforeAction = False Then
                        End If
                    Catch ex As Exception
                        oApplication.StatusBar.SetText("Choose From List Event Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    Finally
                    End Try

                Case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE

                    Try

                    Catch ex As Exception
                        oApplication.StatusBar.SetText("FORM_CLOSE Event Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    End Try

                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                    Try

                    Catch ex As Exception
                        oApplication.StatusBar.SetText("ITEM_PRESSED Event Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    End Try
                Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                    Try
                        Select Case pVal.ItemUID

                        End Select
                    Catch ex As Exception
                        oApplication.StatusBar.SetText("Last Focus Event Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    Finally
                    End Try
                Case SAPbouiCOM.BoEventTypes.et_VALIDATE
                    Try
                        Select Case pVal.ColUID
                        End Select
                    Catch ex As Exception
                        oApplication.StatusBar.SetText("Validate Event Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    End Try
                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    Try
                        Select Case pVal.ItemUID
                            Case "1"
                                If pVal.BeforeAction = False And frmRecFromProduction.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then

                                End If
                        End Select
                    Catch ex As Exception
                        oApplication.StatusBar.SetText("Item Press Event Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    End Try
                Case SAPbouiCOM.BoEventTypes.et_CLICK
                    Try
                        Select Case pVal.ItemUID
                            Case "B_InvTrns"
                                If pVal.BeforeAction = False And frmRecFromProduction.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                    Try
                                        frmRecFromProduction.Freeze(True)
                                        Dim Docentry As String = ""
                                        oGFun.Msg("Loading Data.....", "S", "W")
                                        oApplication.ActivateMenuItem("3080")
                                        GC.Collect()
                                        Dim FrmInvTrns As SAPbouiCOM.Form = oApplication.Forms.ActiveForm
                                        Dim oMatrix1 As SAPbouiCOM.Matrix = FrmInvTrns.Items.Item("23").Specific
                                        Dim PlnDocentry = oDBDSHeader.GetValue("DocEntry", 0)
                                        Dim WareHouse As String = ""
                                        For i = 1 To oMatrix.RowCount
                                            oMatrix1.Columns.Item("1").Cells.Item(oMatrix1.VisualRowCount).Specific.value = oMatrix.Columns.Item("1").Cells.Item(i).Specific.value
                                            oMatrix1.Columns.Item("10").Cells.Item(oMatrix1.VisualRowCount - 1).Specific.value = CDbl(oMatrix.Columns.Item("9").Cells.Item(i).Specific.value)
                                            oMatrix1.Columns.Item("1470001039").Cells.Item(oMatrix1.VisualRowCount - 1).Specific.value = oMatrix.Columns.Item("15").Cells.Item(i).Specific.value
                                            oMatrix1.Columns.Item("U_RecptEntry").Cells.Item(oMatrix1.VisualRowCount - 1).Specific.value = PlnDocentry
                                            oMatrix1.Columns.Item("U_ProQty").Cells.Item(oMatrix1.VisualRowCount - 1).Specific.value = oGFun.getSingleValue("select isnull(sum(Quantity),0) 'Quantity' from WTR1  inner join oign a on a.DocEntry=U_RecptEntry where  ItemCode='" & oMatrix.Columns.Item("1").Cells.Item(i).Specific.value & "' and a.DocEntry='" & PlnDocentry & "' group by U_RecptEntry,ItemCode")
                                            If i = 1 Then WareHouse = oMatrix.Columns.Item("15").Cells.Item(i).Specific.value
                                        Next
                                        FrmInvTrns.Items.Item("18").Specific.value = WareHouse
                                        frmRecFromProduction.Freeze(False)
                                    Catch ex As Exception
                                        oApplication.StatusBar.SetText("Inventory Transfer :" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                        frmRecFromProduction.Freeze(False)
                                    End Try
                                End If
                        End Select
                    Catch ex As Exception
                        oApplication.StatusBar.SetText("Click Event Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        BubbleEvent = False
                    Finally
                    End Try
            End Select
        Catch ex As Exception
            oApplication.StatusBar.SetText("Item Event Method Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
        End Try
    End Sub
    Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            Select Case BusinessObjectInfo.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                    Try
                        If BusinessObjectInfo.ActionSuccess And frmRecFromProduction.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                            Try
                                If BusinessObjectInfo.ActionSuccess = True Then
                                    Dim rset As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    Dim str As String = " "
                                    For i As Integer = 1 To oMatrix.VisualRowCount - 1
                                        If oMatrix.Columns.Item("61").Cells.Item(i).Specific.value <> "" Then
                                            str = str + " if '" & oMatrix.Columns.Item("69").Cells.Item(i).Specific.value & "'='C'begin " & _
                                                        " Update [@MIPL_PPN1] set U_DCompQty= U_DCompQty + " & oMatrix.Columns.Item("9").Cells.Item(i).Specific.value & " where U_PoDocEntry='" & oDBDSDetail.GetValue("BaseEntry", i - 1) & "'" & _
                                                        " end else if '" & oMatrix.Columns.Item("69").Cells.Item(i).Specific.value & "' ='R' begin " & _
                                                        " Update [@MIPL_PPN1] set U_RPlaQty=U_RPlaQty+" & oMatrix.Columns.Item("9").Cells.Item(i).Specific.value & " where U_PoDocEntry='" & oDBDSDetail.GetValue("BaseEntry", i - 1) & "' end "
                                        End If
                                    Next
                                    rset.DoQuery(str)
                                End If
                            Catch ex As Exception
                                oApplication.StatusBar.SetText("Updating Daily Planning Event Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                BubbleEvent = False
                            End Try
                        End If
                    Catch ex As Exception
                        oApplication.StatusBar.SetText("Form Data ADD Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        BubbleEvent = False
                    Finally
                    End Try
            End Select
        Catch ex As Exception
            oApplication.StatusBar.SetText("Form Data Event Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
        End Try
    End Sub
    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub
End Class
