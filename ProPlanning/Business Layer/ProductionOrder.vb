Public Class ProductionOrder
    Public frmRecFromProduction As SAPbouiCOM.Form
    Dim oDBDSHeader, oDBDSDetail As SAPbouiCOM.DBDataSource
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
    Dim oProgBar As SAPbouiCOM.ProgressBar
    Sub LoadForm()
        Try
            frmRecFromProduction = oApplication.Forms.ActiveForm
            oDBDSHeader = frmRecFromProduction.DataSources.DBDataSources.Item(0)
            oDBDSDetail = frmRecFromProduction.DataSources.DBDataSources.Item(1)
            oMatrix = frmRecFromProduction.Items.Item("37").Specific
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
                            Case "1"

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
    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.MenuUID
                Case "1282"
                    Me.InitForm()
                Case "1293"
                Case "1284"
                    Try

                    Catch ex As Exception

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
                    Try
                        If BusinessObjectInfo.ActionSuccess And frmRecFromProduction.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then

                        End If
                    Catch ex As Exception
                        oApplication.StatusBar.SetText("Form Data ADD Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        BubbleEvent = False
                    Finally
                    End Try
                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE
                    Try
                        If BusinessObjectInfo.ActionSuccess = True Then
                            Dim rset As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            Dim str As String = "Update [@MIPL_PPN1] set U_PoStatus='" & oDBDSHeader.GetValue("Status", 0) & "', U_DCPlaQty = '" & oDBDSHeader.GetValue("PlannedQty", 0) & "' where U_PoDocEntry='" & oDBDSHeader.GetValue("DocEntry", 0) & "' and LineId ='" & oDBDSHeader.GetValue("U_PPLNLineNum", 0) & "'"
                            rset.DoQuery(str)
                        End If
                    Catch ex As Exception
                        oApplication.StatusBar.SetText("Updating Daily Planning Event Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        BubbleEvent = False
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
