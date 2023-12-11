Imports SDKLib881
Module EventHandler
    Public WithEvents oApplication As SAPbouiCOM.Application
    Public oForm As SAPbouiCOM.Form

#Region " ... 1) Menu Event ..."

    Private Sub oApplication_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) Handles oApplication.MenuEvent
        Try
            If pVal.BeforeAction = False Then
                Select Case pVal.MenuUID
                    Case ProductionPlanningFormId
                        oProductionPlan.LoadForm()
                    Case MonthPlanningFormId
                        oProductionMthPlan.LoadForm()
                        '    Case PROMenuId
                        '        oProductionOrder.LoadForm()
                    Case PRORecptMenuId
                        oReceiptFromPro.LoadForm()
                End Select
                oForm = oApplication.Forms.ActiveForm
                Select Case pVal.MenuUID
                    Case "1282", "1281", "1292", "1293", "1287", "519", "1284", "1286", "5890", "1290", "1289", "1288", "1291"
                        Select Case oForm.TypeEx
                            Case ProductionPlanningFormId
                                oProductionPlan.MenuEvent(pVal, BubbleEvent)
                            Case MonthPlanningFormId
                                oProductionMthPlan.MenuEvent(pVal, BubbleEvent)
                        End Select
                        Select Case pVal.MenuUID
                            Case "1282", "1284"
                                'Select Case oForm.TypeEx
                                '    'Case PROTypeEx
                                '    '    oProductionOrder.MenuEvent(pVal, BubbleEvent)
                                'End Select
                        End Select
                End Select
            ElseIf pVal.BeforeAction = True Then
                Select Case pVal.MenuUID
                    Case "1282", "1281", "1292", "1293", "1287", "519", "1284", "1286", "5890", "1290", "1289", "1288", "1291"
                        Select Case oForm.TypeEx
                            Case ProductionPlanningFormId
                                oProductionPlan.MenuEvent(pVal, BubbleEvent)
                            Case MonthPlanningFormId
                                oProductionMthPlan.MenuEvent(pVal, BubbleEvent)
                        End Select
                End Select
            End If

            'If pVal.MenuUID = "526" Then
            '    oCompany.Disconnect()
            '    oApplication.StatusBar.SetText(addonName & " AddOn is Disconnected . . .", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            '    End
            'End If
        Catch ex As Exception
            oApplication.StatusBar.SetText("Application Menu Event Failed : " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
        End Try
    End Sub
#End Region

#Region " ... 2) Item Event ..."

    Private Sub oApplication_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles oApplication.ItemEvent
        Try
            'If pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD Then Exit Sub
            'If pVal.InnerEvent = True Then Exit Sub
            Select Case pVal.FormTypeEx 'pVal.FormUID
                Case ProductionPlanningFormId
                    oProductionPlan.ItemEvent(FormUID, pVal, BubbleEvent)
                Case MonthPlanningFormId
                    oProductionMthPlan.ItemEvent(FormUID, pVal, BubbleEvent)
                Case PRORecptTypeEx
                    oReceiptFromPro.ItemEvent(FormUID, pVal, BubbleEvent)
            End Select
            'Select Case pVal.FormTypeEx
            '    'Case PROTypeEx
            '    '    oProductionOrder.ItemEvent(FormUID, pVal, BubbleEvent)
            '    Case PRORecptTypeEx
            '        oReceiptFromPro.ItemEvent(FormUID, pVal, BubbleEvent)
            'End Select
        Catch ex As Exception
            oApplication.StatusBar.SetText("Application ItemEvent Failed : " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
        End Try
    End Sub

#End Region

#Region " ... 3) FormDataEvent ..."

    Private Sub oApplication_FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean) Handles oApplication.FormDataEvent
        Try
            Select Case BusinessObjectInfo.FormTypeEx 'FormUID
                Case ProductionPlanningFormId
                    oProductionPlan.FormDataEvent(BusinessObjectInfo, BubbleEvent)
                Case MonthPlanningFormId
                    oProductionMthPlan.FormDataEvent(BusinessObjectInfo, BubbleEvent)
            End Select

            'Select Case BusinessObjectInfo.FormTypeEx
            '    Case PROTypeEx
            '        oProductionOrder.FormDataEvent(BusinessObjectInfo, BubbleEvent)
            '    Case PRORecptTypeEx
            '        oReceiptFromPro.FormDataEvent(BusinessObjectInfo, BubbleEvent)
            'End Select
        Catch ex As Exception
            oApplication.StatusBar.SetText("Application FormDataEvent Failed : " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
        End Try

    End Sub
#End Region

#Region " ... 4) Status Bar Event ..."
    Public Sub oApplication_StatusBarEvent(ByVal Text As String, ByVal MessageType As SAPbouiCOM.BoStatusBarMessageType) Handles oApplication.StatusBarEvent
        Try
            If MessageType = SAPbouiCOM.BoStatusBarMessageType.smt_Warning Or MessageType = SAPbouiCOM.BoStatusBarMessageType.smt_Error Then
                System.Media.SystemSounds.Asterisk.Play()
            End If
        Catch ex As Exception
            oApplication.StatusBar.SetText(addonName & " StatusBarEvent Event Failed : " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
        End Try
    End Sub
#End Region

#Region " ... 5) Set Event Filter ..."
    Public Sub SetEventFilter()
        Try
            Dim oFilters As SAPbouiCOM.EventFilters
            Dim oFilter As SAPbouiCOM.EventFilter

            'oFilters = New SAPbouiCOM.EventFilters
            'oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST)
            'oFilter.AddEx("PPLN") 'Daily Planning Form
            'oApplication.SetFilter(oFilters)
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        Finally
        End Try
    End Sub
#End Region

#Region " ... 6) Right Click Event ..."

    Private Sub oApplication_RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean) Handles oApplication.RightClickEvent
        Try
            If eventInfo.BeforeAction Then
                Select Case oApplication.Forms.ActiveForm.TypeEx
                    Case "PPLN"
                        oForm = oApplication.Forms.ActiveForm
                        Dim objMatrix As SAPbouiCOM.Matrix
                        objMatrix = oForm.Items.Item("Matrix1").Specific
                        If eventInfo.ItemUID = "" And oForm.Items.Item("CmbSt").Specific.selected.value = "O" Then
                            oForm.EnableMenu("1284", True) 'Cancel
                            oForm.EnableMenu("1286", True) 'Close
                        Else
                            oForm.EnableMenu("1284", False) 'Cancel
                            oForm.EnableMenu("1286", False) 'Close
                        End If
                        Try
                            If eventInfo.ItemUID = "" Then Exit Try
                            If oForm.Items.Item(eventInfo.ItemUID).Specific.String <> "" Then
                                oForm.EnableMenu("772", True)  'Copy
                            ElseIf oForm.Items.Item(eventInfo.ItemUID).Specific.String = "" Then
                                oForm.EnableMenu("773", True)  'Paste
                            End If
                        Catch ex As Exception
                            objMatrix = oForm.Items.Item(eventInfo.ItemUID).Specific
                            If eventInfo.Row <= 0 Then If oForm.Mode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then oForm.EnableMenu("772", True) : oForm.EnableMenu("784", True) : Exit Try
                            If objMatrix.Columns.Item(eventInfo.ColUID).Cells.Item(eventInfo.Row).Specific.String <> "" Then
                                oForm.EnableMenu("772", True)  'Copy
                            ElseIf objMatrix.Columns.Item(eventInfo.ColUID).Cells.Item(eventInfo.Row).Specific.String = "" Then
                                oForm.EnableMenu("773", True)  'Paste
                            End If
                            If eventInfo.ItemUID = "Matrix1" And objMatrix.Columns.Item("ProEntry").Cells.Item(eventInfo.Row).Specific.String <> "" Then
                                oForm.EnableMenu("1293", False)
                            Else
                                oForm.EnableMenu("1293", True)
                            End If
                            If eventInfo.ItemUID = "Matrix1" And objMatrix.Columns.Item("ItemC").Cells.Item(objMatrix.VisualRowCount).Specific.String <> "" Then
                                oForm.EnableMenu("1293", True)
                            End If
                        End Try
                        'Try
                        '    objMatrix = oForm.Items.Item(eventInfo.ItemUID).Specific
                        '    If objMatrix.Item.Type = SAPbouiCOM.BoFormItemTypes.it_MATRIX Then
                        '        If eventInfo.Row = 0 Then oForm.EnableMenu("784", True) : Exit Try
                        '        If objMatrix.Columns.Item(eventInfo.ColUID).Cells.Item(eventInfo.Row).Specific.String <> "" Then
                        '            oForm.EnableMenu("772", True)  'Copy
                        '        Else
                        '            oForm.EnableMenu("772", False)
                        '        End If
                        '    End If
                        '    If eventInfo.ItemUID = "Matrix1" And objMatrix.Columns.Item("ProEntry").Cells.Item(eventInfo.Row).Specific.String <> "" Then
                        '        oForm.EnableMenu("1293", False)
                        '    Else
                        '        oForm.EnableMenu("1293", True)
                        '    End If
                        'Catch ex As Exception
                        '    If oForm.Items.Item(eventInfo.ItemUID).Specific.String <> "" Then
                        '        oForm.EnableMenu("772", True)  'Copy
                        '    Else
                        '        oForm.EnableMenu("772", False)
                        '    End If
                        'End Try

                End Select
            Else
                oForm.EnableMenu("772", False)
                oForm.EnableMenu("1293", False)
                oForm.EnableMenu("1284", False) 'Cancel
                oForm.EnableMenu("1286", False) 'Close
                oForm.EnableMenu("1293", False)
            End If


        Catch ex As Exception
            'oApplication.StatusBar.SetText(addonName & " : Right Click Event Failed : " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
        End Try
    End Sub


#End Region

    '#Region " ... 7) Application Event ..."
    '    Private Sub oApplication_AppEvent(ByVal EventType As SAPbouiCOM.BoAppEventTypes) Handles oApplication.AppEvent
    '        Try
    '            Select Case EventType
    '                Case SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged, SAPbouiCOM.BoAppEventTypes.aet_LanguageChanged, SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition, SAPbouiCOM.BoAppEventTypes.aet_ShutDown
    '                    System.Windows.Forms.Application.Exit()
    '            End Select

    '        Catch ex As Exception
    '            oApplication.StatusBar.SetText("Application Event Failed: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
    '        Finally
    '        End Try

    '    End Sub
    '#End Region

#Region " App Event " 'added by Tamizh 08-Jan-2021
    Private Sub oApplication_AppEvent(ByVal EventType As SAPbouiCOM.BoAppEventTypes) Handles oApplication.AppEvent
        If EventType = SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged Or EventType = SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition Or EventType = SAPbouiCOM.BoAppEventTypes.aet_ShutDown Then
            Try
                'objUIXml.LoadMenuXML("RemoveMenu.xml", Mukesh.SBOLib.UIXML.enuResourceType.Embeded)
                'If oCompany.Connected Then oCompany.Disconnect()
                'oCompany = Nothing
                'oApplication = Nothing
                'System.Runtime.InteropServices.Marshal.ReleaseComObject(oCompany)
                'System.Runtime.InteropServices.Marshal.ReleaseComObject(oApplication)
                'GC.Collect()

                If oCompany.Connected Then oCompany.Disconnect()
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oCompany)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oApplication)
                oCompany = Nothing
                'oApplication = Nothing
                GC.Collect()
                System.Windows.Forms.Application.Exit()
                End
            Catch ex As Exception
            End Try
            End
        End If
    End Sub

#End Region

#Region "...8) Layout Event ..."

    Private Sub objApplication_LayoutKeyEvent(ByRef eventInfo As SAPbouiCOM.LayoutKeyInfo, ByRef BubbleEvent As Boolean) Handles oApplication.LayoutKeyEvent

        'BubbleEvent = True
        'If eventInfo.BeforeAction = True Then
        If eventInfo.FormUID.Contains(ProductionPlanningFormId) Then
            oProductionPlan.LayoutKeyEvent(eventInfo, BubbleEvent)
        End If
        'End If
    End Sub
#End Region


End Module
