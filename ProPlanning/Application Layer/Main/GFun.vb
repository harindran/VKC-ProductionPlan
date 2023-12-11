Imports System
Imports System.Reflection
Imports System.IO
Imports SAPbouiCOM.Framework
''' <summary>
''' Globally whatever Function and method do you want define here 
''' We can use any class and module from here  
''' </summary>
''' <remarks></remarks>
''' 

Public Class GFun
    Inherits GVariables

    ' Public WithEvents oApplication As SAPbouiCOM.Application
    Private Shared intTotalFormCount As Integer = 0
    Sub New(ByVal addon_Name As String)
        addonName = addon_Name

    End Sub

#Region " ...  Connect to SAP ..."

    Public Sub SetApplication()
        Try
            Dim oGUI As New SAPbouiCOM.SboGuiApi
            oGUI.AddonIdentifier = ""
            oGUI.Connect(ConnectionString)
            oApplication = oGUI.GetApplication()
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        Finally
        End Try
    End Sub

    Public Function CookieConnect() As Integer
        Try
            Dim strCkie, strContext As String
            oCompany = New SAPbobsCOM.Company
            Debug.Print(oCompany.CompanyDB)
            strCkie = oCompany.GetContextCookie()
            strContext = oApplication.Company.GetConnectionContext(strCkie)
            CookieConnect = oCompany.SetSboLoginContext(strContext)
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
            Return -1
        Finally
        End Try
    End Function

    Public Function ConnectionContext() As Integer
        Try
            Dim strErrorCode As String
            If oCompany.Connected = True Then oCompany.Disconnect()

            oApplication.StatusBar.SetText("Connecting The " & addonName & " Addon With The Company..........      Please Wait ..........", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            strErrorCode = oCompany.Connect
            ConnectionContext = strErrorCode
            If strErrorCode = 0 Then
                If isValidLicense() Then
                    oApplication.StatusBar.SetText("ADDON for " & addonName & " Module - Connecion Established  !!! ", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                    System.Media.SystemSounds.Asterisk.Play()
                    'AddLogo()
                    Return 0

                Else
                    oApplication.StatusBar.SetText("Failed To Connect, Please Check The License Configuration....." & oCompany.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return -1
                End If

            End If
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
            Return -1
        Finally
        End Try
    End Function

    Public Sub Intialize(ByVal args() As String)
        Try
            Dim objapplication As Application
            If (args.Length < 1) Then objapplication = New Application Else objapplication = New Application(args(0))
            oApplication = Application.SBO_Application

            oGFun.HWKEY = HWKEY
            If isValidLicense() Then
                oApplication.StatusBar.SetText("Connecting The " & addonName & " Addon With The Company..........      Please Wait ..........", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                'objApplication.StatusBar.SetText("Establishing Company Connection Please Wait...", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                oCompany = Application.SBO_Application.Company.GetDICompany()
                'oApplication = oGFun.oApplication
                'oCompany = oGFun.oCompany
                Try
                    Dim oTableCreation As New TableCreation
                    EventHandler.SetEventFilter()
                    oApplication.StatusBar.SetText(" Menu Creation Starting........", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    'oGFun.AddXML("Menu.xml")
                    CreateMenu("", 3, "Daily Plan", SAPbouiCOM.BoMenuType.mt_STRING, "PPLN", oApplication.Menus.Item("4352"))
                    CreateMenu("", 4, "Monthly Plan", SAPbouiCOM.BoMenuType.mt_STRING, "MPLN", oApplication.Menus.Item("4352"))

                    oApplication.StatusBar.SetText("Menu Creation Success.......", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

                Catch ex As Exception
                    oApplication.MessageBox(ex.Message)
                    End
                End Try
                oProductionPlan.addReporttype()
                oApplication.StatusBar.SetText("ADDON for " & addonName & " Module - Connection Established  !!! ", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                'objApplication.StatusBar.SetText("Addon Connected Successfully..!!!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                objapplication.Run()
            Else
                oApplication.StatusBar.SetText("Addon Disconnected due to license mismatch..!!!", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If

            'System.Windows.Forms.Application.Run()
        Catch ex As Exception
            oApplication.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
        End Try
    End Sub

    Private Function CreateMenu(ByVal ImagePath As String, ByVal Position As Int32, ByVal DisplayName As String, ByVal MenuType As SAPbouiCOM.BoMenuType, ByVal UniqueID As String, ByVal ParentMenu As SAPbouiCOM.MenuItem) As SAPbouiCOM.MenuItem
        Try
            Dim oMenuPackage As SAPbouiCOM.MenuCreationParams
            oMenuPackage = oApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
            oMenuPackage.Image = ImagePath
            oMenuPackage.Position = Position
            oMenuPackage.Type = MenuType
            oMenuPackage.UniqueID = UniqueID
            oMenuPackage.String = DisplayName
            ParentMenu.SubMenus.AddEx(oMenuPackage)
        Catch ex As Exception
            oApplication.StatusBar.SetText("Menu Already Exists", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_None)
        End Try
        Return ParentMenu.SubMenus.Item(UniqueID)
    End Function


    Function isValidLicense() As Boolean
        Try
            oApplication.Menus.Item("257").Activate()
            Dim CrrHWKEY As String = oApplication.Forms.ActiveForm.Items.Item("79").Specific.Value.ToString.Trim
            oApplication.Forms.ActiveForm.Close()

            For i As Integer = 0 To HWKEY.Length - 1
                If HWKEY(i).Trim = CrrHWKEY.Trim Then
                    Return True
                End If
            Next

            MsgBox("Installing Add-On failed due to License mismatch", MsgBoxStyle.OkOnly, "License Management")
            ' End
            Return False

            'Dim HKEY As String = ""
            'Dim LicInf As SBOLICENSELib.LicenseInfo = New SBOLICENSELib.LicenseInfo()
            'Dim ww = SBOLICENSELib.ILicenseInfo.GetHardwareKey(HKEY)
            'SBOLICENSELib.LicenseInfoClass.GetHardwareKey(HKEY)
            'ILicenseInfo.
            'LicInf.GetHardwareKey(HKEY)
            'If HKEY.Trim.Equals("X0736535264") = False Then   ' Suriya LapTop Hardware Key
            ' P2023351072 me
            'T2129002141 Suriya Server
            'If HKEY.Trim.Equals("P2023351072") = False Then ' Novateur LapTop Hardware Key
            '    If oCompany.Connected Then oCompany.Disconnect()
            '    System.Windows.Forms.Application.Exit()
            '    MsgBox("Installing Add-On failed due to License mismatch")
            '    End
            'End If
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
            Return True
        End Try
    End Function

    Sub AddLogo()

        'Logo - Changed by Sankaralakshmi
        Try
            'Path
            Dim oItem As SAPbouiCOM.Item
            Dim frm As SAPbouiCOM.Form
            frm = oApplication.Forms.GetFormByTypeAndCount("169", 1)
            frm.Height = frm.Height + 75
            Try
                frm.Items.Add("p_Logo", SAPbouiCOM.BoFormItemTypes.it_PICTURE)
            Catch ex As Exception
            End Try
            oItem = frm.Items.Item("p_Logo")
            oItem.Top = frm.Items.Item("6").Top + frm.Items.Item("6").Width + 125
            oItem.Left = frm.Items.Item("6").Left
            oItem.Width = frm.Items.Item("6").Width - 20
            oItem.LinkTo = "6"
            oItem.Height = 40
            Dim opic As SAPbouiCOM.PictureBox
            opic = oItem.Specific
            opic.Picture = System.Windows.Forms.Application.StartupPath & "\Logo.bmp"
        Catch ex As Exception
            ' oApplication.StatusBar.SetText("Add Logo Method Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End Try

        'Try
        '    Try
        '        oApplication.Forms.Item("Logo").Close()
        '    Catch ex As Exception
        '    End Try
        '    Dim frm As SAPbouiCOM.Form
        '    LoadXML(frm, "Logo", "Logo.xml")
        '    frm = oApplication.Forms.Item("Logo")
        '    frm.Left = 2000
        '    frm.Top = 0
        '    Dim ptr As SAPbouiCOM.PictureBox
        '    ptr = frm.Items.Item("1").Specific
        '    ptr.Picture = System.Windows.Forms.Application.StartupPath & "\Logo.bmp"

        'Catch ex As Exception
        '    oApplication.StatusBar.SetText("Add Logo Method Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        'Finally
        'End Try
    End Sub


    Sub AddXML(ByVal pathstr As String)
        Try
            Dim xmldoc As New Xml.XmlDocument

            'Dim stackTrace As New Diagnostics.StackFrame(0)
            'Dim ss = stackTrace.GetMethod.Name

            Dim asm As Assembly = Assembly.GetExecutingAssembly()
            Dim location As String = asm.FullName
            Dim appName As String = System.IO.Path.GetDirectoryName(location)
            Dim stream As System.IO.Stream
            Dim s = System.Reflection.Assembly.GetCallingAssembly.GetName().Name + "." + pathstr
            Try
                stream = System.Reflection.Assembly.GetCallingAssembly().GetManifestResourceStream(System.Reflection.Assembly.GetCallingAssembly.GetName().Name + "." + pathstr)
                Dim tempstreamreader As New System.IO.StreamReader(stream, True)
            Catch ex As Exception
                stream = System.Reflection.Assembly.GetEntryAssembly().GetManifestResourceStream(System.Reflection.Assembly.GetEntryAssembly.GetName().Name + "." + pathstr)
            End Try

            Dim streamreader As New System.IO.StreamReader(stream, True)
            xmldoc.LoadXml(streamreader.ReadToEnd())
            streamreader.Close()
            oApplication.LoadBatchActions(xmldoc.InnerXml)
        Catch ex As Exception
            oApplication.StatusBar.SetText("AddXML Method Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
        End Try
    End Sub

#End Region

#Region "       Common For Data Base Creation ...   "

    Public Function UDOExists(ByVal code As String) As Boolean
        GC.Collect()
        Dim v_UDOMD As SAPbobsCOM.UserObjectsMD
        Dim v_ReturnCode As Boolean
        v_UDOMD = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)
        v_ReturnCode = v_UDOMD.GetByKey(code)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(v_UDOMD)
        v_UDOMD = Nothing
        Return v_ReturnCode
    End Function

    Function CreateTable(ByVal TableName As String, ByVal TableDesc As String, ByVal TableType As SAPbobsCOM.BoUTBTableType) As Boolean
        CreateTable = False
        Dim v_RetVal As Long
        Dim v_ErrCode As Long
        Dim v_ErrMsg As String = ""
        Try
            If Not Me.TableExists(TableName) Then
                Dim v_UserTableMD As SAPbobsCOM.UserTablesMD
                oApplication.StatusBar.SetText("Creating Table " + TableName + " ...................", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                v_UserTableMD = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables)
                v_UserTableMD.TableName = TableName
                v_UserTableMD.TableDescription = TableDesc
                v_UserTableMD.TableType = TableType
                v_RetVal = v_UserTableMD.Add()
                If v_RetVal <> 0 Then
                    oCompany.GetLastError(v_ErrCode, v_ErrMsg)
                    oApplication.StatusBar.SetText("Failed to Create Table " & TableDesc & v_ErrCode & " " & v_ErrMsg, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(v_UserTableMD)
                    v_UserTableMD = Nothing
                    Return False
                Else
                    oApplication.StatusBar.SetText("[" & TableName & "] - " & TableDesc & " Created Successfully!!!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(v_UserTableMD)
                    v_UserTableMD = Nothing
                    Return True
                End If
            Else
                GC.Collect()
                Return False
            End If
        Catch ex As Exception
            oApplication.StatusBar.SetText(addonName & ":> " & ex.Message & " @ " & ex.Source)
        End Try
    End Function

    Function ColumnExists(ByVal TableName As String, ByVal FieldID As String) As Boolean
        Try
            Dim rs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim oFlag As Boolean = True
            rs.DoQuery("Select 1 from [CUFD] Where TableID='" & Trim(TableName) & "' and AliasID='" & Trim(FieldID) & "'")
            If rs.EoF Then oFlag = False
            System.Runtime.InteropServices.Marshal.ReleaseComObject(rs)
            rs = Nothing
            GC.Collect()
            Return oFlag
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
            Return False
        End Try
    End Function

    Function UDFExists(ByVal TableName As String, ByVal FieldID As String) As Boolean
        Try
            Dim rs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim oFlag As Boolean = True
            Dim aa = "Select 1 from [CUFD] Where TableID='" & Trim(TableName) & "' and AliasID='" & Trim(FieldID) & "'"
            rs.DoQuery("Select 1 from [CUFD] Where TableID='" & Trim(TableName) & "' and AliasID='" & Trim(FieldID) & "'")
            If rs.EoF Then oFlag = False
            System.Runtime.InteropServices.Marshal.ReleaseComObject(rs)
            rs = Nothing
            GC.Collect()
            Return oFlag
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
            Return False
        End Try
    End Function

    Function TableExists(ByVal TableName As String) As Boolean
        Dim oTables As SAPbobsCOM.UserTablesMD
        Dim oFlag As Boolean
        oTables = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables)
        oFlag = oTables.GetByKey(TableName)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(oTables)
        Return oFlag
    End Function

    Function CreateUserFieldsComboBox(ByVal TableName As String, ByVal FieldName As String, ByVal FieldDescription As String, ByVal type As SAPbobsCOM.BoFieldTypes, Optional ByVal size As Long = 0, Optional ByVal subType As SAPbobsCOM.BoFldSubTypes = SAPbobsCOM.BoFldSubTypes.st_None, Optional ByVal LinkedTable As String = "", Optional ByVal ComboValidValues As String(,) = Nothing, Optional ByVal DefaultValidValues As String = "") As Boolean
        Try
            'If TableName.StartsWith("@") = False Then
            If Not Me.UDFExists(TableName, FieldName) Then
                Dim v_UserField As SAPbobsCOM.UserFieldsMD
                v_UserField = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)
                v_UserField.TableName = TableName
                v_UserField.Name = FieldName
                v_UserField.Description = FieldDescription
                v_UserField.Type = type
                If type <> SAPbobsCOM.BoFieldTypes.db_Date Then
                    If size <> 0 Then
                        v_UserField.Size = size
                    End If
                End If
                If subType <> SAPbobsCOM.BoFldSubTypes.st_None Then
                    v_UserField.SubType = subType
                End If

                For i As Int16 = 0 To ComboValidValues.GetLength(0) - 1
                    If i > 0 Then v_UserField.ValidValues.Add()
                    v_UserField.ValidValues.Value = ComboValidValues(i, 0)
                    v_UserField.ValidValues.Description = ComboValidValues(i, 1)
                Next
                If DefaultValidValues <> "" Then v_UserField.DefaultValue = DefaultValidValues

                If LinkedTable <> "" Then v_UserField.LinkedTable = LinkedTable
                v_RetVal = v_UserField.Add()
                If v_RetVal <> 0 Then
                    oCompany.GetLastError(v_ErrCode, v_ErrMsg)
                    oApplication.StatusBar.SetText("Failed to add UserField " & FieldDescription & " - " & v_ErrCode & " " & v_ErrMsg, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(v_UserField)
                    v_UserField = Nothing
                    Return False
                Else
                    oApplication.StatusBar.SetText(" & TableName & - " & FieldDescription & " added successfully!!!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(v_UserField)
                    v_UserField = Nothing
                    Return True
                End If

            Else
                Return False
            End If
            ' End If
        Catch ex As Exception
            oApplication.MessageBox(ex.Message)
            Return False
        End Try
    End Function

    Function CreateUserFields(ByVal TableName As String, ByVal FieldName As String, ByVal FieldDescription As String, ByVal type As SAPbobsCOM.BoFieldTypes, Optional ByVal size As Long = 0, Optional ByVal subType As SAPbobsCOM.BoFldSubTypes = SAPbobsCOM.BoFldSubTypes.st_None, Optional ByVal LinkedTable As String = "", Optional ByVal DefaultValue As String = "") As Boolean
        Try
            If TableName.StartsWith("@") = True Then
                If Not Me.ColumnExists(TableName, FieldName) Then
                    Dim v_UserField As SAPbobsCOM.UserFieldsMD
                    v_UserField = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)
                    v_UserField.TableName = TableName
                    v_UserField.Name = FieldName
                    v_UserField.Description = FieldDescription
                    v_UserField.Type = type
                    If type <> SAPbobsCOM.BoFieldTypes.db_Date Then
                        If size <> 0 Then
                            v_UserField.Size = size
                        End If
                    End If
                    If subType <> SAPbobsCOM.BoFldSubTypes.st_None Then
                        v_UserField.SubType = subType
                    End If
                    If LinkedTable <> "" Then v_UserField.LinkedTable = LinkedTable
                    If DefaultValue <> "" Then v_UserField.DefaultValue = DefaultValue

                    v_RetVal = v_UserField.Add()
                    If v_RetVal <> 0 Then
                        oCompany.GetLastError(v_ErrCode, v_ErrMsg)
                        oApplication.StatusBar.SetText("Failed to add UserField masterid" & v_ErrCode & " " & v_ErrMsg, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(v_UserField)
                        v_UserField = Nothing
                        Return False
                    Else
                        oApplication.StatusBar.SetText("[" & TableName & "] - " & FieldDescription & " added successfully!!!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(v_UserField)
                        v_UserField = Nothing
                        Return True
                    End If
                Else
                    Return False
                End If
            End If

            If TableName.StartsWith("@") = False Then
                If Not Me.UDFExists(TableName, FieldName) Then
                    Dim v_UserField As SAPbobsCOM.UserFieldsMD = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)
                    v_UserField.TableName = TableName
                    v_UserField.Name = FieldName
                    v_UserField.Description = FieldDescription
                    v_UserField.Type = type
                    If type <> SAPbobsCOM.BoFieldTypes.db_Date Then
                        If size <> 0 Then
                            v_UserField.Size = size
                        End If
                    End If
                    If subType <> SAPbobsCOM.BoFldSubTypes.st_None Then
                        v_UserField.SubType = subType
                    End If
                    If LinkedTable <> "" Then v_UserField.LinkedTable = LinkedTable
                    v_RetVal = v_UserField.Add()
                    If v_RetVal <> 0 Then
                        oCompany.GetLastError(v_ErrCode, v_ErrMsg)
                        oApplication.StatusBar.SetText("Failed to add UserField " & FieldDescription & " - " & v_ErrCode & " " & v_ErrMsg, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(v_UserField)
                        v_UserField = Nothing
                        Return False
                    Else
                        oApplication.StatusBar.SetText(" & TableName & - " & FieldDescription & " added successfully!!!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(v_UserField)
                        v_UserField = Nothing
                        Return True
                    End If

                Else
                    Return False
                End If
            End If
        Catch ex As Exception
            oApplication.MessageBox(ex.Message)
            Return False
        End Try
        Return True
    End Function

    Function RegisterUDO(ByVal UDOCode As String, ByVal UDOName As String, ByVal UDOType As SAPbobsCOM.BoUDOObjType, ByVal FindField As String(,), ByVal UDOHTableName As String, Optional ByVal UDODTableName As String = "", Optional ByVal ChildTable As String = "", Optional ByVal ChildTable1 As String = "", _
    Optional ByVal ChildTable2 As String = "", Optional ByVal ChildTable3 As String = "", Optional ByVal ChildTable4 As String = "", Optional ByVal ChildTable5 As String = "", _
    Optional ByVal ChildTable6 As String = "", Optional ByVal ChildTable7 As String = "", Optional ByVal ChildTable8 As String = "", Optional ByVal ChildTable9 As String = "", _
    Optional ByVal LogOption As SAPbobsCOM.BoYesNoEnum = SAPbobsCOM.BoYesNoEnum.tNO) As Boolean
        Dim ActionSuccess As Boolean = False
        Try
            RegisterUDO = False
            Dim v_udoMD As SAPbobsCOM.UserObjectsMD
            v_udoMD = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)
            v_udoMD.CanCancel = SAPbobsCOM.BoYesNoEnum.tYES
            v_udoMD.CanClose = SAPbobsCOM.BoYesNoEnum.tYES
            v_udoMD.CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tNO
            v_udoMD.CanDelete = SAPbobsCOM.BoYesNoEnum.tYES
            v_udoMD.CanFind = SAPbobsCOM.BoYesNoEnum.tYES
            v_udoMD.CanLog = SAPbobsCOM.BoYesNoEnum.tNO
            v_udoMD.CanYearTransfer = SAPbobsCOM.BoYesNoEnum.tYES
            v_udoMD.ManageSeries = SAPbobsCOM.BoYesNoEnum.tYES
            v_udoMD.Code = UDOCode
            v_udoMD.Name = UDOName
            v_udoMD.TableName = UDOHTableName

            If UDODTableName <> "" Then
                v_udoMD.ChildTables.TableName = UDODTableName
                v_udoMD.ChildTables.Add()
            End If

            If ChildTable <> "" Then
                v_udoMD.ChildTables.TableName = ChildTable
                v_udoMD.ChildTables.Add()
            End If
            If ChildTable1 <> "" Then
                v_udoMD.ChildTables.TableName = ChildTable1
                v_udoMD.ChildTables.Add()
            End If
            If ChildTable2 <> "" Then
                v_udoMD.ChildTables.TableName = ChildTable2
                v_udoMD.ChildTables.Add()
            End If
            If ChildTable3 <> "" Then
                v_udoMD.ChildTables.TableName = ChildTable3
                v_udoMD.ChildTables.Add()
            End If
            If ChildTable4 <> "" Then
                v_udoMD.ChildTables.TableName = ChildTable4
                v_udoMD.ChildTables.Add()
            End If
            If ChildTable5 <> "" Then
                v_udoMD.ChildTables.TableName = ChildTable5
                v_udoMD.ChildTables.Add()
            End If
            If ChildTable6 <> "" Then
                v_udoMD.ChildTables.TableName = ChildTable6
                v_udoMD.ChildTables.Add()
            End If
            If ChildTable7 <> "" Then
                v_udoMD.ChildTables.TableName = ChildTable7
                v_udoMD.ChildTables.Add()
            End If
            If ChildTable8 <> "" Then
                v_udoMD.ChildTables.TableName = ChildTable8
                v_udoMD.ChildTables.Add()
            End If
            If ChildTable9 <> "" Then
                v_udoMD.ChildTables.TableName = ChildTable9
                v_udoMD.ChildTables.Add()
            End If

            If LogOption = SAPbobsCOM.BoYesNoEnum.tYES Then
                v_udoMD.CanLog = SAPbobsCOM.BoYesNoEnum.tYES
                v_udoMD.LogTableName = "A" & UDOHTableName
            End If
            v_udoMD.ObjectType = UDOType
            For i As Int16 = 0 To FindField.GetLength(0) - 1
                If i > 0 Then v_udoMD.FindColumns.Add()
                v_udoMD.FindColumns.ColumnAlias = FindField(i, 0)
                v_udoMD.FindColumns.ColumnDescription = FindField(i, 1)
            Next

            If v_udoMD.Add() = 0 Then
                RegisterUDO = True
                If oCompany.InTransaction Then oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                oApplication.StatusBar.SetText("Successfully Registered UDO >" & UDOCode & ">" & UDOName & " >" & oCompany.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            Else
                oApplication.StatusBar.SetText("Failed to Register UDO >" & UDOCode & ">" & UDOName & " >" & oCompany.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                MessageBox.Show(oCompany.GetLastErrorDescription)
                RegisterUDO = False
            End If
            System.Runtime.InteropServices.Marshal.ReleaseComObject(v_udoMD)
            v_udoMD = Nothing
            GC.Collect()
            If ActionSuccess = False And oCompany.InTransaction Then oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
        Catch ex As Exception
            If oCompany.InTransaction Then oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            Return False
        End Try
    End Function

    Function RegisterUDOForDefaultForm(ByVal UDOCode As String, ByVal UDOName As String, ByVal UDOType As SAPbobsCOM.BoUDOObjType, ByVal FindField As String(,), ByVal UDOHTableName As String, Optional ByVal UDODTableName As String = "", Optional ByVal ChildTable As String = "", _
   Optional ByVal LogOption As SAPbobsCOM.BoYesNoEnum = SAPbobsCOM.BoYesNoEnum.tNO) As Boolean
        Dim ActionSuccess As Boolean = False
        Try
            RegisterUDOForDefaultForm = False
            Dim v_udoMD As SAPbobsCOM.UserObjectsMD
            v_udoMD = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)
            v_udoMD.CanCancel = SAPbobsCOM.BoYesNoEnum.tYES
            v_udoMD.CanClose = SAPbobsCOM.BoYesNoEnum.tYES
            v_udoMD.CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tYES
            v_udoMD.CanDelete = SAPbobsCOM.BoYesNoEnum.tYES
            v_udoMD.CanFind = SAPbobsCOM.BoYesNoEnum.tYES
            v_udoMD.CanLog = SAPbobsCOM.BoYesNoEnum.tNO
            v_udoMD.CanYearTransfer = SAPbobsCOM.BoYesNoEnum.tYES
            v_udoMD.ManageSeries = SAPbobsCOM.BoYesNoEnum.tYES
            v_udoMD.Code = UDOCode
            v_udoMD.Name = UDOName
            v_udoMD.TableName = UDOHTableName

            If UDODTableName <> "" Then
                v_udoMD.ChildTables.TableName = UDODTableName
                v_udoMD.ChildTables.Add()
            End If

            If ChildTable <> "" Then
                v_udoMD.ChildTables.TableName = ChildTable
                v_udoMD.ChildTables.Add()
            End If
            If LogOption = SAPbobsCOM.BoYesNoEnum.tYES Then
                v_udoMD.CanLog = SAPbobsCOM.BoYesNoEnum.tYES
                v_udoMD.LogTableName = "A" & UDOHTableName
            End If
            v_udoMD.ObjectType = UDOType
            For i As Int16 = 0 To FindField.GetLength(0) - 1
                If i > 0 Then v_udoMD.FindColumns.Add()
                v_udoMD.FindColumns.ColumnAlias = FindField(i, 0)
                v_udoMD.FindColumns.ColumnDescription = FindField(i, 1)
            Next
            For i As Int16 = 0 To FindField.GetLength(0) - 1
                If i > 0 Then v_udoMD.FormColumns.Add()
                v_udoMD.FormColumns.FormColumnAlias = FindField(i, 0)
                v_udoMD.FormColumns.FormColumnDescription = FindField(i, 1)
            Next

            If v_udoMD.Add() = 0 Then
                RegisterUDOForDefaultForm = True
                If oCompany.InTransaction Then oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                oApplication.StatusBar.SetText("Successfully Registered UDO >" & UDOCode & ">" & UDOName & " >" & oCompany.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            Else
                oApplication.StatusBar.SetText("Failed to Register UDO >" & UDOCode & ">" & UDOName & " >" & oCompany.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                'MessageBox.Show(oCompany.GetLastErrorDescription)
                RegisterUDOForDefaultForm = False
            End If
            System.Runtime.InteropServices.Marshal.ReleaseComObject(v_udoMD)
            v_udoMD = Nothing
            GC.Collect()
            If ActionSuccess = False And oCompany.InTransaction Then oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
        Catch ex As Exception
            If oCompany.InTransaction Then oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            Return False
        End Try
    End Function

#End Region

#Region " ...  Common Function for DB ..."

    Function getSingleValue(ByVal TblName As String, ByVal ValFldNa As String, ByVal Conditions As String) As String
        Try
            Dim rset As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim strReturnVal As String = ""
            Dim strQuery = "SELECT " & ValFldNa & " FROM " & TblName & IIf(Conditions.Trim() = "", "", " WHERE ") & Conditions
            rset.DoQuery(strQuery)
            Return IIf(rset.RecordCount > 0, rset.Fields.Item(0).Value.ToString(), "")
        Catch ex As Exception
            oApplication.StatusBar.SetText(" Get Single Value Function Failed : ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Return ""
        End Try
    End Function

    Function getSingleValue(ByVal strSQL As String) As String
        Try
            Dim rset As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim strReturnVal As String = ""
            rset.DoQuery(strSQL)
            Return IIf(rset.RecordCount > 0, rset.Fields.Item(0).Value.ToString(), "")
        Catch ex As Exception
            oApplication.StatusBar.SetText(" Get Single Value Function Failed :  " & ex.Message + strSQL, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Return ""
        End Try
    End Function

    Function DoQuery(ByVal strSql As String) As SAPbobsCOM.Recordset
        Try
            Dim rsetCode As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            rsetCode.DoQuery(strSql)

            Return rsetCode
        Catch ex As Exception
            oApplication.StatusBar.SetText("Execute Query Function Failed:" & ex.Message + strSql, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Return Nothing
        Finally
        End Try
    End Function

    Function DoExQuery(ByVal strSql As String) As Boolean
        Try
            Dim rsetCode As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            rsetCode.DoQuery(strSql)
            Return True
        Catch ex As Exception
            Throw ex
            oApplication.StatusBar.SetText("Execute Query Function Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Return False
        Finally
        End Try
    End Function
    Function isDuplicate(ByVal oEditText As SAPbouiCOM.EditText, ByVal strTableName As String, ByVal strFildName As String, ByVal strMessage As String) As Boolean
        Try
            Dim rsetPayMethod As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim blReturnVal As Boolean = False
            Dim strQuery As String
            '  If oEditText.Value.Equals("") Then

            ' oApplication.StatusBar.SetText(strMessage & " : Should Not Be left Empty", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            ' Return False
            '  End If

            strQuery = "SELECT * FROM " & strTableName & " WHERE UPPER(" & strFildName & ")=UPPER('" & oEditText.Value & "')"
            rsetPayMethod.DoQuery(strQuery)

            If rsetPayMethod.RecordCount > 0 Then
                oEditText.Active = True
                oApplication.StatusBar.SetText(strMessage & " [ " & oEditText.Value & " ] : Already Exist in Table...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                Return False
            End If
            Return True

        Catch ex As Exception
            oApplication.StatusBar.SetText(" isDuplicate Function Failed : ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Return False
        End Try
    End Function

#End Region

#Region "Production"
    Function GetWONO(ByVal ProNo As String) As String
        Dim BaseNum = ""
        Dim rs1 As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim StrSql1 As String = "select * from OWOR where docnum='" & ProNo & "' "
        rs1.DoQuery(StrSql1)


        If rs1.Fields.Item("U_ParentPONo").Value <> "" Then
            Dim FatherNo = getSingleValue("select U_ParentPONo from OWOR where DocNum='" & ProNo & "' ")
            BaseNum = GetWONO(FatherNo)
            ' Exit Function
        Else
            BaseNum = getSingleValue("select U_BaseNum from OWOR where DocEntry='" & getSingleValue("select max(Docentry) from OWOR where docnum='" & ProNo & "'") & "'")


        End If
        Return BaseNum
    End Function
    Function GetParentPONO(ByVal ProNo As String) As String
        Dim BaseNum = ""
        Dim rs1 As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim StrSql1 As String = "select * from OWOR where docnum='" & ProNo & "' "
        rs1.DoQuery(StrSql1)


        If rs1.Fields.Item("U_ParentPONo").Value <> "" Then
            Dim FatherNo = getSingleValue("select U_ParentPONo from OWOR where DocNum='" & ProNo & "' ")
            BaseNum = GetParentPONO(FatherNo)
            ' Exit Function
        Else
            BaseNum = getSingleValue("select DocNum from OWOR where DocEntry='" & getSingleValue("select max(Docentry) from OWOR where docnum='" & ProNo & "'") & "'")


        End If
        Return BaseNum
    End Function
#End Region

#Region " ...  Common Function for Forms ..."

    Public Function LoadScreenXML(ByVal FileName As String, ByVal Type As enuResourceType, ByVal FormType As String) As SAPbouiCOM.Form
        intTotalFormCount += 1
        Return LoadScreenXML(FileName, Type, FormType, FormType & intTotalFormCount)
    End Function

    Public Function LoadScreenXML(ByVal FileName As String, ByVal Type As enuResourceType, ByVal FormType As String, ByVal FormUID As String) As SAPbouiCOM.Form
        Dim objForm As SAPbouiCOM.Form
        Dim objXML As New Xml.XmlDocument
        Dim strResource As String
        Dim objFrmCreationPrams As SAPbouiCOM.FormCreationParams

        If Type = enuResourceType.Content Then
            objXML.Load(FileName)
            objFrmCreationPrams = oApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams)
            objFrmCreationPrams.FormType = FormType
            objFrmCreationPrams.UniqueID = FormUID
            objFrmCreationPrams.XmlData = objXML.InnerXml
            objForm = oApplication.Forms.AddEx(objFrmCreationPrams)
        Else
            strResource = System.Reflection.Assembly.GetExecutingAssembly().GetName().Name & "." & FileName
            objXML.Load(System.Reflection.Assembly.GetExecutingAssembly.GetManifestResourceStream(strResource))
            objFrmCreationPrams = oApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams)
            objFrmCreationPrams.FormType = FormType
            objFrmCreationPrams.UniqueID = FormUID
            objFrmCreationPrams.XmlData = objXML.InnerXml
            objForm = oApplication.Forms.AddEx(objFrmCreationPrams)
        End If

        Return objForm
    End Function

    Public Enum enuResourceType
        Embeded
        Content
    End Enum

    Sub LoadXML(ByVal Form As SAPbouiCOM.Form, ByVal FormId As String, ByVal FormXML As String)
        Try
            AddXML(FormXML)
            Form = oApplication.Forms.Item(FormId)
            Form.Select()
        Catch ex As Exception
            oApplication.StatusBar.SetText("LoadXML Method Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
        End Try
    End Sub

    Function LoadComboBoxSeries(ByVal oComboBox As SAPbouiCOM.ComboBox, ByVal UDOID As String) As Boolean
        Try
            oComboBox.ValidValues.LoadSeries(UDOID, SAPbouiCOM.BoSeriesMode.sf_Add)
            oComboBox.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
        Catch ex As Exception
            oApplication.StatusBar.SetText("LoadComboBoxSeries Function Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Return False
        Finally
        End Try
        Return True
    End Function

    Function LoadDocumentDate(ByVal oEditText As SAPbouiCOM.EditText) As Boolean
        Try
            oEditText.Active = True
            oEditText.String = "A"
        Catch ex As Exception
            oApplication.StatusBar.SetText("LoadDocumentDate Function Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Return False
        Finally
        End Try
        Return True
    End Function

    Sub SetComboBoxValueRefresh(ByVal oComboBox As SAPbouiCOM.ComboBox, ByVal strQry As String)
        Try
            Dim rsetValidValue As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim intCount As Integer = oComboBox.ValidValues.Count
            ' Remove the Combo Box Value Based On Count ...
            If intCount > 0 Then
                While intCount > 0
                    oComboBox.ValidValues.Remove(intCount - 1, SAPbouiCOM.BoSearchKey.psk_Index)
                    intCount = intCount - 1
                End While
            End If

            rsetValidValue.DoQuery(strQry)
            rsetValidValue.MoveFirst()
            For j As Integer = 0 To rsetValidValue.RecordCount - 1
                oComboBox.ValidValues.Add(rsetValidValue.Fields.Item(0).Value, rsetValidValue.Fields.Item(1).Value)
                rsetValidValue.MoveNext()
            Next

        Catch ex As Exception
            Msg("SetComboBoxValueRefresh Method Faild : " & ex.Message)
        Finally
        End Try
    End Sub

    Function setComboBoxValue(ByVal oComboBox As SAPbouiCOM.ComboBox, ByVal strQry As String) As Boolean
        Try
            If oComboBox.ValidValues.Count = 0 Then
                Dim rsetValidValue As SAPbobsCOM.Recordset = DoQuery(strQry)
                rsetValidValue.MoveFirst()
                For j As Integer = 0 To rsetValidValue.RecordCount - 1
                    oComboBox.ValidValues.Add(rsetValidValue.Fields.Item(0).Value, rsetValidValue.Fields.Item(1).Value)
                    rsetValidValue.MoveNext()
                Next
            End If

            ' If oComboBox.ValidValues.Count > 0 Then oComboBox.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
        Catch ex As Exception
            oApplication.StatusBar.SetText("setComboBoxValue Function Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Return True
        Finally
        End Try
        Return True
    End Function

    Function setLocationCombo(ByVal oComboBox As SAPbouiCOM.ComboBox) As Boolean
        Try
            setComboBoxValue(oComboBox, "Select Code, Location from OLCT Where Code in ( Select U_LocCode from [@MIPL_WIPGT1] where isnull(U_Active,'N') ='Y' )")

        Catch ex As Exception
            oApplication.StatusBar.SetText("setLocationCombo Function Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Return True
        Finally
        End Try
        Return True
    End Function

    Function GetCodeGeneration(ByVal TableName As String) As Integer
        Try
            Dim rsetCode As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim strCode As String = "Select ISNULL(Max(ISNULL(DocEntry,0)),0) + 1 Code From " & Trim(TableName) & ""
            rsetCode.DoQuery(strCode)
            Return CInt(rsetCode.Fields.Item("Code").Value)
        Catch ex As Exception
            oApplication.StatusBar.SetText("GetCodeGeneration Function Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Return True
        Finally
        End Try
    End Function

    Sub SetNewLine(ByVal oMatrix As SAPbouiCOM.Matrix, ByVal oDBDSDetail As SAPbouiCOM.DBDataSource, Optional ByVal RowID As Integer = 1, Optional ByVal ColumnUID As String = "")
        Try
            If ColumnUID.Equals("") = False Then
                If oMatrix.VisualRowCount > 0 Then
                    If oMatrix.Columns.Item(ColumnUID).Cells.Item(RowID).Specific.Value.Equals("") = False And RowID = oMatrix.VisualRowCount Then
                        oMatrix.FlushToDataSource()
                        oMatrix.AddRow()
                        oDBDSDetail.InsertRecord(oDBDSDetail.Size)
                        oDBDSDetail.Offset = oMatrix.VisualRowCount - 1
                        oDBDSDetail.SetValue("LineID", oDBDSDetail.Offset, oMatrix.VisualRowCount)

                        oMatrix.SetLineData(oMatrix.VisualRowCount)
                        oMatrix.FlushToDataSource()
                    End If
                Else
                    oMatrix.FlushToDataSource()
                    oMatrix.AddRow()
                    oDBDSDetail.InsertRecord(oDBDSDetail.Size)
                    oDBDSDetail.Offset = oMatrix.VisualRowCount - 1
                    oDBDSDetail.SetValue("LineID", oDBDSDetail.Offset, oMatrix.VisualRowCount)

                    oMatrix.SetLineData(oMatrix.VisualRowCount)
                    oMatrix.FlushToDataSource()
                End If

            Else
                oMatrix.FlushToDataSource()
                oMatrix.AddRow()
                oDBDSDetail.InsertRecord(oDBDSDetail.Size)
                oDBDSDetail.Offset = oMatrix.VisualRowCount - 1
                oDBDSDetail.SetValue("LineID", oDBDSDetail.Offset, oMatrix.VisualRowCount)

                oMatrix.SetLineData(oMatrix.VisualRowCount)
                oMatrix.FlushToDataSource()
            End If
        Catch ex As Exception
            oApplication.StatusBar.SetText("SetNewLine Method Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
        End Try
    End Sub

    Public Sub Matrix_Addrow(ByVal omatrix As SAPbouiCOM.Matrix, Optional ByVal colname As String = "", Optional ByVal rowno_name As String = "", Optional ByVal Error_Needed As Boolean = False)
        Try
            Dim addrow As Boolean = False

            If omatrix.VisualRowCount = 0 Then addrow = True : GoTo addrow
            If colname = "" Then addrow = True : GoTo addrow
            If omatrix.Columns.Item(colname).Cells.Item(omatrix.VisualRowCount).Specific.string <> "" Then addrow = True : GoTo addrow

addrow:
            If addrow = True Then
                omatrix.AddRow(1)
                omatrix.ClearRowData(omatrix.VisualRowCount)
                If rowno_name <> "" Then omatrix.Columns.Item("LineId").Cells.Item(omatrix.VisualRowCount).Specific.string = omatrix.VisualRowCount
            Else
                If Error_Needed = True Then oApplication.SetStatusBarMessage("Already Empty Row Available", SAPbouiCOM.BoMessageTime.bmt_Short, True)
            End If
        Catch ex As Exception

        End Try
    End Sub

    Sub setEdittextCFL(ByVal oForm As SAPbouiCOM.Form, ByVal UId As String, ByVal strCFL_ID As String, ByVal strCFL_Obj As String, ByVal strCFL_Alies As String)
        Try

            Dim oCFL As SAPbouiCOM.ChooseFromListCreationParams
            oCFL = oApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
            oCFL.UniqueID = strCFL_ID
            oCFL.ObjectType = strCFL_Obj
            oForm.ChooseFromLists.Add(oCFL)

            Dim txt As SAPbouiCOM.EditText = oForm.Items.Item(UId).Specific
            txt.ChooseFromListUID = strCFL_ID
            txt.ChooseFromListAlias = strCFL_Alies

        Catch ex As Exception
            oApplication.StatusBar.SetText("Set EditText CFL Method Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
        End Try
    End Sub

    Sub setEditTextColumnCFL(ByVal oForm As SAPbouiCOM.Form, ByVal oMatrix As SAPbouiCOM.Matrix, ByVal UId As String, ByVal strCFL_ID As String, ByVal strCFL_Obj As String, ByVal strCFL_Alies As String)
        Try

            Dim oCFL As SAPbouiCOM.ChooseFromListCreationParams
            oCFL = oApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
            oCFL.UniqueID = strCFL_ID
            oCFL.ObjectType = strCFL_Obj
            oForm.ChooseFromLists.Add(oCFL)

            oMatrix.Columns.Item(UId).ChooseFromListUID = strCFL_ID
            oMatrix.Columns.Item(UId).ChooseFromListAlias = strCFL_Alies

        Catch ex As Exception
            oApplication.StatusBar.SetText("Set EditText CFL Method Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
        End Try
    End Sub

    Sub ChooseFromListFilteration(ByVal oForm As SAPbouiCOM.Form, ByVal strCFL_ID As String, ByVal strCFL_Alies As String, ByVal strQuery As String)
        Try
            Dim oCFL As SAPbouiCOM.ChooseFromList = oForm.ChooseFromLists.Item(strCFL_ID)
            Dim oConds As SAPbouiCOM.Conditions
            Dim oCond As SAPbouiCOM.Condition
            Dim oEmptyConds As New SAPbouiCOM.Conditions
            Dim rsetCFL As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oCFL.SetConditions(oEmptyConds)
            oConds = oCFL.GetConditions()

            rsetCFL.DoQuery(strQuery)
            rsetCFL.MoveFirst()
            For i As Integer = 1 To rsetCFL.RecordCount
                If i = (rsetCFL.RecordCount) Then
                    oCond = oConds.Add()
                    oCond.Alias = strCFL_Alies
                    oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCond.CondVal = Trim(rsetCFL.Fields.Item(0).Value)
                Else
                    oCond = oConds.Add()
                    oCond.Alias = strCFL_Alies
                    oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCond.CondVal = Trim(rsetCFL.Fields.Item(0).Value)
                    oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR
                End If
                rsetCFL.MoveNext()
            Next
            If rsetCFL.RecordCount = 0 Then
                oCond = oConds.Add()
                oCond.Alias = strCFL_Alies
                oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_NONE
                oCond.CondVal = "-1"
            End If
            oCFL.SetConditions(oConds)
        Catch ex As Exception
            oApplication.StatusBar.SetText("Choose FromList Filter Global Fun. Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
        End Try
    End Sub

    Sub DeleteRow(ByVal oMatrix As SAPbouiCOM.Matrix, ByVal oDBDSDetail As SAPbouiCOM.DBDataSource)
        Try
            oMatrix.FlushToDataSource()

            For i As Integer = 1 To oMatrix.VisualRowCount
                oMatrix.GetLineData(i)
                oDBDSDetail.Offset = i - 1
                oDBDSDetail.SetValue("LineID", oDBDSDetail.Offset, i)
                oMatrix.SetLineData(i)
                oMatrix.FlushToDataSource()
            Next
            'oDBDSDetail.RemoveRecord(oDBDSDetail.Size - 1)
            'oMatrix.LoadFromDataSource()

        Catch ex As Exception
            oApplication.StatusBar.SetText("DeleteRow  Method Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
        End Try
    End Sub

    Function getDateWhFarmate(ByVal docdate As String)
        ' Dim dt = oApplication.Forms.Item(0).Items.Item("15").Specific.Value.ToString
        getDateWhFarmate = docdate.Substring(4, 2) & "/" & docdate.Substring(6, 2) & "/" & docdate.Substring(0, 4)
    End Function

    Sub DeleteEmptyRowInFormDataEvent(ByVal oMatrix As SAPbouiCOM.Matrix, ByVal ColumnUID As String, ByVal oDBDSDetail As SAPbouiCOM.DBDataSource)
        Try
            If oMatrix.VisualRowCount > 0 Then
                If oMatrix.Columns.Item(ColumnUID).Cells.Item(oMatrix.VisualRowCount).Specific.Value.Equals("") Then
                    oMatrix.DeleteRow(oMatrix.VisualRowCount)
                    oDBDSDetail.RemoveRecord(oDBDSDetail.Size - 1)
                    oMatrix.FlushToDataSource()
                End If
            End If
        Catch ex As Exception
            oApplication.StatusBar.SetText("Delete Empty RowIn Function Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
        End Try
    End Sub

    Sub setDocNum(ByVal ofrm As SAPbouiCOM.Form)
        Try
            Dim StrDocNum As Long
            Dim strSerialCode As String = ofrm.Items.Item("c_Series").Specific.Selected.Value
            StrDocNum = ofrm.BusinessObject.GetNextSerialNumber(strSerialCode, ofrm.BusinessObject.Type)
            ofrm.DataSources.DBDataSources.Item(0).SetValue("DocNum", 0, StrDocNum)
        Catch ex As Exception
            oApplication.StatusBar.SetText("Set DocNum Function Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
        End Try
    End Sub
   

    Sub setRightMenu(ByVal strMenuUID As String, ByVal strMenuName As String)
        Try
            Dim MenuItem As SAPbouiCOM.MenuItem = oApplication.Menus.Item("1280") 'Data'
            Dim Menu As SAPbouiCOM.Menus = MenuItem.SubMenus
            Dim MenuParam As SAPbouiCOM.MenuCreationParams = oApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)

            MenuParam.Type = SAPbouiCOM.BoMenuType.mt_STRING
            MenuParam.UniqueID = strMenuUID
            MenuParam.String = strMenuName
            MenuParam.Enabled = True
            If MenuItem.SubMenus.Exists(strMenuUID) = False Then Menu.AddEx(MenuParam)

        Catch ex As Exception
            oApplication.StatusBar.SetText("SubMenuAddEx Method Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
        End Try
    End Sub

    Sub RemoveRightMenu(ByVal strMenuID As String)
        Try
            If oApplication.Menus.Item("1280").SubMenus.Exists(strMenuID) Then oApplication.Menus.Item("1280").SubMenus.RemoveEx(strMenuID)

        Catch ex As Exception
            oApplication.StatusBar.SetText("SubMenusRemoveEx Method Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
        End Try
    End Sub

    Public Function FormExist(ByVal FormID As String) As Boolean
        FormExist = False
        For Each uid As SAPbouiCOM.Form In oApplication.Forms
            If uid.UniqueID = FormID Then
                FormExist = True
                Exit Function
            End If
        Next
        If FormExist Then
            oApplication.Forms.Item(FormID).Visible = True
            oApplication.Forms.Item(FormID).Select()
        End If
    End Function

    ' Convert Hours
    Function ConvertHrs(ByVal Value As String, ByVal TimeMesurement As String) As Double
        Try

            Value = IIf(Value.ToString.Trim = "", 0, Value)
            Dim ValueTimeHr As Double = 0
            Select Case TimeMesurement
                Case "S"
                    ValueTimeHr = Value / 3600
                Case "M"
                    ValueTimeHr = Value / 60
                Case "H"
                    ValueTimeHr = Value
                Case "D"
                    ValueTimeHr = Value * 24
            End Select

            Return ValueTimeHr

        Catch ex As Exception
            oApplication.StatusBar.SetText("Convert Hours Method Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Return False
        Finally
        End Try
    End Function

    Function FindMinCode(ByVal TableName As String, ByVal FieldName As String, ByVal KeyField As String, ByVal KeyValue As String) As String
        Try
            Dim strsql = " select " & FieldName & " from " & _
                        " ( select ROW_NUMBER() over(Order by " & FieldName & ") ROWNUMBER , " & FieldName & " from " & TableName & " where isnull(convert(varchar," & KeyField & "),'N') ='" & KeyValue & "' " & _
                        " group by " & FieldName & ",convert(varchar," & KeyField & ") ) AA  where ROWNUMBER = (Select Min(ROWNUMBER) from ( select ROW_NUMBER() " & _
                        " over(Order by " & FieldName & ") ROWNUMBER , " & FieldName & " from " & TableName & " where isnull(convert(varchar," & KeyField & "),'N') ='" & KeyValue & "' " & _
                        " group by " & FieldName & ",convert(varchar," & KeyField & ") ) TMP1) "
            Dim returnValue = getSingleValue(strsql)
            Return returnValue
        Catch ex As Exception
            oApplication.StatusBar.SetText("Find Min Code Method Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Return False
        End Try
    End Function

    Function FindMaxCode(ByVal TableName As String, ByVal FieldName As String, ByVal KeyField As String, ByVal KeyValue As String) As String
        Try
            Dim strsql = " select " & FieldName & " from " & _
                        " ( select ROW_NUMBER() over(Order by " & FieldName & ") ROWNUMBER , " & FieldName & " from " & TableName & " where isnull(convert(varchar," & KeyField & "),'N') ='" & KeyValue & "' " & _
                        " group by " & FieldName & ",convert(varchar," & KeyField & ") ) AA  where ROWNUMBER = (Select Max(ROWNUMBER) from ( select ROW_NUMBER() " & _
                        " over(Order by " & FieldName & ") ROWNUMBER , " & FieldName & " from " & TableName & " where isnull(convert(varchar," & KeyField & "),'N') ='" & KeyValue & "' " & _
                        " group by " & FieldName & ",convert(varchar," & KeyField & ") ) TMP1) "
            Dim returnValue = getSingleValue(strsql)
            Return returnValue
        Catch ex As Exception
            oApplication.StatusBar.SetText("Find Next Code Method Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Return False
        End Try
    End Function

    Function FindNextCode(ByVal TableName As String, ByVal FieldName As String, ByVal KeyField As String, ByVal KeyValue As String, ByVal CurrentValue As String) As String
        Try
            If CurrentValue.Trim = "" Then
                Return FindMinCode(TableName, FieldName, KeyField, KeyValue)
            Else

                'Find the Current Row
                '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                Dim strsql = " select convert(int,ROWNUMBER) from ( select ROW_NUMBER() over(Order by " & _
                   FieldName & ") ROWNUMBER , " & FieldName & " from " & TableName & " where isnull(convert(varchar," & KeyField & "),'N') ='" & KeyValue & "' " & _
                   " group by " & FieldName & ",convert(varchar," & KeyField & ") ) AA  where " & FieldName & " = '" & CurrentValue & "' "

                Dim CurrentRow = getSingleValue(strsql)
                '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

                'Max Row Number
                '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                strsql = " select Max(convert(int,ROWNUMBER)) from ( select ROW_NUMBER() over(Order by " & _
                   FieldName & ") ROWNUMBER , " & FieldName & " from " & TableName & " where isnull(convert(varchar," & KeyField & "),'N') ='" & KeyValue & "' " & _
                   " group by " & FieldName & ",convert(varchar," & KeyField & ") ) AA  "

                Dim MaxRow = getSingleValue(strsql)
                MaxRow = IIf(MaxRow.Trim = "", 0, MaxRow)
                '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

                Dim NextRow As Integer
                CurrentRow = IIf(CurrentRow.Trim = "", 0, CurrentRow)
                If CInt(CurrentRow) = 0 Or CInt(CurrentRow) = CInt(MaxRow) Then
                    Return ""
                Else
                    NextRow = CInt(CurrentRow) + 1
                End If

                strsql = " select " & FieldName & " from ( select ROW_NUMBER() over(Order by " & _
                   FieldName & ") ROWNUMBER , " & FieldName & " from " & TableName & " where isnull(convert(varchar," & KeyField & "),'N') ='" & KeyValue & "' " & _
                   " group by " & FieldName & ",convert(varchar," & KeyField & ") ) AA  where  ROWNUMBER = '" & NextRow & "' "

                Dim returnValue As String = getSingleValue(strsql)
                Return returnValue

            End If
        Catch ex As Exception
            oApplication.StatusBar.SetText("Find Max Code Method Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Return False
        End Try
    End Function

    Function FindPreviousCode(ByVal TableName As String, ByVal FieldName As String, ByVal KeyField As String, ByVal KeyValue As String, ByVal CurrentValue As String) As String
        Try
            If CurrentValue.Trim = "" Then
                Return FindMaxCode(TableName, FieldName, KeyField, KeyValue)
            Else

                'Find the Current Row
                '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                Dim strsql = "select convert(int, ROWNUMBER) from ( select ROW_NUMBER() over(Order by " & _
                   FieldName & ") ROWNUMBER , " & FieldName & " from " & TableName & " where isnull(convert(varchar," & KeyField & "),'N') ='" & KeyValue & "' " & _
                   " group by " & FieldName & ",convert(varchar," & KeyField & ") ) AA  where " & FieldName & " = '" & CurrentValue & "' "

                Dim CurrentRow = getSingleValue(strsql)
                '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

                'Min Row Number
                '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                strsql = " select Min(convert(int,ROWNUMBER)) from ( select ROW_NUMBER() over(Order by " & _
                   FieldName & ") ROWNUMBER , " & FieldName & " from " & TableName & " where isnull(convert(varchar," & KeyField & "),'N') ='" & KeyValue & "' " & _
                   " group by " & FieldName & ",convert(varchar," & KeyField & ") ) AA  "

                Dim MinRow = getSingleValue(strsql)
                MinRow = IIf(MinRow.Trim = "", 0, MinRow)
                '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

                Dim NextRow As Integer
                CurrentRow = IIf(CurrentRow.Trim = "", 0, CurrentRow)
                If CInt(CurrentRow) = 0 Or CInt(CurrentRow) = CInt(MinRow) Then
                    Return ""
                Else
                    NextRow = CInt(CurrentRow) - 1
                End If

                strsql = " select " & FieldName & " from ( select ROW_NUMBER() over(Order by " & _
                   FieldName & ") ROWNUMBER , " & FieldName & " from " & TableName & " where isnull(convert(varchar," & KeyField & "),'N') ='" & KeyValue & "' " & _
                   " group by " & FieldName & ",convert(varchar," & KeyField & ") ) AA  where  ROWNUMBER = '" & NextRow & "' "

                Dim returnValue As String = getSingleValue(strsql)
                Return returnValue

            End If
        Catch ex As Exception
            oApplication.StatusBar.SetText("Find Pervious Code Method Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Return False
        End Try
    End Function

#End Region

#Region " ...  Common Function User Comunication ..."

    Sub Msg(ByVal strMsg As String, Optional ByVal msgTime As String = "S", Optional ByVal errType As String = "W")
        Dim time As SAPbouiCOM.BoMessageTime
        Dim msgType As SAPbouiCOM.BoStatusBarMessageType
        Select Case errType.ToUpper()
            Case "E"
                msgType = SAPbouiCOM.BoStatusBarMessageType.smt_Error
            Case "W"
                msgType = SAPbouiCOM.BoStatusBarMessageType.smt_Warning
            Case "N"
                msgType = SAPbouiCOM.BoStatusBarMessageType.smt_None
            Case "S"
                msgType = SAPbouiCOM.BoStatusBarMessageType.smt_Success
            Case Else
                msgType = SAPbouiCOM.BoStatusBarMessageType.smt_Warning
        End Select
        Select Case msgTime.ToUpper()
            Case "M"
                time = SAPbouiCOM.BoMessageTime.bmt_Medium
            Case "S"
                time = SAPbouiCOM.BoMessageTime.bmt_Short
            Case "L"
                time = SAPbouiCOM.BoMessageTime.bmt_Long
            Case Else
                time = SAPbouiCOM.BoMessageTime.bmt_Medium
        End Select
        oApplication.StatusBar.SetText(strMsg, time, msgType)
    End Sub
    Sub MsgBox(ByVal strMsg As String, Optional ByVal msgTime As String = "S", Optional ByVal errType As String = "W")
        Dim time As SAPbouiCOM.BoMessageTime
        Dim msgType As SAPbouiCOM.BoStatusBarMessageType
        Select Case errType.ToUpper()
            Case "E"
                msgType = SAPbouiCOM.BoStatusBarMessageType.smt_Error
            Case "W"
                msgType = SAPbouiCOM.BoStatusBarMessageType.smt_Warning
            Case "N"
                msgType = SAPbouiCOM.BoStatusBarMessageType.smt_None
            Case "S"
                msgType = SAPbouiCOM.BoStatusBarMessageType.smt_Success
            Case Else
                msgType = SAPbouiCOM.BoStatusBarMessageType.smt_Warning
        End Select
        Select Case msgTime.ToUpper()
            Case "M"
                time = SAPbouiCOM.BoMessageTime.bmt_Medium
            Case "S"
                time = SAPbouiCOM.BoMessageTime.bmt_Short
            Case "L"
                time = SAPbouiCOM.BoMessageTime.bmt_Long
            Case Else
                time = SAPbouiCOM.BoMessageTime.bmt_Medium
        End Select
        'oApplication.MessageBox(strMsg, time, msgType)
        oApplication.MessageBox(strMsg,, "OK")
        'oApplication.StatusBar.SetText(strMsg, time, msgType)
    End Sub
    'Changed by kannan
    Public Sub WriteSMSLog(ByVal Str As String)
        Dim fs As FileStream
        Dim chatlog As String = oApplication.StartupPath & "\Log_" & Today.ToString("yyyyMMdd") & ".txt"
        If File.Exists(chatlog) Then
        Else
            fs = New FileStream(chatlog, FileMode.Create, FileAccess.Write)
            fs.Close()
        End If
        ' Dim objReader As New System.IO.StreamReader(chatlog)
        Dim sdate As String
        sdate = Now
        'objReader.Close()
        If System.IO.File.Exists(chatlog) = True Then
            Dim objWriter As New System.IO.StreamWriter(chatlog, True)
            objWriter.WriteLine(sdate & " : " & Str)
            objWriter.Close()
        Else
            Dim objWriter As New System.IO.StreamWriter(chatlog, False)
            ' MsgBox("Failed to send message!")
        End If
    End Sub

#End Region

#Region "          Attachment Option          "
    Dim BankFileName = ""

    Sub AddAttachment(ByVal oMatAttach As SAPbouiCOM.Matrix, ByVal oDBDSAttch As SAPbouiCOM.DBDataSource, ByVal oDBDSHeader As SAPbouiCOM.DBDataSource)
        Try
            If oMatAttach.VisualRowCount > 0 Then
                Dim rsetAttCount As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Dim oAttachment As SAPbobsCOM.Attachments2 = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oAttachments2)
                Dim oAttchLines As SAPbobsCOM.Attachments2_Lines
                oAttchLines = oAttachment.Lines
                oMatAttach.FlushToDataSource()
                rsetAttCount.DoQuery("Select Count(*) From ATC1 Where AbsEntry = '" & Trim(oDBDSHeader.GetValue("U_AtcEntry", 0)) & "'")

                If Trim(rsetAttCount.Fields.Item(0).Value).Equals("0") Then
                    For i As Integer = 1 To oMatAttach.VisualRowCount
                        If i > 1 Then oAttchLines.Add()
                        oDBDSAttch.Offset = i - 1
                        oAttchLines.SourcePath = Trim(oDBDSAttch.GetValue("U_ScrPath", oDBDSAttch.Offset))
                        oAttchLines.FileName = Trim(oDBDSAttch.GetValue("U_FileName", oDBDSAttch.Offset))
                        oAttchLines.FileExtension = Trim(oDBDSAttch.GetValue("U_FileExt", oDBDSAttch.Offset))
                        oAttchLines.Override = SAPbobsCOM.BoYesNoEnum.tYES
                    Next
                    oAttachment.Add()
                    Dim rsetAttch As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    rsetAttch.DoQuery("Select  Case When Count(*) > 0 Then  Max(AbsEntry) Else 0 End AbsEntry  From ATC1")
                    oDBDSHeader.SetValue("U_AtcEntry", 0, rsetAttch.Fields.Item(0).Value)
                Else
                    oAttachment.GetByKey(Trim(oDBDSHeader.GetValue("U_AtcEntry", 0)))
                    For i As Integer = 1 To oMatAttach.VisualRowCount
                        If oAttchLines.Count < i Then oAttchLines.Add()
                        oDBDSAttch.Offset = i - 1
                        oAttchLines.SetCurrentLine(i - 1)
                        oAttchLines.SourcePath = Trim(oDBDSAttch.GetValue("U_ScrPath", oDBDSAttch.Offset))
                        oAttchLines.FileName = Trim(oDBDSAttch.GetValue("U_FileName", oDBDSAttch.Offset))
                        oAttchLines.FileExtension = Trim(oDBDSAttch.GetValue("U_FileExt", oDBDSAttch.Offset))
                        oAttchLines.Override = SAPbobsCOM.BoYesNoEnum.tYES
                    Next
                    oAttachment.Update()
                End If
            End If
            'Delete the Attachment Rows...
            Dim rsetDelete As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            rsetDelete.DoQuery("Delete From ATC1 Where AbsEntry = '" & Trim(oDBDSHeader.GetValue("U_AtcEntry", 0)) & "' And Line >'" & oMatAttach.VisualRowCount & "' ")

        Catch ex As Exception
            oApplication.StatusBar.SetText("AddAttachment Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
        End Try
    End Sub

    Sub DeleteRowAttachment(ByVal oForm As SAPbouiCOM.Form, ByVal oMatrix As SAPbouiCOM.Matrix, ByVal oDBDSAttch As SAPbouiCOM.DBDataSource, ByVal SelectedRowID As Integer)
        Try
            oDBDSAttch.RemoveRecord(SelectedRowID - 1)
            oMatrix.DeleteRow(SelectedRowID)
            oMatrix.FlushToDataSource()

            For i As Integer = 1 To oMatrix.VisualRowCount
                oMatrix.GetLineData(i)
                oDBDSAttch.Offset = i - 1

                oDBDSAttch.SetValue("LineID", oDBDSAttch.Offset, i)
                oDBDSAttch.SetValue("U_TrgtPath", oDBDSAttch.Offset, Trim(oMatrix.Columns.Item("TrgtPath").Cells.Item(i).Specific.Value))
                oDBDSAttch.SetValue("U_ScrPath", oDBDSAttch.Offset, Trim(oMatrix.Columns.Item("Path").Cells.Item(i).Specific.Value))
                oDBDSAttch.SetValue("U_FileName", oDBDSAttch.Offset, Trim(oMatrix.Columns.Item("FileName").Cells.Item(i).Specific.Value))
                oDBDSAttch.SetValue("U_FileExt", oDBDSAttch.Offset, Trim(oMatrix.Columns.Item("FileExt").Cells.Item(i).Specific.Value))
                oDBDSAttch.SetValue("U_Date", oDBDSAttch.Offset, Trim(oMatrix.Columns.Item("Date").Cells.Item(i).Specific.Value))
                oMatrix.SetLineData(i)
                oMatrix.FlushToDataSource()
            Next
            'oDBDSAttch.RemoveRecord(oDBDSAttch.Size - 1)
            oMatrix.LoadFromDataSource()

            oForm.Items.Item("b_display").Enabled = False
            oForm.Items.Item("b_delete").Enabled = False

            If oForm.Mode <> SAPbouiCOM.BoFormMode.fm_ADD_MODE Then oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE

        Catch ex As Exception
            oApplication.StatusBar.SetText("DeleteRowAttachment Method Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
        End Try
    End Sub

    Sub OpenAttachment1(ByVal oEdit As SAPbouiCOM.EditText)
        Try
            'Open Attachment File
            OpenFile(oEdit.Value, oEdit.Value)


        Catch ex As Exception
            oApplication.StatusBar.SetText("OpenAttachment Method Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
        End Try
    End Sub

    Function SetAttachMentFile(ByVal oEdit As SAPbouiCOM.EditText) As Boolean

        'Log_init()

        Try



            If oCompany.AttachMentPath.Length <= 0 Then
                oApplication.StatusBar.SetText("Attchment folder not defined, or Attchment folder has been changed or removed. [Message 131-102]")
                Return False
            End If



            Dim strFileName As String = FindFile()



            If strFileName.Equals("") = False Then



                Dim FileExist() As String = strFileName.Split("\")
                Dim FileDestPath As String = oCompany.AttachMentPath & FileExist(FileExist.Length - 1)


                If File.Exists(FileDestPath) Then
                    Dim LngRetVal As Long = oApplication.MessageBox("A file with this name already exists,would you like to replace this?  " & FileDestPath & " will be replaced.", 1, "Yes", "No")
                    If LngRetVal <> 1 Then Return False
                End If


                Dim fileNameExt() As String = FileExist(FileExist.Length - 1).Split(".")
                Dim ScrPath As String = oCompany.AttachMentPath
                ScrPath = ScrPath.Substring(0, ScrPath.Length - 1)
                Dim TrgtPath As String = strFileName.Substring(0, strFileName.LastIndexOf("\"))


                oEdit.Value = TrgtPath + "\" + fileNameExt(0) + "." + fileNameExt(1)



                'oDBDSAttch.SetValue("U_TrgtPath", oDBDSAttch.Offset, ScrPath)               
                'oDBDSAttch.SetValue("U_ScrPath", oDBDSAttch.Offset, TrgtPath)
                'oDBDSAttch.SetValue("U_FileName", oDBDSAttch.Offset, fileNameExt(0))
                'oDBDSAttch.SetValue("U_FileExt", oDBDSAttch.Offset, fileNameExt(1))
                'oDBDSAttch.SetValue("U_Date", oDBDSAttch.Offset, GetServerDate())

            End If
            Return True
        Catch ex As Exception

            oApplication.StatusBar.SetText("Set AttachMent File Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Return False
        Finally
        End Try
    End Function

    Function SetAttachMentFile(ByVal oForm As SAPbouiCOM.Form, ByVal oDBDSHeader As SAPbouiCOM.DBDataSource, ByVal oMatrix As SAPbouiCOM.Matrix, ByVal oDBDSAttch As SAPbouiCOM.DBDataSource) As Boolean
        Try
            If oCompany.AttachMentPath.Length <= 0 Then
                oApplication.StatusBar.SetText("Attchment folder not defined, or Attchment folder has been changed or removed. [Message 131-102]")
                Return False
            End If

            Dim strFileName As String = FindFile()
            If strFileName.Equals("") = False Then
                Dim FileExist() As String = strFileName.Split("\")
                Dim FileDestPath As String = oCompany.AttachMentPath & FileExist(FileExist.Length - 1)

                If File.Exists(FileDestPath) Then
                    Dim LngRetVal As Long = oApplication.MessageBox("A file with this name already exists,would you like to replace this?  " & FileDestPath & " will be replaced.", 1, "Yes", "No")
                    If LngRetVal <> 1 Then Return False
                End If
                Dim fileNameExt() As String = FileExist(FileExist.Length - 1).Split(".")
                Dim ScrPath As String = oCompany.AttachMentPath
                ScrPath = ScrPath.Substring(0, ScrPath.Length - 1)
                Dim TrgtPath As String = strFileName.Substring(0, strFileName.LastIndexOf("\"))

                oMatrix.AddRow()
                oMatrix.FlushToDataSource()
                oDBDSAttch.Offset = oDBDSAttch.Size - 1
                oDBDSAttch.SetValue("LineID", oDBDSAttch.Offset, oMatrix.VisualRowCount)
                oDBDSAttch.SetValue("U_TrgtPath", oDBDSAttch.Offset, ScrPath)
                oDBDSAttch.SetValue("U_ScrPath", oDBDSAttch.Offset, TrgtPath)
                oDBDSAttch.SetValue("U_FileName", oDBDSAttch.Offset, fileNameExt(0))
                oDBDSAttch.SetValue("U_FileExt", oDBDSAttch.Offset, fileNameExt(1))
                oDBDSAttch.SetValue("U_Date", oDBDSAttch.Offset, GetServerDate())
                oMatrix.SetLineData(oDBDSAttch.Size)
                oMatrix.FlushToDataSource()
                If oForm.Mode <> SAPbouiCOM.BoFormMode.fm_ADD_MODE Then oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
            End If
            Return True
        Catch ex As Exception
            oApplication.StatusBar.SetText("Set AttachMent File Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Return False
        Finally
        End Try
    End Function

    Sub OpenAttachment(ByVal oMatrix As SAPbouiCOM.Matrix, ByVal oDBDSAttch As SAPbouiCOM.DBDataSource, ByVal PvalRow As Integer)
        Try
            If PvalRow <= oMatrix.VisualRowCount And PvalRow <> 0 Then
                Dim RowIndex As Integer = oMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_RowOrder) - 1
                Dim strServerPath, strClientPath As String

                strServerPath = Trim(oDBDSAttch.GetValue("U_TrgtPath", RowIndex)) + "\" + Trim(oDBDSAttch.GetValue("U_FileName", RowIndex)) + "." + Trim(oDBDSAttch.GetValue("U_FileExt", RowIndex))
                strClientPath = Trim(oDBDSAttch.GetValue("U_ScrPath", RowIndex)) + "\" + Trim(oDBDSAttch.GetValue("U_FileName", RowIndex)) + "." + Trim(oDBDSAttch.GetValue("U_FileExt", RowIndex))
                'Open Attachment File
                OpenFile(strServerPath, strClientPath)
            End If

        Catch ex As Exception
            oApplication.StatusBar.SetText("OpenAttachment Method Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
        End Try
    End Sub

    Sub AttchButtonEnable(ByVal oForm As SAPbouiCOM.Form, ByVal Matrix As SAPbouiCOM.Matrix, ByVal PvalRow As Integer)
        Try
            If PvalRow <= Matrix.VisualRowCount And PvalRow <> 0 Then
                Matrix.SelectRow(PvalRow, True, False)
                If Matrix.IsRowSelected(PvalRow) = True Then
                    oForm.Items.Item("b_display").Enabled = True
                    oForm.Items.Item("b_delete").Enabled = True
                Else
                    oForm.Items.Item("b_display").Enabled = False
                    oForm.Items.Item("b_delete").Enabled = False
                End If
            End If
        Catch ex As Exception
            oApplication.StatusBar.SetText("Attach Button Enble Function...")
        End Try
    End Sub

    Public Function FindFile() As String


        Dim ShowFolderBrowserThread As Threading.Thread
        Try
            ShowFolderBrowserThread = New Threading.Thread(AddressOf ShowFolderBrowser)



            If ShowFolderBrowserThread.ThreadState = System.Threading.ThreadState.Unstarted Then


                ShowFolderBrowserThread.SetApartmentState(System.Threading.ApartmentState.STA)
                ShowFolderBrowserThread.Start()
            ElseIf ShowFolderBrowserThread.ThreadState = System.Threading.ThreadState.Stopped Then

                ShowFolderBrowserThread.Start()
                ShowFolderBrowserThread.Join()
            End If



            While ShowFolderBrowserThread.ThreadState = Threading.ThreadState.Running
                System.Windows.Forms.Application.DoEvents()
                ShowFolderBrowserThread.Sleep(100)
            End While



            If BankFileName <> "" Then



                Return BankFileName
            End If



        Catch ex As Exception

            oApplication.MessageBox("File Find  Method Failed : " & ex.Message)
        End Try
        Return ""
    End Function

    Public Sub OpenFile(ByVal ServerPath As String, ByVal ClientPath As String)
        Try
            Dim oProcess As System.Diagnostics.Process = New System.Diagnostics.Process
            Try
                oProcess.StartInfo.FileName = ServerPath
                oProcess.Start()
            Catch ex1 As Exception
                Try
                    oProcess.StartInfo.FileName = ClientPath
                    oProcess.Start()
                Catch ex2 As Exception
                    oApplication.StatusBar.SetText("" & ex2.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                Finally
                End Try
            Finally
            End Try
        Catch ex As Exception
            oApplication.StatusBar.SetText("" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
        End Try
    End Sub

    Function GetServerDate() As String
        Try
            Dim rsetBob As SAPbobsCOM.SBObob = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge)
            Dim rsetServerDate As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            rsetServerDate = rsetBob.Format_StringToDate(oApplication.Company.ServerDate())

            Return CDate(rsetServerDate.Fields.Item(0).Value).ToString("yyyyMMdd")
            'Return "20120215"
        Catch ex As Exception
            Return ""
        Finally
        End Try
    End Function

    Public Sub ShowFolderBrowser()



        Dim MyProcs() As System.Diagnostics.Process




        Dim OpenFile As New OpenFileDialog
        Try
            OpenFile.Multiselect = False



            OpenFile.Filter = "All files(*.)|*.*" '   "|*.*"
            Dim filterindex As Integer = 0
            Try
                filterindex = 0
            Catch ex As Exception
            End Try
            OpenFile.FilterIndex = filterindex
            OpenFile.RestoreDirectory = True

            MyProcs = Process.GetProcessesByName("SAP Business One")
            Try

            Catch
            End Try

            ' *******  Modified on 09-Mar-2012 By parthiban ********

            ' If two or more company opened at the same time,  Dialog is  not opening 
            ' Changed Conditon   to >= 1
            ' Added Condition --If comname(1).ToString.Trim.ToUpper = com Then -- to open dialog
            ' only for this company

            'If MyProcs.Length = 1 Then
            If MyProcs.Length >= 1 Then

                For i As Integer = 0 To MyProcs.Length - 1
                    Dim comname As String() = MyProcs(i).MainWindowTitle.ToString.Split("-")

                    'Open dialog only for the company where the button is clicked
                    Dim com As String = oCompany.CompanyName.ToString.Trim.ToUpper
                    If comname(1).ToString.Trim.ToUpper = com Then
                        Dim MyWindow As New WindowWrapper(MyProcs(i).MainWindowHandle)

                        Dim ret As Windows.Forms.DialogResult = OpenFile.ShowDialog(MyWindow)
                        If ret = Windows.Forms.DialogResult.OK Then

                            BankFileName = OpenFile.FileName
                            'OpenFile.Dispose()


                        Else
                            System.Windows.Forms.Application.ExitThread()
                        End If
                    End If
                Next
            Else
            End If
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
            BankFileName = ""
        Finally
            OpenFile.Dispose()
        End Try
    End Sub

    Public Class WindowWrapper

        Implements System.Windows.Forms.IWin32Window
        Private _hwnd As IntPtr

        Public Sub New(ByVal handle As IntPtr)
            _hwnd = handle
        End Sub

        Public ReadOnly Property Handle() As System.IntPtr Implements System.Windows.Forms.IWin32Window.Handle
            Get
                Return _hwnd
            End Get
        End Property

    End Class
#End Region

    Public Function ConvertRecordset(ByVal SAPRecordset As SAPbobsCOM.Recordset) As System.Data.DataTable
        '\ This function will take an SAP recordset from the SAPbobsCOM library and convert it to a more
        '\ easily used ADO.NET datatable which can be used for data binding much easier.
        Dim dtTable As New System.Data.DataTable
        Dim NewCol As System.Data.DataColumn
        Dim NewRow As System.Data.DataRow
        Dim ColCount As Integer
        Try
            For ColCount = 0 To SAPRecordset.Fields.Count - 1
                NewCol = New System.Data.DataColumn(SAPRecordset.Fields.Item(ColCount).Name)
                dtTable.Columns.Add(NewCol)
            Next

            Do Until SAPRecordset.EoF

                NewRow = dtTable.NewRow
                'populate each column in the row we're creating
                For ColCount = 0 To SAPRecordset.Fields.Count - 1

                    NewRow.Item(SAPRecordset.Fields.Item(ColCount).Name) = SAPRecordset.Fields.Item(ColCount).Value

                Next

                'Add the row to the datatable
                dtTable.Rows.Add(NewRow)


                SAPRecordset.MoveNext()
            Loop

            Return dtTable

        Catch ex As Exception
            MsgBox(ex.Message & Chr(10) & "Error converting SAP Recordset to DataTable", MsgBoxStyle.Exclamation)
            Exit Function
        End Try
    End Function

#Region " ConvertDate"
    Function ConvertDate(ByVal _date As String) As Date
        Try
            Return getSingleValue(" Select Convert(DateTime,'" & _date & "') Dt ")

        Catch ex As Exception
            oApplication.StatusBar.SetText(" ConvertDate Function Failed :  " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Return ""
        End Try
    End Function
#End Region
End Class
