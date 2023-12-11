Imports System.Net
Imports System.IO
Imports System.Security.Cryptography
Imports SAPbouiCOM.Framework
Module Root
    Public t As Threading.Thread
    'Public WithEvents objApplication As SAPbouiCOM.Application
    Sub Main(ByVal args() As String)
        Try
            oGFun.Intialize(args)

        Catch ex As Exception

        End Try
        'Try
        '    oGFun.SetApplication()
        'oApplication = oGFun.oApplication
        '    If Not oGFun.CookieConnect() = 0 Then
        '        oApplication.MessageBox("DI Api Conection Failed")
        '        End
        '    End If
        '    oGFun.HWKEY = HWKEY
        '    If Not oGFun.ConnectionContext() = 0 Then
        '        System.Windows.Forms.MessageBox.Show("Failed to Connect Company", addonName)
        '        If oGFun.oCompany.Connected Then oGFun.oCompany.Disconnect()
        '        System.Windows.Forms.Application.Exit()
        '        End
        '    End If
        '    oCompany = oGFun.oCompany
        'Catch ex As Exception
        '    System.Windows.Forms.MessageBox.Show("Application Not Found", addonName)
        '    System.Windows.Forms.Application.ExitThread()
        'Finally
        'End Try
        'Try
        '    Try
        '        Dim oTableCreation As New TableCreation
        '        EventHandler.SetEventFilter()
        '        oApplication.StatusBar.SetText(" Menu Creation Starting........", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        '        oGFun.AddXML("Menu.xml")
        '        oApplication.StatusBar.SetText("Menu Creation Successs.......", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        '    Catch ex As Exception
        '        System.Windows.Forms.MessageBox.Show(ex.Message)
        '        System.Windows.Forms.Application.ExitThread()
        '    Finally
        '    End Try
        '    oApplication.StatusBar.SetText("Connected.......", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        '    'Application.Run()
        'Catch ex As Exception
        '    oApplication.StatusBar.SetText(addonName & " Main Method Failed : ", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        'Finally
        'End Try
        'oProductionPlan.addReporttype()
    End Sub


    'Function isValidLicense() As Boolean
    '    Try

    '        oApplication.Menus.Item("257").Activate()
    '        Dim CrrHWKEY As String = oApplication.Forms.ActiveForm.Items.Item("79").Specific.Value.ToString.Trim
    '        oApplication.Forms.ActiveForm.Close()
    '        For i As Integer = 0 To HWKEY.Length - 1
    '            If HWKEY(i).Trim = CrrHWKEY.Trim Then
    '                Return True
    '            End If
    '        Next
    '        MsgBox("Add-on installation failed due to license mismatch", MsgBoxStyle.OkOnly, "License Management")
    '        Return False
    '    Catch ex As Exception
    '        oApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
    '    End Try
    '    Return True
    'End Function

End Module



