Imports Microsoft.VisualBasic.CompilerServices
Imports SAPbobsCOM
Imports SAPbouiCOM
Imports SBO.SboAddOnBase
Imports System
Imports System.Runtime.CompilerServices
Imports System.Runtime.InteropServices


Public Class SBOEDUConfig
    Inherits SBOBaseObject

    Public Sub New(pAddOn As SboAddon, pForm As IForm)
        MyBase.New(pAddOn, pForm)
    End Sub

    Protected Overrides Function IsReady(ByRef pErrNo As Integer, ByRef pErrMsg As String) As Boolean
        Return MyBase.IsReady(pErrNo, pErrMsg)
    End Function

    Protected Overrides Sub AddDataSource()
        MyBase.AddDataSource()
        Me.m_DBDataSource0 = Me.m_Form.DataSources.DBDataSources.Item("@OWA_EDUCONFIG")
        Dim flag As Boolean = Not Me.getOffset("ORYX", "Code", Me.m_DBDataSource0)
        If flag Then
            Me.m_DBDataSource0.InsertRecord(0)
        End If
    End Sub

    Protected Overrides Sub EnableToolBarButtons()
        MyBase.EnableToolBarButtons()
    End Sub

    Public Overrides Sub OnItemClickAfter(sboObject As Object, pVal As SBOItemEventArg)
        MyBase.OnItemClickAfter(RuntimeHelpers.GetObjectValue(sboObject), pVal)
        Dim itemUID As String = pVal.ItemUID
        Dim flag As Boolean = Operators.CompareString(itemUID, "btnSave", False) = 0
        If flag Then
            Me.SaveConfig()
        End If
    End Sub

    Protected Sub SaveConfig()
        Try
            Dim generalService As GeneralService = Me.m_ParentAddon.SboCompany.GetCompanyService().GetGeneralService("OWAEDUConfig")
            Dim generalData As GeneralData = CType(generalService.GetDataInterface(GeneralServiceDataInterfaces.gsGeneralData), GeneralData)
            Dim generalDataParams As GeneralDataParams = CType(generalService.GetDataInterface(GeneralServiceDataInterfaces.gsGeneralDataParams), GeneralDataParams)
            Dim flag As Boolean
            Try
                generalDataParams.SetProperty("Code", "ORYX")
                generalData = generalService.GetByParams(generalDataParams)
            Catch expr_53 As Exception
                ProjectData.SetProjectError(expr_53)
                flag = True
                ProjectData.ClearProjectError()
            End Try
            Dim flag2 As Boolean = flag
            If flag2 Then
                generalData.SetProperty("Code", "ORYX")
                generalData.SetProperty("Name", "ORYX")
            End If
            Dim arg_A6_0 As Short = 0S
            ' The following expression was wrapped in a checked-expression
            Dim num As Short = CShort((Me.m_DBDataSource0.Fields.Count - 1))
            Dim num2 As Short = arg_A6_0
            While True
                Dim arg_110_0 As Short = num2
                Dim num3 As Short = num
                If arg_110_0 > num3 Then
                    Exit While
                End If
                Dim name As String = Me.m_DBDataSource0.Fields.Item(num2).Name
                Dim vtValue As String = Me.m_DBDataSource0.GetValue(name, 0).Trim()
                flag2 = Not name.StartsWith("U_")
                If Not flag2 Then
                    generalData.SetProperty(name, vtValue)
                End If
                num2 += 1S
            End While
            flag2 = flag
            If flag2 Then
                generalService.Add(generalData)
            Else
                generalService.Update(generalData)
            End If
            Marshal.FinalReleaseComObject(generalData)
            Marshal.FinalReleaseComObject(generalDataParams)
            Marshal.FinalReleaseComObject(generalService)
            Me.m_SboApplication.StatusBar.SetText("Data Saved Successfully", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success)
        Catch expr_15C As Exception
            ProjectData.SetProjectError(expr_15C)
            Dim ex As Exception = expr_15C
            Me.m_SboApplication.StatusBar.SetText(ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
            ProjectData.ClearProjectError()
        End Try
    End Sub
End Class

