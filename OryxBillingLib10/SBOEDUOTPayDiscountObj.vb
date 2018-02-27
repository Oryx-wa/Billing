Imports Microsoft.VisualBasic.CompilerServices
Imports SAPbobsCOM
Imports SAPbouiCOM
Imports SBO.SboAddOnBase
Imports System
Imports System.Runtime.CompilerServices


Public Class SBOEDUOTPayDiscountObj
    Inherits SBOBaseObject

    Private lblName As StaticText

    Private grdStudCl As Grid

    Private txtSessn As EditText

    Private strSQL As String

    Private editCol As GridColumn

    Private txtCode As EditText

    Private lAdd As Boolean

    Private strDocEntry As String

    Private strClass As String

    Public Sub New(pAddOn As SboAddon, pForm As IForm)
        MyBase.New(pAddOn, pForm)
        Me.strSQL = "Select U_CardCode, U_CardName, DocEntry, 1 AS NewRec from [@OWA_EDUOTPHISTROWS] "
        Me.lAdd = False
    End Sub

    Protected Overrides Sub EnableToolBarButtons()
        MyBase.EnableToolBarButtons()
    End Sub

    Protected Overrides Sub AddDataSource()
        MyBase.AddDataSource()
        Me.m_DBDataSource0 = Me.m_Form.DataSources.DBDataSources.Item("@OWA_EDUOTPHISTORY")
        Me.m_DBDataSource1 = Me.m_Form.DataSources.DBDataSources.Item("@OWA_EDUOTPHISTROWS")
        Me.m_DBDataSource3 = Me.m_Form.DataSources.DBDataSources.Item("OCRD")
        Me.m_DataTable0 = Me.m_Form.DataSources.DataTables.Item("DT_0")
    End Sub

    Public Overrides Sub OnChooseFromListAfter(sboObject As Object, pVal As SBOItemEventArg)
        MyBase.OnChooseFromListAfter(RuntimeHelpers.GetObjectValue(sboObject), pVal)
        ' The following expression was wrapped in a checked-statement
        Try
            Dim dataTable As DataTable = Nothing
            Dim itemUID As String = pVal.ItemUID
            Dim flag As Boolean = Operators.CompareString(itemUID, "btnAdd", False) = 0
            Dim flag3 As Boolean
            Dim text As String
            If flag Then
                Dim arg_43_1 As String = pVal.FormUID
                Dim flag2 As Boolean = False
                dataTable = Me.HandleChooseFromListEvent(arg_43_1, pVal, flag2, flag3)
                flag = (dataTable Is Nothing)
                If flag Then
                    Return
                End If
            Else
                Dim arg_6A_1 As String = pVal.FormUID
                Dim flag2 As Boolean = False
                text = Me.HandleChooseFromListEvent(arg_6A_1, pVal, flag2)
                flag = String.IsNullOrEmpty(text)
                If flag Then
                    Return
                End If
            End If
            Dim itemUID2 As String = pVal.ItemUID
            flag = (Operators.CompareString(itemUID2, "txtSessn", False) = 0)
            If flag Then
                Dim flag4 As Boolean = Operators.CompareString(text, Nothing, False) <> 0
                If flag4 Then
                    Me.m_DBDataSource0.SetValue("U_Session", Me.m_DBDataSource0.Offset, text)
                End If
            Else
                Dim flag4 As Boolean = Operators.CompareString(itemUID2, "btnAdd", False) = 0
                If flag4 Then
                    flag = flag3
                    If flag Then
                        Dim num As Integer = Me.m_DataTable0.Rows.Count
                        Dim arg_122_0 As Integer = 0
                        Dim num2 As Integer = dataTable.Rows.Count - 1
                        Dim num3 As Integer = arg_122_0
                        While True
                            Dim arg_1DE_0 As Integer = num3
                            Dim num4 As Integer = num2
                            If arg_1DE_0 > num4 Then
                                Exit While
                            End If
                            Me.m_DataTable0.Rows.Add(1)
                            Dim text2 As String = Conversions.ToString(dataTable.GetValue("CardCode", num3))
                            Me.m_DataTable0.SetValue("U_CardCode", num, text2)
                            flag4 = Me.getOffset(text2, "CardCode", Me.m_DBDataSource3)
                            If flag4 Then
                                Dim value As String = Me.m_DBDataSource3.GetValue("CardName", Me.m_DBDataSource1.Offset)
                                Me.m_DataTable0.SetValue("U_CardName", num, value.Trim())
                                Me.m_DataTable0.SetValue("NewRec", num, -1)
                            End If
                            num += 1
                            num3 += 1
                        End While
                        flag4 = Operators.ConditionalCompareObjectEqual(Me.m_DataTable0.GetValue(0, 0), "", False)
                        If flag4 Then
                            Me.m_DataTable0.Rows.Remove(0)
                        End If
                        flag4 = (Me.m_Form.Mode = BoFormMode.fm_OK_MODE)
                        If flag4 Then
                            Me.m_Form.Mode = BoFormMode.fm_UPDATE_MODE
                        End If
                    End If
                End If
            End If
        Catch expr_23F As Exception
            ProjectData.SetProjectError(expr_23F)
            Dim ex As Exception = expr_23F
            Me.m_SboApplication.StatusBar.SetText(ex.Message.ToString(), BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Error)
            ProjectData.ClearProjectError()
        End Try
    End Sub

    Protected Overrides Sub OnFormNavigate()
        MyBase.OnFormNavigate()
        Dim flag As Boolean = Me.m_Form.Mode <> BoFormMode.fm_ADD_MODE
        If flag Then
            Me.strDocEntry = Me.m_DBDataSource0.GetValue("DocEntry", Me.m_DBDataSource0.Offset)
            Me.strClass = Me.m_DBDataSource0.GetValue("DocEntry", Me.m_DBDataSource0.Offset)
        Else
            Me.strDocEntry = ""
            Me.strClass = ""
        End If
        Me.FormatGrid()
    End Sub

    Public Overrides Sub OnCustomInit()
        Me.grdStudCl = CType(Me.m_Form.Items.Item("grdStudCl").Specific, Grid)
        Me.txtSessn = CType(Me.m_Form.Items.Item("txtSessn").Specific, EditText)
        MyBase.OnCustomInit()
    End Sub

    Protected Overrides Sub FormatGrid()
        Try
            Dim text As String = Me.m_DBDataSource0.GetValue("DocEntry", Me.m_DBDataSource0.Offset)
            Dim str As String = " Where DocEntry = "
            Dim flag As Boolean = Operators.CompareString(text, String.Empty, False) = 0
            If flag Then
                text = "-1"
            End If
            Me.m_Form.Freeze(True)
            Me.m_DataTable0.Clear()
            Me.m_DataTable0.ExecuteQuery(Me.strSQL + str + text)
            Me.grdStudCl.DataTable = Me.m_DataTable0
            Me.grdStudCl.Columns.Item("U_CardCode").TitleObject.Caption = "Student Code"
            Me.grdStudCl.Columns.Item("U_CardName").TitleObject.Caption = "Student Name"
            Me.grdStudCl.Columns.Item("DocEntry").Visible = False
            Me.grdStudCl.Columns.Item("NewRec").Visible = False
            Me.grdStudCl.Columns.Item("U_CardCode").Editable = False
            Me.grdStudCl.Columns.Item("U_CardName").Editable = False
            Me.editCol = Me.grdStudCl.Columns.Item("U_CardCode")
            NewLateBinding.LateSet(Me.editCol, Nothing, "LinkedObjectType", New Object() {2}, Nothing, Nothing)
            Me.editCol.AffectsFormMode = True
            Me.grdStudCl.AutoResizeColumns()
            Me.m_Form.Freeze(False)
        Catch expr_1A7 As Exception
            ProjectData.SetProjectError(expr_1A7)
            Dim ex As Exception = expr_1A7
            Me.m_ParentAddon.SboApplication.SetStatusBarMessage(ex.Message.ToString(), BoMessageTime.bmt_Medium, True)
            ProjectData.ClearProjectError()
        Finally
            Me.m_Form.Freeze(False)
        End Try
    End Sub

    Public Overrides Sub OnItemClickAfter(sboObject As Object, pVal As SBOItemEventArg)
        MyBase.OnItemClickAfter(RuntimeHelpers.GetObjectValue(sboObject), pVal)
        ' The following expression was wrapped in a checked-statement
        Try
            Dim itemUID As String = pVal.ItemUID
            Dim flag As Boolean = Operators.CompareString(itemUID, "1", False) = 0
            If Not flag Then
                flag = (Operators.CompareString(itemUID, "btnRem", False) = 0)
                If flag Then
                    Dim selectedRows As SelectedRows = Me.grdStudCl.Rows.SelectedRows
                    flag = (selectedRows.Count > 0)
                    If flag Then
                        Dim arg_79_0 As Integer = 0
                        Dim num As Integer = selectedRows.Count - 1
                        Dim num2 As Integer = arg_79_0
                        While True
                            Dim arg_9A_0 As Integer = num2
                            Dim num3 As Integer = num
                            If arg_9A_0 > num3 Then
                                Exit While
                            End If
                            Me.m_DataTable0.Rows.Remove(num2)
                            num2 += 1
                        End While
                        flag = (Me.m_Form.Mode = BoFormMode.fm_OK_MODE And Me.m_Form.Mode <> BoFormMode.fm_ADD_MODE)
                        If flag Then
                            Me.m_Form.Mode = BoFormMode.fm_UPDATE_MODE
                        End If
                    End If
                End If
            End If
        Catch expr_D4 As Exception
            ProjectData.SetProjectError(expr_D4)
            ProjectData.ClearProjectError()
        End Try
    End Sub

    Private Sub CommitDataTable()
        ' The following expression was wrapped in a checked-statement
        Try
            Dim generalService As GeneralService = Me.m_ParentAddon.SboCompany.GetCompanyService().GetGeneralService("OWAOTPHISTORY")
            Dim generalDataParams As GeneralDataParams = CType(generalService.GetDataInterface(GeneralServiceDataInterfaces.gsGeneralDataParams), GeneralDataParams)
            generalDataParams.SetProperty("DocEntry", Me.strDocEntry)
            Dim byParams As GeneralData = generalService.GetByParams(generalDataParams)
            Dim generalDataCollection As GeneralDataCollection = byParams.Child("OWA_EDUOTPHISTROWS")
            Dim arg_6B_0 As Integer = 0
            Dim num As Integer = Me.m_DataTable0.Rows.Count - 1
            Dim num2 As Integer = arg_6B_0
            While True
                Dim arg_F8_0 As Integer = num2
                Dim num3 As Integer = num
                If arg_F8_0 > num3 Then
                    Exit While
                End If
                Dim flag As Boolean = Operators.ConditionalCompareObjectEqual(Me.m_DataTable0.GetValue("NewRec", num2), -1, False)
                If flag Then
                    Dim generalData As GeneralData = generalDataCollection.Add()
                    Dim generalData2 As GeneralData = generalData
                    Dim vtValue As String = Conversions.ToString(Me.m_DataTable0.GetValue("U_CardCode", num2))
                    generalData2.SetProperty("U_CardCode", vtValue)
                    generalData2.SetProperty("U_CardName", RuntimeHelpers.GetObjectValue(Me.m_DataTable0.GetValue("U_CardName", num2)))
                End If
                num2 += 1
            End While
            generalService.Update(byParams)
        Catch expr_108 As Exception
            ProjectData.SetProjectError(expr_108)
            Dim ex As Exception = expr_108
            Me.m_ParentAddon.SboApplication.StatusBar.SetText(ex.Message.ToString(), BoMessageTime.bmt_Long, BoStatusBarMessageType.smt_Error)
            ProjectData.ClearProjectError()
        End Try
    End Sub

    Protected Overrides Function Save(ByRef pErrNo As Integer, ByRef pErrMsg As String) As Boolean
        Dim flag As Boolean = Not MyBase.Save(pErrNo, pErrMsg)
        Dim result As Boolean
        If flag Then
            Me.m_ParentAddon.SboApplication.SetStatusBarMessage(pErrMsg, BoMessageTime.bmt_Medium, True)
            result = False
        Else
            result = True
        End If
        Return result
    End Function

    Protected Overrides Function IsReady(ByRef pErrNo As Integer, ByRef pErrMsg As String) As Boolean
        Dim num As Integer = 0
        Dim flag As Boolean = Not Me.validateDataSource(Me.m_DataTable0, "U_CardCode", num)
        Dim result As Boolean
        If flag Then
            pErrMsg = "Duplicate item codes exits!"
            result = False
        Else
            result = True
        End If
        Return result
    End Function

    Protected Overrides Sub SetConditions()
        MyBase.SetConditions()
        Dim condition As Condition = Me.m_Condition
        condition.[Alias] = "QryGroup1"
        condition.Operation = BoConditionOperation.co_EQUAL
        condition.CondVal = "Y"
        condition.Relationship = BoConditionRelationship.cr_AND
        Me.m_Condition = Me.m_Conditions.Add()
        Dim condition2 As Condition = Me.m_Condition
        condition2.[Alias] = "FrozenFor"
        condition2.Operation = BoConditionOperation.co_NOT_EQUAL
        condition2.CondVal = "Y"
        Me.m_Form.ChooseFromLists.Item("CFL_OCRD").SetConditions(Me.m_Conditions)
    End Sub

    Public Overrides Sub OnDataAddAfter(ByRef pVal As BusinessObjectInfo)
        MyBase.OnDataAddAfter(pVal)
        Try
            Dim query As String = "Select Max(DocEntry) DocEntry from [@OWA_EDUOTPHISTORY] "
            Dim actionSuccess As Boolean = pVal.ActionSuccess
            If actionSuccess Then
                DataTable("DT_Base").Clear()
                DataTable("DT_Base").ExecuteQuery(query)
                Me.strDocEntry = Conversions.ToString(DataTable("DT_Base").GetValue(0, 0))
                query = "Delete [@OWA_EDUOTPHISTROWS] Where U_CardCode = 'N/A'"
                DataTable("DT_Base").ExecuteQuery(query)
                Me.CommitDataTable()
            End If
        Catch expr_82 As Exception
            ProjectData.SetProjectError(expr_82)
            ProjectData.ClearProjectError()
        End Try
    End Sub

    Public Overrides Sub OnDataUpdateAfter(ByRef pVal As BusinessObjectInfo)
        MyBase.OnDataUpdateAfter(pVal)
        Me.CommitDataTable()
    End Sub
End Class

