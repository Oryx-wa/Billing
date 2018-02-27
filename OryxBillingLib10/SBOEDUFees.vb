Imports Microsoft.VisualBasic.CompilerServices
Imports SAPbouiCOM
Imports SBO.SboAddOnBase
Imports System
Imports System.Runtime.CompilerServices


Public Class SBOEDUFees
    Inherits SBOBaseObject

    Private txtfee As Column

    Private dgschfes As Matrix

    Private txtyrCode As EditText

    Private txtterm As EditText

    Private keystage As EditText

    Private docentry As EditText

    Private btnCopy As ButtonCombo

    Private btnFees As Button

    Private intCopyValue As Integer

    Public Sub New(pAddOn As SboAddon, pForm As IForm)
        MyBase.New(pAddOn, pForm)
    End Sub

    Protected Overrides Sub AddDataSource()
        MyBase.AddDataSource()
        Me.m_DBDataSource0 = Me.m_Form.DataSources.DBDataSources.Item("@OWA_EDUSCHOOLFEES")
        Me.m_DBDataSource1 = Me.m_Form.DataSources.DBDataSources.Item("@OWA_EDUFEEGROUP")
        Me.m_DBDataSource2 = Me.m_Form.DataSources.DBDataSources.Item("@OWA_EDUSCHFEEROWS")
        Me.m_DBDataSource3 = Me.m_Form.DataSources.DBDataSources.Item("@OWA_EDUTERM")
        Me.m_DBDataSource4 = Me.m_Form.DataSources.DBDataSources.Item("@OWA_EDUSESSIONS")
        Me.m_DataTable0 = Me.m_Form.DataSources.DataTables.Item("DT_ROWS")
    End Sub

    Public Overrides Sub OnCustomInit()
        MyBase.OnCustomInit()
    End Sub

    Protected Overrides Sub OnFormInit()
        MyBase.OnFormInit()
        Me.dgschfes = CType(Me.m_Form.Items.Item("dgschfes").Specific, Matrix)
        Me.txtfee = Me.dgschfes.Columns.Item("colFee")
        Me.btnCopy = CType(Me.m_Form.Items.Item("btnCopy").Specific, ButtonCombo)
        Me.btnFees = CType(Me.m_Form.Items.Item("btnFees").Specific, Button)
        Me.btnCopy.Item.AffectsFormMode = False
        Me.btnCopy.ExpandType = BoExpandType.et_DescriptionOnly
        Me.btnCopy.ValidValues.Add("1", "Copy To")
        Me.btnCopy.ValidValues.Add("2", "Copy From")
        Me.btnCopy.Item.SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 1, BoModeVisualBehavior.mvb_True)
        Me.btnCopy.Item.SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 14, BoModeVisualBehavior.mvb_False)
        Me.intCopyValue = -1
    End Sub

    Public Overrides Sub OnChooseFromListAfter(sboObject As Object, pVal As SBOItemEventArg)
        MyBase.OnChooseFromListAfter(RuntimeHelpers.GetObjectValue(sboObject), pVal)
        ' The following expression was wrapped in a checked-statement
        Try
            Dim dataTable As DataTable = Nothing
            Dim arg_20_1 As String = pVal.FormUID
            Dim flag As Boolean = False
            Dim text As String = Me.HandleChooseFromListEvent(arg_20_1, pVal, flag)
            Dim flag2 As Boolean = String.IsNullOrEmpty(text)
            If Not flag2 Then
                Dim itemUID As String = pVal.ItemUID
                Dim arg_79_0 As Boolean
                If Operators.CompareString(itemUID, "keystage", False) <> 0 AndAlso Operators.CompareString(itemUID, "txtyrCode", False) <> 0 Then
                    If Operators.CompareString(itemUID, "btnFees", False) <> 0 Then
                        arg_79_0 = False
                        GoTo IL_78
                    End If
                End If
                arg_79_0 = True
IL_78:
                flag2 = arg_79_0
                If flag2 Then
                    Dim arg_8D_1 As String = pVal.FormUID
                    flag = False
                    text = Me.HandleChooseFromListEvent(arg_8D_1, pVal, flag)
                    flag2 = String.IsNullOrEmpty(text)
                    If flag2 Then
                        Return
                    End If
                Else
                    Dim arg_B8_1 As String = pVal.FormUID
                    flag = False
                    Dim flag3 As Boolean
                    dataTable = Me.HandleChooseFromListEvent(arg_B8_1, pVal, flag, flag3)
                    flag2 = (dataTable Is Nothing)
                    If flag2 Then
                        Return
                    End If
                End If
                Dim itemUID2 As String = pVal.ItemUID
                flag2 = (Operators.CompareString(itemUID2, "keystage", False) = 0)
                Dim flag4 As Boolean
                If flag2 Then
                    Me.m_DBDataSource0.SetValue("U_FeeGroup", Me.m_DBDataSource0.Offset, text)
                Else
                    flag2 = (Operators.CompareString(itemUID2, "txtyrCode", False) = 0)
                    If flag2 Then
                        Me.m_DBDataSource0.SetValue("U_yearCode", Me.m_DBDataSource0.Offset, text)
                    Else
                        flag2 = (Operators.CompareString(itemUID2, "dgschfes", False) = 0)
                        If flag2 Then
                            Dim colUID As String = pVal.ColUID
                            flag4 = (Operators.CompareString(colUID, "colFee", False) = 0)
                            If flag4 Then
                                Try
                                    Me.m_Form.Freeze(True)
                                    Me.dgschfes.FlushToDataSource()
                                    Dim num As Integer = Me.m_DBDataSource2.Size
                                    Dim num2 As Integer = pVal.Row
                                    Dim arg_1C9_0 As Integer = 0
                                    Dim num3 As Integer = dataTable.Rows.Count - 1
                                    Dim num4 As Integer = arg_1C9_0
                                    While True
                                        Dim arg_280_0 As Integer = num4
                                        Dim num5 As Integer = num3
                                        If arg_280_0 > num5 Then
                                            Exit While
                                        End If
                                        flag4 = (num4 = 0)
                                        If flag4 Then
                                            num2 = pVal.Row - 1
                                        Else
                                            Me.m_DBDataSource2.InsertRecord(num2)
                                        End If
                                        Dim text2 As String = Conversions.ToString(dataTable.GetValue("AcctCode", num4))
                                        Me.m_DBDataSource2.SetValue("U_acccode", num2, text2)
                                        Me.m_DBDataSource2.SetValue("U_feeItemID", num2, text2)
                                        Me.getOffset(text2, "AcctCode", DBDS("OACT"))
                                        Me.m_DBDataSource2.SetValue("U_itemDescription", num2, DBDS("OACT").GetValue("AcctName", 0).Trim())
                                        num += 1
                                        num2 += 1
                                        num4 += 1
                                    End While
                                    Me.dgschfes.LoadFromDataSource()
                                    Me.m_Form.Freeze(False)
                                Catch expr_2A0 As Exception
                                    ProjectData.SetProjectError(expr_2A0)
                                    Dim ex As Exception = expr_2A0
                                    Me.m_SboApplication.StatusBar.SetText(ex.Message, BoMessageTime.bmt_Long, BoStatusBarMessageType.smt_Error)
                                    ProjectData.ClearProjectError()
                                Finally
                                    Me.m_Form.Freeze(False)
                                End Try
                            End If
                        Else
                            flag4 = (Operators.CompareString(itemUID2, "btnFees", False) = 0)
                            If flag4 Then
                                flag2 = (Me.intCopyValue = 1)
                                If flag2 Then
                                End If
                            End If
                        End If
                    End If
                End If
                flag4 = (Me.m_Form.Mode = BoFormMode.fm_OK_MODE)
                If flag4 Then
                    Me.m_Form.Mode = BoFormMode.fm_UPDATE_MODE
                End If
            End If
        Catch expr_32E As Exception
            ProjectData.SetProjectError(expr_32E)
            Dim ex2 As Exception = expr_32E
            Me.m_SboApplication.StatusBar.SetText(ex2.Message.ToString(), BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Error)
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
        Dim flag As Boolean = Not MyBase.IsReady(pErrNo, pErrMsg)
        ' The following expression was wrapped in a checked-statement
        Dim result As Boolean
        If flag Then
            result = False
        Else
            Dim num As Integer = Me.m_DBDataSource2.Size - 1
            While True
                Dim arg_68_0 As Integer = num
                Dim num2 As Integer = 0
                If arg_68_0 < num2 Then
                    Exit While
                End If
                flag = (Operators.CompareString(Me.m_DBDataSource2.GetValue("U_acccode", num), "", False) = 0)
                If flag Then
                    Me.m_DBDataSource2.RemoveRecord(num)
                End If
                num += -1
            End While
            Dim num3 As Integer = 0
            flag = Not Me.validateDataSource(Me.m_DBDataSource2, "U_acccode", num3)
            If flag Then
                pErrMsg = "Duplicate item codes exits!"
                result = False
            Else
                flag = (Me.m_Form.Mode = BoFormMode.fm_ADD_MODE)
                If flag Then
                    Dim text As String = "Select COUNT(*) nCnt from [@OWA_EDUSCHOOLFEES] Where U_keystage = 'x1'  and U_yearCode = 'x3'"
                    text = text.Replace("x1", Me.m_DBDataSource0.GetValue("U_keystage", Me.m_DBDataSource0.Offset).Trim())
                    text = text.Replace("x3", Me.m_DBDataSource0.GetValue("U_yearCode", Me.m_DBDataSource0.Offset).Trim())
                    Dim dataTable As DataTable = Me.ExecuteSQLDT(text)
                    flag = Operators.ConditionalCompareObjectNotEqual(dataTable.GetValue(0, 0), 0, False)
                    If flag Then
                        pErrMsg = "This entry already exists"
                        result = False
                        Return result
                    End If
                End If
                result = True
            End If
        End If
        Return result
    End Function

    Protected Overrides Sub EnableToolBarButtons()
        MyBase.EnableToolBarButtons()
        Me.m_Form.EnableMenu("1293", True)
    End Sub

    Protected Overrides Sub SetConditions()
        MyBase.SetConditions()
        Dim condition As Condition = Me.m_Condition
        condition.[Alias] = "Postable"
        condition.Operation = BoConditionOperation.co_EQUAL
        condition.CondVal = "Y"
        condition.Relationship = BoConditionRelationship.cr_AND
        Me.m_Condition = Me.m_Conditions.Add()
        Dim condition2 As Condition = Me.m_Condition
        condition2.[Alias] = "GroupMask"
        condition2.Operation = BoConditionOperation.co_EQUAL
        condition2.CondVal = "4"
        Me.m_Form.ChooseFromLists.Item("cflItems").SetConditions(Me.m_Conditions)
    End Sub

    Public Overrides Sub OnComboSelectAfter(sboObject As Object, pVal As SBOItemEventArg)
        MyBase.OnComboSelectAfter(RuntimeHelpers.GetObjectValue(sboObject), pVal)
        ' The following expression was wrapped in a checked-statement
        Try
            Dim itemUID As String = pVal.ItemUID
            Dim flag As Boolean = Operators.CompareString(itemUID, "btnCopy", False) = 0
            If flag Then
                Dim selected As ValidValue = Me.btnCopy.Selected
                Me.m_Form.Freeze(True)
                flag = (Operators.CompareString(selected.Value, "2", False) = 0)
                If flag Then
                    Me.btnCopy.Caption = "Copy From"
                    Me.intCopyValue = Conversions.ToInteger(selected.Value)
                    Me.btnFees.Item.Visible = True
                    Me.btnFees.Item.Click(BoCellClickType.ct_Regular)
                    Me.btnFees.Item.Visible = False
                Else
                    Me.btnCopy.Caption = "Copy To"
                    Me.m_DataTable0.Rows.Clear()
                    Dim num As Integer = Me.m_DBDataSource2.Size - 1
                    Me.m_DataTable0.Rows.Add(Me.m_DBDataSource2.Size)
                    Dim arg_115_0 As Integer = 0
                    Dim num2 As Integer = num
                    Dim num3 As Integer = arg_115_0
                    While True
                        Dim arg_18D_0 As Integer = num3
                        Dim num4 As Integer = num2
                        If arg_18D_0 > num4 Then
                            Exit While
                        End If
                        Me.m_DataTable0.SetValue("ItemCode", num3, Me.m_DBDataSource2.GetValue("U_feeItemID", num3))
                        Me.m_DataTable0.SetValue("ItemName", num3, Me.m_DBDataSource2.GetValue("U_itemDescription", num3))
                        Me.m_DataTable0.SetValue("Amount", num3, Me.m_DBDataSource2.GetValue("U_feeAmount", num3))
                        num3 += 1
                    End While
                    Me.m_SboApplication.ActivateMenuItem("1282")
                    Me.m_DBDataSource2.Clear()
                    Dim arg_1B0_0 As Integer = 0
                    Dim num5 As Integer = num
                    num3 = arg_1B0_0
                    While True
                        Dim arg_247_0 As Integer = num3
                        Dim num4 As Integer = num5
                        If arg_247_0 > num4 Then
                            Exit While
                        End If
                        Me.m_DBDataSource2.InsertRecord(num3)
                        Me.m_DBDataSource2.SetValue("U_feeItemID", num3, Conversions.ToString(Me.m_DataTable0.GetValue("ItemCode", num3)))
                        Me.m_DBDataSource2.SetValue("U_itemDescription", num3, Conversions.ToString(Me.m_DataTable0.GetValue("ItemName", num3)))
                        Me.m_DBDataSource2.SetValue("U_feeAmount", num3, Conversions.ToString(Me.m_DataTable0.GetValue("Amount", num3)))
                        num3 += 1
                    End While
                    Me.dgschfes.LoadFromDataSource()
                    Me.m_SboApplication.StatusBar.SetText("Enter keystage session and term information", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
                End If
                Me.m_Form.Freeze(False)
            End If
        Catch expr_281 As Exception
            ProjectData.SetProjectError(expr_281)
            Dim ex As Exception = expr_281
            Me.m_SboApplication.StatusBar.SetText(ex.ToString(), BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
            Me.m_Form.Freeze(False)
            ProjectData.ClearProjectError()
        End Try
    End Sub

    Public Overrides Sub OnComboSelectBefore(sboObject As Object, pVal As SBOItemEventArg, ByRef BubbleEvent As Boolean)
        MyBase.OnComboSelectBefore(RuntimeHelpers.GetObjectValue(sboObject), pVal, BubbleEvent)
        Dim itemUID As String = pVal.ItemUID
        Dim flag As Boolean = Operators.CompareString(itemUID, "btnCopy", False) = 0
        If flag Then
            Dim selected As ValidValue = Me.btnCopy.Selected
            Select Case Me.m_Form.Mode
                Case BoFormMode.fm_ADD_MODE
                    flag = (Operators.CompareString(selected.Value, "2", False) = 0)
                    If flag Then
                        BubbleEvent = False
                    Else
                        BubbleEvent = True
                    End If
            End Select
            flag = (Me.m_Form.Mode = BoFormMode.fm_ADD_MODE)
            If flag Then
            End If
        End If
    End Sub

    Private Sub CalculateTotal()
        Try
            Me.dgschfes.FlushToDataSource()
            Dim arg_1D_0 As Integer = 0
            ' The following expression was wrapped in a checked-expression
            Dim num As Integer = Me.m_DBDataSource2.Size - 1
            Dim num2 As Integer = arg_1D_0
            Dim num4 As Double
            Dim feeP, feeInt, Topup, TopupC, Sandwich As Double
            While True
                Dim arg_44_0 As Integer = num2
                Dim num3 As Integer = num
                If arg_44_0 > num3 Then
                    Exit While
                End If
                num4 += Conversions.ToDouble(Me.m_DBDataSource2.GetValue("U_feeAmount", num2))
                feeP += Conversions.ToDouble(Me.m_DBDataSource2.GetValue("U_feeAmountP", num2))
                feeInt += Conversions.ToDouble(Me.m_DBDataSource2.GetValue("U_feeAmountInt", num2))
                Topup += Conversions.ToDouble(Me.m_DBDataSource2.GetValue("U_feeAmtTp", num2))
                TopupC += Conversions.ToDouble(Me.m_DBDataSource2.GetValue("U_feeAmtTpC", num2))
                Sandwich += Conversions.ToDouble(Me.m_DBDataSource2.GetValue("U_feeSand", num2))

                ' The following expression was wrapped in a checked-statement
                num2 += 1
            End While
            Me.m_Form.DataSources.UserDataSources.Item("UD_Total").ValueEx = Conversions.ToString(num4)
            Me.m_Form.DataSources.UserDataSources.Item("UD_Fee").ValueEx = Conversions.ToString(feeP)
            Me.m_Form.DataSources.UserDataSources.Item("UD_Int").ValueEx = Conversions.ToString(feeInt)
            Me.m_Form.DataSources.UserDataSources.Item("UD_Topup").ValueEx = Conversions.ToString(Topup)
            Me.m_Form.DataSources.UserDataSources.Item("UD_Tpc").ValueEx = Conversions.ToString(TopupC)
            Me.m_Form.DataSources.UserDataSources.Item("UD_Sand").ValueEx = Conversions.ToString(Sandwich)
        Catch expr_6E As Exception
            ProjectData.SetProjectError(expr_6E)
            Dim ex As Exception = expr_6E
            Me.m_SboApplication.StatusBar.SetText(ex.ToString(), BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
            Me.m_Form.DataSources.UserDataSources.Item("UD_Total").ValueEx = Conversions.ToString(0)
            ProjectData.ClearProjectError()
        End Try
    End Sub

    Protected Overrides Sub OnFormNavigate()
        MyBase.OnFormNavigate()
        Me.CalculateTotal()
    End Sub

    Public Overrides Sub OnDataUpdateAfter(ByRef pVal As BusinessObjectInfo)
        MyBase.OnDataUpdateAfter(pVal)
        Me.CalculateTotal()
    End Sub

    Public Overrides Sub OnDataAddAfter(ByRef pVal As BusinessObjectInfo)
        MyBase.OnDataAddAfter(pVal)
        Me.CalculateTotal()
    End Sub
    Public Overrides Sub OnDataAddBefore(ByRef pVal As BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Me.m_DBDataSource2.Clear()
        Me.dgschfes.FlushToDataSource()
        MyBase.OnDataAddBefore(pVal, BubbleEvent)
    End Sub

    Public Overrides Sub OnDataUpdateBefore(ByRef pVal As BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Me.m_DBDataSource2.Clear()
        Me.dgschfes.FlushToDataSource()
        MyBase.OnDataUpdateBefore(pVal, BubbleEvent)
    End Sub

    Public Overrides Sub OnItemValidateAfter(sboObject As Object, pVal As SBOItemEventArg)
        MyBase.OnItemValidateAfter(RuntimeHelpers.GetObjectValue(sboObject), pVal)
        Dim itemUID As String = pVal.ItemUID
        Dim flag As Boolean = Operators.CompareString(itemUID, "dgschfes", False) = 0
        If flag Then
            Dim flag2 As Boolean = Operators.CompareString(pVal.ColUID, "C_0_4", False) = 0
            If flag2 Then
                Me.CalculateTotal()
            End If
        End If
    End Sub

    Public Overrides Sub OnItemPressedBefore(sboObject As Object, pVal As SBOItemEventArg, ByRef BubbleEvent As Boolean)

        MyBase.OnItemPressedBefore(sboObject, pVal, BubbleEvent)
        Dim flag As Boolean = pVal.Row = Me.dgschfes.RowCount + 1
        If flag Then


        End If

    End Sub
End Class

