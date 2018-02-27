Imports Microsoft.VisualBasic.CompilerServices
Imports SAPbobsCOM
Imports SAPbouiCOM
Imports SBO.SboAddOnBase
Imports System
Imports System.Runtime.CompilerServices


Public Class SBOEDUClassHistoryObj
    Inherits SBOBaseObject

    Private cboClass As ComboBox

    Private lblName As StaticText

    Private grdStudCl As Grid

    Private txtKeyst As EditText

    Private txtSessn As EditText

    Private strSQL As String

    Private editCol As GridColumn

    Private txtCode As EditText

    Private lAdd As Boolean

    Private strDocEntry As String

    Private strClass As String

    Private btnCopy As ButtonCombo

    Private intCopyValue As Integer

    Private lAdded As Boolean

    Public Sub New(pAddOn As SboAddon, pForm As IForm)
        MyBase.New(pAddOn, pForm)
        Me.strSQL = "Select U_CardCode, U_CardName, Code, U_DocEntry, 1 AS NewRec from [@OWA_EDUCLAHISTROWS] "
        Me.lAdd = False
    End Sub

    Protected Overrides Sub EnableToolBarButtons()
        MyBase.EnableToolBarButtons()
        Me.m_Form.EnableMenu("1292", True)
        Me.m_Form.EnableMenu("1293", True)
    End Sub

    Protected Overrides Sub AddDataSource()
        MyBase.AddDataSource()
        Me.m_DBDataSource0 = Me.m_Form.DataSources.DBDataSources.Item("@OWA_EDUCLAHISTORY")
        Me.m_DBDataSource1 = Me.m_Form.DataSources.DBDataSources.Item("@OWA_EDUCLAHISTROWS")
        Me.m_DBDataSource2 = Me.m_Form.DataSources.DBDataSources.Item("OHEM")
        Me.m_DBDataSource3 = Me.m_Form.DataSources.DBDataSources.Item("OCRD")
        Me.m_DataTable0 = Me.m_Form.DataSources.DataTables.Item("DT_0")
        Me.m_DataTable1 = Me.m_Form.DataSources.DataTables.Item("DT_1")
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
            flag = (Operators.CompareString(itemUID2, "txtKeyst", False) = 0)
            If flag Then
                Me.m_DBDataSource0.SetValue("U_Keystage", Me.m_DBDataSource0.Offset, text)
                Dim strWhere As String = "Code = '" + text + "'"
                Me.fillCombo("U_clsCode", "U_clsDesc", "@OWA_EDUKEYSTALINES", Me.cboClass.ValidValues, strWhere, False, BoFieldTypes.db_Alpha, False, "", True)
            Else
                flag = (Operators.CompareString(itemUID2, "txtSessn", False) = 0)
                If flag Then
                    Me.m_DBDataSource0.SetValue("U_Session", Me.m_DBDataSource0.Offset, text)
                Else
                    flag = (Operators.CompareString(itemUID2, "btnAdd", False) = 0)
                    If flag Then
                        Dim flag4 As Boolean = flag3
                        If flag4 Then
                            Dim num As Integer = Me.m_DataTable0.Rows.Count
                            Dim arg_183_0 As Integer = 0
                            Dim num2 As Integer = dataTable.Rows.Count - 1
                            Dim num3 As Integer = arg_183_0
                            While True
                                Dim arg_2A9_0 As Integer = num3
                                Dim num4 As Integer = num2
                                If arg_2A9_0 > num4 Then
                                    Exit While
                                End If
                                Me.m_DataTable0.Rows.Add(1)
                                Dim text2 As String = Conversions.ToString(dataTable.GetValue("CardCode", num3))
                                Me.m_DataTable0.SetValue("U_CardCode", num, text2)
                                flag4 = Me.getOffset(text2, "CardCode", Me.m_DBDataSource3)
                                If flag4 Then
                                    Dim value As String = Me.m_DBDataSource3.GetValue("CardName", Me.m_DBDataSource2.Offset)
                                    Me.m_DataTable0.SetValue("U_CardName", num, value.Trim())
                                    flag4 = (Me.m_Form.Mode <> BoFormMode.fm_ADD_MODE)
                                    If flag4 Then
                                        Me.m_DataTable0.SetValue("Code", num, Me.strDocEntry.ToString().Trim() + text2.Trim())
                                        Me.m_DataTable0.SetValue("U_DocEntry", num, Me.strDocEntry)
                                    Else
                                        Me.m_DataTable0.SetValue("Code", num, "xNew")
                                    End If
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
                        Me.m_Form.DataSources.UserDataSources.Item("UD_0").ValueEx = Conversions.ToString(Me.m_DataTable0.Rows.Count)
                    Else
                        Dim flag4 As Boolean = Operators.CompareString(itemUID2, "btnFees", False) = 0
                        If flag4 Then
                            flag = (Me.intCopyValue = 1)
                            If flag Then
                            End If
                        End If
                    End If
                End If
            End If
        Catch expr_36C As Exception
            ProjectData.SetProjectError(expr_36C)
            Dim ex As Exception = expr_36C
            Me.m_SboApplication.StatusBar.SetText(ex.Message.ToString(), BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Error)
            ProjectData.ClearProjectError()
        End Try
    End Sub

    Protected Overrides Sub OnFormNavigate()
        MyBase.OnFormNavigate()
        Try
            Dim flag As Boolean = Me.m_Form.Mode <> BoFormMode.fm_ADD_MODE
            If flag Then
                Me.strDocEntry = Me.m_DBDataSource0.GetValue("DocEntry", Me.m_DBDataSource0.Offset)
                Me.strClass = Me.m_DBDataSource0.GetValue("DocEntry", Me.m_DBDataSource0.Offset)
            Else
                Me.strDocEntry = ""
                Me.strClass = ""
            End If
            Dim text As String = Me.m_DBDataSource0.GetValue("DocEntry", Me.m_DBDataSource0.Offset)
            flag = (Operators.CompareString(text, String.Empty, False) = 0)
            If flag Then
                text = "-1"
            End If
            Dim str As String = " Where U_DocEntry = "
            Me.m_Form.Freeze(True)
            Me.m_DataTable0.Clear()
            Me.m_DataTable0.ExecuteQuery(Me.strSQL + str + text)
            Me.grdStudCl.DataTable = Me.m_DataTable0
            Me.FormatGrid()
            Me.m_Form.Freeze(False)
        Catch expr_110 As Exception
            ProjectData.SetProjectError(expr_110)
            Dim ex As Exception = expr_110
            Me.m_SboApplication.StatusBar.SetText(ex.ToString(), BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
            Me.m_Form.Freeze(False)
            ProjectData.ClearProjectError()
        End Try
    End Sub

    Protected Overrides Sub OnFormInit()
        MyBase.OnFormInit()
        Me.cboClass = CType(Me.m_Form.Items.Item("cboClass").Specific, ComboBox)
        Me.grdStudCl = CType(Me.m_Form.Items.Item("grdStudCl").Specific, Grid)
        Me.txtSessn = CType(Me.m_Form.Items.Item("txtSessn").Specific, EditText)
        Me.txtKeyst = CType(Me.m_Form.Items.Item("txtKeyst").Specific, EditText)
        Me.btnCopy = CType(Me.m_Form.Items.Item("btnCopy").Specific, ButtonCombo)
        Me.btnCopy.Item.AffectsFormMode = False
        Me.btnCopy.ExpandType = BoExpandType.et_DescriptionOnly
        Me.btnCopy.ValidValues.Add("1", "Copy To")
        Me.btnCopy.ValidValues.Add("2", "Copy From")
        Me.btnCopy.Item.SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 1, BoModeVisualBehavior.mvb_True)
        Me.btnCopy.Item.SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 14, BoModeVisualBehavior.mvb_False)
        Me.intCopyValue = -1
    End Sub

    Protected Overrides Sub FormatGrid()
        Try
            Me.grdStudCl.Columns.Item("U_CardCode").TitleObject.Caption = "Student Code"
            Me.grdStudCl.Columns.Item("U_CardName").TitleObject.Caption = "Student Name"
            Me.grdStudCl.Columns.Item("U_DocEntry").Visible = False
            Me.grdStudCl.Columns.Item("Code").Visible = False
            Me.grdStudCl.Columns.Item("NewRec").Visible = False
            Me.grdStudCl.Columns.Item("U_CardCode").Editable = False
            Me.grdStudCl.Columns.Item("U_CardName").Editable = False
            Me.editCol = Me.grdStudCl.Columns.Item("U_CardCode")
            NewLateBinding.LateSet(Me.editCol, Nothing, "LinkedObjectType", New Object() {2}, Nothing, Nothing)
            Me.editCol.AffectsFormMode = True
            Me.grdStudCl.AutoResizeColumns()
            Me.m_Form.DataSources.UserDataSources.Item("UD_0").ValueEx = Conversions.ToString(Me.m_DataTable0.Rows.Count)
        Catch expr_169 As Exception
            ProjectData.SetProjectError(expr_169)
            Dim ex As Exception = expr_169
            Me.m_ParentAddon.SboApplication.SetStatusBarMessage(ex.Message.ToString(), BoMessageTime.bmt_Medium, True)
            ProjectData.ClearProjectError()
        Finally
            Me.m_Form.Freeze(False)
        End Try
    End Sub

    Public Overrides Sub OnComboSelectAfter(sboObject As Object, pVal As SBOItemEventArg)
        MyBase.OnComboSelectAfter(RuntimeHelpers.GetObjectValue(sboObject), pVal)
        ' The following expression was wrapped in a checked-statement
        Try
            Dim itemUID As String = pVal.ItemUID
            Dim flag As Boolean = Operators.CompareString(itemUID, "btnCopy", False) = 0
            If flag Then
                Me.m_Form.Freeze(True)
                Dim selected As SAPbouiCOM.ValidValue = Me.btnCopy.Selected
                Me.m_Form.Freeze(True)
                flag = (Operators.CompareString(selected.Value, "2", False) = 0)
                If flag Then
                    Me.btnCopy.Caption = "Copy From"
                Else
                    Me.btnCopy.Caption = "Copy To"
                    Me.m_DataTable1.CopyFrom(Me.m_DataTable0)
                    Me.m_SboApplication.ActivateMenuItem("1282")
                    Dim arg_D2_0 As Integer = 0
                    Dim num As Integer = Me.m_DataTable1.Rows.Count - 1
                    Dim num2 As Integer = arg_D2_0
                    While True
                        Dim arg_F9_0 As Integer = num2
                        Dim num3 As Integer = num
                        If arg_F9_0 > num3 Then
                            Exit While
                        End If
                        Me.m_DataTable1.SetValue("NewRec", num2, -1)
                        num2 += 1
                    End While
                    Me.m_DataTable0.CopyFrom(Me.m_DataTable1)
                End If
                Me.FormatGrid()
                Me.m_Form.Freeze(False)
                Me.m_SboApplication.StatusBar.SetText("Data copied, Enter Class and Session information", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success)
            End If
        Catch expr_140 As Exception
            ProjectData.SetProjectError(expr_140)
            Dim ex As Exception = expr_140
            Me.m_SboApplication.StatusBar.SetText(ex.ToString(), BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
            Me.m_Form.Freeze(False)
            ProjectData.ClearProjectError()
        End Try
    End Sub

    Public Overrides Sub OnItemClickAfter(sboObject As Object, pVal As SBOItemEventArg)
        MyBase.OnItemClickAfter(RuntimeHelpers.GetObjectValue(sboObject), pVal)
        ' The following expression was wrapped in a checked-statement
        Try
            Dim itemUID As String = pVal.ItemUID
            Dim flag As Boolean = Operators.CompareString(itemUID, "btnRem", False) = 0
            If flag Then
                Dim selectedRows As SelectedRows = Me.grdStudCl.Rows.SelectedRows
                flag = (selectedRows.Count > 0)
                If flag Then
                    Dim arg_5D_0 As Integer = 0
                    Dim num As Integer = selectedRows.Count - 1
                    Dim num2 As Integer = arg_5D_0
                    While True
                        Dim arg_7E_0 As Integer = num2
                        Dim num3 As Integer = num
                        If arg_7E_0 > num3 Then
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
        Catch expr_B8 As Exception
            ProjectData.SetProjectError(expr_B8)
            ProjectData.ClearProjectError()
        End Try
    End Sub

    Public Overrides Sub OnItemClickBefore(sboObject As Object, pVal As SBOItemEventArg, ByRef BubbleEvent As Boolean)
        Dim flag As Boolean = Me.m_Form.Mode = BoFormMode.fm_ADD_MODE
        If flag Then
            Me.lAdd = True
        Else
            Me.lAdd = False
        End If
        MyBase.OnItemClickBefore(RuntimeHelpers.GetObjectValue(sboObject), pVal, BubbleEvent)
    End Sub

    Private Sub CommitDataTable()
        ' The following expression was wrapped in a checked-statement
        Try
            Dim userTables As UserTable = Me.m_ParentAddon.getUserTables("@OWA_EDUCLAHISTROWS")
            Me.m_Form.Freeze(True)
            Dim arg_35_0 As Integer = 0
            Dim num As Integer = Me.m_DataTable0.Rows.Count - 1
            Dim num2 As Integer = arg_35_0
            While True
                Dim arg_22B_0 As Integer = num2
                Dim num3 As Integer = num
                If arg_22B_0 > num3 Then
                    Exit While
                End If
                Dim flag As Boolean = Operators.ConditionalCompareObjectNotEqual(Me.m_DataTable0.GetValue("NewRec", num2), 1, False)
                If flag Then
                    Dim userTable As UserTable = userTables
                    Dim text As String = Conversions.ToString(Me.m_DataTable0.GetValue("U_CardCode", num2))
                    flag = Operators.ConditionalCompareObjectEqual(Me.m_DataTable0.GetValue("Code", num2), "xNew", False)
                    If flag Then
                        userTable.Code = Me.strDocEntry.ToString() + text
                        userTable.Name = Me.strDocEntry.ToString() + text
                    Else
                        userTable.Code = Conversions.ToString(Me.m_DataTable0.GetValue("Code", num2))
                        userTable.Name = Conversions.ToString(Me.m_DataTable0.GetValue("Code", num2))
                    End If
                    userTable.UserFields.Fields.Item("U_CardCode").Value = text
                    userTable.UserFields.Fields.Item("U_CardName").Value = RuntimeHelpers.GetObjectValue(Me.m_DataTable0.GetValue("U_CardName", num2))
                    userTable.UserFields.Fields.Item("U_DocEntry").Value = Me.strDocEntry
                    flag = (userTable.Add() <> 0)
                    If flag Then
                        Dim arg_1AB_0 As SAPbobsCOM.ICompany = Me.m_ParentAddon.SboCompany
                        Dim value2 As String
                        Dim value As Integer = Conversions.ToInteger(value2)
                        Dim text2 As String
                        arg_1AB_0.GetLastError(value, text2)
                        value2 = Conversions.ToString(value)
                        Me.m_ParentAddon.SboApplication.StatusBar.SetText(text2, BoMessageTime.bmt_Long, BoStatusBarMessageType.smt_Error)
                    Else
                        Me.m_DataTable0.SetValue("Code", num2, Me.strDocEntry.ToString() + text)
                        Me.m_DataTable0.SetValue("NewRec", num2, 1)
                    End If
                    Dim disposed As Boolean = Me.disposed
                End If
                num2 += 1
            End While
            Me.m_Form.Freeze(False)
            Me.m_Form.Update()
        Catch expr_24B As Exception
            ProjectData.SetProjectError(expr_24B)
            Dim ex As Exception = expr_24B
            Me.m_ParentAddon.SboApplication.StatusBar.SetText(ex.Message.ToString(), BoMessageTime.bmt_Long, BoStatusBarMessageType.smt_Error)
            Me.m_Form.Freeze(False)
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
        Dim result As Boolean
        Try
            Dim flag As Boolean = Operators.CompareString(Me.m_DBDataSource0.GetValue("U_Session", Me.m_DBDataSource0.Offset).Trim(), "", False) = 0
            If flag Then
                pErrMsg = "Please enter the session"
                Me.txtSessn.Active = True
                result = False
                Return result
            End If
            flag = (Operators.CompareString(Me.m_DBDataSource0.GetValue("U_KeyStage", Me.m_DBDataSource0.Offset).Trim(), "", False) = 0)
            If flag Then
                pErrMsg = "Please enter the Key stage"
                Me.txtKeyst.Active = True
                result = False
                Return result
            End If
            flag = (Operators.CompareString(Me.m_DBDataSource0.GetValue("U_Class", Me.m_DBDataSource0.Offset).Trim(), "", False) = 0)
            If flag Then
                pErrMsg = "Please enter the Class"
                Me.cboClass.Active = True
                result = False
                Return result
            End If
            Dim num As Integer = 0
            flag = Not Me.validateDataSource(Me.m_DataTable0, "U_CardCode", num)
            If flag Then
                pErrMsg = "Duplicate item codes exits!"
                result = False
                Return result
            End If
            flag = (Me.m_Form.Mode = BoFormMode.fm_ADD_MODE)
            If flag Then
                Dim text As String = "Select COUNT(*) nCnt from [@OWA_EDUCLAHISTORY] Where U_Session = 'x1' and U_KeyStage = 'x2' and U_Class = 'x3'"
                text = text.Replace("x1", Me.m_DBDataSource0.GetValue("U_Session", Me.m_DBDataSource0.Offset).Trim())
                text = text.Replace("x2", Me.m_DBDataSource0.GetValue("U_KeyStage", Me.m_DBDataSource0.Offset).Trim())
                text = text.Replace("x3", Me.m_DBDataSource0.GetValue("U_Class", Me.m_DBDataSource0.Offset).Trim())
                Dim dataTable As DataTable = Me.ExecuteSQLDT(text)
                flag = Operators.ConditionalCompareObjectNotEqual(dataTable.GetValue(0, 0), 0, False)
                If flag Then
                    pErrMsg = "This class already updated this session"
                    result = False
                    Return result
                End If
            End If
            Dim id As Integer = Me.m_id
        Catch expr_205 As Exception
            ProjectData.SetProjectError(expr_205)
            Dim ex As Exception = expr_205
            pErrMsg = ex.ToString()
            result = False
            ProjectData.ClearProjectError()
            Return result
        End Try
        result = True
        Return result
    End Function

    Public Overrides Sub OnDataAddAfter(ByRef pVal As BusinessObjectInfo)
        MyBase.OnDataAddAfter(pVal)
        Dim query As String = "Select Max(DocEntry) DocEntry from [@OWA_EDUCLAHISTORY]"
        Dim actionSuccess As Boolean = pVal.ActionSuccess
        If actionSuccess Then
            DataTable("DT_Base").Clear()
            DataTable("DT_Base").ExecuteQuery(query)
            Me.strDocEntry = Conversions.ToString(DataTable("DT_Base").GetValue(0, 0))
            Me.CommitDataTable()
            Dim str As String = " Where U_DocEntry = -1 "
            Me.m_Form.Freeze(True)
            Me.m_DataTable0.Clear()
            Me.m_DataTable0.ExecuteQuery(Me.strSQL + str)
            Me.grdStudCl.DataTable = Me.m_DataTable0
            Me.FormatGrid()
            Me.m_Form.Freeze(False)
        End If
    End Sub

    Public Overrides Sub OnDataUpdateAfter(ByRef pVal As BusinessObjectInfo)
        MyBase.OnDataUpdateAfter(pVal)
        Dim actionSuccess As Boolean = pVal.ActionSuccess
        If actionSuccess Then
            Me.CommitDataTable()
        End If
    End Sub

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

    Public Overrides Sub OnDataDeleteBefore(ByRef pVal As BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        MyBase.OnDataDeleteBefore(pVal, BubbleEvent)
        Me.strDocEntry = Me.m_DBDataSource0.GetValue("DocEntry", Me.m_DBDataSource0.Offset)
    End Sub

    Public Overrides Sub OnDataDeleteAfter(ByRef pVal As BusinessObjectInfo)
        MyBase.OnDataDeleteAfter(pVal)
        Try
            Dim actionSuccess As Boolean = pVal.ActionSuccess
            If actionSuccess Then
                Dim text As String = "Delete [@OWA_EDUCLAHISTROWS] Where U_DocEntry = " + Me.strDocEntry
                text = text.Replace("x1", Me.m_DBDataSource0.GetValue("U_Session", Me.m_DBDataSource0.Offset).Trim())
                text = text.Replace("x2", Me.m_DBDataSource0.GetValue("U_KeyStage", Me.m_DBDataSource0.Offset).Trim())
                text = text.Replace("x3", Me.m_DBDataSource0.GetValue("U_Class", Me.m_DBDataSource0.Offset).Trim())
                Dim dataTable As DataTable = Me.ExecuteSQLDT(text)
            End If
        Catch expr_B8 As Exception
            ProjectData.SetProjectError(expr_B8)
            Dim ex As Exception = expr_B8
            Me.m_SboApplication.StatusBar.SetText(ex.ToString(), BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
            ProjectData.ClearProjectError()
        End Try
    End Sub
End Class

