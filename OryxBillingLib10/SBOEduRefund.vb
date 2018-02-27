Imports Microsoft.VisualBasic.CompilerServices
Imports SAPbobsCOM
Imports SAPbouiCOM
Imports SBO.SboAddOnBase
Imports System
Imports System.Runtime.CompilerServices


Public Class SBOEduRefund
    Inherits SBOBaseObject

    Private txtDocEntry As EditText

    Private txtDate As EditText

    Private txtItem As EditText

    Private optInv As OptionBtn

    Private optCred As OptionBtn

    Private lblItemName As StaticText

    Private strSQL As String

    Private TableName As String

    Private strDocEntry As String

    Private grdStd As Grid

    Private editCol As GridColumn

    Private txtVal As EditText

    Private txtRem As EditText

    Private optItem As OptionBtn

    Private optService As OptionBtn

    Private lblAcct As StaticText

    Private lblCap As StaticText

    Private lnkItem As LinkedButton

    Private txtCAcct As EditText

    Private DocType As String

    Public Sub New(pAddOn As SboAddon, pForm As IForm)
        MyBase.New(pAddOn, pForm)
        Me.strSQL = "Select DocNum, CardCode, CardName, DocTotal, U_RefBatch from "
        Me.TableName = ""
        Me.DocType = "Item"
    End Sub

    Protected Overrides Sub AddDataSource()
        MyBase.AddDataSource()
        Me.m_DBDataSource0 = Me.m_Form.DataSources.DBDataSources.Item("@OWA_EDUSCHOOLREF")
        Me.m_DBDataSource1 = Me.m_Form.DataSources.DBDataSources.Item("@OWA_EDUSCHOOLREFDET")
        Me.m_DataTable0 = Me.m_Form.DataSources.DataTables.Item("DT_0")
    End Sub

    Public Overrides Sub OnCustomInit()
        MyBase.OnCustomInit()
    End Sub

    Protected Overrides Sub OnFormInit()
        MyBase.OnFormInit()
        Me.txtDocEntry = CType(Me.m_Form.Items.Item("txtDocEnt").Specific, EditText)
        Me.txtDate = CType(Me.m_Form.Items.Item("txtDate").Specific, EditText)
        Me.txtItem = CType(Me.m_Form.Items.Item("txtItem").Specific, EditText)
        Me.optCred = CType(Me.m_Form.Items.Item("optCred").Specific, OptionBtn)
        Me.optInv = CType(Me.m_Form.Items.Item("optInv").Specific, OptionBtn)
        Me.lblItemName = CType(Me.m_Form.Items.Item("lblItem").Specific, StaticText)
        Me.grdStd = CType(Me.m_Form.Items.Item("grdStd").Specific, Grid)
        Me.txtVal = CType(Me.m_Form.Items.Item("txtVal").Specific, EditText)
        Me.txtRem = CType(Me.m_Form.Items.Item("txtRem").Specific, EditText)
        Me.lnkItem = CType(Me.m_Form.Items.Item("lnkItem").Specific, LinkedButton)
        Me.txtCAcct = CType(Me.m_Form.Items.Item("txtCAcct").Specific, EditText)
        Me.optItem = CType(Me.m_Form.Items.Item("optItem").Specific, OptionBtn)
        Me.optService = CType(Me.m_Form.Items.Item("optService").Specific, OptionBtn)
        Me.lblAcct = CType(Me.m_Form.Items.Item("lblAcct").Specific, StaticText)
        Me.lblCap = CType(Me.m_Form.Items.Item("lblCap").Specific, StaticText)
        Me.optCred.GroupWith("optInv")
        Me.optService.GroupWith("optItem")
    End Sub

    Protected Overrides Sub OnFormNavigate()
        MyBase.OnFormNavigate()
        Try
            Dim flag As Boolean = Me.m_Form.Mode <> BoFormMode.fm_ADD_MODE
            If flag Then
                Me.strDocEntry = Me.m_DBDataSource0.GetValue("DocEntry", Me.m_DBDataSource0.Offset)
            Else
                Me.strDocEntry = "-1"
            End If
            flag = (Operators.CompareString(Me.TableName, "", False) = 0)
            Dim str As String
            If flag Then
                str = " ORIN  Where U_RefBatch = "
            Else
                str = Me.TableName + " Where U_RefBatch = "
            End If
            Me.m_Form.Freeze(True)
            Me.m_DataTable0.Clear()
            Me.m_DataTable0.ExecuteQuery(Me.strSQL + str + Me.strDocEntry)
            Me.grdStd.DataTable = Me.m_DataTable0
            Me.FormatGrid()
            Me.m_Form.Freeze(False)
        Catch expr_E0 As Exception
            ProjectData.SetProjectError(expr_E0)
            Dim ex As Exception = expr_E0
            Me.m_SboApplication.StatusBar.SetText(ex.ToString(), BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
            Me.m_Form.Freeze(False)
            ProjectData.ClearProjectError()
        End Try
    End Sub

    Protected Overrides Sub FormatGrid()
        MyBase.FormatGrid()
        ' The following expression was wrapped in a checked-statement
        Try
            Me.grdStd.Columns.Item("CardCode").TitleObject.Caption = "Student Code"
            Me.grdStd.Columns.Item("CardName").TitleObject.Caption = "Student Name"
            Me.grdStd.Columns.Item("DocNum").TitleObject.Caption = "Doc No."
            Me.grdStd.Columns.Item("DocTotal").TitleObject.Caption = "Amount"
            Me.grdStd.Columns.Item("U_RefBatch").Visible = False
            Dim arg_CD_0 As Integer = 0
            Dim num As Integer = Me.grdStd.Columns.Count - 1
            Dim num2 As Integer = arg_CD_0
            While True
                Dim arg_F8_0 As Integer = num2
                Dim num3 As Integer = num
                If arg_F8_0 > num3 Then
                    Exit While
                End If
                Me.grdStd.Columns.Item(num2).Editable = False
                num2 += 1
            End While
            Dim flag As Boolean = Me.m_Form.Mode = BoFormMode.fm_ADD_MODE
            If flag Then
                Me.grdStd.Columns.Item("DocTotal").Editable = True
            Else
                Me.grdStd.Columns.Item("DocTotal").Editable = False
            End If
            Me.editCol = Me.grdStd.Columns.Item("DocNum")
            flag = (Operators.CompareString(Me.TableName, "OINV", False) = 0)
            If flag Then
                NewLateBinding.LateSet(Me.editCol, Nothing, "LinkedObjectType", New Object() {13}, Nothing, Nothing)
            Else
                NewLateBinding.LateSet(Me.editCol, Nothing, "LinkedObjectType", New Object() {14}, Nothing, Nothing)
            End If
            Me.editCol = Me.grdStd.Columns.Item("CardCode")
            NewLateBinding.LateSet(Me.editCol, Nothing, "LinkedObjectType", New Object() {2}, Nothing, Nothing)
            Me.grdStd.AutoResizeColumns()
        Catch expr_220 As Exception
            ProjectData.SetProjectError(expr_220)
            Dim ex As Exception = expr_220
            Me.m_ParentAddon.SboApplication.SetStatusBarMessage(ex.Message.ToString(), BoMessageTime.bmt_Medium, True)
            ProjectData.ClearProjectError()
        Finally
            Me.m_Form.Freeze(False)
        End Try
    End Sub

    Public Overrides Sub OnChooseFromListAfter(sboObject As Object, pVal As SBOItemEventArg)
        MyBase.OnChooseFromListAfter(RuntimeHelpers.GetObjectValue(sboObject), pVal)
        ' The following expression was wrapped in a checked-statement
        Try
            Dim dataTable As DataTable = Nothing
            Dim itemUID As String = pVal.ItemUID
            Dim flag As Boolean = Operators.CompareString(itemUID, "txtItem", False) = 0 OrElse Operators.CompareString(itemUID, "txtCAcct", False) = 0
            Dim text As String
            If flag Then
                Dim arg_55_1 As String = pVal.FormUID
                Dim flag2 As Boolean = False
                text = Me.HandleChooseFromListEvent(arg_55_1, pVal, flag2)
                flag = String.IsNullOrEmpty(text)
                If flag Then
                    Return
                End If
            Else
                Dim arg_80_1 As String = pVal.FormUID
                Dim flag2 As Boolean = False
                Dim flag3 As Boolean
                dataTable = Me.HandleChooseFromListEvent(arg_80_1, pVal, flag2, flag3)
                flag = (dataTable Is Nothing)
                If flag Then
                    Return
                End If
            End If
            Dim itemUID2 As String = pVal.ItemUID
            flag = (Operators.CompareString(itemUID2, "txtItem", False) = 0)
            If flag Then
                Me.m_DBDataSource0.SetValue("U_itemcode", Me.m_DBDataSource0.Offset, text)
                flag = (Operators.CompareString(Me.DocType, "Items", False) = 0)
                If flag Then
                    Me.getOffset(text, "ItemCode", DBDS("OITM"))
                    Me.lblItemName.Caption = DBDS("OITM").GetValue("ItemName", 0).Trim()
                Else
                    Me.getOffset(text, "AcctCode", DBDS("OACT"))
                    Me.lblItemName.Caption = DBDS("OACT").GetValue("AcctName", 0).Trim()
                End If
            Else
                flag = (Operators.CompareString(itemUID2, "txtCAcct", False) = 0)
                If flag Then
                    Me.m_DBDataSource0.SetValue("U_CAcct", Me.m_DBDataSource0.Offset, text)
                    Me.getOffset(text, "AcctCode", DBDS("OACT"))
                    Me.lblAcct.Caption = DBDS("OACT").GetValue("AcctName", 0).Trim()
                Else
                    flag = (Operators.CompareString(itemUID2, "btnAdd", False) = 0)
                    If flag Then
                        Me.m_Form.Freeze(True)
                        flag = (Me.m_Form.Mode = BoFormMode.fm_ADD_MODE)
                        If flag Then
                            Dim num As Integer = Me.m_DataTable0.Rows.Count
                            Dim num2 As Double = 0.0
                            flag = (Operators.CompareString(Me.txtVal.Value, "", False) <> 0)
                            If flag Then
                                num2 = Conversions.ToDouble(Me.txtVal.Value)
                            End If
                            Dim arg_290_0 As Integer = 0
                            Dim num3 As Integer = dataTable.Rows.Count - 1
                            Dim num4 As Integer = arg_290_0
                            While True
                                Dim arg_351_0 As Integer = num4
                                Dim num5 As Integer = num3
                                If arg_351_0 > num5 Then
                                    Exit While
                                End If
                                Me.m_DataTable0.Rows.Add(1)
                                Dim text2 As String = Conversions.ToString(dataTable.GetValue("CardCode", num4))
                                Me.m_DataTable0.SetValue("CardCode", num, text2)
                                flag = Me.getOffset(text2, "CardCode", DBDS("OCRD"))
                                If flag Then
                                    Dim value As String = DBDS("OCRD").GetValue("CardName", 0)
                                    Me.m_DataTable0.SetValue("CardName", num, value.Trim())
                                    Me.m_DataTable0.SetValue("DocTotal", num, num2)
                                End If
                                num += 1
                                num4 += 1
                            End While
                            flag = Operators.ConditionalCompareObjectEqual(Me.m_DataTable0.GetValue(1, 0), "", False)
                            If flag Then
                                Me.m_DataTable0.Rows.Remove(0)
                            End If
                        End If
                    End If
                End If
            End If
        Catch expr_390 As Exception
            ProjectData.SetProjectError(expr_390)
            Dim ex As Exception = expr_390
            Me.m_SboApplication.StatusBar.SetText(ex.Message.ToString(), BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Error)
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
            Dim flag As Boolean = Operators.CompareString(itemUID, "optInv", False) = 0
            If flag Then
                Me.TableName = "OINV"
            Else
                flag = (Operators.CompareString(itemUID, "optCred", False) = 0)
                If flag Then
                    Me.TableName = "ORIN"
                Else
                    flag = (Operators.CompareString(itemUID, "optService", False) = 0)
                    If flag Then
                        Me.txtItem.ChooseFromListUID = "CFLOACT"
                        Me.txtItem.ChooseFromListAlias = "AcctCode"
                        Me.lnkItem.LinkedObject = BoLinkedObject.lf_GLAccounts
                        Me.DocType = "Service"
                        Me.lblCap.Caption = "Account"
                    Else
                        flag = (Operators.CompareString(itemUID, "optItem", False) = 0)
                        If flag Then
                            Me.txtItem.ChooseFromListUID = "cflItem"
                            Me.txtItem.ChooseFromListAlias = "ItemCode"
                            Me.lnkItem.LinkedObject = BoLinkedObject.lf_Items
                            Me.DocType = "Items"
                            Me.lblCap.Caption = "Item"
                        Else
                            flag = (Operators.CompareString(itemUID, "btnUpdate", False) = 0)
                            If flag Then
                                Me.m_Form.Freeze(True)
                                Dim num As Double = 0.0
                                flag = (Operators.CompareString(Me.txtVal.Value, "", False) <> 0)
                                If flag Then
                                    num = Conversions.ToDouble(Me.txtVal.Value)
                                End If
                                Dim arg_1AD_0 As Integer = 0
                                Dim num2 As Integer = Me.m_DataTable0.Rows.Count - 1
                                Dim num3 As Integer = arg_1AD_0
                                While True
                                    Dim arg_1D4_0 As Integer = num3
                                    Dim num4 As Integer = num2
                                    If arg_1D4_0 > num4 Then
                                        Exit While
                                    End If
                                    Me.m_DataTable0.SetValue("DocTotal", num3, num)
                                    num3 += 1
                                End While
                            Else
                                flag = (Operators.CompareString(itemUID, "btnClear", False) = 0)
                                If flag Then
                                    Me.m_DataTable0.Clear()
                                    Me.OnFormNavigate()
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Catch expr_205 As Exception
            ProjectData.SetProjectError(expr_205)
            Dim ex As Exception = expr_205
            Me.m_SboApplication.StatusBar.SetText(ex.Message.ToString(), BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Error)
            ProjectData.ClearProjectError()
        Finally
            Me.m_Form.Freeze(False)
        End Try
    End Sub

    Public Overrides Sub OnDataAddAfter(ByRef pVal As BusinessObjectInfo)
        MyBase.OnDataAddAfter(pVal)
        Me.OINVCreate()
    End Sub

    Public Overrides Sub OnDataAddBefore(ByRef pVal As BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        MyBase.OnDataAddBefore(pVal, BubbleEvent)
        Dim text As String = ""
        Dim dateTime As DateTime = Me.sboDate(Me.txtDate.Value)
        Dim value As String = Me.txtItem.Value
        Dim value2 As String = Me.txtRem.Value
        Dim value3 As String = Me.txtCAcct.Value
        ' The following expression was wrapped in a checked-statement
        Dim num As Integer = Me.m_DataTable0.Rows.Count - 1
        Dim arg_62_0 As Integer = 0
        Dim num2 As Integer = num
        Dim num3 As Integer = arg_62_0
        Dim flag As Boolean
        While True
            Dim arg_F9_0 As Integer = num3
            Dim num4 As Integer = num2
            If arg_F9_0 > num4 Then
                Exit While
            End If
            Dim left As String = Conversions.ToString(Me.m_DataTable0.GetValue("CardCode", num3))
            Dim num5 As Double = Conversions.ToDouble(Me.m_DataTable0.GetValue("DocTotal", num3))
            flag = (Operators.CompareString(left, "", False) = 0)
            If flag Then
                text = "Please enter enter student information"
            End If
            flag = (num5 = 0.0)
            If flag Then
                text = "Amount cannot be blank"
            End If
            flag = (Operators.CompareString(text, "", False) <= 0)
            If flag Then
                Exit While
            End If
            num3 += 1
        End While
        flag = (Operators.CompareString(value, "", False) = 0)
        If flag Then
            text = "The item to be posted cannot be blank"
        End If
        flag = (Operators.CompareString(Me.TableName, "", False) = 0)
        If flag Then
            text = "Select Trans. Type"
        End If
        flag = (Operators.CompareString(value3, "", False) = 0)
        If flag Then
            text = "Select a control account"
        End If
        flag = (Operators.CompareString(Me.m_DBDataSource0.GetValue("U_DocType", Me.m_DBDataSource0.Offset), "", False) = 0)
        If flag Then
            text = "Select a document type"
        End If
        Dim num6 As Integer
        flag = Not Me.validateDataSource(Me.m_DataTable0, "CardCode", num6)
        If flag Then
            text = "Duplicate entries found"
        End If
        flag = (Operators.CompareString(text, "", False) <> 0)
        If flag Then
            Me.m_SboApplication.StatusBar.SetText(text, BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Error)
            BubbleEvent = False
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
        Me.m_Form.ChooseFromLists.Item("CFLCard").SetConditions(Me.m_Conditions)
        Me.m_Form.ChooseFromLists.Item("cflItem").SetConditions(Me.m_Conditions)
        Dim condition3 As Condition = Me.m_Condition1
        condition3.[Alias] = "ActType"
        condition3.Operation = BoConditionOperation.co_EQUAL
        condition3.CondVal = "I"
        condition3.Relationship = BoConditionRelationship.cr_AND
        Me.m_Condition1 = Me.m_Conditions1.Add()
        Dim condition4 As Condition = Me.m_Condition1
        condition4.[Alias] = "FrozenFor"
        condition4.Operation = BoConditionOperation.co_NOT_EQUAL
        condition4.CondVal = "Y"
        Me.m_Form.ChooseFromLists.Item("CFLOACT").SetConditions(Me.m_Conditions1)
        Dim condition5 As Condition = Me.m_Condition2
        condition5.[Alias] = "LocManTran"
        condition5.Operation = BoConditionOperation.co_EQUAL
        condition5.CondVal = "Y"
        condition5.Relationship = BoConditionRelationship.cr_AND
        Me.m_Condition2 = Me.m_Conditions2.Add()
        Dim condition6 As Condition = Me.m_Condition2
        condition6.[Alias] = "FrozenFor"
        condition6.Operation = BoConditionOperation.co_NOT_EQUAL
        condition6.CondVal = "Y"
        Me.m_Form.ChooseFromLists.Item("CFLCACT").SetConditions(Me.m_Conditions2)
    End Sub

    Private Function OINVCreate() As Boolean
        Dim flag As Boolean = True
        ' The following expression was wrapped in a checked-statement
        Dim result As Boolean
        Try
            Me.m_ParentAddon.SboCompany.StartTransaction()
            Dim flag2 As Boolean = Operators.CompareString(Me.TableName, "OINV", False) = 0
            Dim documents As Documents
            If flag2 Then
                documents = CType(Me.m_ParentAddon.SboCompany.GetBusinessObject(BoObjectTypes.oInvoices), Documents)
            Else
                documents = CType(Me.m_ParentAddon.SboCompany.GetBusinessObject(BoObjectTypes.oCreditNotes), Documents)
            End If
            Dim text As String = "Select Max(DocEntry) from [@OWA_EDUSCHOOLREF]"
            Dim dataTable As DataTable = Me.m_Form.DataSources.DataTables.Item("DT_0")
            dataTable = Me.ExecuteSQLDT(text)
            flag2 = Not dataTable.IsEmpty
            If flag2 Then
                Dim num As Integer = Conversions.ToInteger(dataTable.GetValue(0, 0))
                Dim num2 As Integer = Me.m_DataTable0.Rows.Count - 1
                Dim arg_DB_0 As Integer = 0
                Dim num3 As Integer = num2
                Dim num4 As Integer = arg_DB_0
                Dim text3 As String
                While True
                    Dim arg_397_0 As Integer = num4
                    Dim num5 As Integer = num3
                    If arg_397_0 > num5 Then
                        GoTo IL_39C
                    End If
                    Dim text2 As String = Conversions.ToString(Me.m_DataTable0.GetValue("CardCode", num4))
                    Dim num6 As Double = Conversions.ToDouble(Me.m_DataTable0.GetValue("DocTotal", num4))
                    Dim dateTime As DateTime = Me.sboDate(Me.txtDate.Value)
                    Dim value As String = Me.txtItem.Value
                    Dim value2 As String = Me.txtRem.Value
                    Dim value3 As String = Me.txtCAcct.Value
                    flag2 = (Operators.CompareString(text2, "", False) = 0)
                    If flag2 Then
                        text3 = "Please enter enter student information"
                    End If
                    flag2 = (num6 = 0.0)
                    If flag2 Then
                        text3 = "Amount cannot be blank"
                    End If
                    flag2 = (Operators.CompareString(value, "", False) = 0)
                    If flag2 Then
                        text3 = "The item to be posted cannot be blank"
                    End If
                    flag2 = (Operators.CompareString(Me.m_DBDataSource0.GetValue("U_DocType", Me.m_DBDataSource0.Offset), "", False) = 0)
                    If flag2 Then
                        text3 = "Select a document type"
                    End If
                    flag2 = (Operators.CompareString(text3, "", False) <> 0)
                    If flag2 Then
                        Exit While
                    End If
                    documents.Series = 0
                    documents.CardCode = text2
                    documents.NumAtCard = num.ToString()
                    documents.HandWritten = BoYesNoEnum.tNO
                    documents.PaymentGroupCode = Conversions.ToInteger("-1")
                    documents.DocDate = dateTime
                    documents.DocDueDate = dateTime
                    documents.Comments = value2.Trim()
                    documents.UserFields.Fields.Item("U_RefBatch").Value = num
                    documents.ControlAccount = value3.Trim()
                    flag2 = (Operators.CompareString(Me.DocType, "Items", False) = 0)
                    If flag2 Then
                        documents.DocType = BoDocumentTypes.dDocument_Items
                        documents.Lines.ItemCode = value
                        documents.Lines.UnitPrice = num6
                        documents.Lines.Quantity = 1.0
                        documents.Lines.LineTotal = num6
                    Else
                        documents.DocType = BoDocumentTypes.dDocument_Service
                        documents.Lines.AccountCode = value
                        documents.Lines.LineTotal = num6
                        documents.Lines.ItemDescription = value2.Trim()
                    End If
                    flag2 = (documents.Add() <> 0)
                    If flag2 Then
                        Dim num7 As Integer
                        Me.m_ParentAddon.SboCompany.GetLastError(num7, text3)
                        text3 = text3 + " BP - " + text2
                        Me.m_SboApplication.StatusBar.SetText(text3, BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Error)
                        flag = False
                    End If
                    flag2 = Not flag
                    If flag2 Then
                        GoTo Block_13
                    End If
                    num4 += 1
                End While
                flag = False
Block_13:
IL_39C:
                flag2 = (flag And Me.m_ParentAddon.SboCompany.InTransaction)
                If flag2 Then
                    Me.m_ParentAddon.SboCompany.EndTransaction(BoWfTransOpt.wf_Commit)
                    Me.m_SboApplication.StatusBar.SetText("Transactions posted", BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Success)
                Else
                    Me.m_ParentAddon.SboCompany.EndTransaction(BoWfTransOpt.wf_RollBack)
                    Me.m_SboApplication.StatusBar.SetText(text3, BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Error)
                End If
                result = flag
            Else
                result = False
            End If
        Catch expr_417 As Exception
            ProjectData.SetProjectError(expr_417)
            Dim ex As Exception = expr_417
            Me.m_SboApplication.StatusBar.SetText(ex.Message.ToString(), BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Error)
            Dim flag2 As Boolean = Me.m_ParentAddon.SboCompany.InTransaction
            If flag2 Then
                Me.m_ParentAddon.SboCompany.EndTransaction(BoWfTransOpt.wf_RollBack)
            End If
            result = False
            ProjectData.ClearProjectError()
        Finally
            Dim flag2 As Boolean = Me.m_ParentAddon.SboCompany.InTransaction
            If flag2 Then
                Me.m_ParentAddon.SboCompany.EndTransaction(BoWfTransOpt.wf_RollBack)
            End If
        End Try
        Return result
    End Function
End Class

