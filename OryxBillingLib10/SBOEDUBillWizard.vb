Imports Microsoft.VisualBasic
Imports Microsoft.VisualBasic.CompilerServices
Imports OWA.SBO.OryxSQL
Imports SAPbobsCOM
Imports SAPbouiCOM
Imports SBO.SboAddOnBase
Imports System
Imports System.Collections
Imports System.Runtime.CompilerServices
Imports OryxBillingLib10.OWA.SBO.OryxBillingLib10


Public Class SBOEDUBillWizard
    Inherits SBOWizard

    Private Structure InvHeader
        Public CardCode As String

        Public CardName As String

        Public KeyStage As String

        Public strClass As String

        Public termType As Integer

        Public TermDate As DateTime

        Public strRemarks As String

        Public Term As String

        Public School As String

        Public Dim1 As String

        Public Dim2 As String
    End Structure

    Private Structure DocH
        Public Session As String

        Public BatchDate As DateTime
    End Structure

    Private lblTitle1 As StaticText

    Private lblDesc1 As StaticText

    Private optNew As OptionBtn

    Private optUpdate As OptionBtn

    Private txtDocNum As EditText

    Private btnSelect As Button

    Private btnDelete As Button

    Private lblBatch As StaticText

    Private lblDesc2 As StaticText

    Private lblDesc2_1 As StaticText

    Private lblDesc2_2 As StaticText

    Private lblTitle2 As StaticText

    Private txtSession As EditText

    Private txtDesc As EditText

    Private lblSession As StaticText

    Private lblBDesc As StaticText

    Private btnSaveStd As Button

    Private cboJump As ButtonCombo

    Private cboClass As ComboBox

    Private grdStud As Grid

    Private btnColAll As Button

    Private btnExpAll As Button

    Private btnStd As Button

    Private txtCount As EditText

    Private grdSpecial As Grid

    Private oComboCol As GridColumn

    Private grdSibling As Grid

    Private lSiblingLoaded As Boolean

    Private btnReview As ButtonCombo

    Private lOneTimeLoaded As Boolean

    Private lStaffKLoaded As Boolean

    Private btnEmail As Button

    Private lEmailLoaded As Boolean

    Private grdEmail As Grid

    Private grdReview As Grid

    Private lReviewLoaded As Boolean

    Private btnPost, btnCancelInv As Button

    Private grdInvoice As Grid

    Private lInvoiceLoaded As Boolean

    Private cboPost As ButtonCombo

    Private DocEntry As Integer

    Private aTitlelist As ArrayList

    Private aDescList As ArrayList

    Private aDescList_1 As ArrayList

    Private aDescList_2 As ArrayList

    Private editCol As GridColumn

    Private lAdd As Boolean

    Private oProgBar As ProgressBar

    Private UserDB As UserDataSource

    Private Invoiced As String

    Protected Overrides Sub OnFormInit()
        MyBase.OnFormInit()
        Me.LastPane = 9
        Me.lblTitle1 = CType(Me.m_Form.Items.Item("lblTitle_1").Specific, StaticText)
        Me.lblDesc1 = CType(Me.m_Form.Items.Item("lblDesc_1").Specific, StaticText)
        Me.optNew = CType(Me.m_Form.Items.Item("optNew").Specific, OptionBtn)
        Me.optUpdate = CType(Me.m_Form.Items.Item("optUpdate").Specific, OptionBtn)
        Me.optUpdate.GroupWith("optNew")
        Me.txtDocNum = CType(Me.m_Form.Items.Item("txtDocNum").Specific, EditText)
        Me.btnSelect = CType(Me.m_Form.Items.Item("btnSelect").Specific, Button)
        Me.btnDelete = CType(Me.m_Form.Items.Item("btnDelete").Specific, Button)
        Me.lblTitle2 = CType(Me.m_Form.Items.Item("lblTitle_2").Specific, StaticText)
        Me.lblDesc2 = CType(Me.m_Form.Items.Item("lblDesc_2").Specific, StaticText)
        Me.lblDesc2_1 = CType(Me.m_Form.Items.Item("lblD_2_1").Specific, StaticText)
        Me.lblDesc2_2 = CType(Me.m_Form.Items.Item("lblD_2_2").Specific, StaticText)
        Me.txtSession = CType(Me.m_Form.Items.Item("txtSession").Specific, EditText)
        Me.txtDesc = CType(Me.m_Form.Items.Item("txtDesc").Specific, EditText)
        Me.lblBatch = CType(Me.m_Form.Items.Item("lblBatch").Specific, StaticText)
        Me.lblBDesc = CType(Me.m_Form.Items.Item("lblBDesc").Specific, StaticText)
        Me.lblSession = CType(Me.m_Form.Items.Item("lblSession").Specific, StaticText)
        Me.btnSaveStd = CType(Me.m_Form.Items.Item("btnSaveStd").Specific, Button)
        Me.cboJump = CType(Me.m_Form.Items.Item("cboJump").Specific, ButtonCombo)
        Me.cboJump.Item.AffectsFormMode = False
        Me.cboJump.ExpandType = BoExpandType.et_ValueDescription
        Me.cboJump.ValidValues.Add("2", "Select Batch")
        Me.cboJump.ValidValues.Add("3", "Special Items")
        Me.cboJump.ValidValues.Add("4", "Student List")
        Me.cboJump.ValidValues.Add("5", "Review Discount")
        Me.cboJump.ValidValues.Add("6", "Review Batch")
        Me.cboJump.ValidValues.Add("7", "Send Emails")
        Me.cboJump.ValidValues.Add("8", "Post Invoices")
        Me.optNew.Item.AffectsFormMode = False
        Me.optUpdate.Item.AffectsFormMode = False
        Me.lblTitle1.Item.TextStyle = 1
        Me.lblTitle1.Item.FontSize = 12
        Me.lblTitle2.Item.TextStyle = 1
        Me.lblTitle2.Item.FontSize = 12
        Me.grdStud = CType(Me.m_Form.Items.Item("grdStud").Specific, Grid)
        Me.btnExpAll = CType(Me.m_Form.Items.Item("btnExpAll").Specific, Button)
        Me.btnColAll = CType(Me.m_Form.Items.Item("btnColAll").Specific, Button)
        Me.btnStd = CType(Me.m_Form.Items.Item("btnStd").Specific, Button)
        Me.grdReview = CType(Me.m_Form.Items.Item("grdReview").Specific, Grid)
        Me.btnEmail = CType(Me.m_Form.Items.Item("btnEmail").Specific, Button)
        Me.grdEmail = CType(Me.m_Form.Items.Item("grdEmail").Specific, Grid)
        Me.btnPost = CType(Me.m_Form.Items.Item("btnPost").Specific, Button)
        Me.btnCancelInv = CType(Me.m_Form.Items.Item("btnCancInv").Specific, Button)
        Me.grdInvoice = CType(Me.m_Form.Items.Item("grdInvoice").Specific, Grid)
        Me.fillTextArray()
        Me.m_DataTable0 = Me.m_Form.DataSources.DataTables.Item("DT_Students")
        Me.m_DataTable1 = Me.m_Form.DataSources.DataTables.Item("DT_Special")
        Me.m_DataTable2 = Me.m_Form.DataSources.DataTables.Item("DT_OneTime")
        Me.m_DataTable3 = Me.m_Form.DataSources.DataTables.Item("DT_Summary")
        Me.m_DataTable4 = Me.m_Form.DataSources.DataTables.Item("DT_Invoice")
        Me.UserDB = Me.m_Form.DataSources.UserDataSources.Item("UD_Count")
    End Sub

    Protected Overrides Sub EnableToolBarButtons()
        MyBase.EnableToolBarButtons()
        Me.m_Form.EnableMenu("1292", True)
    End Sub

    Public Sub New(pAddon As SboAddon, pForm As IForm)
        MyBase.New(pAddon, pForm)
        Me.aTitlelist = New ArrayList(8)
        Me.aDescList = New ArrayList(8)
        Me.aDescList_1 = New ArrayList(8)
        Me.aDescList_2 = New ArrayList(8)
        Me.InitSBOServerSQL(New BusObjectInfoSQL(pAddon))
    End Sub

    Protected Overrides Sub AddDataSource()
        MyBase.AddDataSource()
        Me.m_DBDataSource1 = Me.m_Form.DataSources.DBDataSources.Item("@OWA_EDUSESSIONS")
        Me.m_DBDataSource0 = Me.m_Form.DataSources.DBDataSources.Item("@OWA_EDUBILLINGS")
        Me.m_DBDataSource2 = Me.m_Form.DataSources.DBDataSources.Item("@OWA_EDUBILLSPECIAL")
        Me.m_DBDataSource3 = Me.m_Form.DataSources.DBDataSources.Item("@OWA_EDUSESSIONS")
        Me.m_DataTable0 = Me.m_Form.DataSources.DataTables.Item("DT_Students")
        Me.m_DataTable1 = Me.m_Form.DataSources.DataTables.Item("DT_Sibling")
        Me.m_DataTable2 = Me.m_Form.DataSources.DataTables.Item("DT_OneTime")
    End Sub

    Public Overrides Sub OnCustomInit()
        MyBase.OnCustomInit()
    End Sub

    Public Overrides Sub OnChooseFromListAfter(sboObject As Object, pVal As SBOItemEventArg)
        MyBase.OnChooseFromListAfter(RuntimeHelpers.GetObjectValue(sboObject), pVal)
        Try
            Dim arg_20_1 As String = pVal.FormUID
            Dim flag As Boolean = False
            Dim text As String = Me.HandleChooseFromListEvent(arg_20_1, pVal, flag)
            Dim flag2 As Boolean = String.IsNullOrEmpty(text)
            If Not flag2 Then
                Dim itemUID As String = pVal.ItemUID
                flag2 = (Operators.CompareString(itemUID, "txtSession", False) = 0 OrElse Operators.CompareString(itemUID, "matSpecial", False) = 0)
                If flag2 Then
                    Dim arg_7B_1 As String = pVal.FormUID
                    flag = False
                    text = Me.HandleChooseFromListEvent(arg_7B_1, pVal, flag)
                    flag2 = String.IsNullOrEmpty(text)
                    If flag2 Then
                        Return
                    End If
                Else
                    Dim arg_A6_1 As String = pVal.FormUID
                    flag = False
                    Dim flag3 As Boolean
                    Dim dataTable As DataTable = Me.HandleChooseFromListEvent(arg_A6_1, pVal, flag, flag3)
                    flag2 = (dataTable Is Nothing)
                    If flag2 Then
                        Return
                    End If
                End If
                Dim itemUID2 As String = pVal.ItemUID
                flag2 = (Operators.CompareString(itemUID2, "txtSession", False) = 0)
                If flag2 Then
                    Me.m_DBDataSource0.SetValue("U_Session", Me.m_DBDataSource0.Offset, text)
                End If
            End If
        Catch expr_FB As Exception
            ProjectData.SetProjectError(expr_FB)
            Dim ex As Exception = expr_FB
            Me.m_ParentAddon.SboApplication.SetStatusBarMessage(ex.ToString(), BoMessageTime.bmt_Short, True)
            ProjectData.ClearProjectError()
        End Try
    End Sub

    Public Overrides Sub OnItemClickAfter(sboObject As Object, pVal As SBOItemEventArg)
        MyBase.OnItemClickAfter(RuntimeHelpers.GetObjectValue(sboObject), pVal)
        Try
            Dim itemUID As String = pVal.ItemUID
            Dim flag As Boolean = Operators.CompareString(itemUID, "optNew", False) = 0
            If flag Then
                Me.m_Form.Mode = BoFormMode.fm_ADD_MODE
                Me.RefreshPane2Controls(1)
                Me.lAdd = True
            Else
                flag = (Operators.CompareString(itemUID, "optUpdate", False) = 0)
                If flag Then
                    Me.m_Form.Mode = BoFormMode.fm_OK_MODE
                    Me.RefreshPane2Controls(2)
                    Me.lAdd = False
                Else
                    flag = (Operators.CompareString(itemUID, "btnStd", False) = 0)
                    If flag Then
                        Select Case Me.CurrentPane
                            Case 3
                                'Me.RefreshStudentList(True)
                            Case 4
                                Me.RefreshSummary(True)
                            Case 5
                                Me.refreshEmailList()
                        End Select
                    Else
                        flag = (Operators.CompareString(itemUID, "btnExpAll", False) = 0)
                        If flag Then
                            Me.grdStud.Rows.ExpandAll()
                        Else
                            flag = (Operators.CompareString(itemUID, "btnColAll", False) = 0)
                            If flag Then
                                Me.grdStud.Rows.CollapseAll()
                            Else
                                flag = (Operators.CompareString(itemUID, "btnStd", False) = 0)
                                If Not flag Then
                                    flag = (Operators.CompareString(itemUID, "btnOneTime", False) = 0)
                                    If Not flag Then
                                        flag = (Operators.CompareString(itemUID, "btnSaveStd", False) = 0)
                                        If flag Then
                                            Select Case Me.CurrentPane
                                                Case 2
                                                    Me.SaveBatchHeader()
                                                Case 3
                                                    Me.SaveStudentList2()
                                            End Select
                                        Else
                                            flag = (Operators.CompareString(itemUID, "btnDelete", False) = 0)
                                            If flag Then
                                                Me.DeleteBatch()
                                            Else
                                                flag = (Operators.CompareString(itemUID, "btnEmail", False) = 0)
                                                If Not flag Then
                                                    flag = (Operators.CompareString(itemUID, "btnPost", False) = 0)
                                                    If flag Then
                                                        Me.PostInvoices(-1)
                                                    Else
                                                        flag = (Operators.CompareString(itemUID, "btnCancInv", False) = 0)
                                                        If flag Then
                                                            Me.CancelInvoices()
                                                        Else
                                                            flag = (Operators.CompareString(itemUID, "grdSpecial", False) = 0)
                                                            If flag Then
                                                                Dim flag2 As Boolean = pVal.Row = -1
                                                                If flag2 Then
                                                                    Me.m_DataTable1.Rows.Add(1)
                                                                    Dim comboBoxColumn As ComboBoxColumn = CType(Me.grdSpecial.Columns.Item(0), ComboBoxColumn)
                                                                    ' The following expression was wrapped in a checked-expression
                                                                    comboBoxColumn.Click(Me.m_DataTable1.Rows.Count - 1, True, 0)
                                                                End If
                                                            End If
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Catch expr_296 As Exception
            ProjectData.SetProjectError(expr_296)
            ProjectData.ClearProjectError()
        End Try
    End Sub

    Public Overrides Sub OnItemClickBefore(sboObject As Object, pVal As SBOItemEventArg, ByRef BubbleEvent As Boolean)
        MyBase.OnItemClickBefore(RuntimeHelpers.GetObjectValue(sboObject), pVal, BubbleEvent)
        Dim itemUID As String = pVal.ItemUID
        Dim flag As Boolean = Operators.CompareString(itemUID, "grdSpecial", False) = 0
        If flag Then
            ' The following expression was wrapped in a checked-expression
            Dim flag2 As Boolean = pVal.Row = Me.m_DataTable1.Rows.Count - 1
            If flag2 Then
                Me.m_DataTable1.Rows.Add(1)
            End If
        End If
    End Sub

    Public Overrides Sub OnComboSelectAfter(sboObject As Object, pVal As SBOItemEventArg)
        MyBase.OnComboSelectAfter(RuntimeHelpers.GetObjectValue(sboObject), pVal)
        Try
            Dim itemUID As String = pVal.ItemUID
            Dim flag As Boolean = Operators.CompareString(itemUID, "btnReview", False) = 0
            If flag Then
                Me.m_Form.Freeze(True)
                Dim selected As SAPbouiCOM.ValidValue = Me.btnReview.Selected
                Me.btnReview.Caption = selected.Description.ToString()
                Dim value As String = selected.Value
                flag = (Operators.CompareString(value, "1", False) = 0)
                If flag Then
                    Me.RefreshSiblingList()
                Else
                    Me.RefreshOneTime(selected.Value)
                End If
            Else
                flag = (Operators.CompareString(itemUID, "cboPost", False) = 0)
                If flag Then
                    Me.PostInvoices(Conversions.ToInteger(Me.cboPost.Selected.Value))
                    Me.cboPost.Caption = Me.cboPost.Selected.Description
                Else
                    flag = (Operators.CompareString(itemUID, "cboJump", False) = 0)
                    If flag Then
                        Me.m_Form.PaneLevel = Conversions.ToInteger(Me.cboJump.Selected.Value)
                        Me.SetPanelLevel(Conversions.ToInteger(Me.cboJump.Selected.Value))
                        Me.cboJump.Caption = Me.cboJump.Selected.Description
                        Me.UpdatePaneStatus()
                        Me.RetrievePageData()
                    End If
                End If
            End If
        Catch expr_16D As Exception
            ProjectData.SetProjectError(expr_16D)
            Dim ex As Exception = expr_16D
            Me.m_SboApplication.StatusBar.SetText(ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
            ProjectData.ClearProjectError()
        Finally
            Me.m_Form.Freeze(False)
        End Try
    End Sub

    Private Sub fillTextArray()
        Try
            Dim value As String = "Batch Selection"
            Dim value2 As String = "Add a new batch to start school generation fees generation or" + Environment.NewLine
            Dim value3 As String = "Select a new batch by clicking select new batch radio button." + Environment.NewLine
            Dim value4 As String = ""
            Me.aTitlelist.Add(value)
            Me.aDescList.Add(value2)
            Me.aDescList_1.Add(value3)
            Me.aDescList_2.Add(value4)
            value = "Review Students in Batch"
            value2 = "Please take some time to review the table below. "
            value3 = "It should contain all the students in the schools categorised into their key stage and classes."
            value4 = "You can also filter the list using to get more details per class or key stage."
            Me.aTitlelist.Add(value)
            Me.aDescList.Add(value2)
            Me.aDescList_1.Add(value3)
            Me.aDescList_2.Add(value4)
            value = "Review Batch"
            value2 = "The table below displays all the bills to generated in this batch."
            value3 = "Kindly, review the list for accuracy. "
            value4 = "You can go back to the previous screens if you need to make any changes."
            Me.aTitlelist.Add(value)
            Me.aDescList.Add(value2)
            Me.aDescList_1.Add(value3)
            Me.aDescList_2.Add(value4)
            value = "Email Generation"
            value2 = "Students information with parent emails are displayed below."
            value3 = "If you need to update an email, click on the arrow below to update the Business Partner."
            value4 = "Click the refresh button to if emails have been updated."
            Me.aTitlelist.Add(value)
            Me.aDescList.Add(value2)
            Me.aDescList_1.Add(value3)
            Me.aDescList_2.Add(value4)
            value = "Post Bills to General Ledger!"
            value2 = "By Clicking on the Post Bills below, A/R Invoices are generated"
            value3 = "Please click the button only if you have concluded the bills for this session."
            value4 = "Once the button is clicked, the action cannot be undone."
            Me.aTitlelist.Add(value)
            Me.aDescList.Add(value2)
            Me.aDescList_1.Add(value3)
            Me.aDescList_2.Add(value4)
        Catch expr_194 As Exception
            ProjectData.SetProjectError(expr_194)
            ProjectData.ClearProjectError()
        End Try
    End Sub

    Protected Overrides Sub UpdatePaneStatus()
        MyBase.UpdatePaneStatus()
        Dim currentPane As Integer = Me.CurrentPane
        Dim flag As Boolean = currentPane = 1
        ' The following expression was wrapped in a checked-statement
        If Not flag Then
            flag = (Me.CurrentPane >= 9)
            If Not flag Then
                Me.lblTitle2.Caption = Conversions.ToString(Me.aTitlelist(Me.CurrentPane - 2))
                Me.lblDesc2.Caption = Conversions.ToString(Me.aDescList(Me.CurrentPane - 2))
                Me.lblDesc2_1.Caption = Conversions.ToString(Me.aDescList_1(Me.CurrentPane - 2))
                Me.lblDesc2_2.Caption = Conversions.ToString(Me.aDescList_2(Me.CurrentPane - 2))
            End If
        End If
    End Sub

    Protected Overrides Sub RetrievePageData()
        MyBase.RetrievePageData()
        Select Case Me.CurrentPane
            Case 1
                Me.HandlePane1()
            Case 2
                Me.HandlePane2()
            Case 3
                Me.HandlePane3()
            Case 4
                Me.HandlePane4()
            Case 5
                Me.HandlePane5()
            Case 6
                Me.HandlePane6()
        End Select
    End Sub

    Private Sub HandlePane1()
    End Sub

    Private Sub HandlePane2()
    End Sub

    Private Sub HandlePane3()
        Me.RefreshStudentList(False)
    End Sub

    Private Sub HandlePane4()
        Try
            Me.RefreshSummary(False)
        Catch expr_0C As Exception
            ProjectData.SetProjectError(expr_0C)
            ProjectData.ClearProjectError()
        End Try
    End Sub

    Private Sub HandlePane5()
        Dim flag As Boolean = Not Me.lEmailLoaded
        If flag Then
            Me.refreshEmailList()
        End If
    End Sub

    Private Sub HandlePane6()
        Try
            Me.Invoiced = Me.m_DBDataSource0.GetValue("U_Invoiced", Me.m_DBDataSource0.Offset)
            Dim flag As Boolean = Operators.CompareString(Me.Invoiced, "Y", False) = 0
            If flag Then
                Me.btnPost.Item.Enabled = False
            Else
                Me.btnPost.Item.Enabled = True
            End If
            flag = Not Me.lInvoiceLoaded
            If flag Then
                Me.RefreshInvoice()
            End If
        Catch expr_79 As Exception
            ProjectData.SetProjectError(expr_79)
            ProjectData.ClearProjectError()
        End Try
    End Sub

    Private Sub RefreshStudentList(lRefresh As Boolean)
        ' The following expression was wrapped in a checked-statement
        Try
            Me.m_Form.Freeze(True)
            Me.DocEntry = Conversions.ToInteger(Me.m_DBDataSource0.GetValue("DocEntry", Me.m_DBDataSource0.Offset))
            If lRefresh Then
                Me.ExecuteSQLDT("BillWizardStudentList", Me.m_DataTable0, New String() {Me.DocEntry.ToString(), "Y"})
            Else
                Me.ExecuteSQLDT("BillWizardStudentList", Me.m_DataTable0, New String() {Me.DocEntry.ToString(), "N"})
            End If
            Me.grdStud.DataTable = Me.m_DataTable0
            Me.grdStud.CollapseLevel = 1
            Me.grdStud.Columns.Item(0).TitleObject.Caption = "School"
            Me.grdStud.Columns.Item(1).TitleObject.Caption = "Level"
            Me.grdStud.Columns.Item(2).TitleObject.Caption = "Class Name"
            Me.grdStud.Columns.Item(3).TitleObject.Caption = "No. In Class"
            Me.grdStud.Columns.Item(4).TitleObject.Caption = "Student Code"
            Me.grdStud.Columns.Item(5).TitleObject.Caption = "Student Name"
            Me.grdStud.Columns.Item(6).TitleObject.Caption = "Fee Type"
            Me.grdStud.Columns.Item(7).Visible = False
            Me.grdStud.Columns.Item(8).Visible = False
            Me.grdStud.Columns.Item(9).Visible = False
            Me.grdStud.Columns.Item(10).Visible = False
            Me.grdStud.Columns.Item(11).Visible = False
            Dim arg_26C_0 As Integer = 0
            Dim num As Integer = Me.grdStud.Columns.Count - 1
            Dim num2 As Integer = arg_26C_0
            While True
                Dim arg_297_0 As Integer = num2
                Dim num3 As Integer = num
                If arg_297_0 > num3 Then
                    Exit While
                End If
                Me.grdStud.Columns.Item(num2).Editable = False
                num2 += 1
            End While
            Me.editCol = Me.grdStud.Columns.Item("CardCode")
            NewLateBinding.LateSet(Me.editCol, Nothing, "LinkedObjectType", New Object() {2}, Nothing, Nothing)
            Me.grdStud.AutoResizeColumns()
            Me.UserDB.ValueEx = Conversions.ToString(Me.m_DataTable0.Rows.Count)
            Me.m_Form.Freeze(False)
        Catch expr_319 As Exception
            ProjectData.SetProjectError(expr_319)
            Dim ex As Exception = expr_319
            Me.m_SboApplication.StatusBar.SetText(ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
            ProjectData.ClearProjectError()
        End Try
    End Sub

    Private Sub RefreshSiblingList()
        ' The following expression was wrapped in a checked-statement
        Try
            Me.m_Form.Freeze(True)
            Me.DocEntry = Conversions.ToInteger(Me.m_DBDataSource0.GetValue("DocEntry", Me.m_DBDataSource0.Offset))
            Me.ExecuteSQLDT("BillWizardSiblingList", Me.m_DataTable1, New String() {Me.DocEntry.ToString()})
            Me.grdSibling.DataTable = Me.m_DataTable1
            Me.grdSibling.CollapseLevel = 1
            Me.grdSibling.Columns.Item(0).TitleObject.Caption = "Family Name"
            Me.grdSibling.Columns.Item(1).TitleObject.Caption = "Student Code"
            Me.grdSibling.Columns.Item(2).TitleObject.Caption = "Student Name"
            Dim arg_103_0 As Integer = 0
            Dim num As Integer = Me.grdSibling.Columns.Count - 1
            Dim num2 As Integer = arg_103_0
            While True
                Dim arg_12E_0 As Integer = num2
                Dim num3 As Integer = num
                If arg_12E_0 > num3 Then
                    Exit While
                End If
                Me.grdSibling.Columns.Item(num2).Editable = False
                num2 += 1
            End While
            Me.editCol = Me.grdSibling.Columns.Item("CardCode")
            NewLateBinding.LateSet(Me.editCol, Nothing, "LinkedObjectType", New Object() {2}, Nothing, Nothing)
            Me.grdSibling.AutoResizeColumns()
            Me.UserDB.ValueEx = Conversions.ToString(Me.m_DataTable1.Rows.Count)
            Me.m_Form.Freeze(False)
        Catch expr_1B0 As Exception
            ProjectData.SetProjectError(expr_1B0)
            Dim ex As Exception = expr_1B0
            Me.m_SboApplication.StatusBar.SetText(ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
            ProjectData.ClearProjectError()
        End Try
    End Sub

    Private Sub RefreshOneTime(strType As String)
        ' The following expression was wrapped in a checked-statement
        Try
            Dim strProc As String = ""
            Me.m_Form.Freeze(True)
            Me.m_DataTable2.Clear()
            Dim flag As Boolean = Operators.CompareString(strType, "2", False) = 0
            If flag Then
                strProc = "BillWizardOneTimeDiscountList"
            Else
                flag = (Operators.CompareString(strType, "3", False) = 0)
                If flag Then
                    strProc = "BillWizardStaffKidsList"
                Else
                    flag = (Operators.CompareString(strType, "4", False) = 0)
                    If flag Then
                        strProc = "BillWizardNewStudentsLists"
                    End If
                End If
            End If
            Me.DocEntry = Conversions.ToInteger(Me.m_DBDataSource0.GetValue("DocEntry", Me.m_DBDataSource0.Offset))
            Me.ExecuteSQLDT(strProc, Me.m_DataTable2, New String() {Me.DocEntry.ToString()})
            Me.grdSibling.DataTable = Me.m_DataTable2
            Me.grdSibling.CollapseLevel = 2
            Dim arg_102_0 As Integer = 0
            Dim num As Integer = Me.grdSibling.Columns.Count - 1
            Dim num2 As Integer = arg_102_0
            While True
                Dim arg_12E_0 As Integer = num2
                Dim num3 As Integer = num
                If arg_12E_0 > num3 Then
                    Exit While
                End If
                Me.grdSibling.Columns.Item(num2).Editable = False
                num2 += 1
            End While
            Me.editCol = Me.grdSibling.Columns.Item("Student Code")
            NewLateBinding.LateSet(Me.editCol, Nothing, "LinkedObjectType", New Object() {2}, Nothing, Nothing)
            Me.grdSibling.AutoResizeColumns()
            Me.UserDB.ValueEx = Conversions.ToString(Me.m_DataTable2.Rows.Count)
        Catch expr_1A3 As Exception
            ProjectData.SetProjectError(expr_1A3)
            Dim ex As Exception = expr_1A3
            Me.m_SboApplication.StatusBar.SetText(ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
            ProjectData.ClearProjectError()
        Finally
            Me.m_Form.Freeze(False)
        End Try
    End Sub

    Private Sub refreshEmailList()
        ' The following expression was wrapped in a checked-statement
        Try
            Me.m_Form.Freeze(True)
            Me.DocEntry = Conversions.ToInteger(Me.m_DBDataSource0.GetValue("DocEntry", Me.m_DBDataSource0.Offset))
            Me.ExecuteSQLDT("BillWizardEmailList", Me.m_DataTable2, New String() {Me.DocEntry.ToString()})
            Me.grdEmail.DataTable = Me.m_DataTable2
            Me.grdEmail.CollapseLevel = 1
            Dim arg_91_0 As Integer = 0
            Dim num As Integer = Me.grdEmail.Columns.Count - 1
            Dim num2 As Integer = arg_91_0
            While True
                Dim arg_BC_0 As Integer = num2
                Dim num3 As Integer = num
                If arg_BC_0 > num3 Then
                    Exit While
                End If
                Me.grdEmail.Columns.Item(num2).Editable = False
                num2 += 1
            End While
            Me.editCol = Me.grdEmail.Columns.Item("Student Code")
            NewLateBinding.LateSet(Me.editCol, Nothing, "LinkedObjectType", New Object() {2}, Nothing, Nothing)
            Me.grdEmail.AutoResizeColumns()
            Me.UserDB.ValueEx = Conversions.ToString(Me.m_DataTable2.Rows.Count)
            Me.lEmailLoaded = True
        Catch expr_138 As Exception
            ProjectData.SetProjectError(expr_138)
            Dim ex As Exception = expr_138
            Me.m_SboApplication.StatusBar.SetText(ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
            ProjectData.ClearProjectError()
        Finally
            Me.m_Form.Freeze(False)
        End Try
    End Sub

    Private Sub refreshSpecial()
        Try
            Me.m_Form.Freeze(True)
            Me.DocEntry = Conversions.ToInteger(Me.m_DBDataSource0.GetValue("DocEntry", Me.m_DBDataSource0.Offset))
            Dim text As String = "Select U_Class Class,ISNULL(U_Term,0) Term, U_ItemCode [Item Code], U_ItemName [Item Name], U_Amount Amount from [@OWA_EDUBILLSPECIAL]"
            text = text + " Where DocEntry = " + Me.DocEntry.ToString()
            Me.m_DataTable1.Clear()
            Me.m_DataTable1.ExecuteQuery(text)
            Me.grdSpecial.DataTable = Me.m_DataTable1
            Me.editCol = Me.grdSpecial.Columns.Item("Item Code")
            NewLateBinding.LateSet(Me.editCol, Nothing, "LinkedObjectType", New Object() {4}, Nothing, Nothing)
            Me.grdSpecial.Columns.Item("Class").Type = BoGridColumnType.gct_ComboBox
            Dim comboBoxColumn As ComboBoxColumn = CType(Me.grdSpecial.Columns.Item("Class"), ComboBoxColumn)
            Me.fillCombo("U_clsCode", "U_clsDesc", "@OWA_EDUKEYSTALINES", comboBoxColumn.ValidValues, "", False, BoFieldTypes.db_Alpha, False, "", True)
            Me.grdSpecial.AutoResizeColumns()
            Me.UserDB.ValueEx = Conversions.ToString(Me.m_DataTable1.Rows.Count)
            Me.lEmailLoaded = True
        Catch expr_156 As Exception
            ProjectData.SetProjectError(expr_156)
            Dim ex As Exception = expr_156
            Me.m_SboApplication.StatusBar.SetText(ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
            ProjectData.ClearProjectError()
        Finally
            Me.m_Form.Freeze(False)
        End Try
    End Sub

    Private Sub SaveStudentList()
        ' The following expression was wrapped in a checked-statement
        Try
            Me.DocEntry = Conversions.ToInteger(Me.m_DBDataSource0.GetValue("DocEntry", Me.m_DBDataSource0.Offset))
            Dim generalService As GeneralService = Me.m_ParentAddon.SboCompany.GetCompanyService().GetGeneralService("OWAEDUBILLING")
            Dim generalDataParams As GeneralDataParams = CType(generalService.GetDataInterface(GeneralServiceDataInterfaces.gsGeneralDataParams), GeneralDataParams)
            generalDataParams.SetProperty("DocEntry", Me.DocEntry.ToString())
            Dim byParams As GeneralData = generalService.GetByParams(generalDataParams)
            Dim generalDataCollection As GeneralDataCollection = byParams.Child("OWA_EDUBILLSUMM")
            Me.oProgBar = Me.m_SboApplication.StatusBar.CreateProgressBar("Oryx01", Me.m_DataTable0.Rows.Count - 1, False)
            Me.oProgBar.Value = 0
            Dim arg_D1_0 As Integer = 0
            Dim num As Integer = Me.m_DataTable0.Rows.Count - 1
            Dim num2 As Integer = arg_D1_0
            While True
                Dim arg_1DF_0 As Integer = num2
                Dim num3 As Integer = num
                If arg_1DF_0 > num3 Then
                    Exit While
                End If
                Dim flag As Boolean = Operators.ConditionalCompareObjectEqual(Me.m_DataTable0.GetValue("cType", num2), "A", False)
                If flag Then
                    Dim generalData As GeneralData = generalDataCollection.Add()
                    Dim vtValue As String = Conversions.ToString(Me.m_DataTable0.GetValue("CardCode", num2))
                    Dim generalData2 As GeneralData = generalData
                    generalData2.SetProperty("U_CardCode", vtValue)
                    generalData2.SetProperty("U_CardName", RuntimeHelpers.GetObjectValue(Me.m_DataTable0.GetValue("CardName", num2)))
                    generalData2.SetProperty("U_KeyStage", RuntimeHelpers.GetObjectValue(Me.m_DataTable0.GetValue("KeyStage", num2)))
                    generalData2.SetProperty("U_Class", RuntimeHelpers.GetObjectValue(Me.m_DataTable0.GetValue("ClassCode", num2)))
                End If
                Me.oProgBar.Text = Me.m_DataTable0.GetValue("CardCode", num2).ToString() + " added successfully"
                Me.oProgBar.Value = 2
                num2 += 1
            End While
            Me.oProgBar.[Stop]()
            Me.oProgBar = Nothing
            generalService.Update(byParams)
        Catch expr_202 As Exception
            ProjectData.SetProjectError(expr_202)
            Dim ex As Exception = expr_202
            Me.m_SboApplication.StatusBar.SetText(ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
            ProjectData.ClearProjectError()
        Finally
            Dim flag As Boolean = Me.oProgBar IsNot Nothing
            If flag Then
                Me.oProgBar.[Stop]()
                Me.oProgBar = Nothing
            End If
        End Try
    End Sub

    Private Sub SaveBatchHeader()
        Try
            Dim generalService As GeneralService = Me.m_ParentAddon.SboCompany.GetCompanyService().GetGeneralService("OWAEDUBILLING")
            Dim generalDataParams As GeneralDataParams = CType(generalService.GetDataInterface(GeneralServiceDataInterfaces.gsGeneralDataParams), GeneralDataParams)
            Dim text As String = "Select Count('A') from [@OWA_EDUBILLINGS] Where U_Session = 'xxxx'"
            text = text.Replace("xxxx", Me.m_DBDataSource0.GetValue("U_Session", Me.m_DBDataSource0.Offset).Trim())
            Dim dataTable As DataTable = Me.ExecuteSQLDT(text)
            Dim flag As Boolean = Operators.ConditionalCompareObjectGreater(dataTable.GetValue(0, 0), 0, False)
            If flag Then
                Me.m_SboApplication.StatusBar.SetText("Batch for session already exists", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
            Else
                flag = Me.lAdd
                If flag Then
                    Dim generalData As GeneralData = CType(generalService.GetDataInterface(GeneralServiceDataInterfaces.gsGeneralData), GeneralData)
                    generalData.SetProperty("U_BatchDesc", Me.m_DBDataSource0.GetValue("U_BatchDesc", Me.m_DBDataSource0.Offset))
                    generalData.SetProperty("U_Session", Me.m_DBDataSource0.GetValue("U_Session", Me.m_DBDataSource0.Offset))
                    generalService.Add(generalData)
                    Me.m_SboApplication.StatusBar.SetText("Batch Added Successfully", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success)
                    Me.m_DBDataSource0.Clear()
                    Me.m_DBDataSource0.Query(Nothing)
                    ' The following expression was wrapped in a checked-expression
                    Me.m_DBDataSource0.Offset = Me.m_DBDataSource0.Size - 1
                    Me.m_Form.Mode = BoFormMode.fm_OK_MODE
                    Me.optUpdate.Selected = True
                    Me.RefreshPane2Controls(2)
                End If
            End If
        Catch expr_187 As Exception
            ProjectData.SetProjectError(expr_187)
            Dim ex As Exception = expr_187
            Me.m_SboApplication.StatusBar.SetText(ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
            ProjectData.ClearProjectError()
        End Try
    End Sub

    Private Sub SaveStudentList2()
        ' The following expression was wrapped in a checked-statement
        Me.RefreshStudentList(True)
        'Try
        '    Me.RefreshStudentList(True)
        '    Me.DocEntry = Conversions.ToInteger(Me.m_DBDataSource0.GetValue("DocEntry", Me.m_DBDataSource0.Offset))
        '    Dim text As String = " Delete [@OWA_EDUBILLSUMM] Where (U_InvNum is null or U_InvNum = 0) and U_Batch = " + Me.DocEntry.ToString()
        '    Me.ExecuteSQLDT(text)
        '    text = " Delete [@OWA_EDUBILLDET] Where  U_Batch = " + Me.DocEntry.ToString()
        '    text = text + " And U_CardCode not in (Select U_CardCode from [@OWA_EDUBILLSUMM] Where U_Batch = " + Me.DocEntry.ToString() + ")"
        '    Me.ExecuteSQLDT(text)
        '    text = "Select Max(ISNULL(DocEntry,0)) + 1 from [@OWA_EDUBILLSUMM]"
        '    Dim dataTable As DataTable = Me.ExecuteSQLDT(text)
        '    Dim num As Integer = Conversions.ToInteger(dataTable.GetValue(0, 0))
        '    Dim flag As Boolean = num = 0
        '    If flag Then
        '        num += 1
        '    End If
        '    Me.m_Form.Freeze(True)
        '    Me.oProgBar = Me.m_SboApplication.StatusBar.CreateProgressBar("Oryx01", Me.m_DataTable0.Rows.Count - 1, False)
        '    Me.oProgBar.Value = 0
        '    Dim arg_10D_0 As Integer = 0
        '    Dim num2 As Integer = Me.m_DataTable0.Rows.Count - 1
        '    Dim num3 As Integer = arg_10D_0
        '    While True
        '        Dim arg_26D_0 As Integer = num3
        '        Dim num4 As Integer = num2
        '        If arg_26D_0 > num4 Then
        '            Exit While
        '        End If
        '        num += 1
        '        Dim dataTable2 As DataTable = Me.m_DataTable0
        '        If dataTable2.GetValue("cType", num3).ToString() = "A" Then
        '            Me.ExecuteSQLDT("BillWizardSaveSummary", New String() {num.ToString(),
        '                                                              Me.DocEntry.ToString(), "OWAEDUBILLING",
        '                                                              Conversions.ToString(dataTable2.GetValue("CardCode", num3)),
        '                                                              dataTable2.GetValue("CardName", num3).ToString().Replace("'", ""),
        '                                                              dataTable2.GetValue("School", num3).ToString(),
        '                                                              Conversions.ToString(dataTable2.GetValue("ClassCode", num3)),
        '                                                              Conversions.ToString(dataTable2.GetValue("FeeType", num3)),
        '                                                              Conversions.ToString(dataTable2.GetValue("nLevel", num3)),
        '                                                              Conversions.ToString(dataTable2.GetValue("Dim1", num3)),
        '                                                              Conversions.ToString(dataTable2.GetValue("Dim2", num3))})

        '        End If


        '        Me.oProgBar.Text = Me.m_DataTable0.GetValue("CardCode", num3).ToString() + " added successfully"
        '        Me.oProgBar.Value = 2
        '        num3 += 1
        '    End While
        '    Me.m_DBDataSource0.Clear()
        '    Me.m_DBDataSource0.Query(Nothing)
        '    Me.getOffset(Me.DocEntry.ToString(), "DocEntry", Me.m_DBDataSource0)
        '    Me.m_Form.Freeze(False)
        '    Me.oProgBar.[Stop]()
        '    Me.oProgBar = Nothing
        '    Me.RefreshSummary(True)
        'Catch expr_2CA As Exception
        '    ProjectData.SetProjectError(expr_2CA)
        '    Dim ex As Exception = expr_2CA
        '    Me.m_SboApplication.StatusBar.SetText(ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
        '    ProjectData.ClearProjectError()
        'Finally
        '    Dim flag As Boolean = Me.oProgBar IsNot Nothing
        '    If flag Then
        '        Me.oProgBar.[Stop]()
        '        Me.oProgBar = Nothing
        '    End If
        '    Me.m_Form.Freeze(False)
        'End Try
    End Sub

    Private Sub SaveSpecial()
        ' The following expression was wrapped in a checked-statement
        Try
            Me.DocEntry = Conversions.ToInteger(Me.m_DBDataSource0.GetValue("DocEntry", Me.m_DBDataSource0.Offset))
            Dim strSQL As String = " Delete [@OWA_EDUBILLSPECIAL] Where DocEntry = " + Me.DocEntry.ToString()
            Me.ExecuteSQLDT(strSQL)
            Me.m_Form.Freeze(True)
            Dim arg_68_0 As Integer = 0
            Dim num As Integer = Me.m_DataTable1.Rows.Count - 1
            Dim num2 As Integer = arg_68_0
            While True
                Dim arg_167_0 As Integer = num2
                Dim num3 As Integer = num
                If arg_167_0 > num3 Then
                    Exit While
                End If
                Dim flag As Boolean = Operators.ConditionalCompareObjectNotEqual(Me.m_DataTable1.GetValue("Class", num2), "", False)
                If flag Then
                    Dim num4 As Integer = num2 + 1
                    Me.ExecuteSQLDT("BillWizardBatchSpecial", New String() {Me.DocEntry.ToString(), num4.ToString(), "OWAEDUBILLING", Conversions.ToString(Me.m_DataTable1.GetValue("Item Code", num2)), Conversions.ToString(Me.m_DataTable1.GetValue("Item Name", num2)), Me.m_DataTable1.GetValue("Amount", num2).ToString(), Conversions.ToString(Me.m_DataTable1.GetValue("Class", num2)), Me.m_DataTable1.GetValue("Term", num2).ToString()})
                End If
                num2 += 1
            End While
            Me.m_DBDataSource0.Clear()
            Me.m_DBDataSource0.Query(Nothing)
            Me.getOffset(Me.DocEntry.ToString(), "DocEntry", Me.m_DBDataSource0)
            Me.m_Form.Freeze(False)
        Catch expr_1B1 As Exception
            ProjectData.SetProjectError(expr_1B1)
            Dim ex As Exception = expr_1B1
            Me.m_SboApplication.StatusBar.SetText(ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
            ProjectData.ClearProjectError()
        Finally
            Me.m_Form.Freeze(False)
        End Try
    End Sub

    Private Sub DeleteBatch()
        Try
            Dim flag As Boolean = Not Me.lAdd
            If flag Then
                Dim generalService As GeneralService = Me.m_ParentAddon.SboCompany.GetCompanyService().GetGeneralService("OWAEDUBILLING")
                Dim generalDataParams As GeneralDataParams = CType(generalService.GetDataInterface(GeneralServiceDataInterfaces.gsGeneralDataParams), GeneralDataParams)
                Me.DocEntry = Conversions.ToInteger(Me.m_DBDataSource0.GetValue("DocEntry", Me.m_DBDataSource0.Offset))
                generalDataParams.SetProperty("DocEntry", Me.DocEntry.ToString())
                generalService.Delete(generalDataParams)
                Me.m_SboApplication.StatusBar.SetText("Batch Deleted Successfully", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success)
                Me.m_DBDataSource0.Clear()
                Me.m_DBDataSource0.Query(Nothing)
                Me.optUpdate.Selected = True
                Me.RefreshPane2Controls(2)
                Me.m_Form.Mode = BoFormMode.fm_OK_MODE
            End If
        Catch expr_D7 As Exception
            ProjectData.SetProjectError(expr_D7)
            Dim ex As Exception = expr_D7
            Me.m_SboApplication.StatusBar.SetText(ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
            ProjectData.ClearProjectError()
        End Try
    End Sub

    Protected Overrides Function IsReady(ByRef pErrNo As Integer, ByRef pErrMsg As String) As Boolean
        Return MyBase.IsReady(pErrNo, pErrMsg)
    End Function

    Private Sub RefreshSummary(LRefresh As Boolean)
        ' The following expression was wrapped in a checked-statement
        Try
            Dim flag As Boolean = Not LRefresh And Me.lReviewLoaded
            If Not flag Then
                Me.m_Form.Freeze(True)
                Me.DocEntry = Conversions.ToInteger(Me.m_DBDataSource0.GetValue("DocEntry", Me.m_DBDataSource0.Offset))
                If LRefresh Then
                    Me.ExecuteSQLDT("BillWizardBatchSummary", Me.m_DataTable3, New String() {Me.DocEntry.ToString(), "Y"})
                Else
                    Me.ExecuteSQLDT("BillWizardBatchSummary", Me.m_DataTable3, New String() {Me.DocEntry.ToString()})
                End If
                Me.grdReview.DataTable = Me.m_DataTable3
                Me.grdReview.CollapseLevel = 1
                Me.grdReview.Columns.Item(0).TitleObject.Caption = "School"
                Me.grdReview.Columns.Item(1).TitleObject.Caption = "Class"
                Me.grdReview.Columns.Item(2).TitleObject.Caption = "Student Code"
                Me.grdReview.Columns.Item(3).TitleObject.Caption = "Student Name"
                Me.grdReview.Columns.Item(4).TitleObject.Caption = "Total"
                Me.grdReview.Columns.Item(5).TitleObject.Caption = "Level"
                Dim arg_1CA_0 As Integer = 0
                Dim num As Integer = Me.grdReview.Columns.Count - 1
                Dim num2 As Integer = arg_1CA_0
                While True
                    Dim arg_1F5_0 As Integer = num2
                    Dim num3 As Integer = num
                    If arg_1F5_0 > num3 Then
                        Exit While
                    End If
                    Me.grdReview.Columns.Item(num2).Editable = False
                    num2 += 1
                End While
                Me.editCol = Me.grdReview.Columns.Item("U_CardCode")
                NewLateBinding.LateSet(Me.editCol, Nothing, "LinkedObjectType", New Object() {2}, Nothing, Nothing)
                Me.grdReview.AutoResizeColumns()
                Me.UserDB.ValueEx = Conversions.ToString(Me.m_DataTable3.Rows.Count)
                Me.m_DBDataSource0.Clear()
                Me.m_DBDataSource0.Query(Nothing)
                Me.getOffset(Me.DocEntry.ToString(), "DocEntry", Me.m_DBDataSource0)
                Me.lReviewLoaded = True
                Me.m_Form.Freeze(False)
            End If
        Catch expr_2B4 As Exception
            ProjectData.SetProjectError(expr_2B4)
            Dim ex As Exception = expr_2B4
            Me.m_SboApplication.StatusBar.SetText(ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
            ProjectData.ClearProjectError()
        End Try
    End Sub

    Private Sub RefreshPane2Controls(intOption As Integer)
        Try
            Me.m_Form.Freeze(True)
            Dim flag As Boolean = intOption = 1
            If flag Then
                Me.btnDelete.Item.Visible = False
                Me.btnSelect.Item.Visible = False
                Me.txtSession.Item.Enabled = True
            Else
                Me.btnDelete.Item.Visible = True
                Me.btnSelect.Item.Visible = True
                Me.txtSession.Item.Enabled = False
                Me.txtDesc.Item.Enabled = False
            End If
        Catch expr_9B As Exception
            ProjectData.SetProjectError(expr_9B)
            ProjectData.ClearProjectError()
        Finally
            Me.m_Form.Freeze(False)
        End Try
    End Sub

    Private Sub RefreshInvoice()
        ' The following expression was wrapped in a checked-statement
        Try
            Me.DocEntry = Conversions.ToInteger(Me.m_DBDataSource0.GetValue("DocEntry", Me.m_DBDataSource0.Offset))
            Dim text As String = String.Format("Select b.Name [School], c.Name [Class], a.U_CardCode [Student Code], a.U_CardName [Student Name] ," +
             " a.U_TotalFees [Total Fees], U_Level, d.DocEntry [Invoice Number]  from [@OWA_EDUBILLSUMM] a join [@OWA_EDUSCHOOLS] b on a.U_school = b.code " +
             " Join [@OWA_EDUCLASS] c  On a.U_Class = c.Code Left outer join (Select DocEntry from OINV where canceled = 'N' and U_Batch = {0}) d on a.U_invNum = d.DocEntry " +
             " Where U_Batch = {0}  order by 1,2,3 ", Me.DocEntry.ToString())
            'text += " and U_InvNum not in (Select DocEntry from OINV where canceled = 'Y') "

            Me.m_Form.Freeze(True)
            Me.m_DataTable4.Clear()
            Me.m_DataTable4.ExecuteQuery(text)
            Me.grdInvoice.DataTable = Me.m_DataTable4
            Me.grdInvoice.CollapseLevel = 1
            Dim arg_CE_0 As Integer = 0
            Dim num As Integer = Me.grdInvoice.Columns.Count - 1
            Dim num2 As Integer = arg_CE_0
            While True
                Dim arg_F9_0 As Integer = num2
                Dim num3 As Integer = num
                If arg_F9_0 > num3 Then
                    Exit While
                End If
                Me.grdInvoice.Columns.Item(num2).Editable = False
                num2 += 1
            End While
            Me.editCol = Me.grdInvoice.Columns.Item("Student Code")
            NewLateBinding.LateSet(Me.editCol, Nothing, "LinkedObjectType", New Object() {2}, Nothing, Nothing)
            Me.editCol = Me.grdInvoice.Columns.Item("Invoice Number")
            NewLateBinding.LateSet(Me.editCol, Nothing, "LinkedObjectType", New Object() {13}, Nothing, Nothing)
            Me.grdInvoice.AutoResizeColumns()
            Me.UserDB.ValueEx = Conversions.ToString(Me.m_DataTable4.Rows.Count)
            Me.lInvoiceLoaded = True
        Catch expr_1BA As Exception
            ProjectData.SetProjectError(expr_1BA)
            Dim ex As Exception = expr_1BA
            Me.m_SboApplication.StatusBar.SetText(ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
            ProjectData.ClearProjectError()
        Finally
            Me.m_Form.Freeze(False)
        End Try
    End Sub

    Private Sub CancelInvoices()
        Try
            Dim str As String = ""
            Dim text As String = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
            Dim dataTable As DataTable = Me.m_Form.DataSources.DataTables.Item("DT_InvH")
            Dim dataTable2 As DataTable = Me.m_Form.DataSources.DataTables.Item("DT_InvD")
            Dim dataTable3 As DataTable = Me.m_Form.DataSources.DataTables.Item("DT_0")
            Me.ExecuteSQLDT("BillWizardLastBackup", dataTable3, New String() {Me.m_SboApplication.Company.DatabaseName.ToString().Trim(), text})
            Dim flag As Boolean = Not dataTable3.IsEmpty
            If flag Then
                Dim num As Integer = Conversions.ToInteger(dataTable3.GetValue("diff", 0))
                flag = (num > 30)
                If flag Then
                    Me.m_SboApplication.MessageBox("A backup is needed within 30 mins of generating invoices,", 1, "Ok", "", "")
                Else
                    Dim text3 As String = String.Format("Select DocEntry From OINV Where U_Batch = {0} and  canceled = 'N'", Me.DocEntry.ToString())


                    dataTable.Clear()
                    dataTable.ExecuteQuery(text3)
                    flag = dataTable.IsEmpty
                    If flag Then
                        Me.m_SboApplication.StatusBar.SetText("No invoices to post", BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Warning)
                    Else
                        Dim documents As Documents = CType(Me.m_ParentAddon.SboCompany.GetBusinessObject(BoObjectTypes.oInvoices), Documents)

                        For i = 0 To dataTable.Rows.Count - 1
                            Dim key As Integer = Conversions.ToInteger(dataTable.GetValue("DocEntry", i))
                            documents.GetByKey(key)
                            Dim oCancelDoc As SAPbobsCOM.Documents = documents.CreateCancellationDocument()
                            oCancelDoc.Add()
                            Me.m_SboApplication.StatusBar.SetText("Invoice " + key.ToString() + "cancelled successfully", BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Success)
                        Next

                    End If
                End If
            End If

        Catch ex As Exception
            Me.m_SboApplication.StatusBar.SetText(ex.Message, BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Private Sub PostInvoices(Optional TermType As Integer = -1)
        ' The following expression was wrapped in a checked-statement
        Try
            Dim str As String = ""
            Dim text As String = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
            Dim dataTable As DataTable = Me.m_Form.DataSources.DataTables.Item("DT_InvH")
            Dim dataTable2 As DataTable = Me.m_Form.DataSources.DataTables.Item("DT_InvD")
            Dim dataTable3 As DataTable = Me.m_Form.DataSources.DataTables.Item("DT_0")
            Me.ExecuteSQLDT("BillWizardLastBackup", dataTable3, New String() {Me.m_SboApplication.Company.DatabaseName.ToString().Trim(), text})
            Dim flag As Boolean = Not dataTable3.IsEmpty
            If flag Then
                Dim num As Integer = Conversions.ToInteger(dataTable3.GetValue("diff", 0))
                flag = (num > 30)
                If flag Then
                    Me.m_SboApplication.MessageBox("A backup is needed within 30 mins of generating invoices,", 1, "Ok", "", "")
                Else
                    Dim offset As Integer = Me.m_DBDataSource0.Offset
                    Dim value As String = Me.m_DBDataSource0.GetValue("U_BatchDesc", offset)
                    Dim value2 As String = Me.m_DBDataSource0.GetValue("U_Session", offset)
                    Me.DocEntry = Conversions.ToInteger(Me.m_DBDataSource0.GetValue("DocEntry", offset))
                    flag = Me.getOffset(Me.m_DBDataSource0.GetValue("U_Session", offset), "Code", Me.m_DBDataSource3)
                    Dim dateTime As DateTime
                    If flag Then
                        dateTime = Me.sboDate(Me.m_DBDataSource3.GetValue("U_Start", Me.m_DBDataSource3.Offset))
                    End If
                    Dim flag2 As Boolean = True
                    Dim mdocH As SBOEDUBillWizard.DocH
                    mdocH.BatchDate = dateTime
                    mdocH.Session = value2
                    Me.oProgBar = Me.m_SboApplication.StatusBar.CreateProgressBar("Oryx01", (dataTable.Rows.Count - 1) * 3, False)

                    Dim num2 As Integer
                    'text2 = String.Concat(New String() {text2, "  and a.U_Batch =  ", Me.DocEntry.ToString(), ")"})
                    Dim text3 As String = String.Format("Select a.U_CardCode, a.U_CardName, a.U_School, a.U_Class, a.U_level, b.U_Dim1, b.U_Dim2 " +
                     "from [@OWA_EDUBILLSUMM] a join Ocrd b on a.U_cardCode = b.CardCode Where a.U_batch = {0} and b.frozenFor <> 'Y' " +
                     " and a.U_CardCode not in (Select CardCode from OINV where canceled = 'N' and U_Batch = {0}) Order By  a.U_Class", Me.DocEntry.ToString())


                    dataTable.Clear()
                    dataTable.ExecuteQuery(text3)
                    flag = dataTable.IsEmpty
                    If flag Then
                        Me.m_SboApplication.StatusBar.SetText("No invoices to post", BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Warning)
                    Else
                        flag = (Me.oProgBar IsNot Nothing)
                        If flag Then
                            Me.oProgBar.Value = 0
                        End If
                        Dim arg_30A_0 As Integer = 0
                        Dim num3 As Integer = dataTable.Rows.Count - 1
                        Dim num4 As Integer = arg_30A_0
                        Dim flag3 As Boolean
                        While True
                            Dim arg_543_0 As Integer = num4
                            Dim num5 As Integer = num3
                            If arg_543_0 > num5 Then
                                GoTo IL_548
                            End If
                            Dim invH As SBOEDUBillWizard.InvHeader
                            invH.CardCode = Conversions.ToString(dataTable.GetValue("U_CardCode", num4))
                            invH.CardName = Conversions.ToString(dataTable.GetValue("U_CardName", num4))
                            invH.School = Conversions.ToString(dataTable.GetValue("U_School", num4))
                            invH.strClass = Conversions.ToString(dataTable.GetValue("U_Class", num4))
                            invH.Dim1 = Conversions.ToString(dataTable.GetValue("U_Dim1", num4))
                            invH.Dim2 = Conversions.ToString(dataTable.GetValue("U_Dim2", num4))
                            invH.strRemarks = value2 + " - Invoice for " + invH.CardName
                            invH.TermDate = dateTime
                            Me.oProgBar.Text = String.Concat(New String() {"Creating Invoice for - ", invH.CardName.Trim(), " in ", invH.School, " and ", invH.strClass})
                            Dim text4 As String = "Select a.DocEntry, a.U_CardCode, a.U_Amount,  a.U_ItemCode, a.U_ItemName"
                            text4 += " from [@OWA_EDUBILLDET] a  "
                            text4 = text4 + " where a.U_CardCode = '" + invH.CardCode.Trim() + "'"
                            text4 = text4 + " and a.U_Batch = " + Me.DocEntry.ToString()
                            dataTable2.Clear()
                            dataTable2.ExecuteQuery(text4)
                            flag = dataTable2.IsEmpty
                            If Not flag Then
                                Dim num6 As Integer = 0
                                flag = Not Me.OINVCreate(mdocH, invH, dataTable2, str, num6)
                                If flag Then
                                    Exit While
                                End If
                                flag3 = Not Information.IsNothing(Me.oProgBar)
                                If flag3 Then
                                    Me.oProgBar.Text = "Invoice created sucessfully for - " + invH.CardName.Trim()
                                    Dim progressBar As ProgressBar = Me.oProgBar
                                    progressBar.Value += 1
                                End If
                            End If
                            num4 += 1
                        End While
                        flag3 = Me.m_ParentAddon.SboCompany.InTransaction
                        If flag3 Then
                            Me.m_ParentAddon.SboCompany.EndTransaction(BoWfTransOpt.wf_RollBack)
                        End If
                        flag2 = False
IL_548:
                        flag3 = (Me.oProgBar IsNot Nothing)
                        If flag3 Then
                            Me.oProgBar.[Stop]()
                            Me.oProgBar = Nothing
                        End If
                        flag3 = flag2
                        If flag3 Then
                            Me.m_SboApplication.StatusBar.SetText("Invoices for " + value2 + " Created Sucessfully", BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Success)
                            flag3 = Me.m_ParentAddon.SboCompany.InTransaction
                            If flag3 Then
                                Me.m_ParentAddon.SboCompany.EndTransaction(BoWfTransOpt.wf_Commit)
                            End If
                        Else
                            Me.m_SboApplication.StatusBar.SetText(value2 + " - " + str, BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Error)
                            flag3 = Me.m_ParentAddon.SboCompany.InTransaction
                            If flag3 Then
                                Me.m_ParentAddon.SboCompany.EndTransaction(BoWfTransOpt.wf_RollBack)
                            End If
                        End If
                        Dim str2 As String = " [U_InvNum] "
                        Dim text5 As String = " Update [@OWA_EDUBILLSUMM] Set " + str2 + " = b.DocEntry From [@OWA_EDUBILLSUMM] a JOIN "
                        text5 += "(Select a.DocEntry, a.U_Batch, a.U_TermType, a.CardCode from oinv a Where "
                        text5 = String.Concat(New String() {text5, " a.U_Batch = ", Me.DocEntry.ToString(), " and a.U_TermType = ", num2.ToString()})
                        text5 += " ) b On a.U_Batch = b.U_Batch and a.U_CardCode = b.CardCode "
                        Me.ExecuteSQLDT(text5)
                        Me.RefreshInvoice()
                    End If
                End If
            Else
                Me.m_SboApplication.MessageBox("Error getting backup information ", 1, "Ok", "", "")
            End If
        Catch expr_6AD As Exception
            ProjectData.SetProjectError(expr_6AD)
            Dim ex As Exception = expr_6AD
            Me.m_SboApplication.StatusBar.SetText(ex.Message, BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Error)
            Dim flag3 As Boolean = Me.m_ParentAddon.SboCompany.InTransaction
            If flag3 Then
                Me.m_ParentAddon.SboCompany.EndTransaction(BoWfTransOpt.wf_RollBack)
            End If
            ProjectData.ClearProjectError()
        Finally
            Dim flag3 As Boolean = Me.oProgBar IsNot Nothing
            If flag3 Then
                Me.oProgBar.[Stop]()
                Me.oProgBar = Nothing
            End If
            flag3 = Me.m_ParentAddon.SboCompany.InTransaction
            If flag3 Then
                Me.m_ParentAddon.SboCompany.EndTransaction(BoWfTransOpt.wf_Commit)
            End If
        End Try
    End Sub

    Private Function OINVCreate(mdocH As SBOEDUBillWizard.DocH, invH As SBOEDUBillWizard.InvHeader, oDT As DataTable, ByRef errMsg As String, ByRef errCode As Integer) As Boolean
        Dim result As Boolean = True
        Dim documents As Documents = CType(Me.m_ParentAddon.SboCompany.GetBusinessObject(BoObjectTypes.oInvoices), Documents)
        documents.Series = 0
        documents.CardCode = invH.CardCode
        documents.CardName = invH.CardName
        documents.NumAtCard = Me.DocEntry.ToString() + " - " + mdocH.Session
        documents.HandWritten = BoYesNoEnum.tNO
        documents.PaymentGroupCode = Conversions.ToInteger("-1")
        documents.DocDate = invH.TermDate
        documents.DocDueDate = invH.TermDate
        documents.Comments = invH.strRemarks.Trim()
        documents.UserFields.Fields.Item("U_Batch").Value = Me.DocEntry
        documents.DocType = BoDocumentTypes.dDocument_Service
        ' The following expression was wrapped in a checked-statement
        Dim num As Integer = oDT.Rows.Count - 1
        Dim arg_ED_0 As Integer = 0
        Dim num2 As Integer = num
        Dim num3 As Integer = arg_ED_0
        Dim flag As Boolean
        While True
            Dim arg_21B_0 As Integer = num3
            Dim num4 As Integer = num2
            If arg_21B_0 > num4 Then
                Exit While
            End If
            Dim num5 As Double = Conversions.ToDouble(oDT.GetValue("U_Amount", num3))
            documents.Lines.AccountCode = Conversions.ToString(oDT.GetValue("U_ItemCode", num3))
            documents.Lines.ItemDescription = Conversions.ToString(oDT.GetValue("U_ItemName", num3))
            documents.Lines.UnitPrice = num5
            documents.Lines.LineTotal = num5
            documents.Lines.CostingCode = invH.Dim1
            documents.Lines.CostingCode2 = invH.Dim2
            documents.Lines.UserFields.Fields.Item("U_Class").Value = invH.strClass
            documents.Lines.UserFields.Fields.Item("U_School").Value = invH.School
            documents.Lines.UserFields.Fields.Item("U_Session").Value = mdocH.Session
            flag = (num3 <> num)
            If flag Then
                documents.Lines.Add()
            End If
            num3 += 1
        End While
        flag = (documents.Add() <> 0)
        If flag Then
            Me.m_ParentAddon.SboCompany.GetLastError(errCode, errMsg)
            errMsg = String.Concat(New String() {errMsg, " BP - ", invH.CardCode, "Session ", mdocH.Session.ToString()})
            Me.m_SboApplication.StatusBar.SetText(errMsg, BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Error)
            result = False
        End If
        Return result
    End Function
End Class

