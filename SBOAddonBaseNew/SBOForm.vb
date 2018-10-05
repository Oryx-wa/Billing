Option Explicit On

Imports SAPbouiCOM.Framework
Imports SBO.SboAddOnBase
Imports System.Text

Public MustInherit Class SboForm
    Inherits SboFormBase

    Protected DBDataSource0 As SAPbouiCOM.DBDataSource
    Protected DBDataSource1 As SAPbouiCOM.DBDataSource
    Protected UserDataSource0 As SAPbouiCOM.UserDataSource
    Protected DataTable0 As SAPbouiCOM.DataTable
    Protected DataTable1 As SAPbouiCOM.DataTable

    


    Public Overridable Sub InitBase()

        EnableToolBarButtons()
    End Sub

    Public Sub New()

        AddDataSource()
        EnableToolBarButtons()
        'checkpermissions()
    End Sub

   
    
   
    

End Class

Public MustInherit Class SboFormBase
    Inherits UserFormBase

    Private Shared m_intFormCount As Integer
    Protected m_ParentAddon As SBOAddOn
    Protected m_FormUID As String
    Protected m_SboFormType As enSboFormTypes
    Protected m_SboForm As SAPbouiCOM.Form
    Protected m_Formtype As enSAPFormTypes
    Protected sboCompany As SAPbobsCOM.Company
    Protected sboApp As SAPbouiCOM.Application
    Protected fid As String, m_Active As Boolean = True
    Protected aSources As New ArrayList, aMatSources As New ArrayList
    Protected Conditions0 As SAPbouiCOM.Conditions
    Protected Condition0 As SAPbouiCOM.Condition
    Protected m_SBO_LoadFormType As enSBO_LoadFormTypes
    Public mst_FormUIDSon As String, mstOryxFormType As String
    Protected m_parentForm As SboForm
    Protected errCode As Integer, errMsg As String
    Protected ocomboCol As SAPbouiCOM.ComboBoxColumn, oEditCol As SAPbouiCOM.EditTextColumn
    Protected oCFLEvento As SAPbouiCOM.IChooseFromListEvent, oCFL As SAPbouiCOM.ChooseFromList, ochooseTable As SAPbouiCOM.DataTable
    Protected dsPerm As SboAddOnBase.Permissions



#Region "Properties"

    
    Public ReadOnly Property SBOFormType() As enSboFormTypes
        Get
            Return m_SboFormType
        End Get
    End Property
    Public Shared ReadOnly Property Count() As Integer
        Get
            Return m_intFormCount
        End Get
    End Property

    Public ReadOnly Property sboForm() As SAPbouiCOM.Form
        Get
            Return m_SboForm
        End Get
    End Property

    Protected ReadOnly Property SBO_Application() As SAPbouiCOM.Application
        Get
            Return m_ParentAddon.SboApplication
        End Get
    End Property

    Protected ReadOnly Property SBO_Company() As SAPbobsCOM.Company
        Get
            Return m_ParentAddon.SboCompany
        End Get
    End Property

    ReadOnly Property UniqueID() As String
        Get
            Return m_FormUID
        End Get
    End Property
    Public ReadOnly Property IsFormActive() As Boolean
        Get
            Return m_Active
        End Get
    End Property


    Public ReadOnly Property SBO_LoadFormType() As enSBO_LoadFormTypes
        Get
            Return m_SBO_LoadFormType
        End Get
    End Property


    Public ReadOnly Property DataTable(ByVal pst_DataTableUID As String) As SAPbouiCOM.DataTable
        Get
            Return m_SboForm.DataSources.DataTables.Item(pst_DataTableUID)
        End Get
    End Property
    Public ReadOnly Property UDS(ByVal pst_UDS_UID As String) As SAPbouiCOM.UserDataSource
        Get
            Return m_SboForm.DataSources.UserDataSources.Item(pst_UDS_UID)
        End Get
    End Property
    Public ReadOnly Property pCFL(ByVal pVal As SAPbouiCOM.ItemEvent) As SAPbouiCOM.ChooseFromListEvent
        Get
            Return CType(pVal.SAPpVal, SAPbouiCOM.ChooseFromListEvent)
        End Get
    End Property

    Public ReadOnly Property CFL(ByVal pst_UDS_UID As String) As SAPbouiCOM.ChooseFromList
        Get
            Return m_SboForm.ChooseFromLists.Item(pst_UDS_UID)
        End Get
    End Property
    Public ReadOnly Property DBDS(ByVal pbo_Object As Object) As SAPbouiCOM.DBDataSource
        Get
            If CStr(pbo_Object).ToUpper Like "SBO*".ToUpper Then
                pbo_Object = "@" & CStr(pbo_Object)
            End If
            Return m_SboForm.DataSources.DBDataSources.Item(pbo_Object)
        End Get
    End Property
    Public ReadOnly Property USERDS(ByVal pbo_Object As Object) As SAPbouiCOM.UserDataSource
        Get
            Return m_SboForm.DataSources.UserDataSources.Item(pbo_Object)
        End Get
    End Property
    Public ReadOnly Property FormItem(ByVal pst_ItemUID As String) As SAPbouiCOM.Item
        Get
            Return m_SboForm.Items.Item(pst_ItemUID)
        End Get
    End Property
    Public ReadOnly Property FormActiveX(ByVal pst_ItemUID As String) As SAPbouiCOM.ActiveX
        Get
            Return CType(m_SboForm.Items.Item(pst_ItemUID).Specific, SAPbouiCOM.ActiveX)
        End Get
    End Property
    Public ReadOnly Property FormCheckBox(ByVal pst_ItemUID As String) As SAPbouiCOM.CheckBox
        Get
            Return CType(m_SboForm.Items.Item(pst_ItemUID).Specific, SAPbouiCOM.CheckBox)
        End Get
    End Property
    Public ReadOnly Property FormColumn(ByVal pst_MatrixUID As String, ByVal pst_ColumnUID As String) As SAPbouiCOM.Column
        Get
            Return CType(m_SboForm.Items.Item(pst_MatrixUID).Specific, SAPbouiCOM.Matrix).Columns.Item(pst_ColumnUID)
        End Get
    End Property
    Public ReadOnly Property FormCellEdit(ByVal pst_MatrixUID As String, ByVal pst_ColumnUID As String, ByVal pin_Row As Int32) As SAPbouiCOM.EditText
        Get
            Return CType(m_SboForm.Items.Item(pst_MatrixUID).Specific.Columns.Item(pst_ColumnUID).Cells.Item(pin_Row).Specific, SAPbouiCOM.EditText)
        End Get
    End Property
    Public ReadOnly Property FormButton(ByVal pst_ItemUID As String) As SAPbouiCOM.Button
        Get
            Return CType(m_SboForm.Items.Item(pst_ItemUID).Specific, SAPbouiCOM.Button)
        End Get
    End Property

    Public ReadOnly Property FormPicture(ByVal pst_ItemUID As String) As SAPbouiCOM.PictureBox
        Get
            Return CType(m_SboForm.Items.Item(pst_ItemUID).Specific, SAPbouiCOM.PictureBox)
        End Get
    End Property
    Public ReadOnly Property FormOptionBtn(ByVal pst_ItemUID As String) As SAPbouiCOM.OptionBtn
        Get
            Return CType(m_SboForm.Items.Item(pst_ItemUID).Specific, SAPbouiCOM.OptionBtn)
        End Get
    End Property
    Public ReadOnly Property FormCombo(ByVal pst_ItemUID As String) As SAPbouiCOM.ComboBox
        Get
            Return CType(m_SboForm.Items.Item(pst_ItemUID).Specific, SAPbouiCOM.ComboBox)
        End Get
    End Property
    Public ReadOnly Property FormMatrix(ByVal pst_MatrixUID As String) As SAPbouiCOM.Matrix
        Get
            Return CType(m_SboForm.Items.Item(pst_MatrixUID).Specific, SAPbouiCOM.Matrix)
        End Get
    End Property
    Public ReadOnly Property FormTab(ByVal pst_tabUID As String) As SAPbouiCOM.Folder
        Get
            Return CType(m_SboForm.Items.Item(pst_tabUID).Specific, SAPbouiCOM.Folder)
        End Get
    End Property
    Public ReadOnly Property FormLabel(ByVal pst_ItemUID As String) As SAPbouiCOM.StaticText
        Get
            Return CType(m_SboForm.Items.Item(pst_ItemUID).Specific, SAPbouiCOM.StaticText)
        End Get
    End Property
    Public ReadOnly Property FormPaneComboBox(ByVal pst_ItemUID As String) As SAPbouiCOM.PaneComboBox
        Get
            Return CType(m_SboForm.Items.Item(pst_ItemUID).Specific, SAPbouiCOM.PaneComboBox)
        End Get
    End Property
    Public ReadOnly Property FormText(ByVal pst_ItemUID As String) As SAPbouiCOM.EditText
        Get
            Return CType(m_SboForm.Items.Item(pst_ItemUID).Specific, SAPbouiCOM.EditText)
        End Get
    End Property
    Public ReadOnly Property FormLink(ByVal pst_ItemUID As String) As SAPbouiCOM.LinkedButton
        Get
            Return CType(m_SboForm.Items.Item(pst_ItemUID).Specific, SAPbouiCOM.LinkedButton)
        End Get
    End Property
    Public ReadOnly Property FormGrid(ByVal pst_ItemUID As String) As SAPbouiCOM.Grid
        Get
            Return CType(m_SboForm.Items.Item(pst_ItemUID).Specific, SAPbouiCOM.Grid)
        End Get
    End Property
    Public ReadOnly Property FormCellGridValue(ByVal pst_GridUID As String, ByVal pst_ColumnUID As String, ByVal pin_row As Integer) As Object
        Get
            Return m_SboForm.Items.Item(pst_GridUID).Specific.DataTable.GetValue(pst_ColumnUID, m_SboForm.Items.Item(pst_GridUID).Specific.GetDataTableRowIndex(pin_row))
        End Get
    End Property


    Public ReadOnly Property NextNum(ByVal pst_TableName As String) As String
        Get
            Try
                Dim m_sql As String = ""
                m_sql = "SELECT ISNULL(MAX(CAST(code as int)),0) + 1 NextNum"
                m_sql += " FROM [" + pst_TableName.Trim + "]"
                Me.DataTable("DT_Seq").ExecuteQuery(m_sql)
                Return Me.DataTable("DT_Seq").GetValue(0, 0)
            Catch ex As Exception
                Me.m_ParentAddon.SboApplication.MessageBox("Table - " + pst_TableName + " does not exist ")
                Return Nothing
            End Try
        End Get

    End Property
    Public ReadOnly Property NextNum(ByVal pst_TableName As String, ByVal pst_FieldName As String) As String
        Get
            Try
                Dim m_sql As String = ""
                m_sql = "SELECT ISNULL(MAX(CAST(" + pst_FieldName + " as int)),0) + 1 NextNum"
                m_sql += " FROM [" + pst_TableName.Trim + "]"
                Me.DataTable("DT_Seq").ExecuteQuery(m_sql)
                Return Me.DataTable("DT_Seq").GetValue(0, 0)
            Catch ex As Exception
                Me.m_ParentAddon.SboApplication.MessageBox("Table - " + pst_TableName + " does not exist ")
                Return Nothing
            End Try
        End Get

    End Property



#End Region
#Region " Overrides "


    Public Function checkpermissions(ByVal PermissionID As String) As Integer
        Dim perm As SAPbobsCOM.BoPermission
        Try
            Dim sboBOB As SAPbobsCOM.SBObob
            Dim sboRecordset As SAPbobsCOM.Recordset
            sboBOB = sboCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge)

            sboRecordset = sboBOB.GetSystemPermission(sboApp.Company.UserName, PermissionID)
            perm = sboRecordset.Fields.Item(0).Value

        Catch ex As Exception
            errMsg = ex.ToString
            Me.SBO_Application.StatusBar.SetText(errMsg, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
        checkpermissions = perm
    End Function

    Protected Overridable Sub createPermissionDataset()

    End Sub



#End Region
    Public Overrides Function ToString() As String
        Dim aName() As String
        aName = Split(MyBase.ToString(), ".")
        Return aName(aName.Length - 1)
    End Function

    Protected Overrides Sub Finalize()
        'clean up class
        m_ParentAddon = Nothing
        m_SboForm = Nothing
        GC.WaitForPendingFinalizers()
        GC.Collect()
        MyBase.Finalize()
    End Sub
#Region " Functions "
    Protected Overridable Sub AddDataSource()


    End Sub

    Protected Overridable Sub OkAction()

    End Sub

    Protected Overridable Sub NewRecord()

    End Sub

    Protected Overridable Sub BindDataToForm()

    End Sub

    Protected Overridable Sub GetDataFromDataSource()

    End Sub

    Protected Overridable Sub EnableToolBarButtons()

    End Sub

    Protected Function GenerateRandomString(ByVal intLenghtOfString As Integer) As String
        'Create a new StrinBuilder that would hold the random string.
        Dim randomString As StringBuilder = New StringBuilder
        'Create a new instance of the class Random
        Dim randomNumber As Random = New Random
        'Create a variable to hold the generated charater.
        Dim appendedChar As Char
        'Create a loop that would iterate from 0 to the specified value of intLenghtOfString
        For i As Integer = 0 To intLenghtOfString
            'Generate the char and assign it to appendedChar
            appendedChar = Convert.ToChar(Convert.ToInt32(26 * randomNumber.NextDouble()) + 65)
            'Append appendedChar to randomString
            randomString.Append(appendedChar)
        Next
        'Convert randomString to String and return the result.
        Return randomString.ToString()
    End Function

    Public Function FormSonOpen(ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) As Boolean
        Try
            If Not IsNothing(mst_FormUIDSon) Then 'if a son form is open                
                Dim dsa_Form As SAPbouiCOM.Form
                For Each dsa_Form In SBO_Application.Forms
                    'look in application forms collection to check if it's still open
                    If dsa_Form.UniqueID = mst_FormUIDSon Then
                        'Son form found, select it
                        SBO_Application.Forms.Item(mst_FormUIDSon).Select()
                        dsa_Form = Nothing
                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_VALIDATE _
                        Or pVal.EventType = SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST _
                            Or pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE _
                            Or pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD Then
                            'Some event must be allowed even if son form is open
                            '(like transfert data from son form to father form => et_VALIDATE is needed)
                            Return True
                        Else
                            BubbleEvent = False
                            Return True
                        End If
                    End If
                Next
                dsa_Form = Nothing
                'Form not found, so it means it has been closed                
                mst_FormUIDSon = Nothing
                Return False
            End If
            Return False

        Catch ex As Exception
            BubbleEvent = False
            Return True
        End Try
    End Function

    Protected Function GetSQLString(ByVal strProc As String, ByVal ParamArray Parameters() As String) As String
        Try
            Dim Ret As String = Nothing
            Dim strMsg As String = ""
            Dim file As System.IO.Stream = m_ParentAddon.getAppResource(strProc + ".txt")
            Dim strSQL As String = ""
            If file Is Nothing Then
                strMsg = "SQL File - " + strProc + " does not exist"
                SBO_Application.SetStatusBarMessage(strMsg, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                Return ""
            End If
            Dim SReader As New System.IO.StreamReader(file)

            strSQL = SReader.ReadToEnd

            Dim i As Integer, strHolder As String = "OWAPARAM", length As Integer
            length = Parameters.Length
            If length > 0 Then
                For i = 0 To length - 1
                    strSQL = strSQL.Replace(strHolder + (i + 1).ToString, Parameters(i))
                Next
            End If
            Ret = strSQL
            Return Ret
        Catch ex As Exception
            SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
            Return Nothing

        End Try
    End Function

#End Region

#Region " Miscellaneous "
    Protected Sub fillCombo(ByVal strKey As String, ByVal strDesc As String, _
        ByVal strTable As String, ByVal oCombo As SAPbouiCOM.ComboBox, Optional ByVal strWhere As String = "", _
        Optional ByVal bAddNew As Boolean = False, Optional ByVal fieldType As SAPbobsCOM.BoFieldTypes = SAPbobsCOM.BoFieldTypes.db_Alpha, _
        Optional ByVal AddTo As Boolean = False)
        Dim strSQL As String
        Try
            Dim i As Integer
            If Not AddTo Then
                Try
                    For i = 0 To oCombo.ValidValues.Count - 1
                        oCombo.ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index)
                    Next
                Catch ex As Exception

                End Try

            End If

            strSQL = " SELECT " + strKey + "," + strDesc + " FROM [" + strTable + "]"
            If strWhere <> String.Empty Then strSQL += " WHERE " + strWhere
            strSQL += " ORDER BY " + strKey + " DESC "
            DataTable("DT_Base").Clear()
            DataTable("DT_Base").ExecuteQuery(strSQL)

            For i = 0 To DataTable("DT_Base").Rows.Count - 1
                fieldType = DataTable("DT_Base").Columns.Item(strKey).Type
                oCombo.ValidValues.Add(DataTable("DT_Base").GetValue(strKey, i), DataTable("DT_Base").GetValue(strDesc, i))
            Next

            If bAddNew Then
                oCombo.ValidValues.Add("Add New", "Add New")
            End If


        Catch ex As Exception
            'Me.m_ParentAddon.WriteLog(ex.ToString)
            Me.m_ParentAddon.SboApplication.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
        End Try
    End Sub

    Protected Sub fillCombo(ByVal strKey As String, ByVal strDesc As String, _
    ByVal strTable As String, ByVal oValues As SAPbouiCOM.ValidValues, Optional ByVal strWhere As String = "", _
    Optional ByVal bAddNew As Boolean = False, Optional ByVal fieldType As SAPbobsCOM.BoFieldTypes = SAPbobsCOM.BoFieldTypes.db_Alpha)
        Dim strSQL As String
        Try
            Dim i As Integer
            For i = 0 To oValues.Count - 1
                oValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index)
            Next
            strSQL = " SELECT " + strKey + "," + strDesc + " FROM [" + strTable + "]"
            If strWhere <> String.Empty Then strSQL += " WHERE " + strWhere
            DataTable("DT_Base").Clear()
            DataTable("DT_Base").ExecuteQuery(strSQL)

            For i = 0 To DataTable("DT_Base").Rows.Count - 1
                fieldType = DataTable("DT_Base").Columns.Item(strKey).Type
                If String.IsNullOrEmpty(DataTable("DT_Base").GetValue(strKey, i)) Then
                    Continue For
                End If
                oValues.Add(DataTable("DT_Base").GetValue(strKey, i), DataTable("DT_Base").GetValue(strDesc, i))
            Next

            oValues.Add("", "")

        Catch ex As Exception
            'Me.m_ParentAddon.WriteLog(ex.ToString)
            Me.m_ParentAddon.SboApplication.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
        End Try
    End Sub

    Protected Sub fillCombo(ByVal strKey As String, ByVal strDesc As String, _
    ByVal strProc As String, ByVal strParam As String, ByVal oCombo As SAPbouiCOM.ComboBox)

        Dim strSQL As String
        Try
            Dim i As Integer
            For i = 0 To oCombo.ValidValues.Count - 1
                oCombo.ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index)
            Next
            strSQL = " EXEC " + strProc + " " + strParam
            DataTable("DT_Base").Clear()
            DataTable("DT_Base").ExecuteQuery(strSQL)

            For i = 0 To DataTable("DT_Base").Rows.Count - 1
                If String.IsNullOrEmpty(DataTable("DT_Base").GetValue(strKey, i)) Then
                    Continue For
                End If
                oCombo.ValidValues.Add(DataTable("DT_Base").GetValue(strKey, i), DataTable("DT_Base").GetValue(strDesc, i))
            Next

            oCombo.ValidValues.Add("", "")

        Catch ex As Exception
            'Me.m_ParentAddon.WriteLog(ex.ToString)
            Me.m_ParentAddon.SboApplication.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
        End Try
    End Sub

    Protected Sub clearCombo(ByVal oCombo As SAPbouiCOM.ComboBox, Optional ByVal lAddNew As Boolean = True)
        Dim i As Integer
        For i = 0 To oCombo.ValidValues.Count - 1
            oCombo.ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index)
        Next
        If lAddNew Then
            oCombo.ValidValues.Add("Add New", "")
        Else
            oCombo.ValidValues.Add("", "")
        End If
        oCombo.Select(oCombo.ValidValues.Count - 1, SAPbouiCOM.BoSearchKey.psk_Index)
    End Sub

    Protected Overridable Function validatealias(ByVal condVal As String, ByVal AliasName As String, _
         ByVal oDataSource As SAPbouiCOM.DBDataSource) As Boolean
        Conditions0 = New SAPbouiCOM.Conditions
        Condition0 = Conditions0.Add
        Condition0.Alias = AliasName
        Condition0.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
        Condition0.CondVal = condVal
        If oDataSource.Size > 0 Then oDataSource.Clear()
        oDataSource.Query(Conditions0)
        If oDataSource.Size = 0 Then
            validatealias = False
        Else
            validatealias = True
        End If
        Conditions0 = Nothing
        GC.Collect()
    End Function

    Protected Overridable Function validatealias(ByVal condVal As String, ByVal AliasName As String, _
         ByVal oDataSource As SAPbouiCOM.DBDataSource, ByVal condVal1 As String, ByVal aliasName1 As String) As Boolean
        Conditions0 = New SAPbouiCOM.Conditions
        Condition0 = Conditions0.Add
        Condition0.Alias = AliasName
        Condition0.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
        Condition0.CondVal = condVal
        Condition0.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
        Condition0 = Conditions0.Add
        Condition0.Alias = aliasName1
        Condition0.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
        Condition0.CondVal = condVal1

        oDataSource.Clear()
        oDataSource.Query(Conditions0)
        If oDataSource.Size = 0 Then
            validatealias = False
        Else
            validatealias = True
        End If
        Conditions0 = Nothing
        GC.Collect()
    End Function

    Protected Function getOffset(ByVal condVal As String, ByVal AliasName As String, _
         ByVal oDataSource As SAPbouiCOM.DBDataSource) As Boolean

        Conditions0 = New SAPbouiCOM.Conditions
        Condition0 = Conditions0.Add
        Condition0.Alias = AliasName
        Condition0.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
        Condition0.CondVal = condVal
        oDataSource.Clear()
        oDataSource.Query(Conditions0)
        If oDataSource.Size = 0 Then
            getOffset = False
        Else
            getOffset = True
        End If
        Conditions0 = Nothing
        GC.Collect()
    End Function

    Protected Function getOffset(ByVal condVal As String, ByVal AliasName As String, _
        ByVal oDataSource As SAPbouiCOM.DBDataSource, ByVal operation As SAPbouiCOM.BoConditionOperation) As Boolean

        Conditions0 = New SAPbouiCOM.Conditions
        Condition0 = Conditions0.Add
        Condition0.Alias = AliasName
        Condition0.Operation = operation
        Condition0.CondVal = condVal
        oDataSource.Clear()
        oDataSource.Query(Conditions0)
        If oDataSource.Size = 0 Then
            getOffset = False
        Else
            getOffset = True
        End If
        Conditions0 = Nothing
        GC.Collect()
    End Function

    Protected Function getOffset(ByVal condVal As String, ByVal AliasName As String, ByVal condVal1 As String, ByVal AliasName1 As String, _
            ByVal oDataSource As SAPbouiCOM.DBDataSource, ByVal operation As SAPbouiCOM.BoConditionOperation) As Boolean

        Conditions0 = New SAPbouiCOM.Conditions
        Condition0 = Conditions0.Add
        Condition0.BracketOpenNum = 2
        Condition0.Alias = AliasName
        Condition0.Operation = operation
        Condition0.CondVal = condVal.Trim
        Condition0.BracketCloseNum = 1

        Condition0 = Conditions0.Add
        Condition0.BracketOpenNum = 1
        Condition0.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
        Condition0.Alias = AliasName1
        Condition0.Operation = operation
        Condition0.CondVal = condVal1.Trim
        Condition0.BracketCloseNum = 2

        oDataSource.Clear()
        oDataSource.Query(Conditions0)
        If oDataSource.Size = 0 Then
            getOffset = False
        Else
            getOffset = True
        End If
        Conditions0 = Nothing
        GC.Collect()
    End Function

    Public Shared Function sboDate(ByVal thedate As Date) As String
        sboDate = thedate.Year.ToString + _
                            thedate.Month.ToString.PadLeft(2, "0") + _
                            thedate.Day.ToString.PadLeft(2, "0")
    End Function

    Public Shared Function sboDate(ByVal thedate As String) As Date
        Dim strDate As String
        strDate = thedate.Substring(6, 2) + "." + _
            thedate.Substring(4, 2) + "." + thedate.Substring(0, 4)
        sboDate = DateValue(strDate)
    End Function

    Public Overridable Sub parentFormRefresh(ByVal strID As String)

    End Sub

    Protected Sub chooseFromListSystem(ByVal strObjType As String, ByVal multiselection As Boolean, _
                        ByVal strID As String, ByVal conAlias As String, ByVal strConVal As String)

        Try
            Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition

            oCFLs = m_SboForm.ChooseFromLists

            Dim oCFL As SAPbouiCOM.ChooseFromList
            Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams

            oCFLCreationParams = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)

            ' Adding 2 CFL, one for the button and one for the edit text.
            oCFLCreationParams.MultiSelection = multiselection
            oCFLCreationParams.ObjectType = strObjType
            oCFLCreationParams.UniqueID = strID

            oCFL = oCFLs.Add(oCFLCreationParams)
            If conAlias <> "" Then
                ' Adding Conditions to CFL1
                oCons = oCFL.GetConditions()

                oCon = oCons.Add()
                oCon.Alias = conAlias
                oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCon.CondVal = strConVal

                oCFL.SetConditions(oCons)

            End If


        Catch ex As Exception
            'Me.m_ParentAddon.WriteLog(ex.ToString)
            Me.m_ParentAddon.SboApplication.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
        End Try
    End Sub

    Protected Sub chooseFromListSystem(ByVal strObjType As String, ByVal multiselection As Boolean, _
                        ByVal strID As String, ByVal conAlias As String, ByVal strConVal As String, _
                        ByVal conAlias1 As String, ByVal strConVal1 As String)

        Try
            Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition

            oCFLs = m_SboForm.ChooseFromLists

            Dim oCFL As SAPbouiCOM.ChooseFromList
            Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams

            oCFLCreationParams = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)

            ' Adding 2 CFL, one for the button and one for the edit text.
            oCFLCreationParams.MultiSelection = multiselection
            oCFLCreationParams.ObjectType = strObjType
            oCFLCreationParams.UniqueID = strID

            oCFL = oCFLs.Add(oCFLCreationParams)
            If conAlias <> "" Then
                ' Adding Conditions to CFL1
                oCons = oCFL.GetConditions()

                oCon = oCons.Add()
                oCon.Alias = conAlias
                oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCon.CondVal = strConVal
                oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
                oCon = oCons.Add
                oCon.Alias = conAlias1
                oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCon.CondVal = strConVal1

                oCFL.SetConditions(oCons)

            End If


        Catch ex As Exception
            'Me.m_ParentAddon.WriteLog(ex.ToString)
            Me.m_ParentAddon.SboApplication.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
        End Try
    End Sub

    Protected Function HandleChooseFromListEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) As String
        Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
        oCFLEvento = pVal
        Dim sCFL_ID As String
        sCFL_ID = oCFLEvento.ChooseFromListUID
        Dim oCFL As SAPbouiCOM.ChooseFromList
        oCFL = m_SboForm.ChooseFromLists.Item(sCFL_ID)
        Dim val As String = "", oDatatable As SAPbouiCOM.DataTable
        If oCFLEvento.BeforeAction = False Then
            oDatatable = oCFLEvento.SelectedObjects
            If Not oDatatable Is Nothing Then
                Try
                    val = oDatatable.GetValue(0, 0)
                Catch ex As Exception
                    'Me.m_ParentAddon.WriteLog(ex.ToString)
                    Me.m_ParentAddon.SboApplication.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                End Try
            End If

        End If
        Return val
    End Function

    Protected Function HandleChooseFromListEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean, ByRef lData As Boolean) As SAPbouiCOM.DataTable
        Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
        oCFLEvento = pVal
        Dim sCFL_ID As String
        sCFL_ID = oCFLEvento.ChooseFromListUID
        Dim oCFL As SAPbouiCOM.ChooseFromList
        oCFL = m_SboForm.ChooseFromLists.Item(sCFL_ID)
        Dim val As String = ""
        lData = True
        If oCFLEvento.BeforeAction = False Then
            If IsNothing(oCFLEvento.SelectedObjects) Then lData = False
            Return oCFLEvento.SelectedObjects
        End If
        Return Nothing
    End Function

    Protected Sub setconditions(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, _
                        ByVal conAlias As String, ByVal strConVal As String, _
                        ByVal conAlias1 As String, ByVal strConVal1 As String)
        Dim oCons As SAPbouiCOM.Conditions
        Dim oCon As SAPbouiCOM.Condition
        Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
        oCFLEvento = pVal
        Dim sCFL_ID As String
        sCFL_ID = oCFLEvento.ChooseFromListUID
        Dim oCFL As SAPbouiCOM.ChooseFromList
        oCFL = m_SboForm.ChooseFromLists.Item(sCFL_ID)

        If conAlias <> "" Then
            ' Adding Conditions to CFL1
            Dim emptyCon As SAPbouiCOM.Conditions
            emptyCon = New SAPbouiCOM.ConditionsClass()
            oCFL.SetConditions(emptyCon)

            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = conAlias
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = strConVal.Trim
            oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
            oCon = oCons.Add
            oCon.Alias = conAlias1
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = strConVal1.Trim

            oCFL.SetConditions(oCons)
        End If
    End Sub

    Protected Sub setconditions(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, _
    ByVal conString As ArrayList)
        Dim oCons As SAPbouiCOM.Conditions
        Dim oCon As SAPbouiCOM.Condition
        Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
        oCFLEvento = pVal
        Dim sCFL_ID As String
        sCFL_ID = oCFLEvento.ChooseFromListUID
        Dim oCFL As SAPbouiCOM.ChooseFromList
        oCFL = m_SboForm.ChooseFromLists.Item(sCFL_ID)


        If conString.Count > 0 Then
            ' Adding Conditions to CFL1
            Dim emptyCon As SAPbouiCOM.Conditions
            emptyCon = New SAPbouiCOM.ConditionsClass()
            oCFL.SetConditions(emptyCon)
            Dim i As Integer
            Dim cons As ConditionVals = Nothing
            oCons = oCFL.GetConditions()

            For i = 0 To conString.Count - 1
                oCon = oCons.Add()
                cons = CType(conString(i), ConditionVals)
                oCon.Alias = cons.cAlias
                oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCon.CondVal = cons.cValue
                If i <> conString.Count - 1 Then
                    oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
                End If
            Next
            oCFL.SetConditions(oCons)
        End If
    End Sub

#End Region

End Class

