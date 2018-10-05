Imports System.Threading
Public MustInherit Class SBOBaseObject : Implements IDisposable
    Protected m_ParentAddon As SboAddon
    Protected m_Form As SAPbouiCOM.IForm
    Protected m_DBDataSource0 As SAPbouiCOM.DBDataSource
    Protected m_DBDataSource1 As SAPbouiCOM.DBDataSource
    Protected m_DBDataSource2 As SAPbouiCOM.DBDataSource
    Protected m_DBDataSource3 As SAPbouiCOM.DBDataSource
    Protected m_DBDataSource4 As SAPbouiCOM.DBDataSource
    Protected m_DataTable0 As SAPbouiCOM.DataTable
    Protected m_DataTable1 As SAPbouiCOM.DataTable
    Protected m_DataTable2 As SAPbouiCOM.DataTable
    Protected m_DataTable3 As SAPbouiCOM.DataTable
    Protected m_DataTable4 As SAPbouiCOM.DataTable
    Protected m_DataTable5 As SAPbouiCOM.DataTable
    Protected m_DataTable6 As SAPbouiCOM.DataTable

    Protected m_Conditions As SAPbouiCOM.Conditions, m_Conditions1 As SAPbouiCOM.Conditions, m_Conditions2 As SAPbouiCOM.Conditions
    Protected m_Condition As SAPbouiCOM.Condition, m_Condition1 As SAPbouiCOM.Condition, m_Condition2 As SAPbouiCOM.Condition
    Protected WithEvents m_SboApplication As SAPbouiCOM.Application
    Private m_SBOSQL As SBOSQLBase
    Protected m_CurrentLineNo As Integer
    Protected m_SonForm As UserFormBaseClass
    Protected m_id As Integer

    Private m_SonFormName As String

    ' Flag: Has Dispose already been called? 
    Protected disposed As Boolean = False

    ' Public implementation of Dispose pattern callable by consumers. 
    Public Sub Dispose() _
               Implements IDisposable.Dispose
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub

    ' Protected implementation of Dispose pattern. 
    Protected Overridable Sub Dispose(ByVal disposing As Boolean)
        If disposed Then Return

        If disposing Then
            ' Free any other managed objects here. 
            ' 
        End If

        ' Free any unmanaged objects here. 
        m_Form = Nothing
        '
        disposed = True
    End Sub

    Protected Property SonFormName As String
        Get
            Return m_SonFormName
        End Get
        Set(ByVal value As String)
            m_SonFormName = value
        End Set
    End Property

    Public Overridable Sub OnCustomInit()
        OnFormInit()
        Me.AddDataSource()
        Me.SetConditions()
        Me.OnFormLoad()
        Me.FormatGrid()
        Dim Rand As New Random
        Me.m_id = Rand.Next()

    End Sub

    Protected Overridable Sub OnFormInit()
        m_Form.DataSources.DataTables.Add("DT_SeqObj")
        m_Form.DataSources.DataTables.Add("DT_BaseObj")

    End Sub

    Protected Overridable Sub FormatGrid()

    End Sub
    Protected Overridable Function GetServerSQL(ByVal strProc As String, ByVal ParamArray Parameters() As String) As String
        Dim Ret As String = Nothing
        Try
            Ret = SBOServerSQL.GetSQLString(strProc, Parameters)
            If SBOServerSQL.ErrorNo <> 0 Then
                m_ParentAddon.SboApplication.StatusBar.SetText(SBOServerSQL.ErrorMsg, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End If
        Catch ex As Exception
            m_ParentAddon.SboApplication.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
            Ret = Nothing
        End Try
        GetServerSQL = Ret
    End Function

    Protected ReadOnly Property SBOServerSQL As SBOSQLBase
        Get
            Return m_SBOSQL
        End Get

    End Property

    Protected Sub InitSBOServerSQL(ByVal value As SBOSQLBase)
        m_SBOSQL = value
    End Sub
    Protected Overridable Function ExecuteSQL(ByVal strProc As String, ByVal ParamArray Parameters() As String) As SAPbobsCOM.Recordset
        Dim SBO_RecSet As SAPbobsCOM.Recordset = Nothing
        Try
            SBO_RecSet = SBOServerSQL.ExecuteSQL(strProc, Parameters)
            'If SBOServerSQL.ErrorNo <> 0 Then
            '    m_ParentAddon.SboApplication.StatusBar.SetText(SBOServerSQL.ErrorMsg, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            'End If
        Catch ex As Exception

        End Try
        ExecuteSQL = SBO_RecSet
    End Function

    Protected Overridable Function ExecuteSQL(ByVal strSQL As String) As SAPbobsCOM.Recordset
        Dim SBO_RecSet As SAPbobsCOM.Recordset = Nothing
        Try
            SBO_RecSet = SBOServerSQL.ExecuteSQL(strSQL)
            'If SBOServerSQL.ErrorNo <> 0 Then
            '    m_ParentAddon.SboApplication.StatusBar.SetText(SBOServerSQL.ErrorMsg, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            'End If
        Catch ex As Exception

        End Try
        ExecuteSQL = SBO_RecSet
    End Function

    Public Sub MatrixAutoAddRow(ByVal pVal As SAPbouiCOM.SBOItemEventArg)
        Dim oMatrix As SAPbouiCOM.Matrix

        oMatrix = m_Form.Items.Item(pVal.ItemUID).Specific

        If pVal.Row = oMatrix.RowCount + 1 Then
            If pVal.Row = 1 Then
                oMatrix.AddRow(1)
            Else
                oMatrix.AddRow(1, oMatrix.RowCount)
            End If
            oMatrix.Columns.Item(1).Cells.Item(pVal.Row).Click()
        End If
    End Sub

    Protected Overridable Function ExecuteSQLUpdate(ByVal strSQL As String) As Boolean
        Dim SBO_RecSet As SAPbobsCOM.Recordset = Nothing
        Try
            SBO_RecSet = SBOServerSQL.ExecuteSQL(strSQL)
            ExecuteSQLUpdate = True
        Catch ex As Exception
            m_ParentAddon.SboApplication.StatusBar.SetText(ex.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            ExecuteSQLUpdate = False
        End Try
    End Function

    Protected Overridable Function ExecuteSQLDT(ByVal strProc As String, ByVal ParamArray Parameters() As String) As SAPbouiCOM.DataTable
        Dim SBO_RecSet As SAPbobsCOM.Recordset = Nothing
        Try
            Dim SQL As String = SBOServerSQL.GetSQLString(strProc, Parameters)
            m_Form.DataSources.DataTables.Item("DT_BaseObj").ExecuteQuery(SQL)
            ExecuteSQLDT = m_Form.DataSources.DataTables.Item("DT_BaseObj")
        Catch ex As Exception
            m_ParentAddon.SboApplication.StatusBar.SetText(ex.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            ExecuteSQLDT = Nothing
        End Try
    End Function

    Protected Overridable Function ExecuteSQLDT(ByVal strSQL As String) As SAPbouiCOM.DataTable
        Dim SBO_RecSet As SAPbobsCOM.Recordset = Nothing
        Try
            m_Form.DataSources.DataTables.Item("DT_BaseObj").ExecuteQuery(strSQL)
            ExecuteSQLDT = m_Form.DataSources.DataTables.Item("DT_BaseObj")
        Catch ex As Exception
            m_ParentAddon.SboApplication.StatusBar.SetText(ex.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            ExecuteSQLDT = Nothing
        End Try
    End Function

    Protected Overridable Sub ExecuteSQLDT(ByVal strProc As String, ByRef oDT As SAPbouiCOM.DataTable, ByVal ParamArray Parameters() As String)
        Dim SBO_RecSet As SAPbobsCOM.Recordset = Nothing
        'Try
        Dim SQL As String = SBOServerSQL.GetSQLString(strProc, Parameters)
        oDT.Clear()
        oDT.ExecuteQuery(SQL)
        'Catch ex As Exception
        '    m_ParentAddon.SboApplication.StatusBar.SetText(ex.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        '    Throw (ex)
        'End Try
    End Sub

    Protected Overridable Sub ExecuteSQLDT(ByVal strSQL As String, ByRef oDT As SAPbouiCOM.DataTable)
        Dim SBO_RecSet As SAPbobsCOM.Recordset = Nothing
        'Try
        Dim SQL As String = SBOServerSQL.GetSQLString(strSQL)
        oDT.Clear()
        oDT.ExecuteQuery(SQL)
        'Catch ex As Exception
        '    m_ParentAddon.SboApplication.StatusBar.SetText(ex.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        'End Try
    End Sub
    Protected Overridable Sub SetConditions()
        m_Conditions = New SAPbouiCOM.Conditions
        m_Conditions1 = New SAPbouiCOM.Conditions
        m_Conditions2 = New SAPbouiCOM.Conditions
        m_Condition = m_Conditions.Add
        m_Condition1 = m_Conditions1.Add
        m_Condition2 = m_Conditions2.Add
    End Sub

    Protected Overridable Sub OnFormLoad()
        OnFormNavigate()
        FormRefresh()
        EnableToolBarButtons()
    End Sub

    Protected Overridable Sub AddDataSource()

    End Sub

    Protected Overridable Sub OnFormNavigate()
        QueryDBInit()
    End Sub



    Public Overridable Sub OnDataLoadAfter(ByRef pVal As SAPbouiCOM.BusinessObjectInfo)

    End Sub

    Public Overridable Sub OnDataLoadBefore(ByRef pVal As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)

    End Sub

    Public Overridable Sub OnDataAddAfter(ByRef pVal As SAPbouiCOM.BusinessObjectInfo)

    End Sub

    Public Overridable Sub OnDataAddBefore(ByRef pVal As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Dim ErrNo As Integer, ErrMsg As String = ""
        If Not Save(ErrNo, ErrMsg) Then
            m_ParentAddon.SboApplication.StatusBar.SetText(ErrMsg, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            BubbleEvent = False
        End If
    End Sub

    Public Overridable Sub OnDataDeleteAfter(ByRef pVal As SAPbouiCOM.BusinessObjectInfo)

    End Sub

    Public Overridable Sub OnDataDeleteBefore(ByRef pVal As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        BubbleEvent = CanDelete()
    End Sub

    Public Overridable Sub OnDataUpdateAfter(ByRef pVal As SAPbouiCOM.BusinessObjectInfo)

    End Sub

    Public Overridable Sub OnDataUpdateBefore(ByRef pVal As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Dim ErrNo As Integer, ErrMsg As String = ""
        If Not Save(ErrNo, ErrMsg) Then
            m_ParentAddon.SboApplication.StatusBar.SetText(ErrMsg, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            BubbleEvent = False
        End If
    End Sub


    Public Overridable Sub OnLoadAfter(ByVal pVal As SAPbouiCOM.SBOItemEventArg)

    End Sub

    Public Overridable Sub OnLoadBefore(ByVal pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As Boolean)

    End Sub

    Protected Overridable Function CanDelete() As Boolean
        CanDelete = True
    End Function


    Sub New(ByVal pAddon As SboAddon, ByVal pForm As SAPbouiCOM.IForm)
        m_ParentAddon = pAddon
        m_Form = pForm
        m_SboApplication = m_ParentAddon.SboApplication
    End Sub

    Public Overridable Sub OnComponentInit()

    End Sub
    Public Overridable Sub OnLinkedPressedAfter(ByVal sboObject As Object, ByVal pVal As SAPbouiCOM.SBOItemEventArg)

    End Sub

    Public Overridable Sub OnLinkedPressedBefore(ByVal sboObject As Object, ByVal pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As Boolean)

    End Sub

    Public Overridable Sub OnChooseFromListAfter(ByVal sboObject As Object, ByVal pVal As SAPbouiCOM.SBOItemEventArg)

    End Sub

    Public Overridable Sub OnChooseFromListBefore(ByVal sboObject As Object, ByVal pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As Boolean)

    End Sub

    Public Overridable Sub OnItemClickAfter(ByVal sboObject As Object, ByVal pVal As SAPbouiCOM.SBOItemEventArg)

    End Sub

    Public Overridable Sub OnDoubleClickAfter(ByVal sboObject As Object, ByVal pVal As SAPbouiCOM.SBOItemEventArg)

    End Sub

    Public Overridable Sub OnDoubleClickBefore(ByVal sboObject As Object, ByVal pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As Boolean)

    End Sub
    Public Overridable Sub OnItemClickBefore(ByVal sboObject As Object, ByVal pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As Boolean)
        Dim ErrNo As Integer, ErrMsg As String = ""
        Try
            If pVal.ItemUID = "1" Then
                If m_Form.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or m_Form.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                    If Not Save(ErrNo, ErrMsg) Then
                        m_ParentAddon.SboApplication.StatusBar.SetText(ErrMsg, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        BubbleEvent = False
                    End If
                End If
            End If
        Catch ex As Exception

        End Try
    End Sub

    Public Overridable Sub OnItemDoubleClickAfter(ByVal sboObject As Object, ByVal pVal As SAPbouiCOM.SBOItemEventArg)

    End Sub

    Public Overridable Sub OnItemDoubleClickBefore(ByVal sboObject As Object, ByVal pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As Boolean)

    End Sub

    Public Overridable Sub OnItemGotFocusAfter(ByVal sboObject As Object, ByVal pVal As SAPbouiCOM.SBOItemEventArg)

    End Sub

    Public Overridable Sub OnItemKeyDownAfter(ByVal sboObject As Object, ByVal pVal As SAPbouiCOM.SBOItemEventArg)

    End Sub

    Public Overridable Sub OnItemKeyDownBefore(ByVal sboObject As Object, ByVal pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As Boolean)

    End Sub

    Public Overridable Sub OnItemLostFocusAfter(ByVal sboObject As Object, ByVal pVal As SAPbouiCOM.SBOItemEventArg)

    End Sub

    Public Overridable Sub OnItemPickerClickedBefore(ByVal sboObject As Object, ByVal pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As Boolean)

    End Sub

    Public Overridable Sub OnItemPressedAfter(ByVal sboObject As Object, ByVal pVal As SAPbouiCOM.SBOItemEventArg)

    End Sub

    Public Overridable Sub OnItemPressedBefore(ByVal sboObject As Object, ByVal pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As Boolean)

    End Sub

    Public Overridable Sub OnComboSelectAfter(ByVal sboObject As Object, ByVal pVal As SAPbouiCOM.SBOItemEventArg)

    End Sub

    Public Overridable Sub OnComboSelectBefore(ByVal sboObject As Object, ByVal pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As Boolean)

    End Sub

    Public Overridable Sub OnItemValidateAfter(ByVal sboObject As Object, ByVal pVal As SAPbouiCOM.SBOItemEventArg)
    End Sub

    Public Overridable Sub OnItemValidateBefore(ByVal sboObject As Object, ByVal pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As Boolean)

    End Sub

    Protected Overridable Function Save(ByRef pErrNo As Integer, ByRef pErrMsg As String) As Boolean

        Save = IsReady(pErrNo, pErrMsg)
    End Function

    Protected Overridable Function IsReady(ByRef pErrNo As Integer, ByRef pErrMsg As String) As Boolean
        IsReady = True
    End Function

    Protected Function HandleChooseFromListEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As Boolean) As String
        Dim oCFLEvento As SAPbouiCOM.ISBOChooseFromListEventArg

        oCFLEvento = DirectCast(pVal, SAPbouiCOM.ISBOChooseFromListEventArg)

        Dim val As String = "", oDatatable As SAPbouiCOM.DataTable

        oDatatable = oCFLEvento.SelectedObjects

        If Not oDatatable Is Nothing Then
            Try
                val = oDatatable.GetValue(0, 0)
            Catch ex As Exception

                Me.m_ParentAddon.SboApplication.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
            End Try
        End If

        Return val
    End Function

    Protected Function HandleChooseFromListEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As Boolean, ByRef lData As Boolean) As SAPbouiCOM.DataTable
        Dim oCFLEvento As SAPbouiCOM.ISBOChooseFromListEventArg

        oCFLEvento = DirectCast(pVal, SAPbouiCOM.ISBOChooseFromListEventArg)

        Dim oDatatable As SAPbouiCOM.DataTable

        oDatatable = oCFLEvento.SelectedObjects

        If Not oDatatable Is Nothing Then
            Try
                lData = True
            Catch ex As Exception
                Me.m_ParentAddon.SboApplication.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
            End Try
        End If

        Return oDatatable

        'Dim oCFLEvento As SAPbouiCOM.ISBOChooseFromListEventArg
        'oCFLEvento = DirectCast(pVal, SAPbouiCOM.ISBOChooseFromListEventArg)
        'Dim sCFL_ID As String
        'sCFL_ID = oCFLEvento.ChooseFromListUID
        'Dim oCFL As SAPbouiCOM.ChooseFromList
        'oCFL = m_Form.ChooseFromLists.Item(sCFL_ID)
        'Dim val As String = ""
        'lData = True
        'If oCFLEvento.BeforeAction = False Then
        '    If IsNothing(oCFLEvento.SelectedObjects) Then lData = False
        '    Return oCFLEvento.SelectedObjects
        'End If
        'Return Nothing
    End Function


    Protected Function getOffset(ByVal condVal As String, ByVal AliasName As String, _
     ByVal oDataSource As SAPbouiCOM.DBDataSource) As Boolean

        m_Conditions = New SAPbouiCOM.Conditions
        m_Condition = m_Conditions.Add
        m_Condition.Alias = AliasName
        m_Condition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
        m_Condition.CondVal = condVal
        oDataSource.Clear()
        oDataSource.Query(m_Conditions)
        If oDataSource.Size = 0 Then
            getOffset = False
        Else
            getOffset = True
        End If
        m_Conditions = Nothing
        GC.Collect()
    End Function
    Protected Function getOffset(ByVal condVal As String, ByVal AliasName As String, _
        ByVal oDataSource As SAPbouiCOM.DBDataSource, ByVal operation As SAPbouiCOM.BoConditionOperation) As Boolean

        m_Conditions = New SAPbouiCOM.Conditions
        m_Condition = m_Conditions.Add
        m_Condition.Alias = AliasName
        m_Condition.Operation = operation
        m_Condition.CondVal = condVal
        oDataSource.Clear()
        oDataSource.Query(m_Conditions)
        If oDataSource.Size = 0 Then
            getOffset = False
        Else
            getOffset = True
        End If
        m_Conditions = Nothing
        GC.Collect()
    End Function
    Protected Function getOffset(ByVal condVal As String, ByVal AliasName As String, ByVal condVal1 As String, ByVal AliasName1 As String, _
            ByVal oDataSource As SAPbouiCOM.DBDataSource, ByVal operation As SAPbouiCOM.BoConditionOperation) As Boolean

        m_Conditions = New SAPbouiCOM.Conditions
        m_Condition = m_Conditions.Add
        m_Condition.BracketOpenNum = 2
        m_Condition.Alias = AliasName
        m_Condition.Operation = operation
        m_Condition.CondVal = condVal.Trim
        m_Condition.BracketCloseNum = 1

        m_Condition = m_Conditions.Add
        m_Condition.BracketOpenNum = 1
        m_Condition.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
        m_Condition.Alias = AliasName1
        m_Condition.Operation = operation
        m_Condition.CondVal = condVal1.Trim
        m_Condition.BracketCloseNum = 2

        oDataSource.Clear()
        oDataSource.Query(m_Conditions)
        If oDataSource.Size = 0 Then
            getOffset = False
        Else
            getOffset = True
        End If
        m_Conditions = Nothing
        GC.Collect()
    End Function

    Protected Overridable Sub QueryDBInit()

    End Sub

    Protected Overridable Sub FormRefresh()
        Try
            Me.m_Form.Freeze(True)
            For Each item As SAPbouiCOM.Item In m_Form.Items
                item.Update()
            Next
        Catch ex As Exception

        Finally
            m_Form.Freeze(False)

        End Try

    End Sub

    Protected Sub fillCombo(ByVal strKey As String, ByVal strDesc As String, _
       ByVal strTable As String, ByVal oCombo As SAPbouiCOM.Column, Optional ByVal strWhere As String = "", _
       Optional ByVal bAddNew As Boolean = False, Optional ByVal fieldType As SAPbobsCOM.BoFieldTypes = SAPbobsCOM.BoFieldTypes.db_Alpha, _
       Optional ByVal AddTo As Boolean = False, Optional ByVal OrderKey As String = "", Optional ByVal Desc As Boolean = True)
        Dim strSQL As String, Order As String = "DESC"
        Dim i As Integer
        Try
            If OrderKey = "" Then OrderKey = strKey

            If Not AddTo Then
                Try
                    For i = 0 To oCombo.ValidValues.Count - 1
                        oCombo.ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index)
                    Next

                Catch ex As Exception
                    m_ParentAddon.SboApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                End Try

            End If

            If Not Desc Then Order = "ASC"

            strSQL = " SELECT " + strKey + "," + strDesc + " FROM [" + strTable + "]"
            If strWhere <> String.Empty Then strSQL += " WHERE " + strWhere
            strSQL += " ORDER BY " + OrderKey + " " + Order
            DataTable("DT_Base").Clear()
            DataTable("DT_Base").ExecuteQuery(strSQL)

            For i = 0 To DataTable("DT_Base").Rows.Count - 1
                'fieldType = DataTable("DT_Base").Columns.Item(strKey).Type
                oCombo.ValidValues.Add(DataTable("DT_Base").GetValue(strKey, i), DataTable("DT_Base").GetValue(strDesc, i))
            Next




        Catch ex As Exception
            'Me.m_ParentAddon.WriteLog(ex.ToString)
            'SBO_Application.SetSatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
            m_ParentAddon.SboApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

        End Try
    End Sub

    Protected Sub fillCombo(ByVal strKey As String, ByVal strDesc As String, _
    ByVal strTable As String, ByVal oValues As SAPbouiCOM.ValidValues, Optional ByVal strWhere As String = "", _
    Optional ByVal bAddNew As Boolean = False, Optional ByVal fieldType As SAPbobsCOM.BoFieldTypes = SAPbobsCOM.BoFieldTypes.db_Alpha,
    Optional ByVal AddTo As Boolean = False, Optional ByVal OrderKey As String = "", Optional ByVal Desc As Boolean = True)
        Dim strSQL As String, Order As String = "DESC"

        Try
            Dim i As Integer
            If OrderKey = "" Then OrderKey = strKey
            If Not AddTo Then
                Try
                    For i = 0 To oValues.Count - 1
                        oValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index)
                    Next
                Catch ex As Exception
                    'm_ParentAddon.SboApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                End Try

            End If

            If Not Desc Then Order = "ASC"
            strSQL = " SELECT " + strKey + "," + strDesc + " FROM [" + strTable + "]"
            If strWhere <> String.Empty Then strSQL += " WHERE " + strWhere
            strSQL += " Order By " + OrderKey + " " + Order
            DataTable("DT_Base").Clear()
            DataTable("DT_Base").ExecuteQuery(strSQL)

            For i = 0 To DataTable("DT_Base").Rows.Count - 1
                fieldType = DataTable("DT_Base").Columns.Item(strKey).Type
                If String.IsNullOrEmpty(DataTable("DT_Base").GetValue(strKey, i)) Then
                    Continue For
                End If
                oValues.Add(DataTable("DT_Base").GetValue(strKey, i), DataTable("DT_Base").GetValue(strDesc, i))
            Next

            If bAddNew Then
                oValues.Add("Add New", "Add New")
            End If

        Catch ex As Exception
            'Me.m_ParentAddon.WriteLog(ex.ToString)
            m_ParentAddon.SboApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
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
            m_ParentAddon.SboApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Protected Sub fillCombo(ByVal strKey As String, ByVal strDesc As String, _
    ByVal strSQL As String, ByVal oCombo As SAPbouiCOM.ComboBox)


        Try
            Dim i As Integer
            For i = 0 To oCombo.ValidValues.Count - 1
                oCombo.ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index)
            Next

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
            m_ParentAddon.SboApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
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

    Protected ReadOnly Property DataTable(ByVal pst_DataTableUID As String) As SAPbouiCOM.DataTable
        Get
            Return m_Form.DataSources.DataTables.Item(pst_DataTableUID)
        End Get
    End Property
    Public ReadOnly Property UDS(ByVal pst_UDS_UID As String) As SAPbouiCOM.UserDataSource
        Get
            Return m_Form.DataSources.UserDataSources.Item(pst_UDS_UID)
        End Get
    End Property

    Public ReadOnly Property DBDS(ByVal pbo_Object As Object) As SAPbouiCOM.DBDataSource
        Get
            If CStr(pbo_Object).ToUpper Like "SBO*".ToUpper Then
                pbo_Object = "@" & CStr(pbo_Object)
            End If
            Return m_Form.DataSources.DBDataSources.Item(pbo_Object)
        End Get
    End Property

    Public Function sboDate(ByVal thedate As Date) As String
        sboDate = thedate.Year.ToString + _
                            thedate.Month.ToString.PadLeft(2, "0") + _
                            thedate.Day.ToString.PadLeft(2, "0")
    End Function

    Public Function sboDate(ByVal thedate As String) As Date
        Dim strDate As String
        Dim retDate As Date

        Dim dateSep As String = My.Application.Culture.DateTimeFormat.DateSeparator

        strDate = thedate.Substring(6, 2) + dateSep + _
            thedate.Substring(4, 2) + dateSep + thedate.Substring(0, 4)

        Date.TryParse(strDate, retDate)
        sboDate = retDate

        'sboDate = strDate.ToString(My.Application.Culture.DateTimeFormat.ShortDatePattern)
    End Function

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


    Protected Overridable Sub EnableToolBarButtons()
        'm_Form.EnableMenu("1288", True)
        'm_Form.EnableMenu("1289", True)
        'm_Form.EnableMenu("1290", True)
        'm_Form.EnableMenu("1291", True)
    End Sub

    Protected Overridable Sub OnMatrixAddRow()

    End Sub

    Protected Overridable Sub OnMatrixDeleteRow()

    End Sub

    Protected Overridable Sub OnMatrixDeleteAllRows()

    End Sub
    Protected Overridable Sub AddRowToMatrix(ByVal pMatrix As SAPbouiCOM.Matrix)

        pMatrix.AddRow(1)
        pMatrix.SetCellFocus(pMatrix.RowCount, 1)
    End Sub

    Protected Sub DeleteMatrixRow(ByVal pMatrix As SAPbouiCOM.Matrix)
        pMatrix.DeleteRow(m_CurrentLineNo)
        'pMatrix.FlushToDataSource()
    End Sub

    Protected Sub DeleteAllMatrixRows(ByVal pMatrix As SAPbouiCOM.Matrix)
        pMatrix.Clear()
        'pMatrix.FlushToDataSource()
    End Sub
    Protected Sub AddRowToGrid(ByVal pGrid As SAPbouiCOM.Grid, ByVal pDataTable As SAPbouiCOM.DataTable)
        pDataTable.Rows.Add(1)
        pGrid.Click(pDataTable.Rows.Count - 1)
    End Sub


    Private Sub m_SboApplication_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) Handles m_SboApplication.MenuEvent
        Try
            If Me.disposed Then
                Return
            End If
            If pVal.BeforeAction = False Then
                Select Case pVal.MenuUID
                    Case "1292"
                        OnMatrixAddRow()
                    Case "1288", "1289", "1290", "1291"
                        OnFormNavigate()
                    Case "1282"
                        OnMatrixDeleteAllRows()
                        OnFormNavigate()
                    Case "1293"
                        OnMatrixDeleteRow()
                End Select
            End If
        Catch ex As Exception

        End Try
    End Sub

    Protected Function validateDataSource(ByVal oDBDS As SAPbouiCOM.DBDataSource, ByVal Column As String, ByRef RowIndex As Integer) As Boolean
        Try
            Dim ht As Hashtable = New Hashtable(), strKey As String

            For index = 0 To oDBDS.Size - 1
                strKey = oDBDS.GetValue(Column, index)
                If ht.ContainsKey(strKey) Then
                    RowIndex = index
                    Return False
                Else
                    ht.Add(strKey, strKey)
                End If
            Next
            Return True
        Catch ex As Exception
            m_ParentAddon.SboApplication.SetStatusBarMessage(ex.Message.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, True)
            Return False
        End Try

    End Function

    Protected Function validateDataSource(ByVal oDT As SAPbouiCOM.DataTable, ByVal Column As String, ByRef RowIndex As Integer) As Boolean
        Try
            Dim ht As Hashtable = New Hashtable(), strKey As String

            For index = 0 To oDT.Rows.Count - 1
                strKey = oDT.GetValue(Column, index)
                If ht.ContainsKey(strKey) Then
                    RowIndex = index
                    Return False
                Else
                    ht.Add(strKey, strKey)
                End If
            Next
            Return True
        Catch ex As Exception
            m_ParentAddon.SboApplication.SetStatusBarMessage(ex.Message.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, True)
            Return False
        End Try

    End Function


    Protected Overrides Sub Finalize()
        MyBase.Finalize()
        Me.m_Form = Nothing
    End Sub
End Class


