Option Strict Off
Option Explicit On
Imports SAPbouiCOM.Framework
Imports System.IO
Imports System.Reflection



Public MustInherit Class UserFormBaseClass
    Inherits UserFormBase
    Implements SAPbouiCOM.IForm

    Protected WithEvents SBO_Application As SAPbouiCOM.Application
    Protected Friend oDataSource As SAPbouiCOM.DBDataSource, oDataSource2 As SAPbouiCOM.DBDataSource, oDataSource3 As SAPbouiCOM.DBDataSource
    Protected Friend oDataTable As SAPbouiCOM.DataTable, oDataTable2 As SAPbouiCOM.DataTable
    Protected oRecordSet As SAPbobsCOM.Recordset
    Protected Conditions0 As SAPbouiCOM.Conditions
    Protected Condition0 As SAPbouiCOM.Condition
    Protected oCompany As SAPbobsCOM.Company
    Protected oGeneralService As SAPbobsCOM.GeneralService
    Protected oGeneralData As SAPbobsCOM.GeneralData
    Protected oCompanyService As SAPbobsCOM.CompanyService
    Protected oCFLEvento As SAPbouiCOM.IChooseFromListEvent, oCFL As SAPbouiCOM.ChooseFromList, ochooseTable As SAPbouiCOM.DataTable
    Protected errCode As Integer, errMsg As String
    Protected oCFLEventArg As SAPbouiCOM.ISBOChooseFromListEventArg
    Protected m_ParentAddon As SboAddon
    Protected WithEvents m_Matrix As SAPbouiCOM.Matrix
    Protected m_CurrentMatrixRow As Integer
    Protected MatrixDBDataSource As SAPbouiCOM.DBDataSource
    Protected m_BaseObject As SBOBaseObject


    Public Sub New()
        oForm.DataSources.DataTables.Add("DT_Seq")
        oForm.DataSources.DataTables.Add("DT_Base")


    End Sub

    Protected ReadOnly Property DataTable(ByVal pst_DataTableUID As String) As SAPbouiCOM.DataTable
        Get
            Return UIAPIRawForm.DataSources.DataTables.Item(pst_DataTableUID)
        End Get
    End Property
    Public ReadOnly Property UDS(ByVal pst_UDS_UID As String) As SAPbouiCOM.UserDataSource
        Get
            Return UIAPIRawForm.DataSources.UserDataSources.Item(pst_UDS_UID)
        End Get
    End Property

    Public ReadOnly Property CFL(ByVal pst_UDS_UID As String) As SAPbouiCOM.ChooseFromList
        Get
            Return UIAPIRawForm.ChooseFromLists.Item(pst_UDS_UID)
        End Get
    End Property
    Public ReadOnly Property DBDS(ByVal pbo_Object As Object) As SAPbouiCOM.DBDataSource
        Get
            If CStr(pbo_Object).ToUpper Like "SBO*".ToUpper Then
                pbo_Object = "@" & CStr(pbo_Object)
            End If
            Return UIAPIRawForm.DataSources.DBDataSources.Item(pbo_Object)
        End Get
    End Property
    Public ReadOnly Property USERDS(ByVal pbo_Object As Object) As SAPbouiCOM.UserDataSource
        Get
            Return UIAPIRawForm.DataSources.UserDataSources.Item(pbo_Object)
        End Get
    End Property

    Protected Friend Overridable Sub InitBase(pAddOn As SboAddon)
        SBO_Application = SAPbouiCOM.Framework.Application.SBO_Application

        m_ParentAddon = pAddOn
        oCompany = m_ParentAddon.SboCompany


        EnableToolBarButtons()
    End Sub

    Protected Overridable Sub CreateObject(pObjRef As SBOBaseObject)
        m_BaseObject = pObjRef
        m_BaseObject.OnCustomInit()
    End Sub
    Protected ReadOnly Property oForm As SAPbouiCOM.Form
        Get
            Return Me.UIAPIRawForm
        End Get
    End Property

    Public ReadOnly Property BusObjectInfo As SBOBaseObject
        Get
            Return m_BaseObject
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



    Protected Overridable Sub AddRowToMatrix()
        'MatrixDBDataSource.Clear()
        ''If m_CurrentMatrixRow = m_Matrix.RowCount + 1 Then
        'If m_CurrentMatrixRow <= 1 Then
        '    m_Matrix.AddRow(1)

        'Else
        '    m_Matrix.AddRow(1, m_Matrix.RowCount)
        'End If
        'm_Matrix.Columns.Item(1).Cells.Item(m_CurrentMatrixRow).Click()
        'm_Matrix.FlushToDataSource()

        'End If
    End Sub

    Protected Overridable Sub HandleNavigation()

    End Sub

    Protected Overridable Sub EnableToolBarButtons()
        With oForm
            '.EnableMenu("1282", True)
            '.EnableMenu("1283", True)

            '.EnableMenu("1288", True)
            '.EnableMenu("1289", True)
            '.EnableMenu("1290", True)
            '.EnableMenu("1291", True)


            '.EnableMenu("1292", True)
            '.EnableMenu("1293", True)
            '.EnableMenu("1287", True)
        End With
    End Sub

    Public Sub fillCombo(ByVal strKey As String, ByVal strDesc As String, _
        ByVal strTable As String, ByVal oCombo As SAPbouiCOM.ComboBox, Optional ByVal strWhere As String = "", _
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
                    SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
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

            If bAddNew Then
                oCombo.ValidValues.Add("Add New", "Add New")
            End If


        Catch ex As Exception
            'Me.m_ParentAddon.WriteLog(ex.ToString)
            'SBO_Application.SetSatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
            SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

        End Try
    End Sub

    Public Sub fillCombo(ByVal strKey As String, ByVal strDesc As String, _
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
                    SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
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

            If bAddNew Then
                oCombo.ValidValues.Add("Add New", "Add New")
            End If


        Catch ex As Exception
            'Me.m_ParentAddon.WriteLog(ex.ToString)
            'SBO_Application.SetSatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
            SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

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
                    SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
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

            'oValues.Add("", "")

        Catch ex As Exception
            'Me.m_ParentAddon.WriteLog(ex.ToString)
            SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
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
            SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
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

    Public Function getOffset(ByVal condVal As String, ByVal AliasName As String, _
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

    Protected Sub chooseFromListSystem(ByVal strObjType As String, ByVal multiselection As Boolean, _
                     ByVal strID As String, ByVal conAlias As String, ByVal strConVal As String)

        Try
            Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition

            oCFLs = oForm.ChooseFromLists

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
            SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Protected Sub chooseFromListSystem(ByVal strObjType As String, ByVal multiselection As Boolean, _
                        ByVal strID As String, ByVal conAlias As String, ByVal strConVal As String, _
                        ByVal conAlias1 As String, ByVal strConVal1 As String)

        Try
            Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition

            oCFLs = oForm.ChooseFromLists

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
            SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Public Function HandleChooseFromListEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As Boolean) As String
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
    Protected Function HandleChooseFromListEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean, ByRef lData As Boolean) As SAPbouiCOM.DataTable
        Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
        oCFLEvento = pVal
        Dim sCFL_ID As String
        sCFL_ID = oCFLEvento.ChooseFromListUID
        Dim oCFL As SAPbouiCOM.ChooseFromList
        oCFL = oForm.ChooseFromLists.Item(sCFL_ID)
        Dim val As String = ""
        lData = True
        If oCFLEvento.BeforeAction = False Then
            If IsNothing(oCFLEvento.SelectedObjects) Then lData = False
            Return oCFLEvento.SelectedObjects
        End If
        Return Nothing
    End Function


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

    Public Function getUserTables(ByVal tableName As String) As SAPbobsCOM.UserTable
        Dim table As SAPbobsCOM.UserTable, aTableList As SAPbobsCOM.UserTables
        table = Nothing
        aTableList = oCompany.UserTables
        For Each table In aTableList
            If table.TableName = tableName.Substring(1) Then
                'table = oCompany.UserTables.Item(tableName)
                Exit For
            End If
        Next
        Return table
    End Function

    Protected Function GetSQLString(ByVal strProc As String, ByVal ParamArray Parameters() As String) As String
        Try
            Dim a As Assembly = System.Reflection.Assembly.GetCallingAssembly()

            Dim Ret As String = Nothing
            Dim strMsg As String = "", strSQL As String = ""
            strProc = "FleetMgt." & strProc
            Dim SReader As New System.IO.StreamReader(a.GetManifestResourceStream(strProc + ".txt"))

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

    Public Overrides Sub OnInitializeFormEvents()


    End Sub



    Private Sub SBO_Application_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.MenuEvent
        

    End Sub

    Public Property ActiveItem As String Implements SAPbouiCOM.IForm.ActiveItem
        Get

        End Get
        Set(value As String)

        End Set
    End Property

    Public Property AutoManaged As Boolean Implements SAPbouiCOM.IForm.AutoManaged
        Get

        End Get
        Set(value As Boolean)

        End Set
    End Property

    Public ReadOnly Property BorderStyle As SAPbouiCOM.BoFormBorderStyle Implements SAPbouiCOM.IForm.BorderStyle
        Get

        End Get
    End Property

    Public ReadOnly Property BusinessObject As SAPbouiCOM.BusinessObject Implements SAPbouiCOM.IForm.BusinessObject
        Get

        End Get
    End Property

    Public ReadOnly Property ChooseFromLists As SAPbouiCOM.ChooseFromListCollection Implements SAPbouiCOM.IForm.ChooseFromLists
        Get

        End Get
    End Property

    Public Property ClientHeight As Integer Implements SAPbouiCOM.IForm.ClientHeight
        Get

        End Get
        Set(value As Integer)

        End Set
    End Property

    Public Property ClientWidth As Integer Implements SAPbouiCOM.IForm.ClientWidth
        Get

        End Get
        Set(value As Integer)

        End Set
    End Property

    Public Sub Close() Implements SAPbouiCOM.IForm.Close

    End Sub

    Public ReadOnly Property DataBrowser As SAPbouiCOM.DataBrowser Implements SAPbouiCOM.IForm.DataBrowser
        Get

        End Get
    End Property

    Public ReadOnly Property DataSources As SAPbouiCOM.DataSource Implements SAPbouiCOM.IForm.DataSources
        Get

        End Get
    End Property

    Public Property DefButton As String Implements SAPbouiCOM.IForm.DefButton
        Get

        End Get
        Set(value As String)

        End Set
    End Property

    Public Sub EnableFormatSearch() Implements SAPbouiCOM.IForm.EnableFormatSearch

    End Sub

    Public Sub EnableMenu(MenuUID As String, EnableFlag As Boolean) Implements SAPbouiCOM.IForm.EnableMenu

    End Sub

    Public Sub Freeze(newVal As Boolean) Implements SAPbouiCOM.IForm.Freeze

    End Sub

    Public Function GetAsXML() As String Implements SAPbouiCOM.IForm.GetAsXML

    End Function

    Public Property Height As Integer Implements SAPbouiCOM.IForm.Height
        Get

        End Get
        Set(value As Integer)

        End Set
    End Property

    Public ReadOnly Property IsSystem As Boolean Implements SAPbouiCOM.IForm.IsSystem
        Get

        End Get
    End Property

    Public ReadOnly Property Items As SAPbouiCOM.Items Implements SAPbouiCOM.IForm.Items
        Get

        End Get
    End Property

    Public Property Left As Integer Implements SAPbouiCOM.IForm.Left
        Get

        End Get
        Set(value As Integer)

        End Set
    End Property

    Public Property MaxHeight As Integer Implements SAPbouiCOM.IForm.MaxHeight
        Get

        End Get
        Set(value As Integer)

        End Set
    End Property

    Public Property MaxWidth As Integer Implements SAPbouiCOM.IForm.MaxWidth
        Get

        End Get
        Set(value As Integer)

        End Set
    End Property

    Public ReadOnly Property Menu As SAPbouiCOM.Menus Implements SAPbouiCOM.IForm.Menu
        Get

        End Get
    End Property

    Public ReadOnly Property Modal As Boolean Implements SAPbouiCOM.IForm.Modal
        Get

        End Get
    End Property

    Public Property Mode As SAPbouiCOM.BoFormMode Implements SAPbouiCOM.IForm.Mode
        Get

        End Get
        Set(value As SAPbouiCOM.BoFormMode)

        End Set
    End Property

    Public Property PaneLevel As Integer Implements SAPbouiCOM.IForm.PaneLevel
        Get

        End Get
        Set(value As Integer)

        End Set
    End Property

    Public Sub Refresh() Implements SAPbouiCOM.IForm.Refresh

    End Sub

    Public Property ReportType As String Implements SAPbouiCOM.IForm.ReportType
        Get

        End Get
        Set(value As String)

        End Set
    End Property

    Public Sub ResetMenuStatus() Implements SAPbouiCOM.IForm.ResetMenuStatus

    End Sub

    Public Sub Resize(lWidth As Integer, lHeight As Integer) Implements SAPbouiCOM.IForm.Resize

    End Sub

    Public Sub [Select]() Implements SAPbouiCOM.IForm.Select

    End Sub

    Public ReadOnly Property Selected As Boolean Implements SAPbouiCOM.IForm.Selected
        Get

        End Get
    End Property

    Public ReadOnly Property Settings As SAPbouiCOM.FormSettings Implements SAPbouiCOM.IForm.Settings
        Get

        End Get
    End Property

    Public Property State As SAPbouiCOM.BoFormStateEnum Implements SAPbouiCOM.IForm.State
        Get

        End Get
        Set(value As SAPbouiCOM.BoFormStateEnum)

        End Set
    End Property

    Public Property SupportedModes As Integer Implements SAPbouiCOM.IForm.SupportedModes
        Get

        End Get
        Set(value As Integer)

        End Set
    End Property

    Public Property Title As String Implements SAPbouiCOM.IForm.Title
        Get

        End Get
        Set(value As String)

        End Set
    End Property

    Public Property Top As Integer Implements SAPbouiCOM.IForm.Top
        Get

        End Get
        Set(value As Integer)

        End Set
    End Property

    Public ReadOnly Property Type As Integer Implements SAPbouiCOM.IForm.Type
        Get

        End Get
    End Property

    Public ReadOnly Property TypeCount As Integer Implements SAPbouiCOM.IForm.TypeCount
        Get

        End Get
    End Property

    Public ReadOnly Property TypeEx As String Implements SAPbouiCOM.IForm.TypeEx
        Get

        End Get
    End Property

    Public ReadOnly Property UDFFormUID As String Implements SAPbouiCOM.IForm.UDFFormUID
        Get

        End Get
    End Property

    Public ReadOnly Property UniqueID As String Implements SAPbouiCOM.IForm.UniqueID
        Get

        End Get
    End Property

    Public Sub Update() Implements SAPbouiCOM.IForm.Update

    End Sub

    Public Property Visible As Boolean Implements SAPbouiCOM.IForm.Visible
        Get

        End Get
        Set(value As Boolean)

        End Set
    End Property

    Public Property VisibleEx As Boolean Implements SAPbouiCOM.IForm.VisibleEx
        Get

        End Get
        Set(value As Boolean)

        End Set
    End Property

    Public Property Width As Integer Implements SAPbouiCOM.IForm.Width
        Get

        End Get
        Set(value As Integer)

        End Set
    End Property
End Class

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