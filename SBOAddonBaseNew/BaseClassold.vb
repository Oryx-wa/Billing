Option Strict Off
Option Explicit On
Imports SAPbouiCOM.Framework
Imports System.IO
Imports System.Reflection



Public MustInherit Class UserFormBaseClassOld
    Inherits UserFormBase

    Protected WithEvents SBO_Application As SAPbouiCOM.Application
    Protected oDataSource As SAPbouiCOM.DBDataSource, oDataSource2 As SAPbouiCOM.DBDataSource, oDataSource3 As SAPbouiCOM.DBDataSource
    Protected oDataTable As SAPbouiCOM.DataTable, oDataTable2 As SAPbouiCOM.DataTable
    Protected oRecordSet As SAPbobsCOM.Recordset
    Protected oForm As SAPbouiCOM.Form
    Protected oConditions As SAPbouiCOM.Conditions
    Protected oCondition As SAPbouiCOM.Condition
    Protected oCompany As SAPbobsCOM.Company
    Protected oGeneralService As SAPbobsCOM.GeneralService
    Protected oGeneralData As SAPbobsCOM.GeneralData
    Protected oCompanyService As SAPbobsCOM.CompanyService
    Protected oCFLEvento As SAPbouiCOM.IChooseFromListEvent, oCFL As SAPbouiCOM.ChooseFromList, ochooseTable As SAPbouiCOM.DataTable
    Protected errCode As Integer, errMsg As String
    Protected oCFLEventArg As SAPbouiCOM.ISBOChooseFromListEventArg
    Protected m_ParentAddon As SboAddon

    Public Sub New()
        'initbase()
        oForm.DataSources.DataTables.Add("DT_Seq")
        oForm.DataSources.DataTables.Add("DT_Base")
        EnableToolBarButtons()
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

    Public Sub initbase(ByVal pAddOn As SboAddon)
        SBO_Application = SAPbouiCOM.Framework.Application.SBO_Application
        oCompany = SBO_Application.Company.GetDICompany
        oRecordSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oCompanyService = oCompany.GetCompanyService
        oForm = UIAPIRawForm
        m_ParentAddon = pAddOn 'New SboAddon()
    End Sub

    Private Sub EnableToolBarButtons()
        With oForm
            .EnableMenu("1282", True)
            .EnableMenu("1292", True)
            .EnableMenu("1293", True)
            .EnableMenu("1287", True)
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
        oConditions = New SAPbouiCOM.Conditions
        oCondition = oConditions.Add
        oCondition.Alias = AliasName
        oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
        oCondition.CondVal = condVal
        If oDataSource.Size > 0 Then oDataSource.Clear()
        oDataSource.Query(oConditions)
        If oDataSource.Size = 0 Then
            validatealias = False
        Else
            validatealias = True
        End If
        oConditions = Nothing
        GC.Collect()
    End Function
    Protected Overridable Function validatealias(ByVal condVal As String, ByVal AliasName As String, _
         ByVal oDataSource As SAPbouiCOM.DBDataSource, ByVal condVal1 As String, ByVal aliasName1 As String) As Boolean
        oConditions = New SAPbouiCOM.Conditions
        oCondition = oConditions.Add
        oCondition.Alias = AliasName
        oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
        oCondition.CondVal = condVal
        oCondition.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
        oCondition = oConditions.Add
        oCondition.Alias = aliasName1
        oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
        oCondition.CondVal = condVal1

        oDataSource.Clear()
        oDataSource.Query(oConditions)
        If oDataSource.Size = 0 Then
            validatealias = False
        Else
            validatealias = True
        End If
        oConditions = Nothing
        GC.Collect()
    End Function

    Protected Function getOffset(ByVal condVal As String, ByVal AliasName As String, _
         ByVal oDataSource As SAPbouiCOM.DBDataSource) As Boolean

        oConditions = New SAPbouiCOM.Conditions
        oCondition = oConditions.Add
        oCondition.Alias = AliasName
        oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
        oCondition.CondVal = condVal
        oDataSource.Clear()
        oDataSource.Query(oConditions)
        If oDataSource.Size = 0 Then
            getOffset = False
        Else
            getOffset = True
        End If
        oConditions = Nothing
        GC.Collect()
    End Function
    Protected Function getOffset(ByVal condVal As String, ByVal AliasName As String, _
        ByVal oDataSource As SAPbouiCOM.DBDataSource, ByVal operation As SAPbouiCOM.BoConditionOperation) As Boolean

        oConditions = New SAPbouiCOM.Conditions
        oCondition = oConditions.Add
        oCondition.Alias = AliasName
        oCondition.Operation = operation
        oCondition.CondVal = condVal
        oDataSource.Clear()
        oDataSource.Query(oConditions)
        If oDataSource.Size = 0 Then
            getOffset = False
        Else
            getOffset = True
        End If
        oConditions = Nothing
        GC.Collect()
    End Function
    Protected Function getOffset(ByVal condVal As String, ByVal AliasName As String, ByVal condVal1 As String, ByVal AliasName1 As String, _
            ByVal oDataSource As SAPbouiCOM.DBDataSource, ByVal operation As SAPbouiCOM.BoConditionOperation) As Boolean

        oConditions = New SAPbouiCOM.Conditions
        oCondition = oConditions.Add
        oCondition.BracketOpenNum = 2
        oCondition.Alias = AliasName
        oCondition.Operation = operation
        oCondition.CondVal = condVal.Trim
        oCondition.BracketCloseNum = 1

        oCondition = oConditions.Add
        oCondition.BracketOpenNum = 1
        oCondition.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
        oCondition.Alias = AliasName1
        oCondition.Operation = operation
        oCondition.CondVal = condVal1.Trim
        oCondition.BracketCloseNum = 2

        oDataSource.Clear()
        oDataSource.Query(oConditions)
        If oDataSource.Size = 0 Then
            getOffset = False
        Else
            getOffset = True
        End If
        oConditions = Nothing
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


End Class

