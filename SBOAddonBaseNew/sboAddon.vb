
Imports System
Imports System.IO
Imports System.Xml
Imports System.Xml.XPath
'Imports System.Windows.Forms
Imports System.Runtime.Remoting
Imports System.Text
Imports System.Configuration
Imports SAPbouiCOM.Framework


Public Interface ISboAddon

#Region "Declaration for Methods"
    Function getUserTables(ByVal tableName As String) As SAPbobsCOM.UserTable
    Function checkTables(ByVal ds As Tables) As Boolean
    Function InitialiseTables() As Boolean
    Function createTables(ByVal ds As Tables) As Boolean
    Function AddFields(ByVal ds As Tables) As Boolean
    Function AddUDO(ByVal ds As Tables) As Boolean

    'Function ConnectToSBO() As Boolean

    'Function UpdateSBOForm(ByVal XMLDocumentName As String, ByVal FormUID As String) As SAPbouiCOM.Form
    'Overloads Function CreateSBOForm(ByVal XMLDocumentName As String) As SAPbouiCOM.Form
    'Overloads Function CreateSBOForm(ByVal XMLDocumentName As String, ByVal formType As String) As SAPbouiCOM.Form
    Sub SendAddonBusy()

#End Region

#Region "Declaration for Properties"
    ReadOnly Property FilePath() As String
    ReadOnly Property FileName() As String
    ReadOnly Property Connected() As Boolean
    ReadOnly Property Name() As String
    ReadOnly Property HomePath() As String
    ReadOnly Property SboApplication() As SAPbouiCOM.Application
    ReadOnly Property SboCompany() As SAPbobsCOM.Company
    ReadOnly Property SBOForms() As Collection
    ReadOnly Property StartupPath() As String
    'ReadOnly Property AppConfig() As AppConfig
    ReadOnly Property AppNameSpace() As String
    ReadOnly Property AppAssemblyName() As String
    ReadOnly Property Menuds() As DataSet
    Property BlockEvents() As Boolean
    ReadOnly Property sqlConn() As SqlClient.SqlConnection
    Property aTableList() As SboTables
    Property aFieldlist() As SboFields
    Property aUDOList() As SboUDOs
    Property UseDI() As Boolean
    Property PermissionPrefix() As String
    Property TablePrefix() As String
    Property MenuXMLFileName() As String
    Property usePermissions As Boolean


#End Region
End Interface

Public MustInherit Class SboAddon

    Implements ISboAddon

#Region "Variables"
    Const DEVELOPERSCONNECTIONSTRING As String = "0030002C0030002C00530041005000420044005F00440061007400650076002C0050004C006F006D0056004900490056"

    Private oUserFieldsMD As SAPbobsCOM.UserFieldsMD
    Private dt As Tables.usertableDataTable
    Private oUserTablesMD As SAPbobsCOM.UserTablesMD, oUserObjectMD As SAPbobsCOM.UserObjectsMD
    Private Tablerow As Tables.usertableRow, FieldRow As Tables.TableFieldsRow, objectRow As Tables.UserObjectsRow
    Private validRow As Tables.ValidValuesRow 'UserKeyRow As Tables.UserKeysRow
    Private errCode As Integer, errMsg As String
    Private oProgBar As SAPbouiCOM.ProgressBar
    Private oUserKeysMD As SAPbobsCOM.UserKeysMD
    Private m_permissions As Permissions
    Private m_tablePrefix As String
    Private m_permissionPrefix As String
    Private m_MenuXMLFileName As String
    Private m_useDI As Boolean = True
    Private m_usePermissions As Boolean = False
    Private m_assembly As System.Reflection.Assembly = Nothing
    Private aTablesCreated As New ArrayList

    Protected m_Namespace As String
    Protected m_AssemblyName As String
    Protected WithEvents m_SboApplication As SAPbouiCOM.Application
    Protected m_SboCompany As SAPbobsCOM.Company
    Protected m_SboAddons As Collection
    Protected m_StartUpPath As String
    Protected m_CompanyName As String
    Protected m_Name As String
    Protected m_AddOnHomePath As String
    Protected m_SBOForms As Collection
    Protected m_Connected As Boolean
    Protected colFormsCurrentlyOpen As Collection
    Protected m_sqlConnection As New SqlClient.SqlConnection
    Protected m_useSequence As Boolean = True
    Protected m_AppConfig As AppSettingsReader
    Protected m_EventsBlocked As Boolean
    Protected m_ds As New DataSet
    Protected m_ds2 As New DataSet
    Protected m_file As System.IO.Stream
    Protected m_aTableList As SboTables = Nothing
    Protected m_fieldlist As SboFields = Nothing
    Protected m_UDOlist As SboUDOs = Nothing
    Private m_FilePath As String = ""
    Private m_FileName As String = ""
    Private m_FileExt As String = ""
    Protected InitTables As Boolean = False
    Protected initPerm As Boolean = False
    Protected m_MenuImageFile As String
    Protected UsePhysicalFiles As Boolean
    Public oApp As Application
    Private m_PermUse As Boolean
    Public sysInfo As String = "", sysFormName As String = ""
#End Region

#Region "Propeties"

    Public ReadOnly Property FilePath() As String Implements ISboAddon.FilePath
        Get
            Return m_FilePath
        End Get
    End Property

    Public ReadOnly Property FileName() As String Implements ISboAddon.FileName
        Get
            Return m_FileName
        End Get
    End Property
    Public ReadOnly Property Connected() As Boolean Implements ISboAddon.Connected
        Get
            Return m_Connected
        End Get

    End Property

    Public ReadOnly Property Name() As String Implements ISboAddon.Name
        Get
            Return m_Name
        End Get
    End Property

    Public ReadOnly Property HomePath() As String Implements ISboAddon.HomePath
        Get
            Return m_AddOnHomePath
        End Get
    End Property

    Public ReadOnly Property SboApplication() As SAPbouiCOM.Application Implements ISboAddon.SboApplication
        Get
            Return SAPbouiCOM.Framework.Application.SBO_Application
        End Get
    End Property

    Public ReadOnly Property SboCompany() As SAPbobsCOM.Company Implements ISboAddon.SboCompany
        Get
            If m_SboCompany Is Nothing Then
                If ConnectToDIAPI() = False Then
                    Dim strMessage As String = " Error connecting to DI API the Addon Will now exit "
                    m_SboApplication.MessageBox(strMessage)
                    System.Windows.Forms.Application.Exit()
                End If
                'Else
                '    m_SboCompany = Nothing
            End If
            Return m_SboCompany
        End Get
    End Property

    Public ReadOnly Property SBOForms() As Collection Implements ISboAddon.SBOForms
        Get
            Return m_SBOForms
        End Get
    End Property

    Public ReadOnly Property StartupPath() As String Implements ISboAddon.StartupPath
        Get
            Return m_StartUpPath
        End Get
    End Property

    Public Property BlockEvents() As Boolean Implements ISboAddon.BlockEvents
        Get
            Return m_EventsBlocked
        End Get
        Set(ByVal Value As Boolean)
            m_EventsBlocked = Value
        End Set
    End Property

    Protected Property AppAssembly() As System.Reflection.Assembly
        Get
            Return m_assembly
        End Get
        Set(ByVal Value As System.Reflection.Assembly)
            m_assembly = Value
        End Set
    End Property

    'Public ReadOnly Property AppConfig() As AppConfig Implements ISboAddOn.AppConfig
    '    Get
    '        Return m_AppConfig
    '    End Get
    'End Property
    Public ReadOnly Property AppNameSpace() As String Implements ISboAddon.AppNameSpace
        Get
            Return m_Namespace
        End Get
    End Property
    Public ReadOnly Property appAssemblyName() As String Implements ISboAddon.AppAssemblyName
        Get
            Return m_AssemblyName
        End Get
    End Property


    Public ReadOnly Property Menuds() As DataSet Implements ISboAddon.Menuds
        Get
            If m_ds.Tables.Count = 0 Then
                m_file = getAppResource("formsandmenus.xml")
                m_ds.ReadXml(m_file, XmlReadMode.Auto)
            End If
            Return m_ds
        End Get
    End Property
    Public ReadOnly Property FormItemDs() As DataSet
        Get
            If m_ds2.Tables.Count = 0 Then
                m_file = getAppResource("formanditems.xml")
                m_ds2.ReadXml(m_file, XmlReadMode.Auto)
            End If
            Return m_ds2
        End Get
    End Property


    Public ReadOnly Property sqlConn() As SqlClient.SqlConnection Implements ISboAddon.sqlConn
        Get
            Dim myConnectionString As String = ""
            Try
                If m_sqlConnection.State = ConnectionState.Closed Then
                    Dim orecset As SAPbobsCOM.Recordset
                    orecset = SboCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    orecset.DoQuery(" SELECT * FROM Z001 ")

                    Dim strPassword As String = orecset.Fields.Item("d002").Value
                    Dim strUser As String = orecset.Fields.Item("d001").Value

                    'Dim server As String = ""

                    'Build the connection and connect to the database
                    myConnectionString = "Persist Security Info=False;Password=" + strPassword _
                     + ";User ID=" + strUser + _
                     ";database=" + SboCompany.CompanyDB.Trim + _
                     ";server=" + SboCompany.Server.Trim + _
                     ";Connect Timeout=0"
                    m_sqlConnection.ConnectionString = myConnectionString
                    'If My.Application..AppMode = "DEBUG" Then
                    '    MsgBox("Connectionstring: " & myConnectionString)
                    'End If
                    m_sqlConnection.Open()
                End If
                Return m_sqlConnection
            Catch ex As SqlClient.SqlException
                'WriteLog(myConnectionString + "  " + ex.ToString)
                Return Nothing
            Catch ex As Exception
                'WriteLog(myConnectionString + "  " + ex.ToString)
                Return Nothing
            End Try

        End Get
    End Property

    Public Property aTableList() As SboTables Implements ISboAddon.aTableList
        Get
            If m_aTableList Is Nothing Then
                Try
                    Dim i As Integer = 0
                    Dim orecset As SAPbobsCOM.Recordset
                    orecset = SboCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    orecset.DoQuery("SELECT * FROM OUTB  ")
                    'If orecset.RecordCount = 0 Then
                    '    m_aTableList = Nothing
                    '    Exit Try
                    'End If
                    m_aTableList = New SboTables
                    orecset.MoveFirst()
                    While Not orecset.EoF
                        i += 1
                        If CType(orecset.Fields.Item("TableName").Value, String).StartsWith(m_tablePrefix) Then

                            m_aTableList.Add(New enTableNamesType(orecset.Fields.Item("TableName").Value, i))
                        End If
                        orecset.MoveNext()
                    End While
                Catch ex As Exception
                    SboApplication.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Long, True)
                    m_aTableList = Nothing
                End Try
            End If
            Return m_aTableList
        End Get
        Set(ByVal Value As SboTables)
            m_aTableList = Nothing
        End Set
    End Property
    Public Property aFieldlist() As SboFields Implements ISboAddon.aFieldlist
        Get
            If m_fieldlist Is Nothing Then
                Try
                    Dim i As Integer = 0
                    'Dim field As enFieldNamesType
                    Dim orecset As SAPbobsCOM.Recordset
                    orecset = SboCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    orecset.DoQuery("SELECT * FROM CUFD WHERE TableID Like '@" + m_tablePrefix + "%' union all" + _
                                    " SELECT * FROM CUFD WHERE TableID not Like '@%'")

                    m_fieldlist = New SboFields
                    orecset.MoveFirst()
                    While Not orecset.EoF
                        i += 1
                        If orecset.Fields.Item("TableID").Value.ToString.StartsWith("@") Then
                            Dim fldName As String = orecset.Fields.Item("AliasID").Value
                            m_fieldlist.Add(New enFieldNamesType(orecset.Fields.Item("AliasID").Value, orecset.Fields.Item("TableID").Value.ToString.Substring(1)))
                        Else
                            m_fieldlist.Add(New enFieldNamesType(orecset.Fields.Item("AliasID").Value, orecset.Fields.Item("TableID").Value))
                        End If

                        orecset.MoveNext()
                    End While
                Catch ex As Exception
                    m_fieldlist = Nothing
                    SboApplication.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Long, True)
                End Try
            End If
            Return m_fieldlist
        End Get
        Set(ByVal Value As SboFields)
            m_fieldlist = Nothing
        End Set
    End Property

    Public Function CreateSonForm(pFormName As String) As UserFormBase
        Dim oForm As UserFormBaseClass
        Dim hdlSample As ObjectHandle
        hdlSample = Activator.CreateInstance(m_AssemblyName, m_Namespace + "." + pFormName)

        If hdlSample Is Nothing Then Return Nothing
        oForm = CType(hdlSample.Unwrap(), UserFormBase)

        oForm.InitBase(Me)

        oForm.Show()
        'oForm.BusObjectInfo
        CreateSonForm = oForm
    End Function

    Public Property aUDOList() As SboUDOs Implements ISboAddon.aUDOList
        Get
            If m_UDOlist Is Nothing Then
                Try
                    Dim i As Integer = 0
                    'Dim field As enFieldNamesType
                    Dim orecset As SAPbobsCOM.Recordset
                    orecset = SboCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    orecset.DoQuery("select * from oudo WHERE TableName Like '" + m_tablePrefix + "%' ")

                    m_UDOlist = New SboUDOs
                    orecset.MoveFirst()
                    While Not orecset.EoF
                        i += 1
                        If CType(orecset.Fields.Item("TableName").Value, String).StartsWith(m_tablePrefix) Then
                            m_UDOlist.Add(New enUDONamesType(orecset.Fields.Item("Code").Value, orecset.Fields.Item("TableName").Value))
                        End If
                        orecset.MoveNext()
                    End While
                Catch ex As Exception
                    m_fieldlist = Nothing
                    SboApplication.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Long, True)
                End Try
            End If
            Return m_UDOlist
        End Get
        Set(ByVal Value As SboUDOs)
            m_fieldlist = Nothing
        End Set
    End Property

    Public Property UseDI() As Boolean Implements ISboAddon.UseDI
        Get
            Return m_useDI
        End Get
        Set(ByVal value As Boolean)
            m_useDI = value
        End Set
    End Property

    Public Property PermissionPrefix() As String Implements ISboAddon.PermissionPrefix
        Get
            Return m_permissionPrefix
        End Get
        Set(ByVal value As String)
            m_permissionPrefix = value
        End Set
    End Property

    Public Property TablePrefix() As String Implements ISboAddon.TablePrefix
        Get
            Return m_tablePrefix
        End Get
        Set(ByVal value As String)
            m_tablePrefix = value
        End Set
    End Property

    Public Property MenuXMLFileName() As String Implements ISboAddon.MenuXMLFileName
        Get
            Return m_MenuXMLFileName
        End Get
        Set(ByVal value As String)
            m_MenuXMLFileName = value
        End Set
    End Property

    Public Property usePermissions() As Boolean Implements ISboAddon.usePermissions
        Get
            Return m_PermUse
        End Get
        Set(ByVal value As Boolean)
            m_PermUse = value
        End Set
    End Property


#End Region

#Region "MustOverrides"
    'Public MustOverride Sub Handle_SBO_ItemEvent(ByVal FormUID As String, ByRef pVal As SBO_ItemEvent, ByRef BubbleEvent As Boolean) Handles ItemEvent
    'Public MustOverride Sub Handle_SBO_AppEvent(ByVal EventType As SBO_AppEvent)
    'Public MustOverride Sub Handle_SBO_MenuEvent(ByRef pVal As SBO_MenuEvent, ByRef BubbleEvent As Boolean) Handles MenuEvent

    'Public MustOverride Sub WriteLog(ByVal strLog As String)
    'Public MustOverride Sub WriteLog(ByVal strLog As String, ByVal strFileName As String)

#End Region

#Region "Methods"

    Public Function getUserTables(ByVal tableName As String) As SAPbobsCOM.UserTable Implements ISboAddon.getUserTables
        Dim oTable As enTableNamesType, table As SAPbobsCOM.UserTable
        table = Nothing
        For Each oTable In Me.aTableList
            If oTable.tableName = tableName.Substring(1) Then
                table = SboCompany.UserTables.Item(oTable.tableIndex - 1)
                Exit For
            End If
        Next
        Return table
    End Function

    Private Function checkTables(ByVal ds As Tables) As Boolean Implements ISboAddon.checkTables
        Dim strName As String
        Dim oTable As enTableNamesType
        Try
            For Each Me.Tablerow In ds.usertable
                For Each oTable In Me.aTableList
                    If Tablerow.Secondary <> -1 Then
                        checkTables = False
                        strName = Tablerow.TableName
                        If strName = oTable.tableName Then
                            checkTables = True
                            Exit For
                        End If
                    End If
                Next
                If checkTables = False Then Exit Function
            Next
        Catch ex As Exception

        End Try

    End Function

    Public Function InitialiseTables() As Boolean Implements ISboAddon.InitialiseTables
        Try
            'If Not InitTables Then Return True
            Dim file As System.IO.Stream = Me.getAppResource("AddonTables.xml")
            Dim ds As New Tables
            ds.ReadXml(file)
            'InitialiseTables = False
            If Me.aTableList Is Nothing Then
                Return False
            End If
            If Me.aFieldlist Is Nothing Then
                Return False
            End If
            If Me.aUDOList Is Nothing Then
                Return False
            End If
            Dim i As Integer = 0, strName As String, strName1 As String
            Dim oTable As enTableNamesType
            Dim ofield As enFieldNamesType
            Dim oUDO As enUDONamesType
            For Each Tablerow In ds.usertable
                For Each oTable In Me.aTableList
                    If Tablerow.Secondary <> -1 Then
                        strName = Tablerow.TableName
                        If strName.ToUpper = oTable.tableName.ToUpper Then
                            Tablerow.Created = "Y"
                            Tablerow.AcceptChanges()
                            Exit For
                        End If
                    End If
                Next
            Next

            For Each FieldRow In ds.TableFields
                For Each ofield In Me.aFieldlist
                    strName1 = FieldRow.Name.Trim
                    strName = FieldRow.usertableRow.TableName.Trim
                    If (strName1.Trim.ToUpper = ofield.FieldName.Trim.ToUpper And strName.ToUpper = ofield.TableName.Trim.ToUpper) Then
                        FieldRow.Created = "Y"
                        FieldRow.AcceptChanges()
                        Exit For
                    End If
                Next
            Next
            For Each objectRow In ds.UserObjects
                For Each oUDO In Me.aUDOList
                    strName1 = objectRow.Code.Trim
                    strName = objectRow.usertableRow.TableName.Trim
                    If (strName1.ToUpper.Trim = oUDO.UDOName.ToUpper.Trim And strName.ToUpper.Trim = oUDO.TableName.ToUpper.Trim) Then
                        objectRow.Created = "Y"
                        objectRow.AcceptChanges()
                        Exit For
                    End If
                Next
            Next
            If Not Me.createTables(ds) Then
                Return False
            End If
            InitialiseTables = True
        Catch ex As Exception
            MsgBox(ex.ToString)
        Finally
            SboApplication.MetadataAutoRefresh = True

        End Try

    End Function

    Protected Function createTables(ByVal ds As Tables) As Boolean Implements ISboAddon.createTables
        createTables = False
        GC.Collect()
        Try
            Dim i As Integer, y As Integer = 0, z As Integer = 0
            Dim strMsg As String
            Dim foundRows As DataRow() = ds.usertable.Select("Created = 'N'")
            y = foundRows.Length
            If y > 0 Then

                strMsg = "Creating user tables... This may take a few minutes"
                SboApplication.SetStatusBarMessage(strMsg, SAPbouiCOM.BoMessageTime.bmt_Short, False)
                SboApplication.MetadataAutoRefresh = False
                oProgBar = SboApplication.StatusBar.CreateProgressBar("Oryx01", y, False)
                oProgBar.Value = 1
                oUserTablesMD = Me.SboCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables)
                For i = -1 To 2
                    For Each Me.Tablerow In ds.usertable
                        If Tablerow.Created = "N" Then
                            If Tablerow.Secondary = i Then
                                If Tablerow.Secondary <> -1 Then
                                    'Update Progress Bar
                                    strMsg = "Creating Table - " + Tablerow.TableName
                                    oProgBar.Text = strMsg

                                    oUserTablesMD = Me.SboCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables)
                                    oUserTablesMD.TableName = Tablerow.TableName
                                    oUserTablesMD.TableDescription = Tablerow.TableDescription
                                    oUserTablesMD.TableType = Tablerow.Tabletype
                                    If oUserTablesMD.Add() <> 0 Then
                                        SboCompany.GetLastError(errCode, errMsg)
                                        SboApplication.MessageBox(errMsg, 1)
                                        Exit Function
                                    Else
                                        z += 1
                                        oProgBar.Value = z
                                        'aTablesCreated.Add(Tablerow.tablename.Trim)
                                    End If
                                End If
                                oUserTablesMD = Nothing
                                GC.Collect()
                                createTables = True
                            End If
                        End If
                    Next
                Next
                oProgBar.Stop()
                oProgBar = Nothing
                strMsg = "Tables Created Successfully"
                SboApplication.SetStatusBarMessage(strMsg, SAPbouiCOM.BoMessageTime.bmt_Short, False)

                oUserTablesMD = Nothing
                GC.Collect()
            End If
            If Not AddFields(ds) Then
                SboApplication.MessageBox("Error creating fields")
                Return False
            End If
            If Not AddUDO(ds) Then
                SboApplication.MessageBox("Error creating User defined objects")
                Return False
            End If
            Return True
        Catch ex As Exception
            SboApplication.MessageBox("Table: " & Tablerow.TableName & " " & ex.Message, 1)
            createTables = False
        Finally
            oUserFieldsMD = Nothing
            oUserTablesMD = Nothing
            oUserObjectMD = Nothing
            oUserKeysMD = Nothing
            GC.Collect()
            aTablesCreated = Nothing
            If Not oProgBar Is Nothing Then
                oProgBar.Stop()
                oProgBar = Nothing
            End If
        End Try
    End Function

    Private Function AddFields(ByVal ds As Tables) As Boolean Implements ISboAddon.AddFields
        Dim strFldName As String = ""
        Try
            Dim i As Integer, y As Integer = 0, z As Integer = 0
            Dim fieldRows As Tables.TableFieldsRow()
            Dim strMsg As String
            Dim foundRows As DataRow() = ds.TableFields.Select("Created = 'N'")
            y = foundRows.Length
            If y <= 0 Then
                AddFields = True
                Exit Function
            End If
            strMsg = "Creating user Fields..."
            SboApplication.SetStatusBarMessage(strMsg, SAPbouiCOM.BoMessageTime.bmt_Short, False)
            SboApplication.MetadataAutoRefresh = False
            oProgBar = SboApplication.StatusBar.CreateProgressBar("Oryx01", y, False)
            oProgBar.Value = 1
            For i = -1 To 2
                For Each Me.Tablerow In ds.usertable
                    If Tablerow.Secondary = i Then
                        fieldRows = Tablerow.GetTableFieldsRows()
                        y = fieldRows.Length
                        If fieldRows.Length > 0 Then
                            For Each FieldRow In fieldRows
                                If FieldRow.Created = "N" And FieldRow.Name <> "" Then
                                    '// Setting the Field's mandatory properties
                                    oUserFieldsMD = SboCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)
                                    oUserFieldsMD.TableName = FieldRow.usertableRow.TableName
                                    oUserFieldsMD.Name = FieldRow.Name
                                    strFldName = FieldRow.Name + " in " + FieldRow.usertableRow.TableName
                                    oUserFieldsMD.Description = FieldRow.Description
                                    oUserFieldsMD.Type = FieldRow.Type

                                    'If Not FieldRow.IsLinkedTableNull Then oUserFieldsMD.LinkedTable = FieldRow.LinkedTable
                                    Select Case FieldRow.Type
                                        Case "0"
                                            oUserFieldsMD.Size = FieldRow.Size
                                            'If Not FieldRow.IsEditTypeNull Then oUserFieldsMD.SubType = FieldRow.EditType
                                            oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
                                        Case "2"
                                            oUserFieldsMD.EditSize = FieldRow.EditSize
                                        Case "4"
                                            'oUserFieldsMD.SubType = FieldRow.EditType
                                        Case "3"
                                            oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
                                    End Select
                                    If Not FieldRow.IsEditTypeNull Then
                                        If FieldRow.EditType <> "" Then
                                            Select Case FieldRow.EditType
                                                Case "S" '- Amount 83
                                                    oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_Sum
                                                Case "R" '- Rate 82	
                                                    oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_Rate
                                                Case "Q" '- Quantity 81
                                                    oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_Quantity
                                                Case "P" '- Price 80
                                                    oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_Price
                                                Case "%" '- Percentage 37
                                                    oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_Percentage
                                                Case "M" '- Measurement 77
                                                    oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_Measurement
                                                Case "B" '- Link 66
                                                    oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_Link
                                                Case "I" '- Image 73
                                                    oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_Image
                                                Case "?" '- Address 63
                                                    oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_Address
                                                Case "T" '- Address 63
                                                    oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_Time
                                                Case "#" '- Address 63
                                                    oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_Phone
                                            End Select
                                        End If
                                    End If

                                    If Not FieldRow.IsDefaultValueNull Then oUserFieldsMD.DefaultValue = FieldRow.DefaultValue
                                    oUserFieldsMD.Mandatory = IIf(FieldRow.Mandatory = "Y", 1, 0)
                                    'If Not FieldRow.IsLinkedTableNull Then
                                    '    oUserFieldsMD.LinkedTable = FieldRow.LinkedTable
                                    'End If
                                    'oUserFieldsMD.EditSize = FieldRow.EditSiz
                                    Dim count As Integer = 1
                                    For Each validRow In FieldRow.GetValidValuesRows
                                        If count <> 1 Then oUserFieldsMD.ValidValues.SetCurrentLine(count - 1)
                                        count += 1
                                        oUserFieldsMD.ValidValues.Value = validRow.Value
                                        oUserFieldsMD.ValidValues.Description = validRow.Description
                                        oUserFieldsMD.ValidValues.Add()
                                    Next

                                    '// Adding the Field to the Table
                                    If oUserFieldsMD.Add <> 0 Then
                                        SboCompany.GetLastError(errCode, errMsg)
                                        strMsg = errMsg + " - " + FieldRow.usertableRow.TableName + " - " + FieldRow.Name
                                        'MsgBox(errMsg + " - " + FieldRow.usertableRow.tablename + " - " + FieldRow.Name, 1)
                                        SboApplication.SetStatusBarMessage(strMsg, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                        'If errCode = -2035 Then
                                        '    Continue For
                                        'Else
                                        '    Return False
                                        'End If
                                        Return False
                                    Else
                                        strMsg = "Field - " + FieldRow.Name + " in " + FieldRow.usertableRow.TableName + " Created "
                                        z += 1
                                        'strMsg += strMsg + " - " + ((z / y) * 100).ToString + " completed"
                                        oProgBar.Text = strMsg
                                        oProgBar.Value = z
                                        'm_frm.backgroundWorker.ReportProgress(100, str)
                                    End If
                                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldsMD)
                                    oUserFieldsMD = Nothing
                                    ''GC.WaitForPendingFinalizers()
                                    GC.Collect()
                                End If
                            Next
                        End If
                    End If
                Next
            Next
            oProgBar.Stop()
            oProgBar = Nothing
            strMsg = "Fields Created Successfully"
            SboApplication.SetStatusBarMessage(strMsg, SAPbouiCOM.BoMessageTime.bmt_Short, False)
            Return True
        Catch ex As Exception
            oUserFieldsMD = Nothing
            SboApplication.MessageBox(ex.ToString)
            GC.Collect()
            Return False
        Finally
            If Not oProgBar Is Nothing Then
                oProgBar.Stop()
                oProgBar = Nothing
            End If
            GC.Collect()
        End Try
    End Function

    Private Function AddUDO(ByVal ds As Tables) As Boolean Implements ISboAddon.AddUDO
        Try
            Dim i As Integer, y As Integer = 0, z As Integer = 0
            Dim strMsg As String
            Dim foundRows As DataRow() = ds.UserObjects.Select("Created = 'N'")
            y = foundRows.Length
            If y <= 0 Then
                AddUDO = True
                Exit Function
            End If
            SboApplication.MetadataAutoRefresh = False
            strMsg = "Creating user defined objects..."
            SboApplication.SetStatusBarMessage(strMsg, SAPbouiCOM.BoMessageTime.bmt_Short, False)
            oProgBar = SboApplication.StatusBar.CreateProgressBar("Oryx01", y, False)
            oProgBar.Value = 1
            Dim objectRows As Tables.UserObjectsRow()
            Dim ochildTable As Tables.UserObjectsChildRow
            For i = -1 To 2
                For Each Me.Tablerow In ds.usertable
                    If Tablerow.Secondary = i Then
                        objectRows = Tablerow.GetUserObjectsRows()
                        If objectRows.Length > 0 Then
                            y = objectRows.Length
                            For Each Me.objectRow In objectRows

                                If objectRow.Created = "N" Then

                                    If Not oUserObjectMD Is Nothing Then
                                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserObjectMD)
                                        oUserObjectMD = Nothing
                                        GC.Collect()
                                    End If
                                    oUserObjectMD = SboCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)
                                    oUserObjectMD.CanCancel = IIf(objectRow.CanCancel = "Y", 1, 0)
                                    oUserObjectMD.CanClose = IIf(objectRow.CanClose = "Y", 1, 0)
                                    oUserObjectMD.CanCreateDefaultForm = IIf(objectRow.CanCreateDefaultForm = "Y", 1, 0)
                                    'oUserObjectMD.FormColumns
                                    oUserObjectMD.CanDelete = IIf(objectRow.CanDelete = "Y", 1, 0)
                                    oUserObjectMD.CanFind = IIf(objectRow.CanFind = "Y", 1, 0)
                                    If objectRow.ObjectType = 1 Then
                                        oUserObjectMD.FindColumns.ColumnAlias = "Code"
                                        oUserObjectMD.FindColumns.Add()
                                        oUserObjectMD.FindColumns.SetCurrentLine(1)
                                        oUserObjectMD.FindColumns.ColumnAlias = "Name"
                                    End If

                                    If objectRow.ObjectType = 3 Then
                                        oUserObjectMD.FindColumns.ColumnAlias = "DocNum"
                                        oUserObjectMD.FindColumns.Add()
                                        oUserObjectMD.FindColumns.SetCurrentLine(1)
                                        oUserObjectMD.FindColumns.ColumnAlias = "DocEntry"
                                    End If

                                    'If objectRow.CanLog = "Y" Then
                                    '    oUserObjectMD.CanLog = IIf(objectRow.CanLog = "Y", 1, 0)
                                    '    oUserObjectMD.LogTableName = objectRow.LogTableName
                                    'End If

                                    oUserObjectMD.CanYearTransfer = IIf(objectRow.CanYearTransfer = "Y", 1, 0)
                                    Dim count As Integer = 1
                                    For Each ochildTable In objectRow.GetUserObjectsChildRows
                                        If count <> 1 Then oUserObjectMD.ChildTables.SetCurrentLine(count - 1)
                                        oUserObjectMD.ChildTables.TableName = ochildTable.TableName
                                        'oUserObjectMD.ChildTables.LogTableName = ochildTable.LogName
                                        oUserObjectMD.ChildTables.Add()
                                        count += 1
                                        'ochildTable.ExtensionName = ""
                                    Next
                                    oUserObjectMD.ManageSeries = IIf(objectRow.ManageSeries = "Y", 1, 0)

                                    oUserObjectMD.Code = objectRow.Code
                                    oUserObjectMD.Name = objectRow.Name
                                    oUserObjectMD.ObjectType = objectRow.ObjectType
                                    oUserObjectMD.TableName = objectRow.usertableRow.TableName

                                    If oUserObjectMD.Add() <> 0 Then
                                        Dim ErrMsg As String = ""
                                        Dim ErrCode As Long = 0
                                        SboCompany.GetLastError(ErrCode, ErrMsg)
                                        strMsg = ErrMsg + " - " + objectRow.usertableRow.TableName + " - " + objectRow.Name
                                        'MsgBox(errMsg + " - " + FieldRow.usertableRow.tablename + " - " + FieldRow.Name, 1)
                                        SboApplication.SetStatusBarMessage(strMsg, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                        'oUserObjectMD = Nothing
                                        GC.Collect()
                                        'Return False
                                    Else
                                        strMsg = "Object - " + objectRow.Name + " in " + objectRow.usertableRow.TableName + " Created "
                                        z += 1
                                        oProgBar.Text = strMsg
                                        oProgBar.Value = z
                                    End If
                                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserObjectMD)
                                End If
                                oUserObjectMD = SboCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)
                            Next
                        End If
                        oUserObjectMD = Nothing
                        GC.Collect()
                    End If
                Next
            Next
            oProgBar.Stop()
            oProgBar = Nothing
            strMsg = "Objects Created Successfully"
            SboApplication.SetStatusBarMessage(strMsg, SAPbouiCOM.BoMessageTime.bmt_Short, False)
            Return True
        Catch ex As Exception
            SboApplication.MessageBox(ex.ToString)
            GC.Collect()
            Return False
        Finally
            If Not oProgBar Is Nothing Then
                oProgBar.Stop()
                oProgBar = Nothing
            End If
        End Try


    End Function


#End Region

#Region "miscellaneous"
    Public Sub SendAddonBusy() Implements ISboAddon.SendAddonBusy
        m_SboApplication.StatusBar.SetText("AddOn busy", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_None)
    End Sub

    Private Sub FreezeMenu(ByVal State As Boolean)
        Dim uiFrmMenu As SAPbouiCOM.Form
        'sometimes it is necessary to freeze the menu
        uiFrmMenu = m_SboApplication.Forms.GetFormByTypeAndCount(enSAPFormTypes.sapBaseMenu, 1)
        uiFrmMenu.Freeze(State)

        If State = False Then
            uiFrmMenu.Update()
            uiFrmMenu.Refresh()
        End If
    End Sub

    Public Sub setFilters(ByVal ofilters As SAPbouiCOM.EventFilters)
        Try
            m_SboApplication.SetFilter(ofilters)
        Catch e As Exception

        End Try

    End Sub

    Public Function getAppResource(ByVal filename As String) As System.IO.Stream
        Dim thisExe As System.Reflection.Assembly
        thisExe = System.Reflection.Assembly.GetCallingAssembly()

        Dim file As System.IO.Stream
        If Me.UsePhysicalFiles Then
            file = New FileStream(Me.StartupPath + "\" + filename, FileMode.Open, FileAccess.Read)
        Else
            file = Me.AppAssembly.GetManifestResourceStream(AppNameSpace + "." + filename)
            If file Is Nothing Then
                thisExe = System.Reflection.Assembly.GetExecutingAssembly()
                file = thisExe.GetManifestResourceStream("SBO.SboAddOnBase." + filename)
            End If
        End If
        Return file
    End Function


    Public Sub ReleaseObject(ByVal o As Object)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(o)
    End Sub

   

    Public Sub AddForm(aForm As Object, strId As String)
        Try
            Dim userFormBaseClass As UserFormBaseClass = CType(aForm, UserFormBaseClass)
            Me.SBOForms.Add(userFormBaseClass, userFormBaseClass.UniqueID, Nothing, Nothing)
        Catch ex As Exception
            
        End Try
    End Sub
#End Region

#Region " Other Functions"

    'Public Function NextSequence(ByVal strTable As String) As String
    '    Dim oTable As SAPbobsCOM.UserTable
    '    oTable = Me.getUserTables("@OWA_SEQUENCES")
    '    Dim ret As String = ""

    '    If oTable.GetByKey(strTable) Then

    '        ret = oTable.UserFields.Fields.Item("U_NextNum").Value
    '    End If
    '    Return ret
    'End Function

    'Public Function UpdateSequence(ByVal strTable As String) As Boolean
    '    Dim oTable As SAPbobsCOM.UserTable, NextSequence As Integer
    '    oTable = Me.getUserTables("@OWA_SEQUENCES")

    '    If oTable.GetByKey(strTable) Then
    '        NextSequence = CInt(oTable.UserFields.Fields.Item("U_NextNum").Value) + 1
    '        oTable.UserFields.Fields.Item("U_NextNum").Value = NextSequence.ToString
    '        If oTable.Update() Then
    '            SboCompany.GetLastError(errCode, errMsg)
    '            SboApplication.SetStatusBarMessage(errMsg, SAPbouiCOM.BoMessageTime.bmt_Short)
    '            Return False
    '        End If
    '    End If
    '    Return True
    'End Function

    Public Function GenerateRandomString(ByVal intLenghtOfString As Integer) As String
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
#End Region

#Region "Functions "
    Public Sub BuildMenus(ByVal xmlFileName As String)
        Dim formCmdCenter As SAPbouiCOM.Form
        Try

            'Get reference to the Command Center form
            formCmdCenter = m_SboApplication.Forms.GetFormByTypeAndCount(169, 1)
            'Freeze it                                                              

            Try ' If the manu already exists this code will fail
                formCmdCenter.Freeze(True)
                Dim oXMLDoc1 As XmlDocument

                oXMLDoc1 = New XmlDocument

                'Get the file 
                Dim file As System.IO.Stream = Me.getAppResource(xmlFileName)

                '// load the content of the XML File
                oXMLDoc1.Load(file)

                '// load the form to the SBO application in one batch
                m_SboApplication.LoadBatchActions(oXMLDoc1.InnerXml)

                'get info on success or faliure
                Dim strBatchResults As String
                strBatchResults = Me.m_SboApplication.GetLastBatchResults()
                strBatchResults = strBatchResults + ""


                ' // Set the menu's icon (18x18 pixel, background color = 192.220.192 (for transparancy)
                Dim oMenuItem As SAPbouiCOM.MenuItem = m_SboApplication.Menus.Item(oXMLDoc1.SelectSingleNode("Application/Menus/action/Menu/@UniqueID").Value)
                'oMenuItem.Image = Me.StartupPath & "\" & m_MenuImageFile


            Catch er As Exception ' Menu already exists
                'WriteLog(er.ToString)
                m_SboApplication.MessageBox(er.Message)
            End Try

            formCmdCenter.Freeze(False)
            formCmdCenter.Update()

        Catch e As System.Exception
            'WriteLog(e.ToString)
            Me.m_SboApplication.MessageBox(e.ToString, 1)
            Me.m_SboApplication.StatusBar.SetText("Failed to add new custom Menu", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        Finally
            formCmdCenter.Freeze(False)
            formCmdCenter.Update()
        End Try


    End Sub

    Protected Function initialisePermissions() As Boolean
        Try
            'If Not initPerm Then Return True
            Dim file As System.IO.Stream = Me.getAppResource("Permissions.xml")
            'Dim ds As New Permissions
            'ds.ReadXml(file)
            GC.Collect()
            Dim strMsg As String = "Initialising user permissions"
            Me.m_SboApplication.StatusBar.SetText(strMsg, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_None)
            Dim ret As Boolean = True
            'Get the file 
            m_permissions = New Permissions

            If IsNothing(file) Then
                m_permissions = Nothing
            Else
                Try
                    m_permissions.ReadXml(file)
                Catch ex As DataException
                    MsgBox(ex.ToString)
                Catch ex As XmlException
                    MsgBox(ex.ToString)
                End Try

                'm_permissions.ReadXml()
            End If
            'Check if permissions have been created
            'Get the Userpermission Tree Object
            Dim mUserPerm As SAPbobsCOM.UserPermissionTree
            mUserPerm = SboCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserPermissionTree)
            Dim row As Permissions.PermRow
            Dim strSQL As String = "Select * from oupt where absid like '" + m_permissionPrefix + "%' "
            Dim orecset As SAPbobsCOM.Recordset
            orecset = SboCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            orecset.DoQuery(strSQL)
            If orecset.RecordCount > 0 Then
                mUserPerm.Browser.Recordset = orecset
                mUserPerm.Browser.MoveFirst()

                While Not mUserPerm.Browser.EoF
                    For Each row In m_permissions.Perm
                        If mUserPerm.PermissionID = row.PermissionID And row.IsItem = "N" Then
                            Dim foundRows() As DataRow
                            foundRows = Menuds.Tables(0).Select(" PermCategory = '" + row.PermissionID + "'")
                            If foundRows.GetUpperBound(0) < mUserPerm.UserPermissionForms.Count Then
                                row.Created = "Y"
                            End If
                        Else
                            If mUserPerm.PermissionID = row.PermissionID And row.IsItem = "Y" Then
                                row.Created = "Y"
                            End If
                        End If
                    Next
                    mUserPerm.Browser.MoveNext()
                End While
            End If
            mUserPerm = Nothing

            'Create the permissions in SBO
            For Each row In m_permissions.Perm
                Dim foundRows() As DataRow
                Dim i As Integer

                mUserPerm = SboCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserPermissionTree)

                If row.Created = "N" Then
                    strMsg = "Creating permission " + row.PermName
                    Me.m_SboApplication.StatusBar.SetText(strMsg, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_None)
                    GC.Collect()
                    mUserPerm.PermissionID = row.PermissionID
                    mUserPerm.Name = row.PermName
                    mUserPerm.Options = row.PermType
                    If Not row.FatherID = "" Then
                        mUserPerm.ParentID = row.FatherID
                    End If
                    foundRows = Menuds.Tables(0).Select(" PermCategory = '" + row.PermissionID + "'")
                    If foundRows.Length >= 1 Then

                        ' Print column 0 of each returned row.
                        For i = 0 To foundRows.Length - 1
                            mUserPerm.UserPermissionForms.SetCurrentLine(i)
                            mUserPerm.UserPermissionForms.FormType = foundRows(i)(2)
                            If i <> foundRows.Length - 1 Then mUserPerm.UserPermissionForms.Add()
                        Next i
                    End If
                    If mUserPerm.Add <> 0 Then
                        SboCompany.GetLastError(errCode, errMsg)
                        errMsg = "Failed to create permission - " + errMsg
                        Me.m_SboApplication.StatusBar.SetText(errMsg, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        ret = False
                        Exit For
                    End If

                Else
                    If Not mUserPerm.GetByKey(row.PermissionID) Then
                        Continue For
                    End If

                    foundRows = FormItemDs.Tables(0).Select(" PermCategory = '" + row.PermissionID + "'")
                    If foundRows Is Nothing Then
                        Continue For
                    End If
                    If foundRows.Length >= 1 Then
                        For i = 0 To foundRows.Length - 1
                            mUserPerm.UserPermissionForms.SetCurrentLine(i)
                            mUserPerm.UserPermissionForms.FormType = foundRows(i)(2)
                            mUserPerm.IsItem = SAPbobsCOM.BoYesNoEnum.tYES
                            If i <> foundRows.Length - 1 Then mUserPerm.UserPermissionForms.Add()
                        Next i

                        If mUserPerm.Update <> 0 Then
                            SboCompany.GetLastError(errCode, errMsg)
                            errMsg = "Failed to create permission - " + errMsg
                            Me.m_SboApplication.StatusBar.SetText(errMsg, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            ret = False
                            Exit For
                        End If
                    End If

                End If
            Next
            mUserPerm = Nothing
            GC.Collect()

            strMsg = "Permission check complete "
            Me.m_SboApplication.StatusBar.SetText(strMsg, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_None)
            Me.m_usePermissions = True
            Return ret
        Catch ex As Exception
            Me.m_SboApplication.StatusBar.SetText("Failed to initialise permissions", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End Try
    End Function

    Protected Function initialise() As Boolean
        Dim thisExe As System.Reflection.Assembly
        thisExe = System.Reflection.Assembly.GetCallingAssembly()
        Me.AppAssembly = thisExe
        initialise = False
        Dim ocom As SAPbobsCOM.Company
        ocom = SboCompany

        If Not Me.InitialiseTables() Then
            SboApplication.MessageBox("Error Initialising, addon will now exit")
            Return False
        End If
        If m_PermUse Then
            If Not Me.initialisePermissions() Then
                SboApplication.MessageBox("Error Initialising, addon will now exit")
                Return False
            End If
        End If
        
        BuildMenus("Menus.xml")
        initialise = True
    End Function

    Private Function ConnectToDIAPI() As Boolean

        Try
            'System.Runtime.InteropServices.Marshal.ReleaseComObject(m_SboCompany)

            Me.m_EventsBlocked = True
            'If m_SboApplication Is Nothing Then Me.ConnectToUIAPI()
            'Create company object
            m_SboCompany = New SAPbobsCOM.Company

            'get connection context
            'strCookie = m_SboCompany.GetContextCookie

            'retrieve connection context string via cookie
            'strConnectionContext = m_SboApplication.Company.GetConnectionContext(strCookie)

            'm_SboCompany.SetSboLoginContext(strConnectionContext)
            m_SboCompany = m_SboApplication.Company.GetDICompany
            If m_SboCompany Is Nothing Then
                'catch error
                Dim ErrCode As Integer
                Dim ErrMessage As String = ""

                m_SboCompany.GetLastError(ErrCode, ErrMessage)

                m_SboApplication.MessageBox("Fehler Nr.: " & ErrMessage & _
                vbCrLf & ErrMessage)

                m_EventsBlocked = False

                m_Connected = False

                Return False
            Else
                m_SboApplication.StatusBar.SetText("Connected to DI API.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

                m_EventsBlocked = False
                Return True
            End If

        Catch excE As Exception
            m_SboCompany = Nothing
            m_SboApplication.StatusBar.SetText(excE.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            'MsgBox(excE.ToString)
            Return False
        End Try
    End Function

    Private Function checkUserPermission(ByVal PermID As String) As Boolean
        Try

            'Load the Permissions Recordset
            Dim sboBOB As SAPbobsCOM.SBObob
            Dim sboRecordset As SAPbobsCOM.Recordset
            Dim ret As Boolean = True
            sboBOB = SboCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge)
            sboRecordset = sboBOB.GetSystemPermission(SboApplication.Company.UserName, PermID)

            Dim perm As SAPbobsCOM.BoPermission
            perm = sboRecordset.Fields.Item(0).Value
            If perm = SAPbobsCOM.BoPermission.boper_None Or perm = SAPbobsCOM.BoPermission.boper_Undefined Then
                ret = False
            End If
            Return ret
        Catch ex As Exception
            Return False
        End Try
    End Function

#End Region

#Region "Handle Events"

    
    Public Overridable Sub Handle_SBO_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) Handles m_SboApplication.MenuEvent
        Try
            If pVal.BeforeAction = True Then
                '// BEFORE MENU ACTION
                Select Case pVal.MenuUID
                    Case "DefCreate"
                        'createUserTables()
                    Case Else
                        ' Hand over navigation menu control to the forms

                End Select
            End If
            If pVal.BeforeAction = False Then
                Dim row As DataRow

                For Each row In Menuds.Tables(0).Rows
                    If row.Item("MenuUID") = pVal.MenuUID Then
                        If Me.m_usePermissions And Me.checkUserPermission(row.Item("PermCategory")) Then
                            Dim strMsg As String = "You do not have permission to view this form"
                            SboApplication.SetStatusBarMessage(strMsg, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            Exit Try
                        End If
                        If Menuds.Tables(0).Columns.Contains("UDO") Then
                            If row.Item("UDO") = "Y" Then
                                'SboApplication.ActivateMenuItem(row.Item("Name"))
                                SboApplication.OpenForm(SAPbouiCOM.BoFormObjectEnum.fo_UserDefinedObject, row.Item("Name"), "")
                            Else
                                Dim hdlSample As ObjectHandle
                                hdlSample = Activator.CreateInstance(m_AssemblyName, m_Namespace + "." + row.Item("Name"))
                                If hdlSample Is Nothing Then Exit Sub

                                Dim oForm As UserFormBaseClass
                                oForm = CType(hdlSample.Unwrap(), UserFormBase)
                                oForm.InitBase(Me)
                                'oForm.mstOryxFormType = row.Item("PermID")
                                'oForm.initbase()
                                oForm.Show()
                            End If

                        Else
                            Dim hdlSample As ObjectHandle
                            hdlSample = Activator.CreateInstance(m_AssemblyName, m_Namespace + "." + row.Item("Name"))
                            If hdlSample Is Nothing Then Exit Sub

                            Dim oForm As UserFormBaseClass
                            oForm = CType(hdlSample.Unwrap(), UserFormBase)
                            oForm.InitBase(Me)
                            'oForm.mstOryxFormType = row.Item("PermID")
                            'oForm.initbase()
                            oForm.Show()
                        End If






                    'Me.m_SBOForms.Add(oForm, oForm.UIAPIRawForm.UniqueID)
                    'Exit Sub
                    End If
                Next

                '    '// AFTER MENU ACTION 
                '    Dim CurrSboForm As sboForm
                '    Dim FormUID As String = m_SboApplication.Forms.ActiveForm.UniqueID
                '    For Each CurrSboForm In m_SBOForms
                '        If CurrSboForm.UniqueID = FormUID Then
                '            Try
                '                CurrSboForm.sboForm.Freeze(True)
                '                CurrSboForm.HANDLE_MENU_EVENTS(pVal, BubbleEvent)
                '            Catch ex As Exception
                '                Throw ex
                '            Finally
                '                CurrSboForm.sboForm.Freeze(False)
                '            End Try
                '            'CurrSboForm.HANDLE_MENU_EVENTS(pVal, BubbleEvent)
                '        End If
                '    Next
            End If

        Catch e As Exception
            'MsgBox(excE.ToString)
            SboApplication.SetStatusBarMessage(e.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Me.BlockEvents = False
        End Try
    End Sub
#End Region


    Public Sub New(ByVal StartUpPath As String, ByVal AddOnName As String)

        oApp = New Application

        m_StartUpPath = StartUpPath
        m_Name = AddOnName
        m_SboApplication = Application.SBO_Application
        If UseDI Then
            If ConnectToDIAPI() Then

                'initialise()
            End If
        End If

    End Sub

    Public Sub New()

    End Sub


#Region "MustOverrides"
    'Public MustOverride Sub Handle_SBO_ItemEvent(ByVal FormUID As String, ByRef pVal As SBO_ItemEvent, ByRef BubbleEvent As Boolean) Handles ItemEvent
    'Public MustOverride Sub Handle_SBO_AppEvent(ByVal EventType As SBO_AppEvent)
    'Public MustOverride Sub Handle_SBO_MenuEvent(ByRef pVal As SBO_MenuEvent, ByRef BubbleEvent As Boolean) Handles MenuEvent

    Public MustOverride Sub WriteLog(ByVal strLog As String)
    Public MustOverride Sub WriteLog(ByVal strLog As String, ByVal strFileName As String)

#End Region
End Class
