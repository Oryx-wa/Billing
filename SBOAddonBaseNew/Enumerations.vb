Option Explicit On 
'*********************************************************************
'* TECHDEMO ADDON CLASS DESCRIPTION                                  *
'* assembly name:       SboAddonBase                                 *
'* classname:           -                                            *
'* classtype:           -                                            *
'*********************************************************************
'* created by:          Lutz Morrien                                 *
'* company:             ocb GmbH, Ahaus                              *
'*                                                                   *
'* date of last change: 09-19-2004                                   *
'* last change by:      Lutz Morrien                                 *
'*                                                                   *
'*********************************************************************
'* Class description:                                                *
'* The AppConfig class is used to provide easy access to the         *
'* App.config file of the .Net application.                          *
'*                                                                   *
'*********************************************************************
'* list of changes and additions:                                    *
'*                                                                   *
'*********************************************************************

Public Enum enSAPFormModes As Integer
    cst_Ok = 1
    cst_Find = 0
    cst_Add = 3
    cst_Update = 2
    cst_View = 4
    cst_Print = 5
    cst_Edit = 6
End Enum

Public Enum enSBO_LoadFormTypes
    XmlFile = 0
    LogicOnly = 1
    GuiByCode = 3
End Enum

Public Enum enSboCheckBoxValues
    Y
    N
End Enum

Public Enum enSAPFormTypes As Integer
    sapSalesInvoice = 133
    sapBusinessPartner = 134             'Stammdaten Geschäftspartner
    sapCompanyDetails = 136              'Firmendaten
    sapGeneralSettings = 138
    sapSalesOrder = 139                  'Auftrag
    sapSalesDelivery = 140               'Lieferung
    sapPurchaseInvoice = 141
    sapPurchaseOrder = 142
    sapPurchaseGoodsReceived = 143
    sapDefineCurrencies = 148
    sapSalesQuotation = 149              'Angebot
    sapItemsManagement = 150             'Artikelverwaltung
    sapBaseMenu = 169                    'Basismenü
    sapDocumentNumbering = 172
    sapDefinePaymentTerms = 177
    sapSalesCreditMemo = 179             'Gutschrift
    sapSalesGoodsReturns = 180           'RetourenVerkauf
    sapPurchaseCreditMemo = 181          'GutschriftEinkauf
    sapPurchaseGoodsReturns = 182
    sapPrintReferences = 183
    sapDocumentsPrintingCriteria = 184
    sapDocumentSettings = 228
    sapGLAccountDetermination = 350
    sapChoosePeriod_End_Closing = 411
    sapDefineAdressFormat = 419
    sapDefineCommissionGroups = 664
    sapDefineSalesPerson = 666
    sapDefineTransactionCodes = 710
    sapDefineProjects = 711
    sapSelectTransactionJournal = 725
    sapChartOfAccounts = 804
    sapChooseCompany = 820
    sapDefineIndexes = 865
    sapDefineForeignCurrencyExchangeRates = 866
    sapDefineTaxGroups = 895
    sapDefineCompleteTransportCosts = 898
    sapChooseOpeningBalances = 923
    sapDefineCountries = 941
    sapAuthorizations = 951
    sapAutomaticSummaryWizard = 953
    sapMarketingDocumentDrafts = 3001
    sapBelegbearbeitung = 4665
    sapDefineUsers = 20700
    sapEmployee = 60100

    'Sbo TechDemo Forms
    'This part of the enumeration keeps track of all type numbers given to
    ' any new form in any addon you create
    STDWelcomeForm = 2000100000
    STDFormSimpleForm = 2000100001
    STDCreditMemoForm = 2000060004
    SDTFormToCome4 = 2000100003
    STDChecksForm = 2000060004
    STDPayRollForm = 2000060005
    STDFleetMgtForm = 2000060006
End Enum

Public Enum enSboFormTypes
    XmlFile = 0
    LogicOnly = 1
    GuiByCode = 3
End Enum



Public Enum enSAPMenuUIDs As Integer
    Belegbearbeitung = 5895
    BenutzerfelderAnzeigen = 6913
    ChooseCompany = 3329
    Drucken = 520
    Seitenansicht = 519
End Enum
Public Structure enUserSourceType
    Public Name As String
    Public SBODataType As SAPbouiCOM.BoDataType
    Public length As Integer
    Public dbColName As String
End Structure
Public Structure enMatrixSourceType
    Public Name As String
    Public SBODataType As SAPbouiCOM.BoDataType
    Public length As Integer
    Public dbColName As String
    Public MatColId As String
End Structure
Public Structure ConditionVals
    Public cAlias As String
    Public cValue As String
End Structure
'Public Structure enTableNamesType
'    Public tableName As String
'    Public tableIndex As Integer
'    Public Sub New(ByVal strTablename As String, ByVal intTableIndex As Integer)
'        Me.tableName = strTablename
'        Me.tableIndex = intTableIndex
'    End Sub
'End Structure

Public Structure enTableNamesType
    Public tableName As String
    Public tableIndex As Integer
    Public Sub New(ByVal strTablename As String, ByVal intTableIndex As Integer)
        Me.tableName = strTablename
        Me.tableIndex = intTableIndex
    End Sub
End Structure
Public Structure enFieldNamesType
    Public FieldName As String
    Public TableName As String
    Public Sub New(ByVal strFieldname As String, ByVal strTableName As String)
        Me.FieldName = strFieldname.Trim
        Me.TableName = strTableName.Trim
    End Sub
End Structure
Public Structure enUDONamesType
    Public UDOName As String
    Public TableName As String
    Public Sub New(ByVal strUDOname As String, ByVal strTableName As String)
        Me.UDOName = strUDOname.Trim
        Me.TableName = strTableName.Trim
    End Sub
End Structure


