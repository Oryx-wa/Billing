
Option Strict Off
Option Explicit On

Imports OWA.SBO.OryxBillingLib10
Imports SAPbouiCOM
Imports SAPbouiCOM.Framework
Imports SBO.SboAddOnBase
Imports System
Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.Runtime.CompilerServices
Imports System.Threading



<FormAttribute("OryxBillingWinx10.AddOnForms_SBOEDUSchoolFees_b1f", "AddOnForms_SBOEDUSchoolFees.b1f")>
Friend Class SBOEDUSchoolFees
    Inherits UserFormBaseClass
    WithEvents StaticText0 As SAPbouiCOM.StaticText
    Private WithEvents EditText0 As SAPbouiCOM.EditText
    Private WithEvents EditText1 As SAPbouiCOM.EditText
    Private WithEvents txtTerm As SAPbouiCOM.EditText
    Private WithEvents Button0 As SAPbouiCOM.Button
    Private WithEvents Button1 As SAPbouiCOM.Button
    Private WithEvents dgschfees As SAPbouiCOM.Matrix
    Private WithEvents keystage As SAPbouiCOM.EditText
    Private WithEvents txtyrCode As SAPbouiCOM.EditText
    Private WithEvents txtFee As SAPbouiCOM.Column
    Private WithEvents Button2 As SAPbouiCOM.Button
    Private WithEvents StaticText1 As SAPbouiCOM.StaticText
    Private WithEvents EditText2 As SAPbouiCOM.EditText
    Private WithEvents btnCopy As SAPbouiCOM.ButtonCombo

    Public Sub New()
        AddHandler MyBase.DataAddBefore, AddressOf Me.SBOEDUSchoolFees_DataAddBefore
        AddHandler MyBase.DataDeleteBefore, AddressOf Me.SBOEDUSchoolFees_DataDeleteBefore
        AddHandler MyBase.DataUpdateBefore, AddressOf Me.SBOEDUSchoolFees_DataUpdateBefore

    End Sub

    Protected Overrides Sub InitBase(pAddOn As SboAddon)
        MyBase.InitBase(pAddOn)
        Me.CreateObject(New SBOEDUFees(pAddOn, Me.UIAPIRawForm))
    End Sub

    Public Overrides Sub OnInitializeComponent()
        Me.StaticText0 = CType(Me.GetItem("Item_0").Specific, SAPbouiCOM.StaticText)
        Me.EditText1 = CType(Me.GetItem("txtyrCode").Specific, SAPbouiCOM.EditText)
        Me.Button1 = CType(Me.GetItem("2").Specific, SAPbouiCOM.Button)
        Me.Button0 = CType(Me.GetItem("1").Specific, SAPbouiCOM.Button)
        Me.keystage = CType(Me.GetItem("txtClass").Specific, SAPbouiCOM.EditText)
        Me.txtyrCode = CType(Me.GetItem("txtyrCode").Specific, SAPbouiCOM.EditText)
        Me.EditText0 = CType(Me.GetItem("docentry").Specific, SAPbouiCOM.EditText)
        Me.dgschfees = CType(Me.GetItem("dgschfes").Specific, SAPbouiCOM.Matrix)
        Me.txtFee = CType(Me.GetItem("dgschfes").Specific, SAPbouiCOM.Matrix).Columns.Item("colFee")
        Me.btnCopy = CType(Me.GetItem("btnCopy").Specific, SAPbouiCOM.ButtonCombo)
        Me.Button2 = CType(Me.GetItem("btnFees").Specific, SAPbouiCOM.Button)
        Me.StaticText1 = CType(Me.GetItem("Item_1").Specific, SAPbouiCOM.StaticText)
        Me.EditText2 = CType(Me.GetItem("txtTotal").Specific, SAPbouiCOM.EditText)
        Me.EditText3 = CType(Me.GetItem("Item_4").Specific, SAPbouiCOM.EditText)
        Me.StaticText2 = CType(Me.GetItem("Item_6").Specific, SAPbouiCOM.StaticText)
        Me.StaticText3 = CType(Me.GetItem("Item_7").Specific, SAPbouiCOM.StaticText)
        Me.StaticText4 = CType(Me.GetItem("Item_8").Specific, SAPbouiCOM.StaticText)
        Me.EditText4 = CType(Me.GetItem("Item_9").Specific, SAPbouiCOM.EditText)
        Me.StaticText5 = CType(Me.GetItem("Item_10").Specific, SAPbouiCOM.StaticText)
        Me.EditText5 = CType(Me.GetItem("Item_11").Specific, SAPbouiCOM.EditText)
        Me.StaticText6 = CType(Me.GetItem("Item_12").Specific, SAPbouiCOM.StaticText)
        Me.EditText6 = CType(Me.GetItem("Item_13").Specific, SAPbouiCOM.EditText)
        Me.StaticText7 = CType(Me.GetItem("Item_14").Specific, SAPbouiCOM.StaticText)
        Me.EditText7 = CType(Me.GetItem("Item_15").Specific, SAPbouiCOM.EditText)
        Me.StaticText8 = CType(Me.GetItem("Item_16").Specific, SAPbouiCOM.StaticText)
        Me.EditText8 = CType(Me.GetItem("Item_17").Specific, SAPbouiCOM.EditText)
        Me.StaticText9 = CType(Me.GetItem("Item_18").Specific, SAPbouiCOM.StaticText)
        Me.EditText9 = CType(Me.GetItem("Item_19").Specific, SAPbouiCOM.EditText)
        Me.StaticText10 = CType(Me.GetItem("Item_20").Specific, SAPbouiCOM.StaticText)
        Me.EditText10 = CType(Me.GetItem("Item_21").Specific, SAPbouiCOM.EditText)
        Me.StaticText11 = CType(Me.GetItem("Item_22").Specific, SAPbouiCOM.StaticText)
        Me.EditText11 = CType(Me.GetItem("Item_23").Specific, SAPbouiCOM.EditText)
        Me.StaticText12 = CType(Me.GetItem("Item_24").Specific, SAPbouiCOM.StaticText)
        Me.EditText12 = CType(Me.GetItem("Item_25").Specific, SAPbouiCOM.EditText)
        Me.StaticText14 = CType(Me.GetItem("Item_28").Specific, SAPbouiCOM.StaticText)
        Me.EditText14 = CType(Me.GetItem("Item_29").Specific, SAPbouiCOM.EditText)
        Me.OnCustomInitialize()

    End Sub

    Public Overrides Sub OnInitializeFormEvents()
        AddHandler DataAddBefore, AddressOf Me.SBOEDUSchoolFees_DataAddBefore
        AddHandler DataDeleteBefore, AddressOf Me.SBOEDUSchoolFees_DataDeleteBefore
        AddHandler DataUpdateBefore, AddressOf Me.SBOEDUSchoolFees_DataUpdateBefore
        AddHandler DataAddBefore, AddressOf Me.SBOEDUSchoolFees_DataAddBefore
        AddHandler DataDeleteBefore, AddressOf Me.SBOEDUSchoolFees_DataDeleteBefore
        AddHandler DataUpdateBefore, AddressOf Me.SBOEDUSchoolFees_DataUpdateBefore
        AddHandler DataAddBefore, AddressOf Me.SBOEDUSchoolFees_DataAddBefore
        AddHandler DataDeleteBefore, AddressOf Me.SBOEDUSchoolFees_DataDeleteBefore
        AddHandler DataUpdateBefore, AddressOf Me.SBOEDUSchoolFees_DataUpdateBefore
        AddHandler LoadAfter, AddressOf Me.Form_LoadAfter
        AddHandler DataAddBefore, AddressOf Me.SBOEDUSchoolFees_DataAddBefore
        AddHandler DataDeleteBefore, AddressOf Me.SBOEDUSchoolFees_DataDeleteBefore
        AddHandler DataUpdateBefore, AddressOf Me.SBOEDUSchoolFees_DataUpdateBefore
        AddHandler DataAddBefore, AddressOf Me.SBOEDUSchoolFees_DataAddBefore
        AddHandler DataDeleteBefore, AddressOf Me.SBOEDUSchoolFees_DataDeleteBefore
        AddHandler DataUpdateBefore, AddressOf Me.SBOEDUSchoolFees_DataUpdateBefore
        AddHandler DataAddBefore, AddressOf Me.SBOEDUSchoolFees_DataAddBefore
        AddHandler DataDeleteBefore, AddressOf Me.SBOEDUSchoolFees_DataDeleteBefore
        AddHandler DataUpdateBefore, AddressOf Me.SBOEDUSchoolFees_DataUpdateBefore
        AddHandler DataAddBefore, AddressOf Me.SBOEDUSchoolFees_DataAddBefore
        AddHandler DataDeleteBefore, AddressOf Me.SBOEDUSchoolFees_DataDeleteBefore
        AddHandler DataUpdateBefore, AddressOf Me.SBOEDUSchoolFees_DataUpdateBefore
        AddHandler DataAddBefore, AddressOf Me.SBOEDUSchoolFees_DataAddBefore
        AddHandler DataDeleteBefore, AddressOf Me.SBOEDUSchoolFees_DataDeleteBefore
        AddHandler DataUpdateBefore, AddressOf Me.SBOEDUSchoolFees_DataUpdateBefore
        AddHandler DataAddBefore, AddressOf Me.SBOEDUSchoolFees_DataAddBefore
        AddHandler DataDeleteBefore, AddressOf Me.SBOEDUSchoolFees_DataDeleteBefore
        AddHandler DataUpdateBefore, AddressOf Me.SBOEDUSchoolFees_DataUpdateBefore
        AddHandler DataUpdateAfter, AddressOf Me.Form_DataUpdateAfter
        AddHandler DataAddBefore, AddressOf Me.SBOEDUSchoolFees_DataAddBefore
        AddHandler DataDeleteBefore, AddressOf Me.SBOEDUSchoolFees_DataDeleteBefore
        AddHandler DataUpdateBefore, AddressOf Me.SBOEDUSchoolFees_DataUpdateBefore
        AddHandler DataAddBefore, AddressOf Me.SBOEDUSchoolFees_DataAddBefore
        AddHandler DataDeleteBefore, AddressOf Me.SBOEDUSchoolFees_DataDeleteBefore
        AddHandler DataUpdateBefore, AddressOf Me.SBOEDUSchoolFees_DataUpdateBefore
        AddHandler DataAddBefore, AddressOf Me.SBOEDUSchoolFees_DataAddBefore
        AddHandler DataDeleteBefore, AddressOf Me.SBOEDUSchoolFees_DataDeleteBefore
        AddHandler DataUpdateBefore, AddressOf Me.SBOEDUSchoolFees_DataUpdateBefore
        AddHandler DataAddBefore, AddressOf Me.SBOEDUSchoolFees_DataAddBefore
        AddHandler DataDeleteBefore, AddressOf Me.SBOEDUSchoolFees_DataDeleteBefore
        AddHandler DataUpdateBefore, AddressOf Me.SBOEDUSchoolFees_DataUpdateBefore
        AddHandler DataAddBefore, AddressOf Me.SBOEDUSchoolFees_DataAddBefore
        AddHandler DataDeleteBefore, AddressOf Me.SBOEDUSchoolFees_DataDeleteBefore
        AddHandler DataUpdateBefore, AddressOf Me.SBOEDUSchoolFees_DataUpdateBefore
        AddHandler DataAddBefore, AddressOf Me.SBOEDUSchoolFees_DataAddBefore
        AddHandler DataDeleteBefore, AddressOf Me.SBOEDUSchoolFees_DataDeleteBefore
        AddHandler DataUpdateBefore, AddressOf Me.SBOEDUSchoolFees_DataUpdateBefore

    End Sub

    Private Sub dgschfees_PressedBefore(sboObject As Object, pVal As SBOItemEventArg, ByRef BubbleEvent As Boolean) Handles dgschfees.PressedBefore
        Me.oForm.DataSources.DBDataSources.Item("@OWA_EDUSCHFEEROWS").Clear()
        Me.dgschfees = CType(Me.oForm.Items.Item("dgschfes").Specific, Matrix)
        ' The following expression was wrapped in a checked-expression
        Dim flag As Boolean = pVal.Row = Me.dgschfees.RowCount + 1
        Dim flag4 As Boolean = Me.dgschfees.RowCount = 0
        If flag Then
            Dim flag2 As Boolean = pVal.Row = 1
            If flag2 Then
                If flag4 Then
                    Me.dgschfees.AddRow(30, -1)
                Else
                    Me.dgschfees.AddRow(1, -1)
                End If

            Else
                Me.dgschfees.AddRow(1, Me.dgschfees.RowCount)
            End If
            Me.dgschfees.Columns.Item(1).Cells.Item(pVal.Row).Click(BoCellClickType.ct_Regular, 0)
        End If
    End Sub

    Private Sub keystage_ChooseFromListAfter(sboObject As Object, pVal As SBOItemEventArg) Handles dgschfees.ChooseFromListAfter
        Me.m_BaseObject.OnChooseFromListAfter(RuntimeHelpers.GetObjectValue(sboObject), pVal)
    End Sub

    Private Sub OnCustomInitialize()
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

    Private Sub ButtonCombo0_ComboSelectAfter(sboObject As Object, pVal As SBOItemEventArg)
        Me.m_BaseObject.OnComboSelectAfter(RuntimeHelpers.GetObjectValue(sboObject), pVal)
    End Sub

    Private Sub SBOEDUSchoolFees_DataAddBefore(ByRef pVal As BusinessObjectInfo, ByRef BubbleEvent As Boolean) Handles Me.DataAddBefore
        Me.m_BaseObject.OnDataAddBefore(pVal, BubbleEvent)
    End Sub

    Private Sub SBOEDUSchoolFees_DataDeleteBefore(ByRef pVal As BusinessObjectInfo, ByRef BubbleEvent As Boolean) Handles Me.DataDeleteBefore
        Me.m_BaseObject.OnDataDeleteBefore(pVal, BubbleEvent)
    End Sub

    Private Sub SBOEDUSchoolFees_DataUpdateBefore(ByRef pVal As BusinessObjectInfo, ByRef BubbleEvent As Boolean) Handles Me.DataUpdateBefore
        Me.m_BaseObject.OnDataUpdateBefore(pVal, BubbleEvent)
    End Sub

    Private Sub dgschfees_ValidateAfter(sboObject As Object, pVal As SBOItemEventArg) Handles dgschfees.ValidateAfter
        Me.m_BaseObject.OnItemValidateAfter(RuntimeHelpers.GetObjectValue(sboObject), pVal)
    End Sub
    Private WithEvents EditText3 As SAPbouiCOM.EditText
    Private WithEvents StaticText2 As SAPbouiCOM.StaticText
    Private WithEvents StaticText3 As SAPbouiCOM.StaticText
    Private WithEvents StaticText4 As SAPbouiCOM.StaticText
    Private WithEvents EditText4 As SAPbouiCOM.EditText
    Private WithEvents StaticText5 As SAPbouiCOM.StaticText
    Private WithEvents EditText5 As SAPbouiCOM.EditText
    Private WithEvents StaticText6 As SAPbouiCOM.StaticText
    Private WithEvents EditText6 As SAPbouiCOM.EditText
    Private WithEvents StaticText7 As SAPbouiCOM.StaticText
    Private WithEvents EditText7 As SAPbouiCOM.EditText
    Private WithEvents StaticText8 As SAPbouiCOM.StaticText
    Private WithEvents EditText8 As SAPbouiCOM.EditText
    Private WithEvents StaticText9 As StaticText
    Private WithEvents EditText9 As EditText

    Private Sub Form_LoadAfter(pVal As SBOItemEventArg)
        'Throw New System.NotImplementedException()

    End Sub

    Private WithEvents StaticText10 As StaticText
    Private WithEvents EditText10 As EditText
    Private WithEvents StaticText11 As StaticText
    Private WithEvents EditText11 As EditText
    Private WithEvents StaticText12 As StaticText
    Private WithEvents EditText12 As EditText

    Private Sub Form_DataUpdateAfter(ByRef pVal As BusinessObjectInfo)
        m_BaseObject.OnDataUpdateAfter(pVal)
    End Sub

    Private WithEvents StaticText14 As StaticText
    Private WithEvents EditText14 As EditText
End Class

