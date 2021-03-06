
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




<FormAttribute("OryxBillingWinx10.AddOnForms_SBOEDUBillingWizard_b1f", "AddOnForms_SBOEDUBillingWizard.b1f")>
Friend Class SBOEDUBillingWizard
    Inherits UserFormBaseClass
    Private WithEvents StaticText0 As SAPbouiCOM.StaticText
    Private WithEvents StaticText6 As SAPbouiCOM.StaticText
    Private WithEvents Button0 As SAPbouiCOM.Button
    Private WithEvents Button1 As SAPbouiCOM.Button
    Private WithEvents Button2 As SAPbouiCOM.Button
    Private WithEvents Button3 As SAPbouiCOM.Button
    Private WithEvents Grid0 As SAPbouiCOM.Grid
    'Private WithEvents Matrix1 As SAPbouiCOM.Matrix
    Private WithEvents Button4 As SAPbouiCOM.Button
    Private WithEvents CheckBox1 As SAPbouiCOM.CheckBox
    'Private WithEvents Matrix2 As SAPbouiCOM.Matrix
    Private WithEvents Button5 As SAPbouiCOM.Button
    Private WithEvents chkMails As SAPbouiCOM.CheckBox
    Private WithEvents Button6 As SAPbouiCOM.Button
    Private WithEvents Button7 As SAPbouiCOM.Button
    Private WithEvents Button8 As SAPbouiCOM.Button
    Private WithEvents Button9 As SAPbouiCOM.Button
    Private WithEvents StaticText4 As SAPbouiCOM.StaticText
    Private WithEvents OptionBtn0 As SAPbouiCOM.OptionBtn
    Private WithEvents OptionBtn1 As SAPbouiCOM.OptionBtn
    Private WithEvents StaticText12 As SAPbouiCOM.StaticText
    Private WithEvents StaticText1 As SAPbouiCOM.StaticText
    Private WithEvents EditText3 As SAPbouiCOM.EditText
    Private WithEvents Button11 As SAPbouiCOM.Button
    Private WithEvents StaticText2 As SAPbouiCOM.StaticText
    Private WithEvents StaticText10 As SAPbouiCOM.StaticText
    Private WithEvents StaticText11 As SAPbouiCOM.StaticText
    Private WithEvents Button12 As SAPbouiCOM.Button
    Private WithEvents Grid1 As SAPbouiCOM.Grid
    Private WithEvents Grid3 As SAPbouiCOM.Grid
    Private WithEvents Grid4 As SAPbouiCOM.Grid
    Private WithEvents StaticText3 As SAPbouiCOM.StaticText
    Private WithEvents StaticText5 As SAPbouiCOM.StaticText
    Private WithEvents StaticText7 As SAPbouiCOM.StaticText
    Private WithEvents StaticText8 As SAPbouiCOM.StaticText
    Private WithEvents EditText0 As SAPbouiCOM.EditText
    Private WithEvents StaticText9 As SAPbouiCOM.StaticText
    Private WithEvents EditText1 As SAPbouiCOM.EditText
    Private WithEvents Button15 As SAPbouiCOM.Button
    Private WithEvents Matrix0 As SAPbouiCOM.Matrix
    Private WithEvents Button16 As SAPbouiCOM.Button
    Private WithEvents Button18 As SAPbouiCOM.Button
    Private WithEvents EditText2 As SAPbouiCOM.EditText
    Private WithEvents StaticText13 As SAPbouiCOM.StaticText
    Private WithEvents ButtonCombo0 As SAPbouiCOM.ButtonCombo
    Private WithEvents Button10 As SAPbouiCOM.Button
    Private WithEvents Grid2 As SAPbouiCOM.Grid
    Private WithEvents Grid5 As SAPbouiCOM.Grid
    Private WithEvents Button14 As SAPbouiCOM.Button
    Private WithEvents ButtonCombo2 As SAPbouiCOM.ButtonCombo
    Private WithEvents ButtonCombo4 As SAPbouiCOM.ButtonCombo
    Private WithEvents Grid6 As SAPbouiCOM.Grid

    Public Sub New()
        'SBOEDUBillingWizard.__ENCAddToList(Me)
    End Sub

    Protected Overrides Sub InitBase(pAddOn As SboAddon)
        MyBase.InitBase(pAddOn)
        Me.CreateObject(New SBOEDUBillWizard(pAddOn, Me.UIAPIRawForm))
    End Sub

    Public Overrides Sub OnInitializeComponent()
        Me.StaticText6 = CType(Me.GetItem("lblStep").Specific, SAPbouiCOM.StaticText)
        Me.Button0 = CType(Me.GetItem("btnFinish").Specific, SAPbouiCOM.Button)
        Me.Button1 = CType(Me.GetItem("2").Specific, SAPbouiCOM.Button)
        Me.Button2 = CType(Me.GetItem("btnBack").Specific, SAPbouiCOM.Button)
        Me.Button3 = CType(Me.GetItem("btnNext").Specific, SAPbouiCOM.Button)
        Me.Grid0 = CType(Me.GetItem("grdStud").Specific, SAPbouiCOM.Grid)
        Me.Button9 = CType(Me.GetItem("btnColAll").Specific, SAPbouiCOM.Button)
        Me.OptionBtn0 = CType(Me.GetItem("optNew").Specific, SAPbouiCOM.OptionBtn)
        Me.OptionBtn1 = CType(Me.GetItem("optUpdate").Specific, SAPbouiCOM.OptionBtn)
        Me.StaticText12 = CType(Me.GetItem("lblDesc_1").Specific, SAPbouiCOM.StaticText)
        Me.StaticText1 = CType(Me.GetItem("lblBatch").Specific, SAPbouiCOM.StaticText)
        Me.EditText3 = CType(Me.GetItem("txtDocNum").Specific, SAPbouiCOM.EditText)
        Me.Button11 = CType(Me.GetItem("btnSelect").Specific, SAPbouiCOM.Button)
        Me.StaticText10 = CType(Me.GetItem("lblTitle_1").Specific, SAPbouiCOM.StaticText)
        Me.StaticText11 = CType(Me.GetItem("lblTitle_2").Specific, SAPbouiCOM.StaticText)
        Me.Grid4 = CType(Me.GetItem("grdReview").Specific, SAPbouiCOM.Grid)
        Me.StaticText3 = CType(Me.GetItem("lblDesc_2").Specific, SAPbouiCOM.StaticText)
        Me.StaticText5 = CType(Me.GetItem("lblD_2_1").Specific, SAPbouiCOM.StaticText)
        Me.StaticText7 = CType(Me.GetItem("lblD_2_2").Specific, SAPbouiCOM.StaticText)
        Me.StaticText8 = CType(Me.GetItem("lblBDesc").Specific, SAPbouiCOM.StaticText)
        Me.EditText0 = CType(Me.GetItem("txtDesc").Specific, SAPbouiCOM.EditText)
        Me.StaticText9 = CType(Me.GetItem("lblSession").Specific, SAPbouiCOM.StaticText)
        Me.EditText1 = CType(Me.GetItem("txtSession").Specific, SAPbouiCOM.EditText)
        Me.Button15 = CType(Me.GetItem("btnStd").Specific, SAPbouiCOM.Button)
        Me.Button16 = CType(Me.GetItem("btnSaveStd").Specific, SAPbouiCOM.Button)
        Me.Button18 = CType(Me.GetItem("btnDelete").Specific, SAPbouiCOM.Button)
        Me.EditText2 = CType(Me.GetItem("txtCount").Specific, SAPbouiCOM.EditText)
        Me.StaticText13 = CType(Me.GetItem("Item_1").Specific, SAPbouiCOM.StaticText)
        Me.Button10 = CType(Me.GetItem("btnEmail").Specific, SAPbouiCOM.Button)
        Me.Grid2 = CType(Me.GetItem("grdEmail").Specific, SAPbouiCOM.Grid)
        Me.Grid5 = CType(Me.GetItem("grdInvoice").Specific, SAPbouiCOM.Grid)
        Me.Button14 = CType(Me.GetItem("btnPost").Specific, SAPbouiCOM.Button)
        Me.ButtonCombo2 = CType(Me.GetItem("cboJump").Specific, SAPbouiCOM.ButtonCombo)
        Me.Button8 = CType(Me.GetItem("btnExpAll").Specific, SAPbouiCOM.Button)
        Me.Button13 = CType(Me.GetItem("btnCancInv").Specific, SAPbouiCOM.Button)
        Me.OnCustomInitialize()

    End Sub

    Public Overrides Sub OnInitializeFormEvents()

    End Sub

    Private Sub OnCustomInitialize()
    End Sub

    Private Sub Button0_ClickAfter(sboObject As Object, pVal As SBOItemEventArg) Handles Button3.ClickAfter, Button0.ClickAfter, Button1.ClickAfter, Button2.ClickAfter, Button4.ClickAfter, Button5.ClickAfter, Button6.ClickAfter, Button7.ClickAfter, Button9.ClickAfter, Button8.ClickAfter, OptionBtn0.ClickAfter, OptionBtn1.ClickAfter, Button15.ClickAfter, Button16.ClickAfter, Button18.ClickAfter,
        ButtonCombo0.ClickAfter, Button10.ClickAfter, Button14.ClickAfter, Grid6.ClickAfter, Button13.ClickAfter
        Me.m_BaseObject.OnItemClickAfter(RuntimeHelpers.GetObjectValue(sboObject), pVal)
    End Sub

    Private Sub Button0_ClickBefore(sboObject As Object, pVal As SBOItemEventArg, ByRef BubbleEvent As Boolean) Handles Button3.ClickBefore, Button0.ClickBefore,
        Button1.ClickBefore, Button2.ClickBefore, Button4.ClickBefore, Button5.ClickBefore, Button6.ClickBefore,
        Button7.ClickBefore, Button9.ClickBefore, Button8.ClickBefore, Grid6.ClickBefore, Button13.ClickBefore
        Me.m_BaseObject.OnItemClickBefore(RuntimeHelpers.GetObjectValue(sboObject), pVal, BubbleEvent)
    End Sub

    Private Sub EditText0_ChooseFromListAfter(sboObject As Object, pVal As SBOItemEventArg) Handles EditText1.ChooseFromListAfter, Matrix0.ChooseFromListAfter
        Me.m_BaseObject.OnChooseFromListAfter(RuntimeHelpers.GetObjectValue(sboObject), pVal)
    End Sub

    Private Sub Matrix0_PressedBefore(sboObject As Object, pVal As SBOItemEventArg, ByRef BubbleEvent As Boolean) Handles Matrix0.PressedBefore

        ' The following expression was wrapped in a checked-expression
        Dim flag As Boolean = pVal.Row = Me.Matrix0.RowCount + 1
        If flag Then
            Dim flag2 As Boolean = pVal.Row = 1
            If flag2 Then
                Me.Matrix0.AddRow(1, -1)
            Else
                Me.Matrix0.AddRow(1, Me.Matrix0.RowCount)
            End If
            Me.Matrix0.Columns.Item(1).Cells.Item(pVal.Row).Click(BoCellClickType.ct_Regular, 0)
        End If
    End Sub

    Private Sub ButtonCombo0_ComboSelectAfter(sboObject As Object, pVal As SBOItemEventArg) Handles ButtonCombo0.ComboSelectAfter, ButtonCombo2.ComboSelectAfter, ButtonCombo4.ComboSelectAfter
        Me.m_BaseObject.OnComboSelectAfter(RuntimeHelpers.GetObjectValue(sboObject), pVal)
    End Sub

    Private WithEvents Button13 As Button
End Class

