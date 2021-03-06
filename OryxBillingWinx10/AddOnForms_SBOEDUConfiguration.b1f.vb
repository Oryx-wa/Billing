
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


<FormAttribute("OryxBillingWinx10.AddOnForms_SBOEDUConfiguration_b1f", "AddOnForms_SBOEDUConfiguration.b1f")>
Friend Class AddOnForms_SBOEDUConfiguration_b1f
    Inherits UserFormBaseClass
    Private WithEvents m_DBDataSource0 As SAPbouiCOM.DBDataSource
    Private WithEvents StaticText2 As SAPbouiCOM.StaticText
    Private WithEvents EditText2 As SAPbouiCOM.EditText
    Private WithEvents StaticText3 As SAPbouiCOM.StaticText
    Private WithEvents EditText3 As SAPbouiCOM.EditText
    Private WithEvents StaticText4 As SAPbouiCOM.StaticText
    Private WithEvents EditText4 As SAPbouiCOM.EditText
    Private WithEvents StaticText5 As SAPbouiCOM.StaticText
    Private WithEvents EditText5 As SAPbouiCOM.EditText
    Private WithEvents Button0 As SAPbouiCOM.Button
    Private WithEvents Button1 As SAPbouiCOM.Button

    Public Sub New()

    End Sub

    Protected Overrides Sub InitBase(pAddOn As SboAddon)
        MyBase.InitBase(pAddOn)
        Me.CreateObject(New SBOEDUConfig(pAddOn, Me.UIAPIRawForm))
    End Sub

    Public Overrides Sub OnInitializeComponent()
        Me.StaticText2 = CType(Me.GetItem("Item_4").Specific, StaticText)
        Me.StaticText3 = CType(Me.GetItem("Item_6").Specific, StaticText)
        Me.StaticText4 = CType(Me.GetItem("Item_8").Specific, StaticText)
        Me.StaticText5 = CType(Me.GetItem("Item_10").Specific, StaticText)
        Me.EditText2 = CType(Me.GetItem("txsbCnt").Specific, EditText)
        Me.EditText3 = CType(Me.GetItem("txsbdcnt").Specific, EditText)
        Me.EditText4 = CType(Me.GetItem("txkidtui").Specific, EditText)
        Me.EditText5 = CType(Me.GetItem("txOneTiP").Specific, EditText)
        Me.Button1 = CType(Me.GetItem("2").Specific, Button)
        Me.Button0 = CType(Me.GetItem("btnSave").Specific, Button)
        'Me.EditText6 = CType(Me.GetItem("Item_1").Specific, EditText)
        'Me.EditText7 = CType(Me.GetItem("txtCode").Specific, EditText)
        Me.OnCustomInitialize()
    End Sub

    Public Overrides Sub OnInitializeFormEvents()
    End Sub

    Private Sub OnCustomInitialize()
    End Sub

    Private Sub Button0_ClickAfter(sboObject As Object, pVal As SBOItemEventArg)
        Me.m_BaseObject.OnItemClickAfter(RuntimeHelpers.GetObjectValue(sboObject), pVal)
    End Sub
End Class

