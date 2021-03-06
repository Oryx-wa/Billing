
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


<FormAttribute("OryxBillingWinx10.AddOnForms_SBOEDUClass_b1f", "AddOnForms_SBOEDUClass.b1f")>
Friend Class SBOEDUClass
    Inherits UserFormBaseClass

    Private WithEvents txtCode As SAPbouiCOM.EditText
    Private WithEvents txtName As SAPbouiCOM.EditText
    Private WithEvents Button0 As SAPbouiCOM.Button
    Private WithEvents Button1 As SAPbouiCOM.Button
    Private WithEvents txtSchool As SAPbouiCOM.EditText


    Public Sub New()
    End Sub

    Protected Overrides Sub InitBase(pAddOn As SboAddon)
        MyBase.InitBase(pAddOn)
        Me.CreateObject(New SBOEDUClassObj(pAddOn, Me.UIAPIRawForm))
    End Sub

    Public Overrides Sub OnInitializeComponent()
        Me.txtCode = CType(Me.GetItem("txtCode").Specific, EditText)
        Me.txtName = CType(Me.GetItem("txtName").Specific, EditText)
        Me.Button0 = CType(Me.GetItem("2").Specific, Button)
        Me.Button1 = CType(Me.GetItem("1").Specific, Button)
        ' Me.StaticText0 = CType(Me.GetItem("lblSchool").Specific, StaticText)
        Me.txtSchool = CType(Me.GetItem("txtSchool").Specific, EditText)
        'Me.StaticText1 = CType(Me.GetItem("Item_4").Specific, StaticText)
        Me.OnCustomInitialize()
    End Sub

    Public Overrides Sub OnInitializeFormEvents()
        AddHandler Me.LoadAfter, AddressOf Me.Form_LoadAfter
    End Sub

    Private Sub Form_LoadAfter(pVal As SBOItemEventArg)
    End Sub

    Private Sub OnCustomInitialize()
    End Sub
End Class

