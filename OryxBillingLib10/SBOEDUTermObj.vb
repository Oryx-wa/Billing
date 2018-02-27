Imports Microsoft.VisualBasic.CompilerServices
Imports SAPbouiCOM
Imports SBO.SboAddOnBase
Imports System
Imports System.Runtime.CompilerServices


Public Class SBOEDUTermObj
    Inherits SBOBaseObject

    Public Sub New(pAddOn As SboAddon, pForm As IForm)
        MyBase.New(pAddOn, pForm)
    End Sub

    Protected Overrides Sub EnableToolBarButtons()
        MyBase.EnableToolBarButtons()
        Me.m_Form.EnableMenu("1292", True)
        Me.m_Form.EnableMenu("1293", True)
    End Sub

    Protected Overrides Sub AddDataSource()
        MyBase.AddDataSource()
        Me.m_DBDataSource0 = Me.m_Form.DataSources.DBDataSources.Item("@OWA_EDUTERM")
        Me.m_DBDataSource1 = Me.m_Form.DataSources.DBDataSources.Item("@OWA_EDUSESSIONS")
    End Sub

    Public Overrides Sub OnChooseFromListAfter(sboObject As Object, pVal As SBOItemEventArg)
        MyBase.OnChooseFromListAfter(RuntimeHelpers.GetObjectValue(sboObject), pVal)
        Dim arg_1C_1 As String = pVal.FormUID
        Dim flag As Boolean = False
        Dim text As String = Me.HandleChooseFromListEvent(arg_1C_1, pVal, flag)
        Dim flag2 As Boolean = String.IsNullOrEmpty(text)
        If Not flag2 Then
            Dim itemUID As String = pVal.ItemUID
            flag2 = (Operators.CompareString(itemUID, "txSesion", False) = 0)
            If flag2 Then
                Me.m_DBDataSource0.SetValue("U_session", Me.m_DBDataSource0.Offset, text)
            End If
        End If
    End Sub
End Class

