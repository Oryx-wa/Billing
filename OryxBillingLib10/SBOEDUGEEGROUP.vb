Imports SAPbouiCOM
Imports SBO.SboAddOnBase
Imports System
Imports System.Runtime.CompilerServices


Public Class SBOEDUGEEGROUP
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
        Me.m_DBDataSource0 = Me.m_Form.DataSources.DBDataSources.Item("@OWA_EDUFEEGROUP")
    End Sub

    Public Overrides Sub OnCustomInit()
        MyBase.OnCustomInit()
        Me.m_DBDataSource0.Query(Nothing)
    End Sub

    Public Overrides Sub OnItemPressedBefore(sboObject As Object, pVal As SBOItemEventArg, ByRef BubbleEvent As Boolean)
        MyBase.OnItemPressedBefore(RuntimeHelpers.GetObjectValue(sboObject), pVal, BubbleEvent)
        Dim itemUID As String = pVal.ItemUID
    End Sub
End Class

