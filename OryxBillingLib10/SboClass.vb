Imports Microsoft.VisualBasic.CompilerServices
Imports SAPbouiCOM
Imports SBO.SboAddOnBase
Imports System
Imports System.Runtime.CompilerServices


Public Class SboClass
    Inherits SBOBaseObject

    Private Matrix0 As Matrix

    Private EditText0 As EditText

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
        Me.m_DBDataSource0 = Me.m_Form.DataSources.DBDataSources.Item("@OWA_EDUCLASS")
        Me.m_DBDataSource1 = Me.m_Form.DataSources.DBDataSources.Item("@OWA_EDUKEYSTAGE")
    End Sub

    Public Overrides Sub OnItemClickBefore(sboObject As Object, pVal As SBOItemEventArg, ByRef BubbleEvent As Boolean)
        MyBase.OnItemClickBefore(RuntimeHelpers.GetObjectValue(sboObject), pVal, BubbleEvent)
        Dim itemUID As String = pVal.ItemUID
        Dim flag As Boolean = Operators.CompareString(itemUID, "1", False) = 0
        If flag Then
            Dim flag2 As Boolean = BubbleEvent
            If flag2 Then
                Me.m_ParentAddon.SboApplication.StatusBar.SetText("Operation Completed Successfully", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success)
                Me.m_Form.Mode = BoFormMode.fm_OK_MODE
            End If
        End If
    End Sub

    Public Overrides Sub OnItemPressedAfter(sboObject As Object, pVal As SBOItemEventArg)
        MyBase.OnItemPressedAfter(RuntimeHelpers.GetObjectValue(sboObject), pVal)
        Me.m_CurrentLineNo = pVal.Row
    End Sub

    Protected Overrides Sub OnMatrixAddRow()
        MyBase.OnMatrixAddRow()
        Me.AddRowToMatrix(Me.Matrix0)
    End Sub

    Protected Overrides Sub OnFormInit()
        MyBase.OnFormInit()
        Me.Matrix0 = CType(Me.m_Form.Items.Item("Matrix1").Specific, Matrix)
        Me.AddRowToMatrix(Me.Matrix0)
    End Sub
End Class

