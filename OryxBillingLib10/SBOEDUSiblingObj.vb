Imports Microsoft.VisualBasic.CompilerServices
Imports SAPbouiCOM
Imports SBO.SboAddOnBase
Imports System
Imports System.Runtime.CompilerServices


Public Class SBOEDUSiblingObj
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
        Me.m_DBDataSource0 = Me.m_Form.DataSources.DBDataSources.Item("@OWA_EDUSIBLING")
        Me.m_DBDataSource1 = Me.m_Form.DataSources.DBDataSources.Item("OCRD")
    End Sub

    Public Overrides Sub OnCustomInit()
        MyBase.OnCustomInit()
        Me.Matrix0 = CType(Me.m_Form.Items.Item("dgsibling").Specific, Matrix)
    End Sub

    Public Overrides Sub OnChooseFromListAfter(sboObject As Object, pVal As SBOItemEventArg)
        MyBase.OnChooseFromListAfter(RuntimeHelpers.GetObjectValue(sboObject), pVal)
        Dim arg_1C_1 As String = pVal.FormUID
        Dim flag As Boolean = False
        Dim text As String = Me.HandleChooseFromListEvent(arg_1C_1, pVal, flag)
        Dim flag2 As Boolean = String.IsNullOrEmpty(text)
        If Not flag2 Then
            Me.Matrix0 = CType(Me.m_Form.Items.Item("dgsibling").Specific, Matrix)
            Dim itemUID As String = pVal.ItemUID
            flag2 = (Operators.CompareString(itemUID, "dgsibling", False) = 0)
            If flag2 Then
                Dim colUID As String = pVal.ColUID
                Dim flag3 As Boolean = Operators.CompareString(colUID, "V_1", False) = 0
                If flag3 Then
                    Me.Matrix0.SetCellWithoutValidation(pVal.Row, pVal.ColUID, text)
                    flag3 = Me.getOffset(text, "CardCode", Me.m_DBDataSource1)
                    If flag3 Then
                        Me.Matrix0.SetCellWithoutValidation(pVal.Row, "V_2", Me.m_DBDataSource1.GetValue("CardName", 0).Trim())
                    End If
                End If
            End If
        End If
    End Sub

    Protected Overrides Function IsReady(ByRef pErrNo As Integer, ByRef pErrMsg As String) As Boolean
        Return MyBase.IsReady(pErrNo, pErrMsg)
    End Function

    Protected Overrides Sub SetConditions()
        MyBase.SetConditions()
        Dim condition As Condition = Me.m_Condition
        condition.[Alias] = "QryGroup1"
        condition.Operation = BoConditionOperation.co_EQUAL
        condition.CondVal = "Y"
        condition.Relationship = BoConditionRelationship.cr_AND
        Me.m_Condition = Me.m_Conditions.Add()
        Dim condition2 As Condition = Me.m_Condition
        condition2.[Alias] = "FrozenFor"
        condition2.Operation = BoConditionOperation.co_NOT_EQUAL
        condition2.CondVal = "Y"
        Me.m_Form.ChooseFromLists.Item("CFL_OCRD").SetConditions(Me.m_Conditions)
    End Sub
End Class

