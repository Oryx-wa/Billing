Imports Microsoft.VisualBasic.CompilerServices
Imports SAPbouiCOM
Imports SBO.SboAddOnBase
Imports System
Imports System.Runtime.CompilerServices


Public Class SBOEDUKeyStageObj
    Inherits SBOBaseObject

    Private drgKeySt As Matrix

    Public Sub New(pAddOn As SboAddon, pForm As IForm)
        MyBase.New(pAddOn, pForm)
    End Sub

    Protected Overrides Sub OnFormInit()
        MyBase.OnFormInit()
        Me.drgKeySt = CType(Me.m_Form.Items.Item("drgKeySt").Specific, Matrix)
    End Sub

    Protected Overrides Sub EnableToolBarButtons()
        MyBase.EnableToolBarButtons()
        Me.m_Form.EnableMenu("1292", True)
        Me.m_Form.EnableMenu("1293", True)
    End Sub

    Protected Overrides Sub AddDataSource()
        MyBase.AddDataSource()
        Me.m_DBDataSource0 = Me.m_Form.DataSources.DBDataSources.Item("@OWA_EDUKEYSTAGE")
    End Sub

    Protected Overrides Function CanDelete() As Boolean
        Dim result As Boolean = False
        Dim flag As Boolean = MyBase.CanDelete()
        If flag Then
            Dim value As String = Me.m_DBDataSource0.GetValue("Code", Me.m_DBDataSource0.Offset)
            Dim strSQL As String = "Select Count(*) Num from [@OWA_EDUCLAHISTORY] Where U_Keystage = '" + value + "'"
            Dim dataTable As DataTable = Me.ExecuteSQLDT(strSQL)
            flag = (dataTable IsNot Nothing)
            Dim flag2 As Boolean
            If flag Then
                flag2 = Operators.ConditionalCompareObjectGreaterEqual(dataTable.GetValue(0, 0), 1, False)
                If flag2 Then
                    Me.m_ParentAddon.SboApplication.SetStatusBarMessage("Key Stage used in Class, Delete not possible", BoMessageTime.bmt_Medium, True)
                End If
            End If
            strSQL = "Select Count(*) Num from [@OWA_EDUCLAHISTORY] Where U_Keystage = '" + value + "'"
            result = False
            dataTable = Me.ExecuteSQLDT(strSQL)
            flag2 = (dataTable IsNot Nothing)
            If flag2 Then
                flag = Operators.ConditionalCompareObjectGreaterEqual(dataTable.GetValue(0, 0), 1, False)
                If flag Then
                    Me.m_ParentAddon.SboApplication.SetStatusBarMessage("Key Stage used in Class, Delete not possible", BoMessageTime.bmt_Medium, True)
                    result = False
                Else
                    result = True
                End If
            End If
        End If
        Return result
    End Function

    Public Overrides Sub OnChooseFromListAfter(sboObject As Object, pVal As SBOItemEventArg)
        MyBase.OnChooseFromListAfter(RuntimeHelpers.GetObjectValue(sboObject), pVal)
        Try
            Dim itemUID As String = pVal.ItemUID
            Dim flag As Boolean = Operators.CompareString(itemUID, "drgKeySt", False) = 0
            If flag Then
                Dim flag2 As Boolean = Operators.CompareString(pVal.ColUID, "colFee", False) = 0
                If flag2 Then
                    Dim arg_56_1 As String = pVal.FormUID
                    Dim flag3 As Boolean = False
                    Dim value As String = Me.HandleChooseFromListEvent(arg_56_1, pVal, flag3)
                    flag2 = String.IsNullOrEmpty(value)
                    If Not flag2 Then
                        Me.drgKeySt.SetCellWithoutValidation(pVal.Row, pVal.ColUID, value)
                    End If
                End If
            End If
        Catch expr_88 As Exception
            ProjectData.SetProjectError(expr_88)
            ProjectData.ClearProjectError()
        End Try
    End Sub
End Class

