Imports Microsoft.VisualBasic.CompilerServices
Imports OWA.SBO.OryxSQL
Imports SAPbouiCOM
Imports SBO.SboAddOnBase
Imports System
Imports System.Runtime.CompilerServices


Public MustInherit Class SBOWizard
    Inherits SBOBaseObject

    Private m_CurrentPane As Integer

    Private m_StartPane As Integer

    Private m_LastPane As Integer

    Private m_LastActivePane As Integer

    Protected m_ErrMsg As String

    Protected m_ErrNo As Integer

    Protected Button0 As Button

    Protected Button1 As Button

    Protected Button2 As Button

    Protected StaticText0 As StaticText

    Protected ReadOnly Property CurrentPane() As Integer
        Get
            Return Me.m_CurrentPane
        End Get
    End Property

    Protected ReadOnly Property StartPane() As Integer
        Get
            Return Me.m_StartPane
        End Get
    End Property

    Protected Property LastPane() As Integer
        Get
            Return Me.m_LastPane
        End Get
        Set(value As Integer)
            Me.m_LastPane = value
        End Set
    End Property

    Protected ReadOnly Property LastActivePane() As Integer
        Get
            Return Me.m_LastActivePane
        End Get
    End Property

    Public Sub New(pAddon As SboAddon, pForm As IForm)
        MyBase.New(pAddon, pForm)
        Me.m_StartPane = 1
        Me.m_LastPane = 9
        Me.m_LastActivePane = 8
        Me.InitSBOServerSQL(New BusObjectInfoSQL(pAddon))
    End Sub

    Protected Overrides Sub OnFormInit()
        MyBase.OnFormInit()
        Me.Button0 = CType(Me.m_Form.Items.Item("btnBack").Specific, Button)
        Me.Button1 = CType(Me.m_Form.Items.Item("btnNext").Specific, Button)
        Me.Button2 = CType(Me.m_Form.Items.Item("btnFinish").Specific, Button)
        Me.StaticText0 = CType(Me.m_Form.Items.Item("lblStep").Specific, StaticText)
    End Sub

    Protected Overrides Sub OnFormLoad()
        MyBase.OnFormLoad()
        Me.m_CurrentPane = 1
    End Sub

    Protected Overridable Sub UpdatePaneStatus()
        ' The following expression was wrapped in a checked-expression
        Me.StaticText0.Caption = " Step " + Conversions.ToString(Me.CurrentPane - Me.StartPane) + " of " + Conversions.ToString(Me.LastActivePane)
        Dim paneLevel As Integer = Me.m_Form.PaneLevel
        Dim flag As Boolean = paneLevel = Me.StartPane
        If flag Then
            Me.Button0.Item.Enabled = False
            Me.Button1.Item.Enabled = True
            Me.Button2.Item.Enabled = Me.IsReady(Me.m_ErrNo, Me.m_ErrMsg)
        Else
            flag = (paneLevel = Me.LastPane)
            If flag Then
                Me.Button0.Item.Enabled = True
                Me.Button1.Item.Enabled = False
                Me.Button2.Item.Enabled = True
            Else
                Me.Button0.Item.Enabled = True
                Me.Button1.Item.Enabled = True
                Me.Button2.Item.Enabled = Me.IsReady(Me.m_ErrNo, Me.m_ErrMsg)
            End If
        End If
    End Sub

    Protected Overrides Function IsReady(ByRef pErrNo As Integer, ByRef pErrMsg As String) As Boolean
        Return False
    End Function

    Protected Overridable Function PageMoveOk(ByRef pErrNo As Integer, ByRef pErrMsg As String) As Boolean
        Return True
    End Function

    Protected Overridable Sub InputRetrieve()
        Me.FormRefresh()
    End Sub

    Public Overrides Sub OnChooseFromListAfter(sboObject As Object, pVal As SBOItemEventArg)
        MyBase.OnChooseFromListAfter(RuntimeHelpers.GetObjectValue(sboObject), pVal)
    End Sub

    Public Overrides Sub OnItemClickBefore(sboObject As Object, pVal As SBOItemEventArg, ByRef BubbleEvent As Boolean)
        MyBase.OnItemClickBefore(RuntimeHelpers.GetObjectValue(sboObject), pVal, BubbleEvent)
        Dim itemUID As String = pVal.ItemUID
        Dim flag As Boolean = Operators.CompareString(itemUID, "btnNext", False) = 0
        If flag Then
            Dim flag2 As Boolean = Me.m_CurrentPane <> 1
            If flag2 Then
                Dim flag3 As Boolean = Not Me.PageMoveOk(Me.m_ErrNo, Me.m_ErrMsg)
                If flag3 Then
                    Me.m_SboApplication.StatusBar.SetText(Me.m_ErrMsg, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
                    BubbleEvent = False
                End If
            End If
        Else
            Dim flag3 As Boolean = Operators.CompareString(itemUID, "btnFinish", False) = 0
            If flag3 Then
                Dim flag2 As Boolean = Not Me.FinalizeWizard(Me.m_ErrNo, Me.m_ErrMsg)
                If flag2 Then
                    Me.m_SboApplication.StatusBar.SetText(Me.m_ErrMsg, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
                    BubbleEvent = False
                End If
            End If
        End If
    End Sub

    Public Overrides Sub OnItemClickAfter(sboObject As Object, pVal As SBOItemEventArg)
        MyBase.OnItemClickAfter(RuntimeHelpers.GetObjectValue(sboObject), pVal)
        Dim itemUID As String = pVal.ItemUID
        Dim flag As Boolean = Operators.CompareString(itemUID, "btnBack", False) = 0
        ' The following expression was wrapped in a checked-statement
        If flag Then
            Me.m_CurrentPane -= 1
            Me.m_Form.PaneLevel = Me.CurrentPane
            Me.UpdatePaneStatus()
            Me.RetrievePageData()
        Else
            flag = (Operators.CompareString(itemUID, "btnNext", False) = 0)
            If flag Then
                Me.m_CurrentPane += 1
                Me.m_Form.PaneLevel = Me.CurrentPane
                Me.UpdatePaneStatus()
                Me.RetrievePageData()
            End If
        End If
    End Sub

    Protected Sub SetPanelLevel(PanelLevel As Integer)
        Me.m_CurrentPane = PanelLevel
    End Sub

    Protected Overridable Sub RetrievePageData()
    End Sub

    Public Overridable Function FinalizeWizard(ByRef pErrNo As Integer, ByRef pErrMsg As String) As Boolean
        Dim flag2 As Boolean
        Dim flag As Boolean = flag2
        If flag Then
            Me.m_Form.Close()
        End If
        Return flag2
    End Function

    Protected Overridable Sub AddNewSonForm()
        Me.m_SonForm = CType(Me.m_ParentAddon.CreateSonForm(Me.SonFormName), UserFormBaseClass)
    End Sub

    Protected Overridable Sub EditSonForm()
        Me.AddNewSonForm()
    End Sub
End Class

