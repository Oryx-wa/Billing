Imports SAPbouiCOM
Imports SBO.SboAddOnBase
Imports System


Public Class SBOEDUClassObj
    Inherits SBOBaseObject

    Public Sub New(pAddOn As SboAddon, pForm As IForm)
        MyBase.New(pAddOn, pForm)
    End Sub

    Protected Overrides Sub EnableToolBarButtons()
        MyBase.EnableToolBarButtons()
        Me.m_Form.EnableMenu("1292", True)
        Me.m_Form.EnableMenu("1293", True)
    End Sub
End Class

