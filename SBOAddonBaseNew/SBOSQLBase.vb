Public MustInherit Class SBOSQLBase

    Protected MustOverride Function GetAppResource(filename As String) As System.IO.Stream

    Public Function GetSQLString(ByVal strProc As String, ByVal ParamArray Parameters() As String) As String
        Dim Ret As String = Nothing
        Dim strSQL As String = ""
        Dim file As System.IO.Stream = GetAppResource(strProc + ".txt")

        Try
            m_ErrMsg = "'"
            m_ErrNo = 0

            If file Is Nothing Then
                m_ErrMsg = "SQL File - " + strProc + " does not exist"
                m_ErrNo = -10000
                Return ""
            End If

            Dim SReader As New System.IO.StreamReader(file)

            strSQL = SReader.ReadToEnd

            Dim i As Integer, strHolder As String = "OWAPARAM", length As Integer
            length = Parameters.Length
            If length > 0 Then
                For i = 0 To length - 1
                    strSQL = strSQL.Replace(strHolder + (i + 1).ToString, Parameters(i))
                Next
            End If
            Ret = strSQL

        Catch ex As Exception
            m_ErrNo = -10000
            m_ErrMsg = ex.Message
            Ret = Nothing

        End Try
        GetSQLString = Ret
    End Function


    Private m_ErrNo As Integer, m_ErrMsg As String
    Public ReadOnly Property ErrorNo As Integer
        Get
            Return m_ErrNo
        End Get
    End Property

    Public ReadOnly Property ErrorMsg As String
        Get
            Return m_ErrMsg
        End Get
    End Property

    Public Function ExecuteSQL(ByVal strProc As String, ByVal ParamArray Parameters() As String) As SAPbobsCOM.Recordset
        Dim SBO_RecSet As SAPbobsCOM.Recordset = Nothing
        Dim sSQL As String
        Try
            m_ErrMsg = "'"
            m_ErrNo = 0
            SBO_RecSet = m_AddOn.SboCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            sSQL = GetSQLString(strProc, Parameters)
            SBO_RecSet.DoQuery(sSQL)
        Catch ex As Exception
            m_ErrMsg = ex.Message
            m_ErrNo = -10000
        End Try
        ExecuteSQL = SBO_RecSet
    End Function


    Private m_AddOn As SboAddon
    Sub New(pAddOn As SboAddon)
        m_AddOn = pAddOn
    End Sub

End Class
