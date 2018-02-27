Imports Microsoft.VisualBasic.CompilerServices
Imports SBO.SboAddOnBase
Imports System
Imports System.IO
Imports System.Reflection


Public Class BusObjectInfoSQL
    Inherits SBOSQLBase

    Public Sub New(pAddOn As SboAddon)
        MyBase.New(pAddOn)
    End Sub

    Protected Overrides Function GetAppResource(filename As String) As Stream
        Dim result As Stream = Nothing
        Try
            Dim executingAssembly As Assembly = Assembly.GetExecutingAssembly()
            result = executingAssembly.GetManifestResourceStream("OWA.SBO.OryxSQL." + filename)
        Catch expr_1E As Exception
            ProjectData.SetProjectError(expr_1E)
            ProjectData.ClearProjectError()
        End Try
        Return result
    End Function
End Class

