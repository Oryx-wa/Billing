Imports Microsoft.VisualBasic
Imports SAPbouiCOM
Imports SBO.SboAddOnBase
Imports System
Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.IO
Imports System.Runtime.CompilerServices
Imports System.Threading
Imports System.Windows.Forms


Public Class PayrollAddOn
    Inherits SboAddon

    Private WithEvents SBO_Application As SAPbouiCOM.Application

    Sub New()
        'SBO_Application = Application.SBO_Application
    End Sub

    Public Sub New(ByVal StartUpPath As String, ByVal AddonName As String, ByRef pbo_RunApplication As Boolean)

        MyBase.New(StartUpPath, AddonName)
        m_Namespace = "OWA.SBO.OryxBillingWinx10"
        m_AssemblyName = "OryxBillingWinx10"
        TablePrefix = "OWA"
        PermissionPrefix = "OWA_EDU"
        MenuXMLFileName = "PayrollMenus.xml"
        usePermissions = False



        If IsNothing(m_SboApplication) Then
            pbo_RunApplication = False
            Exit Sub
        Else

            If Not initialise() Then
                pbo_RunApplication = False
                Exit Sub
            End If

        End If
        oApp.Run()
        pbo_RunApplication = True
        'Me.setFilters(oFilters)

    End Sub
    <STAThread()>
    Public Sub Main()

    End Sub


    Public Overrides Sub WriteLog(ByVal strLog As String)
        Dim appstartPath As String = Windows.Forms.Application.StartupPath + "\PayrollLog.txt"
        Dim file As New System.IO.StreamWriter(appstartPath, True)
        file.Write(strLog + ",,,,,,,")
        file.Close()
    End Sub

    Public Overrides Sub WriteLog(ByVal strLog As String, ByVal strFileName As String)
        Dim appstartPath As String = Windows.Forms.Application.StartupPath + strFileName
        Dim file As New System.IO.StreamWriter(appstartPath, True)
        file.Write(strLog)
        file.Close()
    End Sub
End Class

