

Imports System.Management   'Project > Add Reference...

Module Mod_System

    Private Declare Function ExitWindowsEx Lib "user32.dll" (ByVal uFlags As Long, ByVal dwreserved As Long) As Long
    Private Declare Function SetSuspendState Lib "Powrprof" (ByVal Hibernate As Long, ByVal ForceCritical As Long, ByVal DisableWakeEvent As Long) As Long
    Private Declare Function LockWorkStation Lib "user32.dll" () As Long
    Private Declare Function AdjustTokenPrivileges Lib "advapi32.dll" (ByVal _
        TokenHandle As Long, ByVal DisableAllPrivileges As Long, ByVal NewState As  _
        TOKEN_PRIVILEGES, ByVal BufferLength As Long, ByVal PreviousState As  _
        TOKEN_PRIVILEGES, ByVal ReturnLength As Long) As Long
    Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
    Private Declare Function OpenProcessToken Lib "advapi32" (ByVal ProcessHandle _
        As Long, ByVal DesiredAccess As Long, ByVal TokenHandle As Long) As Long
    Private Declare Function LookupPrivilegeValue Lib "advapi32" Alias _
        "LookupPrivilegeValueA" (ByVal lpSystemName As String, ByVal lpName As _
        String, ByVal lpLuid As LUID) As Long

    Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
        (ByVal hwnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, _
         ByVal lParam As Integer) As Integer

    Private Const SHERB_NOCONFIRMATION As Long = &H1
    Private Const SHERB_NOPROGRESSUI As Long = &H2
    Private Const SHERB_NOSOUND As Long = &H4

    Private Declare Auto Function SHEmptyRecycleBin Lib "shell32.dll" Alias "SHEmptyRecycleBinA" (ByVal hwnd As Integer, ByVal pszRootPath As String, ByVal dwFlags As Integer) As Long
    Private Declare Auto Function SHUpdateRecycleBinIcon Lib "shell32.dll" () As Long

    Private Structure LUID
        Dim LowPart As Long
        Dim HighPart As Long
    End Structure

    Private Structure LUID_AND_ATTRIBUTES
        Dim pLuid As LUID
        Dim Attributes As Long
    End Structure

    Private Structure TOKEN_PRIVILEGES
        Dim PrivilegeCount As Long
        Dim Privileges() As LUID_AND_ATTRIBUTES
    End Structure

    Private Enum PowerFlags
        LogOff = 0
        Reboot = 2
        ShutDown = 1
        Force = 4
    End Enum


    Private Enum ShutDown
        LogOff = 0
        Shutdown = 1
        Reboot = 2
        PowerOff = 8

        ForceLogOff = 4
        ForceShutdown = 5
        ForceReboot = 6
        ForcePowerOff = 12

    End Enum

    Private Const SC_SCREENSAVE = &HF140&
    Private Const WM_SYSCOMMAND = &H112
    Private Const MONITOR_ON = -1&
    Private Const SC_MONITORPOWER = &HF170&
    Private Const MONITOR_OFF = 2&





    Public Function power_Shutdown() As Long
        SetPowerMode(ShutDown.ForceShutdown)
    End Function

    Public Function power_Restart() As Long
        Try
            SetPowerMode(ShutDown.ForceReboot)
        Catch ex As Exception

        End Try

    End Function

    Public Function power_LogOff() As Long
        SetPowerMode(ShutDown.ForceLogOff)
    End Function

    Public Function power_Hibernate() As Long
        Application.SetSuspendState(PowerState.Hibernate, True, True)
    End Function

    Public Function power_StandBy() As Long
        Application.SetSuspendState(PowerState.Suspend, True, True)
    End Function

    Public Function power_LockComputer() As Long
        power_LockComputer = LockWorkStation()
        'error =0
    End Function




    Public Sub power_RunScreenSaver()
        Call SendMessage(FrmMain.Handle, 274, 61760, 20&)
    End Sub

    Public Sub power_TurnOffMonitor()
        Call SendMessage(FrmMain.Handle, WM_SYSCOMMAND, SC_MONITORPOWER, MONITOR_OFF)
    End Sub

    Public Sub power_TurnOnMonitor()
        Call SendMessage(FrmMain.Handle, WM_SYSCOMMAND, SC_MONITORPOWER, MONITOR_ON)
    End Sub

    Public Sub EmptyRecycleBin()
        SHEmptyRecycleBin(FrmMain.Handle.ToInt32, "", SHERB_NOCONFIRMATION)
        SHUpdateRecycleBinIcon()
    End Sub


    Private Sub SetPowerMode(ByVal Flags As ShutDown)
        Dim W32_OS As New ManagementClass("Win32_OperatingSystem")
        Dim inParams, outParams As ManagementBaseObject
        Dim result As Integer

        W32_OS.Scope.Options.EnablePrivileges = True

        Dim obj As ManagementObject

        For Each obj In W32_OS.GetInstances
            inParams = obj.GetMethodParameters("Win32Shutdown")
            inParams("Flags") = Flags 'ShutDown.ForceShutdown
            inParams("Reserved") = 0
            outParams = obj.InvokeMethod("Win32Shutdown", inParams, Nothing)

            result = Convert.ToInt32(outParams("returnValue"))
            If result = 0 Then
                Throw New System.ComponentModel.Win32Exception(result)
            End If
        Next
    End Sub

End Module
