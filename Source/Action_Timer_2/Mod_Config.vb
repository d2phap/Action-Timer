
Imports Microsoft.Win32
Imports Microsoft.Win32.Registry

Module Mod_Config

    Public Sub StartWithWindows_Enable()
        SetValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run", _
                Application.ProductName, Application.ExecutablePath, _
                RegistryValueKind.String)
    End Sub


    Public Sub StartWithWindows_Disable()
        SetValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run", _
                  Application.ProductName, "", RegistryValueKind.String)
    End Sub


    Public Sub SaveAllSettingToReg()
        Dim hkey As String = "HKEY_CURRENT_USER\Software\PhapSoftware\" & _
            Application.ProductName

        'lưu vị trí của Form Main------------
        If FrmMain.Top > -1 And FrmMain.Top < Screen.PrimaryScreen.WorkingArea.Bottom - 100 And _
            FrmMain.Left > -1 And FrmMain.Left < Screen.PrimaryScreen.WorkingArea.Right - 50 Then

            SetValue(hkey, "FrmMain_Top", FrmMain.Top, RegistryValueKind.DWord)
            SetValue(hkey, "FrmMain_Left", FrmMain.Left, RegistryValueKind.DWord)

        End If

        'lưu vị trí của Form Reminder--------
        If FrmReminder.Top > -1 And FrmMain.Top < Screen.PrimaryScreen.WorkingArea.Bottom - 100 And _
            FrmReminder.Left > -1 And FrmMain.Left < Screen.PrimaryScreen.WorkingArea.Right - 50 Then

            SetValue(hkey, "FrmReminder_Top", FrmReminder.Top, RegistryValueKind.DWord)
            SetValue(hkey, "FrmReminder_Left", FrmReminder.Left, RegistryValueKind.DWord)

        End If

        'lưu QUICK ACTION--------------------
        SetValue(hkey, "Qui_ChkConfirm", FrmMain.ChkConfirm.Checked.ToString, RegistryValueKind.String)

        'lưu REMINDER------------------------
        SetValue(hkey, "Rem_ChkTip", FrmMain.ChkTip.Checked.ToString, RegistryValueKind.String)
        SetValue(hkey, "Rem_ChkMail", FrmMain.ChkEMail.Checked.ToString, RegistryValueKind.String)

        'lưu SETTING-------------------------
        SetValue(hkey, "Set_ChkSound", FrmMain.ChkSetSound.Checked.ToString, RegistryValueKind.String)
        SetValue(hkey, "Set_TxtSound_Pro", FrmMain.TxtSetPro.Text, RegistryValueKind.String)
        SetValue(hkey, "Set_TxtSound_Rem", FrmMain.TxtSetRem.Text, RegistryValueKind.String)
        SetValue(hkey, "Set_TxtSound_Sys", FrmMain.TxtSetSys.Text, RegistryValueKind.String)
        SetValue(hkey, "Set_Volume", FrmMain.Vol1.Value, RegistryValueKind.DWord)
        SetValue(hkey, "Set_ChkGiaHan", FrmMain.ChkGiaHan.Checked.ToString, RegistryValueKind.String)
        SetValue(hkey, "Set_GiaHan", FrmMain.NumGiaHan.Value, RegistryValueKind.DWord)
        SetValue(hkey, "Set_ChkStartWithWindows", FrmMain.ChkStartWithWindows.Checked.ToString, RegistryValueKind.String)
        SetValue(hkey, "Set_ChkMinimize", FrmMain.ChkMinimize.Checked.ToString, RegistryValueKind.String)
        SetValue(hkey, "Set_ChkAutoMinimize", FrmMain.ChkAutoMinimize.Checked.ToString, RegistryValueKind.String)
        SetValue(hkey, "Set_ChkHideTooltip", FrmMain.ChkHideTooltip.Checked.ToString, RegistryValueKind.String)

    End Sub


    Public Sub LoadAllSettingFromReg()
        Dim hkey As String = "HKEY_CURRENT_USER\Software\PhapSoftware\" & _
            Application.ProductName

        'load vị trí form Main-----------
        FrmMain.Top = GetValue(hkey, "FrmMain_Top", 21)
        FrmMain.Left = GetValue(hkey, "FrmMain_left", 75)

        'load QUICK ACTION---------------
        FrmMain.ChkConfirm.Checked = GetValue(hkey, "Qui_ChkConfirm", True)

        'load REMINDER-------------------
        FrmMain.ChkTip.Checked = GetValue(hkey, "Rem_ChkTip", True)
        FrmMain.ChkEMail.Checked = GetValue(hkey, "Rem_ChkMail", False)

        'load SETTING--------------------
        FrmMain.ChkSetSound.Checked = GetValue(hkey, "Set_ChkSound", False)
        FrmMain.TxtSetPro.Text = GetValue(hkey, "Set_TxtSound_Pro", "")
        FrmMain.TxtSetRem.Text = GetValue(hkey, "Set_TxtSound_Rem", "")
        FrmMain.TxtSetSys.Text = GetValue(hkey, "Set_TxtSound_Sys", "")
        FrmMain.Vol1.Value = GetValue(hkey, "Set_Volume", 1000)
        FrmMain.ChkGiaHan.Checked = GetValue(hkey, "Set_ChkGiaHan", True)
        FrmMain.NumGiaHan.Value = GetValue(hkey, "Set_GiaHan", 15)
        FrmMain.ChkStartWithWindows.Checked = GetValue(hkey, "Set_ChkStartWithWindows", False)
        FrmMain.ChkMinimize.Checked = GetValue(hkey, "Set_ChkMinimize", False)
        FrmMain.ChkAutoMinimize.Checked = GetValue(hkey, "Set_ChkAutoMinimize", False)
        FrmMain.ChkHideTooltip.Checked = GetValue(hkey, "Set_ChkHideTooltip", False)

    End Sub

    Public Sub PreparingWorking()

        Dim INI As New ClassIniFile(AppPath() & "Action List.ini")

        'tiếp tục công việc dang dở
        If INI.GetString("ACTION STATUS", "Continue", "False") = "True" Then

            FrmMain.T1.Enabled = True

            FrmMain.ChkStart.Enabled = False
            FrmMain.ChkStart.Text = " "

            FrmMain.ChkPause.Enabled = True
            FrmMain.ChkPause.Text = "Tạm dừng"
            FrmMain.ChkSkip.Enabled = True
            FrmMain.ChkSkip.Text = "Bỏ qua"
            FrmMain.ChkStop.Enabled = True
            FrmMain.ChkStop.Text = "Dừng lại"

        Else
            FrmMain.T1.Enabled = False

            FrmMain.ChkStart.Enabled = True
            FrmMain.ChkStart.Text = "Thực hiện"

            FrmMain.ChkPause.Enabled = False
            FrmMain.ChkPause.Text = " "
            FrmMain.ChkSkip.Enabled = False
            FrmMain.ChkSkip.Text = " "
            FrmMain.ChkStop.Enabled = False
            FrmMain.ChkStop.Text = " "

        End If

        'THỰC THI DỮ LIỆU ĐÃ LOAD TỪ REG===================================

        'SETTING------------------------
        

        If FrmMain.ChkSetSound.Checked = True Then
            FrmMain.TxtSetPro.Enabled = True
            FrmMain.TxtSetRem.Enabled = True
            FrmMain.TxtSetSys.Enabled = True

            FrmMain.BtnBrowsePro.Enabled = True
            FrmMain.BtnBrowseRem.Enabled = True
            FrmMain.BtnBrowseSys.Enabled = True

            FrmMain.Vol1.Enabled = True
            FrmMain.PicAudioBack.Enabled = True
            FrmMain.PicAudioPlay.Enabled = True
            FrmMain.PicAudioStop.Enabled = True
            FrmMain.PicAudioNext.Enabled = True
        Else
            FrmMain.TxtSetPro.Enabled = False
            FrmMain.TxtSetRem.Enabled = False
            FrmMain.TxtSetSys.Enabled = False

            FrmMain.BtnBrowsePro.Enabled = False
            FrmMain.BtnBrowseRem.Enabled = False
            FrmMain.BtnBrowseSys.Enabled = False

            FrmMain.Vol1.Enabled = False
            FrmMain.PicAudioBack.Enabled = False
            FrmMain.PicAudioPlay.Enabled = False
            FrmMain.PicAudioStop.Enabled = False
            FrmMain.PicAudioNext.Enabled = False
        End If

        If FrmMain.ChkGiaHan.Checked = True Then
            FrmMain.NumGiaHan.Enabled = True
        Else
            FrmMain.NumGiaHan.Enabled = False
        End If

        'REMINDER-----------------------
        If FrmMain.ChkTip.Checked = True Then
            FrmMain.TxtReminder.Enabled = True
            FrmMain.BtnClearReminder.Enabled = True
            FrmMain.BtnOpenReminder.Enabled = True
            FrmMain.BtnSaveReminder.Enabled = True
        Else
            FrmMain.TxtReminder.Enabled = False
            FrmMain.BtnClearReminder.Enabled = False
            FrmMain.BtnOpenReminder.Enabled = False
            FrmMain.BtnSaveReminder.Enabled = False
        End If

        If FrmMain.ChkEMail.Checked = True Then
            FrmMain.TxtEMail.Enabled = True
            FrmMain.BtnNewEMail.Enabled = True
        Else
            FrmMain.TxtEMail.Enabled = False
            FrmMain.BtnNewEMail.Enabled = False
        End If

    End Sub

  

End Module
