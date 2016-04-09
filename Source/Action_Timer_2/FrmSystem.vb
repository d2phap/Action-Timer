Public Class FrmSystem


    Private Sub FrmSystem_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'tạo aero glass form
        Dim AERO As New AeroGlass
        AERO.Apply_AeroGlass(Me)

        Select Case ProSys.Tag
            Case "Shutdown"
                Me.Text = "Đang chuẩn bị tắt máy tính ..."
            Case "Restart"
                Me.Text = "Đang chuẩn bị khởi động lại máy tính ..."
            Case "LogOff"
                Me.Text = "Đang chuẩn bị đăng xuất khỏi tài khoản ..."
            Case "StandBy"
                Me.Text = "Đang chuẩn bị chuyển sang chế độ chờ ..."
            Case "Hibernate"
                Me.Text = "Đang chuẩn bị chuyển sang chế độ ngủ đông ..."

        End Select

        If FrmMain.ChkGiaHan.Checked = True Then
            ProSys.Maximum = FrmMain.NumGiaHan.Value * 100
        Else
            ProSys.Maximum = 150
        End If

        ProSys.Value = 0
        TSys.Tag = 0
        TSys.Enabled = True
    End Sub

    Private Sub TSys_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TSys.Tick

        ProSys.Value += 1

        If ProSys.Value = ProSys.Maximum Then

            Me.Hide()
            TSys.Enabled = False

            Select Case ProSys.Tag
                Case "Shutdown"
                    Mod_Config.SaveAllSettingToReg()
                    Mod_System.power_Shutdown()
                Case "Restart"
                    Mod_Config.SaveAllSettingToReg()
                    Mod_System.power_Restart()
                Case "LogOff"
                    Mod_Config.SaveAllSettingToReg()
                    Mod_System.power_LogOff()
                Case "StandBy"
                    Mod_System.power_StandBy()
                Case "Hibernate"
                    Mod_System.power_Hibernate()
                Case "LockComputer"
                    Mod_System.power_LockComputer()
                Case "TurnOffMonitor"
                    Mod_System.power_TurnOffMonitor()

            End Select


        End If
    End Sub

    Private Sub BtnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnCancel.Click
        TSys.Enabled = False
        Me.Close()
        Me.Dispose()
    End Sub

   
End Class