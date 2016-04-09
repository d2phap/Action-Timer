

Imports System
Imports System.Diagnostics

Public Class FrmMain

    '___________________________________________________________________________
    '===========================================================================
    Private Sub PicStatus_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PicStatus.Click

        PicStatus.Image = Action_Timer_2.My.Resources.Resource1.MD_STATUS
        PicStatus.Tag = "click"

        PicQuickAction.Image = Nothing
        PicProcessing.Image = Nothing
        PicReminder.Image = Nothing
        PicSystem.Image = Nothing

        PicActionList.Image = Nothing
        PicSetting.Image = Nothing
        PicAbout.Image = Nothing

        PanTime.BringToFront()
        PanStatus.BringToFront()

        PicQuickAction.Tag = "normal"
        PicProcessing.Tag = "normal"
        PicReminder.Tag = "normal"
        PicSystem.Tag = "normal"
        PicActionList.Tag = "normal"
        PicSetting.Tag = "normal"
        PicAbout.Tag = "normal"

    End Sub

    Private Sub PicStatus_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PicStatus.MouseDown
        PicStatus.Image = Action_Timer_2.My.Resources.Resource1.MD_STATUS
    End Sub

    Private Sub PicStatus_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles PicStatus.MouseEnter
        PicStatus.Image = Action_Timer_2.My.Resources.Resource1.MM_TAB

        If ChkHideTooltip.Checked = True Then

            Me.Text = StrDup(10, " ") & "- Trạng thái chung của chương trình, cho biết công việc kế" & _
                    " tiếp, công việc vừa làm, báo cáo kết quả,..." & StrDup(27, " ")
            TToolTip.Enabled = True
        End If

    End Sub

    Private Sub PicStatus_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles PicStatus.MouseLeave
        If PicStatus.Tag = "normal" Then
            PicStatus.Image = Nothing
        ElseIf PicStatus.Tag = "click" Then
            PicStatus.Image = Action_Timer_2.My.Resources.Resource1.MD_STATUS
        End If

        TToolTip.Enabled = False
        Me.Text = "Action Timer 2010"
    End Sub
    '___________________________________________________________________________
    '===========================================================================



    '___________________________________________________________________________
    '===========================================================================
    Private Sub PicQuickAction_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PicQuickAction.Click
        PicQuickAction.Image = Action_Timer_2.My.Resources.Resource1.MD_QUICK_ACTION
        PicQuickAction.Tag = "click"

        PicStatus.Image = Nothing
        PicProcessing.Image = Nothing
        PicReminder.Image = Nothing
        PicSystem.Image = Nothing

        PicActionList.Image = Nothing
        PicSetting.Image = Nothing
        PicAbout.Image = Nothing

        PanTime.BringToFront()
        PanQuickAction.BringToFront()

        PicStatus.Tag = "normal"
        PicProcessing.Tag = "normal"
        PicReminder.Tag = "normal"
        PicSystem.Tag = "normal"
        PicActionList.Tag = "normal"
        PicSetting.Tag = "normal"
        PicAbout.Tag = "normal"

    End Sub

    Private Sub PicQuickAction_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PicQuickAction.MouseDown
        PicQuickAction.Image = Action_Timer_2.My.Resources.Resource1.MD_QUICK_ACTION
    End Sub

    Private Sub PicQuickAction_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles PicQuickAction.MouseEnter
        PicQuickAction.Image = Action_Timer_2.My.Resources.Resource1.MM_TAB

        If ChkHideTooltip.Checked = True Then
            Me.Text = StrDup(10, " ") & "- Cung cấp một số chức năng thao tác nhanh với máy tính " & _
                    " và kèm theo chức năng Run của Windows" & StrDup(27, " ")
            TToolTip.Enabled = True
        End If
    End Sub

    Private Sub PicQuickAction_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles PicQuickAction.MouseLeave
        If PicQuickAction.Tag = "normal" Then
            PicQuickAction.Image = Nothing
        ElseIf PicQuickAction.Tag = "click" Then
            PicQuickAction.Image = Action_Timer_2.My.Resources.Resource1.MD_QUICK_ACTION
        End If

        TToolTip.Enabled = False
        Me.Text = "Action Timer 2010"
    End Sub
    '___________________________________________________________________________
    '===========================================================================



    '___________________________________________________________________________
    '===========================================================================
    Private Sub PicProcessing_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PicProcessing.Click
        PicProcessing.Image = Action_Timer_2.My.Resources.Resource1.MD_PROCESS
        PicProcessing.Tag = "click"

        PicQuickAction.Image = Nothing
        PicStatus.Image = Nothing
        PicReminder.Image = Nothing
        PicSystem.Image = Nothing

        PicActionList.Image = Nothing
        PicSetting.Image = Nothing
        PicAbout.Image = Nothing

        PanTime.BringToFront()
        PanProcess.BringToFront()

        PicAdd.Tag = "Quản lý tiến trình"

        PicStatus.Tag = "normal"
        PicQuickAction.Tag = "normal"
        PicReminder.Tag = "normal"
        PicSystem.Tag = "normal"
        PicActionList.Tag = "normal"
        PicSetting.Tag = "normal"
        PicAbout.Tag = "normal"
    End Sub

    Private Sub PicProcessing_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PicProcessing.MouseDown
        PicProcessing.Image = Action_Timer_2.My.Resources.Resource1.MD_PROCESS
    End Sub

    Private Sub PicProcessing_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles PicProcessing.MouseEnter
        PicProcessing.Image = Action_Timer_2.My.Resources.Resource1.MM_TAB

        If ChkHideTooltip.Checked = True Then
            Me.Text = StrDup(10, " ") & "- Xem thông tin chi tiết về các tiến trình đang chạy," & _
                " kèm theo chức năng tắt nó.   - Chạy một file do bạn nhập vào" & _
                "    ** Lưu ý: Bạn phải tạo danh sách công việc mới nếu vừa thực " & _
                "hiện xong một danh sách công việc cũ" & StrDup(20, " ")
            TToolTip.Enabled = True
        End If
    End Sub

    Private Sub PicProcessing_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles PicProcessing.MouseLeave
        If PicProcessing.Tag = "normal" Then
            PicProcessing.Image = Nothing
        ElseIf PicProcessing.Tag = "click" Then
            PicProcessing.Image = Action_Timer_2.My.Resources.Resource1.MD_PROCESS
        End If

        TToolTip.Enabled = False
        Me.Text = "Action Timer 2010"
    End Sub
    '___________________________________________________________________________
    '===========================================================================



    '___________________________________________________________________________
    '===========================================================================
    Private Sub PicReminder_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PicReminder.Click
        PicReminder.Image = Action_Timer_2.My.Resources.Resource1.MD_REMINDER
        PicReminder.Tag = "click"

        PicQuickAction.Image = Nothing
        PicProcessing.Image = Nothing
        PicStatus.Image = Nothing
        PicSystem.Image = Nothing

        PicActionList.Image = Nothing
        PicSetting.Image = Nothing
        PicAbout.Image = Nothing

        PanTime.BringToFront()
        PanReminder.BringToFront()

        PicAdd.Tag = "Nhắc việc"

        PicStatus.Tag = "normal"
        PicProcessing.Tag = "normal"
        PicQuickAction.Tag = "normal"
        PicSystem.Tag = "normal"
        PicActionList.Tag = "normal"
        PicSetting.Tag = "normal"
        PicAbout.Tag = "normal"
    End Sub

    Private Sub PicReminder_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PicReminder.MouseDown
        PicReminder.Image = Action_Timer_2.My.Resources.Resource1.MD_REMINDER
    End Sub

    Private Sub PicReminder_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles PicReminder.MouseEnter
        PicReminder.Image = Action_Timer_2.My.Resources.Resource1.MM_TAB

        If ChkHideTooltip.Checked = True Then
            Me.Text = StrDup(10, " ") & "- Bao gồm chức năng nhắc việc và gửi email từ hộp thư (Gmail)" & _
                "    ** Lưu ý: Bạn phải tạo danh sách công việc mới nếu vừa thực hiện " & _
                "xong một danh sách công việc cũ" & StrDup(20, " ")
            TToolTip.Enabled = True
        End If
    End Sub

    Private Sub PicReminder_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles PicReminder.MouseLeave
        If PicReminder.Tag = "normal" Then
            PicReminder.Image = Nothing
        ElseIf PicReminder.Tag = "click" Then
            PicReminder.Image = Action_Timer_2.My.Resources.Resource1.MD_REMINDER
        End If

        TToolTip.Enabled = False
        Me.Text = "Action Timer 2010"
    End Sub
    '___________________________________________________________________________
    '===========================================================================



    '___________________________________________________________________________
    '===========================================================================
    Private Sub PicSystem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PicSystem.Click
        PicSystem.Image = Action_Timer_2.My.Resources.Resource1.MD_SYSTEM
        PicSystem.Tag = "click"

        PicQuickAction.Image = Nothing
        PicProcessing.Image = Nothing
        PicReminder.Image = Nothing
        PicStatus.Image = Nothing

        PicActionList.Image = Nothing
        PicSetting.Image = Nothing
        PicAbout.Image = Nothing

        PanTime.BringToFront()
        PanSystemAction.BringToFront()

        PicAdd.Tag = "Thao tác với hệ thống"

        PicStatus.Tag = "normal"
        PicProcessing.Tag = "normal"
        PicReminder.Tag = "normal"
        PicQuickAction.Tag = "normal"
        PicActionList.Tag = "normal"
        PicSetting.Tag = "normal"
        PicAbout.Tag = "normal"
    End Sub

    Private Sub PicSystem_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PicSystem.MouseDown
        PicSystem.Image = Action_Timer_2.My.Resources.Resource1.MD_SYSTEM
    End Sub

    Private Sub PicSystem_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles PicSystem.MouseEnter
        PicSystem.Image = Action_Timer_2.My.Resources.Resource1.MM_TAB

        If ChkHideTooltip.Checked = True Then
            Me.Text = StrDup(10, " ") & "- Gồm các chức năng về hệ thống như tắt máy, đăng xuất,..." & _
                "    ** Lưu ý: Bạn phải tạo danh sách công việc mới nếu vừa thực hiện" & _
                " xong một danh sách công việc cũ" & StrDup(20, " ")

            TToolTip.Enabled = True
        End If
    End Sub

    Private Sub PicSystem_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles PicSystem.MouseLeave
        If PicSystem.Tag = "normal" Then
            PicSystem.Image = Nothing
        ElseIf PicSystem.Tag = "click" Then
            PicSystem.Image = Action_Timer_2.My.Resources.Resource1.MD_SYSTEM
        End If

        TToolTip.Enabled = False
        Me.Text = "Action Timer 2010"
    End Sub
    '___________________________________________________________________________
    '===========================================================================



    '___________________________________________________________________________
    '===========================================================================
    Private Sub PicActionList_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PicActionList.Click
        PicActionList.Image = Action_Timer_2.My.Resources.Resource1.MD_ACTION_LIST
        PicActionList.Tag = "click"

        PicQuickAction.Image = Nothing
        PicProcessing.Image = Nothing
        PicReminder.Image = Nothing
        PicSystem.Image = Nothing

        PicStatus.Image = Nothing
        PicSetting.Image = Nothing
        PicAbout.Image = Nothing

        PanTime.BringToFront()
        PanActionList.BringToFront()

        PicStatus.Tag = "normal"
        PicProcessing.Tag = "normal"
        PicReminder.Tag = "normal"
        PicSystem.Tag = "normal"
        PicQuickAction.Tag = "normal"
        PicSetting.Tag = "normal"
        PicAbout.Tag = "normal"
    End Sub

    Private Sub PicActionList_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PicActionList.MouseDown
        PicActionList.Image = Action_Timer_2.My.Resources.Resource1.MD_ACTION_LIST
    End Sub

    Private Sub PicActionList_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles PicActionList.MouseEnter
        PicActionList.Image = Action_Timer_2.My.Resources.Resource1.MM_TAB

        If ChkHideTooltip.Checked = True Then
            Me.Text = StrDup(10, " ") & "- Xem, chỉnh sửa danh sách công việc." & _
                " Để bắt đầu, nhấn Tạo mới" & StrDup(93, " ")
            TToolTip.Enabled = True
        End If
    End Sub

    Private Sub PicActionList_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles PicActionList.MouseLeave
        If PicActionList.Tag = "normal" Then
            PicActionList.Image = Nothing
        ElseIf PicActionList.Tag = "click" Then
            PicActionList.Image = Action_Timer_2.My.Resources.Resource1.MD_ACTION_LIST
        End If

        TToolTip.Enabled = False
        Me.Text = "Action Timer 2010"
    End Sub
    '___________________________________________________________________________
    '===========================================================================



    '___________________________________________________________________________
    '===========================================================================
    Private Sub PicSetting_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PicSetting.Click
        PicSetting.Image = Action_Timer_2.My.Resources.Resource1.MD_SETTING
        PicSetting.Tag = "click"

        PicQuickAction.Image = Nothing
        PicProcessing.Image = Nothing
        PicReminder.Image = Nothing
        PicSystem.Image = Nothing

        PicActionList.Image = Nothing
        PicStatus.Image = Nothing
        PicAbout.Image = Nothing

        PanTime.BringToFront()
        PanSetting.BringToFront()

        PicStatus.Tag = "normal"
        PicProcessing.Tag = "normal"
        PicReminder.Tag = "normal"
        PicSystem.Tag = "normal"
        PicActionList.Tag = "normal"
        PicQuickAction.Tag = "normal"
        PicAbout.Tag = "normal"
    End Sub

    Private Sub PicSetting_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PicSetting.MouseDown
        PicSetting.Image = Action_Timer_2.My.Resources.Resource1.MD_SETTING
    End Sub

    Private Sub PicSetting_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles PicSetting.MouseEnter
        PicSetting.Image = Action_Timer_2.My.Resources.Resource1.MM_TAB

        If ChkHideTooltip.Checked = True Then
            Me.Text = StrDup(10, " ") & "- Thêm sự lựa chọn về âm báo, " & _
                "thời gian gia hạn, chạy ẩn, khởi động cùng Windows" & StrDup(56, " ")
            TToolTip.Enabled = True
        End If
    End Sub

    Private Sub PicSetting_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles PicSetting.MouseLeave
        If PicSetting.Tag = "normal" Then
            PicSetting.Image = Nothing
        ElseIf PicSetting.Tag = "click" Then
            PicSetting.Image = Action_Timer_2.My.Resources.Resource1.MD_SETTING
        End If

        TToolTip.Enabled = False
        Me.Text = "Action Timer 2010"
    End Sub
    '___________________________________________________________________________
    '===========================================================================


    '___________________________________________________________________________
    '===========================================================================
    Private Sub PicAbout_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PicAbout.Click
        PicAbout.Image = Action_Timer_2.My.Resources.Resource1.MD_ABOUT
        PicAbout.Tag = "click"

        PicQuickAction.Image = Nothing
        PicProcessing.Image = Nothing
        PicReminder.Image = Nothing
        PicSystem.Image = Nothing

        PicActionList.Image = Nothing
        PicSetting.Image = Nothing
        PicStatus.Image = Nothing

        PanAbout.BringToFront()

        PicStatus.Tag = "normal"
        PicProcessing.Tag = "normal"
        PicReminder.Tag = "normal"
        PicSystem.Tag = "normal"
        PicActionList.Tag = "normal"
        PicSetting.Tag = "normal"
        PicQuickAction.Tag = "normal"
    End Sub

    Private Sub PicAbout_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PicAbout.MouseDown
        PicAbout.Image = Action_Timer_2.My.Resources.Resource1.MD_ABOUT
    End Sub

    Private Sub PicAbout_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles PicAbout.MouseEnter
        PicAbout.Image = Action_Timer_2.My.Resources.Resource1.MM_FUNCTION

        If ChkHideTooltip.Checked = True Then
            Me.Text = StrDup(10, " ") & "- Xem thông tin chi tiết về Action Timer " & _
                "cũng như tác giả" & StrDup(105, " ")
            TToolTip.Enabled = True
        End If
    End Sub

    Private Sub PicAbout_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles PicAbout.MouseLeave
        If PicAbout.Tag = "normal" Then
            PicAbout.Image = Nothing
        ElseIf PicAbout.Tag = "click" Then
            PicAbout.Image = Action_Timer_2.My.Resources.Resource1.MD_ABOUT
        End If

        TToolTip.Enabled = False
        Me.Text = "Action Timer 2010"
    End Sub

    Private Sub FrmMain_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Mod_Config.SaveAllSettingToReg()
    End Sub

    '___________________________________________________________________________
    '===========================================================================
    Private Sub FrmMain_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        'kiểm tra file Action List có chưa, nếu chưa thì tạo mới
        If System.IO.File.Exists(AppPath() & "Action List.ini") = False Then
            Dim str As String = _
                vbCrLf & _
                "[ACTION STATUS]" & vbCrLf & _
                "Total=0" & vbCrLf & _
                "Next=1" & vbCrLf & _
                "StartAt=1" & vbCrLf & _
                "Pause=False" & vbCrLf & _
                "Continue=False" & vbCrLf & vbCrLf & vbCrLf & _
                "[Success]" & vbCrLf & _
                "Count=0" & vbCrLf & vbCrLf & vbCrLf & _
                "[Error]" & vbCrLf & _
                "Count=0" & vbCrLf & vbCrLf & vbCrLf & _
                "[Skip]" & vbCrLf & _
                "Count=0" & vbCrLf & _
                "SkipAt=0" & vbCrLf & vbCrLf & vbCrLf & _
                "_____________________________________________________________________________________________________________" & vbCrLf & _
                "[ACTION LIST]" & vbCrLf & _
                "Count=0" & vbCrLf

            Mod_Reminder.WriteText(str, AppPath() & "Action List.ini")
        End If

        'load dữ liệu từ REG
        Mod_Config.LoadAllSettingFromReg()

        'chuẩn bị chạy chương trình
        Mod_Config.PreparingWorking()

        PanStatus.Top = 31
        PanStatus.Left = 137

        PanQuickAction.Top = 31
        PanQuickAction.Left = 137

        PanProcess.Top = 31
        PanProcess.Left = 137

        PanReminder.Top = 31
        PanReminder.Left = 137

        PanSystemAction.Top = 31
        PanSystemAction.Left = 137

        PanActionList.Top = 31
        PanActionList.Left = 137

        PanSetting.Top = 31
        PanSetting.Left = 137

        PanAbout.Top = 31
        PanAbout.Left = 137

        PanTime.Top = 443
        PanTime.Left = 137

        PicStatus_Click(sender, e)

        'kiểm tra folder Reminder có chưa, nếu chưa thì tạo mới
        If FolderExists(AppPath() & "Reminder\") = False Then
            CreateNewFolder(AppPath, "Reminder")
        End If

        LoadActionList(AppPath() & "Action List.ini", DvActionList)

        LblVersion.Text = "Phiên bản : " & My.Application.Info.Version.Major & _
            "." & My.Application.Info.Version.Minor & _
            "." & My.Application.Info.Version.Build & _
            "." & My.Application.Info.Version.Revision

        TxtDay.Text = Now.Day & "/" & Now.Month & "/" & Now.Year

    End Sub


    Private Sub TTime_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TTime.Tick
        LblTime.Text = IIf(Len(CStr(TimeOfDay.Hour)) = 1, "0" & CStr(TimeOfDay.Hour), _
                           CStr(TimeOfDay.Hour)) & _
                           ":" & IIf(Len(CStr(TimeOfDay.Minute)) = 1, "0" & _
                            CStr(TimeOfDay.Minute), CStr(TimeOfDay.Minute)) & _
                            ":" & IIf(Len(CStr(TimeOfDay.Second)) = 1, "0" & _
                            CStr(TimeOfDay.Second), CStr(TimeOfDay.Second))

        LblDate.Text = "Ngày " & IIf(Len(CStr(DateTime.Now.Day)) = 1, "0" & DateTime.Now.Day, DateTime.Now.Day) & _
                    " Tháng " & IIf(Len(CStr(DateTime.Now.Month)) = 1, "0" & DateTime.Now.Month, DateTime.Now.Month) & _
                    " Năm " & CStr(DateTime.Now.Year)



        

    End Sub




    '___________________________________________________________________________
    '===========================================================================
    Private Sub PicShutdown_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PicShutdown.MouseDown
        PicShutdown.Image = Action_Timer_2.My.Resources.Resource1.MD_SHUT_DOWN
    End Sub

    Private Sub PicShutdown_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles PicShutdown.MouseEnter
        PicShutdown.Image = Action_Timer_2.My.Resources.Resource1.MM_shut_down
    End Sub

    Private Sub PicShutdown_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles PicShutdown.MouseLeave
        PicShutdown.Image = Action_Timer_2.My.Resources.Resource1.N_shut_down
    End Sub

    Private Sub PicShutdown_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PicShutdown.MouseUp
        PicShutdown.Image = Action_Timer_2.My.Resources.Resource1.MM_shut_down
    End Sub


    '___________________________________________________________________________
    '===========================================================================
    Private Sub PicRestart_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PicRestart.MouseDown
        PicRestart.Image = Action_Timer_2.My.Resources.Resource1.MD_RESTART
    End Sub

    Private Sub PicRestart_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles PicRestart.MouseEnter
        PicRestart.Image = Action_Timer_2.My.Resources.Resource1.MM_restart
    End Sub

    Private Sub PicRestart_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles PicRestart.MouseLeave
        PicRestart.Image = Action_Timer_2.My.Resources.Resource1.N_restart
    End Sub

    Private Sub PicRestart_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PicRestart.MouseUp
        PicRestart.Image = Action_Timer_2.My.Resources.Resource1.MM_restart
    End Sub

    '___________________________________________________________________________
    '===========================================================================
    Private Sub PicLogOff_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PicLogOff.MouseDown
        PicLogOff.Image = Action_Timer_2.My.Resources.Resource1.MD_LOGOFF
    End Sub

    Private Sub PicLogOff_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles PicLogOff.MouseEnter
        PicLogOff.Image = Action_Timer_2.My.Resources.Resource1.MM_log_off
    End Sub

    Private Sub PicLogOff_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles PicLogOff.MouseLeave
        PicLogOff.Image = Action_Timer_2.My.Resources.Resource1.N_log_off
    End Sub

    Private Sub PicLogOff_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PicLogOff.MouseUp
        PicLogOff.Image = Action_Timer_2.My.Resources.Resource1.MM_log_off
    End Sub

    '___________________________________________________________________________
    '===========================================================================
    Private Sub PicStandBy_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PicStandBy.MouseDown
        PicStandBy.Image = Action_Timer_2.My.Resources.Resource1.MD_STAND_BY
    End Sub

    Private Sub PicStandBy_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles PicStandBy.MouseEnter
        PicStandBy.Image = Action_Timer_2.My.Resources.Resource1.MM_stand_by
    End Sub

    Private Sub PicStandBy_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles PicStandBy.MouseLeave
        PicStandBy.Image = Action_Timer_2.My.Resources.Resource1.N_stand_by
    End Sub

    Private Sub PicStandBy_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PicStandBy.MouseUp
        PicStandBy.Image = Action_Timer_2.My.Resources.Resource1.MM_stand_by
    End Sub

    '___________________________________________________________________________
    '===========================================================================
    Private Sub Pichibernate_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PicHibernate.MouseDown
        PicHibernate.Image = Action_Timer_2.My.Resources.Resource1.MD_HIBERNATE
    End Sub

    Private Sub Pichibernate_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles PicHibernate.MouseEnter
        PicHibernate.Image = Action_Timer_2.My.Resources.Resource1.MM_hibernate
    End Sub

    Private Sub Pichibernate_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles PicHibernate.MouseLeave
        PicHibernate.Image = Action_Timer_2.My.Resources.Resource1.N_hibernate
    End Sub

    Private Sub Pichibernate_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PicHibernate.MouseUp
        PicHibernate.Image = Action_Timer_2.My.Resources.Resource1.MM_hibernate
    End Sub

    '___________________________________________________________________________
    '===========================================================================
    Private Sub Piclockcomputer_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PicLockComputer.MouseDown
        PicLockComputer.Image = Action_Timer_2.My.Resources.Resource1.MD_LOCK_COMPUTER
    End Sub

    Private Sub Piclockcomputer_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles PicLockComputer.MouseEnter
        PicLockComputer.Image = Action_Timer_2.My.Resources.Resource1.MM_lock_COMPUTER
    End Sub

    Private Sub Piclockcomputer_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles PicLockComputer.MouseLeave
        PicLockComputer.Image = Action_Timer_2.My.Resources.Resource1.N_lock_COMPUTER
    End Sub

    Private Sub Piclockcomputer_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PicLockComputer.MouseUp
        PicLockComputer.Image = Action_Timer_2.My.Resources.Resource1.MM_lock_COMPUTER
    End Sub

    '___________________________________________________________________________
    '===========================================================================
    Private Sub Picturnoffmonitor_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PicTurnOffMonitor.MouseDown
        PicTurnOffMonitor.Image = Action_Timer_2.My.Resources.Resource1.MD_TURN_OFF_MONITOR
    End Sub

    Private Sub Picturnoffmonitor_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles PicTurnOffMonitor.MouseEnter
        PicTurnOffMonitor.Image = Action_Timer_2.My.Resources.Resource1.MM_turn_off_monitor
    End Sub

    Private Sub Picturnoffmonitor_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles PicTurnOffMonitor.MouseLeave
        PicTurnOffMonitor.Image = Action_Timer_2.My.Resources.Resource1.N_turn_off_monitor
    End Sub

    Private Sub Picturnoffmonitor_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PicTurnOffMonitor.MouseUp
        PicTurnOffMonitor.Image = Action_Timer_2.My.Resources.Resource1.MM_turn_off_monitor
    End Sub

    '___________________________________________________________________________
    '===========================================================================
    Private Sub PicScreenSaver_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PicScreenSaver.MouseDown
        PicScreenSaver.Image = Action_Timer_2.My.Resources.Resource1.MD_SCREEN_SAVER
    End Sub

    Private Sub PicScreenSaver_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles PicScreenSaver.MouseEnter
        PicScreenSaver.Image = Action_Timer_2.My.Resources.Resource1.MM_Screen_Saver
    End Sub

    Private Sub PicScreenSaver_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles PicScreenSaver.MouseLeave
        PicScreenSaver.Image = Action_Timer_2.My.Resources.Resource1.N_Screen_Saver
    End Sub

    Private Sub PicScreenSaver_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PicScreenSaver.MouseUp
        PicScreenSaver.Image = Action_Timer_2.My.Resources.Resource1.MM_Screen_Saver
    End Sub

    '___________________________________________________________________________
    '===========================================================================
    Private Sub PicEmptyRecycle_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PicEmptyRecycle.MouseDown
        PicEmptyRecycle.Image = Action_Timer_2.My.Resources.Resource1.MD_EMPTY_RECYCLE
    End Sub

    Private Sub PicEmptyRecycle_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles PicEmptyRecycle.MouseEnter
        PicEmptyRecycle.Image = Action_Timer_2.My.Resources.Resource1.MM_Empty_Recycle
    End Sub

    Private Sub PicEmptyRecycle_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles PicEmptyRecycle.MouseLeave
        PicEmptyRecycle.Image = Action_Timer_2.My.Resources.Resource1.N_Empty_Recycle
    End Sub

    Private Sub PicEmptyRecycle_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PicEmptyRecycle.MouseUp
        PicEmptyRecycle.Image = Action_Timer_2.My.Resources.Resource1.MM_Empty_Recycle
    End Sub

    
    Private Sub BtnRefresh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnRefresh.Click
        Mod_Process.ListProcess(DViewProcess)
    End Sub


    Private Sub BtnEndProcess_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnEndProcess.Click
        Mod_Process.KillProcess(DViewProcess)
    End Sub

    Private Sub DViewProcess_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DViewProcess.CellClick
        Mod_Process.DataViewCheckBox(DViewProcess, e)
    End Sub


    Private Sub BtnClearReminder_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnClearReminder.Click
        TxtReminder.Text = vbNullString
    End Sub

    Private Sub BtnOpenReminder_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnOpenReminder.Click
        Try

            Dim FileName As String
            FileName = Application.StartupPath & "\Reminder_Temp"
            FileName = FileName.Replace("\\", "\")

            If FileLen(FileName) <> 0 Then
                TxtReminder.Text = ReadText(FileName)
            End If

        Catch ex As Exception
        End Try
    End Sub

    Private Sub BtnSaveReminder_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnSaveReminder.Click
        Dim FileName As String
        FileName = Application.StartupPath & "\Reminder_Temp"
        FileName = FileName.Replace("\\", "\")

        WriteText(TxtReminder.Text, FileName)

    End Sub

  
    


    Private Sub PicShutdown_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PicShutdown.Click
        If ChkConfirm.Checked = True Then
            If MessageBox.Show("Bạn muốn Tắt máy ?" & vbCrLf & _
                                "Nhấn YES - đồng ý, NO - không đồng ý", _
                                "Xác nhận việc Tắt máy", _
                                MessageBoxButtons.YesNo, _
                                MessageBoxIcon.Question) = _
                                Windows.Forms.DialogResult.Yes Then
                Mod_System.power_Shutdown()
            End If
        Else
            Mod_System.power_Shutdown()
        End If

    End Sub

    Private Sub PicRestart_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PicRestart.Click
        If ChkConfirm.Checked = True Then
            If MessageBox.Show("Bạn muốn Khởi động lại máy tính ?" & vbCrLf & _
                                "Nhấn YES - đồng ý, NO - không đồng ý", _
                                "Xác nhận việc Khởi động lại máy tính", _
                                MessageBoxButtons.YesNo, _
                                MessageBoxIcon.Question) = _
                                Windows.Forms.DialogResult.Yes Then
                Mod_System.power_Restart()
            End If
        Else
            Mod_System.power_Restart()
        End If
    End Sub

    Private Sub PicLogOff_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PicLogOff.Click
        If ChkConfirm.Checked = True Then
            If MessageBox.Show("Bạn muốn Đăng xuất ?" & vbCrLf & _
                                "Nhấn YES - đồng ý, NO - không đồng ý", _
                                "Xác nhận việc Đăng xuất", _
                                MessageBoxButtons.YesNo, _
                                MessageBoxIcon.Question) = _
                                Windows.Forms.DialogResult.Yes Then
                Mod_System.power_LogOff()
            End If
        Else
            Mod_System.power_LogOff()
        End If
    End Sub

    Private Sub PicStandBy_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PicStandBy.Click
        If ChkConfirm.Checked = True Then
            If MessageBox.Show("Bạn muốn Chuyển sang chế độ chờ ?" & vbCrLf & _
                                "Nhấn YES - đồng ý, NO - không đồng ý", _
                                "Xác nhận việc Chuyển sang chế độ chờ", _
                                MessageBoxButtons.YesNo, _
                                MessageBoxIcon.Question) = _
                                Windows.Forms.DialogResult.Yes Then
                Mod_System.power_StandBy()
            End If
        Else
            Mod_System.power_StandBy()
        End If
    End Sub

    Private Sub PicHibernate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PicHibernate.Click
        If ChkConfirm.Checked = True Then
            If MessageBox.Show("Bạn muốn Chuyển sang chế độ ngủ đông ?" & vbCrLf & _
                                "Nhấn YES - đồng ý, NO - không đồng ý", _
                                "Xác nhận việc Chuyển sang chế độ ngủ đông", _
                                MessageBoxButtons.YesNo, _
                                MessageBoxIcon.Question) = _
                                Windows.Forms.DialogResult.Yes Then
                Mod_System.power_Hibernate()
            End If
        Else
            Mod_System.power_Hibernate()
        End If
    End Sub

    Private Sub PicLockComputer_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PicLockComputer.Click
        If ChkConfirm.Checked = True Then
            If MessageBox.Show("Bạn muốn Khoá máy tính ?" & vbCrLf & _
                                "Nhấn YES - đồng ý, NO - không đồng ý", _
                                "Xác nhận việc Khoá máy tính", _
                                MessageBoxButtons.YesNo, _
                                MessageBoxIcon.Question) = _
                                Windows.Forms.DialogResult.Yes Then
                Mod_System.power_LockComputer()
            End If
        Else
            Mod_System.power_LockComputer()
        End If
    End Sub

    Private Sub PicTurnOffMonitor_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PicTurnOffMonitor.Click
        If ChkConfirm.Checked = True Then
            If MessageBox.Show("Bạn muốn Tắt màn hình ?" & vbCrLf & _
                                "Nhấn YES - đồng ý, NO - không đồng ý", _
                                "Xác nhận việc Tắt màn hình", _
                                MessageBoxButtons.YesNo, _
                                MessageBoxIcon.Question) = _
                                Windows.Forms.DialogResult.Yes Then
                Mod_System.power_TurnOffMonitor()
            End If
        Else
            Mod_System.power_TurnOffMonitor()
        End If

    End Sub

    Private Sub PicScreenSaver_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PicScreenSaver.Click
        If ChkConfirm.Checked = True Then
            If MessageBox.Show("Bạn muốn Chạy trình bảo vệ màn hình ?" & vbCrLf & _
                                "Nhấn YES - đồng ý, NO - không đồng ý", _
                                "Xác nhận việc Chạy trình bảo vệ màn hình", _
                                MessageBoxButtons.YesNo, _
                                MessageBoxIcon.Question) = _
                                Windows.Forms.DialogResult.Yes Then
                Mod_System.power_RunScreenSaver()
            End If
        Else
            Mod_System.power_RunScreenSaver()
        End If
    End Sub

    Private Sub PicEmptyRecycle_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PicEmptyRecycle.Click
        If ChkConfirm.Checked = True Then
            If MessageBox.Show("Bạn muốn Dọn dẹp thùng rác ?" & vbCrLf & _
                                "Nhấn YES - đồng ý, NO - không đồng ý", _
                                "Xác nhận việc Dọn dẹp thùng rác", _
                                MessageBoxButtons.YesNo, _
                                MessageBoxIcon.Question) = _
                                Windows.Forms.DialogResult.Yes Then
                Mod_System.EmptyRecycleBin()
            End If
        Else
            Mod_System.EmptyRecycleBin()
        End If
    End Sub


    Private Sub TxtQuickRun_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtQuickRun.KeyDown
        If e.KeyCode = 13 Then
            TxtQuickRun.Text = Trim(TxtQuickRun.Text)

            If TxtQuickRun.Text <> vbNullString Then

                Dim cmd As String = TxtQuickRun.Text
                Dim fName() As String

                fName = cmd.Split(" ")
                cmd = ""

                For i As Int16 = 1 To fName.Length - 1
                    cmd &= " " & fName(i)
                Next

                Try
                    System.Diagnostics.Process.Start(fName(0), cmd)
                Catch ex As Exception

                    If Err.Number = 5 Then
                        MessageBox.Show("Chương trình không tìm thấy '" & TxtQuickRun.Text & "'" & _
                            vbCrLf & "Hãy chắc rằng bạn đã gõ đúng tên file và sau đó nhập lại" & vbCrLf & _
                            "Thông tin cụ thể về lỗi này :" & vbCrLf & vbCrLf & _
                            ex.ToString, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error)

                    End If

                End Try
            End If

        End If
    End Sub


    Private Sub BtnOpenRun_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnOpenRun.Click
        DlgOpen.FileName = vbNullString
        DlgOpen.Filter = "All file (*.*)|*.*"

        DlgOpen.Multiselect = False
        DlgOpen.ShowDialog()

        If DlgOpen.CheckFileExists = True Then
            TxtRun.Text = Trim(DlgOpen.FileName)
        Else
            MessageBox.Show("Tập tin không tồn tại", "Lỗi chọn tập tin", _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If

    End Sub

    Private Sub BtnUpdateList_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnUpdateList.Click
        Mod_ActionList.LoadActionList(AppPath() & "Action List.ini", DvActionList)
    End Sub


    Private Sub PicUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PicUpdate.Click

        Dim s As String = CboLoai.Tag
        Dim str() As String = s.Split(",")

        PanActionListUpdate.Visible = False
        DvActionList.Item(CInt(str(0)), CInt(str(1))).Value = CboLoai.Text

        s = CboHanhDong.Tag
        str = s.Split(",")

        If CboLoai.Text = "Thao tác với hệ thống" Then
            DvActionList.Item(CInt(str(0)), CInt(str(1))).Value = CboHanhDong.Text

        ElseIf CboLoai.Text = "Nhắc việc" Then
            Dim _v As String    'tên file, vd: 05~06~2010_09~38~42
            _v = CStr(DvActionList.Item(2, CInt(str(1))).Value) & "_" & _
                CStr(DvActionList.Item(3, CInt(str(1))).Value)
            _v = _v.Replace("/", "~")
            _v = _v.Replace(":", "~")

            If CboHanhDong.Text = "Gửi email" Then
                'lưu nội dung email
                Mod_Reminder.WriteText(TxtMieuTa.Text, Mod_ActionList.AppPath & "Reminder\" & _v & ".mail")
                'ghi lại giá trị ô
                DvActionList.Item(CInt(str(0)), CInt(str(1))).Value = CboHanhDong.Text & " " & Mod_ActionList.AppPath & "Reminder\" & _v & ".mail"
            Else
                Mod_Reminder.WriteText(TxtMieuTa.Text, Mod_ActionList.AppPath & "Reminder\" & _v & ".txt")
                'ghi lại giá trị ô
                DvActionList.Item(CInt(str(0)), CInt(str(1))).Value = CboHanhDong.Text & " " & Mod_ActionList.AppPath & "Reminder\" & _v & ".txt"
            End If

        Else
            DvActionList.Item(CInt(str(0)), CInt(str(1))).Value = CboHanhDong.Text & " " & Trim(TxtMieuTa.Text)

        End If

        Erase str
    End Sub

    Private Sub PicUpdate_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PicUpdate.MouseDown
        PicUpdate.Image = My.Resources.Resource1.MD_UPDATE
    End Sub

    Private Sub PicUpdate_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles PicUpdate.MouseEnter
        PicUpdate.Image = My.Resources.Resource1.MM_UPDATE
    End Sub

    Private Sub PicUpdate_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles PicUpdate.MouseLeave
        PicUpdate.Image = My.Resources.Resource1.N_UPDATE
    End Sub

    Private Sub PicUpdate_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PicUpdate.MouseUp
        PicUpdate.Image = My.Resources.Resource1.MM_UPDATE
    End Sub


    Private Sub CboLoai_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CboLoai.SelectedValueChanged
        CboHanhDong.Items.Clear()
        Select Case CboLoai.Text
            Case "Quản lý tiến trình"
                CboHanhDong.Enabled = True
                CboHanhDong.Items.Add("Chạy")
                CboHanhDong.Items.Add("Tắt tiến trình")
                TxtMieuTa.Enabled = True
                TxtMieuTa.Text = "Nhập tên tiến trình/tập tin"
            Case "Nhắc việc"
                CboHanhDong.Enabled = True
                CboHanhDong.Items.Add("Hiển thị thông báo nhắc việc")
                CboHanhDong.Items.Add("Gửi email")
                TxtMieuTa.Enabled = True
                TxtMieuTa.Text = "Biên soạn nội dung email theo mẫu :" & vbCrLf & _
                    "[EMail]" & vbCrLf & _
                    "To=..." & vbCrLf & _
                    "From=...@gmail.com" & vbCrLf & _
                    "Account=..." & vbCrLf & _
                    "Password=..." & vbCrLf & _
                    "Smtp = smtp.gmail.com" & vbCrLf & vbCrLf & _
                    "[Subject]...[Subject]" & vbCrLf & _
                    "[Body]...[Body]"

            Case "Thao tác với hệ thống"
                CboHanhDong.Enabled = True
                CboHanhDong.Items.Add("Tắt máy tính")
                CboHanhDong.Items.Add("Khởi động lại máy tính")
                CboHanhDong.Items.Add("Đăng xuất khỏi tài khoản")
                CboHanhDong.Items.Add("Chuyển sang chế độ chờ")
                CboHanhDong.Items.Add("Chuyển sang chế độ ngủ đông")
                CboHanhDong.Items.Add("Bật màn hình máy tính")
                CboHanhDong.Items.Add("Tắt màn hình máy tính")
                CboHanhDong.Items.Add("Khoá máy tính")
                TxtMieuTa.Enabled = False
        End Select
    End Sub

    Private Sub DvActionList_CellDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DvActionList.CellDoubleClick
        Try
            If e.ColumnIndex = 5 And e.RowIndex > -1 And _
                DvActionList.Rows(e.RowIndex).ReadOnly = False Then

                PanActionListUpdate.Visible = True
                CboHanhDong.Tag = CStr(e.ColumnIndex) & "," & CStr(e.RowIndex)
                CboLoai.Tag = CStr(e.ColumnIndex - 1) & "," & CStr(e.RowIndex)
                LblEditActionList.Text = "Chỉnh sửa nội dung : " & DvActionList.Item(e.ColumnIndex, e.RowIndex).Value
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub PicCancelEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PicCancelEdit.Click
        PanActionListUpdate.Visible = False
        CboLoai.Text = ""
        CboHanhDong.Text = ""
        CboHanhDong.Enabled = False
        TxtMieuTa.Text = ""
        TxtMieuTa.Enabled = False
    End Sub

    Private Sub PicCancelEdit_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PicCancelEdit.MouseDown
        PicCancelEdit.Image = My.Resources.Resource1.MD_CANCEL
    End Sub

    Private Sub PicCancelEdit_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles PicCancelEdit.MouseEnter
        PicCancelEdit.Image = My.Resources.Resource1.MM_CANCEL
    End Sub

    Private Sub PicCancelEdit_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles PicCancelEdit.MouseLeave
        PicCancelEdit.Image = My.Resources.Resource1.N_CANCEL
    End Sub

    Private Sub PicCancelEdit_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PicCancelEdit.MouseUp
        PicCancelEdit.Image = My.Resources.Resource1.MM_CANCEL
    End Sub


    Private Sub OptSysShutdown_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles OptSysShutdown.MouseDown
        OptSysShutdown.Image = My.Resources.Resource1.MD_SHUT_DOWN
    End Sub

    Private Sub OptSysShutdown_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles OptSysShutdown.MouseEnter
        OptSysShutdown.Image = My.Resources.Resource1.MM_shut_down
    End Sub

    Private Sub OptSysShutdown_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles OptSysShutdown.MouseLeave
        OptSysShutdown.Image = My.Resources.Resource1.N_shut_down
    End Sub

    Private Sub OptSysShutdown_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles OptSysShutdown.MouseUp
        OptSysShutdown.Image = My.Resources.Resource1.MM_shut_down
    End Sub

    Private Sub OptSysRestart_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles OptSysRestart.MouseDown
        OptSysRestart.Image = My.Resources.Resource1.MD_RESTART
    End Sub

    Private Sub OptSysRestart_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles OptSysRestart.MouseEnter
        OptSysRestart.Image = My.Resources.Resource1.MM_restart
    End Sub

    Private Sub OptSysRestart_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles OptSysRestart.MouseLeave
        OptSysRestart.Image = My.Resources.Resource1.N_restart
    End Sub

    Private Sub OptSysRestart_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles OptSysRestart.MouseUp
        OptSysRestart.Image = My.Resources.Resource1.MM_restart
    End Sub

    Private Sub OptSysLogOff_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles OptSysLogOff.MouseDown
        OptSysLogOff.Image = My.Resources.Resource1.MD_LOGOFF
    End Sub

    Private Sub OptSysLogOff_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles OptSysLogOff.MouseEnter
        OptSysLogOff.Image = My.Resources.Resource1.MM_log_off
    End Sub

    Private Sub OptSysLogOff_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles OptSysLogOff.MouseLeave
        OptSysLogOff.Image = My.Resources.Resource1.N_log_off
    End Sub

    Private Sub OptSysLogOff_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles OptSysLogOff.MouseUp
        OptSysLogOff.Image = My.Resources.Resource1.MM_log_off
    End Sub

    Private Sub OptSysStandBy_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles OptSysStandBy.MouseDown
        OptSysStandBy.Image = My.Resources.Resource1.MD_STAND_BY
    End Sub

    Private Sub OptSysStandBy_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles OptSysStandBy.MouseEnter
        OptSysStandBy.Image = My.Resources.Resource1.MM_stand_by
    End Sub

    Private Sub OptSysStandBy_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles OptSysStandBy.MouseLeave
        OptSysStandBy.Image = My.Resources.Resource1.N_stand_by
    End Sub

    Private Sub OptSysStandBy_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles OptSysStandBy.MouseUp
        OptSysStandBy.Image = My.Resources.Resource1.MM_stand_by
    End Sub

    Private Sub OptSysHibernate_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles OptSysHibernate.MouseDown
        OptSysHibernate.Image = My.Resources.Resource1.MD_HIBERNATE
    End Sub

    Private Sub OptSysHibernate_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles OptSysHibernate.MouseEnter
        OptSysHibernate.Image = My.Resources.Resource1.MM_hibernate
    End Sub

    Private Sub OptSysHibernate_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles OptSysHibernate.MouseLeave
        OptSysHibernate.Image = My.Resources.Resource1.N_hibernate
    End Sub

    Private Sub OptSysHibernate_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles OptSysHibernate.MouseUp
        OptSysHibernate.Image = My.Resources.Resource1.MM_hibernate
    End Sub

    Private Sub OptSysTurnOffMonitor_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles OptSysTurnOffMonitor.MouseDown
        OptSysTurnOffMonitor.Image = My.Resources.Resource1.MD_TURN_OFF_MONITOR
    End Sub

    Private Sub OptSysTurnOffMonitor_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles OptSysTurnOffMonitor.MouseEnter
        OptSysTurnOffMonitor.Image = My.Resources.Resource1.MM_turn_off_monitor
    End Sub

    Private Sub OptSysTurnOffMonitor_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles OptSysTurnOffMonitor.MouseLeave
        OptSysTurnOffMonitor.Image = My.Resources.Resource1.N_turn_off_monitor
    End Sub

    Private Sub OptSysTurnOffMonitor_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles OptSysTurnOffMonitor.MouseUp
        OptSysTurnOffMonitor.Image = My.Resources.Resource1.MM_turn_off_monitor
    End Sub

    Private Sub OptSysTurnOnMonitor_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles OptSysTurnOnMonitor.MouseDown
        OptSysTurnOnMonitor.Image = My.Resources.Resource1.MD_TURN_ON_MONITOR
    End Sub

    Private Sub OptSysTurnOnMonitor_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles OptSysTurnOnMonitor.MouseEnter
        OptSysTurnOnMonitor.Image = My.Resources.Resource1.MM_TURN_ON_MONITOR
    End Sub

    Private Sub OptSysTurnOnMonitor_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles OptSysTurnOnMonitor.MouseLeave
        OptSysTurnOnMonitor.Image = My.Resources.Resource1.N_TURN_ON_MONITOR
    End Sub

    Private Sub OptSysTurnOnMonitor_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles OptSysTurnOnMonitor.MouseUp
        OptSysTurnOnMonitor.Image = My.Resources.Resource1.MM_TURN_ON_MONITOR
    End Sub


    Private Sub ChkSetSound_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChkSetSound.CheckedChanged
        If ChkSetSound.Checked = True Then
            TxtSetPro.Enabled = True
            TxtSetRem.Enabled = True
            TxtSetSys.Enabled = True

            BtnBrowsePro.Enabled = True
            BtnBrowseRem.Enabled = True
            BtnBrowseSys.Enabled = True

            Vol1.Enabled = True
            PicAudioBack.Enabled = True
            PicAudioPlay.Enabled = True
            PicAudioStop.Enabled = True
            PicAudioNext.Enabled = True
        Else
            TxtSetPro.Enabled = False
            TxtSetRem.Enabled = False
            TxtSetSys.Enabled = False

            BtnBrowsePro.Enabled = False
            BtnBrowseRem.Enabled = False
            BtnBrowseSys.Enabled = False

            Vol1.Enabled = False
            PicAudioBack.Enabled = False
            PicAudioPlay.Enabled = False
            PicAudioStop.Enabled = False
            PicAudioNext.Enabled = False
        End If
    End Sub

    Private Sub Vol1_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Vol1.ValueChanged
        Try
            Mod_Audio.setVolume(CInt(Vol1.Value))
        Catch ex As Exception
        End Try
    End Sub

    Private Sub PicAudioBack_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PicAudioBack.MouseDown
        PicAudioBack.Image = My.Resources.Resource1.audio_md_back
    End Sub

    Private Sub PicAudioBack_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles PicAudioBack.MouseEnter
        PicAudioBack.Image = My.Resources.Resource1.audio_mm_back
    End Sub

    Private Sub PicAudioBack_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles PicAudioBack.MouseLeave
        PicAudioBack.Image = My.Resources.Resource1.audio_n_back
    End Sub

    Private Sub PicAudioBack_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PicAudioBack.MouseUp
        PicAudioBack.Image = My.Resources.Resource1.audio_mm_back
    End Sub

    Private Sub PicAudioPlay_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PicAudioPlay.MouseDown
        PicAudioPlay.Image = My.Resources.Resource1.audio_md_play
    End Sub

    Private Sub PicAudioPlay_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles PicAudioPlay.MouseEnter
        PicAudioPlay.Image = My.Resources.Resource1.audio_mm_play
    End Sub

    Private Sub PicAudioPlay_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles PicAudioPlay.MouseLeave
        PicAudioPlay.Image = My.Resources.Resource1.audio_n_play
    End Sub

    Private Sub PicAudioPlay_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PicAudioPlay.MouseUp
        PicAudioPlay.Image = My.Resources.Resource1.audio_mm_play
    End Sub

    Private Sub PicAudioStop_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PicAudioStop.MouseDown
        PicAudioStop.Image = My.Resources.Resource1.audio_md_stop
    End Sub

    Private Sub PicAudioStop_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles PicAudioStop.MouseEnter
        PicAudioStop.Image = My.Resources.Resource1.audio_mm_stop
    End Sub

    Private Sub PicAudioStop_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles PicAudioStop.MouseLeave
        PicAudioStop.Image = My.Resources.Resource1.audio_n_stop
    End Sub

    Private Sub PicAudioStop_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PicAudioStop.MouseUp
        PicAudioStop.Image = My.Resources.Resource1.audio_mm_stop
    End Sub

    Private Sub PicAudioNext_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PicAudioNext.MouseDown
        PicAudioNext.Image = My.Resources.Resource1.audio_md_next
    End Sub

    Private Sub PicAudioNext_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles PicAudioNext.MouseEnter
        PicAudioNext.Image = My.Resources.Resource1.audio_mm_next
    End Sub

    Private Sub PicAudioNext_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles PicAudioNext.MouseLeave
        PicAudioNext.Image = My.Resources.Resource1.audio_n_next
    End Sub

    Private Sub PicAudioNext_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PicAudioNext.MouseUp
        PicAudioNext.Image = My.Resources.Resource1.audio_mm_next
    End Sub


    Private Sub BtnBrowsePro_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnBrowsePro.Click
        DlgOpen.FileName = ""
        DlgOpen.Filter = "All File (Audio and video only)|*.flac;*.flv;*.asx;*.wax;*.m3u;*.wpl;*.wvx;*.wmx;*.dvr-ms;*.mid;*.rmi;*.midi;*.mpeg;*.mpg;*.m1v;*.mp2;*.mpa;*.mpe;*.mp4;*.ogg;*.ogm;*.ts;*.m2ts;*.wav;*.snd;*.au;*.aif;*.aifc;*.aiff;*.wma;*.mp3;*.asf;*.wm;*.wmv;*.wmd;*.avi;*.DAT"

        If DlgOpen.ShowDialog() = Windows.Forms.DialogResult.OK Then
            TxtSetPro.Text = DlgOpen.FileName
        End If
    End Sub

    Private Sub BtnBrowseRem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnBrowseRem.Click
        DlgOpen.FileName = ""
        DlgOpen.Filter = "All File (Audio and video only)|*.flac;*.flv;*.asx;*.wax;*.m3u;*.wpl;*.wvx;*.wmx;*.dvr-ms;*.mid;*.rmi;*.midi;*.mpeg;*.mpg;*.m1v;*.mp2;*.mpa;*.mpe;*.mp4;*.ogg;*.ogm;*.ts;*.m2ts;*.wav;*.snd;*.au;*.aif;*.aifc;*.aiff;*.wma;*.mp3;*.asf;*.wm;*.wmv;*.wmd;*.avi;*.DAT"

        If DlgOpen.ShowDialog() = Windows.Forms.DialogResult.OK Then
            TxtSetRem.Text = DlgOpen.FileName
        End If
    End Sub

    Private Sub BtnBrowseSys_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnBrowseSys.Click
        DlgOpen.FileName = ""
        DlgOpen.Filter = "All File (Audio and video only)|*.flac;*.flv;*.asx;*.wax;*.m3u;*.wpl;*.wvx;*.wmx;*.dvr-ms;*.mid;*.rmi;*.midi;*.mpeg;*.mpg;*.m1v;*.mp2;*.mpa;*.mpe;*.mp4;*.ogg;*.ogm;*.ts;*.m2ts;*.wav;*.snd;*.au;*.aif;*.aifc;*.aiff;*.wma;*.mp3;*.asf;*.wm;*.wmv;*.wmd;*.avi;*.DAT"

        If DlgOpen.ShowDialog() = Windows.Forms.DialogResult.OK Then
            TxtSetSys.Text = DlgOpen.FileName
        End If
    End Sub

    Private Sub PicAudioPlay_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PicAudioPlay.Click

        If TxtSetPro.Text = "" And TxtSetRem.Text = "" And TxtSetSys.Text = "" Then Exit Sub

        Select Case PicAudioPlay.Tag
            Case "Pro"
                PicAudioPlay.Tag = "Rem"
                If Trim(TxtSetPro.Text) <> "" Then
                    Mod_Audio.PlayAudio(Trim(TxtSetPro.Text))

                Else
                    PicaudioNext_click(sender, e)
                End If

            Case "Rem"
                PicAudioPlay.Tag = "Sys"
                If Trim(TxtSetRem.Text) <> "" Then
                    Mod_Audio.PlayAudio(Trim(TxtSetRem.Text))

                Else
                    PicaudioNext_click(sender, e)
                End If

            Case "Sys"
                PicAudioPlay.Tag = "Pro"
                If Trim(TxtSetSys.Text) <> "" Then
                    Mod_Audio.PlayAudio(Trim(TxtSetSys.Text))

                Else
                    PicaudioNext_click(sender, e)
                End If

        End Select
        Mod_Audio.setVolume(CInt(Vol1.Value))
    End Sub

    Private Sub PicAudioStop_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PicAudioStop.Click
        Mod_Audio.StopAudio()
    End Sub

    Private Sub PicAudioNext_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PicAudioNext.Click
        PicAudioPlay_Click(sender, e)
    End Sub

    Private Sub PicAudioBack_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PicAudioBack.Click
        'pro -> rem -> sys
        If TxtSetPro.Text = "" And TxtSetRem.Text = "" And TxtSetSys.Text = "" Then Exit Sub

        Select Case PicAudioPlay.Tag
            Case "Pro"
                PicAudioPlay.Tag = "Sys"
                If Trim(TxtSetPro.Text) <> "" Then
                    Mod_Audio.PlayAudio(Trim(TxtSetPro.Text))

                Else
                    PicAudioNext_Click(sender, e)
                End If

            Case "Rem"
                PicAudioPlay.Tag = "Pro"
                If Trim(TxtSetRem.Text) <> "" Then
                    Mod_Audio.PlayAudio(Trim(TxtSetRem.Text))

                Else
                    PicAudioNext_Click(sender, e)
                End If

            Case "Sys"
                PicAudioPlay.Tag = "Rem"
                If Trim(TxtSetSys.Text) <> "" Then
                    Mod_Audio.PlayAudio(Trim(TxtSetSys.Text))

                Else
                    PicAudioNext_Click(sender, e)
                End If

        End Select
    End Sub

    Private Sub TxtH_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TxtH.KeyPress
        'chi cho nhap so
        Try
            If e.KeyChar = Chr(8) Then  'cho phép fím BackSpace hoạt động
                Exit Sub
            End If

            Dim s As String = e.KeyChar
            Dim f As Short = s / 1  'tạo lỗi
        Catch ex As Exception
            e.KeyChar = ""  'huỷ ký tự
        End Try
    End Sub

    Private Sub TxtM_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TxtM.KeyPress
        Try
            If e.KeyChar = Chr(8) Then
                Exit Sub
            End If

            Dim s As String = e.KeyChar
            Dim f As Short = s / 1
        Catch ex As Exception
            e.KeyChar = ""
        End Try
    End Sub

    Private Sub TxtS_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TxtS.KeyPress
        Try
            If e.KeyChar = Chr(8) Then
                Exit Sub
            End If

            Dim s As String = e.KeyChar
            Dim f As Short = s / 1
        Catch ex As Exception
            e.KeyChar = ""
        End Try
    End Sub

    Private Sub TxtDay_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TxtDay.KeyPress
        Try
            If e.KeyChar = Chr(8) Or e.KeyChar = Chr(47) Then
                Exit Sub
            End If

            Dim s As String = e.KeyChar
            Dim f As Short = s / 1
        Catch ex As Exception
            e.KeyChar = ""
        End Try
    End Sub

    Private Sub PicAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PicAdd.Click
        Mod_ActionList.AddAction(DvActionList)

    End Sub

    Private Sub PicAdd_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PicAdd.MouseDown
        PicAdd.Image = My.Resources.Resource1.MD_ADD
    End Sub

    Private Sub PicAdd_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles PicAdd.MouseEnter
        PicAdd.Image = My.Resources.Resource1.MM_ADD
    End Sub

    Private Sub PicAdd_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles PicAdd.MouseLeave
        PicAdd.Image = My.Resources.Resource1.N_ADD
    End Sub

    Private Sub PicAdd_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PicAdd.MouseUp
        PicAdd.Image = My.Resources.Resource1.MM_ADD
    End Sub

    Private Sub OptSysShutdown_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OptSysShutdown.CheckedChanged
        PanSystemAction.Tag = OptSysShutdown.Tag
    End Sub

    Private Sub OptSysRestart_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OptSysRestart.CheckedChanged
        PanSystemAction.Tag = OptSysRestart.Tag
    End Sub

    Private Sub OptSysLogOff_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OptSysLogOff.CheckedChanged
        PanSystemAction.Tag = OptSysLogOff.Tag
    End Sub

    Private Sub OptSysStandBy_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OptSysStandBy.CheckedChanged
        PanSystemAction.Tag = OptSysStandBy.Tag
    End Sub

    Private Sub OptSysHibernate_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OptSysHibernate.CheckedChanged
        PanSystemAction.Tag = OptSysHibernate.Tag
    End Sub

    

    Private Sub OptSysTurnOffMonitor_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OptSysTurnOffMonitor.CheckedChanged
        PanSystemAction.Tag = OptSysTurnOffMonitor.Tag
    End Sub

    Private Sub OptSysTurnOnMonitor_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OptSysTurnOnMonitor.CheckedChanged
        PanSystemAction.Tag = OptSysTurnOnMonitor.Tag
    End Sub

    
    Private Sub T1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles T1.Tick

        If ProNextWork.RightToLeft = Windows.Forms.RightToLeft.No Then
            If ProNextWork.Value < ProNextWork.Maximum Then
                ProNextWork.Value += 1
            Else
                ProNextWork.Value = 1
                ProNextWork.RightToLeft = Windows.Forms.RightToLeft.Yes
            End If
        ElseIf ProNextWork.RightToLeft = Windows.Forms.RightToLeft.Yes Then
            If ProNextWork.Value < ProNextWork.Maximum Then
                ProNextWork.Value += 1
            Else
                ProNextWork.Value = 1
                ProNextWork.RightToLeft = Windows.Forms.RightToLeft.No
            End If
        End If

        ThucHien(T1, DvActionList)
    End Sub

    Private Sub BtnNewList_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnNewList.Click
        If MessageBox.Show("Tạo mới danh sách sẽ làm danh sách công việc cũ bị xoá đi," & _
                           vbCrLf & "Nhấn YES để đồng ý, ngược lại nhấn NO", _
                           "Xác nhận tạo mới danh sách", MessageBoxButtons.YesNo, _
                           MessageBoxIcon.Question) = Windows.Forms.DialogResult.Yes Then

            DvActionList.Rows.Clear()
            Dim str As String = _
                vbCrLf & _
                "[ACTION STATUS]" & vbCrLf & _
                "Total=0" & vbCrLf & _
                "Next=1" & vbCrLf & _
                "StartAt=1" & vbCrLf & _
                "Pause=False" & vbCrLf & _
                "Continue=False" & vbCrLf & vbCrLf & vbCrLf & _
                "[Success]" & vbCrLf & _
                "Count=0" & vbCrLf & vbCrLf & vbCrLf & _
                "[Error]" & vbCrLf & _
                "Count=0" & vbCrLf & vbCrLf & vbCrLf & _
                "[Skip]" & vbCrLf & _
                "Count=0" & vbCrLf & _
                "SkipAt=0" & vbCrLf & vbCrLf & vbCrLf & _
                "_____________________________________________________________________________________________________________" & vbCrLf & _
                "[ACTION LIST]" & vbCrLf & _
                "Count=0" & vbCrLf

            Mod_Reminder.WriteText(str, AppPath() & "Action List.ini")

            'xoá hết file rác
            Dim fso
            fso = CreateObject("Scripting.FileSystemObject")
            fso.DeleteFile(AppPath() & "Reminder\*.*")
            fso = Nothing

        End If
    End Sub

    Private Sub ChkStart_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ChkStart.Click
        T1.Enabled = True
        Dim INI As New ClassIniFile(AppPath() & "Action list.ini")
        INI.WriteString("ACTION STATUS", "Continue", "True")
        Mod_ActionList.LoadActionList(AppPath() & "Action List.ini", DvActionList)

        ChkStart.Enabled = False
        ChkPause.Enabled = True
        ChkSkip.Enabled = True
        ChkStop.Enabled = True

        ChkStart.Text = " "
        ChkPause.Text = "Tạm dừng"
        ChkSkip.Text = "Bỏ qua"
        ChkStop.Text = "Dừng lại"
    End Sub

    Private Sub ChkStart_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles ChkStart.MouseDown
        ChkStart.Image = My.Resources.Resource1.MD_START
    End Sub

    Private Sub ChkStart_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles ChkStart.MouseEnter
        ChkStart.Image = My.Resources.Resource1.MM_START
    End Sub

    Private Sub ChkStart_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles ChkStart.MouseLeave
        ChkStart.Image = My.Resources.Resource1.N_START
    End Sub

    Private Sub ChkStart_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles ChkStart.MouseUp
        ChkStart.Image = My.Resources.Resource1.MM_START
    End Sub

    Private Sub ChkPause_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ChkPause.Click
        Dim INI As New ClassIniFile(AppPath() & "Action List.ini")

        If ChkPause.Text = "Tạm dừng" Then
            ChkPause.Text = "Tiếp tục"
            INI.WriteString("ACTION STATUS", "Pause", "True")
            T1.Enabled = False

        ElseIf ChkPause.Text = "Tiếp tục" Then
            ChkPause.Text = "Tạm dừng"
            INI.WriteString("ACTION STATUS", "Pause", "False")
            T1.Enabled = True

        End If
    End Sub

    Private Sub ChkPause_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles ChkPause.MouseDown
        ChkPause.Image = My.Resources.Resource1.MD_PAUSE
    End Sub

    Private Sub ChkPause_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles ChkPause.MouseEnter
        ChkPause.Image = My.Resources.Resource1.MM_PAUSE
    End Sub

    Private Sub ChkPause_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles ChkPause.MouseLeave
        ChkPause.Image = My.Resources.Resource1.N_PAUSE
    End Sub

    Private Sub ChkPause_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles ChkPause.MouseUp
        ChkPause.Image = My.Resources.Resource1.MM_PAUSE
    End Sub

    Private Sub ChkSkip_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ChkSkip.Click
        Dim INI As New ClassIniFile(AppPath() & "Action List.ini")
        Dim Count As Integer = INI.GetInteger("Skip", "Count", 0)
        Dim StartAt As Integer = INI.GetString("ACTION STATUS", "StartAt", "")
        Dim ListCount As Integer = INI.GetInteger("ACTION LIST", "Count", 0)

        If StartAt < INI.GetInteger("ACTION STATUS", "Total", 0) Then

            INI.WriteString("Skip", CStr(Count + 1), CStr(StartAt))
            'giá trị sẽ skip
            INI.WriteInteger("Skip", "SkipAt", StartAt)
            INI.WriteInteger("Skip", "Count", Count + 1)

            DvActionList.Item(0, StartAt - 1).Value = My.Resources.Resource1.AL_Skip

            For i As Int16 = 1 To ListCount

                If DvActionList.Item(2, StartAt).Value & "_" & _
                    DvActionList.Item(3, StartAt).Value = _
                    INI.GetString("ACTION LIST", CStr(i), "") Then

                    INI.WriteInteger("ACTION STATUS", "Next", i)
                    Exit For

                End If

            Next

            INI.WriteInteger("ACTION STATUS", "StartAt", StartAt + 1)

        End If
    End Sub

    Private Sub ChkSkip_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles ChkSkip.MouseDown
        ChkSkip.Image = My.Resources.Resource1.MD_SKIP
    End Sub

    Private Sub ChkSkip_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles ChkSkip.MouseEnter
        ChkSkip.Image = My.Resources.Resource1.MM_SKIP
    End Sub

    Private Sub ChkSkip_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles ChkSkip.MouseLeave
        ChkSkip.Image = My.Resources.Resource1.N_SKIP
    End Sub

    Private Sub ChkSkip_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles ChkSkip.MouseUp
        ChkSkip.Image = My.Resources.Resource1.MM_SKIP
    End Sub

    Private Sub ChkStop_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ChkStop.Click
        Dim INI As New ClassIniFile(AppPath() & "Action List.ini")
        Dim StartAt As Integer = INI.GetInteger("ACTION STATUS", "StartAt", 0)
        Dim Total As Integer = INI.GetInteger("ACTION STATUS", "Total", 0)

        T1.Enabled = False

        For i As Integer = StartAt To Total

            DvActionList.Item(0, i - 1).Value = My.Resources.Resource1.AL_Skip
            INI.WriteInteger("Skip", CStr(INI.GetInteger("Skip", "Count", 0) + 1), i)
            INI.WriteInteger("Skip", "Count", INI.GetInteger("Skip", "Count", 0) + 1)

        Next

        INI.WriteString("ACTION STATUS", "Continue", "False")
        INI.WriteInteger("ACTION STATUS", "Next", 1)
        INI.WriteInteger("ACTION STATUS", "StartAt", 1)
        INI.WriteInteger("Skip", "SkipAt", 0)


        ChkStart.Enabled = True
        ChkPause.Enabled = False
        ChkSkip.Enabled = False
        ChkStop.Enabled = False

        ChkStart.Text = "Thực hiện"
        ChkPause.Text = " "
        ChkSkip.Text = " "
        ChkStop.Text = " "

        Pic_Status.Image = My.Resources.Resource1.status_stop
        LblStatus.Text = "Trạng thái : Không hoạt động"
        ProNextWork.Value = 0

    End Sub

    Private Sub ChkStop_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles ChkStop.MouseDown
        ChkStop.Image = My.Resources.Resource1.MD_STOP
    End Sub

    Private Sub ChkStop_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles ChkStop.MouseEnter
        ChkStop.Image = My.Resources.Resource1.MM_STOP
    End Sub

    Private Sub ChkStop_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles ChkStop.MouseLeave
        ChkStop.Image = My.Resources.Resource1.N_STOP
    End Sub

    Private Sub ChkStop_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles ChkStop.MouseUp
        ChkStop.Image = My.Resources.Resource1.MM_STOP
    End Sub

    Private Sub ChkStop_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChkStop.CheckedChanged

    End Sub

    Private Sub ChkGiaHan_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChkGiaHan.CheckedChanged
        NumGiaHan.Enabled = ChkGiaHan.Checked
    End Sub

    Private Sub OptSysLockComputer_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OptSysLockComputer.CheckedChanged
        PanSystemAction.Tag = OptSysLockComputer.Tag
    End Sub

    Private Sub OptSysLockComputer_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles OptSysLockComputer.MouseDown
        OptSysLockComputer.Image = My.Resources.Resource1.MD_LOCK_COMPUTER
    End Sub

    Private Sub OptSysLockComputer_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles OptSysLockComputer.MouseEnter
        OptSysLockComputer.Image = My.Resources.Resource1.MM_lock_COMPUTER
    End Sub

    Private Sub OptSysLockComputer_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles OptSysLockComputer.MouseLeave
        OptSysLockComputer.Image = My.Resources.Resource1.N_lock_COMPUTER
    End Sub

    Private Sub OptSysLockComputer_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles OptSysLockComputer.MouseUp
        OptSysLockComputer.Image = My.Resources.Resource1.MM_lock_COMPUTER
    End Sub

    Private Sub ChkTip_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChkTip.CheckedChanged
        TxtReminder.Enabled = ChkTip.Checked
        BtnClearReminder.Enabled = ChkTip.Checked
        BtnOpenReminder.Enabled = ChkTip.Checked
        BtnSaveReminder.Enabled = ChkTip.Checked
    End Sub

    Private Sub ChkEMail_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChkEMail.CheckedChanged
        TxtEMail.Enabled = ChkEMail.Checked
        BtnNewEMail.Enabled = ChkEMail.Checked
    End Sub

    Private Sub BtnNewEMail_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnNewEMail.Click
        TxtEMail.Text = "[EMail]" & vbCrLf & _
                    "To=..." & vbCrLf & _
                    "From=...@gmail.com" & vbCrLf & _
                    "Account=..." & vbCrLf & _
                    "Password=..." & vbCrLf & _
                    "Smtp = smtp.gmail.com" & vbCrLf & vbCrLf & _
                    "[Subject]...[Subject]" & vbCrLf & _
                    "[Body]...[Body]"
    End Sub

    Private Sub Tray_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Tray.MouseClick
        'If e.Button = Windows.Forms.MouseButtons.Right Then
        'PopupMenu1.Show(MousePosition.X, MousePosition.Y)

        ' End If
    End Sub


    Private Sub Tray_MouseDoubleClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Tray.MouseDoubleClick
        If e.Button = Windows.Forms.MouseButtons.Left Then

            TTime.Enabled = True
            Me.Show()
            Me.WindowState = FormWindowState.Normal

            Tray.Visible = False

        End If
    End Sub

    Private Sub PicMinimize_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PicMinimize.Click

        TTime.Enabled = False
        Me.WindowState = FormWindowState.Minimized
        Me.Hide()

        Tray.Visible = True

    End Sub

    Private Sub PicMinimize_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PicMinimize.MouseDown
        PicMinimize.Image = My.Resources.Resource1.md_minimize
    End Sub

    Private Sub PicMinimize_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles PicMinimize.MouseEnter
        PicMinimize.Image = My.Resources.Resource1.mm_minimize
    End Sub

    Private Sub PicMinimize_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles PicMinimize.MouseLeave
        PicMinimize.Image = My.Resources.Resource1.n_minimize
    End Sub

    Private Sub PicMinimize_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PicMinimize.MouseUp
        PicMinimize.Image = My.Resources.Resource1.mm_minimize
    End Sub

    
    Private Sub mnuShow_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuShow.Click

        TTime.Enabled = True
        Me.Show()
        Me.WindowState = FormWindowState.Normal
        Tray.Visible = False

    End Sub

    Private Sub mnuWebVisit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuWebVisit.Click
        Process.Start("http://www.phapsoftware.hotgoo.net")
    End Sub

    Private Sub mnuExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuExit.Click
        Mod_Config.SaveAllSettingToReg()
        Application.Exit()
    End Sub

    Private Sub BtnSaveToActionList_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnSaveToActionList.Click
        Dim str As String = _
                vbCrLf & _
                "[ACTION STATUS]" & vbCrLf & _
                "Total=0" & vbCrLf & _
                "Next=1" & vbCrLf & _
                "StartAt=1" & vbCrLf & _
                "Pause=False" & vbCrLf & _
                "Continue=False" & vbCrLf & vbCrLf & vbCrLf & _
                "[Success]" & vbCrLf & _
                "Count=0" & vbCrLf & vbCrLf & vbCrLf & _
                "[Error]" & vbCrLf & _
                "Count=0" & vbCrLf & vbCrLf & vbCrLf & _
                "[Skip]" & vbCrLf & _
                "Count=0" & vbCrLf & _
                "SkipAt=0" & vbCrLf & vbCrLf & vbCrLf & _
                "_____________________________________________________________________________________________________________" & vbCrLf & _
                "[ACTION LIST]" & vbCrLf & _
                "Count=0" & vbCrLf

        Mod_Reminder.WriteText(str, AppPath() & "Action List.ini")
        Mod_ActionList.UpdateActionList(AppPath() & "Action List.ini", DvActionList)
    End Sub

   
    Private Sub LblLastWork_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles LblLastWork.MouseMove
        Tip1.SetToolTip(LblLastWork, LblLastWork.Text.Replace("Công việc vừa làm : ", ""))
    End Sub

  
    Private Sub ChkStartWithWindows_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChkStartWithWindows.CheckedChanged
        If ChkStartWithWindows.Checked = True Then
            StartWithWindows_Enable()
        Else
            StartWithWindows_Disable()
        End If
    End Sub


    Private Sub FrmMain_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown
        'ẩn xuống khay
        If ChkMinimize.Checked = True Then
            TTime.Enabled = False
            Me.WindowState = FormWindowState.Minimized
            Hide()
            Tray.Visible = True
            Tray.ShowBalloonTip(5000, "Action Timer 2010", _
                    "Phần mềm Action Timer 2010 đã được khởi động" & vbCrLf & _
                    "Chương trình sẽ chạy ẩn dưới khay hệ thống, right-click vào biểu " & _
                    vbCrLf & "tượng của chương trình để hiện menu", ToolTipIcon.Info)
        End If
    End Sub

    Private Sub FrmMain_SizeChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.SizeChanged
        'ẩn xuống khay mỗi khi minimize
        If Me.WindowState = FormWindowState.Minimized And _
            ChkAutoMinimize.Checked = True Then
            TTime.Enabled = False
            Me.Hide()
            Tray.Visible = True
        End If
    End Sub

    Private Sub mnuHelp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuHelp.Click
        Process.Start("mailto:xeko_necromancer@zing.vn")
    End Sub

    
    Private Sub LlblHomepage_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LlblHomepage.LinkClicked
        Process.Start("http://www.phapsoftware.hotgoo.net")
    End Sub

    Private Sub LlblEMail_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LlblEMail.LinkClicked
        Process.Start("mailto:xeko_necromancer@zing.vn")
    End Sub

    Private Sub ChkSkip_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChkSkip.CheckedChanged

    End Sub

    Private Sub TToolTip_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TToolTip.Tick
        Dim x As String
        Dim y As String

        'gán x = 1 ký tự đầu dòng văn bản
        x = Me.Text.Substring(0, 1)
        'gán y là phần còn lại
        y = Microsoft.VisualBasic.Right(Me.Text, Me.Text.Length - 1)

        'Hiển thị trở lại textbox theo thứ tự ngược lại.
        Me.Text = y + x
    End Sub

    Private Sub PopupMenu1_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles PopupMenu1.MouseLeave
        'PopupMenu1.Hide()
        'PopupMenu1.Close()
    End Sub

End Class
