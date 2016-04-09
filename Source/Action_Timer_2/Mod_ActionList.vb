
Imports System.IO

Module Mod_ActionList

    Public Function AppPath() As String
        Dim str As String

        str = My.Application.Info.DirectoryPath & "\"
        AppPath = str.Replace("\\", "\")

    End Function


    'Lấy dữ liệu từ file ActionList đổ vào DataGridView
    Public Sub LoadActionList(ByVal file As String, ByVal DView As DataGridView)
        On Error Resume Next

        Dim INI As New ClassIniFile(file)

        Dim _Total As Integer = INI.GetInteger("ACTION STATUS", "Total", 0)
        Dim _StartAt As Integer = INI.GetInteger("ACTON STATUS", "StartAT", 0)

        Dim _Success As Integer = INI.GetInteger("Success", "Count", 0)
        Dim _Skip As Integer = INI.GetInteger("Skip", "Count", 0)
        Dim _Error As Integer = INI.GetInteger("Error", "Count", 0)

        Dim _ATCount As Integer = INI.GetInteger("ACTION LIST", "Count", 0)

        If _Total > 0 Then

            Dim i As Integer = 0
            Dim j As Integer = 0
            Dim Value As String
            Dim SubValue As String
            Dim str() As String
            Dim dem As Integer = 0

            DView.Rows.Clear()
            DView.Rows.Add(_Total)

            For i = 0 To _ATCount - 1
                Value = INI.GetString("ACTION LIST", CStr(i + 1), "") 'lấy chuỗi Ngày Giờ

                If Value.Length <> 0 Then
                    Dim _AL_Total As Integer = INI.GetInteger(Value, "Total", 0)

                    For j = 0 To _AL_Total - 1
                        SubValue = INI.GetString(Value, CStr(j + 1), "")    'lấy chuỗi công việc
                        str = Value.Split("_")  'cắt chuỗi ngày tháng

                        DView.Item(2, dem).Value = str(0)   'load dữ liệu vào cột Ngày Thực Hiện
                        DView.Item(3, dem).Value = str(1)   'load dữ liệu vào cột Thời Gian Thực Hiện

                        str = SubValue.Split(">")

                        Application.DoEvents()

                        Select Case str(1)  'Chọn loại công việc

                            Case "Pro" 'Process
                                DView.Item(4, dem).Value = "Quản lý tiến trình" 'load dữ liệu vào cột Phân Loại

                                'Load dữ liệu vào cột Công việc
                                If str(2) = "Kill" Then
                                    DView.Item(5, dem).Value = "Tắt tiến trình " & str(3)
                                ElseIf str(2) = "Run" Then
                                    DView.Item(5, dem).Value = "Chạy " & str(3)
                                End If

                            Case "Rem" 'Reminder
                                DView.Item(4, dem).Value = "Nhắc việc"

                                If str(2) = "Tip" Then
                                    DView.Item(5, dem).Value = "Hiển thị thông báo nhắc việc " & str(3)
                                ElseIf str(2) = "Mail" Then
                                    DView.Item(5, dem).Value = "Gửi email " & str(3)
                                End If

                            Case "Sys" 'System Action
                                DView.Item(4, dem).Value = "Thao tác với hệ thống"

                                Select Case str(2)
                                    Case "Shutdown"
                                        DView.Item(5, dem).Value = "Tắt máy tính"
                                    Case "Restart"
                                        DView.Item(5, dem).Value = "Khởi động lại máy tính"
                                    Case "LogOff"
                                        DView.Item(5, dem).Value = "Đăng xuất khỏi tài khoản"
                                    Case "StandBy"
                                        DView.Item(5, dem).Value = "Chuyển sang chế độ chờ"
                                    Case "Hibernate"
                                        DView.Item(5, dem).Value = "Chuyển sang chế độ ngủ đông"
                                        'Case "UnHibernate"
                                        'DView.Item(5, dem).Value = "Thoát khỏi chế độ ngủ đông"
                                    Case "TurnOffMonitor"
                                        DView.Item(5, dem).Value = "Tắt màn hình máy tính"
                                    Case "TurnOnMonitor"
                                        DView.Item(5, dem).Value = "Bật màn hình máy tính"
                                    Case "LockComputer"
                                        DView.Item(5, dem).Value = "Khoá máy tính"
                                End Select

                        End Select

                        'Đổ dữ liệu vào cột STT
                        DView.Item(1, dem).Value = dem + 1
                        'Lưu giá trị hình ảnh
                        DView.Item(1, dem).Tag = "AL_Normal"

                        'Tăng thêm 1 dòng
                        dem += 1
                    Next

                End If
                'Xoá dữ liệu trong mảng str
                ReDim str(4)
            Next


            'load dữ liệu vào cột hình ảnh (status)
            Dim Num As Integer   'lưu thông tin vị trí

            If _Success <> 0 Then
                For i = 1 To _Success
                    Num = INI.GetInteger("Success", CStr(i), 0)
                    DView.Item(0, Num - 1).Value = My.Resources.Resource1.AL_Success
                    DView.Item(0, Num - 1).Tag = "AL_Success"
                    DView.Rows(Num - 1).ReadOnly = True 'ko cho edit dòng này
                Next

            End If

            If _Skip <> 0 Then
                For i = 1 To _Skip
                    Num = INI.GetInteger("Skip", CStr(i), 0)
                    DView.Item(0, Num - 1).Value = My.Resources.Resource1.AL_Skip
                    DView.Item(0, Num - 1).Tag = "AL_Skip"
                    DView.Rows(Num - 1).ReadOnly = True 'ko cho edit dòng này
                Next
            End If

            If _Error <> 0 Then
                For i = 1 To _Error
                    Num = INI.GetInteger("Error", CStr(i), 0)
                    DView.Item(0, Num - 1).Value = My.Resources.Resource1.AL_Error
                    DView.Item(0, Num - 1).Tag = "AL_Error"
                    DView.Rows(Num - 1).ReadOnly = True 'ko cho edit dòng này
                Next
            End If

            Num = INI.GetInteger("ACTION STATUS", "StartAt", 0)
            If Num <> 0 And Num <= _Total - 1 And _
                INI.GetString("ACTION STATUS", "Continue", "") = "True" Then

                DView.Item(0, Num - 1).Value = My.Resources.Resource1.AL_NextTo
                DView.Item(0, Num - 1).Tag = "AL_NextTo"
                DView.Rows(Num - 1).ReadOnly = True 'ko cho edit dòng này
            End If

            'xoá mảng
            Erase str
        End If

    End Sub

    'Lưu dữ liệu từ DataGridView vào file ActionList

    Public Sub UpdateActionList(ByVal File As String, ByVal DView As DataGridView)

        Dim INI As New ClassIniFile(File)
        Dim i As Integer
        Dim _Value As String
        Dim str As String
        Dim _Count As Integer = 0


        INI.WriteInteger("Success", "Count", 0)
        INI.WriteInteger("Skip", "Count", 0)
        INI.GetInteger("Error", "Count", 0)

        For i = 0 To DView.Rows.Count - 1

            _Value = CStr(i + 1) & ">"
            str = Trim(DView.Item(5, i).Value)

            Select Case DView.Item(4, i).Value
                Case "Quản lý tiến trình"
                    _Value &= "Pro>"

                    If str.Substring(0, Len("Chạy")) = "Chạy" Then
                        _Value &= "Run>" & Trim(str.Substring(Len("Chạy")))
                    ElseIf str.Substring(0, Len("Tắt tiến trình")) = "Tắt tiến trình" Then
                        _Value &= "Kill>" & Trim(str.Substring(Len("Tắt tiến trình")))
                    End If

                Case "Nhắc việc"
                    _Value &= "Rem>"

                    If str.Substring(0, Len("Gửi email")) = "Gửi email" Then
                        _Value &= "Mail>" & Trim(str.Substring(Len("Gửi email")))
                    ElseIf str.Substring(0, Len("Hiển thị thông báo nhắc việc")) = _
                        "Hiển thị thông báo nhắc việc" Then
                        _Value &= "Tip>" & Trim(str.Substring(Len("Hiển thị thông báo nhắc việc")))
                    End If

                Case "Thao tác với hệ thống"
                    _Value &= "Sys>"
                    Select Case str
                        Case "Tắt máy tính"
                            _Value &= "Shutdown"
                        Case "Khởi động lại máy tính"
                            _Value &= "Restart"
                        Case "Đăng xuất khỏi tài khoản"
                            _Value &= "LogOff"
                        Case "Chuyển sang chế độ chờ"
                            _Value &= "StandBy"
                        Case "Chuyển sang chế độ ngủ đông"
                            _Value &= "Hibernate"
                            'Case "Thoát khỏi chế độ ngủ đông"
                            '_Value &= "UnHibernate"
                        Case "Tắt màn hình máy tính"
                            _Value &= "TurnOffMonitor"
                        Case "Bật màn hình máy tính"
                            _Value &= "TurnOnMonitor"
                        Case "Khoá máy tính"
                            _Value &= "LockComputer"
                    End Select

            End Select

            INI.WriteINI_MultiKey("ACTION LIST", DView.Item(2, i).Value & "_" & _
                                  DView.Item(3, i).Value, _Value)

            If DView.Item(0, i).Tag = "AL_Success" Then
                _Count = INI.GetInteger("Success", "Count", 0)
                INI.WriteInteger("Success", CStr(_Count + 1), i + 1)
                INI.WriteInteger("Success", "Count", _Count + 1)

            ElseIf DView.Item(0, i).Tag = "AL_Skip" Then
                _Count = INI.GetInteger("Skip", "Count", 0)
                INI.WriteInteger("Skip", CStr(_Count + 1), i + 1)
                INI.WriteInteger("Skip", "Count", _Count + 1)

            ElseIf DView.Item(0, i).Tag = "AL_Error" Then
                _Count = INI.GetInteger("Error", "Count", 0)
                INI.WriteInteger("Error", CStr(_Count + 1), i + 1)
                INI.WriteInteger("Error", "Count", _Count + 1)

            ElseIf DView.Item(0, i).Tag = "AL_NextTo" Then
                INI.WriteInteger("ACTION STATUS", "Next", i + 1)

            End If

        Next

        INI.WriteInteger("ACTION STATUS", "Total", DView.Rows.Count) 'ghi Total

    End Sub


    Public Sub AddAction(ByVal DView As DataGridView)

        With FrmMain

            Dim INI As New ClassIniFile(AppPath() & "Action List.ini")
            Dim sDay As String = .TxtDay.Text
            Dim str() As String = .TxtDay.Text.Split("/")
            .Err1.Clear()

            'kiểm tra xem thời gian có hợp lệ không ?
            If Len(.TxtH.Text) = 0 Or Len(.TxtM.Text) = 0 Or _
                Len(.TxtS.Text) = 0 Or Len(.TxtDay.Text) < 8 Then
                .Err1.SetError(.PicAdd, "Nhập chưa đủ dữ liệu ngày giờ...")
                Exit Sub
            Else

                .Err1.Clear()
                If CInt(.TxtH.Text) > 23 Or 0 > CInt(.TxtH.Text) Then
                    .Err1.SetError(.TxtH, "Giá trị không hợp lệ!")
                    Exit Sub

                ElseIf CInt(.TxtM.Text) > 59 Or 0 > CInt(.TxtM.Text) Then
                    .Err1.SetError(.TxtM, "Giá trị không hợp lệ!")
                    Exit Sub

                ElseIf CInt(.TxtS.Text) > 59 Or 0 > CInt(.TxtS.Text) Then
                    .Err1.SetError(.TxtS, "Giá trị không hợp lệ!")
                    Exit Sub

                ElseIf sDay.Contains("/") = False Then
                    .Err1.SetError(.TxtDay, "Ngày tháng phải được ngăn cách bởi '/'")
                    Exit Sub

                ElseIf str.Length <> 3 Then
                    .Err1.SetError(.TxtDay, "Phân cách giữa ngày tháng năm không hợp lệ")
                    Exit Sub

                ElseIf Len(str(0)) = 0 Or Len(str(1)) = 0 Or Len(str(2)) <> 4 Or _
                    CInt(str(0)) > 31 Or CInt(str(0)) < 1 Or CInt(str(1)) > 12 Or _
                    CInt(str(1)) < 1 Then
                    .Err1.SetError(.TxtDay, "Giá trị không hợp lệ!")
                    Exit Sub

                End If
            End If

            '-----------------THÊM DỮ LIỆU VÀO DVIEW VÀ FILE INI-----------------------

            Dim i As Integer
            Dim sDem As String
            Dim _Value As String

            'thêm vào cột Ngày thực hiện
            Dim Ngay As String = .TxtDay.Text
            'thêm vào cột Thời gian thực hiện
            Dim Gio As String = .TxtH.Text & ":" & .TxtM.Text & ":" & .TxtS.Text

            Select Case .PicAdd.Tag
                Case "Quản lý tiến trình"

                    sDem = "0"

                    For i = 0 To .DViewProcess.Rows.Count - 1
                        If .DViewProcess.Item(0, i).Value = CheckState.Checked Then

                            DView.Rows.Add()
                            DView.Item(0, DView.Rows.Count - 1).Tag = "AL_Normal"
                            DView.Item(1, DView.Rows.Count - 1).Value = DView.Rows.Count
                            DView.Item(2, DView.Rows.Count - 1).Value = Ngay
                            DView.Item(3, DView.Rows.Count - 1).Value = Gio
                            DView.Item(4, DView.Rows.Count - 1).Value = .PicAdd.Tag
                            DView.Item(5, DView.Rows.Count - 1).Value = _
                                "Tắt tiến trình " & .DViewProcess.Item(4, i).Value

                            sDem = "1"

                            'ghi vào INI
                            _Value = CStr(DView.Rows.Count) & ">Pro>Kill>" & _
                                .DViewProcess.Item(4, i).Value
                            INI.WriteINI_MultiKey("ACTION LIST", Ngay & "_" & Gio, _Value)

                            .DViewProcess.Item(0, i).Value = CheckState.Unchecked
                        End If
                    Next

                    If Trim(.TxtRun.Text) <> "" Then

                        DView.Rows.Add()
                        DView.Item(0, DView.Rows.Count - 1).Tag = "AL_Normal"
                        DView.Item(1, DView.Rows.Count - 1).Value = DView.Rows.Count
                        DView.Item(2, DView.Rows.Count - 1).Value = Ngay
                        DView.Item(3, DView.Rows.Count - 1).Value = Gio
                        DView.Item(4, DView.Rows.Count - 1).Value = .PicAdd.Tag
                        DView.Item(5, DView.Rows.Count - 1).Value = _
                            "Chạy " & .TxtRun.Text.Trim

                        sDem = "1"

                        'ghi vào INI
                        _Value = CStr(DView.Rows.Count) & ">Pro>Run>" & .TxtRun.Text.Trim
                        INI.WriteINI_MultiKey("ACTION LIST", Ngay & "_" & Gio, _Value)

                        .TxtRun.Text = ""
                    End If

                    If sDem = "0" Then
                        .Err1.SetError(.PicAdd, "Bạn chưa chọn tiến trình nào hoặc chưa nhập tên, đường dẫn")
                    End If

                Case "Nhắc việc"
                    Dim tam As String = "0"
                    If .ChkTip.Checked = True And Trim(.TxtReminder.Text) <> "" Then

                        sDem = Ngay.Replace("/", "~") & "_" & Gio.Replace(":", "~")
                        Mod_Reminder.WriteText(.TxtReminder.Text, _
                                               Mod_ActionList.AppPath & _
                                               "Reminder\" & sDem & ".txt")
                        DView.Rows.Add()
                        DView.Item(0, DView.Rows.Count - 1).Tag = "AL_Normal"
                        DView.Item(1, DView.Rows.Count - 1).Value = DView.Rows.Count
                        DView.Item(2, DView.Rows.Count - 1).Value = Ngay
                        DView.Item(3, DView.Rows.Count - 1).Value = Gio
                        DView.Item(4, DView.Rows.Count - 1).Value = .PicAdd.Tag
                        DView.Item(5, DView.Rows.Count - 1).Value = _
                            "Hiển thị thông báo nhắc việc " & _
                            Mod_ActionList.AppPath & "Reminder\" & sDem & ".txt"


                        'ghi vào INI
                        _Value = CStr(DView.Rows.Count) & ">Rem>Tip>" & _
                            Mod_ActionList.AppPath & "Reminder\" & sDem & ".txt"
                        INI.WriteINI_MultiKey("ACTION LIST", Ngay & "_" & Gio, _Value)

                        tam = "1"

                    End If

                    If .ChkEMail.Checked = True And .TxtEMail.Text.Trim <> "" Then
                        sDem = Ngay.Replace("/", "~") & "_" & Gio.Replace(":", "~")
                        Mod_Reminder.WriteText(.TxtEMail.Text, _
                                               Mod_ActionList.AppPath & _
                                               "Reminder\" & sDem & ".mail")

                        DView.Rows.Add()
                        DView.Item(0, DView.Rows.Count - 1).Tag = "AL_Normal"
                        DView.Item(1, DView.Rows.Count - 1).Value = DView.Rows.Count
                        DView.Item(2, DView.Rows.Count - 1).Value = Ngay
                        DView.Item(3, DView.Rows.Count - 1).Value = Gio
                        DView.Item(4, DView.Rows.Count - 1).Value = .PicAdd.Tag
                        DView.Item(5, DView.Rows.Count - 1).Value = _
                            "Gửi email " & _
                            Mod_ActionList.AppPath & "Reminder\" & sDem & ".mail"


                        'ghi vào INI
                        _Value = CStr(DView.Rows.Count) & ">Rem>Mail>" & _
                            Mod_ActionList.AppPath & "Reminder\" & sDem & ".mail"
                        INI.WriteINI_MultiKey("ACTION LIST", Ngay & "_" & Gio, _Value)

                        tam = "1"
                    End If


                    sDem = ""
                    If tam = "0" Then
                        .Err1.SetError(.PicAdd, "Bạn chưa chọn công việc cụ thể")
                    End If
                    tam = Nothing

                Case "Thao tác với hệ thống"

                    If .PanSystemAction.Tag <> "" Then
                        str = .PanSystemAction.Tag.ToString.Split("_")

                        DView.Rows.Add()
                        DView.Item(0, DView.Rows.Count - 1).Tag = "AL_Normal"
                        DView.Item(1, DView.Rows.Count - 1).Value = DView.Rows.Count
                        DView.Item(2, DView.Rows.Count - 1).Value = Ngay
                        DView.Item(3, DView.Rows.Count - 1).Value = Gio
                        DView.Item(4, DView.Rows.Count - 1).Value = .PicAdd.Tag
                        DView.Item(5, DView.Rows.Count - 1).Value = str(0)

                        'ghi vào INI
                        _Value = CStr(DView.Rows.Count) & ">Sys>" & str(1)
                        INI.WriteINI_MultiKey("ACTION LIST", Ngay & "_" & Gio, _Value)

                        .PanSystemAction.Tag = ""

                        .OptSysHibernate.Checked = False
                        .OptSysLockComputer.Checked = False
                        .OptSysLogOff.Checked = False
                        .OptSysRestart.Checked = False
                        .OptSysShutdown.Checked = False
                        .OptSysStandBy.Checked = False
                        .OptSysTurnOffMonitor.Checked = False
                        .OptSysTurnOnMonitor.Checked = False
                        '.OptSysUnHibernate.Checked = False

                    Else
                        .Err1.SetError(.PicAdd, "Bạn chưa chọn công việc về hệ thống")
                    End If

            End Select
            'Ghi khoá Total
            INI.WriteInteger("ACTION STATUS", "Total", DView.Rows.Count)

            'xoá mảng str()
            Erase str
        End With
    End Sub



    Public Sub ThucHien(ByVal T1 As Timer, ByVal DView As DataGridView)
        'Thực hiện dựa vào dữ liệu ở file INI là chính
        'Sau đó mới update dữ liệu ở DView
        'update chủ yếu là cột Status 
        On Error Resume Next

        Dim INI As New ClassIniFile(AppPath() & "Action List.ini")

        If IO.File.Exists(AppPath() & "Action List.ini") = False Then
            FrmReminder.TxtTip.Text = "Lỗi không tìm thấy file Action List.ini" & _
                             vbNewLine & vbNewLine & _
                             "Chương trình không thể thực hiện tiếp công việc của bạn," & _
                             vbNewLine & "vì thế chương trình sẽ dừng lại tại đây"
            FrmReminder.Show()
            GoTo finish
        End If

        'pause
        If INI.GetString("ACTION STATUS", "Pause", "") = "True" Then
            FrmMain.Pic_Status.Image = My.Resources.Resource1.status_pause
            FrmMain.LblStatus.Text = "Trạng thái : Tạm dừng"
            Exit Sub
        End If

        FrmMain.Pic_Status.Image = My.Resources.Resource1.status_run
        FrmMain.LblStatus.Text = "Trạng thái : Đang thực hiện"

        Dim H%, M%, S%  'giờ phút giây
        Dim cH%, cM%, cS%   'giờ phút giây hiện tại

        Dim Ngay%, Thang%, Nam% 'ngày tháng năm
        Dim cNgay%, cThang%, cNam%  'ngày tháng năm hiện tại

        Dim NextTo% 'index của công việc sắp tới
        Dim Total%  'tổng số công việc

        Dim dem As String = ""
        Dim str() As String
        Dim st() As String

        NextTo = INI.GetInteger("ACTION STATUS", "Next", 0) 'vi tri thực hiện
        Total = INI.GetInteger("ACTION LIST", "Count", 0) 'tổng

        'công việc kế tiếp ở STATUS-------------------------------------------------

        dem = INI.GetString("ACTION STATUS", "StartAt", 1)

        FrmMain.TxtNextWork.Text = _
            "Tên công việc : " & DView.Item(5, CInt(dem) - 1).Value & vbCrLf & _
            "Thuộc nhóm : " & DView.Item(4, CInt(dem) - 1).Value & vbCrLf & _
            "Thời gian thực hiện : " & DView.Item(3, CInt(dem) - 1).Value & " " & _
                DView.Item(2, CInt(dem) - 1).Value

        FrmMain.TxtResult.Text = _
            "Tổng số công việc : " & DView.Rows.Count & vbCrLf & _
            "Đã thực hiện : " & INI.GetString("Success", "Count", "0") & vbCrLf & _
            "Chưa thực hiện : " & DView.Rows.Count - CInt(dem) + 1 & vbCrLf & _
            "Công việc bị lỗi : " & INI.GetString("Error", "Count", "0") & vbCrLf & _
            "Tiến độ công việc : " & Math.Round(((CInt(dem) - 1) / DView.Rows.Count) * 100, 2) & " %"

        'công việc vưa làm
        If dem > 1 Then
            FrmMain.LblLastWork.Text = "Công việc vừa làm : " & DView.Item(5, CInt(dem) - 2).Value
        Else
            FrmMain.LblLastWork.Text = "Công việc vừa làm : chưa có"
        End If
        '-----------------------------------------------------------------------

        'so sánh công việc kế tiếp và tổng số
        If NextTo > Total Or Total = 0 Then 'hoàn thành xong rồi
finish:
            FrmMain.ChkStart.Enabled = True
            FrmMain.ChkPause.Enabled = False
            FrmMain.ChkSkip.Enabled = False
            FrmMain.ChkStop.Enabled = False

            FrmMain.ChkStart.Text = "Thực hiện"
            FrmMain.ChkPause.Text = " "
            FrmMain.ChkSkip.Text = " "
            FrmMain.ChkStop.Text = " "

            INI.WriteString("ACTION STATUS", "Continue", "False")
            INI.WriteInteger("ACTION STATUS", "Next", 1)
            INI.WriteInteger("ACTION STATUS", "StartAt", 1)
            INI.WriteString("ACTION STATUS", "Pause", "False")

            FrmMain.Pic_Status.Image = My.Resources.Resource1.status_stop
            FrmMain.LblStatus.Text = "Trạng thái : Không hoạt động"
            FrmMain.ProNextWork.Value = FrmMain.ProNextWork.Maximum

            FrmMain.TxtNextWork.Text = _
                "Tên công việc : không có" & vbCrLf & _
                "Thuộc nhóm : ..." & vbCrLf & _
                "Thời gian thực hiện : ..."

            FrmMain.TxtResult.Text = _
                "Tổng số công việc : " & DView.Rows.Count & vbCrLf & _
                "Đã thực hiện : " & INI.GetString("Success", "Count", "0") & vbCrLf & _
                "Chưa thực hiện : 0" & vbCrLf & _
                "Công việc bị lỗi : " & INI.GetString("Error", "Count", "0") & vbCrLf & _
                "Tiến độ công việc : 100 %"

            T1.Enabled = False
            Exit Sub
        End If

        dem = INI.GetString("ACTION LIST", CInt(NextTo), "") 'lấy chuỗi ngày giờ là subSection
        str = dem.Split("_") 'cắt ra 2 chuỗi ngày giờ
        st = str(0).Split("/") 'lấy chuỗi ngày cắt ra

        Ngay = CInt(st(0))
        Thang = CInt(st(1))
        Nam = CInt(st(2))

        st = str(1).Split(":") 'lấy chuỗi giờ cắt ra

        H = CInt(st(0))
        M = CInt(st(1))
        S = CInt(st(2))

        cH = DateAndTime.Hour(Now())
        cM = DateAndTime.Minute(Now())
        cS = DateAndTime.Second(Now())

        cNgay = DateAndTime.Day(Now())
        cThang = DateAndTime.Month(Now())
        cNam = DateAndTime.Year(Now())


        '-----------------------XÉT ĐIỀU KIỆN VÀ THỰC HIỆN--------------------------
        If Ngay = cNgay And Thang = cThang And Nam = cNam And _
            H = cH And M = cM And S = cS Then

            Dim i As Integer
            Dim strVal As String
            Dim Report As Integer    'báo cáo thành công
            Dim SubCount As Integer 'tổng công việc thực hiện cùng 1 lúc
            SubCount = INI.GetInteger(dem, "Total", 1) 'lấy dữ liệu

            For i = 1 To SubCount

                strVal = INI.GetString(dem, CStr(i), "") 'lấy chuỗi công việc
                ReDim str(4)
                str = strVal.Split(">") 'cắt công việc  

                If INI.GetString("Skip", "SkipAt", "") <> str(0) Then


                    'phân loại công việc
                    Select Case str(1)
                        Case "Pro"

                            If FrmMain.ChkSetSound.Checked = True And _
                                       Trim(FrmMain.TxtSetPro.Text) <> "" Then
                                setVolume(FrmMain.Vol1.Value)
                                PlayAudio(Trim(FrmMain.TxtSetPro.Text))
                            End If

                            Select Case str(2)
                                Case "Run"

                                    If StartPro(str(3).Trim) <> 0 Then
                                        Report = INI.GetInteger("Success", "Count", 0) 'lấy giá trị của Key Count
                                        INI.WriteInteger("Success", CStr(Report + 1), str(0)) 'ghi Key thành công
                                        INI.WriteInteger("Success", "Count", Report + 1) 'ghi trở lại Key Count
                                        DView.Item(0, CInt(str(0)) - 1).Value = My.Resources.Resource1.AL_Success
                                    Else
                                        Report = INI.GetInteger("Error", "Count", 0) 'lấy giá trị của Key Count
                                        INI.WriteInteger("Error", CStr(Report + 1), str(0)) 'ghi Key thất bại
                                        INI.WriteInteger("Error", "Count", Report + 1) 'ghi trở lại Key Count
                                        DView.Item(0, CInt(str(0)) - 1).Value = My.Resources.Resource1.AL_Error
                                    End If

                                Case "Kill"

                                    If KillProcessByPath(str(3).Trim, FrmMain.DViewProcess) <> 0 Then
                                        Report = INI.GetInteger("Success", "Count", 0) 'lấy giá trị của Key Count
                                        INI.WriteInteger("Success", CStr(Report + 1), str(0)) 'ghi Key thành công
                                        INI.WriteInteger("Success", "Count", Report + 1) 'ghi trở lại Key Count
                                        DView.Item(0, CInt(str(0)) - 1).Value = My.Resources.Resource1.AL_Success
                                    Else
                                        Report = INI.GetInteger("Error", "Count", 0) 'lấy giá trị của Key Count
                                        INI.WriteInteger("Error", CStr(Report + 1), str(0)) 'ghi Key thất bại
                                        INI.WriteInteger("Error", "Count", Report + 1) 'ghi trở lại Key Count
                                        DView.Item(0, CInt(str(0)) - 1).Value = My.Resources.Resource1.AL_Error
                                    End If

                            End Select

                        Case "Rem"

                            If FrmMain.ChkSetSound.Checked = True And _
                                       Trim(FrmMain.TxtSetRem.Text) <> "" Then
                                setVolume(FrmMain.Vol1.Value)
                                PlayAudio(Trim(FrmMain.TxtSetRem.Text))
                            End If

                            Select Case str(2)
                                Case "Tip"

                                    If IO.File.Exists(str(3).Trim) = True Then
                                        FrmReminder.TxtTip.Text = _
                                            "[" & H & ":" & M & ":" & S & "  " & Ngay & "/" & Thang & "/" & _
                                            Nam & "]__________________________________________________________" & _
                                            vbCrLf & vbCrLf & ReadText(str(3)) & vbCrLf & _
                                            vbCrLf & vbCrLf & vbCrLf & FrmReminder.TxtTip.Text


                                        FrmReminder.Show()
                                        Report = INI.GetInteger("Success", "Count", 0) 'lấy giá trị của Key Count
                                        INI.WriteInteger("Success", CStr(Report + 1), str(0)) 'ghi Key thành công
                                        INI.WriteInteger("Success", "Count", Report + 1) 'ghi trở lại Key Count
                                        DView.Item(0, CInt(str(0)) - 1).Value = My.Resources.Resource1.AL_Success
                                    Else
                                        Report = INI.GetInteger("Error", "Count", 0) 'lấy giá trị của Key Count
                                        INI.WriteInteger("Error", CStr(Report + 1), str(0)) 'ghi Key thất bại
                                        INI.WriteInteger("Error", "Count", Report + 1) 'ghi trở lại Key Count
                                        DView.Item(0, CInt(str(0)) - 1).Value = My.Resources.Resource1.AL_Error
                                    End If

                                Case "Mail"
                                    Dim kq As Integer = Mod_Reminder.Send_Email(str(3).Trim)

                                    If kq = 1 Then

                                        Report = INI.GetInteger("Success", "Count", 0) 'lấy giá trị của Key Count
                                        INI.WriteInteger("Success", CStr(Report + 1), str(0)) 'ghi Key thành công
                                        INI.WriteInteger("Success", "Count", Report + 1) 'ghi trở lại Key Count
                                        DView.Item(0, CInt(str(0)) - 1).Value = My.Resources.Resource1.AL_Success

                                    ElseIf kq = 0 Then

                                        Report = INI.GetInteger("Error", "Count", 0) 'lấy giá trị của Key Count
                                        INI.WriteInteger("Error", CStr(Report + 1), str(0)) 'ghi Key thất bại
                                        INI.WriteInteger("Error", "Count", Report + 1) 'ghi trở lại Key Count
                                        DView.Item(0, CInt(str(0)) - 1).Value = My.Resources.Resource1.AL_Error

                                    End If

                            End Select

                        Case "Sys"

                            Report = INI.GetInteger("Success", "Count", 0) 'lấy giá trị của Key Count
                            INI.WriteInteger("Success", CStr(Report + 1), str(0)) 'ghi Key thành công
                            INI.WriteInteger("Success", "Count", Report + 1) 'ghi trở lại Key Count
                            FrmMain.DvActionList.Item(0, CInt(str(0)) - 1).Value = My.Resources.Resource1.AL_Success

                            If FrmMain.ChkSetSound.Checked = True And _
                                Trim(FrmMain.TxtSetSys.Text) <> "" Then

                                setVolume(FrmMain.Vol1.Value)
                                PlayAudio(Trim(FrmMain.TxtSetSys.Text))
                            End If

                            Select Case str(2)
                                'Case "UnHibernate"
                                '## code UnHibernate ##
                                Case "TurnOnMonitor"
                                    Mod_System.power_TurnOnMonitor()
                                Case Else
                                    FrmSystem.ProSys.Tag = str(2)
                                    FrmSystem.Show()

                            End Select

                    End Select

                End If


                If CInt(str(0)) < DView.Rows.Count Then
                    'đặt biểu tượng Next |=>
                    DView.Item(0, CInt(str(0))).Value = My.Resources.Resource1.AL_NextTo
                    INI.WriteInteger("ACTION STATUS", "StartAt", CInt(str(0)) + 1)
                End If

                My.Application.DoEvents()

            Next

            'Thực hiện công việc kế tiếp
            INI.WriteInteger("ACTION STATUS", "Next", NextTo + 1)

        Else    'nếu điều kiện không đúng
            Dim i As Integer = INI.GetInteger("ACTION STATUS", "StartAt", 0)
            If i <= DView.Rows.Count Then
                'đặt biểu tượng Next |=>
                DView.Item(0, i - 1).Value = My.Resources.Resource1.AL_NextTo

            End If

        End If

        'xoá mảng giải phóng bộ nhớ
        Erase str
        Erase st

    End Sub


End Module
