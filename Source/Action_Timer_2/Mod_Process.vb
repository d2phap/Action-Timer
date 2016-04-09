
Imports System
Imports System.Diagnostics


Module Mod_Process

    Public Sub ListProcess(ByVal DataView As DataGridView)
        Dim running() As Process = Process.GetProcesses()

        If running.Length > 0 Then

            Dim Cnt As Integer
            Dim i As Integer = 0

            DataView.Rows.Clear()

            For Cnt = 0 To running.Length - 1

                Try
                    DataView.Rows.Add()

                    DataView.Item(1, i).Value = running(Cnt).Id
                    DataView.Item(2, i).Value = running(Cnt).ProcessName
                    DataView.Item(3, i).Value = running(Cnt).MainWindowTitle
                    DataView.Item(4, i).Value = running(Cnt).MainModule.FileName
                    DataView.Item(5, i).Value = running(Cnt).BasePriority
                    DataView.Item(6, i).Value = running(Cnt).PagedMemorySize.ToString
                    DataView.Item(7, i).Value = running(Cnt).HandleCount
                    DataView.Item(8, i).Value = running(Cnt).StartTime.ToString
                    DataView.Item(9, i).Value = running(Cnt).Responding
                    DataView.Item(10, i).Value = running(Cnt).Modules.Count
                Catch ex As Exception
                End Try

                i += 1
            Next Cnt

            FrmMain.LblProcessCount.Text = "Danh sách các tiến trình : " & running.Length

        End If
    End Sub


    Public Sub KillProcess(ByVal DataView As DataGridView)
        Dim i As Integer
        Dim pro, T() As Process

        For i = 0 To DataView.RowCount - 1
            If DataView.Item(0, i).Value = CheckState.Checked Then

                T = Process.GetProcessesByName(DataView.Item(2, i).Value.ToString)
                For Each pro In T
                    pro.Kill()
                Next

            End If
        Next

        ListProcess(DataView)
    End Sub


    Public Sub DataViewCheckBox(ByVal DataView As DataGridView, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
        If e.ColumnIndex = 0 And e.RowIndex <> -1 Then
            If DataView.Item(e.ColumnIndex, e.RowIndex).Value = CheckState.Checked Then
                DataView.Item(e.ColumnIndex, e.RowIndex).Value = CheckState.Unchecked
            ElseIf DataView.Item(e.ColumnIndex, e.RowIndex).Value = CheckState.Unchecked Then
                DataView.Item(e.ColumnIndex, e.RowIndex).Value = CheckState.Checked
            End If
        End If
    End Sub

    Public Function KillProcessByPath(ByVal Path As String, ByVal DView As DataGridView) As Short

        Dim pro As Process
        Dim i As Integer

        KillProcessByPath = 0
        ListProcess(FrmMain.DViewProcess)

        For i = 0 To DView.Rows.Count - 1

            If Path = DView.Item(4, i).Value Then
                pro = Process.GetProcessById(CInt(DView.Item(1, i).Value))
                pro.Kill()
                KillProcessByPath = 1
                Exit For
            End If

        Next

    End Function

    Public Function StartPro(ByVal File As String) As Integer

        Try
            StartPro = Shell(File, AppWinStyle.NormalFocus)
        Catch ex As Exception
            StartPro = 0
        End Try

    End Function


End Module
