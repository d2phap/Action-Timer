



Imports System.Runtime.InteropServices


Public Class AeroGlass

    'xác định aero có bật không?
    <Runtime.InteropServices.DllImport("dwmapi.dll", PreserveSig:=False)> _
    Public Shared Function DwmIsCompositionEnabled() As Boolean
    End Function

    'khai báo tạo aero
    <StructLayout(LayoutKind.Sequential)> _
    Public Structure MARGINS
        Public cxLeftWidth As Integer
        Public cxRightWidth As Integer
        Public cyTopHeight As Integer
        Public cyBottomheight As Integer
    End Structure

    <DllImport("dwmapi.dll")> _
    Public Shared Function DwmExtendFrameIntoClientArea(ByVal hWnd As IntPtr, ByRef pMarinset As MARGINS) As Integer
    End Function


    Public Sub Apply_AeroGlass(ByVal Frm As Form, Optional ByVal cyLeft As Integer = -1, _
                        Optional ByVal cyRight As Integer = -1, _
                        Optional ByVal cyTop As Integer = -1, _
                       Optional ByVal cyBottom As Integer = -1)

        Try
            'mặc định là áp dụng hết form

            If DwmIsCompositionEnabled() = True Then
                Frm.BackColor = Color.Black

                Dim margins As MARGINS = New MARGINS

                margins.cxLeftWidth = cyLeft
                margins.cxRightWidth = cyRight
                margins.cyTopHeight = cyTop
                margins.cyBottomheight = cyBottom

                Dim hwnd As IntPtr = Frm.Handle
                Dim result As Integer = DwmExtendFrameIntoClientArea(hwnd, margins)

            End If

        Catch ex As Exception
        End Try

    End Sub

End Class
