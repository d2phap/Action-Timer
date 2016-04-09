Public Class FrmReminder

    Private Sub BtnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnClose.Click
        Me.Hide()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        TxtTip.Text = vbNullString
    End Sub

    Private Sub FrmReminder_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Dim AERO As New AeroGlass
        AERO.Apply_AeroGlass(Me)

        Me.Top = Microsoft.Win32.Registry.GetValue("HKEY_CURRENT_USER\Software\PhapSoftware\" & Application.ProductName, "FrmReminder_Top", 0)
        Me.Left = Microsoft.Win32.Registry.GetValue("HKEY_CURRENT_USER\Software\PhapSoftware\" & Application.ProductName, "FrmReminder_Left", 0)

    End Sub
End Class