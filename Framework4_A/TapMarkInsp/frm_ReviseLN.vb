Public Class frm_ReviseLN

    Private Sub TextBox1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox1.KeyDown

        With Me
            If e.KeyCode = Keys.Enter Then
                TapMarkInsp.SettingLot.ReviseLotNo = .TextBox1.Text.ToUpper
                .TextBox1.Text = ""
                .Close()
            End If

        End With

    End Sub

    Private Sub frm_ReviseLN_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown

        TapMarkInsp.SettingLot.ReviseLotNo = ""
        Me.TextBox1.Focus()

    End Sub

End Class