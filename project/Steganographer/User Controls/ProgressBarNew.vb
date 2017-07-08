Public Class ProgressBarNew

    Protected Overrides Sub WndProc(ByRef M As Message)

        MyBase.WndProc(M)

        If M.Msg = &HF Then    'WM_PAINT

            Dim G As Graphics = Graphics.FromHwnd(Me.Handle)
            Dim Font As New Font("Microsoft Sans Serif", 8.25, FontStyle.Regular, GraphicsUnit.Point)
            Dim StrSz As SizeF = G.MeasureString(Me.Value.ToString + " %", Font)
            Dim X As Integer = (Me.Width / 2) - (StrSz.Width / 2)
            Dim Y As Integer = (Me.Height / 2) - (StrSz.Height / 2)
            G.DrawString(Me.Value.ToString + " %", Font, Brushes.Black, X, Y)
            G.Flush()
            Font.Dispose()

        End If

    End Sub

End Class
