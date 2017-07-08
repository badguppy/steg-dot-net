Public Class Frm_Main

    Private Sub Cmd_Welcome_Next_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cmd_Welcome_Next.Click

        If Opt_Enc.Checked Then
            Me.Panel_Enc_Input.BringToFront()
        Else
            Me.Panel_Dec_Input.BringToFront()
        End If

    End Sub

    Private Sub Cmd_Enc_Input_Back_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cmd_Enc_Input_Back.Click
        Me.Panel_Welcome.BringToFront()
    End Sub

    Private Sub Cmd_Enc_Input_Browse_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cmd_Enc_Input_Browse.Click
        If Me.DlgOpen_Enc_Input.ShowDialog = Windows.Forms.DialogResult.Cancel Then Exit Sub
        Me.Lbl_Enc_Input.Text = Me.DlgOpen_Enc_Input.FileName
        Enc_FileSize = FileIO.FileSystem.GetFileInfo(Me.DlgOpen_Enc_Input.FileName).Length
        Me.Lbl_Enc_Input_Size.Text = String.Format("{0:###,###,###,###,###} Bytes", Enc_FileSize)
        Me.Lbl_Enc_Input_Size2.Text = Me.Lbl_Enc_Input_Size.Text
        Me.Lbl_Enc_In_Proceed.Visible = True
        Me.Cmd_Enc_Input_Next.Enabled = True
        Me.Lbl_Enc_Input_Lbl.Visible = True
        Me.Lbl_Enc_Input.Visible = True
        Me.Lbl_Enc_Input_Size.Visible = True
        Me.Lbl_Enc_Input_SizeLbl.Visible = True
    End Sub

    Private Sub Cmd_Enc_Input_Next_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cmd_Enc_Input_Next.Click
        Enc_RefreshListView()
        Me.Panel_Enc_Options.BringToFront()
    End Sub

    Private Sub Cmd_Enc_Opt_Back_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cmd_Enc_Opt_Back.Click
        Me.Panel_Enc_Input.BringToFront()
    End Sub

    Private Sub Cmd_Enc_Opt_Next_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cmd_Enc_Opt_Next.Click

        If Me.Opt_Enc_Bin.Checked Then
            Enc_ComponentsPerByte = 8
        ElseIf Me.Opt_Enc_Dec.Checked Then
            Enc_ComponentsPerByte = 3
        Else
            Enc_ComponentsPerByte = 2
        End If

        If Me.Opt_Enc_24BPP.Checked Then
            Enc_ComponentsPerPixel = 3
        Else
            Enc_ComponentsPerPixel = 4
        End If

        Enc_RefreshListView()
        Me.Panel_Enc_Output.BringToFront()

    End Sub

    Private Sub Txt_Enc_Password_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Txt_Enc_Password.TextChanged
        If Txt_Enc_Password.Text = "" Then
            Me.Cmd_Enc_Opt_Next.Enabled = False
        Else
            Me.Cmd_Enc_Opt_Next.Enabled = True
        End If
    End Sub

    Private Sub Cmd_Enc_Save_Back_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cmd_Enc_Save_Back.Click
        Me.Panel_Enc_Options.BringToFront()
    End Sub

    Private Sub Cmd_Enc_Save_Browse_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cmd_Enc_Save_Browse.Click
        If Me.DlgDirOpen_Enc_Output.ShowDialog = Windows.Forms.DialogResult.Cancel Then Exit Sub
        Me.Lbl_Enc_Save_Path.Text = Me.DlgDirOpen_Enc_Output.SelectedPath
        If (Not Me.Lbl_Enc_Carrier_Error.Visible) And (Me.List_Enc_Carriers.Items.Count > 0) Then Me.Cmd_Enc_Save_Start.Enabled = True
        Me.Lbl_Enc_Save_Lbl.Visible = True
        Me.Lbl_Enc_Save_Path.Visible = True
    End Sub

    Private Sub Cmd_Enc_Carriers_Add_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cmd_Enc_Carriers_Add.Click

        If Me.DlgOpen_Enc_Carrier.ShowDialog() = Windows.Forms.DialogResult.Cancel Then Exit Sub

        For Each Path As String In Me.DlgOpen_Enc_Carrier.FileNames
            If Not Enc_CarrierPaths.ContainsKey(Path) Then Enc_CarrierPaths.Add(Path, "")
        Next Path
        Enc_RefreshListView()

    End Sub

    Private Sub Cmd_Enc_Carriers_Remove_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cmd_Enc_Carriers_Remove.Click

        For Each Item As ListViewItem In Me.List_Enc_Carriers.SelectedItems
            Enc_CarrierPaths.Remove(Item.Text)
        Next
        Enc_RefreshListView()

    End Sub

    Private Sub List_Enc_Carriers_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles List_Enc_Carriers.SelectedIndexChanged
        If Me.List_Enc_Carriers.SelectedItems.Count > 0 Then
            Me.Cmd_Enc_Carriers_Remove.Enabled = True
        Else
            Me.Cmd_Enc_Carriers_Remove.Enabled = False
        End If
    End Sub

    Private Sub Cmd_Cancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cmd_Cancel.Click
        If MsgBox("Abort encryption ?", MsgBoxStyle.YesNo + MsgBoxStyle.Question, "Are You Sure ?") = MsgBoxResult.Yes Then
            Me.Progress_Enc.Value = 0
            Me.Panel_Enc_Output.BringToFront()
        End If
    End Sub

    Private Sub Cmd_Enc_Save_Start_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cmd_Enc_Save_Start.Click
        Me.Panel_Enc_Progress.BringToFront()
        Enc_Encrypt(Me.DlgOpen_Enc_Input.FileName, Me.DlgDirOpen_Enc_Output.SelectedPath, Me.Txt_Enc_Password.Text)
    End Sub

    Private Sub Frm_Main_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If Not (MsgBox("Are you sure you want to exit ?", MsgBoxStyle.YesNoCancel + MsgBoxStyle.Question, "Exit ?") = MsgBoxResult.Yes) Then e.Cancel = True
    End Sub

    Private Sub Cmd_Dec_Cancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cmd_Dec_Cancel.Click
        If MsgBox("Abort decryption ?", MsgBoxStyle.YesNo + MsgBoxStyle.Question, "Are You Sure ?") = MsgBoxResult.Yes Then
            Me.Progress_Dec.Value = 0
            Me.Panel_Dec_Input.BringToFront()
        End If
    End Sub

    Private Sub Txt_Dec_Password_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Txt_Dec_Password.TextChanged
        If Txt_Dec_Password.Text = "" Then
            Me.Cmd_dec_Input_Next.Enabled = False
        Else
            If Me.List_Dec_Carriers.Items.Count > 0 Then Me.Cmd_Dec_Input_Next.Enabled = True
        End If
    End Sub

    Private Sub Cmd_Dec_Input_Back_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cmd_Dec_Input_Back.Click
        Me.Panel_Welcome.BringToFront()
    End Sub

    Private Sub Cmd_Dec_Input_Next_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cmd_Dec_Input_Next.Click
        If Dec_ValidateCarriers() Then Me.Panel_Dec_Output.BringToFront()
    End Sub

    Private Sub Cmd_Dec_Carriers_Add_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cmd_Dec_Carriers_Add.Click

        If Me.DlgOpen_Dec_Carrier.ShowDialog() = Windows.Forms.DialogResult.Cancel Then Exit Sub

        Dim Count As Integer = 0
        Dim ItemArr() As ListViewItem
        For Each Path As String In Me.DlgOpen_Dec_Carrier.FileNames
            If Not Me.List_Dec_Carriers.Items.ContainsKey(Path) Then
                ReDim Preserve ItemArr(Count)
                ItemArr(Count) = New ListViewItem(Path)
                ItemArr(Count).Name = Path
                Count += 1
            End If
        Next Path

        If Count > 0 Then
            Me.List_Dec_Carriers.Items.AddRange(ItemArr)
            Me.List_Dec_Carriers.Refresh()
            If Me.Txt_Dec_Password.Text <> "" Then Me.Cmd_Dec_Input_Next.Enabled = True
        Else
            Me.Cmd_Dec_Input_Next.Enabled = False
        End If

    End Sub

    Private Sub Cmd_Dec_Carriers_Remove_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cmd_Dec_Carriers_Remove.Click
        Me.Cmd_Dec_Carriers_Remove.Enabled = False
        For Each Item As ListViewItem In Me.List_Dec_Carriers.SelectedItems
            Item.Remove()
        Next
        If Me.List_Dec_Carriers.Items.Count = 0 Then Me.Cmd_Dec_Input_Next.Enabled = False
        List_Dec_Carriers.Refresh()
    End Sub

    Private Sub List_Dec_Carriers_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles List_Dec_Carriers.SelectedIndexChanged
        If Me.List_Dec_Carriers.SelectedItems.Count > 0 Then
            Me.Cmd_Dec_Carriers_Remove.Enabled = True
        Else
            Me.Cmd_Dec_Carriers_Remove.Enabled = False
        End If
    End Sub

    Private Sub Cmd_Dec_Output_Back_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cmd_Dec_Output_Back.Click
        Me.Panel_Dec_Input.BringToFront()
    End Sub

    Private Sub Cmd_Dec_Save_Browse_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cmd_Dec_Save_Browse.Click
        If Me.DlgDirOpen_Dec_Output.ShowDialog = Windows.Forms.DialogResult.Cancel Then Exit Sub
        Me.Lbl_Dec_Save_Path.Text = Me.DlgDirOpen_Dec_Output.SelectedPath
        Me.Cmd_Dec_Start.Enabled = True
        Me.Lbl_Dec_Save_Lbl.Visible = True
        Me.Lbl_Dec_Save_Path.Visible = True
    End Sub

    Private Sub Cmd_Dec_Start_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cmd_Dec_Start.Click
        Me.Panel_Dec_Progress.BringToFront()
        Dec_Decrypt(Me.DlgDirOpen_Dec_Output.SelectedPath)
    End Sub

End Class
