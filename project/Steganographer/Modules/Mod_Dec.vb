Module Mod_Dec

#Region "Global Variables"

    Public Dec_CarrierPaths As Hashtable
    Public Dec_FileSize As Long = 0
    Public Dec_FileName As String = ""

    Private Dec_DataBuffer As String
    Private Dec_DataBufferAux As String
    Private F As IO.FileStream

    Private Dec_ComponentsPerByte As Integer = 0
    Private Dec_ComponentsPerPixel As Integer = 0
    Private Dec_PasswordBits As String
    Private Dec_PasswordBits_Curr As Integer = 0
    Private Dec_SortedCarrierPaths As Hashtable
    Private Dec_CurrPixelX As Integer = 0
    Private Dec_CurrPixelY As Integer = 0
    Private Dec_CarrierPath_Curr As Integer = 0
    Private Dec_CarrierPath_Total As Integer = 0
    Private Dec_CarrierBMP As Bitmap
    Private Dec_DoneSize As Long = 0
    Private Dec_PasswordBits_Last As Integer = 0

    Private Structure CarrierHeader
        Public Sequence_Total As Integer
        Public Sequence_Curr As Integer
        Public FileSize As Long
        Public Filename As String
        Public ComponentsPerByte As Integer
        Public ComponentsPerPixel As Integer
        Public DataStartX As Integer
        Public DataStartY As Integer
    End Structure

#End Region

#Region "Misc Functions"

    Public Function Dec_ValidateCarriers() As Boolean

        Dim Path As String
        Dim Item As ListViewItem
        Dim Header As CarrierHeader

        Dec_CarrierPaths = New Hashtable
        Dec_SortedCarrierPaths = New Hashtable
        Dec_PreparePassword(Frm_Main.Txt_Dec_Password.Text)
        Dec_FileName = ""

        For Each Item In Frm_Main.List_Dec_Carriers.Items

            Path = Item.Name
            Try
                Dec_CarrierBMP = New Bitmap(Path)
            Catch ex As Exception
                Frm_Main.Lbl_Dec_Error.Text = "Unable to open carrier " + Path
                Frm_Main.Lbl_Dec_Error.Visible = True
                MsgBox(ex.Message, MsgBoxStyle.Critical, "Unable to Open Carrier")
                Return False
            End Try

            If Not Dec_ReadHeader(Header, True) Then
                Frm_Main.Lbl_Dec_Error.Text = "Failed to read data. This can be because of wrong password or image does not contain any data."
                Frm_Main.Lbl_Dec_Error.Visible = True
                MsgBox("Failed to read data. This can be because of wrong password or image does not contain any data." + vbCrLf + Path, MsgBoxStyle.Critical, "Unable to Read Carrier")
                Dec_CarrierBMP.Dispose()
                Return False
            End If

            Dec_CarrierBMP.Dispose()

            If Not Dec_FileName = "" Then
                If (Dec_FileName <> Header.Filename) Or (Dec_FileSize <> Header.FileSize) Or (Dec_CarrierPath_Total <> Header.Sequence_Total) Then
                    Frm_Main.Lbl_Dec_Error.Text = "Carriers do not contain same file. Please select only those carriers containing the same file."
                    Frm_Main.Lbl_Dec_Error.Visible = True
                    MsgBox("Carriers do not contain same file. Please select only those carriers containing the same file." + vbCrLf + Path, MsgBoxStyle.Critical, "Multiple Files to Extract")
                    Return False
                End If
            Else
                Dec_FileName = Header.Filename
                Dec_FileSize = Header.FileSize
                Dec_CarrierPath_Total = Header.Sequence_Total
                Dec_ComponentsPerByte = Header.ComponentsPerByte
                Dec_ComponentsPerPixel = Header.ComponentsPerPixel
            End If

            If Dec_SortedCarrierPaths.Contains(Header.Sequence_Curr) Then
                Frm_Main.Lbl_Dec_Error.Text = "Carriers do not contain same file. Please select only those carriers containing the same file."
                Frm_Main.Lbl_Dec_Error.Visible = True
                MsgBox("Carriers do not contain same file. Please select only those carriers containing the same file." + vbCrLf + Path, MsgBoxStyle.Critical, "Multiple Files to Extract")
                Return False
            Else
                Dec_SortedCarrierPaths.Add(Header.Sequence_Curr, Path)
                Dec_CarrierPaths.Add(Path, Header)
            End If

        Next Item

        Dim Count As Integer = 0
        While Count < Dec_CarrierPath_Total
            If Not Dec_SortedCarrierPaths.Contains(Count) Then
                Frm_Main.Lbl_Dec_Error.Text = "Sequence incomplete. The data for the file is stored is more carriers. Please select all of the carrier files."
                Frm_Main.Lbl_Dec_Error.Visible = True
                MsgBox("Sequence incomplete. The data for the file is stored is more carriers. Please select all of the carrier files.", MsgBoxStyle.Critical, "Sequence Incomplete")
                Return False
            End If
            Count += 1
        End While

        Frm_Main.Lbl_Dec_Filename.Text = Dec_FileName
        Frm_Main.Lbl_Dec_Size.Text = String.Format("{0:###,###,###,###,###} Bytes", Dec_FileSize)
        Frm_Main.Lbl_Dec_Error.Visible = False
        Return True

    End Function

#End Region

#Region "Helper Functions"

    Private Function Dec_XORBits(ByVal DataBits As String) As String

        If OverrideXOR Then Return DataBits

        Dim Count As Integer = 0
        Dim XORBits As String = ""
        Dim XORBit As String

        While Count < DataBits.Length
            XORBit = Trim(Mid(DataBits, Count + 1, 1)) Xor Trim(Mid(Dec_PasswordBits, Dec_PasswordBits_Curr + 1, 1))
            XORBits = XORBits + XORBit
            Dec_PasswordBits_Curr += 1
            If Dec_PasswordBits_Curr >= Dec_PasswordBits.Length Then Dec_PasswordBits_Curr = 0
            Count += 1
        End While

        Return XORBits

    End Function

    Private Sub Dec_PreparePassword(ByVal Password As String)

        Dim Count As Integer = 0
        Dim ByteArr(Password.Length - 1) As Byte

        ByteArr = StringToByteArr(Password, Password.Length)
        Dec_PasswordBits = ""

        While Count < Password.Length
            Dec_PasswordBits = Dec_PasswordBits + DecimalToBinary(ByteArr(Count))
            Count += 1
        End While

    End Sub

    Private Function Dec_NextPixel() As Boolean

        Dec_CurrPixelX = Dec_CurrPixelX + 1
        If Dec_CurrPixelX >= Dec_CarrierBMP.Width Then
            Dec_CurrPixelX = 0
            Dec_CurrPixelY = Dec_CurrPixelY + 1
        End If

        If Dec_CurrPixelY >= Dec_CarrierBMP.Height Then
            Return False
        Else
            Return True
        End If

    End Function

    Private Function Dec_ReadComponent(ByVal ComponentValue As Integer, ByVal ComponentsPerByte As Integer) As Integer

        Dim Base As Integer

        If ComponentsPerByte = 8 Then
            Base = 2
        ElseIf ComponentsPerByte = 3 Then
            Base = 10
        Else
            Base = 16
        End If

        Return ComponentValue Mod Base

    End Function

#End Region

#Region "Main Functions"

    Public Sub Dec_Decrypt(ByVal SaveDir As String)

        Dec_CarrierPath_Curr = -1
        Dec_DoneSize = 0
        If Trim(Mid(SaveDir, SaveDir.Length, 1)) <> "\" Then SaveDir = SaveDir + "\"
        Dec_DataBuffer = ""
        Dec_DataBufferAux = ""

        Try
            F = New IO.FileStream(SaveDir + Dec_FileName, IO.FileMode.Create, IO.FileAccess.Write, IO.FileShare.None)
        Catch ex As Exception
            Frm_Main.Lbl_Dec_Error.Text = "Unable to open output file for writing : " + SaveDir + Dec_FileName
            Frm_Main.Lbl_Dec_Error.Visible = True
            MsgBox("Unable to open output file for writing : " + SaveDir + Dec_FileName + vbCrLf + ex.Message, MsgBoxStyle.Critical, "Cannot Extract File")
            Frm_Main.Panel_Dec_Output.BringToFront()
            Exit Sub
        End Try

        ' Start Decryption

NextCarrier:
        If Dec_CarrierPath_Curr < (Dec_CarrierPath_Total - 1) Then
            If Not Dec_NextCarrier() Then
                Frm_Main.Progress_Dec.Value = 0
                Frm_Main.Panel_Dec_Output.BringToFront()
                F.Close()
                F.Dispose()
                Exit Sub
            End If
        Else
            Frm_Main.Lbl_Dec_Error.Text = "Carriers exhausted before required file size could be reached. Check password."
            Frm_Main.Lbl_Dec_Error.Visible = True
            MsgBox("Carriers exhausted before required file size could be reached. Check password.", MsgBoxStyle.Critical, "Carrier Exhausted")
            Frm_Main.Panel_Dec_Output.BringToFront()
            F.Close()
            F.Dispose()
            Exit Sub
        End If

NextPixel:
        If Not Dec_ReadPixel(Dec_CarrierBMP.GetPixel(Dec_CurrPixelX, Dec_CurrPixelY)) Then
            Frm_Main.Lbl_Dec_Error.Text = "Cannot Read Carrier. Invalid Password."
            Frm_Main.Lbl_Dec_Error.Visible = True
            MsgBox("Cannot Read Carrier. Invalid Password.", MsgBoxStyle.Critical, "Decryption Failed.")
            Frm_Main.Panel_Dec_Output.BringToFront()
            F.Close()
            F.Dispose()
            Exit Sub
        End If
        If Dec_DoneSize = Dec_FileSize Then GoTo Done
        Frm_Main.Progress_Dec.Value = (Dec_DoneSize / Dec_FileSize) * 100
        Application.DoEvents()

        If Dec_NextPixel() Then
            GoTo NextPixel
        Else
            GoTo NextCarrier
        End If

        ' Finished
Done:
        F.Close()
        F.Dispose()
        Dec_CarrierBMP.Dispose()

        Frm_Main.Lbl_Dec_Status.Text = "Completed"
        Frm_Main.Cmd_Dec_Cancel.Enabled = False
        MsgBox("Decryption Finished", MsgBoxStyle.Information, "Decryption Completed")
        Frm_Main.Panel_Welcome.BringToFront()
        Frm_Main.Progress_Dec.Value = 0
        Frm_Main.Lbl_Dec_Status.Text = "Decrypting ..."
        Frm_Main.Cmd_Dec_Cancel.Enabled = True
        Exit Sub

    End Sub

    Private Function Dec_ReadPixel(ByVal C As Color) As Boolean

        On Error GoTo Err

        Dim A_Data As Integer = Dec_ReadComponent(C.A, Dec_ComponentsPerByte)
        Dim R_Data As Integer = Dec_ReadComponent(C.R, Dec_ComponentsPerByte)
        Dim G_Data As Integer = Dec_ReadComponent(C.G, Dec_ComponentsPerByte)
        Dim B_Data As Integer = Dec_ReadComponent(C.B, Dec_ComponentsPerByte)

        If Dec_ComponentsPerByte = 2 Then

            If Dec_ComponentsPerPixel = 3 Then

                If Dec_DataBuffer.Length = 0 Then

                    Dim HexArr(1) As String
                    HexArr(0) = R_Data.ToString
                    HexArr(1) = G_Data.ToString
                    Dim Data As Integer = HexToDecimal(HexArr)
                    F.WriteByte(BinaryToDecimal(Dec_XORBits(DecimalToBinary(Data))))
                    Dec_DoneSize += 1
                    Dec_DataBuffer = B_Data.ToString

                Else

                    Dim HexArr(1) As String
                    HexArr(0) = Dec_DataBuffer
                    HexArr(1) = R_Data.ToString
                    Dim Data As Integer = HexToDecimal(HexArr)
                    F.WriteByte(BinaryToDecimal(Dec_XORBits(DecimalToBinary(Data))))
                    Dec_DoneSize += 1
                    Dec_DataBuffer = ""

                    If Dec_DoneSize < Dec_FileSize Then
                        HexArr(0) = G_Data.ToString
                        HexArr(1) = B_Data.ToString
                        Data = HexToDecimal(HexArr)
                        F.WriteByte(BinaryToDecimal(Dec_XORBits(DecimalToBinary(Data))))
                        Dec_DoneSize += 1
                    End If

                End If

            Else

                Dim HexArr(1) As String
                HexArr(0) = A_Data.ToString
                HexArr(1) = R_Data.ToString
                Dim Data As Integer = HexToDecimal(HexArr)
                F.WriteByte(BinaryToDecimal(Dec_XORBits(DecimalToBinary(Data))))
                Dec_DoneSize += 1

                If Dec_DoneSize < Dec_FileSize Then

                    HexArr(0) = G_Data.ToString
                    HexArr(1) = B_Data.ToString
                    Data = HexToDecimal(HexArr)
                    F.WriteByte(BinaryToDecimal(Dec_XORBits(DecimalToBinary(Data))))
                    Dec_DoneSize += 1

                End If

            End If

        ElseIf Dec_ComponentsPerByte = 3 Then

            If Dec_ComponentsPerPixel = 3 Then
                Dim Data As Integer = (R_Data * 100) + (G_Data * 10) + (B_Data)
                F.WriteByte(BinaryToDecimal(Dec_XORBits(DecimalToBinary(Data))))
                Dec_DoneSize += 1
            Else

                If Dec_DataBuffer.Length = 0 Then

                    Dim Data As Integer = (A_Data * 100) + (R_Data * 10) + (G_Data)
                    F.WriteByte(BinaryToDecimal(Dec_XORBits(DecimalToBinary(Data))))
                    Dec_DataBuffer = B_Data.ToString
                    Dec_DoneSize += 1

                ElseIf Dec_DataBufferAux.Length = 0 Then

                    Dim Data As Integer = (Integer.Parse(Dec_DataBuffer) * 100) + (A_Data * 10) + (R_Data)
                    F.WriteByte(BinaryToDecimal(Dec_XORBits(DecimalToBinary(Data))))
                    Dec_DataBuffer = G_Data.ToString
                    Dec_DataBufferAux = B_Data.ToString
                    Dec_DoneSize += 1

                Else

                    Dim Data As Integer = (Integer.Parse(Dec_DataBuffer) * 100) + (Integer.Parse(Dec_DataBufferAux) * 10) + (A_Data)
                    F.WriteByte(BinaryToDecimal(Dec_XORBits(DecimalToBinary(Data))))
                    Dec_DoneSize += 1

                    If Dec_DoneSize < Dec_FileSize Then
                        Data = (R_Data * 100) + (G_Data * 10) + (B_Data)
                        F.WriteByte(BinaryToDecimal(Dec_XORBits(DecimalToBinary(Data))))
                        Dec_DoneSize += 1
                    End If

                    Dec_DataBuffer = ""
                    Dec_DataBufferAux = ""

                End If

            End If

        Else

            If Dec_ComponentsPerPixel = 3 Then
                Dec_DataBuffer = Dec_DataBuffer + R_Data.ToString + G_Data.ToString + B_Data.ToString
            Else
                Dec_DataBuffer = Dec_DataBuffer + A_Data.ToString + R_Data.ToString + G_Data.ToString + B_Data.ToString
            End If

            Dim Data As String
            While (Dec_DataBuffer.Length >= 8) And (Dec_DoneSize < Dec_FileSize)
                Data = Dec_XORBits(Trim(Mid(Dec_DataBuffer, 1, 8)))
                F.WriteByte(BinaryToDecimal(Data))

                If Dec_DataBuffer.Length = 8 Then
                    Dec_DataBuffer = ""
                Else
                    Dec_DataBuffer = Trim(Mid(Dec_DataBuffer, 9))
                End If

                Dec_DoneSize += 1
            End While

        End If

        Return True

Err:
        Return False

    End Function

    Private Function Dec_NextCarrier() As Boolean

        If Dec_CarrierPath_Curr > -1 Then Dec_CarrierBMP.Dispose()
        Dec_CarrierPath_Curr += 1
        Dec_PasswordBits_Last = Dec_PasswordBits_Curr
        Dec_PasswordBits_Curr = 0

        Dim Path As String
        Dim Header As CarrierHeader

        Try
            Path = Dec_SortedCarrierPaths.Item(Dec_CarrierPath_Curr)
            Header = Dec_CarrierPaths.Item(Path)
            Dec_CurrPixelX = Header.DataStartX
            Dec_CurrPixelY = Header.DataStartY
        Catch ex As Exception
            Frm_Main.Lbl_Dec_Error.Text = "Error in image sequence : " + Dec_CarrierPath_Curr.ToString + " / " + Dec_CarrierPath_Total.ToString + ". Wrong carrier or invalid password."
            Frm_Main.Lbl_Dec_Error.Visible = True
            MsgBox("Error in image sequence : " + Dec_CarrierPath_Curr.ToString + " / " + Dec_CarrierPath_Total.ToString + vbCrLf + ex.Message, MsgBoxStyle.Critical, "Carrier Does Not Exist")
            Return False
        End Try

        Try
            Dec_CarrierBMP = New Bitmap(Path)
        Catch ex As Exception
            Frm_Main.Lbl_Dec_Error.Text = "Unable to open carrier " + Path
            Frm_Main.Lbl_Dec_Error.Visible = True
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Unable to Open Carrier")
            Return False
        End Try

        Dec_PasswordBits_Curr = Dec_PasswordBits_Last
        Return True

    End Function

    Private Function Dec_ReadHeader(ByRef Header As CarrierHeader, Optional ByVal ResetXY As Boolean = False) As Boolean

        Dim Clr As Color
        Dim Data As String

        If ResetXY Then
            Dec_CurrPixelX = 0
            Dec_CurrPixelY = 0
        End If
        Dec_PasswordBits_Last = Dec_PasswordBits_Curr
        Dec_PasswordBits_Curr = 0

        'Read PixelFormat and Noise Mode

        Clr = Dec_CarrierBMP.GetPixel(Dec_CurrPixelX, Dec_CurrPixelY)
        Header.ComponentsPerPixel = IIf(Dec_XORBits(Dec_ReadComponent(Clr.R, 8)) = "0", 3, 4)
        If Dec_XORBits(Dec_ReadComponent(Clr.G, 8)) = "1" Then
            Header.ComponentsPerByte = 8
        Else
            If Dec_XORBits(Dec_ReadComponent(Clr.B, 8)) = "1" Then
                Header.ComponentsPerByte = 3
            Else
                Header.ComponentsPerByte = 2
            End If
        End If
        If Not Dec_NextPixel() Then Return False
        Dec_PasswordBits_Curr = 0

        ' Read Total Carriers

        Clr = Dec_CarrierBMP.GetPixel(Dec_CurrPixelX, Dec_CurrPixelY)
        Data = Dec_XORBits(Dec_ReadComponent(Clr.R, 8).ToString + Dec_ReadComponent(Clr.G, 8).ToString + Dec_ReadComponent(Clr.B, 8).ToString)
        Header.Sequence_Total = BinaryToDecimal(Data) + 1
        If Not Dec_NextPixel() Then Return False
        Dec_PasswordBits_Curr = 0

        ' Read Current Carriers

        Clr = Dec_CarrierBMP.GetPixel(Dec_CurrPixelX, Dec_CurrPixelY)
        Data = Dec_XORBits(Dec_ReadComponent(Clr.R, 8).ToString + Dec_ReadComponent(Clr.G, 8).ToString + Dec_ReadComponent(Clr.B, 8).ToString)
        Header.Sequence_Curr = BinaryToDecimal(Data)
        If Not Dec_NextPixel() Then Return False
        Dec_PasswordBits_Curr = 0

        ' Logic Check for Sequence
        If Header.Sequence_Curr >= Header.Sequence_Total Then Return False

        ' Read File Size

        Dim Count As Integer = 0
        Data = ""

        While Count < 9
            Clr = Dec_CarrierBMP.GetPixel(Dec_CurrPixelX, Dec_CurrPixelY)
            Data = Data + Dec_XORBits(Dec_ReadComponent(Clr.R, 8))
            Data = Data + Dec_XORBits(Dec_ReadComponent(Clr.G, 8))
            Data = Data + Dec_XORBits(Dec_ReadComponent(Clr.B, 8))

            If Not Dec_NextPixel() Then Return False
            Count += 1
        End While

        Header.FileSize = BinaryToDecimal(Data)
        Dec_PasswordBits_Curr = 0
        ' Dont Need Next Pixel due to entry controlled loop

        ' Read File Name Length

        Dim FileNameL As Integer
        Count = 0
        Data = ""
        Dec_PasswordBits_Curr = 0

        While Count < 2
            Clr = Dec_CarrierBMP.GetPixel(Dec_CurrPixelX, Dec_CurrPixelY)
            Data = Data + Dec_XORBits(Dec_ReadComponent(Clr.R, 8))
            Data = Data + Dec_XORBits(Dec_ReadComponent(Clr.G, 8))
            Data = Data + Dec_XORBits(Dec_ReadComponent(Clr.B, 8))

            If Not Dec_NextPixel() Then Return False
            Count += 1
        End While
        FileNameL = BinaryToDecimal(Data)
        Dec_PasswordBits_Curr = 0
        ' Dont Need Next Pixel due to entry controlled loop

        ' Read File Name

        Dim FileNameByteArr(FileNameL - 1) As Byte
        Data = ""
        Count = 0
        Header.Filename = ""

        While Count < FileNameL

            Clr = Dec_CarrierBMP.GetPixel(Dec_CurrPixelX, Dec_CurrPixelY)
            Data = Data + Dec_XORBits(Dec_ReadComponent(Clr.R, 8))
            Data = Data + Dec_XORBits(Dec_ReadComponent(Clr.G, 8))
            Data = Data + Dec_XORBits(Dec_ReadComponent(Clr.B, 8))

            If Data.Length >= 8 Then
                Header.Filename = Header.Filename + Chr(BinaryToDecimal(Trim(Mid(Data, 1, 8))))
                If Data.Length > 8 Then
                    Data = Trim(Mid(Data, 9))
                Else
                    Data = ""
                End If
                Count += 1
            End If

            If Not Dec_NextPixel() Then Return False
        End While
        Dec_PasswordBits_Curr = 0
        ' Dont Need Next Pixel due to entry controlled loop

        'Logic Check on Valid File name
        If Header.Filename.Contains(":") Or Header.Filename.Contains("!") Or Header.Filename.Contains("*") Or Header.Filename.Contains("%") Or Header.Filename.Contains("|") _
        Or Header.Filename.Contains("<") Or Header.Filename.Contains(">") Or Header.Filename.Contains("{") Or Header.Filename.Contains("}") Or Header.Filename.Contains("/") Or Header.Filename.Contains("\") _
        Or Header.Filename.Contains("?") Then Return False

        ' Verify Checksum File Length
        Count = 0
        Data = ""

        While Count < 9
            Clr = Dec_CarrierBMP.GetPixel(Dec_CurrPixelX, Dec_CurrPixelY)
            Data = Data + Dec_XORBits(Dec_ReadComponent(Clr.R, 8))
            Data = Data + Dec_XORBits(Dec_ReadComponent(Clr.G, 8))
            Data = Data + Dec_XORBits(Dec_ReadComponent(Clr.B, 8))

            If Not Dec_NextPixel() Then Return False
            Count += 1
        End While
        If Header.FileSize <> BinaryToDecimal(Data) Then Return False
        Dec_PasswordBits_Curr = 0

        ' If Not Dec_NextPixel() Then Return False
        Header.DataStartX = Dec_CurrPixelX
        Header.DataStartY = Dec_CurrPixelY
        Return True

    End Function

#End Region



End Module
