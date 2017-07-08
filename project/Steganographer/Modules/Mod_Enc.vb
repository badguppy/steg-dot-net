Module Mod_Enc

#Region "Global Variables"

    Public Enc_CarrierPaths As New Hashtable
    Public Enc_TotalCapacity As Long = 0
    Public Enc_ComponentsPerByte As Integer = 0
    Public Enc_ComponentsPerPixel As Integer = 0
    Public Enc_FileSize As Long = 0
    Public Enc_FileName As String = ""

    Private Enc_PasswordBits As String
    Private Enc_PasswordBits_Curr As Integer = 0
    Private Enc_SortedCarrierPaths As String()
    Private Enc_CurrPixelX As Integer = 0
    Private Enc_CurrPixelY As Integer = 0
    Private Enc_CurrColor As Color
    Private Enc_A As Integer
    Private Enc_R As Integer
    Private Enc_G As Integer
    Private Enc_B As Integer
    Private Enc_ComponentCurr As Integer = 0
    Private Enc_CarrierPath_Curr As Integer = 0
    Private Enc_CarrierPath_Total As Integer = 0
    Private Enc_CarrierBMP As Bitmap
    Private Enc_FinalBMP As Bitmap
    Private Enc_DoneSize As Long = 0

    Private Enc_PasswordBits_Last As Integer = 0

#End Region

#Region "Misc Functions"

    Public Function Enc_RefreshListView() As Boolean

        Dim Count As Integer = 0
        Dim ItemArr() As ListViewItem
        Dim BMP As Bitmap
        Dim Capacity As Long = 0
        Dim HTNew As New Hashtable

        Enc_RefreshListView = True
        Enc_TotalCapacity = 0
        Frm_Main.List_Enc_Carriers.Items.Clear()
        Frm_Main.Cmd_Enc_Carriers_Remove.Enabled = False

        For Each Path As String In Enc_CarrierPaths.Keys
            Try

                BMP = New Bitmap(Path)
                Capacity = ((BMP.Width * BMP.Height * Enc_ComponentsPerPixel) / Enc_ComponentsPerByte) - Path.Length - Enc_GetHeaderSize(Frm_Main.Lbl_Enc_Input.Text) 'Header Width
                Enc_TotalCapacity += Capacity

                ReDim Preserve ItemArr(Count)
                ItemArr(Count) = New ListViewItem(Path)
                ItemArr(Count).SubItems.Add(BMP.Width.ToString + "x" + BMP.Height.ToString)
                ItemArr(Count).SubItems.Add(String.Format("{0:###,###,###,###,###}", Capacity))
                ItemArr(Count).Name = Path
                HTNew.Add(Path, Capacity)

                BMP.Dispose()
                Count += 1

            Catch ex As Exception
                MsgBox("Failed to open " + Path + ". " + ex.Message, MsgBoxStyle.Critical, "Failed to Open Carrier Image")
                Enc_RefreshListView = False
            End Try
        Next Path

        Enc_CarrierPaths = HTNew

        If Count > 0 Then

            Frm_Main.List_Enc_Carriers.Items.AddRange(ItemArr)
            Frm_Main.List_Enc_Carriers.Refresh()

            Frm_Main.Lbl_Enc_Capacity.Text = String.Format("{0:###,###,###,###,###} Bytes", Enc_TotalCapacity)
            If Enc_FileSize > Enc_TotalCapacity Then
                Frm_Main.Lbl_Enc_Carrier_Error.Visible = True
                Frm_Main.Cmd_Enc_Save_Start.Enabled = False
                Enc_RefreshListView = False
            Else
                Frm_Main.Lbl_Enc_Carrier_Error.Visible = False
                If Frm_Main.Lbl_Enc_Save_Path.Visible Then Frm_Main.Cmd_Enc_Save_Start.Enabled = True
            End If

        Else
            Frm_Main.Lbl_Enc_Carrier_Error.Visible = False
            Frm_Main.Lbl_Enc_Capacity.Text = "0 Bytes"
            Frm_Main.Cmd_Enc_Save_Start.Enabled = False
        End If

    End Function

#End Region

#Region "Helper Functions"

    Private Function Enc_XORBits(ByVal DataBits As String) As String

        If OverrideXOR Then Return DataBits

        Dim Count As Integer = 0
        Dim XORBits As String = ""
        Dim XORBit As String

        While Count < DataBits.Length
            XORBit = Trim(Mid(DataBits, Count + 1, 1)) Xor Trim(Mid(Enc_PasswordBits, Enc_PasswordBits_Curr + 1, 1))
            XORBits = XORBits + XORBit
            Enc_PasswordBits_Curr += 1
            If Enc_PasswordBits_Curr >= Enc_PasswordBits.Length Then Enc_PasswordBits_Curr = 0
            Count += 1
        End While

        Return XORBits

    End Function

    Private Function Enc_SortCarrierPaths() As Boolean

        ' Refresh Carrier Path Capacities and Perform FileExists Alike Functions
        If Not Enc_RefreshListView() Then Return False

        Dim Count As Integer = 0
        Dim Pass As Integer = 0
        Dim Path As String

        ' Copy Carrier Paths from HT to Array

        ReDim Enc_SortedCarrierPaths(Enc_CarrierPaths.Keys.Count - 1)
        For Each Path In Enc_CarrierPaths.Keys
            Enc_SortedCarrierPaths(Count) = Path
            Count += 1
        Next

        ' Sort Array of Carrier Paths by Capacity

        While Pass < Enc_CarrierPaths.Keys.Count
            Count = 0
            While Count < Enc_CarrierPaths.Keys.Count - Pass - 1
                If Enc_CarrierPaths.Item(Enc_SortedCarrierPaths(Count)) < Enc_CarrierPaths.Item(Enc_SortedCarrierPaths(Count + 1)) Then
                    Path = Enc_SortedCarrierPaths(Count)
                    Enc_SortedCarrierPaths(Count) = Enc_SortedCarrierPaths(Count + 1)
                    Enc_SortedCarrierPaths(Count + 1) = Path
                End If
                Count = Count + 1
            End While
            Pass = Pass + 1
        End While

        'Determine Maximum Carriers Necessary, Leave out the rest

        Dim Capacity As Long = 0
        Count = 0
        Enc_CarrierPath_Total = 0

        While Capacity < Enc_FileSize
            Capacity += Enc_CarrierPaths.Item(Enc_SortedCarrierPaths(Count))
            Enc_CarrierPath_Total += 1
            Count += 1
        End While

        If Enc_CarrierPath_Total < Enc_SortedCarrierPaths.Count Then MsgBox("You have selected more carrier images than necessary. Few carriers will not be used.", MsgBoxStyle.Information, "Too Many Carriers")
        Return True

    End Function

    Private Sub Enc_PreparePassword(ByVal Password As String)

        Dim Count As Integer = 0
        Dim ByteArr(Password.Length - 1) As Byte

        ByteArr = StringToByteArr(Password, Password.Length)
        Enc_PasswordBits = ""

        While Count < Password.Length
            Enc_PasswordBits = Enc_PasswordBits + DecimalToBinary(ByteArr(Count))
            Count += 1
        End While

    End Sub

    Private Function Enc_NextPixel(Optional ByVal ExitOnPixelExhaustion As Boolean = False) As Boolean

        Enc_CurrPixelX = Enc_CurrPixelX + 1
        If Enc_CurrPixelX >= Enc_FinalBMP.Width Then
            Enc_CurrPixelX = 0
            Enc_CurrPixelY = Enc_CurrPixelY + 1
        End If

        If Enc_CurrPixelY >= Enc_FinalBMP.Height Then
            If ExitOnPixelExhaustion Then
                MsgBox("Bitmap pixels exhausted. Please use a carrier large enough to accomodate the header, prefferably larger than 256 pixels.", MsgBoxStyle.Critical, "Encryption Failed")
                Frm_Main.Progress_Enc.Value = 0
                Frm_Main.Panel_Enc_Output.BringToFront()
                Return False
            Else
                Return False
            End If
        Else
            Return True
        End If

    End Function

    Private Function Enc_WriteComponent(ByVal ComponentValue As Integer, ByVal DataValue As Integer, ByVal ComponentsPerByte As Integer) As Integer

        Dim Base As Integer

        If ComponentsPerByte = 8 Then
            Base = 2
        ElseIf ComponentsPerByte = 3 Then
            Base = 10
        Else
            Base = 16
        End If

        Dim R As Integer = ComponentValue Mod Base
        If R <> DataValue Then
            Dim Diff As Double = R - DataValue
            If ComponentValue - Diff > 255 Then
                ComponentValue = ComponentValue - Base - Diff
            ElseIf ComponentValue - Diff < 0 Then
                ComponentValue = ComponentValue + Base - Diff
            Else
                ComponentValue = ComponentValue - Diff
            End If
        End If

        Return ComponentValue

    End Function

    Private Function Enc_GetHeaderSize(ByVal Path As String) As Long
        Return 23 + (GetFileName(Path).Length * 8 / 3)
    End Function

#End Region

#Region "Main Functions"

    Public Sub Enc_Encrypt(ByVal InFile As String, ByVal SaveDir As String, ByVal Password As String)

        ' Initialize Variables For a New Encryption Session

        If Not Enc_SortCarrierPaths() Then
            MsgBox("Error opening carrier images or images capacity is less than the size of the file to be encrypted.", MsgBoxStyle.Critical, "Encryption Failed")
            Frm_Main.Progress_Enc.Value = 0
            Frm_Main.Panel_Enc_Output.BringToFront()
            Exit Sub
        End If

        Enc_PreparePassword(Password)
        Enc_PasswordBits_Curr = 0
        Enc_CurrPixelX = 0
        Enc_CurrPixelY = 0
        Enc_ComponentCurr = 0
        Enc_CarrierPath_Curr = -1
        Enc_DoneSize = 0
        Enc_FileName = GetFileName(InFile)

        ' Open Input File

        Dim F As IO.FileStream
        Try
            F = New IO.FileStream(InFile, IO.FileMode.Open, IO.FileAccess.Read, IO.FileShare.Read)
        Catch ex As Exception
            MsgBox("Error opening input file for reading. " + ex.Message, MsgBoxStyle.Critical, "Encryption Failed")
            Frm_Main.Progress_Enc.Value = 0
            Frm_Main.Panel_Enc_Output.BringToFront()
            Exit Sub
        End Try

        ' Open a New Carrier and Initialize Carrier BMP and Final BMP

        If Not Enc_NextCarrier(SaveDir) Then
            MsgBox("Error opening carrier file for reading. ", MsgBoxStyle.Critical, "Encryption Failed")
            Frm_Main.Progress_Enc.Value = 0
            Frm_Main.Panel_Enc_Output.BringToFront()
            F.Close()
            F.Dispose()
            Exit Sub
        End If

        ' Main Loop

        While Enc_DoneSize < F.Length
            If Not Enc_WriteByte(BinaryToDecimal(Enc_XORBits(DecimalToBinary(F.ReadByte()))), SaveDir) Then
                MsgBox("Error opening carrier file for reading. ", MsgBoxStyle.Critical, "Encryption Failed")
                Frm_Main.Progress_Enc.Value = 0
                Frm_Main.Panel_Enc_Output.BringToFront()
                F.Close()
                F.Dispose()
                Exit Sub
            End If
            Enc_DoneSize += 1

            Frm_Main.Progress_Enc.Value = (Enc_DoneSize / Enc_FileSize) * 100
            Application.DoEvents()
        End While

        ' Finished

        F.Close()
        F.Dispose()

        While Enc_NextPixel()
            Enc_FinalBMP.SetPixel(Enc_CurrPixelX, Enc_CurrPixelY, Enc_CarrierBMP.GetPixel(Enc_CurrPixelX, Enc_CurrPixelY))
        End While

        If Not Enc_SaveCarrier(SaveDir) Then
            MsgBox("Error saving output.", MsgBoxStyle.Critical, "Encryption Failed")
            Frm_Main.Progress_Enc.Value = 0
            Frm_Main.Panel_Enc_Output.BringToFront()
            Exit Sub
        End If

        Frm_Main.Lbl_Enc_Status.Text = "Completed"
        Frm_Main.Cmd_Cancel.Enabled = False
        MsgBox("Encryption Finished", MsgBoxStyle.Information, "Encryption Completed")
        Frm_Main.Panel_Welcome.BringToFront()
        Frm_Main.Progress_Enc.Value = 0
        Frm_Main.Lbl_Enc_Status.Text = "Encrypting ..."
        Frm_Main.Cmd_Cancel.Enabled = True

    End Sub

    Private Function Enc_WriteByte(ByVal Value As Integer, ByVal SavePath As String) As Boolean

        'Create Data Array

        Dim Count As Integer = 0
        Dim DataArr(Enc_ComponentsPerByte - 1) As String

        If Enc_ComponentsPerByte = 8 Then
            Dim Bin As String = DecimalToBinary(Value)
            Count = 0
            While Count < Bin.Length
                DataArr(Count) = Trim(Mid(Bin, Count + 1, 1))
                Count += 1
            End While
        ElseIf Enc_ComponentsPerByte = 3 Then
            Dim Dec As String = Value.ToString
            While Dec.Length < 3
                Dec = "0" + Dec
            End While
            DataArr(0) = Dec(0)
            DataArr(1) = Dec(1)
            DataArr(2) = Dec(2)
        Else
            DataArr = DecimalToHex(Value)
        End If

        ' Write Data Array

        Count = 0
        While Count < Enc_ComponentsPerByte

RetryWrite:
            If Enc_ComponentCurr = 0 Then
                Enc_A = Enc_WriteComponent(Enc_A, Integer.Parse(DataArr(Count)), Enc_ComponentsPerByte)
            ElseIf Enc_ComponentCurr = 1 Then
                Enc_R = Enc_WriteComponent(Enc_R, Integer.Parse(DataArr(Count)), Enc_ComponentsPerByte)
            ElseIf Enc_ComponentCurr = 2 Then
                Enc_G = Enc_WriteComponent(Enc_G, Integer.Parse(DataArr(Count)), Enc_ComponentsPerByte)
            ElseIf Enc_ComponentCurr = 3 Then
                Enc_B = Enc_WriteComponent(Enc_B, Integer.Parse(DataArr(Count)), Enc_ComponentsPerByte)
            Else
                ' Need Next Pixel

                ' Write Current Pixel
                Enc_FinalBMP.SetPixel(Enc_CurrPixelX, Enc_CurrPixelY, Color.FromArgb(Enc_A, Enc_R, Enc_G, Enc_B))

                ' Increment Pixel Coordinate
                If Not Enc_NextPixel() Then
                    ' If Bitmap Exhausted Then Save Current BMP, Get Next Carrier
                    If Not Enc_NextCarrier(SavePath) Then Return False
                End If

                ' Ready Next Pixel
                If Enc_ComponentsPerPixel = 3 Then
                    Enc_ComponentCurr = 1
                Else
                    Enc_ComponentCurr = 0
                End If
                Enc_CurrColor = Enc_CarrierBMP.GetPixel(Enc_CurrPixelX, Enc_CurrPixelY)
                Enc_A = Enc_CurrColor.A
                Enc_R = Enc_CurrColor.R
                Enc_G = Enc_CurrColor.G
                Enc_B = Enc_CurrColor.B

                ' Write Data
                GoTo RetryWrite

            End If

            If (Enc_FileSize = Enc_DoneSize + 1) And (Count = Enc_ComponentsPerByte - 1) Then Enc_FinalBMP.SetPixel(Enc_CurrPixelX, Enc_CurrPixelY, Color.FromArgb(Enc_A, Enc_R, Enc_G, Enc_B))
            Enc_ComponentCurr += 1
            Count += 1

        End While

        Return True

    End Function

    Private Function Enc_WriteHeader() As Boolean

        Dim Clr As Color
        Dim R As Integer
        Dim G As Integer
        Dim B As Integer

        Enc_PasswordBits_Curr = 0

        'Write PixelFormat and Noise Mode

        Clr = Enc_CarrierBMP.GetPixel(Enc_CurrPixelX, Enc_CurrPixelY)
        R = Enc_WriteComponent(Clr.R, Enc_XORBits(IIf(Enc_ComponentsPerPixel = 3, 0, 1)), 8)
        G = Enc_WriteComponent(Clr.G, Enc_XORBits(IIf(Enc_ComponentsPerByte = 8, 1, 0)), 8)
        B = Enc_WriteComponent(Clr.B, Enc_XORBits(IIf(Enc_ComponentsPerByte = 3, 1, 0)), 8)
        Enc_FinalBMP.SetPixel(Enc_CurrPixelX, Enc_CurrPixelY, Color.FromArgb(Clr.A, R, G, B))
        If Not Enc_NextPixel(True) Then Return False
        Enc_PasswordBits_Curr = 0

        ' Write Total Carriers

        Dim TotalCarrierBin As String = Enc_XORBits(DecimalToBinary(Enc_CarrierPath_Total - 1, True, 3))
        Clr = Enc_CarrierBMP.GetPixel(Enc_CurrPixelX, Enc_CurrPixelY)
        R = Enc_WriteComponent(Clr.R, Integer.Parse(TotalCarrierBin(0)), 8)
        G = Enc_WriteComponent(Clr.G, Integer.Parse(TotalCarrierBin(1)), 8)
        B = Enc_WriteComponent(Clr.B, Integer.Parse(TotalCarrierBin(2)), 8)
        Enc_FinalBMP.SetPixel(Enc_CurrPixelX, Enc_CurrPixelY, Color.FromArgb(Clr.A, R, G, B))
        If Not Enc_NextPixel(True) Then Return False
        Enc_PasswordBits_Curr = 0

        ' Write Curr Carrier

        Dim CurrCarrierBin As String = Enc_XORBits(DecimalToBinary(Enc_CarrierPath_Curr, True, 3))
        Clr = Enc_CarrierBMP.GetPixel(Enc_CurrPixelX, Enc_CurrPixelY)
        R = Enc_WriteComponent(Clr.R, Integer.Parse(CurrCarrierBin(0)), 8)
        G = Enc_WriteComponent(Clr.G, Integer.Parse(CurrCarrierBin(1)), 8)
        B = Enc_WriteComponent(Clr.B, Integer.Parse(CurrCarrierBin(2)), 8)
        Enc_FinalBMP.SetPixel(Enc_CurrPixelX, Enc_CurrPixelY, Color.FromArgb(Clr.A, R, G, B))
        If Not Enc_NextPixel(True) Then Return False
        Enc_PasswordBits_Curr = 0

        ' Write File Length

        Dim FileSizeBin As String = Enc_XORBits(DecimalToBinary(Enc_FileSize, True, 27))
        Dim CurrComponent As Integer = 1
        Dim Count As Integer = 0

        While Count < FileSizeBin.Length

            If CurrComponent = 4 Then
                Enc_FinalBMP.SetPixel(Enc_CurrPixelX, Enc_CurrPixelY, Color.FromArgb(Clr.A, R, G, B))
                If Not Enc_NextPixel(True) Then Return False
                Clr = Enc_CarrierBMP.GetPixel(Enc_CurrPixelX, Enc_CurrPixelY)
                G = Clr.G
                B = Clr.B
                CurrComponent = 1
            End If

            If CurrComponent = 1 Then
                R = Enc_WriteComponent(Clr.R, Integer.Parse(FileSizeBin(Count)), 8)
            ElseIf CurrComponent = 2 Then
                G = Enc_WriteComponent(Clr.G, Integer.Parse(FileSizeBin(Count)), 8)
            Else
                B = Enc_WriteComponent(Clr.B, Integer.Parse(FileSizeBin(Count)), 8)
            End If

            CurrComponent += 1
            Count += 1
        End While

        Enc_FinalBMP.SetPixel(Enc_CurrPixelX, Enc_CurrPixelY, Color.FromArgb(Clr.A, R, G, B))
        If Not Enc_NextPixel(True) Then Return False
        Enc_PasswordBits_Curr = 0

        ' Write FileName Length

        Dim FileNameL As Integer = Enc_FileName.Length
        Dim FileNameLBin As String = Enc_XORBits(DecimalToBinary(FileNameL, True, 6))
        Clr = Enc_CarrierBMP.GetPixel(Enc_CurrPixelX, Enc_CurrPixelY)
        R = Enc_WriteComponent(Clr.R, Integer.Parse(FileNameLBin(0)), 8)
        G = Enc_WriteComponent(Clr.G, Integer.Parse(FileNameLBin(1)), 8)
        B = Enc_WriteComponent(Clr.B, Integer.Parse(FileNameLBin(2)), 8)
        Enc_FinalBMP.SetPixel(Enc_CurrPixelX, Enc_CurrPixelY, Color.FromArgb(Clr.A, R, G, B))
        If Not Enc_NextPixel(True) Then Return False
        Clr = Enc_CarrierBMP.GetPixel(Enc_CurrPixelX, Enc_CurrPixelY)
        R = Enc_WriteComponent(Clr.R, Integer.Parse(FileNameLBin(3)), 8)
        G = Enc_WriteComponent(Clr.G, Integer.Parse(FileNameLBin(4)), 8)
        B = Enc_WriteComponent(Clr.B, Integer.Parse(FileNameLBin(5)), 8)
        Enc_FinalBMP.SetPixel(Enc_CurrPixelX, Enc_CurrPixelY, Color.FromArgb(Clr.A, R, G, B))
        If Not Enc_NextPixel(True) Then Return False
        Enc_PasswordBits_Curr = 0

        ' Write FileName

        Dim FileNameByteArr() As Byte = StringToByteArr(Enc_FileName, Enc_FileName.Length)
        Dim FileNameBin As String = ""
        Count = 0
        While Count < FileNameL
            FileNameBin = FileNameBin + DecimalToBinary(Asc(Enc_FileName(Count)))
            Count += 1
        End While

        CurrComponent = 1
        Count = 0
        FileNameBin = Enc_XORBits(FileNameBin)

        While Count < FileNameBin.Length

            If CurrComponent = 4 Then
                Enc_FinalBMP.SetPixel(Enc_CurrPixelX, Enc_CurrPixelY, Color.FromArgb(Clr.A, R, G, B))
                If Not Enc_NextPixel(True) Then Return False
                Clr = Enc_CarrierBMP.GetPixel(Enc_CurrPixelX, Enc_CurrPixelY)
                G = Clr.G
                B = Clr.B
                CurrComponent = 1
            End If

            If CurrComponent = 1 Then
                R = Enc_WriteComponent(Clr.R, Integer.Parse(FileNameBin(Count)), 8)
            ElseIf CurrComponent = 2 Then
                G = Enc_WriteComponent(Clr.G, Integer.Parse(FileNameBin(Count)), 8)
            Else
                B = Enc_WriteComponent(Clr.B, Integer.Parse(FileNameBin(Count)), 8)
            End If

            CurrComponent += 1
            Count += 1
        End While

        Enc_FinalBMP.SetPixel(Enc_CurrPixelX, Enc_CurrPixelY, Color.FromArgb(Clr.A, R, G, B))
        If Not Enc_NextPixel(True) Then Return False
        Enc_PasswordBits_Curr = 0

        ' Write Checksum File Length
        CurrComponent = 1
        Count = 0

        While Count < FileSizeBin.Length

            If CurrComponent = 4 Then
                Enc_FinalBMP.SetPixel(Enc_CurrPixelX, Enc_CurrPixelY, Color.FromArgb(Clr.A, R, G, B))
                If Not Enc_NextPixel(True) Then Return False
                Clr = Enc_CarrierBMP.GetPixel(Enc_CurrPixelX, Enc_CurrPixelY)
                G = Clr.G
                B = Clr.B
                CurrComponent = 1
            End If

            If CurrComponent = 1 Then
                R = Enc_WriteComponent(Clr.R, Integer.Parse(FileSizeBin(Count)), 8)
            ElseIf CurrComponent = 2 Then
                G = Enc_WriteComponent(Clr.G, Integer.Parse(FileSizeBin(Count)), 8)
            Else
                B = Enc_WriteComponent(Clr.B, Integer.Parse(FileSizeBin(Count)), 8)
            End If

            CurrComponent += 1
            Count += 1
        End While

        Enc_FinalBMP.SetPixel(Enc_CurrPixelX, Enc_CurrPixelY, Color.FromArgb(Clr.A, R, G, B))
        If Not Enc_NextPixel(True) Then Return False
        Enc_PasswordBits_Curr = Enc_PasswordBits_Last

        Return True

    End Function

    Private Function Enc_SaveCarrier(ByVal SavePath As String) As Boolean

        If (Enc_CarrierPath_Curr = -1) Then Return True
        Try
            If Trim(Mid(SavePath, SavePath.Length, 1)) <> "\" Then SavePath = SavePath + "\"
            Enc_FinalBMP.Save(SavePath + GetFileName(Enc_SortedCarrierPaths(Enc_CarrierPath_Curr)) + ".png")
            Enc_FinalBMP.Dispose()
            Enc_CarrierBMP.Dispose()
        Catch ex As Exception
            Return False
        End Try
        Return True

    End Function

    Private Function Enc_NextCarrier(ByVal SavePath As String) As Boolean

        If Not Enc_SaveCarrier(SavePath) Then Return False
        Enc_CarrierPath_Curr += 1
        Enc_PasswordBits_Last = Enc_PasswordBits_Curr
        Enc_PasswordBits_Curr = 0

        Try

            Enc_CarrierBMP = New Bitmap(Enc_SortedCarrierPaths(Enc_CarrierPath_Curr))

            Dim PxFormat As System.Drawing.Imaging.PixelFormat
            If Enc_ComponentsPerPixel = 3 Then
                PxFormat = Imaging.PixelFormat.Format24bppRgb
            Else
                PxFormat = Imaging.PixelFormat.Format32bppArgb
            End If
            Enc_FinalBMP = New Bitmap(Enc_CarrierBMP.Width, Enc_CarrierBMP.Height, PxFormat)

            Enc_CurrPixelX = 0
            Enc_CurrPixelY = 0
            If Not Enc_WriteHeader() Then Return False
            If Enc_ComponentsPerPixel = 3 Then
                Enc_ComponentCurr = 1
            Else
                Enc_ComponentCurr = 0
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Encryption Failed")
            Return False
        End Try
        
        Return True

    End Function

#End Region

End Module
