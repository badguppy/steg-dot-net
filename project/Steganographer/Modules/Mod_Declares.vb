Module Mod_Declares

    Public OverrideXOR As Boolean = False

#Region "User Defined Helper Functions"

    <DebuggerStepThrough()> Public Function DecimalToBinary(ByVal Value As Long, Optional ByVal ConstDigits As Boolean = True, Optional ByVal Digits As Integer = 8) As String

        Dim Str As String = ""

        While Value > 1
            Str = (Value Mod 2).ToString + Str
            Value = Value \ 2
        End While

        If Value = 1 Then Str = Value.ToString + Str
        If ConstDigits Then
            While Str.Length < Digits
                Str = "0" + Str
            End While
        End If

        Return Str

    End Function

    <DebuggerStepThrough()> Public Function DecimalToHex(ByVal Value As Long) As String()

        If Value > 255 Then Throw New Exception("DecimalToHex() IS Not Supposed To Take Inputs Greater Than 255")

        Dim Str() As String = {"0", "0"}

        If Value < 16 Then
            Str(1) = Value.ToString
            Return Str
        Else
            Str(1) = (Value Mod 16).ToString
            Value = Value \ 16
            Str(0) = Value.ToString
            Return Str
        End If

    End Function

    <DebuggerStepThrough()> Public Function BinaryToDecimal(ByVal Value As String) As Long

        Dim D As Long = 0
        Dim Count As Integer = 0

        Value = StrReverse(Value)

        While Count < Value.Length
            D = D + (Trim(Mid(Value, Count + 1, 1)) * Math.Pow(2, Count))
            Count += 1
        End While

        Return D

    End Function

    <DebuggerStepThrough()> Public Function HexToDecimal(ByVal Value As String()) As Integer

        Dim D As Integer = 0
        Dim Count As Integer = 0

        Array.Reverse(Value)

        While Count < Value.Length
            D = D + Value(Count) * Math.Pow(16, Count)
            Count += 1
        End While

        Return D

    End Function

    <DebuggerStepThrough()> Public Function ByteArrToString(ByVal ByteArr() As Byte, ByVal Total As Long) As Char()

        Dim CharArr(Total - 1) As Char
        Dim Count As Integer = 0
        While Count < Total
            CharArr(Count) = Chr(ByteArr(Count))
            Count = Count + 1
        End While

        Return CharArr

    End Function

    <DebuggerStepThrough()> Public Function StringToByteArr(ByVal CharArr() As Char, ByVal Total As Long) As Byte()

        Dim ByteArr(Total - 1) As Byte
        Dim Count As Integer = 0
        While Count < Total
            ByteArr(Count) = Asc(CharArr(Count))
            Count = Count + 1
        End While

        Return ByteArr

    End Function

    Public Function GetFileName(ByVal Path As String) As String
        Return Trim(Mid(Path, Path.LastIndexOf("\") + 2))
    End Function

#End Region

End Module
