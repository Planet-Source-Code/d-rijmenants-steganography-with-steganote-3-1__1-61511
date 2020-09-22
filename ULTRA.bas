Attribute VB_Name = "ULTRA"
'-----------------------------------------------------------------
'
'            ULTRA v1.0.3 Cryptgraphic Algorithm (c) 2004
'
'               Developed and written by D. Rijmenants
'
'-----------------------------------------------------------------
'
' COPYRIGHT NOTICE
'
' ULTRA v1.0.3 is free code and can be used and distributed under
' the following restrictions: It is forbidden to use this code for
' commercial purpose, sell, lease or make profit of copies or parts
' of this code, or make use of this code for other means than legal.
' It is not allowed to use the name ULTRA if changes are made to the
' code, resulting in poor encryption, different data processing or
' an other encryption format than the original ULTRA code.
'
' When using this code in your program, give credit to the author by
' noting 'ULTRA v1.0.3 by D. Rijmenants' in the About and Help sections
' of your program and in your program code, and notify the author at
' mailadress: DR.Defcom@telenet.be
'
' This code may only be used when agreeding these conditions.
'
' THIS CODE IS SUPPLIED "AS IS" AND WITHOUT WARRANTIES OF ANY KIND,
' EITHER EXPRESSED OR IMPLIED, WITH RESPECT TO THIS CODE, ITS QUALITY,
' PERFORMANCE, OR FITNESS FOR ANY PARTICULAR PURPOSE. THE ENTIRE RISK
' AS TO ITS QUALITY AND PERFORMANCE IS WITH THE USER. IN NO EVENT WILL
' THE AUTOR BE LIABLE FOR ANY DIRECT, INDIRECT, OR CONSEQUENTIAL DAMAGES
' RESULTING OUT OF THE USE OF OR INABILITY TO USE THIS CODE.

' IMPORTANT NOTICE.
' ABOUT RESTRICTIONS ON IMPORTING STRONG ENCRYPTION ALGORITHMS:
' ULTRA IS A STREAM CIPHER ENCRYPTION ALGORITHM THAT USES A LENGTH-
' VARIABLE KEY. IN SOME COUNTRIES IMPORT OF THIS TYPE OF SOFTWARE IS
' FORBIDDEN BY LAW OR HAS LEGAL RESTRICTIONS.
' CHECK FOR LEGAL RESTRICTIONS ON THIS SUBJECT IN YOUR COUNTRY.
'
' (c) Rijmenants 2004
'
'-----------------------------------------------------------------

Private K1(0 To 462)  As Integer
Private S1            As Integer
Private P1            As Integer

Private K2(0 To 250)  As Integer
Private P2            As Integer
Private S2            As Integer

Private K3(0 To 180)  As Integer
Private S3            As Integer
Private P3            As Integer

Private FEEDBACK      As Byte
Private SeedString    As String
Private SteganoByte   As Byte
Public UltraReturnValue As Integer

Private Const PR1 = 463
Private Const PR2 = 251
Private Const PR3 = 181

Private aDecTab(255) As Integer
Private aEncTab(63)  As Byte

Private Type HUFFMANTREE
  ParentNode As Integer
  RightNode As Integer
  LeftNode As Integer
  Value As Integer
  Weight As Long
End Type

Private Type byteArray
  Count As Byte
  Data() As Byte
End Type

Private Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

' ------------------------------------------------------------
'            ULTRA Encryption algorithm functions
' ------------------------------------------------------------

Public Sub SetKey(ByVal aKey As String)
'transposition key array initialisations
Dim i           As Long
Dim j           As Long
Dim KEYLen      As Long
Dim KEY1()      As Byte
Dim KEY2(16)    As Byte
Dim KEY3(22)    As Byte
Dim KEYPCC()    As Byte
Dim tmp         As Integer
Dim PCCLen      As Integer
' setup key1 - variable
KEYLen = Len(aKey)
KEY1() = StrConv(aKey, vbFromUnicode)
For i = 0 To PR1 - 1
    K1(i) = i
Next
P1 = 0
S1 = 0
For i = 0 To PR1 - 1
    j = (j + K1(i) + KEY1(i Mod KEYLen)) Mod PR1
    tmp = K1(i)
    K1(i) = K1(j)
    K1(j) = tmp
Next
' setup key2 - 136 bits
For i = 0 To PR1 - 1
    KEY2(i Mod 17) = KEY2(i Mod 17) Xor (K1(i) And 255)
Next
For i = 0 To PR2 - 1
    K2(i) = i
Next
P2 = 0
S2 = 0
For i = 0 To PR2 - 1
    j = (j + K2(i) + KEY2(i Mod 17)) Mod PR2
    tmp = K2(i)
    K2(i) = K2(j)
    K2(j) = tmp
Next
' setup key3 - 184 bits
For i = 0 To PR2 - 1
    KEY3(i Mod 23) = KEY3(i Mod 23) Xor (K2(i) And 255)
Next
For i = 0 To PR3 - 1
    K3(i) = i
Next i
S2 = 0
P2 = 0
For i = 0 To PR3 - 1
    j = (j + K3(i) + KEY3(i Mod 23)) Mod PR3
    tmp = K3(i)
    K3(i) = K3(j)
    K3(j) = tmp
Next
S3 = 0
P3 = 0
FEEDBACK = 0
aKey = ""
End Sub

Private Function FnULTRA(FB As Byte) As Byte
'Ultra function
Dim TS As Integer
Dim OUT1 As Byte
Dim OUT2 As Integer
Dim OUT3 As Integer
P1 = (P1 + 1) Mod PR1
S1 = (S1 + K1(P1) + FB) Mod PR1
TS = K1(P1)
K1(P1) = K1(S1)
K1(S1) = TS
OUT1 = K1((K1(P1) + K1(S1)) Mod PR1) Mod 256
P2 = (P2 + 1) Mod PR2
S2 = (S2 + K2(P2) + OUT1) Mod PR2
TS = K2(P2)
K2(P2) = K2(S2)
K2(S2) = TS
OUT2 = K2((K2(P2) + K2(S2)) Mod PR2) Mod 256
P3 = (P3 + 1) Mod PR3
S3 = (S3 + K3(P3) + OUT2) Mod PR3
TS = K3(P3)
K3(P3) = K3(S3)
K3(S3) = TS
OUT3 = K3((K3(P3) + K3(S3)) Mod PR3) Mod 256
FnULTRA = (OUT1 + OUT2 + OUT3) Mod 256
End Function

Private Function EncodeByte(aByte As Byte) As Byte
EncodeByte = aByte Xor FnULTRA(FEEDBACK)
FEEDBACK = EncodeByte
End Function

Private Function DecodeByte(aByte As Byte) As Byte
Dim tmpbyte As Byte
tmpbyte = aByte
DecodeByte = aByte Xor FnULTRA(FEEDBACK)
FEEDBACK = tmpbyte
End Function

' ------------------------------------------------------------
'                     Random Dummy generating
' ------------------------------------------------------------

Private Function RandomDummy() As String
'setup dummy string, between 16 and 255 bytes, first byte contains dummylenght
Dim rndKey As String
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim q As Byte
Dim SizeDummy As Integer
RandomDummy = ""
Randomize
SizeDummy = Int(224 * Rnd) + 32
If Len(SeedString) > 0 Then
    For k = 1 To Len(SeedString)
        SizeDummy = SizeDummy Xor Asc(Mid(SeedString, k, 1))
    Next
    End If
Do While SizeDummy > 255
    SizeDummy = SizeDummy - 224
Loop
If SizeDummy < 32 Then SizeDummy = SizeDummy + 224
For k = 1 To SizeDummy - 1
    RandomDummy = RandomDummy & Chr(Int((256 * Rnd)))
Next
j = 1
For k = 1 To 16
    rndKey = ""
    For i = 1 To 16
        q = Int((256 * Rnd))
        If Len(SeedString) > 0 Then q = q Xor Asc(Mid(SeedString, j, 1))
        j = j + 1: If j > Len(SeedString) Then j = 1
        rndKey = rndKey & Chr(q)
    Next i
    Call SetKey(rndKey)
    For i = 1 To Len(RandomDummy)
        q = Asc(Mid(RandomDummy, i, 1))
        If k Mod 3 = 0 Then
            q = DecodeByte(q)
            Else
            q = EncodeByte(q)
            End If
        Mid(RandomDummy, i, 1) = Chr(q)
    Next i
Next k
RandomDummy = Chr(SizeDummy) & RandomDummy
End Function

Public Sub RandomFeed(ByVal x As Single, ByVal y As Single)
'this sub enables the user to feed random data to seedstring
Static Xp As Single
Static Yp As Single
If x = Xp And y = Yp Then Exit Sub
Xp = x: Yp = y
SeedString = SeedString & Chr((x Xor y) And 255)
If Len(SeedString) > 251 Then SeedString = Mid(SeedString, 2): Exit Sub
' show rnd in statusbar
frmMain.lblRND.Caption = "RND " & Trim(Int(Len(SeedString) / 251 * 100)) & "%"
End Sub

' ------------------------------------------------------------
'                   Compression functions
' ------------------------------------------------------------

Private Function HuffDecodeString(Text As String) As String
Dim byteArray() As Byte
byteArray() = StrConv(Text, vbFromUnicode)
Call HuffDecodeByte(byteArray, Len(Text))
HuffDecodeString = StrConv(byteArray(), vbUnicode)
End Function

Private Function HuffEncodeString(Text As String) As String
Dim byteArray() As Byte
byteArray() = StrConv(Text, vbFromUnicode)
Call HuffEncodeByte(byteArray, Len(Text))
HuffEncodeString = StrConv(byteArray(), vbUnicode)
End Function

Private Sub HuffEncodeByte(byteArray() As Byte, ByteLen As Long)
Dim i As Long, j As Long, Char As Byte, BitPos As Byte, lNode1 As Long
Dim lNode2 As Long, lNodes As Long, lLength As Long, Count As Integer
Dim lWeight1 As Long, lWeight2 As Long, Result() As Byte, ByteValue As Byte
Dim ResultLen As Long, bytes As byteArray, NodesCount As Integer, NewProgress As Integer
Dim BitValue(0 To 7) As Byte, CharCount(0 To 255) As Long
Dim Nodes(0 To 511) As HUFFMANTREE, CharValue(0 To 255) As byteArray
'set identification
If (ByteLen = 0) Then
    ReDim Preserve byteArray(0 To ByteLen + 3)
    If (ByteLen > 0) Then Call CopyMem(byteArray(4), byteArray(0), ByteLen)
    byteArray(0) = 72
    byteArray(1) = 69
    byteArray(2) = 48
    byteArray(3) = 13
    Exit Sub
End If
ReDim Result(0 To 522)
Result(0) = 72
Result(1) = 69
Result(2) = 51
Result(3) = 13
ResultLen = 4
'get frequency off all bytes
For i = 0 To (ByteLen - 1)
    CharCount(byteArray(i)) = CharCount(byteArray(i)) + 1
Next
'put freq in nodes
For i = 0 To 255
    If (CharCount(i) > 0) Then
        With Nodes(NodesCount)
            .Weight = CharCount(i)
            .Value = i
            .LeftNode = -1
            .RightNode = -1
            .ParentNode = -1
        End With
        NodesCount = NodesCount + 1
    End If
Next

For lNodes = NodesCount To 2 Step -1
    lNode1 = -1: lNode2 = -1
    For i = 0 To (NodesCount - 1)
        If (Nodes(i).ParentNode = -1) Then
            If (lNode1 = -1) Then
                lWeight1 = Nodes(i).Weight
                lNode1 = i
            ElseIf (lNode2 = -1) Then
                lWeight2 = Nodes(i).Weight
                lNode2 = i
            ElseIf (Nodes(i).Weight < lWeight1) Then
                If (Nodes(i).Weight < lWeight2) Then
                    If (lWeight1 < lWeight2) Then
                        lWeight2 = Nodes(i).Weight
                        lNode2 = i
                    Else
                        lWeight1 = Nodes(i).Weight
                        lNode1 = i
                    End If
                Else
                    lWeight1 = Nodes(i).Weight
                    lNode1 = i
                End If
            ElseIf (Nodes(i).Weight < lWeight2) Then
                lWeight2 = Nodes(i).Weight
                lNode2 = i
            End If
        End If
    Next
    
    With Nodes(NodesCount)
        .Weight = lWeight1 + lWeight2
        .LeftNode = lNode1
        .RightNode = lNode2
        .ParentNode = -1
        .Value = -1
    End With
    
    Nodes(lNode1).ParentNode = NodesCount
    Nodes(lNode2).ParentNode = NodesCount
    NodesCount = NodesCount + 1
Next

ReDim bytes.Data(0 To 255)
Call CreateBitSequences(Nodes(), NodesCount - 1, bytes, CharValue)

For i = 0 To 255
    If (CharCount(i) > 0) Then lLength = lLength + CharValue(i).Count * CharCount(i)
Next
lLength = IIf(lLength Mod 8 = 0, lLength \ 8, lLength \ 8 + 1)
  
If ((lLength = 0) Or (lLength > ByteLen)) Then
    ReDim Preserve byteArray(0 To ByteLen + 3)
    Call CopyMem(byteArray(4), byteArray(0), ByteLen)
    byteArray(0) = 72
    byteArray(1) = 69
    byteArray(2) = 48
    byteArray(3) = 13
    Exit Sub
End If
'calculate CRC
Char = 0
For i = 0 To (ByteLen - 1)
    Char = Char Xor byteArray(i)
Next
Result(ResultLen) = Char
ResultLen = ResultLen + 1
Call CopyMem(Result(ResultLen), ByteLen, 4)
ResultLen = ResultLen + 4
BitValue(0) = 2 ^ 0
BitValue(1) = 2 ^ 1
BitValue(2) = 2 ^ 2
BitValue(3) = 2 ^ 3
BitValue(4) = 2 ^ 4
BitValue(5) = 2 ^ 5
BitValue(6) = 2 ^ 6
BitValue(7) = 2 ^ 7
Count = 0
For i = 0 To 255
    If (CharValue(i).Count > 0) Then Count = Count + 1
Next
Call CopyMem(Result(ResultLen), Count, 2)
ResultLen = ResultLen + 2
Count = 0
For i = 0 To 255
    If (CharValue(i).Count > 0) Then
        Result(ResultLen) = i
        ResultLen = ResultLen + 1
        Result(ResultLen) = CharValue(i).Count
        ResultLen = ResultLen + 1
        Count = Count + 16 + CharValue(i).Count
    End If
Next
  
ReDim Preserve Result(0 To ResultLen + Count \ 8)
  
BitPos = 0
ByteValue = 0
For i = 0 To 255
    With CharValue(i)
        If (.Count > 0) Then
            For j = 0 To (.Count - 1)
                If (.Data(j)) Then ByteValue = ByteValue + BitValue(BitPos)
                BitPos = BitPos + 1
                If (BitPos = 8) Then
                    Result(ResultLen) = ByteValue
                    ResultLen = ResultLen + 1
                    ByteValue = 0
                    BitPos = 0
                End If
            Next
        End If
    End With
Next
If (BitPos > 0) Then
    Result(ResultLen) = ByteValue
    ResultLen = ResultLen + 1
End If
  
ReDim Preserve Result(0 To ResultLen - 1 + lLength)
  
Char = 0
BitPos = 0
For i = 0 To (ByteLen - 1)
    With CharValue(byteArray(i))
        For j = 0 To (.Count - 1)
            If (.Data(j) = 1) Then Char = Char + BitValue(BitPos)
            BitPos = BitPos + 1
            If (BitPos = 8) Then
                Result(ResultLen) = Char
                ResultLen = ResultLen + 1
                BitPos = 0
                Char = 0
            End If
        Next
    End With
Next

If (BitPos > 0) Then
    Result(ResultLen) = Char
    ResultLen = ResultLen + 1
End If
ReDim byteArray(0 To ResultLen - 1)
Call CopyMem(byteArray(0), Result(0), ResultLen)
End Sub

Private Sub HuffDecodeByte(byteArray() As Byte, ByteLen As Long)
Dim i As Long, j As Long, Pos As Long, Char As Byte, CurrPos As Long
Dim Count As Integer, CheckSum As Byte, Result() As Byte, BitPos As Integer
Dim NodeIndex As Long, ByteValue As Byte, ResultLen As Long, NodesCount As Long
Dim lResultLen As Long, NewProgress As Integer, BitValue(0 To 7) As Byte
Dim Nodes(0 To 511) As HUFFMANTREE, CharValue(0 To 255) As byteArray
    
If (byteArray(0) <> 72) Or (byteArray(1) <> 69) Or (byteArray(3) <> 13) Then
ElseIf (byteArray(2) = 48) Then
    Call CopyMem(byteArray(0), byteArray(4), ByteLen - 4)
    ReDim Preserve byteArray(0 To ByteLen - 5)
    Exit Sub
ElseIf (byteArray(2) <> 51) Then
    Err.Raise vbObjectError, "HuffmanDecode()", "The data either was not compressed with HE3 or is corrupt (identification string not found)"
    Exit Sub
End If

CurrPos = 5
CheckSum = byteArray(CurrPos - 1)
CurrPos = CurrPos + 1

Call CopyMem(ResultLen, byteArray(CurrPos - 1), 4)
CurrPos = CurrPos + 4
lResultLen = ResultLen
If (ResultLen = 0) Then Exit Sub
ReDim Result(0 To ResultLen - 1)
Call CopyMem(Count, byteArray(CurrPos - 1), 2)
CurrPos = CurrPos + 2

For i = 1 To Count
    With CharValue(byteArray(CurrPos - 1))
        CurrPos = CurrPos + 1
        .Count = byteArray(CurrPos - 1)
        CurrPos = CurrPos + 1
        ReDim .Data(0 To .Count - 1)
    End With
Next

BitValue(0) = 2 ^ 0
BitValue(1) = 2 ^ 1
BitValue(2) = 2 ^ 2
BitValue(3) = 2 ^ 3
BitValue(4) = 2 ^ 4
BitValue(5) = 2 ^ 5
BitValue(6) = 2 ^ 6
BitValue(7) = 2 ^ 7

ByteValue = byteArray(CurrPos - 1)
CurrPos = CurrPos + 1
BitPos = 0

For i = 0 To 255
    With CharValue(i)
        If (.Count > 0) Then
            For j = 0 To (.Count - 1)
                If (ByteValue And BitValue(BitPos)) Then .Data(j) = 1
                BitPos = BitPos + 1
                If (BitPos = 8) Then
                    ByteValue = byteArray(CurrPos - 1)
                    CurrPos = CurrPos + 1
                    BitPos = 0
                End If
            Next
        End If
    End With
Next

If (BitPos = 0) Then CurrPos = CurrPos - 1
 
NodesCount = 1
Nodes(0).LeftNode = -1
Nodes(0).RightNode = -1
Nodes(0).ParentNode = -1
Nodes(0).Value = -1

For i = 0 To 255
    Call CreateTree(Nodes(), NodesCount, i, CharValue(i))
Next

ResultLen = 0
    For CurrPos = CurrPos To ByteLen
        ByteValue = byteArray(CurrPos - 1)
        For BitPos = 0 To 7
            If (ByteValue And BitValue(BitPos)) Then NodeIndex = Nodes(NodeIndex).RightNode Else NodeIndex = Nodes(NodeIndex).LeftNode
            If (Nodes(NodeIndex).Value > -1) Then
                Result(ResultLen) = Nodes(NodeIndex).Value
                ResultLen = ResultLen + 1
                If (ResultLen = lResultLen) Then GoTo DecodeFinished
                NodeIndex = 0
            End If
        Next
    Next

DecodeFinished:
    'check CRC
    Char = 0
    For i = 0 To (ResultLen - 1)
        Char = Char Xor Result(i)
    Next
    If (Char <> CheckSum) Then UltraReturnValue = 5
    ReDim byteArray(0 To ResultLen - 1)
    Call CopyMem(byteArray(0), Result(0), ResultLen)
End Sub

Private Sub CreateBitSequences(Nodes() As HUFFMANTREE, ByVal NodeIndex As Integer, bytes As byteArray, CharValue() As byteArray)
    Dim NewBytes As byteArray
    If (Nodes(NodeIndex).Value > -1) Then
        CharValue(Nodes(NodeIndex).Value) = bytes
        Exit Sub
    End If
    If (Nodes(NodeIndex).LeftNode > -1) Then
        NewBytes = bytes
        NewBytes.Data(NewBytes.Count) = 0
        NewBytes.Count = NewBytes.Count + 1
        Call CreateBitSequences(Nodes(), Nodes(NodeIndex).LeftNode, NewBytes, CharValue)
    End If
    If (Nodes(NodeIndex).RightNode > -1) Then
        NewBytes = bytes
        NewBytes.Data(NewBytes.Count) = 1
        NewBytes.Count = NewBytes.Count + 1
        Call CreateBitSequences(Nodes(), Nodes(NodeIndex).RightNode, NewBytes, CharValue)
    End If
End Sub

Private Sub CreateTree(Nodes() As HUFFMANTREE, NodesCount As Long, Char As Long, bytes As byteArray)
Dim a As Integer
Dim NodeIndex As Long
NodeIndex = 0
For a = 0 To (bytes.Count - 1)
    If (bytes.Data(a) = 0) Then
        If (Nodes(NodeIndex).LeftNode = -1) Then
            Nodes(NodeIndex).LeftNode = NodesCount
            Nodes(NodesCount).ParentNode = NodeIndex
            Nodes(NodesCount).LeftNode = -1
            Nodes(NodesCount).RightNode = -1
            Nodes(NodesCount).Value = -1
            NodesCount = NodesCount + 1
        End If
        NodeIndex = Nodes(NodeIndex).LeftNode
    ElseIf (bytes.Data(a) = 1) Then
        If (Nodes(NodeIndex).RightNode = -1) Then
            Nodes(NodeIndex).RightNode = NodesCount
            Nodes(NodesCount).ParentNode = NodeIndex
            Nodes(NodesCount).LeftNode = -1
            Nodes(NodesCount).RightNode = -1
            Nodes(NodesCount).Value = -1
            NodesCount = NodesCount + 1
        End If
        NodeIndex = Nodes(NodeIndex).RightNode
    Else
        Stop
    End If
Next
Nodes(NodeIndex).Value = Char
End Sub

' ------------------------------------------------------------
'                   Base 64 Radix functions
' ------------------------------------------------------------

Private Function PadString(strData As String) As String
' Pad data string to next multiple of 8 bytes
Dim nLen As Long
Dim sPad As String
Dim nPad As Integer
nLen = Len(strData)
nPad = ((nLen \ 8) + 1) * 8 - nLen
sPad = String(nPad, Chr(nPad))
PadString = strData & sPad
End Function

Private Function UnpadString(strData As String) As String
' Strip padding
Dim nLen As Long
Dim nPad As Long
nLen = Len(strData)
If nLen = 0 Then Exit Function
nPad = Asc(Right(strData, 1))
If nPad > 8 Then nPad = 0
UnpadString = Left(strData, nLen - nPad)
End Function

Private Function EncodeStr64(encString As String, ByVal MaxPerLine As Integer) As String
' Return radix64 encoding of string of binary values
Dim abOutput()  As Byte
Dim sLast       As String
Dim B(3)        As Byte
Dim j           As Integer
Dim CharCount   As Integer
Dim iIndex      As Long
Dim Umax        As Long
Dim i As Long, nLen As Long, nQuants As Long
EncodeStr64 = ""
nLen = Len(encString)
nQuants = nLen \ 3
iIndex = 0
If MaxPerLine < 10 Then MaxPerLine = 10
Umax = nQuants + 1
Call MakeEncTab
If (nQuants > 0) Then
    ReDim abOutput(nQuants * 4 - 1)
    For i = 0 To nQuants - 1
        For j = 0 To 2
            B(j) = Asc(Mid(encString, (i * 3) + j + 1, 1))
        Next
        Call EncodeQuantumB(B)
        abOutput(iIndex) = B(0)
        abOutput(iIndex + 1) = B(1)
        abOutput(iIndex + 2) = B(2)
        abOutput(iIndex + 3) = B(3)
        CharCount = CharCount + 4
        ' insert CRLF if max char per line is reached
        If CharCount >= MaxPerLine Then
            ReDim Preserve abOutput(UBound(abOutput) + 2)
            CharCount = 0
            abOutput(iIndex + 4) = 13
            abOutput(iIndex + 5) = 10
            iIndex = iIndex + 6
            Else
            iIndex = iIndex + 4
            End If
    Next
    EncodeStr64 = StrConv(abOutput, vbUnicode)
End If
Select Case nLen Mod 3
Case 0
    sLast = ""
Case 1
    B(0) = Asc(Mid(encString, nLen, 1))
    B(1) = 0
    B(2) = 0
    Call EncodeQuantumB(B)
    sLast = StrConv(B(), vbUnicode)
    ' Replace last 2 with =
    sLast = Left(sLast, 2) & "=="
Case 2
    B(0) = Asc(Mid(encString, nLen - 1, 1))
    B(1) = Asc(Mid(encString, nLen, 1))
    B(2) = 0
    Call EncodeQuantumB(B)
    sLast = StrConv(B(), vbUnicode)
    ' Replace last with =
    sLast = Left(sLast, 3) & "="
End Select
EncodeStr64 = EncodeStr64 & sLast
End Function

Private Function DecodeStr64(decString As String) As String
' Return string of decoded binary values from radix64 string
Dim abDecoded() As Byte
Dim d(3)    As Byte
Dim c       As Integer
Dim di      As Integer
Dim i       As Long
Dim nLen    As Long
Dim iIndex  As Long
Dim Umax    As Long
nLen = Len(decString)
If nLen < 4 Then
    Exit Function
End If
ReDim abDecoded(((nLen \ 4) * 3) - 1)
Umax = nLen
iIndex = 0
di = 0
Call MakeDecTab
For i = 1 To Len(decString)
    c = CByte(Asc(Mid(decString, i, 1)))
    c = aDecTab(c)
    If c >= 0 Then
        d(di) = CByte(c)
        di = di + 1
        If di = 4 Then
            abDecoded(iIndex) = SHL2(d(0)) Or (SHR4(d(1)) And &H3)
            iIndex = iIndex + 1
            abDecoded(iIndex) = SHL4(d(1) And &HF) Or (SHR2(d(2)) And &HF)
            iIndex = iIndex + 1
            abDecoded(iIndex) = SHL6(d(2) And &H3) Or d(3)
            iIndex = iIndex + 1
            If d(3) = 64 Then
                iIndex = iIndex - 1
                abDecoded(iIndex) = 0
            End If
            If d(2) = 64 Then
                iIndex = iIndex - 1
                abDecoded(iIndex) = 0
            End If
            di = 0
        End If
    End If
Next i
DecodeStr64 = StrConv(abDecoded(), vbUnicode)
DecodeStr64 = Left(DecodeStr64, iIndex)
End Function

Private Sub EncodeQuantumB(B() As Byte)
Dim b0 As Byte, B1 As Byte, B2 As Byte, B3 As Byte
b0 = SHR2(B(0)) And &H3F
B1 = SHL4(B(0) And &H3) Or (SHR4(B(1)) And &HF)
B2 = SHL2(B(1) And &HF) Or (SHR6(B(2)) And &H3)
B3 = B(2) And &H3F
B(0) = aEncTab(b0)
B(1) = aEncTab(B1)
B(2) = aEncTab(B2)
B(3) = aEncTab(B3)
End Sub

Private Function MakeDecTab()
' Set up Radix 64 decoding table
Dim t As Integer
Dim c As Integer
For c = 0 To 255
    aDecTab(c) = -1
Next
t = 0
For c = Asc("A") To Asc("Z")
    aDecTab(c) = t
    t = t + 1
Next
For c = Asc("a") To Asc("z")
    aDecTab(c) = t
    t = t + 1
Next
For c = Asc("0") To Asc("9")
    aDecTab(c) = t
    t = t + 1
Next
c = Asc("+")
aDecTab(c) = t
t = t + 1
c = Asc("/")
aDecTab(c) = t
t = t + 1
c = Asc("=")
aDecTab(c) = t
End Function

Private Function MakeEncTab()
' Set up Radix 64 encoding table in bytes
Dim i As Integer
Dim c As Integer
i = 0
For c = Asc("A") To Asc("Z")
    aEncTab(i) = c
    i = i + 1
Next
For c = Asc("a") To Asc("z")
    aEncTab(i) = c
    i = i + 1
Next
For c = Asc("0") To Asc("9")
    aEncTab(i) = c
    i = i + 1
Next
c = Asc("+")
aEncTab(i) = c
i = i + 1
c = Asc("/")
aEncTab(i) = c
i = i + 1
End Function

Private Function SHL2(ByVal bytValue As Byte) As Byte
' Shift 8-bit value to left by 2 bits
SHL2 = (bytValue * &H4) And &HFF
End Function

Private Function SHL4(ByVal bytValue As Byte) As Byte
' Shift 8-bit value to left by 4 bits
SHL4 = (bytValue * &H10) And &HFF
End Function

Private Function SHL6(ByVal bytValue As Byte) As Byte
' Shift 8-bit value to left by 6 bits
SHL6 = (bytValue * &H40) And &HFF
End Function

Private Function SHR2(ByVal bytValue As Byte) As Byte
' Shift 8-bit value to right by 2 bits
SHR2 = bytValue \ &H4
End Function

Private Function SHR4(ByVal bytValue As Byte) As Byte
' Shift 8-bit value to right by 4 bits
SHR4 = bytValue \ &H10
End Function

Private Function SHR6(ByVal bytValue As Byte) As Byte
' Shift 8-bit value to right by 6 bits
SHR6 = bytValue \ &H40
End Function

' ------------------------------------------------------------
'                   Steganography functions
' ------------------------------------------------------------

Public Sub SetSteganoKey(ByVal KeyString As String)
Call SetKey(KeyString)
SteganoByte = 0
End Sub

Public Function SetSteganoText(TextIn As String) As String
Dim DummyString As String
Dim checkByte1 As Byte
Dim checkByte2 As Byte
If TextIn = "" Then Exit Function
SetSteganoText = HuffEncodeString(TextIn)
'create dummy header
DummyString = RandomDummy
checkByte1 = Asc(Mid(DummyString, Len(DummyString) - 1, 1))
checkByte2 = Asc(Mid(DummyString, Len(DummyString), 1))
'add dummy and check bytes
SetSteganoText = DummyString & SetSteganoText & Chr(checkByte1) & Chr(checkByte2)
End Function

Public Function GetSteganoText(TextIn As String) As String
Dim DummyString As String
Dim SizeDummy As Integer
Dim checkByte1 As Byte
Dim checkByte2 As Byte
UltraReturnValue = 0
SizeDummy = Asc(Left(TextIn, 1))
If SizeDummy > Len(TextIn) - 2 Then GoTo errDecode
'get dummy string
DummyString = Left(TextIn, SizeDummy)
'get checkbytes
checkByte1 = Asc(Mid(DummyString, Len(DummyString) - 1, 1))
checkByte2 = Asc(Mid(DummyString, Len(DummyString), 1))
'check decryption
If Asc(Mid(TextIn, Len(TextIn) - 1, 1)) = checkByte1 And _
 Asc(Mid(TextIn, Len(TextIn), 1)) = checkByte2 Then
    GetSteganoText = Mid(TextIn, SizeDummy + 1, (Len(TextIn) - 2) - SizeDummy)
    GetSteganoText = HuffDecodeString(GetSteganoText)
    Else
    GoTo errDecode
    End If
Exit Function
errDecode:
GetSteganoText = ""
UltraReturnValue = 40
End Function

Public Function GetSteganoPix() As Byte
'get next value from fnUltra
GetSteganoPix = FnULTRA(SteganoByte)
FEEDBACK = GetSteganoPix
End Function

Public Sub SetSteganoByte(ByVal aByte As Byte)
'set the feedback byte for fnUltra input
SteganoByte = aByte
End Sub

