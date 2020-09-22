Attribute VB_Name = "modStegano"
Option Explicit

Public DataPicX
Public DataPicY

Dim Xp As Long
Dim Yp As Long
Dim Cp As Integer
Dim Pixel As Long

Private Xmax As Long
Private Ymax As Long
Private NmbrPix As Long

Private R As Byte
Private G As Byte
Private B As Byte

Private UsedPix() As Byte

Dim PicOverflow As Boolean

Public Sub SaveDataPicture()
'save the data into image
Dim strData As String
Dim B1 As Double
Dim B2 As Double
Dim B3 As Double
Dim k As Double
Dim sLen As Double
Dim sFilename As String
Dim blnErase As Boolean
Dim convertToBmp As Boolean
Dim killName As String
Dim FileBuffer() As Byte
Dim fileO As Integer

Cp = 0

If GetFileExt(gstrImageName) <> "bmp" Then
    retVal = MsgBox("The data will be saved as bitmap named """ & CutFileExt(CutFilePath(gstrImageName)) & ".bmp""." & vbCrLfLf & "Dou you want to delete the original file """ & CutFilePath(gstrImageName) & """?", vbQuestion + vbYesNoCancel)
    If retVal = vbCancel Then Exit Sub
    If retVal = vbYes Then blnErase = True
    End If

'set image
frmMain.picStegano.AutoSize = True
frmMain.picStegano.Picture = LoadPicture(gstrImageName)
frmMain.picStegano.Refresh
frmMain.picStegano.AutoSize = False

'get size of image
DataPicX = frmMain.picStegano.ScaleWidth
DataPicY = frmMain.picStegano.ScaleHeight
Xmax = DataPicX - 1
Ymax = DataPicY - 1

'set savename
If GetFileExt(gstrImageName) <> "bmp" Then
    convertToBmp = True
    killName = gstrImageName
    gstrImageName = CutFileExt(gstrImageName) & ".bmp"
    'check for overwrite
    If FileExist(gstrImageName) Then
        retVal = MsgBox("""" & CutFilePath(gstrImageName) & """ already exists. StegaNote will overwrite this file.", vbInformation + vbOKCancel)
        If retVal = vbCancel Then Exit Sub
    End If
End If
    
    
PicOverflow = False
Randomize


On Error GoTo errHandleFile
'get text and file, to strData and compress
Screen.MousePointer = 11
If frmMain.txtMain.Text <> "" Then
    'text to strData
    strData = "TEXT" & Chr(0) & frmMain.txtMain.Text
    Else
    ' no text, add chr(0)
    strData = "TEXT" & Chr(0) & Chr(0)
    End If
If gstrFileName <> "" Then
    If FileLen(gstrFileName) > 0 Then
        'open file and read bytes into buffer array
        fileO = FreeFile
        Open gstrFileName For Binary As #fileO
            ReDim FileBuffer(0 To LOF(fileO) - 1)
            Get #fileO, , FileBuffer()
        Close #fileO
        strData = strData & "FILE" & Chr(0) & CutFilePath(gstrFileName) & Chr(0)
        strData = strData & StrConv(FileBuffer(), vbUnicode)
    Else
        If FileLen(gstrFileName) = 0 Then MsgBox """" & CutFilePath(gstrFileName) & """ doesn't contain any information and will not be inserted.", vbInformation
        'no file, add chr(0)
        strData = strData & "FILE" & Chr(0) & Chr(0)
    End If
Else
    strData = strData & "FILE" & Chr(0) & Chr(0)
End If

'compress and add header
strData = SetSteganoText(strData)
sLen = Len(strData)
Screen.MousePointer = 0

'set key
Call SetSteganoKey(gstrActiveKey)

'setup array for used pixs
NmbrPix = (DataPicX * DataPicY) - 1
ReDim UsedPix(NmbrPix)

'check data size to fit in image
If (sLen + 2) * 8 > (DataPicX * DataPicY) * 3 Then
    MsgBox "Image too small to store all data." & vbCrLf & "Please select a larger image or decrease the amount of data to save.", vbCritical
    Exit Sub
    End If

'get first pixelset
Call GetNextPixel

    
'set cancel flag
gblnCancelProcces = False

'set text lenght
B1 = sLen And 255
B2 = (sLen And 65280) / 256
B3 = (sLen And 16711680) / 65536

'write data lengt to image
WritePixel (B1)
WritePixel (B2)
WritePixel (B3)

ProgressShow frmMain.picProgBar, 0

'start writing all bits to image
For k = 1 To sLen
    WritePixel (Asc(Mid(strData, k, 1)))
    If PicOverflow = True Then
        MsgBox "Image too small to store all data." & vbCrLfLf & "Data merge overflow.", vbCritical ' "Data merge overflow.":
        Exit Sub
    End If
    ProgressShow frmMain.picProgBar, k / sLen
    If k Mod 10 = 0 Then DoEvents
    If gblnCancelProcces = True Then Exit For
Next k
'set last pixel RGB
frmMain.picStegano.PSet (Xp, Yp), RGB(R, G, B)

'refresh image
frmMain.picStegano.Picture = frmMain.picStegano.Image

'abort message
If gblnCancelProcces = True Then
    MsgBox "Saving data in image aborted by user.", vbInformation '"Merging data and image aborted."
    ProgressShow frmMain.picProgBar, 0
    Exit Sub
    End If

On Error GoTo errHandle

Screen.MousePointer = 11
If gstrImageName <> "" Then
    SavePicture frmMain.picStegano.Picture, gstrImageName
    gblnTextHasChanged = False
End If
Screen.MousePointer = 0

If GetFileExt(gstrImageName) <> "bmp" Then
    convertToBmp = True
    killName = gstrImageName
    gstrImageName = CutFileExt(gstrImageName) & ".bmp"
    End If
    
ProgressShow frmMain.picProgBar, 0
    
'if no errors, original deleted if not bmp
On Error GoTo errHandleKill
If convertToBmp = True And blnErase = True Then Kill killName
gblnTextHasChanged = False
Exit Sub

errHandle:
Screen.MousePointer = 0
MsgBox "Failed saving " & CutFilePath(gstrImageName) & vbCrLfLf & Err.Description, vbCritical
ProgressShow frmMain.picProgBar, 0
Exit Sub

errHandleKill:
Screen.MousePointer = 0
MsgBox "Failed deleting sourcefile " & CutFilePath(killName) & vbCrLf & vbCrLf & Err.Description, vbCritical  'Failed deleting:
ProgressShow frmMain.picProgBar, 0

Exit Sub

errHandleFile:
Screen.MousePointer = 0
MsgBox "Failed reading data from " & CutFilePath(gstrFileName) & vbCrLf & vbCrLf & Err.Description, vbCritical
ProgressShow frmMain.picProgBar, 0

End Sub

Public Sub LoadDataPicture()
'extract data from image
Dim strData As String
Dim B1 As Double
Dim B2 As Double
Dim B3 As Double
Dim sLen As Double
Dim tmp As String
Dim fByte As Byte
Dim tmpFileName As String
Dim tmpfolder As String
Dim FileBuffer() As Byte
Dim fileO As Integer
Dim Pos As Long
On Error GoTo errHandle
Dim k As Single
Screen.MousePointer = 11
'load image and its size

Screen.MousePointer = 0

'set image
Screen.MousePointer = 11
frmMain.picStegano.AutoSize = True
frmMain.picStegano.Picture = LoadPicture(gstrImageName)
frmMain.picStegano.Refresh
frmMain.picStegano.AutoSize = False
Screen.MousePointer = 0

'get size of image
DataPicX = frmMain.picStegano.ScaleWidth
DataPicY = frmMain.picStegano.ScaleHeight

Xmax = DataPicX - 1
Ymax = DataPicY - 1
Cp = 0

Call SetSteganoKey(gstrActiveKey)
'setup array for used pixs
NmbrPix = (DataPicX * DataPicY) - 1
ReDim UsedPix(NmbrPix)
'get first pixelset
Call GetNextPixel

'get number of bytes
B1 = ReadPixel
B2 = ReadPixel
B3 = ReadPixel

sLen = B1 + (B2 * 256) + (B3 * 65536)


'check image bitsize
If ((sLen + 2) * 8) > (DataPicX * DataPicY) * 3 Or sLen = 0 Then
    'corrupted data lenght
    MsgBox "Failed reading the data from " & CutFilePath(gstrImageName) & vbCrLfLf & "This may be caused by the following:" & vbCrLfLf & "- Wrong passphraze or passphraze contains errors." & vbCrLf & "- The data in the image contains errors or is corrupted." & vbCrLf & "- The image doesn't contain any data.", vbCritical
    Exit Sub
    End If

gblnCancelProcces = False
ProgressShow frmMain.picProgBar, 0
PicOverflow = False
For k = 1 To sLen
    fByte = ReadPixel
    If PicOverflow = True Then
        MsgBox "Failed reading the data from " & CutFilePath(gstrImageName) & vbCrLfLf & "This may be caused by the following:" & vbCrLfLf & "- Wrong passphraze or passphraze contains errors." & vbCrLf & "- The data in the image contains errors or is corrupted." & vbCrLf & "- The image doesn't contain any data.", vbCritical
        gstrActiveKey = ""
        Exit Sub
        End If
    strData = strData & Chr(fByte)
    ProgressShow frmMain.picProgBar, k / sLen
    DoEvents
    If gblnCancelProcces = True Then Exit For
Next k

If gblnCancelProcces = True Then
    'Extracting data from image aborted.
    MsgBox "Reading data from image aborted by user.", vbInformation
    Exit Sub
    End If

ProgressShow frmMain.picProgBar, 0
Screen.MousePointer = 11

'check header and decompres
strData = GetSteganoText(strData)

Screen.MousePointer = 0

On Error GoTo errHandleSaveFile

If UltraReturnValue = 0 Then
    'get text from strData
    strData = Mid(strData, 6)
    Pos = InStr(1, strData, "FILE" & Chr(0))
    
    If Pos > 1 Then
        frmMain.txtMain.Text = Left(strData, Pos - 1)
        Else
        frmMain.txtMain.Text = ""
    End If
    gblnTextHasChanged = False
    'get file from strData
    strData = Mid(strData, Pos + 5)
    If Len(strData) > 1 Then
        Pos = InStr(1, strData, Chr(0))
        tmpFileName = Left(strData, Pos - 1)
        strData = Mid(strData, Pos + 1)
        FileBuffer() = StrConv(strData, vbFromUnicode)
        'get save directory
askAgain:
        retVal = MsgBox("Save """ & tmpFileName & """ in the same folder as the carrier image?", vbQuestion + vbYesNoCancel)
        If retVal = vbYes Then
            'save file in carrier folder
            gstrFileName = GetFilePath(gstrImageName) & tmpFileName
        ElseIf retVal = vbNo Then
            'show folder brows dialog
            tmpfolder = Browse("Select folder to save " & tmpFileName)
            If tmpfolder <> "" Then
                gstrFileName = tmpfolder & tmpFileName
                Else
                'gstrFileName = GetFilePath(gstrImageName) & tmpFileName
                Exit Sub
            End If
        Else
            Exit Sub
        End If
        'save the file
        fileO = FreeFile
        Screen.MousePointer = 11
        Open gstrFileName For Binary As #fileO
            Put #fileO, , FileBuffer()
        Close #fileO
        Screen.MousePointer = 0
        Else
        gstrFileName = ""
    End If
    Call GetFile
Else
    MsgBox "Failed reading the data from " & CutFilePath(gstrImageName) & vbCrLfLf & "This may be caused by the following:" & vbCrLfLf & "- Wrong passphraze or passphraze contains errors." & vbCrLf & "- The data in the image contains errors or is corrupted." & vbCrLf & "- The image doesn't contain any data.", vbCritical
    Screen.MousePointer = 0
    gstrFileName = ""
    frmMain.txtMain.Text = ""
    Call GetFile
End If

Exit Sub

errHandle:
Screen.MousePointer = 0
'"Failed loading image: "
MsgBox "Failed loading the source image." & vbCrLf & vbCrLf & Err.Description, vbCritical

errHandleSaveFile:
Screen.MousePointer = 0
'"Failed loading image: "
MsgBox "Failed saving the extracted file." & vbCrLfLf & Err.Description, vbCritical
gstrFileName = ""
End Sub

Public Sub EraseImage()
Dim blnErase As Boolean
Dim convertToBmp As Boolean
Dim killName As String

If GetFileExt(gstrImageName) <> "bmp" Then
    MsgBox "Cannot erase data from """ & CutFilePath(gstrImageName) & """" & vbCrLfLf & "This file is no data carrier image.", vbCritical
    Exit Sub
End If

gblnCancelProcces = False
ProgressShow frmMain.picProgBar, 0
'set image
frmMain.picStegano.AutoSize = True
frmMain.picStegano.Picture = LoadPicture(gstrImageName)
frmMain.picStegano.Refresh
frmMain.picStegano.AutoSize = False

'get size of image
DataPicX = frmMain.picStegano.ScaleWidth
DataPicY = frmMain.picStegano.ScaleHeight

Xmax = DataPicX - 1
Ymax = DataPicY - 1
Cp = 0

NmbrPix = (DataPicX * DataPicY) - 1
For Xp = 1 To DataPicX
    For Yp = 1 To DataPicY
        'first read in the original colors
        Pixel = frmMain.picStegano.Point(Xp, Yp)
        R = Pixel And &HFF&
        G = (Pixel And &HFF00&) \ &H100&
        B = (Pixel And &HFF0000) \ &H10000
        R = (R And &HFE) Or Int(2 * Rnd)
        G = (G And &HFE) Or Int(2 * Rnd)
        B = (B And &HFE) Or Int(2 * Rnd)
        frmMain.picStegano.PSet (Xp, Yp), RGB(R, G, B)
        ProgressShow frmMain.picProgBar, (Xp * Yp) / NmbrPix
        DoEvents
        If gblnCancelProcces = True Then Exit For
    Next
Next
'refresh image
frmMain.picStegano.Picture = frmMain.picStegano.Image

If gblnCancelProcces = True Then
    'Extracting data from image aborted.
    MsgBox "Erasing data from image cancelled by user." & vbCrLfLf & "No data was removed.", vbExclamation
    Exit Sub
    End If

ProgressShow frmMain.picProgBar, 0

On Error GoTo errHandle


Screen.MousePointer = 11
If gstrImageName <> "" Then
    SavePicture frmMain.picStegano.Picture, gstrImageName
    gblnTextHasChanged = False
End If
Screen.MousePointer = 0
ProgressShow frmMain.picProgBar, 0

MsgBox "Erasing data from """ & CutFilePath(gstrImageName) & """ completed.", vbInformation

Exit Sub

errHandle:
Screen.MousePointer = 0
MsgBox "Failed erasing data from " & CutFilePath(killName) & vbCrLf & vbCrLf & Err.Description, vbCritical
ProgressShow frmMain.picProgBar, 0

End Sub


Public Sub WritePixel(ByVal aByte As Byte)
Dim currentBit As Integer
Dim Bit As Byte
Dim k As Integer

currentBit = 1
'get all bits of this byte
For k = 1 To 8
    If aByte And currentBit Then Bit = 1 Else Bit = 0
    'select which color to change
    Select Case Cp
        Case 0
            'first read in the original colors
            Pixel = frmMain.picStegano.Point(Xp, Yp)
            R = Pixel And &HFF&
            G = (Pixel And &HFF00&) \ &H100&
            B = (Pixel And &HFF0000) \ &H10000
            'set red
            R = (R And &HFE) Or Bit
        Case 1
            'set green
            G = (G And &HFE) Or Bit
        Case 2
            'set blue
            B = (B And &HFE) Or Bit
    End Select
    frmMain.picStegano.PSet (Xp, Yp), RGB(R, G, B)
    currentBit = currentBit * 2
    'next color, if r, g and b used, get next pixel
    Cp = Cp + 1
    If Cp > 2 Then
        Cp = 0
        'get next pixel position
        Call GetNextPixel
        End If
Next
'set ULTRA feedback byte
Call SetSteganoByte(aByte)
End Sub


Public Function ReadPixel() As Byte
Dim currentBit As Integer
Dim Bit As Byte
Dim k As Integer
Dim Pixel As Long
Dim R As Byte
Dim G As Byte
Dim B As Byte

Xmax = DataPicX - 1
Ymax = DataPicY - 1
currentBit = 1
'read in a complete byte (8 bits)
For k = 1 To 8
    'get all colors rgb
    Pixel = frmMain.picStegano.Point(Xp, Yp)
    R = Pixel And &HFF&
    G = (Pixel And &HFF00&) \ &H100&
    B = (Pixel And &HFF0000) \ &H10000
    Select Case Cp
        Case 0
            'read red
            Bit = (R And &H1)
        Case 1
            'read green
            Bit = (G And &H1)
        Case 2
            'read blue
            Bit = (B And &H1)
        End Select
    'add read pixel to byte
    If Bit Then ReadPixel = ReadPixel Or currentBit
    'set for next bit in byte
    currentBit = currentBit * 2
    'get next color
    Cp = Cp + 1
    If Cp > 2 Then
        Cp = 0
        'if all colors of this pixel used, get next pixel position
        Call GetNextPixel
        End If
Next k
'set ULTRA feedback byte
Call SetSteganoByte(ReadPixel)
End Function

Private Sub GetNextPixel()
'get pixel positions from ULTRA
Dim NewX As Long
Dim NewY As Long
Dim PixNr As Long
Dim FreePix As Long
Dim k As Long
Dim v1 As Long
Dim v2 As Long
'pull 4 byte from ULTRA and resize to image XY
v1 = GetSteganoPix
v2 = GetSteganoPix
NewX = (((v1 * 256) + v2) + 1) Mod DataPicX
v1 = GetSteganoPix
v2 = GetSteganoPix
NewY = (((v1 * 256) + v2) + 1) Mod DataPicY
'set pixel nr
PixNr = (NewY * DataPicX) + NewX
'check if pix has been generated before
If UsedPix(PixNr) = 0 Then
    'set newly used pix
    UsedPix(PixNr) = 1
    Xp = NewX
    Yp = NewY
    Else
    'seek next free pix
    FreePix = -1
    For k = PixNr To NmbrPix
        If UsedPix(k) = 0 Then FreePix = k: Exit For
    Next k
    If FreePix = -1 Then
    'restart search from begin array
        For k = 0 To PixNr
            If UsedPix(k) = 0 Then FreePix = k: Exit For
        Next k
        End If
    If FreePix = -1 Then
        PicOverflow = True
        Exit Sub
        End If
    'set new pixs
   Yp = Int(FreePix / DataPicX)
   Xp = FreePix - (Yp * DataPicX)
   UsedPix(FreePix) = 1
End If
End Sub
