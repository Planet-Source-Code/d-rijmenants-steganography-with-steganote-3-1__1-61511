Attribute VB_Name = "modPrint"
'-----------------------------------------------------------------
'
'                Print module v3 (C) Dirk Rijmenants 2004
'
'   sub PrintString
'   PrintString Text, leftfmargin, rightmargin, topmargin, bottommargin
'   margins are long values 0-100 percent
'
'-----------------------------------------------------------------
Option Explicit

Public Sub PrintString(printVar As String, leftMargePrcnt As Long, rightMargePrcnt As Long, topMargePrcnt As Long, bottomMargePrcnt As Long)
Dim lMarge As Long
Dim rMarge As Long
Dim tMarge As Long
Dim bMarge As Long
Dim printLijn As String
Dim staPos  As Long
Dim endPos As Long
Dim txtHoogte As Long
Dim printHoogte As Long
Dim objectHoogte As Long
Dim objectBreedte As Long
Dim currYpos As Long
Dim cutChar As String
Dim k As Long
Dim cutPos As Long
frmPrintProgress.ProgressBar1.Max = Len(printVar) + 4
txtHoogte = Printer.TextHeight("AbgWq")
lMarge = Int((Printer.Width / 100) * leftMargePrcnt)
rMarge = Int((Printer.Width / 100) * rightMargePrcnt)
tMarge = Int((Printer.Height / 100) * topMargePrcnt)
bMarge = Int((Printer.Height / 100) * bottomMargePrcnt)
objectHoogte = Printer.Height - tMarge - bMarge
objectBreedte = Printer.Width - lMarge - rMarge
Printer.CurrentY = tMarge
staPos = 1
endPos = 0
Do

'get next line to crlf
endPos = InStr(staPos, printVar, vbCrLf)
If endPos <> 0 Then
    printLijn = Mid(printVar, staPos, endPos - staPos)
    Else
    printLijn = Mid(printVar, staPos)
    endPos = Len(printVar)
    End If
    
'check lenght one line
If Printer.TextWidth(printLijn) <= objectBreedte Then
    'line ok, keep line as it is
    staPos = endPos + 2
    Else
    'line to big, try to cut of at space or other signs within limits
    cutPos = 0
    For k = 1 To Len(printLijn)
        cutChar = Mid(printLijn, k, 1)
        If cutChar = " " Or cutChar = "." Or cutChar = "," Or cutChar = ":" Or cutChar = ")" Then
            If Printer.TextWidth(Left(printLijn, k)) > objectBreedte Then Exit For
            cutPos = k
        End If
    Next k
    'check result search for space
    If cutPos > 1 Then
        'cut off on space
        printLijn = Mid(printVar, staPos, cutPos)
        staPos = staPos + cutPos
        Else
        'no cut-character found within limits, so cut line on paperwidth
        For k = 1 To Len(printLijn)
            If Printer.TextWidth(Left(printLijn, k)) > objectBreedte Then Exit For
        Next k
        printLijn = Mid(printVar, staPos, k - 1)
        staPos = staPos + (k - 1)
    End If
End If

'print line
Printer.CurrentX = lMarge
currYpos = Printer.CurrentY + txtHoogte
If currYpos > tMarge + objectHoogte Then
    Printer.NewPage
    Printer.CurrentY = tMarge
    Printer.CurrentX = lMarge
    End If
Printer.Print printLijn
'check for cancel
DoEvents
If gblnCancelPrint = True Then
    Printer.KillDoc
    Exit Do
    End If
'Update ProgressBar
frmPrintProgress.ProgressBar1.Value = staPos
'frmPrintProgress.ProgressBar1.Refresh
Loop While staPos < Len(printVar)
Printer.EndDoc
frmPrintProgress.ProgressBar1.Value = 0
End Sub


