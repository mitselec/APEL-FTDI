VERSION 5.00
Begin VB.Form Test_Form 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Apel Test Program "
   ClientHeight    =   2505
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4725
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2505
   ScaleWidth      =   4725
   StartUpPosition =   3  'Windows Default
   Begin VB.HScrollBar divscroll 
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   1680
      Width           =   2535
   End
   Begin VB.HScrollBar bytescroll 
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   1200
      Width           =   2535
   End
   Begin VB.CommandButton closedev 
      Caption         =   "Close Device"
      Height          =   615
      Left            =   2280
      TabIndex        =   1
      Top             =   240
      Width           =   1455
   End
   Begin VB.CommandButton opendev 
      Caption         =   "Open Device"
      Height          =   615
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   1455
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   4200
      Top             =   240
   End
   Begin VB.Label divlabel 
      Caption         =   "Divisor"
      Height          =   255
      Left            =   3240
      TabIndex        =   5
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label bytelabel 
      Caption         =   "Bytes"
      Height          =   255
      Left            =   3240
      TabIndex        =   3
      Top             =   1200
      Width           =   975
   End
End
Attribute VB_Name = "Test_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub opendev_Click()

Dim J As Long
Dim S As String * 64
J = 0

closedev_Click
DoEvents
'Finds the First FTDI module connected to the pc

aErr = FT_GetNumDevices(J, 0, FT_LIST_BY_NUMBER_ONLY)
J = 0
aErr = FT_ListDevices(J, S, (FT_OPEN_BY_DESCRIPTION Or FT_LIST_BY_INDEX))
aErr = FT_Open(J, aHandle)
aErr = FT_SetUSBParameters(aHandle, 65536, 65536)
aErr = FT_SetLatencyTimer(aHandle, 2&)
aErr = FT_SetBitMode(aHandle, &H0, &H0) 'Reset Device
aErr = FT_SetBitMode(aHandle, &HFF, &H4) 'Sychronous for 232R
aErr = FT_SetDivisor(aHandle, CLng(divscroll.Value))

DoEvents
Timer1.Enabled = True
End Sub

Private Sub closedev_Click()
'Shuts down the device

Timer1.Enabled = False
DoEvents
If aHandle <> 0 Then
    aErr = FT_Close(aHandle)
End If
End Sub

Private Sub Form_Load()
Dim counter As Long
'Fills a 32k buffer with alternating 1's & 0's
For counter = 1 To 32768 Step 2
    Mid$(aTxBuf, counter, 1) = Chr$(&HFF)
    Mid$(aTxBuf, counter + 1, 1) = Chr$(&H0)
Next counter

Me.Show
DoEvents
bytescroll.Value = 0
bytescroll_Change
End Sub

Function SentBytes(xBytes As Long)
'Function used to send bytes to the module

Dim J As Long
Dim sRX As Long
Dim sTX As Long
Dim sEV As Long
Dim rBuf As String * 32768
Dim rBytes As Long
Dim S As String * 64


If aHandle <> 0 Then
    aErr = FT_Purge(aHandle, FT_PURGE_RX Or FT_PURGE_TX)
    aErr = FT_Write(aHandle, aTxBuf, xBytes, J)
    aErr = FT_GetStatus(aHandle, sRX, sTX, sEV)
    aErr = FT_Read(aHandle, rBuf, sRX, rBytes)
    'Debug.Print "Bytes Wtitten:" & J; rBytes; sRX; sTX; sEV
End If

End Function

Private Sub Form_Unload(Cancel As Integer)
closedev_Click
End Sub

Private Sub bytescroll_Change()
bytelabel.Caption = "Bytes :" & xBytes
xBytes = bytescroll.Value
End Sub

Private Sub divscroll_Change()
divlabel.Caption = "Divisor :" & divscroll.Value
closedev_Click
For counter = 1 To 100
DoEvents
Next counter
opendev_Click
End Sub

Private Sub Timer1_Timer()
SentBytes xBytes
End Sub
