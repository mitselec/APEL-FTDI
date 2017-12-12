Attribute VB_Name = "Module1"
Option Explicit

'Public fMainForm As DEMO_EEPROM
'===========================
'CLASSIC INTERFACE
'===========================
Public Declare Function FT_ListDevices Lib "FTD2XX.DLL" (ByVal arg1 As Long, ByVal arg2 As String, ByVal dwFlags As Long) As Long
Public Declare Function FT_GetNumDevices Lib "FTD2XX.DLL" Alias "FT_ListDevices" (ByRef arg1 As Long, ByVal arg2 As String, ByVal dwFlags As Long) As Long

Public Declare Function FT_Open Lib "FTD2XX.DLL" (ByVal intDeviceNumber As Integer, ByRef lngHandle As Long) As Long
Public Declare Function FT_OpenEx Lib "FTD2XX.DLL" (ByVal arg1 As String, ByVal arg2 As Long, ByRef lngHandle As Long) As Long
Public Declare Function FT_Close Lib "FTD2XX.DLL" (ByVal lngHandle As Long) As Long
Public Declare Function FT_Read Lib "FTD2XX.DLL" (ByVal lngHandle As Long, ByVal lpszBuffer As String, ByVal lngBufferSize As Long, ByRef lngBytesReturned As Long) As Long
Public Declare Function FT_Write Lib "FTD2XX.DLL" (ByVal lngHandle As Long, ByVal lpszBuffer As String, ByVal lngBufferSize As Long, ByRef lngBytesWritten As Long) As Long
Public Declare Function FT_WriteByte Lib "FTD2XX.DLL" Alias "FT_Write" (ByVal lngHandle As Long, ByRef lpszBuffer As Any, ByVal lngBufferSize As Long, ByRef lngBytesWritten As Long) As Long
Public Declare Function FT_SetBaudRate Lib "FTD2XX.DLL" (ByVal lngHandle As Long, ByVal lngBaudRate As Long) As Long
Public Declare Function FT_SetUSBParameters Lib "FTD2XX.DLL" (ByVal lngHandle As Long, ByVal InDataSize As Long, ByVal OutDataSize As Long) As Long
Public Declare Function FT_SetDataCharacteristics Lib "FTD2XX.DLL" (ByVal lngHandle As Long, ByVal byWordLength As Byte, ByVal byStopBits As Byte, ByVal byParity As Byte) As Long
Public Declare Function FT_SetFlowControl Lib "FTD2XX.DLL" (ByVal lngHandle As Long, ByVal intFlowControl As Integer, ByVal byXonChar As Byte, ByVal byXoffChar As Byte) As Long
Public Declare Function FT_SetDtr Lib "FTD2XX.DLL" (ByVal lngHandle As Long) As Long
Public Declare Function FT_ClrDtr Lib "FTD2XX.DLL" (ByVal lngHandle As Long) As Long
Public Declare Function FT_SetRts Lib "FTD2XX.DLL" (ByVal lngHandle As Long) As Long
Public Declare Function FT_ClrRts Lib "FTD2XX.DLL" (ByVal lngHandle As Long) As Long
Public Declare Function FT_GetModemStatus Lib "FTD2XX.DLL" (ByVal lngHandle As Long, ByRef lngModemStatus As Long) As Long
Public Declare Function FT_SetChars Lib "FTD2XX.DLL" (ByVal lngHandle As Long, ByVal byEventChar As Byte, ByVal byEventCharEnabled As Byte, ByVal byErrorChar As Byte, ByVal byErrorCharEnabled As Byte) As Long
Public Declare Function FT_Purge Lib "FTD2XX.DLL" (ByVal lngHandle As Long, ByVal lngMask As Long) As Long
Public Declare Function FT_SetTimeouts Lib "FTD2XX.DLL" (ByVal lngHandle As Long, ByVal lngReadTimeout As Long, ByVal lngWriteTimeout As Long) As Long
Public Declare Function FT_GetQueueStatus Lib "FTD2XX.DLL" (ByVal lngHandle As Long, ByRef lngRxBytes As Long) As Long
Public Declare Function FT_SetBreakOn Lib "FTD2XX.DLL" (ByVal lngHandle As Long) As Long
Public Declare Function FT_SetBreakOff Lib "FTD2XX.DLL" (ByVal lngHandle As Long) As Long
Public Declare Function FT_GetStatus Lib "FTD2XX.DLL" (ByVal lngHandle As Long, ByRef lngRxBytes As Long, ByRef lngTxBytes As Long, ByRef lngEventsDWord As Long) As Long
Public Declare Function FT_SetEventNotification Lib "FTD2XX.DLL" (ByVal lngHandle As Long, ByVal dwEventMask As Long, ByVal pVoid As Long) As Long
Public Declare Function FT_ResetDevice Lib "FTD2XX.DLL" (ByVal lngHandle As Long) As Long
Public Declare Function FT_SetDivisor Lib "FTD2XX.DLL" (ByVal lngHandle As Long, ByVal shDivisor As Long) As Long

'Public Declare Function FT_GetEventStatus Lib "FTD2XX.DLL" (ByVal lngHandle As Long, ByRef lngEventsDWord As Long) As Long

 Public Declare Function FT_GetBitMode Lib "FTD2XX.DLL" (ByVal lngHandle As Long, ByRef intData As Any) As Long
 Public Declare Function FT_SetBitMode Lib "FTD2XX.DLL" (ByVal lngHandle As Long, ByVal intMask As Byte, ByVal intMode As Byte) As Long

Public Declare Function FT_SetLatencyTimer Lib "FTD2XX.DLL" (ByVal Handle As Long, ByVal pucTimer As Byte) As Long
Public Declare Function FT_GetLatencyTimer Lib "FTD2XX.DLL" (ByVal Handle As Long, ByRef ucTimer As Long) As Long



'=============================
'FT_W32 API
'=============================

Public Declare Function FT_W32_CreateFile Lib "FTD2XX.DLL" (ByVal lpszName As String, ByVal dwAccess As Long, ByVal dwShareMode As Long, ByRef lpSecurityAttributes As LPSECURITY_ATTRIBUTES, ByVal dwCreate As Long, ByVal dwAttrsAndFlags As Long, ByVal hTemplate As Long) As Long
Public Declare Function FT_W32_CloseHandle Lib "FTD2XX.DLL" (ByVal ftHandle As Long) As Long
Public Declare Function FT_W32_ReadFile Lib "FTD2XX.DLL" (ByVal lngHandle As Long, ByVal lpszBuffer As String, ByVal lngBufferSize As Long, ByRef lngBytesReturned As Long, ByRef lpftOverlapped As lpOverlapped) As Long
Public Declare Function FT_W32_WriteFile Lib "FTD2XX.DLL" (ByVal lngHandle As Long, ByVal lpszBuffer As String, ByVal lngBufferSize As Long, ByRef lngBytesWritten As Long, ByRef lpftOverlapped As lpOverlapped) As Long
Public Declare Function FT_W32_GetOverlappedResult Lib "FTD2XX.DLL" (ByVal lngHandle As Long, ByRef lpftOverlapped As lpOverlapped, ByRef lpdwBytesTransferred As Long, ByVal bWait As Boolean) As Long
Public Declare Function FT_W32_GetCommState Lib "FTD2XX.DLL" (ByVal lngHandle, ByRef lpftDCB As FTDCB) As Long
Public Declare Function FT_W32_SetCommState Lib "FTD2XX.DLL" (ByVal lngHandle, ByRef lpftDCB As FTDCB) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Declare Function CreateEvent Lib "kernel32" Alias "CreateEventA" (ByVal lpEventAttributes As Long, ByVal bManualReset As Long, ByVal bInitialState As Long, ByVal lpName As String) As Long
Public Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Public Declare Function SetEvent Lib "kernel32" (ByVal hHandle As Long) As Long
Public Declare Function CreateThread Lib "kernel32" (lpThreadAttributes As Any, ByVal dwStackSize As Long, ByVal lpStartAddress As Long, lpParameter As Any, ByVal dwCreationFlags As Long, lpThreadID As Long) As Long
Public Declare Function TerminateThread Lib "kernel32" (ByVal hThread As Long, ByVal dwExitCode As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

'====================================================================
'APIGID32.DLL by DESAWARE Inc. (www.desaware.com), see Dan Appleman's
'"Visual Basic Programmer's Guide to the WIN32-API"; here used to get
'the addresses of the VB-bytearrays:
'====================================================================
Public Declare Function agGetAddressForObject& Lib "apigid32.dll" (object As Any)

'==============================================================
'Declarations for the EEPROM-accessing functions in FTD2XX.dll:
'==============================================================
Public Declare Function FT_EE_Program Lib "FTD2XX.DLL" (ByVal lngHandle As Long, ByRef lpData As FT_PROGRAM_DATA) As Long
Public Declare Function FT_EE_Read Lib "FTD2XX.DLL" (ByVal lngHandle As Long, ByRef lpData As FT_PROGRAM_DATA) As Long
Public Declare Function FT_EE_UASize Lib "FTD2XX.DLL" (ByVal lngHandle As Long, ByRef lpdwSize As Long) As Long
Public Declare Function FT_EE_UAWrite Lib "FTD2XX.DLL" (ByVal lngHandle As Long, ByVal pucData As String, ByVal dwDataLen As Long) As Long
Public Declare Function FT_EE_UARead Lib "FTD2XX.DLL" (ByVal lngHandle As Long, ByVal pucData As String, ByVal dwDataLen As Long, ByRef lpdwBytesRead As Long) As Long

Public Type LPSECURITY_ATTRIBUTES
  nLength As Long
  lpSecurityDescriptor As Long
  bInheritHandle As Long
End Type

Public Type lpOverlapped
  Internal As Long
  InternalHigh As Long
  Offset As Long
  OffsetHigh As Long
  hEvent As Long
End Type

Public Type FTDCB
    DCBlength As Long                   'sizeof (FTDCB)
    BaudRate As Long                    '9600
'    fBinary As Long                     '= 1 Binary mode (skip EOF check)
'    fParity As Long                     '= 1 Enable parity checking
'    fOutxCtsFlow As Long                '= 1 CTS handshaking on output
'    fOutxDsrFlow As Long                '= 1 DSR handshaking on output
'    fDtrControl As Long                 '= 2 DTR flow control
'    fDsrSensitivity As Long             '= 1 DSR Sensitivity
'    fTXContinueOnXoff As Long           '= 1 Continue TX when Xoff sent
'    fOutX As Long                       '= 1 Enable output X-on/X-off
'    fInX As Long                        '= 1 Enable input X-on/X-off
'    fErrorChar As Long                  '= 1 Enable error replacement
'    fNull As Long                       '= 1 Enable null stripping
'    fRtsControl As Long                 '= 2 RTS flow control
'    fAbortOnError As Long               '= 1 Abort all reads and writes on error
'    fDummy2 As Long                     '= 17 Reserved
'    wReserved As Integer                'Not currently used
'    XonLim As Integer                   'Transmit X-on threshold
'    XoffLim As Integer                  'Transmit X-off threshold
'    ByteSize As Byte                    'Number of bits/ byte, 7-8
'    Parity As Byte                      '0-4= None, Odd, Even, Mark, Space
'    StopBits As Byte                    '0, 2 = 1, 2
'    XonChar As Byte                     'TX and RX X-ON character
'    XoffChar As Byte                    'TX and RX X-OFF character
'    ErrorChar As Byte                   'Eror replacement char
'    EofChar As Byte                     'End of input Character
'    EvtChar As Byte                     'Received event character
'    wReserved1 As Integer               'BCD (0x0200 => USB2)
End Type




'====================================================================
'Type definition as equivalent for C-structure "ft_program_data" used
'in FT_EE_READ and FT_EE_WRITE;
'ATTENTION! The variables "Manufacturer", "ManufacturerID",
'"Description" and "SerialNumber" are passed as POINTERS to
'locations of Bytearrays. Each Byte in these arrays will be
'filled with one character of the whole string.
'(See below, calls to "agGetAddressForObject")
'=====================================================================


Public Type FT_PROGRAM_DATA

    signature1 As Long                  '0x00000000
    signature2 As Long                  '0xFFFFFFFF
    version As Long                     '0

    VendorId As Integer                 '0x0403
    ProductId As Integer                '0x6001
    Manufacturer As Long                '32 "FTDI"
    ManufacturerId As Long              '16 "FT"
    Description As Long                 '64 "USB HS Serial Converter"
    SerialNumber As Long                '16 "FT000001" if fixed, or NULL
    MaxPower As Integer                 ' // 0 < MaxPower <= 500
    PnP As Integer                      ' // 0 = disabled, 1 = enabled
    SelfPowered As Integer              ' // 0 = bus powered, 1 = self powered
    RemoteWakeup As Integer             ' // 0 = not capable, 1 = capable
     'Rev4 extensions:
    Rev4 As Byte                        ' // true if Rev4 chip, false otherwise
    IsoIn As Byte                       ' // true if in endpoint is isochronous
    IsoOut As Byte                      ' // true if out endpoint is isochronous
    PullDownEnable As Byte              ' // true if pull down enabled
    SerNumEnable As Byte                ' // true if serial number to be used
    USBVersionEnable As Byte            ' // true if chip uses USBVersion
    USBVersion As Integer               ' // BCD (0x0200 => USB2)
    'FT2232C extensions:
    Rev5 As Byte                        'non-zero if Rev5 chip, zero otherwise
    IsoInA As Byte                      'non-zero if in endpoint is isochronous
    IsoInB As Byte                      'non-zero if in endpoint is isochronous
    IsoOutA As Byte                     'non-zero if out endpoint is isochronous
    IsoOutB As Byte                     'non-zero if out endpoint is isochronous
    PullDownEnable5 As Byte             'non-zero if pull down enabled
    SerNumEnable5 As Byte               'non-zero if serial number to be used
    USBVersionEnable5 As Byte           'non-zero if chip uses USBVersion
    USBVersion5 As Integer              'BCD 0x110 = USB 1.1, BCD 0x200 = USB 2.0
    AlsHighCurrent As Byte              'non-zero if interface is high current
    BlsHighCurrent As Byte              'non-zero if interface is high current
    IFAlsFifo As Byte                   'non-zero if interface is 245 FIFO
    IFAlsFifoTar As Byte                'non-zero if interface is 245 FIFO CPU target
    IFAlsFastSer As Byte                'non-zero if interface is Fast Serial
    AlsVCP As Byte                      'non-zero if interface is to use VCP drivers
    IFBlsFifo As Byte                   'non-zero if interface is 245 FIFO
    IFBlsFifoTar As Byte                'non-zero if interface is 245 FIFO CPU target
    IFBlsFastSer As Byte                'non-zero if interface is Fast Serial
    BlsVCP As Byte                      'non-zero if interface is to use VCP drivers
    'FT232R extensions
    UseExtOSC As Byte                   'non-zero use ext osc
    HighDriveIOs As Byte                'non-zero to use High Drive IO's
    EndPointSize As Byte                '64 Do not change
    PullDownEnableR As Byte             'non-zeero if pull down enabled
    SerNumEnableR As Byte               'non-zero if pull serial number enabled
    InvertTXD As Byte                   'non-zero if invert TXD
    InvertRXD As Byte                   'non-zero if invert RXD
    InvertRTS As Byte                   'non-zero if invert RTS
    InvertCTS As Byte                   'non-zero if invert CTS
    InvertDTR As Byte                   'non-zero if invert DTR
    InvertDSR As Byte                   'non-zero if invert DSR
    InvertDCD As Byte                   'non-zero if invert DCD
    InvertRI As Byte                    'non-zero if invert RI
    Cbus0 As Byte                       'Cbus Mux control
    Cbus1 As Byte                       'Cbus Mux control
    Cbus2 As Byte                       'Cbus Mux control
    Cbus3 As Byte                       'Cbus Mux control
    Cbus4 As Byte                       'Cbus Mux control
    RIsVCP As Byte                      'non-zero if using VCP driver
    
    
End Type



' Return codes
Public Const FT_OK = 0
Public Const FT_INVALID_HANDLE = 1
Public Const FT_DEVICE_NOT_FOUND = 2
Public Const FT_DEVICE_NOT_OPENED = 3
Public Const FT_IO_ERROR = 4
Public Const FT_INSUFFICIENT_RESOURCES = 5
Public Const FT_INVALID_PARAMETER = 6
Public Const FT_INVALID_BAUD_RATE = 7
Public Const FT_DEVICE_NOT_OPENED_FOR_ERASE = 8
Public Const FT_DEVICE_NOT_OPENED_FOR_WRITE = 9
Public Const FT_FAILED_TO_WRITE_DEVICE = 10
Public Const FT_EEPROM_READ_FAILED = 11
Public Const FT_EEPROM_WRITE_FAILED = 12
Public Const FT_EEPROM_ERASE_FAILED = 13
Public Const FT_EEPROM_NOT_PRESENT = 14
Public Const FT_EEPROM_NOT_PROGRAMMED = 15
Public Const FT_INVALID_ARGS = 16
Public Const FT_NOT_SUPPORTED = 17
Public Const FT_OTHER_ERROR = 18

' Word Lengths
Public Const FT_BITS_8 = 8
Public Const FT_BITS_7 = 7

' Stop Bits
Public Const FT_STOP_BITS_1 = 0
Public Const FT_STOP_BITS_1_5 = 1
Public Const FT_STOP_BITS_2 = 2

' Parity
Public Const FT_PARITY_NONE = 0
Public Const FT_PARITY_ODD = 1
Public Const FT_PARITY_EVEN = 2
Public Const FT_PARITY_MARK = 3
Public Const FT_PARITY_SPACE = 4

' Flow Control
Public Const FT_FLOW_NONE = &H0
Public Const FT_FLOW_RTS_CTS = &H100
Public Const FT_FLOW_DTR_DSR = &H200
Public Const FT_FLOW_XON_XOFF = &H400

' Purge rx and tx buffers
Public Const FT_PURGE_RX = 1
Public Const FT_PURGE_TX = 2

' Modem Status
Public Const FT_MODEM_STATUS_CTS = &H10
Public Const FT_MODEM_STATUS_DSR = &H20
Public Const FT_MODEM_STATUS_RI = &H40
Public Const FT_MODEM_STATUS_DCD = &H80

Public Const FT_EVENT_RXCHAR As Long = 1
Public Const FT_EVENT_MODEM_STATUS = 2

Const WAIT_ABANDONED As Long = &H80
Const WAIT_FAILD As Long = &HFFFFFFFF
Const WAIT_OBJECT_0 As Long = &H0
Const WAIT_TIMEOUT As Long = &H102

' Flags for FT_ListDevices
Public Const FT_LIST_BY_NUMBER_ONLY = &H80000000
Public Const FT_LIST_BY_INDEX = &H40000000
Public Const FT_LIST_ALL = &H20000000

' Flags for FT_OpenEx
Public Const FT_OPEN_BY_SERIAL_NUMBER = 1
Public Const FT_OPEN_BY_DESCRIPTION = 2


Private Const INFINITE As Long = 1000   '&HFFFFFFFF

Global hThread As Long
Global hThreadID As Long
Global hEvent As Long
Global EventMask As Long

Global lngHandle As Long
 
    

'Sub Main()
'    Set fMainForm = New DEMO_EEPROM
'    fMainForm.Show
'End Sub

