VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "USB �����q�T����-SetupDiGetClassDevsc�禡"
   ClientHeight    =   3180
   ClientLeft      =   255
   ClientTop       =   330
   ClientWidth     =   5445
   LinkTopic       =   "Form1"
   ScaleHeight     =   3180
   ScaleWidth      =   5445
   Begin VB.Frame fraSendAndReceive 
      Caption         =   "��J/��X���"
      Height          =   855
      Left            =   0
      TabIndex        =   1
      Top             =   120
      Width           =   5415
      Begin VB.CommandButton cmdOnce 
         Caption         =   "�T�w"
         Height          =   495
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   5175
      End
   End
   Begin VB.ListBox lstResults 
      Height          =   2010
      ItemData        =   "USBHID.frx":0000
      Left            =   0
      List            =   "USBHID.frx":0002
      TabIndex        =   0
      Top             =   1080
      Width           =   5415
   End
   Begin VB.Timer tmrContinuousDataCollect 
      Left            =   4440
      Top             =   1080
   End
   Begin VB.Timer tmrDelay 
      Enabled         =   0   'False
      Left            =   120
      Top             =   11400
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'�����]�p�P���-�ϥ�Visual Basic�ACH19
'�HSetupDiGetClassDevs�]�^�禡����USB HID�˸m
'PC�D���PHID�˸m���q�T�{�� -> �ק�� Jan Axelson (jan@lvr.com)

Option Explicit

Dim Capabilities As HIDP_CAPS
Dim DataString As String
Dim DetailData As Long
Dim DetailDataBuffer() As Byte
Dim DeviceAttributes As HIDD_ATTRIBUTES
Dim DevicePathName As String
Dim DeviceInfoSet As Long
Dim ErrorString As String
Dim HidDevice As Long
Dim LastDevice As Boolean
Dim MyDeviceDetected As Boolean
Dim MyDeviceInfoData As SP_DEVINFO_DATA
Dim MyDeviceInterfaceDetailData As SP_DEVICE_INTERFACE_DETAIL_DATA
Dim MyDeviceInterfaceData As SP_DEVICE_INTERFACE_DATA
Dim Needed As Long
Dim OutputReportData(7) As Byte
Dim PreparsedData As Long
Dim Result As Long
Dim Timeout As Boolean

'�]�wVID/PID�X�H�ũM�˸m����{���X�PINF�w���ɮ�

Const MyVendorID = &H1234
Const MyProductID = &H2468

Function FindTheHid() As Boolean
'����@�t�C��API�禡�I�s�A�H���һݪ�HID�s�ո˸m
'�p�G������HID�˸m���ܡA�N�Ǧ^�u�ȡA�Ϥ��N�Ǧ^����

Dim Count As Integer
Dim GUIDString As String
Dim HidGuid As GUID
Dim MemberIndex As Long

LastDevice = False
MyDeviceDetected = False

'******************************************************************************
'HidD_GetHidGuid�禡
'���o�Ҧ��s���W��HID�˸m��GUID
'�Ǧ^��: HidGuid����GUID
'�o�Ө禡���|�^��Result�ƭ�
'���O�o�Ө禡�|�H�禡���覡�ӥ[�H�ŧi�A�H�ϱo�i�H�P��L��API�禡�ۮe
'******************************************************************************

Result = HidD_GetHidGuid(HidGuid)
Call DisplayResultOfAPICall("GetHidGuid")

'==== ��� GUID ====
GUIDString = _
    Hex$(HidGuid.Data1) & "-" & _
    Hex$(HidGuid.Data2) & "-" & _
    Hex$(HidGuid.Data3) & "-"

For Count = 0 To 7
    '�ھ�GUID�ŧi���榡�A�ݽT�O�C�@��GUID��8�Ӧ줸�����2�Ӧr��
    If HidGuid.Data4(Count) >= &H10 Then
        GUIDString = GUIDString & Hex$(HidGuid.Data4(Count)) & " "
    Else
        GUIDString = GUIDString & "0" & Hex$(HidGuid.Data4(Count)) & " "
    End If
Next Count

lstResults.AddItem "  GUID for system HIDs: "
lstResults.AddItem "  " & GUIDString

'******************************************************************************
'SetupDiGetClassDevs
'�^�ǡG �w��w�w�˪��˸m�^�Ǥ@�ӫ��V�˸m�T������handle
'�ݨD�G��GetHidGuid�Ҧ^�Ǫ�HidGuid
'******************************************************************************

DeviceInfoSet = SetupDiGetClassDevs _
    (HidGuid, _
    vbNullString, _
    0, _
    (DIGCF_PRESENT Or DIGCF_DEVICEINTERFACE))
    
Call DisplayResultOfAPICall("SetupDiClassDevs")
DataString = GetDataString(DeviceInfoSet, 32)
lstResults.AddItem "  DeviceInfoSet:" & DeviceInfoSet

End Function

Private Function GetDataString _
    (Address As Long, _
    Bytes As Long) _
As String

'���^�O����Retrieves a string of length Bytes from memory, beginning at Address.
'Adapted from Dan Appleman's "Win32 API Puzzle Book"

Dim Offset As Integer
Dim Result$
Dim ThisByte As Byte

For Offset = 0 To Bytes - 1
    Call RtlMoveMemory(ByVal VarPtr(ThisByte), ByVal Address + Offset, 1)
    If (ThisByte And &HF0) = 0 Then
        Result$ = Result$ & "0"
    End If
    Result$ = Result$ & Hex$(ThisByte) & " "
Next Offset

GetDataString = Result$
End Function

Private Function GetErrorString _
    (ByVal LastError As Long) _
As String

'�Ǧ^�W�@�ӿ��~�����~�T��

Dim Bytes As Long
Dim ErrorString As String
ErrorString = String$(129, 0)
Bytes = FormatMessage _
    (FORMAT_MESSAGE_FROM_SYSTEM, _
    0&, _
    LastError, _
    0, _
    ErrorString$, _
    128, _
    0)
    
'�q�T�������L CR�PLF�r���A�]����h2�Ӧr��

If Bytes > 2 Then
    GetErrorString = Left$(ErrorString, Bytes - 2)
End If

End Function

Private Sub cmdOnce_Click()
lstResults.Clear
Call FindTheHid
End Sub
Private Sub DisplayResultOfAPICall(FunctionName As String)

'�e�{API�I�s�����G

Dim ErrorString As String

lstResults.AddItem ""
ErrorString = GetErrorString(Err.LastDllError)
lstResults.AddItem FunctionName
lstResults.AddItem "  Result = " & ErrorString

End Sub

Private Sub Form_Load()
frmMain.Show
tmrDelay.Enabled = False

End Sub


