VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "USB�����q�T����-SetupDiGetDeviceInterfaceDetail�禡"
   ClientHeight    =   3180
   ClientLeft      =   255
   ClientTop       =   330
   ClientWidth     =   5445
   LinkTopic       =   "Form1"
   ScaleHeight     =   3180
   ScaleWidth      =   5445
   Begin VB.Frame fraSendAndReceive 
      Caption         =   "��J�P��X���"
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
'�HSetupDiGetDeviceInterfaceDetail�]�^�禡����USB HID�˸m
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

'******************************************************************************
'SetupDiEnumDeviceInterfaces�禡
'�^�ǡG �]�t�F���V�@�өҰ�����˸m��SP_DEVICE_INTERFACE_DATA���c��handle
'�ݨD�G
'1. �qSetupDiGetClassDevs�禡�Ҧ^�Ǫ�DeviceInfoSet
'2. �qGetHidGuid�Ҧ^�Ǫ�HidGuid
'3. �]�w���˸m�����ޭ�
'******************************************************************************

'�H0�}�l�A�å[�H���W����S���˸m�Q�����쬰��
MemberIndex = 0

Do
    '�bMyDeviceInterfaceData���c����cbSize���������Q�]�w
    '�o�ӵ��c�O�H�줸�լ��p�ƼƭȡA�o�Ӥj�p�ȬO28�Ӧ줸��
    MyDeviceInterfaceData.cbSize = LenB(MyDeviceInterfaceData)
    Result = SetupDiEnumDeviceInterfaces _
        (DeviceInfoSet, _
        0, _
        HidGuid, _
        MemberIndex, _
        MyDeviceInterfaceData)
    
    Call DisplayResultOfAPICall("SetupDiEnumDeviceInterfaces")
    If Result = 0 Then LastDevice = True
    
    '�p�G�˸m�s�b���ܡA��ܦ^�Ǫ��T��
    If Result <> 0 Then
        lstResults.AddItem "  DeviceInfoSet for device #" & CStr(MemberIndex) & ": "
        lstResults.AddItem "  cbSize = " & CStr(MyDeviceInterfaceData.cbSize)
        lstResults.AddItem _
            "  InterfaceClassGuid.Data1 = " & Hex$(MyDeviceInterfaceData.InterfaceClassGuid.Data1)
        lstResults.AddItem _
            "  InterfaceClassGuid.Data2 = " & Hex$(MyDeviceInterfaceData.InterfaceClassGuid.Data2)
        lstResults.AddItem _
            "  InterfaceClassGuid.Data3 = " & Hex$(MyDeviceInterfaceData.InterfaceClassGuid.Data3)
        lstResults.AddItem _
            "  Flags = " & Hex$(MyDeviceInterfaceData.Flags)

        
'******************************************************************************
'SetupDiGetDeviceInterfaceDetail�禡
'�^�ǡG�]�t���˸m�T����SP_DEVICE_INTERFACE_DETAIL_DATA���c
'�`�N�ƶ��G���F���^�������T���A�I�s���禡�⦸
'�Ĥ@���Ǧ^�һݪ����c�j�p
'�ĤG���Ǧ^�bDeviceInfoSet��ƪ����о�
'�ݨD�G��SetupDiGetClassDevs�Ҧ^�Ǫ�DeviceInfoSet
 '�H�Υ�SetupDiEnumDeviceInterfaces�Ҧ^�Ǫ�SP_DEVICE_INTERFACE_DATA���c
'*******************************************************************************
        
        MyDeviceInfoData.cbSize = Len(MyDeviceInfoData)
        Result = SetupDiGetDeviceInterfaceDetail _
           (DeviceInfoSet, _
           MyDeviceInterfaceData, _
           0, _
           0, _
           Needed, _
           0)
        
        DetailData = Needed
            
        Call DisplayResultOfAPICall("SetupDiGetDeviceInterfaceDetail")
        lstResults.AddItem "  (OK to say too small)"
        lstResults.AddItem "  Required buffer size for the data: " & Needed
        
        '�x�s�o�ӵ��c���j�p��
        MyDeviceInterfaceDetailData.cbSize = _
            Len(MyDeviceInterfaceDetailData)
        
        '�ϥΦ줸�հ}�C�h�ŧiMyDeviceInterfaceDetailData���c���O����
        ReDim DetailDataBuffer(Needed)
        
        '�x�s�b�}�C���e�|�Ӧ줸�աAcbSize
         Call RtlMoveMemory _
            (DetailDataBuffer(0), _
            MyDeviceInterfaceDetailData, _
            4)
        
        '�A�@���I�sSetupDiGetDeviceInterfaceDetail�禡
        '�o�@���A�O�ǻ�DetailDataBuffer�Ĥ@�Ӥ�������}
        '�Ӧ��^�Ǫ��O�bDetailData�һݪ��w�İϤj�p
        
        Result = SetupDiGetDeviceInterfaceDetail _
           (DeviceInfoSet, _
           MyDeviceInterfaceData, _
           VarPtr(DetailDataBuffer(0)), _
           DetailData, _
           Needed, _
           0)
        
        Call DisplayResultOfAPICall(" Result of second call: ")
        lstResults.AddItem "  MyDeviceInterfaceDetailData.cbSize: " & _
            CStr(MyDeviceInterfaceDetailData.cbSize)
        
        '�N�줸�հ}�C�ഫ���r��
        DevicePathName = CStr(DetailDataBuffer())
        '�ഫ��Unicode�X�榡.
        DevicePathName = StrConv(DevicePathName, vbUnicode)
        '�q�}�l�B�簣cbSize (4 bytes).
        DevicePathName = Right$(DevicePathName, Len(DevicePathName) - 4)
        lstResults.AddItem "  Device pathname: "
        lstResults.AddItem "    " & DevicePathName
                
        '******************************************************************************
        'CreateFile�禡
        '�^�ǡG
        '�^�ǥΨӭP��Ū���P�g�J���˸m��handle
        '�ݨD:
        '��SetupDiGetDeviceInterfaceDetail�禡�ҶǦ^��DevicePathName
        '******************************************************************************
    
        HidDevice = CreateFile _
            (DevicePathName, _
            GENERIC_READ Or GENERIC_WRITE, _
            (FILE_SHARE_READ Or FILE_SHARE_WRITE), _
            0, _
            OPEN_EXISTING, _
            0, _
            0)
            
        'Call DisplayResultOfAPICall("CreateFile")
        'lstResults.AddItem "  Returned handle: " & Hex$(HidDevice) & "h"
        
        '��o�̡A�w�g���ڭ̥��b�M�䪺�˸m
        
        '******************************************************************************
        'HidD_GetAttributes
        '�ݨD�G
         '�ݭn�q�˸m���T���A�H�αqCreatFile�禡�Ҧ^�Ǫ�Handle
        '�^�ǡG
        '�^�ǥ]�t�F�c���VID�A���~PID�H�β��~�����X��HIDD_ATTRIBUTES���c
        '�ڭ̥i�H�ϥγo�ӰT���ӨM�w�O�_�o�Ӹ˸m�O�ڭ̩ҭn�M�䪺.
        '******************************************************************************
        
        '�]�w"Size"�ݩʬ��b���c�����줸�ռƭ�
        DeviceAttributes.Size = LenB(DeviceAttributes)
        Result = HidD_GetAttributes _
            (HidDevice, _
            DeviceAttributes)
            
        'Call DisplayResultOfAPICall("HidD_GetAttributes")
        'If Result <> 0 Then
        '    lstResults.AddItem "  HIDD_ATTRIBUTES structure filled without error."
        'Else
        '    lstResults.AddItem "  Error in filling HIDD_ATTRIBUTES structure."
        'End If
    
        'lstResults.AddItem "  Structure size: " & DeviceAttributes.Size
        'lstResults.AddItem "  Vendor ID: " & Hex$(DeviceAttributes.VendorID)
        'lstResults.AddItem "  Product ID: " & Hex$(DeviceAttributes.ProductID)
        'lstResults.AddItem "  Version Number: " & Hex$(DeviceAttributes.VersionNumber)
        
        '��X�O�_�o�Ӹ˸m�P�ڭ̥��b�M�䪺�˸m�k�XFind out if the device matches the one we're looking for.
        If (DeviceAttributes.VendorID = MyVendorID) And _
            (DeviceAttributes.ProductID = MyProductID) Then
        '        lstResults.AddItem "  My device detected"
                MyDeviceDetected = True
        Else
                MyDeviceDetected = False
                '�p�G���O�ڭ̩ҷQ�n���˸m�A�N������handle.
                Result = CloseHandle _
                    (HidDevice)
                DisplayResultOfAPICall ("CloseHandle")
        End If
End If
    '����a�M�䪽��ڭ̤w�o�{���˸m�A�άO�w�g�S���ѤU���˸m�n�ˬd

    MemberIndex = MemberIndex + 1
    
Loop Until (LastDevice = True) Or (MyDeviceDetected = True)
    
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


