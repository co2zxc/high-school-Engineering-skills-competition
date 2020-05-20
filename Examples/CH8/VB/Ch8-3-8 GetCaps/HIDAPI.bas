Attribute VB_Name = "HIDAPI"
'�����]�p�P���-�ϥ�Visual Basic�ACH19
'API��������Ҳ��ɮ�: �C�X�Ҧ��n�Ψ쪺API�禡

'******************************************************************************
'API�`��, �ڭ̫��Ӧr�����ǥ[�H�C�X
'******************************************************************************

'�q 98ddk\inc\win98\setupapi.h���ҩw�q��
Public Const DIGCF_PRESENT = &H2
Public Const DIGCF_DEVICEINTERFACE = &H10

Public Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Public Const GENERIC_READ = &H80000000
Public Const GENERIC_WRITE = &H40000000
Public Const FILE_SHARE_READ = &H1
Public Const FILE_SHARE_WRITE = &H2

'Typedef enum �w�qHidP_Report_Type �@�վ�Ʊ`��
'�ŧi�o�Ǭ���ơ]Int�^���A (16-bits)

Public Const HidP_Input = 0
Public Const HidP_Output = 1
Public Const HidP_Feature = 2

Public Const OPEN_EXISTING = 3

'******************************************************************************
'API�I�s���ϥΪ̩w�q, �ڭ̫��Ӧr�����ǥ[�H�C�X
'******************************************************************************

Public Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type

Public Type HIDD_ATTRIBUTES
    Size As Long
    VendorID As Integer
    ProductID As Integer
    VersionNumber As Integer
End Type

'�ϥΦbhidpi.h���ҩw�q���ŧi
'HID �˸m����O

Public Type HIDP_CAPS
    Usage As Integer
    UsagePage As Integer
    InputReportByteLength As Integer
    OutputReportByteLength As Integer
    FeatureReportByteLength As Integer
    Reserved(16) As Integer
    NumberLinkCollectionNodes As Integer
    NumberInputButtonCaps As Integer
    NumberInputValueCaps As Integer
    NumberInputDataIndices As Integer
    NumberOutputButtonCaps As Integer
    NumberOutputValueCaps As Integer
    NumberOutputDataIndices As Integer
    NumberFeatureButtonCaps As Integer
    NumberFeatureValueCaps As Integer
    NumberFeatureDataIndices As Integer
End Type

'HID�ƭȯ�O���w�q�GHidP_Value_Caps
'�p�GIsRange�O����, UsageMin �OUsage, �H�� UsageMax �å��ϥ�
'�p�G IsStringRange �O����, StringMin�O�r����ޭ�,StringMax �å��ϥ�
'�p�G IsDesignatorRange �O����, DesignatorMin�Odesignator����,DesignatorMax �å��ϥ�

Public Type HidP_Value_Caps
    UsagePage As Integer
    ReportID As Byte
    IsAlias As Long
    BitField As Integer
    LinkCollection As Integer
    LinkUsage As Integer
    LinkUsagePage As Integer
    IsRange As Long
    IsStringRange As Long
    IsDesignatorRange As Long
    IsAbsolute As Long
    HasNull As Long
    Reserved As Byte
    BitSize As Integer
    ReportCount As Integer
    Reserved2 As Integer
    Reserved3 As Integer
    Reserved4 As Integer
    Reserved5 As Integer
    Reserved6 As Integer
    LogicalMin As Long
    LogicalMax As Long
    PhysicalMin As Long
    PhysicalMax As Long
    UsageMin As Integer
    UsageMax As Integer
    StringMin As Integer
    StringMax As Integer
    DesignatorMin As Integer
    DesignatorMax As Integer
    DataIndexMin As Integer
    DataIndexMax As Integer
End Type


Public Type SP_DEVICE_INTERFACE_DATA
   cbSize As Long
   InterfaceClassGuid As GUID
   Flags As Long
   Reserved As Long
End Type

Public Type SP_DEVICE_INTERFACE_DETAIL_DATA
    cbSize As Long
    DevicePath As Byte
End Type

Public Type SP_DEVINFO_DATA
    cbSize As Long
    ClassGuid As GUID
    DevInst As Long
    Reserved As Long
End Type

'******************************************************************************
'HID API ���, �ڭ̫��Ӧr�����ǥ[�H�C�X
'******************************************************************************

Public Declare Function CloseHandle _
    Lib "kernel32" _
    (ByVal hObject As Long) _
As Long

Public Declare Function CreateFile _
    Lib "kernel32" _
    Alias "CreateFileA" _
    (ByVal lpFileName As String, _
    ByVal dwDesiredAccess As Long, _
    ByVal dwShareMode As Long, _
    ByRef lpSecurityAttributes As Long, _
    ByVal dwCreationDisposition As Long, _
    ByVal dwFlagsAndAttributes As Long, _
    ByVal hTemplateFile As Long) _
As Long

Public Declare Function FormatMessage _
    Lib "kernel32" _
    Alias "FormatMessageA" _
    (ByVal dwFlags As Long, _
    ByRef lpSource As Any, _
    ByVal dwMessageId As Long, _
    ByVal dwLanguageZId As Long, _
    ByVal lpBuffer As String, _
    ByVal nSize As Long, _
    ByVal Arguments As Long) _
As Long

Public Declare Function HidD_FreePreparsedData _
    Lib "hid.dll" _
    (ByRef PreparsedData As Long) _
As Long

Public Declare Function HidD_GetAttributes _
    Lib "hid.dll" _
    (ByVal HidDeviceObject As Long, _
    ByRef Attributes As HIDD_ATTRIBUTES) _
As Long

'���F�@�P�ʡA�HAPI�禡�[�H�ŧi
'�b���A���Ǧ^����� (�i�H�����Ǧ^��)

Public Declare Function HidD_GetHidGuid _
    Lib "hid.dll" _
    (ByRef HidGuid As GUID) _
As Long


Public Declare Function HidD_GetPreparsedData _
    Lib "hid.dll" _
    (ByVal HidDeviceObject As Long, _
    ByRef PreparsedData As Long) _
As Long

Public Declare Function HidP_GetCaps _
    Lib "hid.dll" _
    (ByVal PreparsedData As Long, _
    ByRef Capabilities As HIDP_CAPS) _
As Long

Public Declare Function HidP_GetValueCaps _
    Lib "hid.dll" _
    (ByVal ReportType As Integer, _
    ByRef ValueCaps As Byte, _
    ByRef ValueCapsLength As Integer, _
    ByVal PreparsedData As Long) _
As Long
       

Public Declare Function ReadFile _
    Lib "kernel32" _
    (ByVal hFile As Long, _
    ByRef lpBuffer As Byte, _
    ByVal nNumberOfBytesToRead As Long, _
    ByRef lpNumberOfBytesRead As Long, _
    ByVal lpOverlapped As Long) _
As Long

Public Declare Function RtlMoveMemory _
    Lib "kernel32" _
    (dest As Any, _
    src As Any, _
    ByVal Count As Long) _
As Long

Public Declare Function SetupDiDestroyDeviceInfoList _
    Lib "setupapi.dll" _
    (ByVal DeviceInfoSet As Long) _
As Long

Public Declare Function SetupDiEnumDeviceInterfaces _
    Lib "setupapi.dll" _
    (ByVal DeviceInfoSet As Long, _
    ByVal DeviceInfoData As Long, _
    ByRef InterfaceClassGuid As GUID, _
    ByVal MemberIndex As Long, _
    ByRef DeviceInterfaceData As SP_DEVICE_INTERFACE_DATA) _
As Long

Public Declare Function SetupDiGetClassDevs _
    Lib "setupapi.dll" _
    Alias "SetupDiGetClassDevsA" _
    (ByRef ClassGuid As GUID, _
    ByVal Enumerator As String, _
    ByVal hwndParent As Long, _
    ByVal Flags As Long) _
As Long

Public Declare Function SetupDiGetDeviceInterfaceDetail _
   Lib "setupapi.dll" _
   Alias "SetupDiGetDeviceInterfaceDetailA" _
   (ByVal DeviceInfoSet As Long, _
   ByRef DeviceInterfaceData As SP_DEVICE_INTERFACE_DATA, _
   ByVal DeviceInterfaceDetailData As Long, _
   ByVal DeviceInterfaceDetailDataSize As Long, _
   ByRef RequiredSize As Long, _
   ByVal DeviceInfoData As Long) _
As Long
    
Public Declare Function WriteFile _
    Lib "kernel32" _
    (ByVal hFile As Long, _
    ByRef lpBuffer As Byte, _
    ByVal nNumberOfBytesToWrite As Long, _
    ByRef lpNumberOfBytesWritten As Long, _
    ByVal lpOverlapped As Long) _
As Long

Public Declare Function HidD_GetFeature _
    Lib "hid.dll" _
    (ByVal HidDeviceObject As Long, _
    ByRef ReportBuffer As Byte, _
    ByVal ReportBufferLength As Long) _
As Long

Public Declare Function HidD_SetFeature _
    Lib "hid.dll" _
    (ByVal HidDeviceObject As Long, _
    ByRef ReportBuffer As Byte, _
    ByVal ReportBufferLength As Long) _
As Long


        



 