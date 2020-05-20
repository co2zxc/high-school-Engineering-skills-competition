VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "USB_HID"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'�����]�p�P���-�ϥ�Visual Basic�ACH20
 'DLL�ɮת��]�p�P�s�@
'�]�p�@�򥻪��ʺA�쵲�禡�w�ɮ�
'�]�t OpenDevice()�AReadDevice()�AGetDeviceCapabilities�]�^�AWriteDevice()�A
'ClaseDevice()���禡

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
Dim PreparsedData As Long
Dim OutputReportData(7) As Byte
Dim Result As Long
Dim Timeout As Boolean
'Dim finddev As Boolean
'Dim h As Integer
Dim VendorID As Integer
Dim ProductID As Integer

'���լO�_��usbhid�˸m
Public Static Function OpenDevice(VendorID, ProductID) As Boolean
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

''******************************************************************************
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
 
 

 
    If Result = 0 Then LastDevice = True
    
    '�p�G�˸m�s�b���ܡA��ܦ^�Ǫ��T��
    If Result <> 0 Then
    
'******************************************************************************
'SetupDiGetDeviceInterfaceDetail�禡
'�^�ǡG�]�t���˸m�T����SP_DEVICE_INTERFACE_DETAIL_DATA���c
'�`�N�ƶ��G���F���^�������T���A�I�s���禡�⦸
'�Ĥ@���Ǧ^�һݪ����c�j�p
'�ĤG���Ǧ^�bDeviceInfoSet��ƪ����о�
'�ݨD�G��SetupDiGetClassDevs�Ҧ^�Ǫ�DeviceInfoSet
 '�H�Υ�SetupDiEnumDeviceInterfaces�Ҧ^�Ǫ�SP_DEVICE_INTERFACE_DATA���c
'*******************************************************************************
        
        Result = SetupDiGetDeviceInterfaceDetail _
           (DeviceInfoSet, _
           MyDeviceInterfaceData, _
           0, _
           0, _
           Needed, _
           0)
        
        DetailData = Needed
            
        
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
           
        '�N�줸�հ}�C�ഫ���r��
        DevicePathName = CStr(DetailDataBuffer())
        
        '�ഫ��Unicode�X�榡.
        DevicePathName = StrConv(DevicePathName, vbUnicode)
        
        '�q�}�l�B�簣cbSize (4 bytes).
        DevicePathName = Right$(DevicePathName, Len(DevicePathName) - 4)
                        
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

        Result = HidD_GetAttributes _
            (HidDevice, _
            DeviceAttributes)
            
    
        
      '��X�O�_�o�Ӹ˸m�P�ڭ̥��b�M�䪺�˸m�k�X
        
        If (DeviceAttributes.VendorID = VendorID) And _
            (DeviceAttributes.ProductID = ProductID) Then
                MyDeviceDetected = True
        Else
                MyDeviceDetected = False

        End If
End If
    

Loop Until (LastDevice = True) Or (MyDeviceDetected = True)

If MyDeviceDetected = True Then
    OpenDevice = True
    Call GetDeviceCapabilities
    ''lstResults.AddItem " Device  found."
Else
    ''lstResults.AddItem " Device not found."
    OpenDevice = False
End If
 
End Function


'�o��˸m�\��T��
Private Sub GetDeviceCapabilities()

'******************************************************************************
'HidD_GetPreparsedData
'�^�ǡG�]�t���V�@�����󦹸˸m��O�T�����w�İϪ����о�
'�ݨD�G��CreateFile�Ҧ^�Ǫ�handle
'�`�N�ƶ��G�b�����ݪ����B�z�w�İ�
'��HidP_GetCaps�P��L��API�禡�ݭn�@�ӫ��V���w�İϪ����о�
'*******************************************************************************

Dim ppData(29) As Byte
Dim ppDataString As Variant


'�Q��R����ơ]PreparsedData�^���о��O���V�禡�ҫŧi���w�İ�
Result = HidD_GetPreparsedData _
    (HidDevice, _
    PreparsedData)

'�N���PreparsedData����ƫ�����줸�հ}�C

Result = RtlMoveMemory _
    (ppData(0), _
    PreparsedData, _
    30)


'******************************************************************************
'HidP_GetCaps
'�D�X�˸m����O
'���з�HID���˸m�A�p��L�άO�n�쵥�A���i�H��즹�˸m�S�w����O
'�ӹ��ۭq���˸m�Athe software will probably know what the device is capable of,
'so this call only verifies the information.
'�ݨD�G�[�\�o�ǰT�����w�İϪ�����
'�o�ӫ��о��O��HidD_GetPreparsedData�禡�Ҧ^�Ǫ�.
'�^��: �]�t�����T������O�����c
'******************************************************************************
Result = HidP_GetCaps _
    (PreparsedData, _
    Capabilities)
    
'******************************************************************************
'HidP_GetValueCaps
'�^�ǡG�Ǧ^�@�ӥ]�t�FHidP_ValueCaps���c�}�C���w�İ�
'�`�N�ƶ��G�C�@�ӵ��c�w�q�F�@�ӼƭȪ���O
'�o�����ε{���ä��ϥΦ����
'******************************************************************************

'���줸�հ}�C�O�����b�o���c��
Dim ValueCaps(1024) As Byte

Result = HidP_GetValueCaps _
    (HidP_Input, _
    ValueCaps(0), _
    Capabilities.NumberInputValueCaps, _
    PreparsedData)
    

'���F�ϥΦ���ơA�N�줸�հ}�C�����쵲�c�}�C��
End Sub

'�������
Public Static Function ReadDevice() As Variant
'�q�˸mŪ�����
Dim Count
Dim NumberOfBytesRead As Long

'�ŧi�����i�Ϊ��w�İϡA'Byte 0�O���iID�A�ҥH�ݴ�h1
Dim ReadBuffer() As Byte
Dim UBoundReadBuffer As Integer
Dim ByteValue(8) As String
Dim recevicedata As Variant
'******************************************************************************
'HidD_GetFeature
'�Ǧ^: �bReadBuffer�����i�y�z��.
'�ݨD: ��CreateFile�ҶǦ^���˸mhandle,
'�`�N�ƶ��G�o�S�ʳ��i�����׬O��HidP_GetCaps�禡�Ҧ^�Ǫ��A�B�H�줸�լ����
'******************************************************************************

'ReadFile�O�϶��I�s
'�o�����ε{���|"��"���A����˸m�e�X�һݪ���ƶq����
'���F�קK"��"�F�A�ڭ̥����T�w�˸m�`�O����ƨӰe�X

ReDim ReadBuffer(Capabilities.FeatureReportByteLength - 1)
NumberOfBytesRead = Capabilities.FeatureReportByteLength
'ReadBuffer�}�C�O�H0���ҩl��, �]���ݱN�줸�ժ��ƥش�h1
'Result = ReadFile _
'    (HidDevice, _
'    ReadBuffer(0), _
'    CLng(Capabilities.InputReportByteLength), _
'    NumberOfBytesRead, _
'    0)
Result = HidD_GetFeature _
    (HidDevice, _
    ReadBuffer(0), _
    NumberOfBytesRead)

For Count = 1 To UBound(ReadBuffer)
     '�W�[�e�ɬ�0��0 - Fh�ƭȫe
    If Len(Hex$(ReadBuffer(Count))) < 2 Then
        ByteValue(Count - 1) = "0" & Hex$(ReadBuffer(Count))
    Else
        ByteValue(Count - 1) = Hex$(ReadBuffer(Count))
    End If

    
Next Count
recevicedata = ByteValue()
ReadDevice = recevicedata
End Function


Public Static Function WriteDevice(senddata() As Byte)
'�ǰe���

Dim NumberOfBytesToSend As Long
Dim NumberOfBytesWritten As Long
Dim SendBuffer() As Byte
Dim Count As Integer
Dim NumberOfBytesRead As Long
Dim ReadBuffer() As Byte

'�o��SendBuffer�}�C�O�q0�}�l��,�ҥH�ݴ�h1�Ӧ줸�ժ��ƥ�
ReDim SendBuffer(Capabilities.FeatureReportByteLength - 1)

'******************************************************************************
'HidD_SetFeature�禡
'�e�X���i���˸m�ϥ�
'�Ǧ^�G���\�άO����
'�ݨD�G��CreatFile�禡�Ҧ^�Ǫ�handle
'��HidP_GetCaps�禡�Ǧ^�S�ʳ��i������
'******************************************************************************

'�Ĥ@�Ӧ줸�լOReport ID
SendBuffer(0) = 0

'�U�@�Ӧ줸�լO���������
For Count = 1 To Capabilities.FeatureReportByteLength - 1
    SendBuffer(Count) = senddata(Count - 1)
Next Count
NumberOfBytesWritten = Capabilities.FeatureReportByteLength
''SendBuffer(1) = senddata
'Result = WriteFile _
'    (HidDevice, _
'    SendBuffer(0), _
'    CLng(Capabilities.OutputReportByteLength), _
'    NumberOfBytesWritten, _
'    0)
Result = HidD_SetFeature _
    (HidDevice, _
    SendBuffer(0), _
    NumberOfBytesWritten)
    
ReDim ReadBuffer(Capabilities.FeatureReportByteLength - 1)
End Function


Public Static Function CloseDevice()
Result = CloseHandle _
    (HidDevice)
Result = SetupDiDestroyDeviceInfoList _
    (DeviceInfoSet)
Result = HidD_FreePreparsedData _
    (PreparsedData)
End Function



