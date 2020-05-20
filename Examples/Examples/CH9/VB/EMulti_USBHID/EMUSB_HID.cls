VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EMUSB_HID"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'�����]�p�P���-�ϥ�Visual Basic�ACH20
'DLL�ɮת��]�p�P�s�@
'�]�p�@�i�����ʺA�쵲�禡�w�ɮסA�i�H���@�ӧO��ƥ[�H��J�P��X��USB�˸m��
'�]�t FindCidPid()�AOpenDevice()�AReadDevice()�AGetDeviceCapabilities�]�^
'WriteByte()�A�PClaseByte()���禡

Option Explicit
'�w�q���ϰ��ܼ�
Dim Capabilities As HIDP_CAPS
Dim DataString As String
Dim DetailData As Long
Dim DetailDataBuffer() As Byte
Dim DeviceAttributes As HIDD_ATTRIBUTES
Dim DevicePathName As String
Dim DeviceInfoSet As Long
Dim ErrorString As String
Dim hiddevice As Long
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
Dim finddev As Boolean
Dim h As Integer
Dim VendorID As Integer
Dim ProductID As Integer
Dim seldev As Long
'Opendevice(�}�Ҹ˸m)�G
Public Static Function OpenDevice(seldev, VendorID, ProductID) As Boolean
'�w�q�ܼ�
Dim Count As Integer
Dim GUIDString As String
Dim HidGuid As GUID
Dim MemberIndex As Long
'�w�] LastDevice �M MyDeviceDetected �� Flase ��
LastDevice = False
MyDeviceDetected = False

'�I�sHidD_GetHidGuid�禡�A�N�Ȧ^��
Result = HidD_GetHidGuid(HidGuid)
'�I�sSetupDiGetClassDevs�禡�A�N�Ȧ^��
DeviceInfoSet = SetupDiGetClassDevs _
    (HidGuid, _
    vbNullString, _
    0, _
    (DIGCF_PRESENT Or DIGCF_DEVICEINTERFACE))

'�q0�}�l�W�[�A����S���˸m�Q�����C
 MemberIndex = seldev
 Do
    'MyDeviceInterfaceData��cbSize element�j�p�����Q�]�w��28 bytes�C
    MyDeviceInterfaceData.cbSize = LenB(MyDeviceInterfaceData)
    Result = SetupDiEnumDeviceInterfaces _
        (DeviceInfoSet, _
        0, _
        HidGuid, _
        MemberIndex, _
        MyDeviceInterfaceData)
 
 
 
    If Result = 0 Then LastDevice = True
    
    '�Y�˸m�Q�����h��ܪ�^���T���C
    If Result <> 0 Then
        
        Result = SetupDiGetDeviceInterfaceDetail _
           (DeviceInfoSet, _
           MyDeviceInterfaceData, _
           0, _
           0, _
           Needed, _
           0)
        
        DetailData = Needed
            
        
        '�x�s���c���j�p�C
        MyDeviceInterfaceDetailData.cbSize = _
            Len(MyDeviceInterfaceDetailData)
        
        '�ϥΰ}�C�Ӥ��tMyDeviceInterfaceDetailData structure�O����C
        ReDim DetailDataBuffer(Needed)
        '�x�scbSize�b�e�|��bytes���}�C�C
        Call RtlMoveMemory _
            (DetailDataBuffer(0), _
            MyDeviceInterfaceDetailData, _
            4)
        
        '�A���I�sSetupDiGetDeviceInterfaceDetail
        '�o���A�ǰeDetailDataBuffer���Ĥ@�ӭȩҦs�񪺦�}�C
        '���K�^�ǩһ�buffer���j�p��DetailData
        Result = SetupDiGetDeviceInterfaceDetail _
           (DeviceInfoSet, _
           MyDeviceInterfaceData, _
           VarPtr(DetailDataBuffer(0)), _
           DetailData, _
           Needed, _
           0)
           
        'byte��r��C
        DevicePathName = CStr(DetailDataBuffer())
        '�ഫUnicode
        DevicePathName = StrConv(DevicePathName, vbUnicode)
        'cbSize���_�l�B�}�l�M���C
        DevicePathName = Right$(DevicePathName, Len(DevicePathName) - 4)
        '�I�sCreateFile�禡�A�N�Ȧ^��
        hiddevice = CreateFile _
            (DevicePathName, _
            GENERIC_READ Or GENERIC_WRITE, _
            (FILE_SHARE_READ Or FILE_SHARE_WRITE), _
            0, _
            OPEN_EXISTING, _
            0, _
            0)
        '�I�sHidD_GetAttributes�禡�A�N�Ȧ^��
        Result = HidD_GetAttributes _
            (hiddevice, _
            DeviceAttributes)
            
   
        '�ݬݳo�˸m���S���ŦX�ڭ̷Q�n���C
        
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
Else
    OpenDevice = False
End If
 
End Function
'Closedevice(�����˸m)�G
Public Static Function CloseDevice()
'�I�sCloseHandle�禡�A�N�ȶǵ�Reasult
Result = CloseHandle _
    (hiddevice)
'�I�sSetupDiDestroyDeviceInfoList�禡�A�N�ȶǵ�Reasult
Result = SetupDiDestroyDeviceInfoList _
    (DeviceInfoSet)
'�I�sHidD_FreePreparsedData�禡�A�N�ȶǵ�Reasult
Result = HidD_FreePreparsedData _
    (PreparsedData)
End Function
'Findhiddevice(�M��˸m)�G
Public Static Function FindHidDevice()
'�w�q�ܼƫ��A
Dim HidGuid As GUID
Dim MemberIndex As Long
Dim dev1 As Long
Dim i As Long
'�I�sHidD_GetHidGuid�禡�A�N�ȶǵ�Reasult
Result = HidD_GetHidGuid(HidGuid)
'�I�sSetupDiGetClassDevs�禡�A�N�ȶǵ�DeviceInfoSet
DeviceInfoSet = SetupDiGetClassDevs _
    (HidGuid, _
    vbNullString, _
    0, _
    (DIGCF_PRESENT Or DIGCF_DEVICEINTERFACE))
    
 i = 0
 '
For dev1 = 0 To 127
 MemberIndex = dev1
MyDeviceInterfaceData.cbSize = LenB(MyDeviceInterfaceData)
 '�I�sSetupDiEnumDeviceInterfaces�禡�A�N�ȶǵ�Reasult
 Result = SetupDiEnumDeviceInterfaces _
        (DeviceInfoSet, _
        0, _
        HidGuid, _
        MemberIndex, _
        MyDeviceInterfaceData)
 '�p�GResult������0�A�hi=i+1
 If Result <> 0 Then
 i = i + 1
 End If
 Next
 FindHidDevice = i
End Function
'GetDeviceCapabilities(���o�˸m��O)�G
Private Sub GetDeviceCapabilities()
'�w�q�ܼƫ��A
Dim ppData(29) As Byte
Dim ppDataString As Variant

'�I�sHidD_GetPreparsedData�禡�A�N�ȶǵ�Reasult
Result = HidD_GetPreparsedData _
    (hiddevice, _
    PreparsedData)

'�I�sRtlMoveMemory�禡�A�N�ȶǵ�Reasult
Result = RtlMoveMemory _
    (ppData(0), _
    PreparsedData, _
    30)

'�I�sHidP_GetCaps�禡�A�N�ȶǵ�Reasult
Result = HidP_GetCaps _
    (PreparsedData, _
    Capabilities)


Dim ValueCaps(1024) As Byte
'�I�sHidP_GetValueCaps�禡�A�N�ȶǵ�Reasult
Result = HidP_GetValueCaps _
    (HidP_Input, _
    ValueCaps(0), _
    Capabilities.NumberInputValueCaps, _
    PreparsedData)
    
End Sub
Function ReadByte(SeleByte) As String

'Read data from the device.

Dim Count
Dim NumberOfBytesRead As Long
Dim ReadBuffer() As Byte
Dim UBoundReadBuffer As Integer
Dim ByteValue(7) As String


'ReadBuffer�}�C�O�H0���ҩl��, �]���ݱN�줸�ժ��ƥش�h1
ReDim ReadBuffer(Capabilities.FeatureReportByteLength - 1)

NumberOfBytesRead = Capabilities.FeatureReportByteLength
'��Ū���w�İϪ����w���줸�ժ���}
'Result = ReadFile _
'    (HidDevice, _
'    ReadBuffer(0), _
'    CLng(Capabilities.InputReportByteLength), _
'    NumberOfBytesRead, _
'    0)
Result = HidD_GetFeature _
    (hiddevice, _
    ReadBuffer(0), _
    NumberOfBytesRead)


  '�N���ন16�i��
For Count = 1 To UBound(ReadBuffer)
    If Len(Hex$(ReadBuffer(Count))) < 2 Then
        ByteValue(Count - 1) = "0" & Hex$(ReadBuffer(Count))
    Else
        ByteValue(Count - 1) = Hex$(ReadBuffer(Count))
    End If
Next Count

'�Ȱw��ҭnŪ�����ĴX�������ƭȡA�[�HŪ���öǦ^
    ReadByte = ByteValue(SeleByte - 1)
    

End Function 'Readdevice(Ū���˸m���)�G
Public Static Function ReadDevice() As String()
'�w�q�ܼƫ��A
Dim Count
Dim NumberOfBytesRead As Long
Dim ReadBuffer() As Byte
Dim UBoundReadBuffer As Integer
Dim ByteValue(7) As String


'�]ReadBuffer�}�C�q0�}�l�A�ҥH�ȴ�1
ReDim ReadBuffer(Capabilities.FeatureReportByteLength - 1)
    
NumberOfBytesRead = Capabilities.FeatureReportByteLength
'�NReadBuffer���Ĥ@�줸�զs���m�A�A�I�sReadFile�禡�A�N�ȶǦ^Reasult
'Result = ReadFile _
'    (HidDevice, _
'    ReadBuffer(0), _
'    CLng(Capabilities.InputReportByteLength), _
'    NumberOfBytesRead, _
'    0)
Result = HidD_GetFeature _
    (hiddevice, _
    ReadBuffer(0), _
    NumberOfBytesRead)

'�N���ন16�i��
For Count = 1 To UBound(ReadBuffer)
    If Len(Hex$(ReadBuffer(Count))) < 2 Then
        ByteValue(Count - 1) = "0" & Hex$(ReadBuffer(Count))
    Else
        ByteValue(Count - 1) = Hex$(ReadBuffer(Count))
    End If
Next Count

ReadDevice = ByteValue
End Function
'Writebyte(��ܶǰe�������)�G
Public Static Function WriteByte(senddata() As Byte, SeleByte As Integer)
'�w�q�ܼƫ��A
Dim NumberOfBytesToSend As Long
Dim NumberOfBytesWritten As Long
Dim SendBuffer() As Byte
Dim Count As Integer
Dim NumberOfBytesRead As Long
'Dim ReadBuffer() As Byte

'�]SendBuffer�}�C�q0�}�l�A�ҥH�ȴ�1
ReDim SendBuffer(Capabilities.FeatureReportByteLength - 1)

'�Ĥ@�Ӧ줸�լ�Report ID�A���᪺�줸�լ����
SendBuffer(0) = 0
'�Nsenddata���Ȱe�JSendBuffer�A�[�Jcontrolwrite�ӼW�[��ܸ�ƿﶵ
'�Ycontrolwrite��0�A�h�ǰe��1���A�Y��n�A�ǰe��n��
For Count = 1 To Capabilities.FeatureReportByteLength - 1
    If Count = SeleByte Then
    SendBuffer(Count) = senddata(SeleByte - 1)
    Else
    SendBuffer(Count) = 0
    End If
Next Count

NumberOfBytesWritten = Capabilities.FeatureReportByteLength
'�I�sWriteFile�禡�A�N�Ȧ^��
'Result = WriteFile _
'    (HidDevice, _
'    SendBuffer(0), _
'    CLng(Capabilities.OutputReportByteLength), _
'    NumberOfBytesWritten, _
'    0)
Result = HidD_SetFeature _
    (hiddevice, _
    SendBuffer(0), _
    NumberOfBytesWritten)

ReDim ReadBuffer(Capabilities.FeatureReportByteLength - 1)
End Function
'Writedevice(�ǰe���)�G
Public Static Function WriteDevice(senddata() As Byte)
'�w�q�ܼƫ��A
Dim NumberOfBytesToSend As Long
Dim NumberOfBytesWritten As Long
Dim SendBuffer() As Byte
Dim Count As Integer
Dim NumberOfBytesRead As Long
Dim ReadBuffer() As Byte

'�]SendBuffer�}�C�q0�}�l�A�ҥH�ȴ�1
ReDim SendBuffer(Capabilities.FeatureReportByteLength - 1)

'�Ĥ@�Ӧ줸�լ�Report ID�A���᪺�줸�լ����
SendBuffer(0) = 0

'�bSendBuffer()�w�İϤ��A�ȱN�n�]�w���ĴX�����줸�ձN�H�g�J�A��l��J0
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
    (hiddevice, _
    SendBuffer(0), _
    NumberOfBytesWritten)

ReDim ReadBuffer(Capabilities.FeatureReportByteLength - 1)
End Function


