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
'介面設計與實習-使用Visual Basic，CH20
'DLL檔案的設計與製作
'設計一進階的動態鏈結函式庫檔案，可以對單一個別資料加以輸入與輸出至USB裝置中
'包含 FindCidPid()，OpenDevice()，ReadDevice()，GetDeviceCapabilities（）
'WriteByte()，與ClaseByte()等函式

Option Explicit
'定義全區域變數
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
'Opendevice(開啟裝置)：
Public Static Function OpenDevice(seldev, VendorID, ProductID) As Boolean
'定義變數
Dim Count As Integer
Dim GUIDString As String
Dim HidGuid As GUID
Dim MemberIndex As Long
'預設 LastDevice 和 MyDeviceDetected 為 Flase 值
LastDevice = False
MyDeviceDetected = False

'呼叫HidD_GetHidGuid函式，將值回傳
Result = HidD_GetHidGuid(HidGuid)
'呼叫SetupDiGetClassDevs函式，將值回傳
DeviceInfoSet = SetupDiGetClassDevs _
    (HidGuid, _
    vbNullString, _
    0, _
    (DIGCF_PRESENT Or DIGCF_DEVICEINTERFACE))

'從0開始增加，直到沒有裝置被偵察到。
 MemberIndex = seldev
 Do
    'MyDeviceInterfaceData的cbSize element大小必須被設定成28 bytes。
    MyDeviceInterfaceData.cbSize = LenB(MyDeviceInterfaceData)
    Result = SetupDiEnumDeviceInterfaces _
        (DeviceInfoSet, _
        0, _
        HidGuid, _
        MemberIndex, _
        MyDeviceInterfaceData)
 
 
 
    If Result = 0 Then LastDevice = True
    
    '若裝置被移除則顯示返回的訊息。
    If Result <> 0 Then
        
        Result = SetupDiGetDeviceInterfaceDetail _
           (DeviceInfoSet, _
           MyDeviceInterfaceData, _
           0, _
           0, _
           Needed, _
           0)
        
        DetailData = Needed
            
        
        '儲存結構的大小。
        MyDeviceInterfaceDetailData.cbSize = _
            Len(MyDeviceInterfaceDetailData)
        
        '使用陣列來分配MyDeviceInterfaceDetailData structure記憶體。
        ReDim DetailDataBuffer(Needed)
        '儲存cbSize在前四個bytes的陣列。
        Call RtlMoveMemory _
            (DetailDataBuffer(0), _
            MyDeviceInterfaceDetailData, _
            4)
        
        '再次呼叫SetupDiGetDeviceInterfaceDetail
        '這次，傳送DetailDataBuffer的第一個值所存放的位址。
        '順便回傳所需buffer的大小到DetailData
        Result = SetupDiGetDeviceInterfaceDetail _
           (DeviceInfoSet, _
           MyDeviceInterfaceData, _
           VarPtr(DetailDataBuffer(0)), _
           DetailData, _
           Needed, _
           0)
           
        'byte轉字串。
        DevicePathName = CStr(DetailDataBuffer())
        '轉換Unicode
        DevicePathName = StrConv(DevicePathName, vbUnicode)
        'cbSize的起始處開始清除。
        DevicePathName = Right$(DevicePathName, Len(DevicePathName) - 4)
        '呼叫CreateFile函式，將值回傳
        hiddevice = CreateFile _
            (DevicePathName, _
            GENERIC_READ Or GENERIC_WRITE, _
            (FILE_SHARE_READ Or FILE_SHARE_WRITE), _
            0, _
            OPEN_EXISTING, _
            0, _
            0)
        '呼叫HidD_GetAttributes函式，將值回傳
        Result = HidD_GetAttributes _
            (hiddevice, _
            DeviceAttributes)
            
   
        '看看這裝置有沒有符合我們想要的。
        
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
'Closedevice(結束裝置)：
Public Static Function CloseDevice()
'呼叫CloseHandle函式，將值傳給Reasult
Result = CloseHandle _
    (hiddevice)
'呼叫SetupDiDestroyDeviceInfoList函式，將值傳給Reasult
Result = SetupDiDestroyDeviceInfoList _
    (DeviceInfoSet)
'呼叫HidD_FreePreparsedData函式，將值傳給Reasult
Result = HidD_FreePreparsedData _
    (PreparsedData)
End Function
'Findhiddevice(尋找裝置)：
Public Static Function FindHidDevice()
'定義變數型態
Dim HidGuid As GUID
Dim MemberIndex As Long
Dim dev1 As Long
Dim i As Long
'呼叫HidD_GetHidGuid函式，將值傳給Reasult
Result = HidD_GetHidGuid(HidGuid)
'呼叫SetupDiGetClassDevs函式，將值傳給DeviceInfoSet
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
 '呼叫SetupDiEnumDeviceInterfaces函式，將值傳給Reasult
 Result = SetupDiEnumDeviceInterfaces _
        (DeviceInfoSet, _
        0, _
        HidGuid, _
        MemberIndex, _
        MyDeviceInterfaceData)
 '如果Result不等於0，則i=i+1
 If Result <> 0 Then
 i = i + 1
 End If
 Next
 FindHidDevice = i
End Function
'GetDeviceCapabilities(取得裝置能力)：
Private Sub GetDeviceCapabilities()
'定義變數型態
Dim ppData(29) As Byte
Dim ppDataString As Variant

'呼叫HidD_GetPreparsedData函式，將值傳給Reasult
Result = HidD_GetPreparsedData _
    (hiddevice, _
    PreparsedData)

'呼叫RtlMoveMemory函式，將值傳給Reasult
Result = RtlMoveMemory _
    (ppData(0), _
    PreparsedData, _
    30)

'呼叫HidP_GetCaps函式，將值傳給Reasult
Result = HidP_GetCaps _
    (PreparsedData, _
    Capabilities)


Dim ValueCaps(1024) As Byte
'呼叫HidP_GetValueCaps函式，將值傳給Reasult
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


'ReadBuffer陣列是以0為啟始的, 因此需將位元組的數目減去1
ReDim ReadBuffer(Capabilities.FeatureReportByteLength - 1)

NumberOfBytesRead = Capabilities.FeatureReportByteLength
'傳讀取緩衝區的指定的位元組的位址
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


  '將值轉成16進制
For Count = 1 To UBound(ReadBuffer)
    If Len(Hex$(ReadBuffer(Count))) < 2 Then
        ByteValue(Count - 1) = "0" & Hex$(ReadBuffer(Count))
    Else
        ByteValue(Count - 1) = Hex$(ReadBuffer(Count))
    End If
Next Count

'僅針對所要讀取的第幾筆中的數值，加以讀取並傳回
    ReadByte = ByteValue(SeleByte - 1)
    

End Function 'Readdevice(讀取裝置資料)：
Public Static Function ReadDevice() As String()
'定義變數型態
Dim Count
Dim NumberOfBytesRead As Long
Dim ReadBuffer() As Byte
Dim UBoundReadBuffer As Integer
Dim ByteValue(7) As String


'因ReadBuffer陣列從0開始，所以值減1
ReDim ReadBuffer(Capabilities.FeatureReportByteLength - 1)
    
NumberOfBytesRead = Capabilities.FeatureReportByteLength
'將ReadBuffer的第一位元組存放位置，再呼叫ReadFile函式，將值傳回Reasult
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

'將值轉成16進制
For Count = 1 To UBound(ReadBuffer)
    If Len(Hex$(ReadBuffer(Count))) < 2 Then
        ByteValue(Count - 1) = "0" & Hex$(ReadBuffer(Count))
    Else
        ByteValue(Count - 1) = Hex$(ReadBuffer(Count))
    End If
Next Count

ReadDevice = ByteValue
End Function
'Writebyte(選擇傳送那筆資料)：
Public Static Function WriteByte(senddata() As Byte, SeleByte As Integer)
'定義變數型態
Dim NumberOfBytesToSend As Long
Dim NumberOfBytesWritten As Long
Dim SendBuffer() As Byte
Dim Count As Integer
Dim NumberOfBytesRead As Long
'Dim ReadBuffer() As Byte

'因SendBuffer陣列從0開始，所以值減1
ReDim SendBuffer(Capabilities.FeatureReportByteLength - 1)

'第一個位元組為Report ID，之後的位元組為資料
SendBuffer(0) = 0
'將senddata的值送入SendBuffer，加入controlwrite來增加選擇資料選項
'若controlwrite為0，則傳送第1筆，若為n，傳送第n筆
For Count = 1 To Capabilities.FeatureReportByteLength - 1
    If Count = SeleByte Then
    SendBuffer(Count) = senddata(SeleByte - 1)
    Else
    SendBuffer(Count) = 0
    End If
Next Count

NumberOfBytesWritten = Capabilities.FeatureReportByteLength
'呼叫WriteFile函式，將值回傳
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
'Writedevice(傳送資料)：
Public Static Function WriteDevice(senddata() As Byte)
'定義變數型態
Dim NumberOfBytesToSend As Long
Dim NumberOfBytesWritten As Long
Dim SendBuffer() As Byte
Dim Count As Integer
Dim NumberOfBytesRead As Long
Dim ReadBuffer() As Byte

'因SendBuffer陣列從0開始，所以值減1
ReDim SendBuffer(Capabilities.FeatureReportByteLength - 1)

'第一個位元組為Report ID，之後的位元組為資料
SendBuffer(0) = 0

'在SendBuffer()緩衝區中，僅將要設定的第幾筆的位元組將以寫入，其餘填入0
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



