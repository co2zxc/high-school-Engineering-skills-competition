VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MUSB_HID"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'介面設計與實習-使用Visual Basic，CH20
 'DLL檔案的設計與製作
'設計一進階的動態鏈結函式庫檔案，可以對多個USB HID裝置同時加以控制
'包含 FindCidPid()，OpenDevice()，ReadDevice()，GetDeviceCapabilities（）
'WriteDevice()，與ClaseDevice()等函式

Option Explicit
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
Dim SelecDev As Long
'取得Vid與Pid碼
Public Static Function FindVidPid(SelecDev) As String
'執行一系列的API函式呼叫，以找到所需的HID群組裝置
'如果偵測到HID裝置的話，就傳回真值，反之就傳回假值
Dim Count As Integer
Dim GUIDString As String
Dim HidGuid As GUID
Dim MemberIndex As Long
Dim c, pid, vid
LastDevice = False
MyDeviceDetected = False


Result = HidD_GetHidGuid(HidGuid)


DeviceInfoSet = SetupDiGetClassDevs(HidGuid, vbNullString, 0, (DIGCF_PRESENT Or DIGCF_DEVICEINTERFACE))

'******************************************************************************
'SetupDiEnumDeviceInterfaces
'On return, MyDeviceInterfaceData contains the handle to a
'SP_DEVICE_INTERFACE_DATA structure for a detected device.
'Requires:
'the DeviceInfoSet returned in SetupDiGetClassDevs.
'the HidGuid returned in GetHidGuid.
'An index to specify a device.
'******************************************************************************

'Begin with 0 and increment until no more devices are detected.
 MemberIndex = SelecDev

    'The cbSize element of the MyDeviceInterfaceData structure must be set to
    'the structure's size in bytes. The size is 28 bytes.
    MyDeviceInterfaceData.cbSize = LenB(MyDeviceInterfaceData)
    
    Result = SetupDiEnumDeviceInterfaces(DeviceInfoSet, 0, HidGuid, MemberIndex, MyDeviceInterfaceData)
    '此一函式主要是辨識每一個hid的介面，而設定外部輸入的變數seldev來判斷所要偵測的裝置是否為hid介面。進而加以控制
 

 
    If Result = 0 Then LastDevice = True
    
    'If a device exists, display the information returned.
    If Result <> 0 Then
    
        '******************************************************************************
        'SetupDiGetDeviceInterfaceDetail
        'Returns: an SP_DEVICE_INTERFACE_DETAIL_DATA structure
        'containing information about a device.
        'To retrieve the information, call this function twice.
        'The first time returns the size of the structure in Needed.
        'The second time returns a pointer to the data in DeviceInfoSet.
        'Requires:
        'A DeviceInfoSet returned by SetupDiGetClassDevs and
        'an SP_DEVICE_INTERFACE_DATA structure returned by SetupDiEnumDeviceInterfaces
        '*******************************************************************************
        
        Result = SetupDiGetDeviceInterfaceDetail(DeviceInfoSet, MyDeviceInterfaceData, 0, 0, Needed, 0)
        
        DetailData = Needed
            
        
        'Store the structure's size.
        MyDeviceInterfaceDetailData.cbSize = _
            Len(MyDeviceInterfaceDetailData)
        
        'Use a byte array to allocate memory for
        'the MyDeviceInterfaceDetailData structure
        ReDim DetailDataBuffer(Needed)
        'Store cbSize in the first four bytes of the array.
        Call RtlMoveMemory _
            (DetailDataBuffer(0), _
            MyDeviceInterfaceDetailData, _
            4)
        
        'Call SetupDiGetDeviceInterfaceDetail again.
        'This time, pass the address of the first element of DetailDataBuffer
        'and the returned required buffer size in DetailData.
        Result = SetupDiGetDeviceInterfaceDetail _
           (DeviceInfoSet, _
           MyDeviceInterfaceData, _
           VarPtr(DetailDataBuffer(0)), _
           DetailData, _
           Needed, _
           0)
           
        'Convert the byte array to a string.
        DevicePathName = CStr(DetailDataBuffer())
        'Convert to Unicode.
        DevicePathName = StrConv(DevicePathName, vbUnicode)
        'Strip cbSize (4 bytes) from the beginning.
        DevicePathName = Right$(DevicePathName, Len(DevicePathName) - 4)
        FindVidPid = DevicePathName
        DevicePathName = Right$(DevicePathName, Len(DevicePathName) - 4)
        c = Split(DevicePathName, "pid_", 2, vbBinaryCompare)
        pid = Left(c(1), 4)
        c = Split(DevicePathName, "vid_", 2, vbBinaryCompare)
        vid = Left(c(1), 4)
        FindVidPid = CStr(vid & pid)
    End If
End Function



'測試是否有usbhid裝置
Public Static Function OpenDevice(SelecDev, VendorID, ProductID) As Boolean
'Uses API calls to locate the desired HID-class device.
'Returns True if the device is detected, False if not detected.
Dim Count As Integer
Dim GUIDString As String
Dim HidGuid As GUID
Dim MemberIndex As Long
LastDevice = False
MyDeviceDetected = False

'******************************************************************************
'HidD_GetHidGuid函式
'取得所有連接上的HID裝置的GUID
'傳回值: HidGuid內的GUID
'這個函式不會回傳Result數值
'但是這個函式會以函式的方式來加以宣告，以使得可以與其他的API函式相容
'******************************************************************************
Result = HidD_GetHidGuid(HidGuid)

'******************************************************************************
'SetupDiGetClassDevs
'回傳： 針對已安裝的裝置回傳一個指向裝置訊息集的handle
'需求：由GetHidGuid所回傳的HidGuid
'******************************************************************************

DeviceInfoSet = SetupDiGetClassDevs(HidGuid, vbNullString, 0, (DIGCF_PRESENT Or DIGCF_DEVICEINTERFACE))

''******************************************************************************
'SetupDiEnumDeviceInterfaces函式
'回傳： 包含了指向一個所偵測到裝置的SP_DEVICE_INTERFACE_DATA結構的handle
'需求：
'1. 從SetupDiGetClassDevs函式所回傳的DeviceInfoSet
'2. 從GetHidGuid所回傳的HidGuid
'3. 設定此裝置的索引值
'******************************************************************************

'以0開始，並加以遞增直到沒有裝置被偵測到為止
 MemberIndex = SelecDev
 Do
   '在MyDeviceInterfaceData結構中的cbSize元素必須被設定
    '這個結構是以位元組為計數數值，這個大小值是28個位元組
    MyDeviceInterfaceData.cbSize = LenB(MyDeviceInterfaceData)
    
    Result = SetupDiEnumDeviceInterfaces(DeviceInfoSet, 0, HidGuid, MemberIndex, MyDeviceInterfaceData)
    '此一函式主要是辨識每一個hid的介面，而設定外部輸入的變數seldev來判斷所要偵測的裝置是否為hid介面。進而加以控制
 

 
    If Result = 0 Then LastDevice = True
    
   '如果裝置存在的話，顯示回傳的訊息
    If Result <> 0 Then
    
  '******************************************************************************
'SetupDiGetDeviceInterfaceDetail函式
'回傳：包含此裝置訊息的SP_DEVICE_INTERFACE_DETAIL_DATA結構
'注意事項：為了取回相關的訊息，呼叫此函式兩次
'第一次傳回所需的結構大小
'第二次傳回在DeviceInfoSet資料的指標器
'需求：由SetupDiGetClassDevs所回傳的DeviceInfoSet
 '以及由SetupDiEnumDeviceInterfaces所回傳的SP_DEVICE_INTERFACE_DATA結構
'*******************************************************************************
        
        Result = SetupDiGetDeviceInterfaceDetail(DeviceInfoSet, MyDeviceInterfaceData, 0, 0, Needed, 0)
        
        DetailData = Needed
            
        
        '儲存這個結構的大小值
        MyDeviceInterfaceDetailData.cbSize = _
            Len(MyDeviceInterfaceDetailData)
        
       '使用位元組陣列去宣告MyDeviceInterfaceDetailData結構的記憶體
        ReDim DetailDataBuffer(Needed)
        
        '儲存在陣列的前四個位元組，cbSize
        Call RtlMoveMemory _
            (DetailDataBuffer(0), _
            MyDeviceInterfaceDetailData, _
            4)
        
        '再一次呼叫SetupDiGetDeviceInterfaceDetail函式
        '這一此，是傳遞DetailDataBuffer第一個元素的位址
        '而此回傳的是在DetailData所需的緩衝區大小
        Result = SetupDiGetDeviceInterfaceDetail _
           (DeviceInfoSet, _
           MyDeviceInterfaceData, _
           VarPtr(DetailDataBuffer(0)), _
           DetailData, _
           Needed, _
           0)
           
         '將位元組陣列轉換成字串
        DevicePathName = CStr(DetailDataBuffer())
        
        '轉換成Unicode碼格式.
        DevicePathName = StrConv(DevicePathName, vbUnicode)
        
        '從開始處剔除cbSize (4 bytes).
        DevicePathName = Right$(DevicePathName, Len(DevicePathName) - 4)
                        
        '******************************************************************************
        'CreateFile函式
        '回傳：
        '回傳用來致能讀取與寫入此裝置的handle
        '需求:
        '由SetupDiGetDeviceInterfaceDetail函式所傳回的DevicePathName
        '******************************************************************************
    
        hiddevice = CreateFile _
            (DevicePathName, _
            GENERIC_READ Or GENERIC_WRITE, _
            (FILE_SHARE_READ Or FILE_SHARE_WRITE), _
            0, _
            OPEN_EXISTING, _
            0, _
            0)
            
        '到這裡，已經找到我們正在尋找的裝置
        
        '******************************************************************************
        'HidD_GetAttributes
        '需求：
         '需要從裝置的訊息，以及從CreatFile函式所回傳的Handle
        '回傳：
        '回傳包含了販售商VID，產品PID以及產品版本碼的HIDD_ATTRIBUTES結構
        '我們可以使用這個訊息來決定是否這個裝置是我們所要尋找的.
        '******************************************************************************
        
        '設定"Size"屬性為在結構中的位元組數值

        
        Result = HidD_GetAttributes _
            (hiddevice, _
            DeviceAttributes)
     
        '此一函式主要是測試我們所要控制裝置的產品碼與廠商碼是否與目標裝置一樣，因此我們需要由將產品碼與廠商碼輸入給函式。
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

'得到裝置功能訊息
Public Sub GetDeviceCapabilities()

'******************************************************************************
'HidD_GetPreparsedData
'回傳：包含指向一個關於此裝置能力訊息的緩衝區的指標器
'需求：由CreateFile所回傳的handle
'注意事項：在此不需直接處理緩衝區
'但HidP_GetCaps與其他的API函式需要一個指向此緩衝區的指標器
'*******************************************************************************

Dim ppData(29) As Byte
Dim ppDataString As Variant

'被剖析的資料（PreparsedData）指標器是指向函式所宣告的緩衝區
Result = HidD_GetPreparsedData _
    (hiddevice, _
    PreparsedData)

'將位於PreparsedData的資料拷貝到位元組陣列

Result = RtlMoveMemory _
    (ppData(0), _
    PreparsedData, _
    30)

'******************************************************************************
'HidP_GetCaps
'求出裝置的能力
'對於標準HID的裝置，如鍵盤或是搖桿等，都可以找到此裝置特定的能力
'而對於自訂的裝置，the software will probably know what the device is capable of,
'so this call only verifies the information.
'需求：涵蓋這些訊息的緩衝區的指標
'這個指標器是由HidD_GetPreparsedData函式所回傳的.
'回傳: 包含相關訊息的能力之結構
'******************************************************************************
Result = HidP_GetCaps _
    (PreparsedData, _
    Capabilities)

'******************************************************************************
'HidP_GetValueCaps
'回傳：傳回一個包含了HidP_ValueCaps結構陣列的緩衝區
'注意事項：每一個結構定義了一個數值的能力
'這個應用程式並不使用此資料
'******************************************************************************

'此位元組陣列是維持在這結構中

Dim ValueCaps(1024) As Byte

Result = HidP_GetValueCaps _
    (HidP_Input, _
    ValueCaps(0), _
    Capabilities.NumberInputValueCaps, _
    PreparsedData)
    
'為了使用這資料，將位元組陣列拷貝到結構中
End Sub
'接收資料
Public Static Function ReadDevice() As Variant
'從裝置讀取資料
Dim Count
Dim NumberOfBytesRead As Long

'宣告給報告用的緩衝區，'Byte 0是報告ID，所以需減去1
Dim ReadBuffer() As Byte
Dim UBoundReadBuffer As Integer
Dim ByteValue(8) As String
Dim recevicedata As Variant
'******************************************************************************
'HidD_GetFeature
'傳回: 在ReadBuffer的報告描述元.
'需求: 由CreateFile所傳回的裝置handle,
'注意事項：這特性報告的長度是由HidP_GetCaps函式所回傳的，且以位元組為單位
'******************************************************************************

'ReadFile是區塊呼叫
'這個應用程式會"掛"住，直到裝置送出所需的資料量為止
'為了避免"掛"了，我們必須確定裝置總是有資料來送出

ReDim ReadBuffer(Capabilities.FeatureReportByteLength - 1)
NumberOfBytesRead = Capabilities.FeatureReportByteLength
'ReadBuffer陣列是以0為啟始的, 因此需將位元組的數目減去1
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

For Count = 1 To UBound(ReadBuffer)
     '增加前導為0到0 - Fh數值前
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
'傳送資料

Dim NumberOfBytesToSend As Long
Dim NumberOfBytesWritten As Long
Dim SendBuffer() As Byte
Dim Count As Integer
Dim NumberOfBytesRead As Long
Dim ReadBuffer() As Byte

'這個SendBuffer陣列是從0開始的,所以需減去1個位元組的數目
ReDim SendBuffer(Capabilities.FeatureReportByteLength - 1)

'******************************************************************************
'HidD_SetFeature函式
'送出報告給裝置使用
'傳回：成功或是失敗
'需求：由CreatFile函式所回傳的handle
'由HidP_GetCaps函式傳回特性報告的長度
'******************************************************************************

'第一個位元組是Report ID
SendBuffer(0) = 0

'下一個位元組是相關的資料
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

'關閉所裝置
Public Static Function CloseDevice()
'當程式結束前所必須執行的動作，全部放在這個副程式內

'關閉為此HID裝置所開啟的handle
Result = CloseHandle(hiddevice)
Result = SetupDiDestroyDeviceInfoList(DeviceInfoSet)
Result = HidD_FreePreparsedData(PreparsedData)
End Function
'找尋HID的裝置共有多少個
Public Static Function FindHidDevice() As String
Dim HidGuid As GUID
Dim MemberIndex As Long
Dim dev1 As Long
Dim i As Long
Result = HidD_GetHidGuid(HidGuid)
DeviceInfoSet = SetupDiGetClassDevs _
    (HidGuid, _
    vbNullString, _
    0, _
    (DIGCF_PRESENT Or DIGCF_DEVICEINTERFACE))
    
 i = 0
 For dev1 = 0 To 127
 MemberIndex = dev1
 MyDeviceInterfaceData.cbSize = LenB(MyDeviceInterfaceData)
 Result = SetupDiEnumDeviceInterfaces _
        (DeviceInfoSet, _
        0, _
        HidGuid, _
        MemberIndex, _
        MyDeviceInterfaceData)
 If Result <> 0 Then
 i = i + 1
 End If
 Next

 FindHidDevice = i
 
End Function




