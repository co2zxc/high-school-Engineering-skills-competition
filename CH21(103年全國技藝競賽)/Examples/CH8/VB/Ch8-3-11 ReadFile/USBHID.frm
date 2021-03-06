VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "USB介面通訊測試-ReadFile函式"
   ClientHeight    =   3180
   ClientLeft      =   255
   ClientTop       =   330
   ClientWidth     =   5445
   LinkTopic       =   "Form1"
   ScaleHeight     =   3180
   ScaleWidth      =   5445
   Begin VB.Frame fraBytesReceived 
      Caption         =   "輸入位元組"
      Height          =   855
      Left            =   1560
      TabIndex        =   5
      Top             =   120
      Width           =   1575
      Begin VB.TextBox txtBytesReceived 
         Height          =   375
         Left            =   360
         TabIndex        =   6
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame fraBytesToSend 
      Caption         =   "輸出位元組"
      Height          =   855
      Left            =   0
      TabIndex        =   3
      Top             =   120
      Width           =   1455
      Begin VB.TextBox txtBytesSend 
         Height          =   375
         Left            =   360
         TabIndex        =   4
         Text            =   "00"
         Top             =   360
         Width           =   732
      End
   End
   Begin VB.Frame fraSendAndReceive 
      Caption         =   "輸入與輸出資料"
      Height          =   855
      Left            =   3240
      TabIndex        =   1
      Top             =   120
      Width           =   1935
      Begin VB.CommandButton cmdOnce 
         Caption         =   "確定"
         Height          =   495
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1695
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
'介面設計與實習-使用Visual Basic，CH19
'以WriteFile（）函式測試USB HID裝置
'PC主機與HID裝置的通訊程式 -> 修改自 Jan Axelson (jan@lvr.com)

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

'設定VID/PID碼以符和裝置韌體程式碼與INF安裝檔案

Const MyVendorID = &H1234
Const MyProductID = &H2468

Function FindTheHid() As Boolean
'執行一系列的API函式呼叫，以找到所需的HID群組裝置
'如果偵測到HID裝置的話，就傳回真值，反之就傳回假值

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
Call DisplayResultOfAPICall("GetHidGuid")

'==== 顯示 GUID ====
GUIDString = _
    Hex$(HidGuid.Data1) & "-" & _
    Hex$(HidGuid.Data2) & "-" & _
    Hex$(HidGuid.Data3) & "-"

For Count = 0 To 7
    '根據GUID宣告的格式，需確保每一個GUID的8個位元組顯示2個字元
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
'回傳： 針對已安裝的裝置回傳一個指向裝置訊息集的handle
'需求：由GetHidGuid所回傳的HidGuid
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
'SetupDiEnumDeviceInterfaces函式
'回傳： 包含了指向一個所偵測到裝置的SP_DEVICE_INTERFACE_DATA結構的handle
'需求：
'1. 從SetupDiGetClassDevs函式所回傳的DeviceInfoSet
'2. 從GetHidGuid所回傳的HidGuid
'3. 設定此裝置的索引值
'******************************************************************************

'以0開始，並加以遞增直到沒有裝置被偵測到為止
MemberIndex = 0

Do
    '在MyDeviceInterfaceData結構中的cbSize元素必須被設定
    '這個結構是以位元組為計數數值，這個大小值是28個位元組
    MyDeviceInterfaceData.cbSize = LenB(MyDeviceInterfaceData)
    Result = SetupDiEnumDeviceInterfaces _
        (DeviceInfoSet, _
        0, _
        HidGuid, _
        MemberIndex, _
        MyDeviceInterfaceData)
    
    Call DisplayResultOfAPICall("SetupDiEnumDeviceInterfaces")
    If Result = 0 Then LastDevice = True
    
    '如果裝置存在的話，顯示回傳的訊息
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
'SetupDiGetDeviceInterfaceDetail函式
'回傳：包含此裝置訊息的SP_DEVICE_INTERFACE_DETAIL_DATA結構
'注意事項：為了取回相關的訊息，呼叫此函式兩次
'第一次傳回所需的結構大小
'第二次傳回在DeviceInfoSet資料的指標器
'需求：由SetupDiGetClassDevs所回傳的DeviceInfoSet
 '以及由SetupDiEnumDeviceInterfaces所回傳的SP_DEVICE_INTERFACE_DATA結構
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
        
        Call DisplayResultOfAPICall(" Result of second call: ")
        lstResults.AddItem "  MyDeviceInterfaceDetailData.cbSize: " & _
            CStr(MyDeviceInterfaceDetailData.cbSize)
        
        '將位元組陣列轉換成字串
        DevicePathName = CStr(DetailDataBuffer())
        '轉換成Unicode碼格式.
        DevicePathName = StrConv(DevicePathName, vbUnicode)
        '從開始處剔除cbSize (4 bytes).
        DevicePathName = Right$(DevicePathName, Len(DevicePathName) - 4)
        lstResults.AddItem "  Device pathname: "
        lstResults.AddItem "    " & DevicePathName
                
        '******************************************************************************
        'CreateFile函式
        '回傳：
        '回傳用來致能讀取與寫入此裝置的handle
        '需求:
        '由SetupDiGetDeviceInterfaceDetail函式所傳回的DevicePathName
        '******************************************************************************
    
        HidDevice = CreateFile _
            (DevicePathName, _
            GENERIC_READ Or GENERIC_WRITE, _
            (FILE_SHARE_READ Or FILE_SHARE_WRITE), _
            0, _
            OPEN_EXISTING, _
            0, _
            0)
            
        Call DisplayResultOfAPICall("CreateFile")
        lstResults.AddItem "  Returned handle: " & Hex$(HidDevice) & "h"
        
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
        DeviceAttributes.Size = LenB(DeviceAttributes)
        Result = HidD_GetAttributes _
            (HidDevice, _
            DeviceAttributes)
            
        Call DisplayResultOfAPICall("HidD_GetAttributes")
        If Result <> 0 Then
            lstResults.AddItem "  HIDD_ATTRIBUTES structure filled without error."
        Else
            lstResults.AddItem "  Error in filling HIDD_ATTRIBUTES structure."
        End If
    
        lstResults.AddItem "  Structure size: " & DeviceAttributes.Size
        lstResults.AddItem "  Vendor ID: " & Hex$(DeviceAttributes.VendorID)
        lstResults.AddItem "  Product ID: " & Hex$(DeviceAttributes.ProductID)
        lstResults.AddItem "  Version Number: " & Hex$(DeviceAttributes.VersionNumber)
        
        '找出是否這個裝置與我們正在尋找的裝置吻合Find out if the device matches the one we're looking for.
        If (DeviceAttributes.VendorID = MyVendorID) And _
            (DeviceAttributes.ProductID = MyProductID) Then
                lstResults.AddItem "  My device detected"
                MyDeviceDetected = True
        Else
                MyDeviceDetected = False
                '如果不是我們所想要的裝置，就關掉其handle.
                Result = CloseHandle _
                    (HidDevice)
                DisplayResultOfAPICall ("CloseHandle")
        End If
End If
    '持續地尋找直到我們已發現此裝置，或是已經沒有剩下的裝置要檢查

    MemberIndex = MemberIndex + 1
    
Loop Until (LastDevice = True) Or (MyDeviceDetected = True)

If MyDeviceDetected = True Then
    FindTheHid = True
Else
    lstResults.AddItem " Device not found."
End If

End Function

Private Function GetDataString _
    (Address As Long, _
    Bytes As Long) _
As String

'取回記憶體Retrieves a string of length Bytes from memory, beginning at Address.
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

'傳回上一個錯誤的錯誤訊息

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
    
'從訊息中跳過 CR與LF字元，因此減去2個字元

If Bytes > 2 Then
    GetErrorString = Left$(ErrorString, Bytes - 2)
End If

End Function

Private Sub cmdOnce_Click()
lstResults.Clear
Call ReadAndWriteToDevice
End Sub
Private Sub DisplayResultOfAPICall(FunctionName As String)

'呈現API呼叫的結果

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

Private Sub Form_Unload(Cancel As Integer)
Call Shutdown
End Sub

Private Sub GetDeviceCapabilities()

'******************************************************************************
'HidD_GetPreparsedData
'回傳：包含指向一個關於此裝置能力訊息的緩衝區的指標器
'需求：由CreateFile所回傳的handle
'注意事項：在此不需直接處理緩衝區
'但HidP_GetCaps與其他的API函式需要一個指向此緩衝區的指標器
'******************************************************************************

Dim ppData(29) As Byte
Dim ppDataString As Variant

'被剖析的資料（PreparsedData）指標器是指向函式所宣告的緩衝區
Result = HidD_GetPreparsedData _
    (HidDevice, _
    PreparsedData)
Call DisplayResultOfAPICall("HidD_GetPreparsedData")


'將位於PreparsedData的資料拷貝到位元組陣列

Result = RtlMoveMemory _
    (ppData(0), _
    PreparsedData, _
    30)

Call DisplayResultOfAPICall("RtlMoveMemory")

ppDataString = ppData()

'將資料轉換成Unicode格式
ppDataString = StrConv(ppDataString, vbUnicode)

'******************************************************************************
'HidP_GetCaps
'求出裝置的能力
'對於標準HID的裝置，如鍵盤或是搖桿等，都可以找到此裝置特定的能力
'而對於自訂的裝置，這個程式僅可能知道這裝置的能力是什麼,這個呼叫也僅驗證相關的資訊而已。
'需求：涵蓋這些訊息的緩衝區的指標
'這個指標器是由HidD_GetPreparsedData函式所回傳的.
'回傳: 包含相關訊息的能力之結構
'******************************************************************************
Result = HidP_GetCaps _
    (PreparsedData, _
    Capabilities)

Call DisplayResultOfAPICall("HidP_GetCaps")
lstResults.AddItem "  Last error: " & ErrorString
lstResults.AddItem "  Usage: " & Hex$(Capabilities.Usage)
lstResults.AddItem "  Usage Page: " & Hex$(Capabilities.UsagePage)
lstResults.AddItem "  Input Report Byte Length: " & Capabilities.InputReportByteLength
lstResults.AddItem "  Output Report Byte Length: " & Capabilities.OutputReportByteLength
lstResults.AddItem "  Feature Report Byte Length: " & Capabilities.FeatureReportByteLength
lstResults.AddItem "  Number of Link Collection Nodes: " & Capabilities.NumberLinkCollectionNodes
lstResults.AddItem "  Number of Input Button Caps: " & Capabilities.NumberInputButtonCaps
lstResults.AddItem "  Number of Input Value Caps: " & Capabilities.NumberInputValueCaps
lstResults.AddItem "  Number of Input Data Indices: " & Capabilities.NumberInputDataIndices
lstResults.AddItem "  Number of Output Button Caps: " & Capabilities.NumberOutputButtonCaps
lstResults.AddItem "  Number of Output Value Caps: " & Capabilities.NumberOutputValueCaps
lstResults.AddItem "  Number of Output Data Indices: " & Capabilities.NumberOutputDataIndices
lstResults.AddItem "  Number of Feature Button Caps: " & Capabilities.NumberFeatureButtonCaps
lstResults.AddItem "  Number of Feature Value Caps: " & Capabilities.NumberFeatureValueCaps
lstResults.AddItem "  Number of Feature Data Indices: " & Capabilities.NumberFeatureDataIndices

'******************************************************************************
'HidP_GetValueCaps
'回傳：傳回一個包含了HidP_ValueCaps結構陣列的緩衝區
'注意事項：每一個結構定義了一個數值的能力
'這個應用程式並不使用此資料
'******************************************************************************

'此位元組陣列是維持在這結構中
Dim ValueCaps(1023) As Byte

Result = HidP_GetValueCaps _
    (HidP_Input, _
    ValueCaps(0), _
    Capabilities.NumberInputValueCaps, _
    PreparsedData)
   
Call DisplayResultOfAPICall("HidP_GetValueCaps")


'為了使用此資料，將位元組陣列拷貝到結構陣列中

End Sub

Private Sub ReadAndWriteToDevice()
'送出1個位元組給裝置，並且再讀回來.

Dim DeviceDetected As Boolean
Dim temp As String
'報告的Header
'lstResults.AddItem "HID Test Report"
'lstResults.AddItem Format(Now, "general date")


If Len(Trim(txtBytesSend.Text)) < 2 Then
txtBytesSend.Text = "0" & txtBytesSend.Text
ElseIf Len(Trim(txtBytesSend.Text)) > 2 Then
txtBytesSend.Text = Right(txtBytesSend.Text, 2)
End If

txtBytesSend.Text = UCase(txtBytesSend.Text)
OutputReportData(0) = AHex(txtBytesSend.Text)  'cboByte0.ListIndex
'OutputReportData(1) = cboByte1.ListIndex

'找尋裝置
DeviceDetected = FindTheHid
If DeviceDetected = True Then
    '學習裝置的能力
    Call GetDeviceCapabilities
    '將報告寫入給裝置
    Call WriteReport
        
    '增加延遲的時間允許主機有時間去輪詢所回傳的資料
    Timeout = False
    tmrDelay.Interval = 100
    tmrDelay.Enabled = True
    Do
        DoEvents
    Loop Until Timeout = True
    '從裝置讀取報告
    Call ReadReport
Else
End If

'Scroll to the bottom of the list box.
'lstResults.ListIndex = lstResults.ListCount - 1

End Sub

Private Sub ReadReport()

'Read data from the device.

Dim Count
Dim NumberOfBytesRead As Long
'Allocate a buffer for the report.
'Byte 0 is the report ID.
Dim ReadBuffer() As Byte
Dim UBoundReadBuffer As Integer

'******************************************************************************
'ReadFile
'傳回: 在ReadBuffer的報告描述元.
'需求: 由CreateFile所傳回的裝置handle,
'注意事項：這輸入報告的長度是由HidP_GetCaps函式所回傳的，且以位元組為單位
'******************************************************************************

'ReadFile是區塊呼叫
'這個應用程式會"掛"住，直到裝置送出所需的資料量為止
'為了避免"掛"了，我們必須確定裝置總是有資料來送出

Dim ByteValue As String

'ReadBuffer陣列是以0為啟始的, 因此需將位元組的數目減去1
ReDim ReadBuffer(Capabilities.FeatureReportByteLength - 1)
    
NumberOfBytesRead = Capabilities.FeatureReportByteLength

'傳讀取緩衝區的第一個位元組的位址
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
Call DisplayResultOfAPICall("HidD_GetFeature")

lstResults.AddItem " Report ID: " & ReadBuffer(0)
lstResults.AddItem " Report Data:"

txtBytesReceived.Text = ""
For Count = 1 To UBound(ReadBuffer)
    'Add a leading 0 to values 0 - Fh.
    If Len(Hex$(ReadBuffer(Count))) < 2 Then
        ByteValue = "0" & Hex$(ReadBuffer(Count))
    Else
        ByteValue = Hex$(ReadBuffer(Count))
    End If
    lstResults.AddItem " " & ByteValue
Next Count

    '將所接收到的位元組放置在文字盒內
    'txtBytesReceived.SelStart = Len(txtBytesReceived.Text)
    'txtBytesReceived.SelText = ByteValue & vbCrLf
    txtBytesReceived.Text = Hex$(ReadBuffer(1))
End Sub

Private Sub Shutdown()
'當程式結束前所必須執行的動作，全部放在這個副程式內

'關閉為此HID裝置所開啟的handle
Result = CloseHandle _
    (HidDevice)
Call DisplayResultOfAPICall("CloseHandle (HidDevice)")

'使用SetupDiGetClassDevs函式來釋放所佔用的記憶體
'若是非零代表成功

Result = SetupDiDestroyDeviceInfoList _
    (DeviceInfoSet)
Call DisplayResultOfAPICall("DestroyDeviceInfoList")

Result = HidD_FreePreparsedData _
    (PreparsedData)
Call DisplayResultOfAPICall("HidD_FreePreparsedData")

End Sub

Private Sub tmrDelay_Timer()
Timeout = True
tmrDelay.Enabled = False
End Sub

Private Sub WriteReport()
'送出資料給裝置

Dim Count As Integer
Dim NumberOfBytesRead As Long
Dim NumberOfBytesToSend As Long
Dim NumberOfBytesWritten As Long
Dim ReadBuffer() As Byte
Dim SendBuffer() As Byte

'這個SendBuffer陣列是從0開始的,所以需減去1個位元組的數目
ReDim SendBuffer(Capabilities.FeatureReportByteLength - 1)

'******************************************************************************
'WriteFile函式
'送出報告給裝置使用
'傳回：成功或是失敗
'需求：由CreatFile函式所回傳的handle
'由HidP_GetCaps函式傳回輸出報告的長度
'******************************************************************************

'第一個位元組是Report ID
SendBuffer(0) = 0

'下一個位元組是相關的資料
For Count = 1 To Capabilities.FeatureReportByteLength - 1
    SendBuffer(Count) = OutputReportData(Count - 1)
Next Count

NumberOfBytesWritten = Capabilities.FeatureReportByteLength

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
Call DisplayResultOfAPICall("HidD_SetFeature")

lstResults.AddItem " FeatureReportByteLength = " & Capabilities.FeatureReportByteLength
lstResults.AddItem " NumberOfBytesWritten = " & NumberOfBytesWritten
lstResults.AddItem " Report ID: " & SendBuffer(0)
lstResults.AddItem " Report Data:"

For Count = 1 To UBound(SendBuffer)
    lstResults.AddItem " " & Hex$(SendBuffer(Count))
Next Count

End Sub

'原作者:布雷克
Private Function AHex(a) 'As String)   '16轉10進制
    Dim sum, i
    For i = Len(a) To 1 Step -1
        If Asc(Mid(a, i, 1)) < 58 And Asc(Mid(a, i, 1)) > 47 Then
            sum = sum + Val(Mid(a, i, 1)) * 16 ^ (Len(a) - i)
        End If
        If Asc(Mid(a, i, 1)) < 71 And Asc(Mid(a, i, 1)) > 64 Then
            sum = sum + (Asc(Mid(a, i, 1)) - 55) * 16 ^ (Len(a) - i)
        End If
    Next i
    AHex = sum
End Function

