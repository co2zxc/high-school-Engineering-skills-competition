VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "USB介面通訊測試-SetupDiGetDeviceInterfaceDetail函式"
   ClientHeight    =   3180
   ClientLeft      =   255
   ClientTop       =   330
   ClientWidth     =   5445
   LinkTopic       =   "Form1"
   ScaleHeight     =   3180
   ScaleWidth      =   5445
   Begin VB.Frame fraSendAndReceive 
      Caption         =   "輸入與輸出資料"
      Height          =   855
      Left            =   0
      TabIndex        =   1
      Top             =   120
      Width           =   5415
      Begin VB.CommandButton cmdOnce 
         Caption         =   "確定"
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
'介面設計與實習-使用Visual Basic，CH19
'以SetupDiGetDeviceInterfaceDetail（）函式測試USB HID裝置
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
            
        'Call DisplayResultOfAPICall("CreateFile")
        'lstResults.AddItem "  Returned handle: " & Hex$(HidDevice) & "h"
        
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
        
        '找出是否這個裝置與我們正在尋找的裝置吻合Find out if the device matches the one we're looking for.
        If (DeviceAttributes.VendorID = MyVendorID) And _
            (DeviceAttributes.ProductID = MyProductID) Then
        '        lstResults.AddItem "  My device detected"
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
Call FindTheHid
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


