VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "USB 介面通訊測試-SetupDiGetClassDevsc函式"
   ClientHeight    =   3180
   ClientLeft      =   255
   ClientTop       =   330
   ClientWidth     =   5445
   LinkTopic       =   "Form1"
   ScaleHeight     =   3180
   ScaleWidth      =   5445
   Begin VB.Frame fraSendAndReceive 
      Caption         =   "輸入/輸出資料"
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
'以SetupDiGetClassDevs（）函式測試USB HID裝置
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


