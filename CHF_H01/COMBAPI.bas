Attribute VB_Name = "COMBAPI"

Public Declare Function OpenUsbDevice Lib "USBHIO.dll" (ByVal MyUsbVendorID As Integer, ByVal MyUsbProductID As Integer) As Boolean
Public Declare Sub OutDataCtrl Lib "USBHIO.dll" (ByVal OutData As Byte, ByVal OutControl As Byte)

Public Declare Sub OutDataEightByte Lib "USBHIO.dll" (ByVal OutData1 As Byte, ByVal OutData2 As Byte, ByVal OutData3 As Byte, _
                                                      ByVal OutData4 As Byte, ByVal OutData5 As Byte, ByVal OutData6 As Byte, _
                                                      ByVal OutData7 As Byte, ByVal OutData8 As Byte)

Public Declare Sub ReadData Lib "USBHIO.dll" (ByRef ReadBuffer() As Byte)
Public Declare Sub CloseUsbDevice Lib "USBHIO.dll" ()

Public a As Integer
Public b(99) As Byte
Public c As Integer
Public ReadKeyData(8) As Byte

Public Const VendorID = &H1234
Public Const ProductID = &H2468
'

Public Sub OutRedLED(ByVal RedData As Byte)
    OutDataCtrl RedData, &H20
End Sub
'

Public Sub ShowRedLED(ByVal RedLED As Byte)
    Dim tt As Byte
    For tt = 0 To 7
        If ((RedLED And 2 ^ tt) = 2 ^ tt) Then
            frmMain.Shape2(tt).FillColor = &HFF&
        Else
            frmMain.Shape2(tt).FillColor = &H0&
        End If
    Next tt
End Sub
'

Public Sub ShowInitLEDs()
    Dim ss As Integer
    For ss = 0 To 7
        frmMain.Shape2(ss).FillStyle = 1
    Next ss
End Sub
'

Public Sub ShowNormalLEDs()
    Dim ss As Integer
    For ss = 0 To 7
        frmMain.Shape2(ss).FillStyle = 0
    Next ss
End Sub
'
