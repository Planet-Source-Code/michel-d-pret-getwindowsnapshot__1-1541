<div align="center">

## GetWindowSnapShot


</div>

### Description

This allows a VB program to capture either the screen or the program window.

It has been tested under Win95 and NT4.0. It derives from a routine by Dan Appleman (VisualBasic 5.0 Programmer's Guide to the WIN32 API, page 303) which unfortunately does not work reliably under all conditions. Dan Appleman's exhaustive preliminary tutorial, though, is all it takes to understand the code.
 
### More Info
 
mode - 0 = screen, 1 = window

a reference to an image control

'Create a form, define two command controls and an image control, insert the following code:

Private Sub Command1_Click()

GetWindowSnapShot 0, Image1

End Sub

Private Sub Command2_Click()

GetWindowSnapShot 1, Image1

End Sub

no known side effect


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Michel D\. PRET](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/michel-d-pret.md)
**Level**          |Unknown
**User Rating**    |4.7 (14 globes from 3 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/michel-d-pret-getwindowsnapshot__1-1541/archive/master.zip)

### API Declarations

```
Declare Function MapVirtualKey Lib "user32" Alias "MapVirtualKeyA" (ByVal wCode As Long, ByVal wMapType As Long) As Long
Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" _
     (LpVersionInformation As OSVERSIONINFO) As Long
Public Const VK_MENU = &H12
Public Const KEYEVENTF_KEYUP = &H2
Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128   ' Maintenance string for PSS usage
End Type
```


### Source Code

```
'Insert this in a module:
Public Sub GetWindowSnapShot(Mode As Long, ThisImage As Image)
 ' mode = 0 -> Screen snapshot
 ' mode = 1 -> Window snapshot
 Dim altscan%, NT As Boolean, nmode As Long
 NT = IsNT
 If Not NT Then
  If Mode = 0& Then Mode = 1& Else Mode = 0&
 End If
 If NT And Mode = 0 Then
   keybd_event vbKeySnapshot, 0&, 0&, 0&
 Else
   altscan = MapVirtualKey(VK_MENU, 0)
   keybd_event VK_MENU, altscan, 0, 0
   DoEvents
   keybd_event vbKeySnapshot, Mode, 0&, 0&
 End If
 DoEvents
 ThisImage = Clipboard.GetData(vbCFBitmap)
 keybd_event VK_MENU, altscan, KEYEVENTF_KEYUP, 0
End Sub
Public Function IsNT() As Boolean
 Dim verinfo As OSVERSIONINFO
 verinfo.dwOSVersionInfoSize = Len(verinfo)
 If (GetVersionEx(verinfo)) = 0 Then Exit Function
 If verinfo.dwPlatformId = 2 Then IsNT = True
End Function
```

