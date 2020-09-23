<div align="center">

## Get window handlle \(hWnd\) from process ID


</div>

### Description

A module for retrieving the handle number (hWnd) of a window providing only the process id number (PID). This code is seeking handles that match the given PID and returns the hanlde that refers to a visible window. But, since not all the processes running are windowed or some processes may have multiple windows, may not work for every case.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Evripidis](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/evripidis.md)
**Level**          |Intermediate
**User Rating**    |4.7 (28 globes from 6 users)
**Compatibility**  |VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/evripidis-get-window-handlle-hwnd-from-process-id__1-63954/archive/master.zip)

### API Declarations

```
Private Declare Function EnumWindows Lib "user32" _
 (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Boolean
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" _
 (ByVal Hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" _
 (ByVal Hwnd As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
  (ByVal Hwnd As Long, ByVal wIndx As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" _
 (ByVal Hwnd As Long, lpdwProcessId As Long) As Long
```


### Source Code

```
'EVRIS JAN 2006
'vcr4545@yahoo.com
Option Explicit
Private Const GWL_STYLE = (-16)
Private Const GWL_EXSTYLE = (-20)
Private Const WS_VISIBLE = &H10000000
Private Const WS_EX_APPWINDOW = &H40000
Private Type HWND_TEXT
 Window_Handle As Long
 Window_Title As String
End Type
Private ColCounter As Long
Private HandleTextCollection() As HWND_TEXT
Private Declare Function EnumWindows Lib "user32" _
  (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Boolean
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" _
  (ByVal Hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" _
  (ByVal Hwnd As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
    (ByVal Hwnd As Long, ByVal wIndx As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" _
  (ByVal Hwnd As Long, lpdwProcessId As Long) As Long
Private Function EnumWindowsProc(ByVal Hwnd As Long, ByVal lParam As Long) As Boolean
 Dim Title As String
 Dim r As Long
 r = GetWindowTextLength(Hwnd)
 Title = Space(r)
 GetWindowText Hwnd, Title, r + 1
 'Add to type array
 ColCounter = ColCounter + 1
 ReDim Preserve HandleTextCollection(ColCounter)
 HandleTextCollection(ColCounter).Window_Handle = Hwnd
 HandleTextCollection(ColCounter).Window_Title = Title
 EnumWindowsProc = True
End Function
Private Sub EnumAllWindows()
 ColCounter = 0
 ReDim HandleTextCollection(0)
 EnumWindows AddressOf EnumWindowsProc, ByVal 0&
 'At this point, HandleTextCollection() array holds the handles and window titles
 'of all windows enumerated
End Sub
Public Function GetWindowHandleFromPID(ByVal AppPID As Long) As Long
 Dim AppPID_HWND() As HWND_TEXT   'Will store only handles related to AppPID
 Dim CNT As Long
 Dim HND As Long
 Dim n As Long
 Dim r As Long
 Dim i As Long
 Dim TaskID As Long
 Dim TheHandle As Long
 'STEP 1: Find all window handles currently running
 'and put them into HandleTextCollection() array
 Call EnumAllWindows
 'STEP 2: Filter handles and keep only theese tha match with AppPID pid number
 'and put them into AppPID_HWND() array
 'GetWindowThreadProcessId returns by reference the PID for each handle given
 CNT = 0
 ReDim AppPID_HWND(0)
 For i = 1 To UBound(HandleTextCollection())
  HND = HandleTextCollection(i).Window_Handle
  n = GetWindowThreadProcessId(HND, TaskID)
  If TaskID = AppPID Then
   'Handle matches AppPID
   CNT = CNT + 1
   ReDim Preserve AppPID_HWND(CNT)
   AppPID_HWND(CNT).Window_Handle = HND
   AppPID_HWND(CNT).Window_Title = HandleTextCollection(i).Window_Title
  End If
 Next i
 'STEP 3: From all the handles related to AppPID, search for the handle
 'that refers to a window that is visible.
 For i = 1 To UBound(AppPID_HWND())
  'Must be a visible window
  n = GetWindowLong(AppPID_HWND(i).Window_Handle, GWL_STYLE)
  'And a top-level window onto the taskbar since the window is visible
  r = GetWindowLong(AppPID_HWND(i).Window_Handle, GWL_EXSTYLE)
  If (n) And (WS_VISIBLE) Then
   If (r) And (WS_EX_APPWINDOW) Then
    'Is it the right hWnd ?
    'TODO: GetTitleBarInfo
    TheHandle = AppPID_HWND(i).Window_Handle
    Exit For
   End If
  End If
 Next i
 GetWindowHandleFromPID = TheHandle
End Function
```

