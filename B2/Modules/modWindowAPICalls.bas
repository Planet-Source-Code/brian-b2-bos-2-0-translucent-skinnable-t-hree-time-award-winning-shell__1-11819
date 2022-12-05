Attribute VB_Name = "modWindowAPICalls"
'Function for resizing desktop
Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long

'Constants for SystemParametersInfo
Public Const SPI_SETWORKAREA = 47
Public Const SPI_GETWORKAREA = 48
Public Const SPIF_SENDCHANGE = &H2

Public Declare Function SetWindowPos Lib "user32" (ByVal hWND As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
'Constants for SetWindowPos:
Public Const HWND_TOPMOST = -1
Public Const SW_MINIMIZE = 6
Public Const SW_SHOWDEFAULT = 10
Public Const SW_SHOWMAXIMIZED = 3
Public Const SW_SHOWMINIMIZED = 2
Public Const SW_SHOWMINNOACTIVE = 7
Public Const SW_SHOWNA = 8
Public Const SW_SHOWNOACTIVATE = 4
Public Const SW_SHOWNORMAL = 1
Public Const SWP_FRAMECHANGED = &H20
Public Const SWP_DRAWFRAME = SWP_FRAMECHANGED
Public Const SWP_HIDEWINDOW = &H80
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_NOCOPYBITS = &H100
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOOWNERZORDER = &H200
Public Const SWP_NOREDRAW = &H8
Public Const SWP_NOREPOSITION = SWP_NOOWNERZORDER
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOZORDER = &H4
Public Const SWP_SHOWWINDOW = &H40

Public Declare Function GetForegroundWindow Lib "user32" () As Long

'Function for listing windows
Private Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Private Declare Function IsIconic Lib "user32" (ByVal hWND As Long) As Long
Private Declare Function IsWindowVisible Lib "user32" (ByVal hWND As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hWND As Long) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hWND As Long, ByVal wCmd As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWND As Long, ByVal nIndex As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWND As Long, ByVal lpString As String, ByVal cch As Long) As Long

'Constants for the function
Private Const SW_SHOW = 5
Private Const SW_RESTORE = 9
Private Const GW_OWNER = 4
Private Const GWL_HWNDPARENT = (-8)
Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_TOOLWINDOW = &H80
Private Const WS_EX_APPWINDOW = &H40000

Public WindowName() As String
Public WindowID() As Long

Private WindowCaption As String
Public AWI As Integer
Private AWH As Long
Public Declare Function RedrawWindow Lib "user32" (ByVal hWND As Long, lprcUpdate As RECT, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Declare Function ShowWindow Lib "user32" (ByVal hWND As Long, ByVal nCmdShow As Long) As Long
Public Const RDW_INVALIDATE = &H1
Public Const RDW_INTERNALPAINT = &H2
Public Const RDW_ERASE = &H4
Public Const RDW_VALIDATE = &H8
Public Const RDW_NOINTERNALPAINT = &H10
Public Const RDW_NOERASE = &H20
Public Const RDW_NOCHILDREN = &H40
Public Const RDW_ALLCHILDREN = &H80
Public Const RDW_UPDATENOW = &H100
Public Const RDW_ERASENOW = &H200
Public Const RDW_FRAME = &H400
Public Const RDW_NOFRAME = &H800




Public Function GetWindows() As Variant
    ReDim WindowName(0)
    ReDim WindowID(0)

    Call EnumWindows(AddressOf AddToArray, 1) 'The "AddToArray" function will be called once for each window
    WindowName = RemoveFirstElement(ShellSort(WindowName))
    WindowID = RemoveFirstElement(WindowID)
    
    AWI = GetAWI
    
    GetWindows = WindowName
End Function

Public Function ActivateWindow(hWND As Long)
    If AWI <> -1 Then
        If WindowID(AWI) = hWND Then 'The window is already active. Minimize it
            ShowWindow hWND, SW_MINIMIZE
            Exit Function
        End If
    End If
    
    If IsIconic(hWND) Then ShowWindow hWND, SW_RESTORE
    SetWindowPos hWND, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOREDRAW Or SWP_NOSIZE Or SWP_NOREPOSITION Or SWP_NOZORDER
End Function

Private Function AddToArray(ByVal hWND As Long, ByVal lParam As Long) As Long
    If IsWindowVisible(hWND) Then
        If GetParent(hWND) = 0 Then
            HasNoOwner = (GetWindow(hWND, GW_OWNER) = 0)
            WindowStyle = GetWindowLong(hWND, GWL_EXSTYLE)
            If (((WindowStyle And WS_EX_TOOLWINDOW) = 0) And HasNoOwner) Or ((WindowStyle And WS_EX_APPWINDOW) And Not HasNoOwner) Then
                WindowCaption = Space(256)
                GetWindowText hWND, WindowCaption, Len(WindowCaption)
                x = UBound(WindowName) + 1
                ReDim Preserve WindowName(x)
                ReDim Preserve WindowID(x + 1)
                WindowName(x) = RTrim(WindowCaption)
                WindowID(x) = hWND
            End If
        End If
    End If
    AddToArray = True
End Function

Function RemoveFirstElement(InArray As Variant) As Variant
    If UBound(InArray) <> 0 Then
        For i = 1 To UBound(InArray)
            InArray(i - 1) = InArray(i)
        Next
        ReDim Preserve InArray(UBound(InArray) - 1)
    End If
    RemoveFirstElement = InArray
End Function

Function ShellSort(sort As Variant)
span = UBound(sort) \ 2 + 1
Do While span > 0
    For i = span To UBound(sort) + 1
        j = i - span - 1
        For j = (i - span - 1) To 1 Step -span
            If sort(j) <= sort(j + span) Then Exit For
            'swap array elements that are out of order
            temp = sort(j)
            sort(j) = sort(j + span)
            sort(j + span) = temp
            
            temp2 = WindowID(j)
            WindowID(j) = WindowID(j + span)
            WindowID(j + span) = temp2
        Next j
    Next i
    span = span \ 2
Loop
ShellSort = sort
End Function

Private Function GetAWI()
    WindowCaption = Space(256)
    GetWindowText GetForegroundWindow, WindowCaption, Len(WindowCaption)
    WindowCaption = RTrim(WindowCaption)
    Tmp = -1
    For i = 0 To UBound(WindowName)
        If WindowName(i) = WindowCaption Then
            Tmp = i
        End If
    Next
    GetAWI = Tmp
End Function
