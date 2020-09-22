Attribute VB_Name = "Module1"
Public Declare Function RegisterClass Lib "user32" Alias "RegisterClassA" (Class As WNDCLASS) As Long
Public Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Sub PostQuitMessage Lib "user32" (ByVal nExitCode As Long)
Public Declare Function GetMessage Lib "user32" Alias "GetMessageA" (lpMsg As Msg, ByVal hwnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long) As Long
Public Declare Function TranslateMessage Lib "user32" (lpMsg As Msg) As Long
Public Declare Function DispatchMessage Lib "user32" Alias "DispatchMessageA" (lpMsg As Msg) As Long
Public Declare Function UnregisterClass Lib "user32" Alias "UnregisterClassA" (ByVal lpClassName As String, ByVal hInstance As Long) As Long
Public Declare Function LoadIcon Lib "user32" Alias "LoadIconA" (ByVal hInstance As Long, ByVal lpIconName As String) As Long
Public Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As String) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long
Public Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long


Public Type POINTAPI
        x As Long
        y As Long
End Type

Public Const SW_SHOWNORMAL = 1
Public Type Msg
    hwnd As Long
    message As Long
    wParam As Long
    lParam As Long
    time As Long
    pt As POINTAPI
End Type


Public Type WNDCLASS
    style As Long
    lpfnwndproc As Long
    cbClsextra As Long
    cbWndExtra2 As Long
    hInstance As Long
    hIcon As Long
    hCursor As Long
    hbrBackground As Long
    lpszMenuName As String
    lpszClassName As String
End Type

Public Const SW_HIDE = 0
Public Const MB_OK = &H0&
Public Const MB_ICONHAND = &H10&
Public Const ES_MULTILINE = &H4&
Public Const LB_RESETCONTENT = &H184
Public Const LB_ADDSTRING = &H180
Public Const SW_SHOW = 5
Public Const WM_LBUTTONUP = &H202
Public Const GWL_WNDPROC = (-4)
Public Const CW_USEDEFAULT = &H80000000
Public Const WS_CHILD = &H40000000
Public Const IDI_APPLICATION = 32512&
Public Const CS_VREDRAW = &H1
Public Const CS_HREDRAW = &H2
Public Const IDC_ARROW = 32512&
Public Const WS_OVERLAPPED = &H0&
Public Const WS_CAPTION = &HC00000
Public Const WS_BORDER = &H800000
Public Const WS_DLGFRAME = &H400000
Public Const WS_SYSMENU = &H80000
Public Const WS_THICKFRAME = &H40000
Public Const WS_MINIMIZEBOX = &H20000
Public Const WM_DESTROY = &H2
Public Const WS_OVERLAPPEDWINDOW = (WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_MINIMIZEBOX Or WS_BORDER)
Public Const WS_EX_CLIENTEDGE = &H200&
Public Const WM_GETTEXT = &HD
Dim gButOldProc As Long
Dim gButOldProc2 As Long
Dim gButOldProc3 As Long
Dim mListOldProc As Long
Dim mhWnd As Long
Dim mEdit As Long
Dim mList As Long
Dim mButton2 As Long
Dim mButton3 As Long
Dim mStatic1 As Long
Dim mGroup As Long
Dim mStatic2 As Long
Dim mButton As Long
Dim mAbout As Long



Public Sub Main()
Dim szMessage As Msg
    
  'makes sure that the window was created
  If MakeWindow = False Then Exit Sub
              
      'makes sure all the controls were created
      If MakeClass Then
        'do the following functions (api) while the program is running
        'gets all of the windows messages sent to the program
         Do While GetMessage(szMessage, 0&, 0&, 0&)
            'translates the windows messages so our program can read them
            Call TranslateMessage(szMessage)
            'sends all the commands to the default window procedure
            Call DispatchMessage(szMessage)
         Loop
      End If
    
    'when the program ends, unregister our custom class we made
    Call UnregisterClass("CustomClass", App.hInstance)
End Sub

Public Function MakeWindow() As Boolean

Dim mWind As WNDCLASS

'declaring all the variables to register our new class
With mWind

    .hInstance = App.hInstance 'set the class's application instance
    .hCursor = LoadCursor(App.hInstance, IDC_ARROW) 'loads the default cursor
    .hIcon = LoadIcon(App.hInstance, IDI_APPLICATION) 'loads the default icon
    .lpfnwndproc = GetAddress(AddressOf WndProc) 'sets the default window procedure
    .lpszClassName = "CustomClass" 'name of our new class
    .style = CS_VREDRAW Or CS_HREDRAW 'window styles
    .hbrBackground = 5 'default background color
End With
    
    MakeWindow = RegisterClass(mWind) <> 0
    
End Function


Public Function MakeClass() As Boolean
    
    'creates our actual dialog window
    mhWnd = CreateWindowEx(0&, "CustomClass", "API Class Example", WS_OVERLAPPEDWINDOW, CW_USEDEFAULT, CW_USEDEFAULT, 370, 320, 0&, 0&, App.hInstance, ByVal 0&)
    'creates first button
    mButton = CreateWindowEx(0&, "Button", "&Add to list", WS_CHILD, 20, 20, 100, 35, mhWnd, 0&, App.hInstance, 0&)
    'creates second button
    mButton2 = CreateWindowEx(0&, "Button", "&Clear list", WS_CHILD, 130, 20, 100, 35, mhWnd, 0&, App.hInstance, 0&)
    'creates third button
    mButton3 = CreateWindowEx(0&, "Button", "A&bout", WS_CHILD, 240, 20, 100, 35, mhWnd, 0&, App.hInstance, 0&)
    'creates static (label)
    mStatic1 = CreateWindowEx(0&, "Static", "Add List Item:", WS_CHILD, 190, 70, 100, 30, mhWnd, 0&, App.hInstance, 0&)
    'creates an edit (vb's know it as textbox)
    mEdit = CreateWindowEx(WS_EX_CLIENTEDGE, "Edit", "New item for Listbox", WS_CHILD Or ES_MULTILINE, 190, 90, 150, 100, mhWnd, 0&, App.hInstance, 0&)
    'creates a listbox
    mList = CreateWindowEx(WS_EX_CLIENTEDGE, "Listbox", "", WS_CHILD, 20, 70, 150, 200, mhWnd, 0&, App.hInstance, 0&)
    'creates our second static (label)
    mStatic2 = CreateWindowEx(0&, "Static", "API Class example by Armen Shimoon. Code free to use by anybody.", WS_CHILD, 190, 200, 150, 75, mhWnd, 0&, App.hInstance, 0&)
    
    
    'now we make all of the "windows" visible
    Call ShowWindow(mhWnd, SW_SHOWNORMAL)
    Call ShowWindow(mButton, SW_SHOWNORMAL)
    Call ShowWindow(mEdit, SW_SHOWNORMAL)
    Call ShowWindow(mList, SW_SHOWNORMAL)
    Call ShowWindow(mButton2, SW_SHOWNORMAL)
    Call ShowWindow(mStatic1, SW_SHOWNORMAL)
    Call ShowWindow(mButton3, SW_SHOWNORMAL)
    Call ShowWindow(mStatic2, SW_SHOWNORMAL)

    'this adds the first string to our listbox
    Call SendMessage(mList, LB_ADDSTRING, 7, ByVal "Listbox")
    
    
    'this function retreives the address of the default window procedures of the buttons
    gButOldProc = GetWindowLong(mButton, GWL_WNDPROC)
    gButOldProc2 = GetWindowLong(mButton2, GWL_WNDPROC)
    gButOldProc3 = GetWindowLong(mButton3, GWL_WNDPROC)
    
    'this function now sets a new window procedure to handle windows messages
    Call SetWindowLong(mButton, GWL_WNDPROC, GetAddress(AddressOf ButtonWndProc))
    Call SetWindowLong(mButton2, GWL_WNDPROC, GetAddress(AddressOf ButtonWndProc2))
    Call SetWindowLong(mButton3, GWL_WNDPROC, GetAddress(AddressOf ButtonWndProc3))
   
    
    
    'any integer higher than 0 is true
    MakeClass = (mhWnd <> 0)
End Function

Public Function ButtonWndProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim ttxt As String
    'makes a case to handle the windows messages
    Select Case uMsg&
       'when the user left-mouse button clicks a the first button
       Case WM_LBUTTONUP:
              'get the text from the edit
              ttxt = Space(100)
              Call SendMessage(mEdit, WM_GETTEXT, ByVal 100, ByVal ttxt)
              ttxt = RTrim(ttxt)
              'add that text to the listbox
              Call SendMessage(mList, LB_ADDSTRING, ByVal Len(ttxt), ByVal ttxt)
       End Select

'since we only want to handle 1 message sent by windows, we change the procedure back to the default
ButtonWndProc = CallWindowProc(gButOldProc, hwnd&, uMsg&, wParam&, lParam&)
End Function

Public Function ButtonWndProc2(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim lbCount As Integer
    'case to handle the windows messages
    Select Case uMsg&
       Case WM_LBUTTONUP:
            'clears the listbox
            Call SendMessage(mList, LB_RESETCONTENT, 0&, 0&)
    End Select
'set the default window procedure
ButtonWndProc2 = CallWindowProc(gButOldProc2, hwnd&, uMsg&, wParam&, lParam&)
End Function

Public Function ButtonWndProc3(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim lbCount As Integer
Dim aStatic1 As Long

    Select Case uMsg&
       Case WM_LBUTTONUP:
            'when the user clicks the about button, we want to make a new window with "About" information
            mAbout = CreateWindowEx(0&, "CustomClass", "About API Class", WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU, CW_USEDEFAULT, CW_USEDEFAULT, 200, 120, 0&, 0&, App.hInstance, ByVal 0&)
            aStatic1 = CreateWindowEx(0&, "Static", "Written by Armen Shimoon. Free code to use anywhere by anybody. Please remember to vote! ©2001 Shimoon.", WS_CHILD, 10, 5, 194, 114, mAbout, 0&, App.hInstance, 0&)
            Call ShowWindow(mAbout, SW_SHOWNORMAL)
            Call ShowWindow(aStatic1, SW_SHOWNORMAL)
    End Select
'set the default window procedure
ButtonWndProc3 = CallWindowProc(gButOldProc2, hwnd&, uMsg&, wParam&, lParam&)
End Function

Public Function WndProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim strTemp As String
    Select Case uMsg&
       'when the user exits the program this is called
       Case WM_DESTROY:
          'when the user closes the "About" dialog, it sends a WM_DESTROY msg to here aswell
          'if the window that sent the close message is the about window, just ignore it and hide the window
          If hwnd = mAbout Then
          Call ShowWindow(mAbout, SW_HIDE)
          'else means it was sent by the main window, which means we quit
          Else
          'tells windows we wish to quit
          Call PostQuitMessage(0&)
          End If
       End Select
'sets the default window procedure
WndProc = DefWindowProc(hwnd&, uMsg&, wParam&, lParam&)
End Function



Public Function GetAddress(ByVal lpAddr As Long) As Long
'this gives us a pointer to the function called
GetAddress = lpAddr&
End Function
