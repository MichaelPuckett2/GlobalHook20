Imports System.Runtime.InteropServices

''' <summary>
''' The Base Class of Every Hook Class within GlobalHook 2.0
''' </summary>
''' <remarks></remarks>
Public MustInherit Class HookBase
    Implements IDisposable

    Public Const I_LOVE_GOD_FOREVER As String = "Thank you God for all the wisdom, love, mercy, and forgiveness you have given me.  " &
                                                "Thank you for the talent to do what all I can do and I pray that I do it all to gloriy you.  " &
                                                "Thank you so much for my family.  I love you and praise you and lift your name higher forever and ever. I love you God with all my heart."

    ''' <summary>
    ''' The value used to store the HookType.
    ''' </summary>
    ''' <remarks></remarks>
    Private ReadOnly idHook As HookType

    ''' <summary>
    ''' Initializes the HookBase Class and determines the type of hook that will be installed.
    ''' </summary>
    ''' <param name="_idHook">The HookType to be installed.</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal _idHook As HookType)
        idHook = _idHook
    End Sub

    ''' <summary>
    ''' The Handle to our Application instantiated in the constructor. This is passed in to the SetWindowHookEx as the hMod parameter.
    ''' </summary>
    ''' <remarks></remarks>
    Private Shared Handle As IntPtr = Process.GetCurrentProcess().MainModule.BaseAddress

    ''' <summary>
    ''' The handle that SetWindowsHookEx returns to us as a pointer to our Hook from within the user32.dll.
    ''' </summary>
    ''' <remarks></remarks>
    Private HookHandle As IntPtr = Nothing

    ''' <summary>
    ''' The Delegate passed in to the SetWindowsHookEx as the callback variable.
    ''' [in] Pointer to the hook procedure. If the dwThreadId parameter is zero or specifies the identifier of a thread created by a different process, the lpfn parameter must point to a hook procedure in a DLL.
    ''' Otherwise, lpfn can point to a hook procedure in the code associated with the current process. 
    ''' </summary>
    ''' <param name="nCode">[in] Specifies a code the hook procedure uses to determine how to process the message. If nCode is less than zero, the hook procedure must pass the message to the CallNextHookEx function without further processing and should return the value returned by CallNextHookEx. This parameter can be one of the following values. 
    ''' HC_ACTION
    ''' The wParam and lParam parameters contain information about a keyboard message.</param>
    ''' <param name="wParam">[in] Specifies the identifier of the keyboard message. This parameter can be one of the following messages: WM_KEYDOWN, WM_KEYUP, WM_SYSKEYDOWN, or WM_SYSKEYUP. </param>
    ''' <param name="lParam">[in] Pointer to a KBDLLHOOKSTRUCT structure. </param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Delegate Function HookProc(ByVal nCode As Integer, ByVal wParam As IntPtr, <[In]()> ByVal lParam As IntPtr) As IntPtr

    ''' <summary>
    ''' The SetWindowsHookEx function installs an application-defined hook procedure into a hook chain. You would install a hook procedure to monitor the system for certain types of events.
    ''' These events are associated either with a specific thread or with all threads in the same desktop as the calling thread. 
    ''' </summary>
    ''' <param name="hook">[in] Specifies the type of hook procedure to be installed. This parameter can be one of the following values. 
    ''' WH_CALLWNDPROC
    ''' Installs a hook procedure that monitors messages before the system sends them to the destination window procedure. For more information, see the CallWndProc hook procedure.
    ''' WH_CALLWNDPROCRET
    ''' Installs a hook procedure that monitors messages after they have been processed by the destination window procedure. For more information, see the CallWndRetProc hook procedure.
    ''' WH_CBT
    ''' Installs a hook procedure that receives notifications useful to a computer-based training (CBT) application. For more information, see the CBTProc hook procedure.
    ''' WH_DEBUG
    ''' Installs a hook procedure useful for debugging other hook procedures. For more information, see the DebugProc hook procedure.
    ''' WH_FOREGROUNDIDLE
    ''' Installs a hook procedure that will be called when the application's foreground thread is about to become idle. This hook is useful for performing low priority tasks during idle time. For more information, see the ForegroundIdleProc hook procedure. 
    ''' WH_GETMESSAGE
    ''' Installs a hook procedure that monitors messages posted to a message queue. For more information, see the GetMsgProc hook procedure.
    ''' WH_JOURNALPLAYBACK
    ''' Installs a hook procedure that posts messages previously recorded by a WH_JOURNALRECORD hook procedure. For more information, see the JournalPlaybackProc hook procedure.
    ''' WH_JOURNALRECORD
    ''' Installs a hook procedure that records input messages posted to the system message queue. This hook is useful for recording macros. For more information, see the JournalRecordProc hook procedure.
    ''' WH_KEYBOARD
    ''' Installs a hook procedure that monitors keystroke messages. For more information, see the KeyboardProc hook procedure.
    ''' WH_KEYBOARD_LL
    ''' Windows NT/2000/XP: Installs a hook procedure that monitors low-level keyboard input events. For more information, see the LowLevelKeyboardProc hook procedure.
    ''' WH_MOUSE
    ''' Installs a hook procedure that monitors mouse messages. For more information, see the MouseProc hook procedure.
    ''' WH_MOUSE_LL
    ''' Windows NT/2000/XP: Installs a hook procedure that monitors low-level mouse input events. For more information, see the LowLevelMouseProc hook procedure.
    ''' WH_MSGFILTER
    ''' Installs a hook procedure that monitors messages generated as a result of an input event in a dialog box, message box, menu, or scroll bar. For more information, see the MessageProc hook procedure.
    ''' WH_SHELL
    ''' Installs a hook procedure that receives notifications useful to shell applications. For more information, see the ShellProc hook procedure.
    ''' WH_SYSMSGFILTER
    ''' Installs a hook procedure that monitors messages generated as a result of an input event in a dialog box, message box, menu, or scroll bar. The hook procedure monitors these messages for all applications in the same desktop as the calling thread. For more information, see the SysMsgProc hook procedure.</param>
    ''' <param name="lpfn">[in] Pointer to the hook procedure. If the dwThreadId parameter is zero or specifies the identifier of a thread created by a different process, the lpfn parameter must point to a hook procedure in a DLL. Otherwise, lpfn can point to a hook procedure in the code associated with the current process. </param>
    ''' <param name="hMod">[in] Handle to the DLL containing the hook procedure pointed to by the lpfn parameter. The hMod parameter must be set to NULL if the dwThreadId parameter specifies a thread created by the current process and if the hook procedure is within the code associated with the current process. </param>
    ''' <param name="dwThreadId">[in] Specifies the identifier of the thread with which the hook procedure is to be associated. If this parameter is zero, the hook procedure is associated with all existing threads running in the same desktop as the calling thread. </param>
    ''' <returns>If the function succeeds, the return value is the handle to the hook procedure.
    ''' If the function fails, the return value is NULL. To get extended error information, call GetLastError.</returns>
    ''' <remarks></remarks>
    <DllImport("user32.dll")>
    Private Shared Function SetWindowsHookEx(ByVal hook As Integer, ByVal lpfn As HookProc, ByVal hMod As IntPtr, ByVal dwThreadId As UInteger) As IntPtr
    End Function

    ''' <summary>
    ''' The CallNextHookEx function passes the hook information to the next hook procedure in the current hook chain. A hook procedure can call this function either before or after processing the hook information. 
    ''' </summary>
    ''' <param name="hhk">[in] Windows 95/98/ME: Handle to the current hook. An application receives this handle as a result of a previous call to the SetWindowsHookEx function. 
    ''' Windows NT/XP/2003: Ignored.</param>
    ''' <param name="nCode">[in] Specifies the hook code passed to the current hook procedure. The next hook procedure uses this code to determine how to process the hook information.</param>
    ''' <param name="wParam">[in] Specifies the wParam value passed to the current hook procedure. The meaning of this parameter depends on the type of hook associated with the current hook chain.</param>
    ''' <param name="lParam">[in] Specifies the lParam value passed to the current hook procedure. The meaning of this parameter depends on the type of hook associated with the current hook chain.</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <DllImport("user32.dll")>
    Private Shared Function CallNextHookEx(ByVal hhk As IntPtr, ByVal nCode As Integer, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As IntPtr
    End Function

    ''' <summary>
    ''' The UnhookWindowsHookEx function removes a hook procedure installed in a hook chain by the SetWindowsHookEx function. 
    ''' </summary>
    ''' <param name="hhk">[in] Handle to the hook to be removed. This parameter is a hook handle obtained by a previous call to SetWindowsHookEx.</param>
    ''' <returns>If the function succeeds, the return value is nonzero.
    ''' If the function fails, the return value is zero. To get extended error information, call GetLastError.</returns>
    ''' <remarks></remarks>
    <DllImport("user32.dll")>
    Private Shared Function UnhookWindowsHookEx(ByVal hhk As IntPtr) As Boolean
    End Function

    ''' <summary>
    ''' The variable used to instantiate the HookProc Delegate. When instantiated it should refer to the OnHookProc Method.
    ''' </summary>
    ''' <remarks></remarks>
    Private HookProc_Delegated As HookProc = New HookProc(AddressOf MainHookProc)

    ''' <summary>
    ''' Enumeration of WindowsMessages Retrieved from http://pinvoke.net/default.aspx/Enums/WindowsMessages.html
    ''' </summary>
    ''' <remarks>At early writing this DLL was built with custom Enumerations for each class but after reviewing the extensive work put in place for this Enumberation to include
    ''' commenting it was decided this would be the best way to go.  It has been adopted at PInvoke.Net and is currently being used by many developers world wide although not ever value in this enumeration
    ''' has been tested here during the building of this DLL.  The values used in our Hook Classes however have been tested.</remarks>
    Public Enum WindowsMessages
        ''' <summary>
        '''The WM_ACTIVATE message is sent when a window is being activated or deactivated. This message is sent first to the window procedure of the top-level window being deactivated; it is then sent to the window procedure of the top-level window being activated.
        ''' </summary>
        ''' <remarks></remarks>
        WM_ACTIVATE = &H6

        ''' <summary>
        '''The WM_ACTIVATEAPP message is sent when a window belonging to a different application than the active window is about to be activated. The message is sent to the application whose window is being activated and to the application whose window is being deactivated.
        ''' </summary>
        ''' <remarks></remarks>
        WM_ACTIVATEAPP = &H1C

        ''' <summary>
        '''Definition Needed
        ''' </summary>
        ''' <remarks></remarks>
        WM_AFXFIRST = &H360

        ''' <summary>
        '''Definition Needed
        ''' </summary>
        ''' <remarks></remarks>
        WM_AFXLAST = &H37F

        ''' <summary>
        '''The WM_APP constant is used by applications to help define private messages usually of the form WM_APP+X where X is an integer value.
        ''' </summary>
        ''' <remarks></remarks>
        WM_APP = &H8000

        ''' <summary>
        '''The WM_ASKCBFORMATNAME message is sent to the clipboard owner by a clipboard viewer window to request the name of a CF_OWNERDISPLAY clipboard format.
        ''' </summary>
        ''' <remarks></remarks>
        WM_ASKCBFORMATNAME = &H30C

        ''' <summary>
        '''The WM_CANCELJOURNAL message is posted to an application when a user cancels the application's journaling activities. The message is posted with a NULL window handle.
        ''' </summary>
        ''' <remarks></remarks>
        WM_CANCELJOURNAL = &H4B

        ''' <summary>
        '''The WM_CANCELMODE message is sent to cancel certain modes such as mouse capture. For example the system sends this message to the active window when a dialog box or message box is displayed. Certain functions also send this message explicitly to the specified window regardless of whether it is the active window. For example the EnableWindow function sends this message when disabling the specified window.
        ''' </summary>
        ''' <remarks></remarks>
        WM_CANCELMODE = &H1F

        ''' <summary>
        '''The WM_CAPTURECHANGED message is sent to the window that is losing the mouse capture.
        ''' </summary>
        ''' <remarks></remarks>
        WM_CAPTURECHANGED = &H215

        ''' <summary>
        '''The WM_CHANGECBCHAIN message is sent to the first window in the clipboard viewer chain when a window is being removed from the chain.
        ''' </summary>
        ''' <remarks></remarks>
        WM_CHANGECBCHAIN = &H30D

        ''' <summary>
        '''An application sends the WM_CHANGEUISTATE message to indicate that the user interface (UI) state should be changed.
        ''' </summary>
        ''' <remarks></remarks>
        WM_CHANGEUISTATE = &H127

        ''' <summary>
        '''The WM_CHAR message is posted to the window with the keyboard focus when a WM_KEYDOWN message is translated by the TranslateMessage function. The WM_CHAR message contains the character code of the key that was pressed.
        ''' </summary>
        ''' <remarks></remarks>
        WM_CHAR = &H102

        ''' <summary>
        '''Sent by a list box with the LBS_WANTKEYBOARDINPUT style to its owner in response to a WM_CHAR message.
        ''' </summary>
        ''' <remarks></remarks>
        WM_CHARTOITEM = &H2F

        ''' <summary>
        '''The WM_CHILDACTIVATE message is sent to a child window when the user clicks the window's title bar or when the window is activated moved or sized.
        ''' </summary>
        ''' <remarks></remarks>
        WM_CHILDACTIVATE = &H22

        ''' <summary>
        '''An application sends a WM_CLEAR message to an edit control or combo box to delete (clear) the current selection if any from the edit control.
        ''' </summary>
        ''' <remarks></remarks>
        WM_CLEAR = &H303

        ''' <summary>
        '''The WM_CLOSE message is sent as a signal that a window or an application should terminate.
        ''' </summary>
        ''' <remarks></remarks>
        WM_CLOSE = &H10

        ''' <summary>
        '''The WM_COMMAND message is sent when the user selects a command item from a menu when a control sends a notification message to its parent window or when an accelerator keystroke is translated.
        ''' </summary>
        ''' <remarks></remarks>
        WM_COMMAND = &H111

        ''' <summary>
        '''The WM_COMPACTING message is sent to all top-level windows when the system detects more than 12.5 percent of system time over a 30- to 60-second interval is being spent compacting memory. This indicates that system memory is low.
        ''' </summary>
        ''' <remarks></remarks>
        WM_COMPACTING = &H41

        ''' <summary>
        '''The system sends the WM_COMPAREITEM message to determine the relative position of a new item in the sorted list of an owner-drawn combo box or list box. Whenever the application adds a new item the system sends this message to the owner of a combo box or list box created with the CBS_SORT or LBS_SORT style.
        ''' </summary>
        ''' <remarks></remarks>
        WM_COMPAREITEM = &H39

        ''' <summary>
        '''The WM_CONTEXTMENU message notifies a window that the user clicked the right mouse button (right-clicked) in the window.
        ''' </summary>
        ''' <remarks></remarks>
        WM_CONTEXTMENU = &H7B

        ''' <summary>
        '''An application sends the WM_COPY message to an edit control or combo box to copy the current selection to the clipboard in CF_TEXT format.
        ''' </summary>
        ''' <remarks></remarks>
        WM_COPY = &H301

        ''' <summary>
        '''An application sends the WM_COPYDATA message to pass data to another application.
        ''' </summary>
        ''' <remarks></remarks>
        WM_COPYDATA = &H4A

        ''' <summary>
        '''The WM_CREATE message is sent when an application requests that a window be created by calling the CreateWindowEx or CreateWindow function. (The message is sent before the function returns.) The window procedure of the new window receives this message after the window is created but before the window becomes visible.
        ''' </summary>
        ''' <remarks></remarks>
        WM_CREATE = &H1

        ''' <summary>
        '''The WM_CTLCOLORBTN message is sent to the parent window of a button before drawing the button. The parent window can change the button's text and background colors. However only owner-drawn buttons respond to the parent window processing this message.
        ''' </summary>
        ''' <remarks></remarks>
        WM_CTLCOLORBTN = &H135

        ''' <summary>
        '''The WM_CTLCOLORDLG message is sent to a dialog box before the system draws the dialog box. By responding to this message the dialog box can set its text and background colors using the specified display device context handle.
        ''' </summary>
        ''' <remarks></remarks>
        WM_CTLCOLORDLG = &H136

        ''' <summary>
        '''An edit control that is not read-only or disabled sends the WM_CTLCOLOREDIT message to its parent window when the control is about to be drawn. By responding to this message the parent window can use the specified device context handle to set the text and background colors of the edit control.
        ''' </summary>
        ''' <remarks></remarks>
        WM_CTLCOLOREDIT = &H133

        ''' <summary>
        '''Sent to the parent window of a list box before the system draws the list box. By responding to this message the parent window can set the text and background colors of the list box by using the specified display device context handle.
        ''' </summary>
        ''' <remarks></remarks>
        WM_CTLCOLORLISTBOX = &H134

        ''' <summary>
        '''The WM_CTLCOLORMSGBOX message is sent to the owner window of a message box before Windows draws the message box. By responding to this message the owner window can set the text and background colors of the message box by using the given display device context handle.
        ''' </summary>
        ''' <remarks></remarks>
        WM_CTLCOLORMSGBOX = &H132

        ''' <summary>
        '''The WM_CTLCOLORSCROLLBAR message is sent to the parent window of a scroll bar control when the control is about to be drawn. By responding to this message the parent window can use the display context handle to set the background color of the scroll bar control.
        ''' </summary>
        ''' <remarks></remarks>
        WM_CTLCOLORSCROLLBAR = &H137

        ''' <summary>
        '''A static control or an edit control that is read-only or disabled sends the WM_CTLCOLORSTATIC message to its parent window when the control is about to be drawn. By responding to this message the parent window can use the specified device context handle to set the text and background colors of the static control.
        ''' </summary>
        ''' <remarks></remarks>
        WM_CTLCOLORSTATIC = &H138

        ''' <summary>
        '''An application sends a WM_CUT message to an edit control or combo box to delete (cut) the current selection if any in the edit control and copy the deleted text to the clipboard in CF_TEXT format.
        ''' </summary>
        ''' <remarks></remarks>
        WM_CUT = &H300

        ''' <summary>
        '''The WM_DEADCHAR message is posted to the window with the keyboard focus when a WM_KEYUP message is translated by the TranslateMessage function. WM_DEADCHAR specifies a character code generated by a dead key. A dead key is a key that generates a character such as the umlaut (double-dot) that is combined with another character to form a composite character. For example the umlaut-O character (?) is generated by typing the dead key for the umlaut character and then typing the O key.
        ''' </summary>
        ''' <remarks></remarks>
        WM_DEADCHAR = &H103

        ''' <summary>
        '''Sent to the owner of a list box or combo box when the list box or combo box is destroyed or when items are removed by the LB_DELETESTRING LB_RESETCONTENT CB_DELETESTRING or CB_RESETCONTENT message. The system sends a WM_DELETEITEM message for each deleted item. The system sends the WM_DELETEITEM message for any deleted list box or combo box item with nonzero item data.
        ''' </summary>
        ''' <remarks></remarks>
        WM_DELETEITEM = &H2D

        ''' <summary>
        '''The WM_DESTROY message is sent when a window is being destroyed. It is sent to the window procedure of the window being destroyed after the window is removed from the screen. This message is sent first to the window being destroyed and then to the child windows (if any) as they are destroyed. During the processing of the message it can be assumed that all child windows still exist.
        ''' </summary>
        ''' <remarks></remarks>
        WM_DESTROY = &H2

        ''' <summary>
        '''The WM_DESTROYCLIPBOARD message is sent to the clipboard owner when a call to the EmptyClipboard function empties the clipboard.
        ''' </summary>
        ''' <remarks></remarks>
        WM_DESTROYCLIPBOARD = &H307

        ''' <summary>
        '''Notifies an application of a change to the hardware configuration of a device or the computer.
        ''' </summary>
        ''' <remarks></remarks>
        WM_DEVICECHANGE = &H219

        ''' <summary>
        '''The WM_DEVMODECHANGE message is sent to all top-level windows whenever the user changes device-mode settings.
        ''' </summary>
        ''' <remarks></remarks>
        WM_DEVMODECHANGE = &H1B

        ''' <summary>
        '''The WM_DISPLAYCHANGE message is sent to all windows when the display resolution has changed.
        ''' </summary>
        ''' <remarks></remarks>
        WM_DISPLAYCHANGE = &H7E

        ''' <summary>
        '''The WM_DRAWCLIPBOARD message is sent to the first window in the clipboard viewer chain when the content of the clipboard changes. This enables a clipboard viewer window to display the new content of the clipboard.
        ''' </summary>
        ''' <remarks></remarks>
        WM_DRAWCLIPBOARD = &H308

        ''' <summary>
        '''The WM_DRAWITEM message is sent to the parent window of an owner-drawn button combo box list box or menu when a visual aspect of the button combo box list box or menu has changed.
        ''' </summary>
        ''' <remarks></remarks>
        WM_DRAWITEM = &H2B

        ''' <summary>
        '''Sent when the user drops a file on the window of an application that has registered itself as a recipient of dropped files.
        ''' </summary>
        ''' <remarks></remarks>
        WM_DROPFILES = &H233

        ''' <summary>
        '''The WM_ENABLE message is sent when an application changes the enabled state of a window. It is sent to the window whose enabled state is changing. This message is sent before the EnableWindow function returns but after the enabled state (WS_DISABLED style bit) of the window has changed.
        ''' </summary>
        ''' <remarks></remarks>
        WM_ENABLE = &HA

        ''' <summary>
        '''The WM_ENDSESSION message is sent to an application after the system processes the results of the WM_QUERYENDSESSION message. The WM_ENDSESSION message informs the application whether the session is ending.
        ''' </summary>
        ''' <remarks></remarks>
        WM_ENDSESSION = &H16

        ''' <summary>
        '''The WM_ENTERIDLE message is sent to the owner window of a modal dialog box or menu that is entering an idle state. A modal dialog box or menu enters an idle state when no messages are waiting in its queue after it has processed one or more previous messages.
        ''' </summary>
        ''' <remarks></remarks>
        WM_ENTERIDLE = &H121

        ''' <summary>
        '''The WM_ENTERMENULOOP message informs an application's main window procedure that a menu modal loop has been entered.
        ''' </summary>
        ''' <remarks></remarks>
        WM_ENTERMENULOOP = &H211

        ''' <summary>
        '''The WM_ENTERSIZEMOVE message is sent one time to a window after it enters the moving or sizing modal loop. The window enters the moving or sizing modal loop when the user clicks the window's title bar or sizing border or when the window passes the WM_SYSCOMMAND message to the DefWindowProc function and the wParam parameter of the message specifies the SC_MOVE or SC_SIZE value. The operation is complete when DefWindowProc returns.
        ''' </summary>
        ''' <remarks></remarks>
        WM_ENTERSIZEMOVE = &H231

        ''' <summary>
        '''The WM_ERASEBKGND message is sent when the window background must be erased (for example when a window is resized). The message is sent to prepare an invalidated portion of a window for painting.
        ''' </summary>
        ''' <remarks></remarks>
        WM_ERASEBKGND = &H14

        ''' <summary>
        '''The WM_EXITMENULOOP message informs an application's main window procedure that a menu modal loop has been exited.
        ''' </summary>
        ''' <remarks></remarks>
        WM_EXITMENULOOP = &H212

        ''' <summary>
        '''The WM_EXITSIZEMOVE message is sent one time to a window after it has exited the moving or sizing modal loop. The window enters the moving or sizing modal loop when the user clicks the window's title bar or sizing border or when the window passes the WM_SYSCOMMAND message to the DefWindowProc function and the wParam parameter of the message specifies the SC_MOVE or SC_SIZE value. The operation is complete when DefWindowProc returns.
        ''' </summary>
        ''' <remarks></remarks>
        WM_EXITSIZEMOVE = &H232

        ''' <summary>
        '''An application sends the WM_FONTCHANGE message to all top-level windows in the system after changing the pool of font resources.
        ''' </summary>
        ''' <remarks></remarks>
        WM_FONTCHANGE = &H1D

        ''' <summary>
        '''The WM_GETDLGCODE message is sent to the window procedure associated with a control. By default the system handles all keyboard input to the control; the system interprets certain types of keyboard input as dialog box navigation keys. To override this default behavior the control can respond to the WM_GETDLGCODE message to indicate the types of input it wants to process itself.
        ''' </summary>
        ''' <remarks></remarks>
        WM_GETDLGCODE = &H87

        ''' <summary>
        '''An application sends a WM_GETFONT message to a control to retrieve the font with which the control is currently drawing its text.
        ''' </summary>
        ''' <remarks></remarks>
        WM_GETFONT = &H31

        ''' <summary>
        '''An application sends a WM_GETHOTKEY message to determine the hot key associated with a window.
        ''' </summary>
        ''' <remarks></remarks>
        WM_GETHOTKEY = &H33

        ''' <summary>
        '''The WM_GETICON message is sent to a window to retrieve a handle to the large or small icon associated with a window. The system displays the large icon in the ALT+TAB dialog and the small icon in the window caption.
        ''' </summary>
        ''' <remarks></remarks>
        WM_GETICON = &H7F

        ''' <summary>
        '''The WM_GETMINMAXINFO message is sent to a window when the size or position of the window is about to change. An application can use this message to override the window's default maximized size and position or its default minimum or maximum tracking size.
        ''' </summary>
        ''' <remarks></remarks>
        WM_GETMINMAXINFO = &H24

        ''' <summary>
        '''Active Accessibility sends the WM_GETOBJECT message to obtain information about an accessible object contained in a server application. Applications never send this message directly. It is sent only by Active Accessibility in response to calls to AccessibleObjectFromPoint AccessibleObjectFromEvent or AccessibleObjectFromWindow. However server applications handle this message.
        ''' </summary>
        ''' <remarks></remarks>
        WM_GETOBJECT = &H3D

        ''' <summary>
        '''An application sends a WM_GETTEXT message to copy the text that corresponds to a window into a buffer provided by the caller.
        ''' </summary>
        ''' <remarks></remarks>
        WM_GETTEXT = &HD

        ''' <summary>
        '''An application sends a WM_GETTEXTLENGTH message to determine the length in characters of the text associated with a window.
        ''' </summary>
        ''' <remarks></remarks>
        WM_GETTEXTLENGTH = &HE

        ''' <summary>
        '''Definition Needed
        ''' </summary>
        ''' <remarks></remarks>
        WM_HANDHELDFIRST = &H358

        ''' <summary>
        '''Definition Needed
        ''' </summary>
        ''' <remarks></remarks>
        WM_HANDHELDLAST = &H35F

        ''' <summary>
        '''Indicates that the user pressed the F1 key. If a menu is active when F1 is pressed WM_HELP is sent to the window associated with the menu; otherwise WM_HELP is sent to the window that has the keyboard focus. If no window has the keyboard focus WM_HELP is sent to the currently active window.
        ''' </summary>
        ''' <remarks></remarks>
        WM_HELP = &H53

        ''' <summary>
        '''The WM_HOTKEY message is posted when the user presses a hot key registered by the RegisterHotKey function. The message is placed at the top of the message queue associated with the thread that registered the hot key.
        ''' </summary>
        ''' <remarks></remarks>
        WM_HOTKEY = &H312

        ''' <summary>
        '''This message is sent to a window when a scroll event occurs in the window's standard horizontal scroll bar. This message is also sent to the owner of a horizontal scroll bar control when a scroll event occurs in the control.
        ''' </summary>
        ''' <remarks></remarks>
        WM_HSCROLL = &H114

        ''' <summary>
        '''The WM_HSCROLLCLIPBOARD message is sent to the clipboard owner by a clipboard viewer window. This occurs when the clipboard contains data in the CF_OWNERDISPLAY format and an event occurs in the clipboard viewer's horizontal scroll bar. The owner should scroll the clipboard image and update the scroll bar values.
        ''' </summary>
        ''' <remarks></remarks>
        WM_HSCROLLCLIPBOARD = &H30E

        ''' <summary>
        '''Windows NT 3.51 and earlier: The WM_ICONERASEBKGND message is sent to a minimized window when the background of the icon must be filled before painting the icon. A window receives this message only if a class icon is defined for the window; otherwise WM_ERASEBKGND is sent. This message is not sent by newer versions of Windows.
        ''' </summary>
        ''' <remarks></remarks>
        WM_ICONERASEBKGND = &H27

        ''' <summary>
        '''Sent to an application when the IME gets a character of the conversion result. A window receives this message through its WindowProc function.
        ''' </summary>
        ''' <remarks></remarks>
        WM_IME_CHAR = &H286

        ''' <summary>
        '''Sent to an application when the IME changes composition status as a result of a keystroke. A window receives this message through its WindowProc function.
        ''' </summary>
        ''' <remarks></remarks>
        WM_IME_COMPOSITION = &H10F

        ''' <summary>
        '''Sent to an application when the IME window finds no space to extend the area for the composition window. A window receives this message through its WindowProc function.
        ''' </summary>
        ''' <remarks></remarks>
        WM_IME_COMPOSITIONFULL = &H284

        ''' <summary>
        '''Sent by an application to direct the IME window to carry out the requested command. The application uses this message to control the IME window that it has created. To send this message the application calls the SendMessage function with the following parameters.
        ''' </summary>
        ''' <remarks></remarks>
        WM_IME_CONTROL = &H283

        ''' <summary>
        '''Sent to an application when the IME ends composition. A window receives this message through its WindowProc function.
        ''' </summary>
        ''' <remarks></remarks>
        WM_IME_ENDCOMPOSITION = &H10E

        ''' <summary>
        '''Sent to an application by the IME to notify the application of a key press and to keep message order. A window receives this message through its WindowProc function.
        ''' </summary>
        ''' <remarks></remarks>
        WM_IME_KEYDOWN = &H290

        ''' <summary>
        '''Definition Needed
        ''' </summary>
        ''' <remarks></remarks>
        WM_IME_KEYLAST = &H10F

        ''' <summary>
        '''Sent to an application by the IME to notify the application of a key release and to keep message order. A window receives this message through its WindowProc function.
        ''' </summary>
        ''' <remarks></remarks>
        WM_IME_KEYUP = &H291

        ''' <summary>
        '''Sent to an application to notify it of changes to the IME window. A window receives this message through its WindowProc function.
        ''' </summary>
        ''' <remarks></remarks>
        WM_IME_NOTIFY = &H282

        ''' <summary>
        '''Sent to an application to provide commands and request information. A window receives this message through its WindowProc function.
        ''' </summary>
        ''' <remarks></remarks>
        WM_IME_REQUEST = &H288

        ''' <summary>
        '''Sent to an application when the operating system is about to change the current IME. A window receives this message through its WindowProc function.
        ''' </summary>
        ''' <remarks></remarks>
        WM_IME_SELECT = &H285

        ''' <summary>
        '''Sent to an application when a window is activated. A window receives this message through its WindowProc function.
        ''' </summary>
        ''' <remarks></remarks>
        WM_IME_SETCONTEXT = &H281

        ''' <summary>
        '''Sent immediately before the IME generates the composition string as a result of a keystroke. A window receives this message through its WindowProc function.
        ''' </summary>
        ''' <remarks></remarks>
        WM_IME_STARTCOMPOSITION = &H10D

        ''' <summary>
        '''The WM_INITDIALOG message is sent to the dialog box procedure immediately before a dialog box is displayed. Dialog box procedures typically use this message to initialize controls and carry out any other initialization tasks that affect the appearance of the dialog box.
        ''' </summary>
        ''' <remarks></remarks>
        WM_INITDIALOG = &H110

        ''' <summary>
        '''The WM_INITMENU message is sent when a menu is about to become active. It occurs when the user clicks an item on the menu bar or presses a menu key. This allows the application to modify the menu before it is displayed.
        ''' </summary>
        ''' <remarks></remarks>
        WM_INITMENU = &H116

        ''' <summary>
        '''The WM_INITMENUPOPUP message is sent when a drop-down menu or submenu is about to become active. This allows an application to modify the menu before it is displayed without changing the entire menu.
        ''' </summary>
        ''' <remarks></remarks>
        WM_INITMENUPOPUP = &H117

        ''' <summary>
        '''The WM_INPUTLANGCHANGE message is sent to the topmost affected window after an application's input language has been changed. You should make any application-specific settings and pass the message to the DefWindowProc function which passes the message to all first-level child windows. These child windows can pass the message to DefWindowProc to have it pass the message to their child windows and so on.
        ''' </summary>
        ''' <remarks></remarks>
        WM_INPUTLANGCHANGE = &H51

        ''' <summary>
        '''The WM_INPUTLANGCHANGEREQUEST message is posted to the window with the focus when the user chooses a new input language either with the hotkey (specified in the Keyboard control panel application) or from the indicator on the system taskbar. An application can accept the change by passing the message to the DefWindowProc function or reject the change (and prevent it from taking place) by returning immediately.
        ''' </summary>
        ''' <remarks></remarks>
        WM_INPUTLANGCHANGEREQUEST = &H50

        ''' <summary>
        '''The WM_KEYDOWN message is posted to the window with the keyboard focus when a nonsystem key is pressed. A nonsystem key is a key that is pressed when the ALT key is not pressed.
        ''' </summary>
        ''' <remarks></remarks>
        WM_KEYDOWN = &H100

        ''' <summary>
        '''This message filters for keyboard messages.
        ''' </summary>
        ''' <remarks></remarks>
        WM_KEYFIRST = &H100

        ''' <summary>
        '''This message filters for keyboard messages.
        ''' </summary>
        ''' <remarks></remarks>
        WM_KEYLAST = &H108

        ''' <summary>
        '''The WM_KEYUP message is posted to the window with the keyboard focus when a nonsystem key is released. A nonsystem key is a key that is pressed when the ALT key is not pressed or a keyboard key that is pressed when a window has the keyboard focus.
        ''' </summary>
        ''' <remarks></remarks>
        WM_KEYUP = &H101

        ''' <summary>
        '''The WM_KILLFOCUS message is sent to a window immediately before it loses the keyboard focus.
        ''' </summary>
        ''' <remarks></remarks>
        WM_KILLFOCUS = &H8

        ''' <summary>
        '''The WM_LBUTTONDBLCLK message is posted when the user double-clicks the left mouse button while the cursor is in the client area of a window. If the mouse is not captured the message is posted to the window beneath the cursor. Otherwise the message is posted to the window that has captured the mouse.
        ''' </summary>
        ''' <remarks></remarks>
        WM_LBUTTONDBLCLK = &H203

        ''' <summary>
        '''The WM_LBUTTONDOWN message is posted when the user presses the left mouse button while the cursor is in the client area of a window. If the mouse is not captured the message is posted to the window beneath the cursor. Otherwise the message is posted to the window that has captured the mouse.
        ''' </summary>
        ''' <remarks></remarks>
        WM_LBUTTONDOWN = &H201

        ''' <summary>
        '''The WM_LBUTTONUP message is posted when the user releases the left mouse button while the cursor is in the client area of a window. If the mouse is not captured the message is posted to the window beneath the cursor. Otherwise the message is posted to the window that has captured the mouse.
        ''' </summary>
        ''' <remarks></remarks>
        WM_LBUTTONUP = &H202

        ''' <summary>
        '''The WM_MBUTTONDBLCLK message is posted when the user double-clicks the middle mouse button while the cursor is in the client area of a window. If the mouse is not captured the message is posted to the window beneath the cursor. Otherwise the message is posted to the window that has captured the mouse.
        ''' </summary>
        ''' <remarks></remarks>
        WM_MBUTTONDBLCLK = &H209

        ''' <summary>
        '''The WM_MBUTTONDOWN message is posted when the user presses the middle mouse button while the cursor is in the client area of a window. If the mouse is not captured the message is posted to the window beneath the cursor. Otherwise the message is posted to the window that has captured the mouse.
        ''' </summary>
        ''' <remarks></remarks>
        WM_MBUTTONDOWN = &H207

        ''' <summary>
        '''The WM_MBUTTONUP message is posted when the user releases the middle mouse button while the cursor is in the client area of a window. If the mouse is not captured the message is posted to the window beneath the cursor. Otherwise the message is posted to the window that has captured the mouse.
        ''' </summary>
        ''' <remarks></remarks>
        WM_MBUTTONUP = &H208

        ''' <summary>
        '''An application sends the WM_MDIACTIVATE message to a multiple-document interface (MDI) client window to instruct the client window to activate a different MDI child window.
        ''' </summary>
        ''' <remarks></remarks>
        WM_MDIACTIVATE = &H222

        ''' <summary>
        '''An application sends the WM_MDICASCADE message to a multiple-document interface (MDI) client window to arrange all its child windows in a cascade format.
        ''' </summary>
        ''' <remarks></remarks>
        WM_MDICASCADE = &H227

        ''' <summary>
        '''An application sends the WM_MDICREATE message to a multiple-document interface (MDI) client window to create an MDI child window.
        ''' </summary>
        ''' <remarks></remarks>
        WM_MDICREATE = &H220

        ''' <summary>
        '''An application sends the WM_MDIDESTROY message to a multiple-document interface (MDI) client window to close an MDI child window.
        ''' </summary>
        ''' <remarks></remarks>
        WM_MDIDESTROY = &H221

        ''' <summary>
        '''An application sends the WM_MDIGETACTIVE message to a multiple-document interface (MDI) client window to retrieve the handle to the active MDI child window.
        ''' </summary>
        ''' <remarks></remarks>
        WM_MDIGETACTIVE = &H229

        ''' <summary>
        '''An application sends the WM_MDIICONARRANGE message to a multiple-document interface (MDI) client window to arrange all minimized MDI child windows. It does not affect child windows that are not minimized.
        ''' </summary>
        ''' <remarks></remarks>
        WM_MDIICONARRANGE = &H228

        ''' <summary>
        '''An application sends the WM_MDIMAXIMIZE message to a multiple-document interface (MDI) client window to maximize an MDI child window. The system resizes the child window to make its client area fill the client window. The system places the child window's window menu icon in the rightmost position of the frame window's menu bar and places the child window's restore icon in the leftmost position. The system also appends the title bar text of the child window to that of the frame window.
        ''' </summary>
        ''' <remarks></remarks>
        WM_MDIMAXIMIZE = &H225

        ''' <summary>
        '''An application sends the WM_MDINEXT message to a multiple-document interface (MDI) client window to activate the next or previous child window.
        ''' </summary>
        ''' <remarks></remarks>
        WM_MDINEXT = &H224

        ''' <summary>
        '''An application sends the WM_MDIREFRESHMENU message to a multiple-document interface (MDI) client window to refresh the window menu of the MDI frame window.
        ''' </summary>
        ''' <remarks></remarks>
        WM_MDIREFRESHMENU = &H234

        ''' <summary>
        '''An application sends the WM_MDIRESTORE message to a multiple-document interface (MDI) client window to restore an MDI child window from maximized or minimized size.
        ''' </summary>
        ''' <remarks></remarks>
        WM_MDIRESTORE = &H223

        ''' <summary>
        '''An application sends the WM_MDISETMENU message to a multiple-document interface (MDI) client window to replace the entire menu of an MDI frame window to replace the window menu of the frame window or both.
        ''' </summary>
        ''' <remarks></remarks>
        WM_MDISETMENU = &H230

        ''' <summary>
        '''An application sends the WM_MDITILE message to a multiple-document interface (MDI) client window to arrange all of its MDI child windows in a tile format.
        ''' </summary>
        ''' <remarks></remarks>
        WM_MDITILE = &H226

        ''' <summary>
        '''The WM_MEASUREITEM message is sent to the owner window of a combo box list box list view control or menu item when the control or menu is created.
        ''' </summary>
        ''' <remarks></remarks>
        WM_MEASUREITEM = &H2C

        ''' <summary>
        '''The WM_MENUCHAR message is sent when a menu is active and the user presses a key that does not correspond to any mnemonic or accelerator key. This message is sent to the window that owns the menu.
        ''' </summary>
        ''' <remarks></remarks>
        WM_MENUCHAR = &H120

        ''' <summary>
        '''The WM_MENUCOMMAND message is sent when the user makes a selection from a menu.
        ''' </summary>
        ''' <remarks></remarks>
        WM_MENUCOMMAND = &H126

        ''' <summary>
        '''The WM_MENUDRAG message is sent to the owner of a drag-and-drop menu when the user drags a menu item.
        ''' </summary>
        ''' <remarks></remarks>
        WM_MENUDRAG = &H123

        ''' <summary>
        '''The WM_MENUGETOBJECT message is sent to the owner of a drag-and-drop menu when the mouse cursor enters a menu item or moves from the center of the item to the top or bottom of the item.
        ''' </summary>
        ''' <remarks></remarks>
        WM_MENUGETOBJECT = &H124

        ''' <summary>
        '''The WM_MENURBUTTONUP message is sent when the user releases the right mouse button while the cursor is on a menu item.
        ''' </summary>
        ''' <remarks></remarks>
        WM_MENURBUTTONUP = &H122

        ''' <summary>
        '''The WM_MENUSELECT message is sent to a menu's owner window when the user selects a menu item.
        ''' </summary>
        ''' <remarks></remarks>
        WM_MENUSELECT = &H11F

        ''' <summary>
        '''The WM_MOUSEACTIVATE message is sent when the cursor is in an inactive window and the user presses a mouse button. The parent window receives this message only if the child window passes it to the DefWindowProc function.
        ''' </summary>
        ''' <remarks></remarks>
        WM_MOUSEACTIVATE = &H21

        'Commented out b/c is the same as WM_MOUSEMOVE
        '''' <summary>
        ''''Use WM_MOUSEFIRST to specify the first mouse message. Use the PeekMessage() Function.
        '''' </summary>
        '''' <remarks></remarks>
        'WM_MOUSEFIRST = &H200

        ''' <summary>
        '''The WM_MOUSEHOVER message is posted to a window when the cursor hovers over the client area of the window for the period of time specified in a prior call to TrackMouseEvent.
        ''' </summary>
        ''' <remarks></remarks>
        WM_MOUSEHOVER = &H2A1

        ''' <summary>
        '''Definition Needed
        ''' </summary>
        ''' <remarks></remarks>
        WM_MOUSELAST = &H20D

        ''' <summary>
        '''The WM_MOUSELEAVE message is posted to a window when the cursor leaves the client area of the window specified in a prior call to TrackMouseEvent.
        ''' </summary>
        ''' <remarks></remarks>
        WM_MOUSELEAVE = &H2A3

        ''' <summary>
        '''The WM_MOUSEMOVE message is posted to a window when the cursor moves. If the mouse is not captured the message is posted to the window that contains the cursor. Otherwise the message is posted to the window that has captured the mouse.
        ''' </summary>
        ''' <remarks></remarks>
        WM_MOUSEMOVE = &H200

        ''' <summary>
        '''The WM_MOUSEWHEEL message is sent to the focus window when the mouse wheel is rotated. The DefWindowProc function propagates the message to the window's parent. There should be no internal forwarding of the message since DefWindowProc propagates it up the parent chain until it finds a window that processes it.
        ''' </summary>
        ''' <remarks></remarks>
        WM_MOUSEWHEEL = &H20A

        ''' <summary>
        '''The WM_MOUSEHWHEEL message is sent to the focus window when the mouse's horizontal scroll wheel is tilted or rotated. The DefWindowProc function propagates the message to the window's parent. There should be no internal forwarding of the message since DefWindowProc propagates it up the parent chain until it finds a window that processes it.
        ''' </summary>
        ''' <remarks></remarks>
        WM_MOUSEHWHEEL = &H20E

        ''' <summary>
        '''The WM_MOVE message is sent after a window has been moved.
        ''' </summary>
        ''' <remarks></remarks>
        WM_MOVE = &H3

        ''' <summary>
        '''The WM_MOVING message is sent to a window that the user is moving. By processing this message an application can monitor the position of the drag rectangle and if needed change its position.
        ''' </summary>
        ''' <remarks></remarks>
        WM_MOVING = &H216

        ''' <summary>
        '''Non Client Area Activated Caption(Title) of the Form
        ''' </summary>
        ''' <remarks></remarks>
        WM_NCACTIVATE = &H86

        ''' <summary>
        '''The WM_NCCALCSIZE message is sent when the size and position of a window's client area must be calculated. By processing this message an application can control the content of the window's client area when the size or position of the window changes.
        ''' </summary>
        ''' <remarks></remarks>
        WM_NCCALCSIZE = &H83

        ''' <summary>
        '''The WM_NCCREATE message is sent prior to the WM_CREATE message when a window is first created.
        ''' </summary>
        ''' <remarks></remarks>
        WM_NCCREATE = &H81

        ''' <summary>
        '''The WM_NCDESTROY message informs a window that its nonclient area is being destroyed. The DestroyWindow function sends the WM_NCDESTROY message to the window following the WM_DESTROY message. WM_DESTROY is used to free the allocated memory object associated with the window.
        ''' </summary>
        ''' <remarks></remarks>
        WM_NCDESTROY = &H82

        ''' <summary>
        '''The WM_NCHITTEST message is sent to a window when the cursor moves or when a mouse button is pressed or released. If the mouse is not captured the message is sent to the window beneath the cursor. Otherwise the message is sent to the window that has captured the mouse.
        ''' </summary>
        ''' <remarks></remarks>
        WM_NCHITTEST = &H84

        ''' <summary>
        '''The WM_NCLBUTTONDBLCLK message is posted when the user double-clicks the left mouse button while the cursor is within the nonclient area of a window. This message is posted to the window that contains the cursor. If a window has captured the mouse this message is not posted.
        ''' </summary>
        ''' <remarks></remarks>
        WM_NCLBUTTONDBLCLK = &HA3

        ''' <summary>
        '''The WM_NCLBUTTONDOWN message is posted when the user presses the left mouse button while the cursor is within the nonclient area of a window. This message is posted to the window that contains the cursor. If a window has captured the mouse this message is not posted.
        ''' </summary>
        ''' <remarks></remarks>
        WM_NCLBUTTONDOWN = &HA1

        ''' <summary>
        '''The WM_NCLBUTTONUP message is posted when the user releases the left mouse button while the cursor is within the nonclient area of a window. This message is posted to the window that contains the cursor. If a window has captured the mouse this message is not posted.
        ''' </summary>
        ''' <remarks></remarks>
        WM_NCLBUTTONUP = &HA2

        ''' <summary>
        '''The WM_NCMBUTTONDBLCLK message is posted when the user double-clicks the middle mouse button while the cursor is within the nonclient area of a window. This message is posted to the window that contains the cursor. If a window has captured the mouse this message is not posted.
        ''' </summary>
        ''' <remarks></remarks>
        WM_NCMBUTTONDBLCLK = &HA9

        ''' <summary>
        '''The WM_NCMBUTTONDOWN message is posted when the user presses the middle mouse button while the cursor is within the nonclient area of a window. This message is posted to the window that contains the cursor. If a window has captured the mouse this message is not posted.
        ''' </summary>
        ''' <remarks></remarks>
        WM_NCMBUTTONDOWN = &HA7

        ''' <summary>
        '''The WM_NCMBUTTONUP message is posted when the user releases the middle mouse button while the cursor is within the nonclient area of a window. This message is posted to the window that contains the cursor. If a window has captured the mouse this message is not posted.
        ''' </summary>
        ''' <remarks></remarks>
        WM_NCMBUTTONUP = &HA8

        ''' <summary>
        '''The WM_NCMOUSEMOVE message is posted to a window when the cursor is moved within the nonclient area of the window. This message is posted to the window that contains the cursor. If a window has captured the mouse this message is not posted.
        ''' </summary>
        ''' <remarks></remarks>
        WM_NCMOUSEMOVE = &HA0

        ''' <summary>
        '''The WM_NCPAINT message is sent to a window when its frame must be painted.
        ''' </summary>
        ''' <remarks></remarks>
        WM_NCPAINT = &H85

        ''' <summary>
        '''The WM_NCRBUTTONDBLCLK message is posted when the user double-clicks the right mouse button while the cursor is within the nonclient area of a window. This message is posted to the window that contains the cursor. If a window has captured the mouse this message is not posted.
        ''' </summary>
        ''' <remarks></remarks>
        WM_NCRBUTTONDBLCLK = &HA6

        ''' <summary>
        '''The WM_NCRBUTTONDOWN message is posted when the user presses the right mouse button while the cursor is within the nonclient area of a window. This message is posted to the window that contains the cursor. If a window has captured the mouse this message is not posted.
        ''' </summary>
        ''' <remarks></remarks>
        WM_NCRBUTTONDOWN = &HA4

        ''' <summary>
        '''The WM_NCRBUTTONUP message is posted when the user releases the right mouse button while the cursor is within the nonclient area of a window. This message is posted to the window that contains the cursor. If a window has captured the mouse this message is not posted.
        ''' </summary>
        ''' <remarks></remarks>
        WM_NCRBUTTONUP = &HA5

        ''' <summary>
        '''The WM_NEXTDLGCTL message is sent to a dialog box procedure to set the keyboard focus to a different control in the dialog box
        ''' </summary>
        ''' <remarks></remarks>
        WM_NEXTDLGCTL = &H28

        ''' <summary>
        '''The WM_NEXTMENU message is sent to an application when the right or left arrow key is used to switch between the menu bar and the system menu.
        ''' </summary>
        ''' <remarks></remarks>
        WM_NEXTMENU = &H213

        ''' <summary>
        '''Sent by a common control to its parent window when an event has occurred or the control requires some information.
        ''' </summary>
        ''' <remarks></remarks>
        WM_NOTIFY = &H4E

        ''' <summary>
        '''Determines if a window accepts ANSI or Unicode structures in the WM_NOTIFY notification message. WM_NOTIFYFORMAT messages are sent from a common control to its parent window and from the parent window to the common control.
        ''' </summary>
        ''' <remarks></remarks>
        WM_NOTIFYFORMAT = &H55

        ''' <summary>
        '''The WM_NULL message performs no operation. An application sends the WM_NULL message if it wants to post a message that the recipient window will ignore.
        ''' </summary>
        ''' <remarks></remarks>
        WM_NULL = &H0

        ''' <summary>
        '''Occurs when the control needs repainting
        ''' </summary>
        ''' <remarks></remarks>
        WM_PAINT = &HF

        ''' <summary>
        '''The WM_PAINTCLIPBOARD message is sent to the clipboard owner by a clipboard viewer window when the clipboard contains data in the CF_OWNERDISPLAY format and the clipboard viewer's client area needs repainting.
        ''' </summary>
        ''' <remarks></remarks>
        WM_PAINTCLIPBOARD = &H309

        ''' <summary>
        '''Windows NT 3.51 and earlier: The WM_PAINTICON message is sent to a minimized window when the icon is to be painted. This message is not sent by newer versions of Microsoft Windows except in unusual circumstances explained in the Remarks.
        ''' </summary>
        ''' <remarks></remarks>
        WM_PAINTICON = &H26

        ''' <summary>
        '''This message is sent by the OS to all top-level and overlapped windows after the window with the keyboard focus realizes its logical palette. This message enables windows that do not have the keyboard focus to realize their logical palettes and update their client areas.
        ''' </summary>
        ''' <remarks></remarks>
        WM_PALETTECHANGED = &H311

        ''' <summary>
        '''The WM_PALETTEISCHANGING message informs applications that an application is going to realize its logical palette.
        ''' </summary>
        ''' <remarks></remarks>
        WM_PALETTEISCHANGING = &H310

        ''' <summary>
        '''The WM_PARENTNOTIFY message is sent to the parent of a child window when the child window is created or destroyed or when the user clicks a mouse button while the cursor is over the child window. When the child window is being created the system sends WM_PARENTNOTIFY just before the CreateWindow or CreateWindowEx function that creates the window returns. When the child window is being destroyed the system sends the message before any processing to destroy the window takes place.
        ''' </summary>
        ''' <remarks></remarks>
        WM_PARENTNOTIFY = &H210

        ''' <summary>
        '''An application sends a WM_PASTE message to an edit control or combo box to copy the current content of the clipboard to the edit control at the current caret position. Data is inserted only if the clipboard contains data in CF_TEXT format.
        ''' </summary>
        ''' <remarks></remarks>
        WM_PASTE = &H302

        ''' <summary>
        '''Definition Needed
        ''' </summary>
        ''' <remarks></remarks>
        WM_PENWINFIRST = &H380

        ''' <summary>
        '''Definition Needed
        ''' </summary>
        ''' <remarks></remarks>
        WM_PENWINLAST = &H38F

        ''' <summary>
        '''Notifies applications that the system typically a battery-powered personal computer is about to enter a suspended mode. Obsolete : use POWERBROADCAST instead
        ''' </summary>
        ''' <remarks></remarks>
        WM_POWER = &H48

        ''' <summary>
        '''Notifies applications that a power-management event has occurred.
        ''' </summary>
        ''' <remarks></remarks>
        WM_POWERBROADCAST = &H218

        ''' <summary>
        '''The WM_PRINT message is sent to a window to request that it draw itself in the specified device context most commonly in a printer device context.
        ''' </summary>
        ''' <remarks></remarks>
        WM_PRINT = &H317

        ''' <summary>
        '''The WM_PRINTCLIENT message is sent to a window to request that it draw its client area in the specified device context most commonly in a printer device context.
        ''' </summary>
        ''' <remarks></remarks>
        WM_PRINTCLIENT = &H318

        ''' <summary>
        '''The WM_QUERYDRAGICON message is sent to a minimized (iconic) window. The window is about to be dragged by the user but does not have an icon defined for its class. An application can return a handle to an icon or cursor. The system displays this cursor or icon while the user drags the icon.
        ''' </summary>
        ''' <remarks></remarks>
        WM_QUERYDRAGICON = &H37

        ''' <summary>
        '''The WM_QUERYENDSESSION message is sent when the user chooses to end the session or when an application calls one of the system shutdown functions. If any application returns zero the session is not ended. The system stops sending WM_QUERYENDSESSION messages as soon as one application returns zero. After processing this message the system sends the WM_ENDSESSION message with the wParam parameter set to the results of the WM_QUERYENDSESSION message.
        ''' </summary>
        ''' <remarks></remarks>
        WM_QUERYENDSESSION = &H11

        ''' <summary>
        '''This message informs a window that it is about to receive the keyboard focus giving the window the opportunity to realize its logical palette when it receives the focus.
        ''' </summary>
        ''' <remarks></remarks>
        WM_QUERYNEWPALETTE = &H30F

        ''' <summary>
        '''The WM_QUERYOPEN message is sent to an icon when the user requests that the window be restored to its previous size and position.
        ''' </summary>
        ''' <remarks></remarks>
        WM_QUERYOPEN = &H13

        ''' <summary>
        '''The WM_QUEUESYNC message is sent by a computer-based training (CBT) application to separate user-input messages from other messages sent through the WH_JOURNALPLAYBACK Hook procedure.
        ''' </summary>
        ''' <remarks></remarks>
        WM_QUEUESYNC = &H23

        ''' <summary>
        '''Once received it ends the application's Message Loop signaling the application to end. It can be sent by pressing Alt+F4 Clicking the X in the upper right-hand of the program or going to File->Exit.
        ''' </summary>
        ''' <remarks></remarks>
        WM_QUIT = &H12

        ''' <summary>
        '''he WM_RBUTTONDBLCLK message is posted when the user double-clicks the right mouse button while the cursor is in the client area of a window. If the mouse is not captured the message is posted to the window beneath the cursor. Otherwise the message is posted to the window that has captured the mouse.
        ''' </summary>
        ''' <remarks></remarks>
        WM_RBUTTONDBLCLK = &H206

        ''' <summary>
        '''The WM_RBUTTONDOWN message is posted when the user presses the right mouse button while the cursor is in the client area of a window. If the mouse is not captured the message is posted to the window beneath the cursor. Otherwise the message is posted to the window that has captured the mouse.
        ''' </summary>
        ''' <remarks></remarks>
        WM_RBUTTONDOWN = &H204

        ''' <summary>
        '''The WM_RBUTTONUP message is posted when the user releases the right mouse button while the cursor is in the client area of a window. If the mouse is not captured the message is posted to the window beneath the cursor. Otherwise the message is posted to the window that has captured the mouse.
        ''' </summary>
        ''' <remarks></remarks>
        WM_RBUTTONUP = &H205

        ''' <summary>
        '''The WM_RENDERALLFORMATS message is sent to the clipboard owner before it is destroyed if the clipboard owner has delayed rendering one or more clipboard formats. For the content of the clipboard to remain available to other applications the clipboard owner must render data in all the formats it is capable of generating and place the data on the clipboard by calling the SetClipboardData function.
        ''' </summary>
        ''' <remarks></remarks>
        WM_RENDERALLFORMATS = &H306

        ''' <summary>
        '''The WM_RENDERFORMAT message is sent to the clipboard owner if it has delayed rendering a specific clipboard format and if an application has requested data in that format. The clipboard owner must render data in the specified format and place it on the clipboard by calling the SetClipboardData function.
        ''' </summary>
        ''' <remarks></remarks>
        WM_RENDERFORMAT = &H305

        ''' <summary>
        '''The WM_SETCURSOR message is sent to a window if the mouse causes the cursor to move within a window and mouse input is not captured.
        ''' </summary>
        ''' <remarks></remarks>
        WM_SETCURSOR = &H20

        ''' <summary>
        '''When the controll got the focus
        ''' </summary>
        ''' <remarks></remarks>
        WM_SETFOCUS = &H7

        ''' <summary>
        '''An application sends a WM_SETFONT message to specify the font that a control is to use when drawing text.
        ''' </summary>
        ''' <remarks></remarks>
        WM_SETFONT = &H30

        ''' <summary>
        '''An application sends a WM_SETHOTKEY message to a window to associate a hot key with the window. When the user presses the hot key the system activates the window.
        ''' </summary>
        ''' <remarks></remarks>
        WM_SETHOTKEY = &H32

        ''' <summary>
        '''An application sends the WM_SETICON message to associate a new large or small icon with a window. The system displays the large icon in the ALT+TAB dialog box and the small icon in the window caption.
        ''' </summary>
        ''' <remarks></remarks>
        WM_SETICON = &H80

        ''' <summary>
        '''An application sends the WM_SETREDRAW message to a window to allow changes in that window to be redrawn or to prevent changes in that window from being redrawn.
        ''' </summary>
        ''' <remarks></remarks>
        WM_SETREDRAW = &HB

        ''' <summary>
        '''Text / Caption changed on the control. An application sends a WM_SETTEXT message to set the text of a window.
        ''' </summary>
        ''' <remarks></remarks>
        WM_SETTEXT = &HC

        ''' <summary>
        '''An application sends the WM_SETTINGCHANGE message to all top-level windows after making a change to the WIN.INI file. The SystemParametersInfo function sends this message after an application uses the function to change a setting in WIN.INI.
        ''' </summary>
        ''' <remarks></remarks>
        WM_SETTINGCHANGE = &H1A

        ''' <summary>
        '''The WM_SHOWWINDOW message is sent to a window when the window is about to be hidden or shown
        ''' </summary>
        ''' <remarks></remarks>
        WM_SHOWWINDOW = &H18

        ''' <summary>
        '''The WM_SIZE message is sent to a window after its size has changed.
        ''' </summary>
        ''' <remarks></remarks>
        WM_SIZE = &H5

        ''' <summary>
        '''The WM_SIZECLIPBOARD message is sent to the clipboard owner by a clipboard viewer window when the clipboard contains data in the CF_OWNERDISPLAY format and the clipboard viewer's client area has changed size.
        ''' </summary>
        ''' <remarks></remarks>
        WM_SIZECLIPBOARD = &H30B

        ''' <summary>
        '''The WM_SIZING message is sent to a window that the user is resizing. By processing this message an application can monitor the size and position of the drag rectangle and if needed change its size or position.
        ''' </summary>
        ''' <remarks></remarks>
        WM_SIZING = &H214

        ''' <summary>
        '''The WM_SPOOLERSTATUS message is sent from Print Manager whenever a job is added to or removed from the Print Manager queue.
        ''' </summary>
        ''' <remarks></remarks>
        WM_SPOOLERSTATUS = &H2A

        ''' <summary>
        '''The WM_STYLECHANGED message is sent to a window after the SetWindowLong function has changed one or more of the window's styles.
        ''' </summary>
        ''' <remarks></remarks>
        WM_STYLECHANGED = &H7D

        ''' <summary>
        '''The WM_STYLECHANGING message is sent to a window when the SetWindowLong function is about to change one or more of the window's styles.
        ''' </summary>
        ''' <remarks></remarks>
        WM_STYLECHANGING = &H7C

        ''' <summary>
        '''The WM_SYNCPAINT message is used to synchronize painting while avoiding linking independent GUI threads.
        ''' </summary>
        ''' <remarks></remarks>
        WM_SYNCPAINT = &H88

        ''' <summary>
        '''The WM_SYSCHAR message is posted to the window with the keyboard focus when a WM_SYSKEYDOWN message is translated by the TranslateMessage function. It specifies the character code of a system character key — that is a character key that is pressed while the ALT key is down.
        ''' </summary>
        ''' <remarks></remarks>
        WM_SYSCHAR = &H106

        ''' <summary>
        '''This message is sent to all top-level windows when a change is made to a system color setting.
        ''' </summary>
        ''' <remarks></remarks>
        WM_SYSCOLORCHANGE = &H15

        ''' <summary>
        '''A window receives this message when the user chooses a command from the Window menu (formerly known as the system or control menu) or when the user chooses the maximize button minimize button restore button or close button.
        ''' </summary>
        ''' <remarks></remarks>
        WM_SYSCOMMAND = &H112

        ''' <summary>
        '''The WM_SYSDEADCHAR message is sent to the window with the keyboard focus when a WM_SYSKEYDOWN message is translated by the TranslateMessage function. WM_SYSDEADCHAR specifies the character code of a system dead key — that is a dead key that is pressed while holding down the ALT key.
        ''' </summary>
        ''' <remarks></remarks>
        WM_SYSDEADCHAR = &H107

        ''' <summary>
        '''The WM_SYSKEYDOWN message is posted to the window with the keyboard focus when the user presses the F10 key (which activates the menu bar) or holds down the ALT key and then presses another key. It also occurs when no window currently has the keyboard focus; in this case the WM_SYSKEYDOWN message is sent to the active window. The window that receives the message can distinguish between these two contexts by checking the context code in the lParam parameter.
        ''' </summary>
        ''' <remarks></remarks>
        WM_SYSKEYDOWN = &H104

        ''' <summary>
        '''The WM_SYSKEYUP message is posted to the window with the keyboard focus when the user releases a key that was pressed while the ALT key was held down. It also occurs when no window currently has the keyboard focus; in this case the WM_SYSKEYUP message is sent to the active window. The window that receives the message can distinguish between these two contexts by checking the context code in the lParam parameter.
        ''' </summary>
        ''' <remarks></remarks>
        WM_SYSKEYUP = &H105

        ''' <summary>
        '''Sent to an application that has initiated a training card with Microsoft Windows Help. The message informs the application when the user clicks an authorable button. An application initiates a training card by specifying the HELP_TCARD command in a call to the WinHelp function.
        ''' </summary>
        ''' <remarks></remarks>
        WM_TCARD = &H52

        ''' <summary>
        '''A message that is sent whenever there is a change in the system time.
        ''' </summary>
        ''' <remarks></remarks>
        WM_TIMECHANGE = &H1E

        ''' <summary>
        '''The WM_TIMER message is posted to the installing thread's message queue when a timer expires. The message is posted by the GetMessage or PeekMessage function.
        ''' </summary>
        ''' <remarks></remarks>
        WM_TIMER = &H113

        ''' <summary>
        '''An application sends a WM_UNDO message to an edit control to undo the last operation. When this message is sent to an edit control the previously deleted text is restored or the previously added text is deleted.
        ''' </summary>
        ''' <remarks></remarks>
        WM_UNDO = &H304

        ''' <summary>
        '''The WM_UNINITMENUPOPUP message is sent when a drop-down menu or submenu has been destroyed.
        ''' </summary>
        ''' <remarks></remarks>
        WM_UNINITMENUPOPUP = &H125

        ''' <summary>
        '''The WM_USER constant is used by applications to help define private messages for use by private window classes usually of the form WM_USER+X where X is an integer value.
        ''' </summary>
        ''' <remarks></remarks>
        WM_USER = &H400

        ''' <summary>
        '''The WM_USERCHANGED message is sent to all windows after the user has logged on or off. When the user logs on or off the system updates the user-specific settings. The system sends this message immediately after updating the settings.
        ''' </summary>
        ''' <remarks></remarks>
        WM_USERCHANGED = &H54

        ''' <summary>
        '''Sent by a list box with the LBS_WANTKEYBOARDINPUT style to its owner in response to a WM_KEYDOWN message.
        ''' </summary>
        ''' <remarks></remarks>
        WM_VKEYTOITEM = &H2E

        ''' <summary>
        '''The WM_VSCROLL message is sent to a window when a scroll event occurs in the window's standard vertical scroll bar. This message is also sent to the owner of a vertical scroll bar control when a scroll event occurs in the control.
        ''' </summary>
        ''' <remarks></remarks>
        WM_VSCROLL = &H115

        ''' <summary>
        '''The WM_VSCROLLCLIPBOARD message is sent to the clipboard owner by a clipboard viewer window when the clipboard contains data in the CF_OWNERDISPLAY format and an event occurs in the clipboard viewer's vertical scroll bar. The owner should scroll the clipboard image and update the scroll bar values.
        ''' </summary>
        ''' <remarks></remarks>
        WM_VSCROLLCLIPBOARD = &H30A

        ''' <summary>
        '''The WM_WINDOWPOSCHANGED message is sent to a window whose size position or place in the Z order has changed as a result of a call to the SetWindowPos function or another window-management function.
        ''' </summary>
        ''' <remarks></remarks>
        WM_WINDOWPOSCHANGED = &H47

        ''' <summary>
        '''The WM_WINDOWPOSCHANGING message is sent to a window whose size position or place in the Z order is about to change as a result of a call to the SetWindowPos function or another window-management function.
        ''' </summary>
        ''' <remarks></remarks>
        WM_WINDOWPOSCHANGING = &H46

        ''' <summary>
        '''An application sends the WM_WININICHANGE message to all top-level windows after making a change to the WIN.INI file. The SystemParametersInfo function sends this message after an application uses the function to change a setting in WIN.INI. Note The WM_WININICHANGE message is provided only for compatibility with earlier versions of the system. Applications should use the WM_SETTINGCHANGE message.
        ''' </summary>
        ''' <remarks></remarks>
        WM_WININICHANGE = &H1A

        ''' <summary>
        '''Definition Needed
        ''' </summary>
        ''' <remarks></remarks>
        WM_XBUTTONDBLCLK = &H20D

        ''' <summary>
        '''Definition Needed
        ''' </summary>
        ''' <remarks></remarks>
        WM_XBUTTONDOWN = &H20B

        ''' <summary>
        '''Definition Needed
        ''' </summary>
        ''' <remarks></remarks>
        WM_XBUTTONUP = &H20C



    End Enum

    ''' <summary>
    ''' Hooks the class when true and Unhooks the class when false. This value should never be set to true consecutively and NEVER
    ''' release this object without calling setting this value to false.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property IsHooked() As Boolean
        Get
            Return HookHandle <> Nothing
        End Get
        Set(ByVal value As Boolean)
            If value Then
                Hook()
            Else
                UnHook()
            End If
        End Set
    End Property

    ''' <summary>
    ''' Called when Hook is called.  Run any extra unhooking code you wish to perform here.
    ''' </summary>
    ''' <remarks></remarks>
    Protected MustOverride Sub OnHook()

    ''' <summary>
    ''' Called when UnHook is called.  Run any extra unhooking code you wish to perform here.
    ''' </summary>
    ''' <remarks></remarks>
    Protected MustOverride Sub OnUnHook()

    ''' <summary>
    ''' A value used to determine whether or not other applications or threads that have also installed this hook type will receive this hook event.
    ''' This is based on the last hook to be installed to the chain, if another application installs a hook of the same type and also does not perform the CallNextHookEx then this application or thread
    ''' will not receive the hook event either. By default and strong recommendation you should ALWAYS call the next hook meaning this value should always be true.
    ''' --This is not the same as the Disabled property which prevents all hook events of this type from being sent to ANY application.
    ''' </summary>
    ''' <value>True to continue processing the hook event to other hooks previously installed, false to dis-continue other hooks installed.</value>
    ''' <returns>Value indicating whether or not you are currently returning previously installed hooks.</returns>
    ''' <remarks>This value may not work for all hooks, check msdn for more information.</remarks>
    Public Property CallNextHook() As Boolean = True

    ''' <summary>
    ''' The Delegate passed in to the SetWindowsHookEx as the callback variable.
    ''' [in] Pointer to the hook procedure. If the dwThreadId parameter is zero or specifies the identifier of a thread created by a different process, the lpfn parameter must point to a hook procedure in a DLL.
    ''' Otherwise, lpfn can point to a hook procedure in the code associated with the current process. 
    ''' </summary>
    ''' <param name="nCode">[in] Specifies a code the hook procedure uses to determine how to process the message. If nCode is less than zero, the hook procedure must pass the message to the CallNextHookEx function without further processing and should return the value returned by CallNextHookEx. This parameter can be one of the following values. 
    ''' HC_ACTION
    ''' The wParam and lParam parameters contain information about a keyboard message.</param>
    ''' <param name="wParam">[in] Specifies the identifier of the keyboard message. This parameter can be one of the following messages: WM_KEYDOWN, WM_KEYUP, WM_SYSKEYDOWN, or WM_SYSKEYUP. </param>
    ''' <param name="lParam">[in] Pointer to a KBDLLHOOKSTRUCT structure. </param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function MainHookProc(ByVal nCode As Integer, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As IntPtr
        If nCode < 0 Then Return CallNextHookEx(HookHandle, nCode, wParam, lParam)
        Dim HBEA As New HookBaseEventArgs(nCode, wParam, lParam)
        RaiseEvent HookProcessing(Me, HBEA)
        If OnHookProc(nCode, wParam, lParam) Or Disabled Or HBEA.Handled Then
            If CallNextHook Then CallNextHookEx(Nothing, nCode, wParam, lParam)
            Return New IntPtr(1)
        Else
            If CallNextHook Then
                Return CallNextHookEx(HookHandle, nCode, wParam, lParam)
            Else
                Return IntPtr.Zero
            End If
        End If
    End Function

    ''' <summary>
    ''' A generic event that runs when the hook is being processed.
    ''' </summary>
    ''' <param name="sender">HookBase</param>
    ''' <param name="e">HookBaseEventArgs</param>
    ''' <remarks></remarks>
    Public Event HookProcessing(ByVal sender As Object, ByVal e As HookBaseEventArgs)

    ''' <summary>
    ''' The Delegate passed in to the SetWindowsHookEx as the callback variable.
    ''' [in] Pointer to the hook procedure. If the dwThreadId parameter is zero or specifies the identifier of a thread created by a different process, the lpfn parameter must point to a hook procedure in a DLL.
    ''' Otherwise, lpfn can point to a hook procedure in the code associated with the current process. 
    ''' </summary>
    ''' <param name="nCode">[in] Specifies a code the hook procedure uses to determine how to process the message. If nCode is less than zero, the hook procedure must pass the message to the CallNextHookEx function without further processing and should return the value returned by CallNextHookEx. This parameter can be one of the following values. 
    ''' HC_ACTION
    ''' The wParam and lParam parameters contain information about a keyboard message.</param>
    ''' <param name="wParam">[in] Specifies the identifier of the keyboard message. This parameter can be one of the following messages: WM_KEYDOWN, WM_KEYUP, WM_SYSKEYDOWN, or WM_SYSKEYUP. </param>
    ''' <param name="lParam">[in] Pointer to a KBDLLHOOKSTRUCT structure. </param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Protected MustOverride Function OnHookProc(ByVal nCode As Integer, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As Boolean

    ''' <summary>
    ''' If true prevents the hook messages from reaching other applications. This will leave the entire keyboard, mouse, ect.. disabled.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Disabled() As Boolean

    Public Enum HookType As Integer
        WH_JOURNALRECORD = 0
        WH_JOURNALPLAYBACK = 1
        WH_KEYBOARD = 2
        WH_GETMESSAGE = 3
        WH_CALLWNDPROC = 4
        WH_CBT = 5
        WH_SYSMSGFILTER = 6
        WH_MOUSE = 7
        WH_HARDWARE = 8
        WH_DEBUG = 9
        WH_SHELL = 10
        WH_FOREGROUNDIDLE = 11
        WH_CALLWNDPROCRET = 12
        ''' <summary>
        ''' Windows NT/2000/XP: Represents a hook procedure that monitors low-level keyboard input events. For more information, see the LowLevelKeyboardProc hook procedure.
        ''' Used as the hook value in SetWindowsHookEx. 
        ''' </summary>
        ''' <remarks></remarks>
        WH_KEYBOARD_LL = 13
        ''' <summary>
        ''' Windows NT/2000/XP: Installs a hook procedure that monitors low-level mouse input events. For more information, see the LowLevelMouseProc hook procedure.
        ''' </summary>
        ''' <remarks></remarks>
        WH_MOUSE_LL = 14
    End Enum

    ''' <summary>
    ''' Hooks the Class.  If the class is already hooked once it is unhooked and an exception is thrown.
    ''' It also resets the KeyPositions and gains the HookHandle with a call to SetWindowsHookEx.
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub Hook()

        If IsHooked Then
            UnHook()
            Throw New Exception("Cannot hook the class twice with the same handle. Class is already hooked by this process or thread and has now been unhooked do prevent system error.") : Return
        End If
        HookHandle = SetWindowsHookEx(idHook, HookProc_Delegated, Handle, Nothing)
        If HookHandle = Nothing Then Throw New Exception("HookKeyBoard Failed. A call to SetWindowsHookEx returned a value of Null")
        OnHook()
    End Sub

    ''' <summary>
    ''' Unhooks the class with a call to UnHookWindowsHookEx.
    ''' Sets the hookHandle as nothing (meaning IsHooked returns false) and all other values should be reset.
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub UnHook()
        UnhookWindowsHookEx(HookHandle)
        HookHandle = Nothing
    End Sub

#Region " IDisposable Support "
    Private disposedValue As Boolean = False
    Protected Overridable Sub Dispose(ByVal disposing As Boolean)
        If Not disposedValue Then
            If disposing Then
                If IsHooked Then UnHook()
                HookProc_Delegated = Nothing
                Handle = Nothing
            End If
        End If
        disposedValue = True
    End Sub


    Public Sub Dispose() Implements IDisposable.Dispose
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
#End Region

    Protected Overrides Sub Finalize()
        Dispose()
        MyBase.Finalize()
    End Sub

End Class