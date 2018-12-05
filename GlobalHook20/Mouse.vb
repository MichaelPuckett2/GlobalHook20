Imports System.Runtime.InteropServices
Imports System.Drawing

''' <summary>
''' Used to perform a low level hook on the mouse then read and fire mouse event information.
''' </summary>
''' <remarks></remarks>
Public Class Mouse
    Inherits HookBase
    Implements IDisposable

    Public Sub New()
        MyBase.New(HookType.WH_MOUSE_LL)
    End Sub

    ''' <summary>
    ''' Runs when Hook is called.
    ''' </summary>
    ''' <remarks></remarks>
    Protected Overrides Sub OnHook()
        StoredButtonPositions = New Dictionary(Of MouseButtons, Boolean)
    End Sub

    ''' <summary>
    ''' Runs when UnHook is called.
    ''' </summary>
    ''' <remarks></remarks>
    Protected Overrides Sub OnUnHook()
        StoredButtonPositions = Nothing
    End Sub

    ''' <summary>
    ''' The address of the delegate passed to SetWindowsHookEx.  Pointed to by LowLevelMouseProc_Delegated, instatiation of LowLevelMouseProc.
    ''' </summary>
    ''' <param name="nCode">[in] Specifies a code the hook procedure uses to determine how to process the message.
    ''' If nCode is less than zero, the hook procedure must pass the message to the CallNextHookEx function without further processing and should return the value returned by CallNextHookEx.
    ''' This parameter can be one of the following values. 
    ''' HC_ACTION
    ''' The wParam and lParam parameters contain information about a mouse message</param>
    ''' <param name="wParam">[in] Specifies the identifier of the mouse message. This parameter can be one of the following messages:
    ''' WM_LBUTTONDOWN, WM_LBUTTONUP, WM_MOUSEMOVE, WM_MOUSEWHEEL, WM_MOUSEHWHEEL, WM_RBUTTONDOWN, or WM_RBUTTONUP.</param>
    ''' <param name="lParam">[in] Pointer to an MSLLHOOKSTRUCT structure.</param>
    ''' <returns>True if that event should be handled and not passed to other applications.</returns>
    ''' <remarks>An application installs the hook procedure by specifying the WH_MOUSE_LL hook type and a pointer to the hook procedure in a call to the SetWindowsHookEx function.
    ''' This hook is called in the context of the thread that installed it. The call is made by sending a message to the thread that installed the hook. Therefore, the thread that installed the hook must have a message loop.
    ''' The hook procedure should process a message in less time than the data entry specified in the LowLevelHooksTimeout value in the following registry key: 
    ''' HKEY_CURRENT_USER\Control Panel\Desktop
    ''' The value is in milliseconds. If the hook procedure does not return during this interval, the system will pass the message to the next hook.
    ''' Note that debug hooks cannot track this type of hook.</remarks>
    Protected Overrides Function OnHookProc(ByVal nCode As Integer, ByVal wParam As IntPtr, <[In]()> ByVal lParam As IntPtr) As Boolean
        'Mouse Event Code Here
        Dim MEA As New MouseEventArgs(nCode, wParam, lParam)

        Select Case wParam.ToInt32

            Case WindowsMessages.WM_MOUSEMOVE
                RaiseEvent Mouse_Move(Me, MEA)

            Case WindowsMessages.WM_MBUTTONDOWN
                StoredButtonPositions(MouseButtons.Middle) = True
                RaiseEvent Mouse_MiddleButtonDown(Me, MEA)

            Case WindowsMessages.WM_MBUTTONUP
                StoredButtonPositions(MouseButtons.Middle) = False
                RaiseEvent Mouse_MiddleButtonUp(Me, MEA)

            Case WindowsMessages.WM_RBUTTONDOWN
                StoredButtonPositions(MouseButtons.Right) = True
                RaiseEvent Mouse_RightButtonDown(Me, MEA)

            Case WindowsMessages.WM_RBUTTONUP
                StoredButtonPositions(MouseButtons.Right) = False
                RaiseEvent Mouse_RightButtonUp(Me, MEA)

            Case WindowsMessages.WM_LBUTTONDOWN
                StoredButtonPositions(MouseButtons.Left) = True
                RaiseEvent Mouse_LeftButtonDown(Me, MEA)

            Case WindowsMessages.WM_LBUTTONUP
                StoredButtonPositions(MouseButtons.Left) = False
                RaiseEvent Mouse_LeftButtonUp(Me, MEA)

            Case WindowsMessages.WM_MOUSEWHEEL
                RaiseEvent Mouse_Wheel(Me, MEA)

            Case WindowsMessages.WM_XBUTTONDOWN
                StoredButtonPositions(MouseButtons.XButton1) = True
                RaiseEvent Mouse_XButtonDown(Me, MEA)

            Case WindowsMessages.WM_XBUTTONUP
                StoredButtonPositions(MouseButtons.XButton1) = False
                RaiseEvent Mouse_XButtonUp(Me, MEA)

            Case WindowsMessages.WM_MOUSEHWHEEL
                RaiseEvent Mouse_MiddleHorizontalButton(Me, MEA)
        End Select

        If wParam.ToInt32 <> WindowsMessages.WM_MOUSEMOVE Then RaiseEvent Mouse_Event(Me, MEA)

        Return MEA.Handled
    End Function

#Region "Mouse Events"

    ''' <summary>
    ''' Fired when the mouse button is used as a generic event. Does not fire for mouse movement.  Use the Mouse_Move event directly to intercept the mouse moving.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Public Event Mouse_Event(ByVal sender As Object, ByVal e As MouseEventArgs)

    ''' <summary>
    ''' Fired when the middle mouse button is pushed down.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Public Event Mouse_MiddleButtonDown(ByVal sender As Object, ByVal e As MouseEventArgs)

    ''' <summary>
    ''' Fired when the middle mouse button is released up.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Public Event Mouse_MiddleButtonUp(ByVal sender As Object, ByVal e As MouseEventArgs)

    ''' <summary>
    ''' Fired when the right mouse button is pressed down.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Public Event Mouse_RightButtonDown(ByVal sender As Object, ByVal e As MouseEventArgs)

    ''' <summary>
    ''' Fired when the right mouse button is released up.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Public Event Mouse_RightButtonUp(ByVal sender As Object, ByVal e As MouseEventArgs)

    ''' <summary>
    ''' Fired when the left mouse button is pressed down.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Public Event Mouse_LeftButtonDown(ByVal sender As Object, ByVal e As MouseEventArgs)

    ''' <summary>
    ''' Fired when the left mouse button is released up.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Public Event Mouse_LeftButtonUp(ByVal sender As Object, ByVal e As MouseEventArgs)

    ''' <summary>
    ''' Fired when the mouse wheel or delta is moved forward or backwards.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Public Event Mouse_Wheel(ByVal sender As Object, ByVal e As MouseEventArgs)

    ''' <summary>
    ''' Fired when the mouse X button is pressed down.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Public Event Mouse_XButtonDown(ByVal sender As Object, ByVal e As MouseEventArgs)

    ''' <summary>
    ''' Fired when the mouse X button is released up.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Public Event Mouse_XButtonUp(ByVal sender As Object, ByVal e As MouseEventArgs)

    ''' <summary>
    ''' Fired when the mouse is moved. Careful using this event b/c latency is an issue here. The longer you tie up this event the longer the mouse cursor waits to be updated on screen.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Public Event Mouse_Move(ByVal sender As Object, ByVal e As MouseEventArgs)

    ''' <summary>
    ''' Fired when the middle horizontal wheel is pressed left, right, or rotated.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Public Event Mouse_MiddleHorizontalButton(ByVal sender As Object, ByVal e As MouseEventArgs)

#End Region

    ''' <summary>
    ''' A generic dictionary of Forms.Keys that stores a boolean for each key pressed = true is the key is down and false when it is released.
    ''' </summary>
    ''' <remarks></remarks>
    Private StoredButtonPositions As Dictionary(Of MouseButtons, Boolean)

    ''' <summary>
    ''' A function that takes a Forms.Keys Argument, tests it and returns True if the key is currently down and false if it is up.
    ''' This method only works when the keyboard is hooked and will not return true for keys that are down before the keyboard was hooked.
    ''' </summary>
    ''' <param name="button"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function IsButtonDown(ByVal button As MouseButtons) As Boolean
        If StoredButtonPositions Is Nothing Then Throw New Exception("The Keyboard must be successfully hooked in order to retrieve the key position")
        If Not StoredButtonPositions.ContainsKey(button) Then Return False
        Return StoredButtonPositions(button)
    End Function

    ''' <summary>
    ''' Takes an Array of Forms.MouseButtons as an argument and test them all and returns True if they are all down and false if any of them are not down.
    ''' This method only works when the mouse is hooked and will not return true for buttons that are down before the mouse was hooked.
    ''' </summary>
    ''' <param name="buttons"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function AreButtonsDown(ByVal ParamArray buttons() As MouseButtons) As Boolean
        If StoredButtonPositions Is Nothing Then Throw New Exception("The mouse must be successfully hooked in order to retrieve the key position")
        Dim IsDown As Boolean
        For Each k As MouseButtons In buttons
            If Not StoredButtonPositions.ContainsKey(k) Then Return False
            IsDown = StoredButtonPositions(k)
            If Not IsDown Then Return False
        Next
        Return IsDown
    End Function

    ''' <summary>
    ''' The MSLLHOOKSTRUCT structure contains information about a low-level mouse input event. 
    ''' </summary>
    ''' <remarks></remarks>
    <StructLayout(LayoutKind.Sequential)> _
        Public Structure MSLLHOOKSTRUCT
        ''' <summary>
        ''' Specifies a POINT structure that contains the x- and y-coordinates of the cursor, in screen coordinates. 
        ''' </summary>
        ''' <remarks></remarks>
        Public pt As Point
        ''' <summary>
        ''' If the message is WM_MOUSEWHEEL, the high-order word of this member is the wheel delta. The low-order word is reserved.
        ''' A positive value indicates that the wheel was rotated forward, away from the user; a negative value indicates that the wheel was rotated backward, toward the user.
        ''' One wheel click is defined as WHEEL_DELTA, which is 120.
        ''' If the message is WM_XBUTTONDOWN, WM_XBUTTONUP, WM_XBUTTONDBLCLK, WM_NCXBUTTONDOWN, WM_NCXBUTTONUP, or WM_NCXBUTTONDBLCLK,
        ''' the high-order word specifies which X button was pressed or released, and the low-order word is reserved. This value can be one or more of the following values. Otherwise, mouseData is not used. 
        ''' XBUTTON1
        ''' The first X button was pressed or released.
        ''' XBUTTON2
        ''' The second X button was pressed or released.
        ''' </summary>
        ''' <remarks></remarks>
        Public mouseData As Integer
        ''' <summary>
        ''' Specifies the event-injected flag. An application can use the following value to test the mouse flags. Value Purpose 
        ''' LLMHF_INJECTED Test the event-injected flag.
        ''' 0
        ''' Specifies whether the event was injected. The value is 1 if the event was injected; otherwise, it is 0.
        ''' 1-15
        ''' Reserved.
        ''' </summary>
        ''' <remarks></remarks>
        Public flags As MSLLHOOKSTRUCTFlags
        ''' <summary>
        ''' Specifies the time stamp for this message. 
        ''' </summary>
        ''' <remarks></remarks>
        Public time As UInteger
        ''' <summary>
        ''' Specifies extra information associated with the message. 
        ''' </summary>
        ''' <remarks></remarks>
        Public dwExtraInfo As IntPtr
    End Structure

    ''' <summary>
    ''' Specifies the event-injected flag. An application can use the following value to test the mouse flags. Value Purpose LLMHF_INJECTED Test the event-injected flag.  
    ''' 0
    ''' Specifies whether the event was injected. The value is 1 if the event was injected; otherwise, it is 0.
    ''' 1-15
    ''' Reserved.
    ''' </summary>
    ''' <remarks></remarks>
    <Flags()> _
    Public Enum MSLLHOOKSTRUCTFlags As UInteger
        ''' <summary>
        ''' Test the event-injected flag.
        ''' </summary>
        ''' <remarks></remarks>
        LLMHF_INJECTED = &H1
    End Enum

End Class
