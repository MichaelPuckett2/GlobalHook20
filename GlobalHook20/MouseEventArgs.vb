Imports System.Runtime.InteropServices

''' <summary>
''' Passed with MouseEvents from the Mouse class with all the information needed to read and handle the current mouse event.
''' </summary>
''' <remarks></remarks>
Public Class MouseEventArgs
    Inherits HookBaseEventArgs

    Private lParamCopied As Mouse.MSLLHOOKSTRUCT

    ''' <summary>
    ''' The constructor for the MouseEventArgs class which initializes the properties and determines certain mouse events.
    ''' </summary>
    ''' <param name="nCode">[in] Specifies a code the hook procedure uses to determine how to process the message.
    ''' If nCode is less than zero, the hook procedure must pass the message to the CallNextHookEx function without further processing and should return the value returned by CallNextHookEx.
    ''' This parameter can be one of the following values. 
    ''' HC_ACTION
    ''' The wParam and lParam parameters contain information about a mouse message</param>
    ''' <param name="wParam">[in] Specifies the identifier of the mouse message. This parameter can be one of the following messages:
    ''' WM_LBUTTONDOWN, WM_LBUTTONUP, WM_MOUSEMOVE, WM_MOUSEWHEEL, WM_MOUSEHWHEEL, WM_RBUTTONDOWN, or WM_RBUTTONUP.</param>
    ''' <param name="lParam">[in] Pointer to an MSLLHOOKSTRUCT structure.</param>
    ''' <remarks>An application installs the hook procedure by specifying the WH_MOUSE_LL hook type and a pointer to the hook procedure in a call to the SetWindowsHookEx function.
    ''' This hook is called in the context of the thread that installed it. The call is made by sending a message to the thread that installed the hook. Therefore, the thread that installed the hook must have a message loop.
    ''' The hook procedure should process a message in less time than the data entry specified in the LowLevelHooksTimeout value in the following registry key: 
    ''' HKEY_CURRENT_USER\Control Panel\Desktop
    ''' The value is in milliseconds. If the hook procedure does not return during this interval, the system will pass the message to the next hook.
    ''' Note that debug hooks cannot track this type of hook.</remarks>
    Public Sub New(ByVal nCode As Integer, ByVal wParam As IntPtr, <[In]()> ByVal lParam As IntPtr)
        MyBase.New(nCode, wParam, lParam)
        lParamCopied = CType(Marshal.PtrToStructure(lParam, GetType(Mouse.MSLLHOOKSTRUCT)), Mouse.MSLLHOOKSTRUCT)
    End Sub

    ''' <summary>
    ''' The location on screen the mouse is currently positioned.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Point() As Drawing.Point
        Get
            Return lParamCopied.pt
        End Get
    End Property

    ''' <summary>
    ''' True if the current mouse event was injected; meaning another application inserted the event and the mouse was not physically used.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Injected() As Boolean
        Get
            Return lParamCopied.flags = (lParamCopied.flags Or Mouse.MSLLHOOKSTRUCTFlags.LLMHF_INJECTED)
        End Get
    End Property

    ''' <summary>
    ''' Returns the direction the wheel was turned if used.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property WheelDirection() As WheelDirections
        Get
            If (WParam.ToInt32 <> HookBase.WindowsMessages.WM_MOUSEWHEEL) Then Return WheelDirections.none
            If (lParamCopied.mouseData > 0) Then Return WheelDirections.WheelForward Else Return WheelDirections.WheelBackward
        End Get
    End Property

    ''' <summary>
    ''' Returns the direction the horizontal wheel tilt or turn if it was used.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property HorizontalWheelDirection() As HorizontalWheelTiltDirections
        Get
            If (WParam.ToInt32 <> HookBase.WindowsMessages.WM_MOUSEHWHEEL) Then Return HorizontalWheelTiltDirections.none
            If (lParamCopied.mouseData > 0) Then Return HorizontalWheelTiltDirections.Right Else Return HorizontalWheelTiltDirections.Left
        End Get
    End Property

    ''' <summary>
    ''' Returns which XButton was pressed if any.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property XButton() As Xbuttons
        Get
            If (WParam.ToInt32 <> HookBase.WindowsMessages.WM_XBUTTONDOWN) And (WParam.ToInt32 <> HookBase.WindowsMessages.WM_XBUTTONUP) Then Return Xbuttons.none
            Return CType(lParamCopied.mouseData, Xbuttons)
        End Get
    End Property

    ''' <summary>
    ''' Represents the XButton if used by the mouse.
    ''' </summary>
    ''' <remarks></remarks>
    Public Enum Xbuttons As Integer
        ''' <summary>
        ''' The XButton was not used.
        ''' </summary>
        ''' <remarks></remarks>
        none = 0
        ''' <summary>
        ''' Xbutton1 was used.
        ''' </summary>
        ''' <remarks></remarks>
        X1 = 65536
        ''' <summary>
        ''' XButton2 was used.
        ''' </summary>
        ''' <remarks></remarks>
        X2 = 131072
    End Enum

    ''' <summary>
    ''' The direction of the wheel movement if used.
    ''' </summary>
    ''' <remarks></remarks>
    Public Enum WheelDirections As Integer
        ''' <summary>
        ''' The wheel was not used.
        ''' </summary>
        ''' <remarks></remarks>
        none
        ''' <summary>
        ''' The wheel was turned forward, away from the user.
        ''' </summary>
        ''' <remarks></remarks>
        WheelForward
        ''' <summary>
        ''' The wheel was turned backwards, towards the user.
        ''' </summary>
        ''' <remarks></remarks>
        WheelBackward
    End Enum

    ''' <summary>
    ''' The direction the wheel is tilted or horizontal wheel is turned.
    ''' </summary>
    ''' <remarks></remarks>
    Public Enum HorizontalWheelTiltDirections As Integer
        none
        Left
        Right
    End Enum

    ''' <summary>
    ''' The mouse button used.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property MouseButtons() As MouseButtons
        Get
            Select Case CType(WParam, HookBase.WindowsMessages)
                Case HookBase.WindowsMessages.WM_LBUTTONDOWN
                    Return MouseButtons.Left

                Case HookBase.WindowsMessages.WM_LBUTTONUP
                    Return MouseButtons.Left

                Case HookBase.WindowsMessages.WM_RBUTTONDOWN
                    Return MouseButtons.Right

                Case HookBase.WindowsMessages.WM_RBUTTONUP
                    Return MouseButtons.Right

                Case HookBase.WindowsMessages.WM_MBUTTONDOWN
                    Return MouseButtons.Middle

                Case HookBase.WindowsMessages.WM_MBUTTONUP
                    Return MouseButtons.Middle

                Case HookBase.WindowsMessages.WM_XBUTTONDOWN
                    If Me.XButton = Xbuttons.X1 Then
                        Return MouseButtons.XButton1
                    Else
                        Return MouseButtons.XButton2
                    End If

                Case HookBase.WindowsMessages.WM_XBUTTONUP
                    If Me.XButton = Xbuttons.X1 Then
                        Return MouseButtons.XButton1
                    Else
                        Return MouseButtons.XButton2
                    End If

                Case Else
                    Return Nothing

            End Select
        End Get
    End Property
End Class
