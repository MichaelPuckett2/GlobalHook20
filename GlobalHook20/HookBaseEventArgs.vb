Public Class HookBaseEventArgs
    Inherits EventArgs

    ''' <summary>
    ''' Instatiates the KeyEventArgs Class with the same signature returned to our HookProc Delegate returned with SetWindowsHookEx.
    ''' </summary>
    ''' <param name="nCode">[in] Specifies a code the hook procedure uses to determine how to process the message. If nCode is less than zero, the hook procedure must pass the message to the CallNextHookEx function without further processing and should return the value returned by CallNextHookEx. This parameter can be one of the following values. 
    ''' HC_ACTION
    ''' The wParam and lParam parameters contain information about a keyboard message.</param>
    ''' <param name="wParam">[in] Specifies the identifier of the keyboard message. This parameter can be one of the following messages: WM_KEYDOWN, WM_KEYUP, WM_SYSKEYDOWN, or WM_SYSKEYUP. </param>
    ''' <param name="lParam">[in] Pointer to a KBDLLHOOKSTRUCT structure. </param>
    ''' <remarks></remarks>
    Public Sub New(ByVal nCode As Integer, ByVal wParam As IntPtr, ByVal lParam As IntPtr)
        Me.LParam = lParam
        Me.NCode = nCode
        Me.WParam = wParam
    End Sub

    ''' <summary>
    ''' For information or wrapping purposes only.
    ''' [in] Specifies a code the hook procedure uses to determine how to process the message. If nCode is less than zero, the hook procedure must pass the message to the CallNextHookEx function without further processing and should return the value returned by CallNextHookEx. This parameter can be one of the following values. 
    ''' HC_ACTION
    ''' The wParam and lParam parameters contain information about a keyboard message.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property NCode() As Integer

    ''' <summary>
    ''' [in] Specifies the identifier of the keyboard message. This parameter can be one of the following messages: WM_KEYDOWN, WM_KEYUP, WM_SYSKEYDOWN, or WM_SYSKEYUP. 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property WParam() As IntPtr

    ''' <summary>
    ''' [in] Pointer to a KBDLLHOOKSTRUCT structure.
    ''' The exact same as KeyboardHookStruct but named to match MSDN to not confuse developers wrapping this DLL.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property LParam() As IntPtr

    ''' <summary>
    ''' Determines if the current key should be handled or not meaning if Handled = True that key event will not pass to any other applications.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Handled() As Boolean

End Class
