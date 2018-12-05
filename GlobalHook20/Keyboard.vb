Imports System.Runtime.InteropServices

''' <summary>
''' Used to perform a low level hook on the keyboard then read and fire key event information.
''' </summary>
''' <remarks></remarks>
Public Class Keyboard
    Inherits HookBase

    Public Sub New()
        MyBase.New(HookType.WH_KEYBOARD_LL)
    End Sub

    ''' <summary>
    ''' Runs when Hook is called.
    ''' </summary>
    ''' <remarks></remarks>
    Protected Overrides Sub OnHook()
        StoredKeyPositions = New Dictionary(Of Keys, Boolean)
    End Sub

    ''' <summary>
    ''' Runs when UnHook is called.
    ''' </summary>
    ''' <remarks></remarks>
    Protected Overrides Sub OnUnHook()
        StoredKeyPositions = Nothing
    End Sub

    ''' <summary>
    ''' The Method address delegated to the HookProc Delegate and assigned to the HookProc_Delegated variable.
    ''' </summary>
    ''' <param name="nCode">[in] Specifies a code the hook procedure uses to determine how to process the message. If nCode is less than zero, the hook procedure must pass the message to the CallNextHookEx function without further processing and should return the value returned by CallNextHookEx. This parameter can be one of the following values. 
    ''' HC_ACTION
    ''' The wParam and lParam parameters contain information about a keyboard message.</param>
    ''' <param name="wParam">[in] Specifies the identifier of the keyboard message. This parameter can be one of the following messages: WM_KEYDOWN, WM_KEYUP, WM_SYSKEYDOWN, or WM_SYSKEYUP. </param>
    ''' <param name="lParam">[in] Pointer to a KBDLLHOOKSTRUCT structure. </param>
    ''' <returns>True if that event should be handled and not passed to other applications.</returns>
    ''' <remarks></remarks>
    Protected Overrides Function OnHookProc(ByVal nCode As Integer, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As Boolean
        Dim wParammm As WindowsMessages = CType(wParam.ToInt32, WindowsMessages)
        Dim lParammm As KBDLLHOOKSTRUCT = CType(Marshal.PtrToStructure(lParam, GetType(KBDLLHOOKSTRUCT)), KBDLLHOOKSTRUCT)

        Dim KEA As New KeyEventArgs(nCode, wParam, lParam)

        RaiseEvent KeyEvent(Me, KEA)

        Select Case wParammm
            Case WindowsMessages.WM_KEYDOWN
                StoredKeyPositions(CType(lParammm.vkCode, Keys)) = True
                RaiseEvent KeyDown(Me, KEA)

            Case WindowsMessages.WM_KEYUP
                StoredKeyPositions(CType(lParammm.vkCode, Keys)) = False
                RaiseEvent KeyUp(Me, KEA)

            Case WindowsMessages.WM_SYSKEYDOWN
                StoredKeyPositions(CType(lParammm.vkCode, Keys)) = True
                RaiseEvent SysKeyDown(Me, KEA)

            Case WindowsMessages.WM_SYSKEYUP
                StoredKeyPositions(CType(lParammm.vkCode, Keys)) = False
                RaiseEvent SysKeyUp(Me, KEA)

        End Select

        Return KEA.Handled
    End Function

#Region "Key Events"

    ''' <summary>
    ''' Generic KeyEvent, raised for every key stroke event. Happens before all other Key Events.
    ''' If Handled is set to True the remaining appropriate events KeyUp, KeyDown, SysKeyUp, and SysKeyDown events will still fire but the Key will still be handled for other applications.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Public Event KeyEvent(ByVal sender As Object, ByVal e As KeyEventArgs)

    ''' <summary>
    ''' Fired when a non-system key is pressed. Will fire multiple times if that key is held down and no other keys are pressed.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Public Event KeyDown(ByVal sender As Object, ByVal e As KeyEventArgs)

    ''' <summary>
    ''' Fired when a non-system key is released.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Public Event KeyUp(ByVal sender As Object, ByVal e As KeyEventArgs)

    ''' <summary>
    ''' Fired when a System Key is Pressed. System keys are the Alt Keys or another Key pressed while the Alt key is still down. Event will fire multiple times if that key is held down and no other keys are pressed.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Public Event SysKeyDown(ByVal sender As Object, ByVal e As KeyEventArgs)

    ''' <summary>
    ''' Fired when a System Key other than the Alt key is released. This is NOT an Alt key but another key being released while the Alt key is still pressed.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Public Event SysKeyUp(ByVal sender As Object, ByVal e As KeyEventArgs)

#End Region

    ''' <summary>
    ''' The Structure returned from the SetWindowsHookEx as the lParam value. Returns information about the key event.
    ''' </summary>
    ''' <remarks></remarks>
    <StructLayout(LayoutKind.Sequential)>
    Public Structure KBDLLHOOKSTRUCT
        ''' <summary>
        ''' Virtual Key Code returned with the key event.
        ''' </summary>
        ''' <remarks></remarks>
        Public vkCode As UInteger
        ''' <summary>
        ''' Hardware Scancode returned with the key event.
        ''' </summary>
        ''' <remarks></remarks>
        Public scanCode As UInteger
        ''' <summary>
        ''' The bit-wise flags returned from the key event as BLDLLHOOKSTRUCTFlags.
        ''' </summary>
        ''' <remarks></remarks>
        Public flags As KBDLLHOOKSTRUCTFlags
        ''' <summary>
        ''' The time of the key event from the time of system start.
        ''' </summary>
        ''' <remarks></remarks>
        Public time As UInteger
        ''' <summary>
        ''' Extra information associated with the key event; currently not used.
        ''' </summary>
        ''' <remarks></remarks>
        Public dwExtraInfo As IntPtr
    End Structure

    ''' <summary>
    ''' The bitwise flags return with the HookProc delegate assigned to the SetWindowsHookProc.
    ''' Specifies the extended-key flag, event-injected flag, context code, and transition-state flag. This member is specified as follows. An application can use the following values to test the keystroke flags. Value Purpose 
    ''' LLKHF_EXTENDED Test the extended-key flag.  
    ''' LLKHF_INJECTED Test the event-injected flag.  
    ''' LLKHF_ALTDOWN Test the context code.  
    ''' LLKHF_UP Test the transition-state flag.
    ''' 0
    ''' Specifies whether the key is an extended key, such as a function key or a key on the numeric keypad. The value is 1 if the key is an extended key; otherwise, it is 0.
    ''' 1-3
    ''' Reserved.
    ''' 4
    ''' Specifies whether the event was injected. The value is 1 if the event was injected; otherwise, it is 0.
    ''' 5
    ''' Specifies the context code. The value is 1 if the ALT key is pressed; otherwise, it is 0.
    ''' 6
    ''' Reserved.
    ''' 7
    ''' Specifies the transition state. The value is 0 if the key is pressed and 1 if it is being released.
    ''' </summary>
    ''' <remarks></remarks>
    <Flags()>
    Public Enum KBDLLHOOKSTRUCTFlags As UInteger
        ''' <summary>
        ''' Test the extended-key flag. 
        ''' </summary>
        ''' <remarks></remarks>
        LLKHF_EXTENDED = &H1
        ''' <summary>
        ''' Test the event-injected flag.
        ''' </summary>
        ''' <remarks></remarks>
        LLKHF_INJECTED = &H10
        ''' <summary>
        ''' Test the context code. 
        ''' </summary>
        ''' <remarks></remarks>
        LLKHF_ALTDOWN = &H20
        ''' <summary>
        ''' Test the transition-state flag. 
        ''' </summary>
        ''' <remarks></remarks>
        LLKHF_UP = &H80
    End Enum

    ''' <summary>
    ''' A generic dictionary of Forms.Keys that stores a boolean for each key pressed = true is the key is down and false when it is released.
    ''' </summary>
    ''' <remarks></remarks>
    Private StoredKeyPositions As Dictionary(Of Keys, Boolean)

    ''' <summary>
    ''' A function that takes a Forms.Keys Argument, tests it and returns True if the key is currently down and false if it is up.
    ''' This method only works when the keyboard is hooked and will not return true for keys that are down before the keyboard was hooked.
    ''' </summary>
    ''' <param name="key"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function IsKeyDown(ByVal key As Keys) As Boolean
        If StoredKeyPositions Is Nothing Then Throw New Exception("The Keyboard must be successfully hooked in order to retrieve the key position")
        If Not StoredKeyPositions.ContainsKey(key) Then Return False
        Return StoredKeyPositions(key)
    End Function

    ''' <summary>
    ''' Takes an Array of Forms.Keys as an argument and test them all and returns True if they are all down and false if any of them are not down.
    ''' This method only works when the keyboard is hooked and will not return true for keys that are down before the keyboard was hooked.
    ''' </summary>
    ''' <param name="keys"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function AreKeysDown(ByVal ParamArray keys() As Keys) As Boolean
        If StoredKeyPositions Is Nothing Then Throw New Exception("The Keyboard must be successfully hooked in order to retrieve the key position")
        Dim IsDown As Boolean
        For Each k As Keys In keys
            If Not StoredKeyPositions.ContainsKey(k) Then Return False
            IsDown = StoredKeyPositions(k)
            If Not IsDown Then Return False
        Next
        Return IsDown
    End Function

End Class
