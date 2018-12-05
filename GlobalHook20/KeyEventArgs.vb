Imports System.Runtime.InteropServices

''' <summary>
''' Passed with KeyEvents from the Keyboard class with all the information needed to read and handle the current key event.
''' </summary>
''' <remarks></remarks>
Public Class KeyEventArgs
    Inherits HookBaseEventArgs

    ''' <summary>
    ''' True if the key is an extended key.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>The flag value in KBDLLHOOKSTRUCTFlags.LLKHF_EXTENDED compared to the flag sent with the event.</remarks>
    Public ReadOnly Property ExtendedKey() As Boolean
        Get
            Return KeyboardHookStruct.flags = (KeyboardHookStruct.flags Or Keyboard.KBDLLHOOKSTRUCTFlags.LLKHF_EXTENDED)
        End Get
    End Property

    ''' <summary>
    '''True if the current key event was injected; meaning another application inserted the event and the keyboard was not physically used.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>The flag value in KBDLLHOOKSTRUCTFlags.LLKHF_INJECTED compared to the flag sent with the event.</remarks>
    Public ReadOnly Property Injected() As Boolean
        Get
            Return KeyboardHookStruct.flags = (KeyboardHookStruct.flags Or Keyboard.KBDLLHOOKSTRUCTFlags.LLKHF_INJECTED)
        End Get
    End Property

    ''' <summary>
    ''' True if the Alt Key is down.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>The flag value in KBDLLHOOKSTRUCTFlags.LLKHF_INJECTED compared to the flag sent with the event.</remarks>
    Public ReadOnly Property AltDown() As Boolean
        Get
            Return KeyboardHookStruct.flags = (KeyboardHookStruct.flags Or Keyboard.KBDLLHOOKSTRUCTFlags.LLKHF_ALTDOWN)
        End Get
    End Property

    ''' <summary>
    ''' True if the Key is up.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>The flag value in KBDLLHOOKSTRUCTFlags.LLKHF_INJECTED compared to the flag sent with the event.</remarks>
    Public ReadOnly Property KeyUp() As Boolean
        Get
            Return KeyboardHookStruct.flags = (KeyboardHookStruct.flags Or Keyboard.KBDLLHOOKSTRUCTFlags.LLKHF_UP)
        End Get
    End Property

    ''' <summary>
    ''' Returns Forms.Key associated with the key event.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Key() As keys
        Get
            Return CType(KeyboardHookStruct.vkCode, Keys)
        End Get
    End Property

    ''' <summary>
    ''' Returns all the ASCII character as Char associated with the key event.
    ''' Returns Nothing if there is no character associated with the key event.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property ASCIIChar() As Char
        Get
            Return ConvertToCorrectASCII(KeyboardHookStruct.vkCode)
        End Get
    End Property

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
        MyBase.New(nCode, wParam, lParam)
        KeyboardHookStruct = CType(Marshal.PtrToStructure(lParam, GetType(Keyboard.KBDLLHOOKSTRUCT)), Keyboard.KBDLLHOOKSTRUCT)
    End Sub

    ''' <summary>
    ''' References the KBDLLHOOKSTRUCT associated with the current key event.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property KeyboardHookStruct() As Keyboard.KBDLLHOOKSTRUCT

    ''' <summary>
    ''' Converts the appropriate vkCode to a Char from the ASCII char.  Returns nothing if no charact is available.
    ''' </summary>
    ''' <param name="vkCode">The vistual key code passed in from the vkCode from within KeyboardHookStruct</param>
    ''' <returns>The appropriate ASCII key character or nothing if there is no character to return.</returns>
    ''' <remarks></remarks>
    Private Function ConvertToCorrectASCII(ByVal vkCode As UInteger) As Char
        'Check and send space character
        If vkCode = 32 Then Return Chr(CInt(vkCode))
        'Check and send A-z; then 1-0, then Shift+ 1-0
        If vkCode >= 65 And vkCode <= 90 Then
            If (My.Computer.Keyboard.CapsLock And Not (My.Computer.Keyboard.ShiftKeyDown)) Or (Not My.Computer.Keyboard.CapsLock And (My.Computer.Keyboard.ShiftKeyDown)) Then
                Return Chr(CInt(vkCode))
            Else
                Return Chr(CInt(vkCode) + 32)
            End If
        ElseIf vkCode >= 48 And vkCode <= 57 Then
            If Not My.Computer.Keyboard.ShiftKeyDown Then
                Return Chr(CInt(vkCode))
            Else
                Select Case vkCode
                    Case 48
                        Return CChar(")")
                    Case 49
                        Return CChar("!")
                    Case 50
                        Return CChar("@")
                    Case 51
                        Return CChar("#")
                    Case 52
                        Return CChar("$")
                    Case 53
                        Return CChar("%")
                    Case 54
                        Return CChar("^")
                    Case 55
                        Return CChar("&")
                    Case 56
                        Return CChar("*")
                    Case 57
                        Return CChar("(")
                End Select
            End If
        End If
        'Check and send all remaining keys with 2 positional characters ie, with and without shift down.
        If My.Computer.Keyboard.ShiftKeyDown Then
            Select Case vkCode
                Case 192
                    Return CChar("~")
                Case 189
                    Return CChar("_")
                Case 187
                    Return CChar("+")
                Case 219
                    Return CChar("{")
                Case 221
                    Return CChar("}")
                Case 220
                    Return CChar("|")
                Case 186
                    Return CChar(":")
                Case 222
                    Return CChar("""")
                Case 188
                    Return CChar("<")
                Case 190
                    Return CChar(">")
                Case 191
                    Return CChar("?")
            End Select
        Else
            Select Case vkCode
                Case 192
                    Return CChar("`")
                Case 189
                    Return CChar("-")
                Case 187
                    Return CChar("=")
                Case 219
                    Return CChar("[")
                Case 221
                    Return CChar("]")
                Case 220
                    Return CChar("\")
                Case 186
                    Return CChar(";")
                Case 222
                    Return CChar("'")
                Case 188
                    Return CChar(",")
                Case 190
                    Return CChar(".")
                Case 191
                    Return CChar("/")
            End Select
        End If
        'Check and send NumPad characters after checking NumLock and Shift down status.
        If vkCode >= 96 And vkCode <= 105 Then
            If My.Computer.Keyboard.NumLock Then
                If Not My.Computer.Keyboard.ShiftKeyDown Then
                    Return Chr(CInt(vkCode) - 48)
                End If
            End If
        End If
        Select Case vkCode
            Case 111
                Return CChar("/")
            Case 106
                Return CChar("*")
            Case 109
                Return CChar("-")
            Case 107
                Return CChar("+")
        End Select
        'Return Nothing for non character keys.
        Return Nothing
    End Function

    ''' <summary>
    ''' Returns the position or toggled state of system keys suck as Alt, Shift, ScrollLock, NumLock....
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property SysKeys() As Devices.Keyboard
        Get
            Dim kb As New Devices.Keyboard
            Return kb
        End Get
    End Property

End Class