Public Module SerialKey

    Private Const SERIAL_KEY As String = "e1ffc5fc-4819-4340-a49d-065e4885c3bb"
    Public Const I_LOVE_GOD_FOREVER As String = "Thank you God for all the wisdom, love, mercy, and forgiveness you have given me.  " & _
                                                "Thank you for the talent to do what all I can do and I pray that I do it all to gloriy you.  " & _
                                                "Thank you so much for my family.  I love you and praise you and lift your name higher forever and ever. I love you God with all my heart."
    Public Property SerialKey() As String

    Public Sub CheckKey()
        If SerialKey Is Nothing Then Throw New Exception("Serial Key Required")
        If SERIAL_KEY.ToUpper <> SerialKey Then Throw New Exception("Invalid Serial Key")
    End Sub

End Module
