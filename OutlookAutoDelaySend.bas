Private Sub Application_ItemSend(ByVal Item As Object, Cancel As Boolean)
    
    dtDate = FormatDateTime(Now, vbShortDate)
    dtTime = FormatDateTime(Now, vbShortTime)
    
    Debug.Print "Subject: " & Item.Subject
    
    Dim flagDeferredDeliveryTime As Boolean
    flagDeferredDeliveryTime = False
    
    If Not (Item.Importance = olImportanceHigh) Then
        If Item.DeferredDeliveryTime = #1/1/4501# Then
            ' 稼働日でない（週末 or 祝日)
            If Not isWorkday(dtDate) Then
                flagDeferredDeliveryTime = True
            ' 22:00以降の場合（いったん翌日に設定）
            ElseIf dtTime >= FormatDateTime("22:00", vbShortTime) Then
                flagDeferredDeliveryTime = True
                dtDate = DateSerial(Year(dtDate), Month(dtDate), day(dtDate) + 1)
            ' 7:00前の場合
            ElseIf dtTime < FormatDateTime("7:00", vbShortTime) Then
                flagDeferredDeliveryTime = True
            Else
                flagDeferredDeliveryTime = False
            End If
    
            If flagDeferredDeliveryTime Then
                dtDeliveryDate = getNextWorkday(dtDate)
                Item.DeferredDeliveryTime = DateValue(dtDeliveryDate) + TimeValue("7:00")
                Debug.Print "   This mail will be sent after " & Item.DeferredDeliveryTime
            End If

        Else
            Debug.Print "   DeferredDeliveryTime has already been set."
        End If
    Else
        Debug.Print "   Importance is High, so that this mail will be sent immediately."
    End If
        
End Sub


Private Function getNextWorkday(ByRef dtDate) As String

    If Not isWorkday(dtDate) Then
        dtDate = getNextWorkday(DateSerial(Year(dtDate), Month(dtDate), day(dtDate) + 1))
    End If
    getNextWorkday = dtDate

End Function


Private Function isWorkday(ByRef dtDate) As Boolean
    
    If Weekday(dtDate, vbMonday) >= 6 Then
        isWorkday = False
        'Debug.Print dtDate & " is a Weekend"
    ElseIf isHoliday(dtDate) Then
        isWorkday = Fale
        'Debug.Print dtDate & " is a Holiday"
    Else
        isWorkday = True
        'Debug.Print dtDate & " is a Workday"
    End If
    
End Function


Private Function isHoliday(ByRef dtDate) As Boolean
    
    Dim holiday() As Variant
    holiday = Array(#3/21/2018#, #4/30/2018#, #5/3/2018#, #5/4/2018#, #7/16/2018#, #8/13/2018#, #8/14/2018#, #9/17/2018#, #9/24/2018#, #10/8/2018#, #11/23/2018#, #12/24/2018#, #1/1/2018#, #1/14/2018#)
    
    isHoliday = False
    
    Dim i As Integer
    For i = LBound(holiday) To UBound(holiday)
        If Not isHoliday Then
            If FormatDateTime(dtDate, vbShortDate) = FormatDateTime(holiday(i), vbShortDate) Then
                isHoliday = True
            End If
        End If
    Next i
        
End Function
