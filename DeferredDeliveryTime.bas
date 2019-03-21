Option Explicit

Public g_flag_BusinessTrip As Boolean
Public g_flag_SendImmediate As Boolean


Public Sub SendImmediate()
    g_flag_SendImmediate = True
End Sub


Public Sub SetDeliveryTime(ByVal Item As Object)
    Dim dtDate As String
    Dim dtTime As String
    Dim dtDeliveryDate As String
    Dim enableDDT As Boolean
    
    dtDate = FormatDateTime(Now, vbShortDate)
    dtTime = FormatDateTime(Now, vbShortTime)
    
    Debug.Print "[INFO] Application_ItemSend() :"
    Debug.Print "    Subject: " & Item.Subject
    
    If g_flag_BusinessTrip Then
        g_flag_SendImmediate = True '出張中はすぐに送信する
    End If
        
    If g_flag_SendImmediate Then
        Debug.Print "    Deferred Delivery Time option is disabled."
        g_flag_SendImmediate = False    '一度送信したら配信オプションを有効に戻す
        Exit Sub
    End If
    
    enableDDT = False
    
    If Not (Item.Importance = olImportanceHigh) Then
        If Item.DeferredDeliveryTime = #1/1/4501# Then
            ' 稼働日でない（週末 or 祝日)
            If Not isWorkday(dtDate) Then
                enableDDT = True
            ' 22:00以降の場合（いったん翌日に設定）
            ElseIf dtTime >= FormatDateTime("22:00", vbShortTime) Then
                enableDDT = True
                dtDate = DateSerial(Year(dtDate), Month(dtDate), Day(dtDate) + 1)
            ' 7:00前の場合
            ElseIf dtTime < FormatDateTime("7:00", vbShortTime) Then
                enableDDT = True
            Else
                enableDDT = False
            End If
        Else
            Debug.Print "    DeferredDeliveryTime has already been set."
        End If
    Else
        Debug.Print "    Importance is High, so that this mail will be sent immediately."
    End If
    
    If enableDDT Then
        dtDeliveryDate = getNextWorkday(dtDate)
        Item.DeferredDeliveryTime = DateValue(dtDeliveryDate) + TimeValue("7:00")
        Debug.Print "    This mail will be sent after " & Item.DeferredDeliveryTime
    End If
    
End Sub


Private Function isWorkday(ByRef dtDate) As Boolean
    If Weekday(dtDate, vbMonday) >= 6 Then
        isWorkday = False
        'Debug.Print "    " & dtDate & " is a Weekend"
    ElseIf isHoliday(dtDate) Then
        isWorkday = False
        'Debug.Print "    " & dtDate & " is a Holiday"
    Else
        isWorkday = True
        'Debug.Print "    " & dtDate & " is a Workday"
    End If
End Function


Private Function isHoliday(ByRef dtDate) As Boolean
    Dim holiday() As Variant
    holiday = Array(#1/14/2019#, #2/11/2019#, #3/21/2019#, #4/29/2019#, #4/30/2019#, #5/1/2019#, #5/2/2019#, #5/3/2019#, #5/6/2019#, #7/15/2019#, #8/12/2019#, #9/16/2019#, #9/23/2019#, #10/14/2019#, #10/22/2019#, #11/4/2019#, #12/30/2019#, #12/31/2019#, #1/1/2020#, #1/2/2020#, #1/3/2020#, #1/13/2020#)
    
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


Private Function getNextWorkday(ByRef dtDate) As String
    If Not isWorkday(dtDate) Then
        dtDate = getNextWorkday(DateSerial(Year(dtDate), Month(dtDate), Day(dtDate) + 1))
    End If
    getNextWorkday = dtDate
End Function


Public Function getPrevWorkday(ByRef dtDate) As String
    If Not isWorkday(dtDate) Then
        dtDate = getPrevWorkday(DateSerial(Year(dtDate), Month(dtDate), Day(dtDate) - 1))
    End If
    getPrevWorkday = dtDate
End Function
