Attribute VB_Name = "GetOpenMeetingTimes"
Public Sub GetOpenMeetingTimes()
    Dim myNameSpace As Outlook.NameSpace
    Dim myRecipient As Outlook.RECIPIENT
    Dim timeslot, start_time, timeslot_num, FBInfoMerged As String
    Dim MEETING_DATE As Variant
    Dim RECIPIENTS As Variant
    Dim PEOPLE As String
    Dim FBInfo(20)
    Dim LArray() As String
    
    
    MEETING_DATE = Array(#5/24/2022#, #5/25/2022#, #5/26/2022#)
    PEOPLE = ""
    
    
    RECIPIENTS = Split(PEOPLE, "; ")
    Set myNameSpace = Application.GetNamespace("MAPI")
    arrayLen = UBound(RECIPIENTS) - LBound(RECIPIENTS) + 1
    arrayLenDate = UBound(MEETING_DATE) - LBound(MEETING_DATE) + 1

    For o = 0 To arrayLen - 1
        Debug.Print RECIPIENTS(o)
    Next o

    For a = 0 To arrayLenDate - 1
        FBInfoMerged = ""
        Debug.Print MEETING_DATE(a)
        For j = 0 To arrayLen - 1
            FBInfo(j) = Mid(myNameSpace.CreateRecipient(RECIPIENTS(j)).FreeBusy(MEETING_DATE(a), 30, False), 1, 48)
            On Error GoTo ErrorHandler
            'Debug.Print MEETING_DATE(a) & " - " & RECIPIENTS(j)
            'Debug.Print FBInfo(j)
        Next j
        
        For g = 1 To 48
            timeslot_num = 0
            For k = 0 To arrayLen - 1
                If Mid(FBInfo(k), g, 1) <> 0 Then
                    timeslot_num = 1
                    Exit For
                End If
            Next k
            FBInfoMerged = FBInfoMerged & timeslot_num
        Next g
        
        'Debug.Print FBInfoMerged
        
        start_time = "12:00:00 AM"
        For i = 1 To 48
            If Mid(FBInfoMerged, i, 1) = 0 Then
                timeslot = "FREE"
            Else
                timeslot = "BUSY"
            End If
    
            If (TimeValue(start_time) >= TimeValue("7:00:00 AM")) And (TimeValue(start_time) < TimeValue("6:30:00 PM")) And timeslot <> "BUSY" Then
                Debug.Print Format(start_time, "hh:mm AM/PM") & " - " & Format(DateAdd("n", 30, start_time), "hh:mm AM/PM") '& ": " & timeslot
            End If
            
            start_time = DateAdd("n", 30, start_time)
        Next i
    Next a
    
    Exit Sub
ErrorHandler:
    MsgBox "Cannot access the information. "
End Sub
