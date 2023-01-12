Attribute VB_Name = "Mod_Scheduler"
Sub check_availability()
Dim myOlApp                                                                                         As New Outlook.Application
Set myNameSpace = myOlApp.GetNamespace("MAPI")
Set objApp = CreateObject("Outlook.Application")
'Set objItem = Outlook.Application.ActiveExplorer.Selection.Item(1)
Set objItem = objApp.ActiveInspector.CurrentItem
'Set objAttendees = Outlook.Application.ActiveExplorer.Selection.Item(1).Recipients
Set objAttendees = objItem.Recipients
Dim oCurrentUser                                                                                    As Recipient
Dim FreeBusy(20, 2)                                                                                 As String
Dim BusySlot                                                                                        As Long
Dim DateBusySlot                                                                                    As Date
Dim i                                                                                               As Long
Const SlotLength = 30
Dim teller                                                                                          As Integer
Dim eruit                                                                                           As Boolean
eruit = False
Dim aantalre As Integer
Dim StartDate As Date
StartDate = Format(objItem.Start, "dd-mm-yyyy") 'DateAdd("n", SlotLength, objItem.Start)
'Debug.Print StartDate
Dim delta As Variant
delta0 = (DateDiff("n", DataValue(Format(objItem.Start, "dd-mm-yyyy")), objItem.Start) / SlotLength)
st = CDate(CLng(objItem.Start))
delta = ((DateDiff("n", st, objItem.Start)) / SlotLength)
aantalre = 0
For X = 1 To objAttendees.Count
    If (objAttendees(X).Type = 1 And objAttendees(X).Sendable = True) Or objAttendees(X).Index = 1 Then
        Set myRecipient = myNameSpace.CreateRecipient(objAttendees(X).Address)
       On Error Resume Next
        FreeBusy(X, 1) = myRecipient.FreeBusy(StartDate, SlotLength, True)
        'debug.print objAttendees(X).Name & " " & vbCrLf & Left(FreeBusy(X, 1), 200)
                If Err.Number < 0 Then
                       MsgBox "Unable to get Calendar for " & objAttendees(X).Name
                End If
        FreeBusy(X, 2) = objAttendees(X).Name
    aantalre = aantalre + 1
    End If
Next
'debug.print Len(FreeBusy(1, 1))
Dim Message, Title, Default, aantal
Message = "Enter number of weeks (max 4)"    ' Set prompt.
Title = "Input for max number of weeks"    ' Set title.
Default = "1"    ' Set default.
' Display message, title, and default value.
aantal = InputBox(Message, Title, Default)
For re = 0 To objAttendees.Count
For i = delta + 2 To (aantal * 7 * 24 / (SlotLength / 60)) + delta + 2 'delta is the adjustment for the hours, since FreeBusy works with whole days
    teller = 0
    For Y = 1 To 20
        If Len(FreeBusy(Y, 1)) > 1 Then
            If CLng(Mid(FreeBusy(Y, 1), i, 1)) = 0 Or CLng(Mid(FreeBusy(Y, 1), i, 1)) = 1 Then
                teller = teller + 1
            End If
            
            If teller = aantalre - re Then
                BusySlot = (i - 1) * SlotLength
                DateBusySlot = DateAdd("n", BusySlot, StartDate)
                If TimeValue(DateBusySlot) >= TimeValue(#9:00:00 AM#) And TimeValue(DateBusySlot) <= TimeValue(#5:00:00 PM#) And Not (Weekday(DateBusySlot) = vbSaturday Or Weekday(DateBusySlot) = vbSunday) Then
                Debug.Print " first open interval:" & Y & "I: " & FreeBusy(Y, 2) & "i: " & i & _
                   vbCrLf & _
                   Format$(DateBusySlot, "mm\/dd\/yyyy hh:mm AM/PM")
                MsgBox "Found slot for " & aantalre - re & "/" & aantalre & " participants"
                'objItem.Start = Format$(DateBusySlot, "mm\/dd\/yyyy hh:mm AM/PM") 20221229, changed below as well
                objItem.Start = Format$(DateBusySlot, "dd\/mm\/yyyy hh:mm AM/PM")
                eruit = True
                Exit For
            End If        ' close when found
        End If        'close when all recipients have been evaluated
    End If 'close when all recipients in the array have been evaluated
    If eruit = True Then
        Exit For
    End If
Next 'Next Y, next recipient in array
If eruit = True Then
    Exit For
End If
Next 'next time entry in array
If eruit = True Then
    Exit For
End If
Next

Set objApp = Nothing
Set objItem = Nothing
Set objAttendees = Nothing
Set myNameSpace = Nothing


End Sub


