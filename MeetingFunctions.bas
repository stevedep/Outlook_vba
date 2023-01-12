Attribute VB_Name = "MeetingFunctions"
Sub update_subject()
 Set objApp = CreateObject("Outlook.Application")
Set objItem = GetCurrentItem()
    objItem.Subject = InputBox("subject")
    objItem.Save
    Set objApp = Nothing
    Set objItem = Nothing
    
End Sub

Sub deselect()

Dim objApp As Outlook.Application

    Dim objItem As Object
    Dim objAttendees As Outlook.Recipients

    Set objApp = CreateObject("Outlook.Application")
    Set objItem = objApp.ActiveInspector.CurrentItem
    Set objAttendees = objItem.Recipients

    For X = 1 To objAttendees.Count
    On Error Resume Next
            If objAttendees(X).Type = 2 Or objAttendees(X).Type = 3 Then
                objAttendees(X).Sendable = False
               '  MsgBox (objAttendees(x).Type)
            End If
    Next

    Set objApp = Nothing
    Set objItem = Nothing
    Set objAttendees = Nothing

End Sub

Sub GetAttendeeList()
' GetCurrentItem function: https://slipstick.me/9hu-b
Dim objApp As Outlook.Application
Dim objItem As Object
Dim objAttendees As Outlook.Recipients
Dim objAttendeeReq As String
Dim objAttendeeOpt As String
Dim objOrganizer As String
Dim dtStart As Date
Dim dtEnd As Date
Dim strSubject As String
Dim strLocation As String
Dim strNotes As String
Dim strMeetStatus As String
Dim strCopyData As String
Dim strCount  As String
Dim arrAccepted(500) As String
'On Error Resume Next
 
Set objApp = CreateObject("Outlook.Application")

Set oCalendar = Application.Session.GetDefaultFolder(olFolderCalendar)
today = Format(DateAdd("d", -1, Date), "dd/mm/yyyy")
inweek = Format(DateAdd("d", 100, Date), "dd/mm/yyyy")

strFilter = "[Start] > '" & today & "' And [Start] < '" & inweek & "'"
Debug.Print strFilter
 Set calendarItems = oCalendar.Items
calendarItems.IncludeRecurrences = False
Set oItems = calendarItems.Restrict(strFilter)
'oItems.IncludeRecurrences = False
'MsgBox oItems.Count

     
    Set objApp = CreateObject("Outlook.Application")
   'Set objItem = Outlook.Application.ActiveExplorer.Selection.Item(1)

  For Each objItem In oItems
     
    objItem.Subject = objItem.ConversationTopic
    objItem.Save
                            Set objAttendees = objItem.Recipients
                             ia = 0
                             ino = 0
                             it = 0
                             ide = 0
                             a = 0
                            'On Error GoTo EndClean:
                            
                            ' Is it an appointment
                            If objItem.Class <> 26 Then
                              MsgBox "This code only works with meetings."
                            End If
                             
               
                            ' Get The Attendee List
                            For X = 1 To objAttendees.Count
                               strMeetStatus = ""
                              If objAttendees(X).Type = 1 Then
                              a = a + 1
                               Select Case objAttendees(X).MeetingResponseStatus
                                 Case 0
                                   strMeetStatus = "No Response"
                                   ino = ino + 1
                                 Case 1
                                   strMeetStatus = "Organizer"
                                   ino = ino + 1
                                 Case 2
                                   strMeetStatus = "Tentative"
                                   it = it + 1
                                 Case 3
                                   strMeetStatus = "Accepted"
                                   arrAccepted(ia) = objAttendees(X).Name
                                   ia = ia + 1
                                 Case 4
                                   strMeetStatus = "Declined"
                                   ide = ide + 1
                               End Select
                   End If
                            Next
                              
  
                             Strvar = objItem.Subject
                             StartPos = 1
                             If InStr(1, Strvar, ") ", vbTextCompare) > 0 Then
                                StartPos = InStr(1, Strvar, ") ", vbTextCompare)
                             End If
                             sbj = Mid(Strvar, StartPos, Len(Strvar) - StartPos + 1)
                            Dim bodystring As String
                            bodystring = objItem.Body
                        
                        'reset first
                        objItem.Categories = ""
                         Dim catstring  As String
                         catstring = ""
                        
                             
                        If ia >= a Then
                                    objItem.Subject = "(" & ia & "/" & a - 1 & ") " & sbj
                                    Debug.Print "(" & ia & "/" & a - 1 & ") " & sbj
                                   objItem.Categories = "AllAccepted"  '"PartialAcceptance"
                                   objItem.Save
                            ElseIf ia > 0 Then
                                objItem.Subject = "(" & ia & "/-" & ide & "/" & a - 1 & ") " & sbj
                                  ' objItem.Categories = "PartialAcceptance"
                                 If ia / (X - 2) <= 0.2 Then
                                     objItem.Categories = "Red"
                                ElseIf ia / (X - 2) <= 0.4 Then
                                    objItem.Categories = "Orange"
                                ElseIf ia / (X - 2) <= 0.6 Then
                                    objItem.Categories = "Peach"
                                ElseIf ia / (X - 2) <= 0.8 Then
                                    objItem.Categories = "Yellow"
                                ElseIf ia / (X - 2) <= 0.99 Then
                                    objItem.Categories = "LightGreen"
                                End If
                                Debug.Print ia / (X - 2)

                                 Debug.Print catstring & "PartialAcceptance"
                     '              For Each acceptant In arrAccepted
      'show the element in the debug window.
                                        
                              '          If Len(acceptant) > 1 Then
                               '             bodystring = vbTab & acceptant & vbLf & bodystring
                                '        End If
                          '      Next acceptant
                           '     objItem.Body = bodystring
                          '      Debug.Print bodystring
 'strRTF = StrConv(objItem.RTFBody, vbUnicode)
 
 'Debug.Print strRTF
                   '             objItem.RTFBody = "test" & vbNewLine & objItem.RTFBody
                                'onderstaande herstellen
'                                objItem.Categories = "PartialAcceptance"
                                Debug.Print "(" & ia & "/" & a - 1 & ") " & sbj
                                objItem.Save
                            ElseIf ia = 0 Then
                                objItem.Subject = "(" & ia & "/-" & ide & "/" & a - 1 & ") " & sbj
                                Debug.Print "(" & ia & "/" & a - 1 & ") " & sbj
                                objItem.Categories = "NoneAccepted"
                               objItem.Save
                            End If
                            'Set ListAttendees = Application.CreateItem(olMailItem)
                            '  ListAttendees.Body = strCopyData & vbCrLf & strCount
                            '  ListAttendees.Display
                               
                        If ide > 1 Then
                                       objItem.Categories = objItem.Categories & "; DarkRed"
                                       objItem.Save
                        End If
                            Set objApp = Nothing
                            Set objItem = Nothing
                            Set objAttendees = Nothing

    Next


End Sub



Sub test()
Strvar = "(4/7) Churn analytics bi-weekly"
StartPos = InStr(1, Strvar, ")", vbTextCompare)
MsgBox Mid(Strvar, StartPos + 2, Len(Strvar) - StartPost)

End Sub


Function GetCurrentItem() As Object
    Dim objApp As Outlook.Application
           
    Set objApp = Application
    On Error Resume Next
    Select Case TypeName(objApp.ActiveWindow)
        Case "Explorer"
            Set GetCurrentItem = objApp.ActiveExplorer.Selection.Item(1)
        Case "Inspector"
            Set GetCurrentItem = objApp.ActiveInspector.CurrentItem
    End Select
       GetCurrentItem.Start = #9/24/2003 1:30:00 PM#
    Set objApp = Nothing
End Function
