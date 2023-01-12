Attribute VB_Name = "Mod_ShowformwithAttendees"
Sub GetAttendeeList_msgbox()
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
 

    Set objItem = Outlook.Application.ActiveExplorer.Selection.Item(1)
'MsgBox objItem.Categories
                            Set objAttendees = objItem.Recipients
                             ia = 0
                             ino = 0
                             it = 0
                             ide = 0
                            'On Error GoTo EndClean:
                            
                            ' Is it an appointment
                            If objItem.Class <> 26 Then
                              MsgBox "This code only works with meetings."
                              GoTo EndClean:
                            End If
                             
               
                            ' Get The Attendee List
                            For X = 1 To objAttendees.Count
                               strMeetStatus = ""
                               Select Case objAttendees(X).MeetingResponseStatus
                                 Case 0
                                   strMeetStatus = "No Response"
                                   frmAcceptance.lstNone.AddItem objAttendees(X).Name
                                   ino = ino + 1
                                 Case 1
                                   strMeetStatus = "Organizer"
                                   ino = ino + 1
                                 Case 2
                                   strMeetStatus = "Tentative"
                                   frmAcceptance.lstTentative.AddItem objAttendees(X).Name
                                   it = it + 1
                                 Case 3
                                   strMeetStatus = "Accepted"
                                   frmAcceptance.lstAccepted.AddItem objAttendees(X).Name
                                   ia = ia + 1
                                 Case 4
                                   strMeetStatus = "Declined"
                                   frmAcceptance.lstDeclined.AddItem objAttendees(X).Name
                                   ide = ide + 1
                               End Select
                   
                           Next
                              
  frmAcceptance.Show
                               
EndClean:
Debug.Print "gat iets foiut"
                            Set objApp = Nothing
                            Set objItem = Nothing
                            Set objAttendees = Nothing

    


End Sub

