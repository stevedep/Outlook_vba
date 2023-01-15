VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisOutlookSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Public Function URLEncode( _
   StringVal As String, _
   Optional SpaceAsPlus As Boolean = False _
) As String
 
  Dim StringLen As Long: StringLen = Len(StringVal)
 
  If StringLen > 0 Then
    ReDim result(StringLen) As String
    Dim i As Long, CharCode As Integer
    Dim Char As String, Space As String
 
    If SpaceAsPlus Then Space = "+" Else Space = "%20"
 
    For i = 1 To StringLen
      Char = Mid$(StringVal, i, 1)
      CharCode = Asc(Char)
      Select Case CharCode
        Case 97 To 122, 65 To 90, 48 To 57, 45, 46, 95, 126
          result(i) = Char
        Case 32
          result(i) = Space
        Case 0 To 15
          result(i) = "%0" & Hex(CharCode)
        Case Else
          result(i) = "%" & Hex(CharCode)
      End Select
    Next i
    URLEncode = Join(result, "")
  End If
End Function
Public Sub GetFreeBusyInfo()
    Dim myOlApp As New Outlook.Application
    Dim myNameSpace As Outlook.NameSpace
    Dim myRecipient As Outlook.Recipient
    Dim myFBInfo As String
    Set myNameSpace = myOlApp.GetNamespace("MAPI")
    Set myRecipient = myNameSpace.CreateRecipient("Raf Brouwers")
    On Error GoTo ErrorHandler
    myFBInfo = myRecipient.FreeBusy(Now, 60)
    MsgBox myFBInfo
    Exit Sub
ErrorHandler:
        MsgBox "Cannot access the information. "
End Sub

Sub restore_title()
  Dim objApp As Outlook.Application
Dim objItem As Object


    Set objItem = Outlook.Application.ActiveExplorer.Selection.Item(1)
    objItem.Subject = objItem.ConversationTopic
    objItem.Save
    
Set objApp = Nothing
Set objItem = Nothing
End Sub



Sub DemoSetSchedulingStartTime()
    Dim objApp As Outlook.Application
           
    Set objApp = Application
    Select Case TypeName(objApp.ActiveWindow)
        Case "Explorer"
            Set GetCurrentItem = objApp.ActiveExplorer.Selection.Item(1)
        Case "Inspector"
            Set GetCurrentItem = objApp.ActiveInspector.CurrentItem
    End Select
    
    Set oAppt = GetCurrentItem
 
 Set oInsp = oAppt.GetInspector
 
 ' Switch to the Scheduling Assistant tab in that inspector.
 
 oInsp.SetCurrentFormPage ("Scheduling Assistant")
 
 ' Set the appointment start time in the Scheduling Assistant.
 
 'oInsp.SetSchedulingStartTime (DateAdd("m", 1, Now))
    
 'GetCurrentItem.SetSchedulingStartTime (DateAdd("m", 1, Now))
    Set objApp = Nothing
 
 
End Sub



Sub open_webpage()
 
    Dim objApp As Outlook.Application
    Dim objItem As Object
    Dim objAttendees As Outlook.Recipients
    Dim strNames As String
    Dim chromePath As String
     
    Set objApp = CreateObject("Outlook.Application")
    Set objItem = Outlook.Application.ActiveExplorer.Selection.Item(1)
    Set objAttendees = Outlook.Application.ActiveExplorer.Selection.Item(1).Recipients
    chromePath = """C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"""
 
    For X = 1 To objAttendees.Count
    'console.Log objAttendees(x).Adress
            strNames = strNames & Right(objAttendees(X).Address, Len(objAttendees(X).Address) - InStrRev(objAttendees(X).Address, "-")) & ";"
    Next
     
   ' strNames = URLEncode(strNames)
     
    Dim stradres As String
    stradres = "http://localhost/orgchart.html?Names=" & strNames
    Shell (chromePath & " -url " & stradres)
 
     
    Set objApp = Nothing
    Set objItem = Nothing
    Set objAttendees = Nothing
 
End Sub

Sub ReplacebyImportantL1()
   Dim obApp As Object
    Dim NewEmail As MailItem
    Set obApp = Outlook.Application
    Set myOlExp = Application.ActiveExplorer
    myOlExp.Activate
    Set myOlsel = myOlExp.Selection
        c = myOlsel.Item(1).Categories
        c = Replace(c, Mid(c, InStr(c, "ToDo"), 6), "ImportantL1")
        'MsgBox c
        myOlsel.Item(1).Categories = c
        myOlsel.Item(1).Save
     
    Set obApp = Nothing
    Set NewEmail = Nothing
End Sub

Sub ReplacebyImportantL2()
   Dim obApp As Object
    Dim NewEmail As MailItem
    Set obApp = Outlook.Application
    Set myOlExp = Application.ActiveExplorer
    myOlExp.Activate
    Set myOlsel = myOlExp.Selection
        c = myOlsel.Item(1).Categories
        c = Replace(c, Mid(c, InStr(c, "ToDo"), 6), "ImportantL2")
        'MsgBox c
        myOlsel.Item(1).Categories = c
        myOlsel.Item(1).Save
     
    Set obApp = Nothing
    Set NewEmail = Nothing
End Sub

Sub ReplacebyArchive()
   Dim obApp As Object
    Dim NewEmail As MailItem
    Set obApp = Outlook.Application
    Set myOlExp = Application.ActiveExplorer
    myOlExp.Activate
    Set myOlsel = myOlExp.Selection
        c = myOlsel.Item(1).Categories
        c = Replace(c, Mid(c, InStr(c, "ToDo"), 6), "Archive")
        'MsgBox c
        myOlsel.Item(1).Categories = c
        myOlsel.Item(1).Save
     
    Set obApp = Nothing
    Set NewEmail = Nothing
End Sub


Sub cmdRemoveTodoL()

    Dim obApp As Object
    Dim NewEmail As MailItem
    Set obApp = Outlook.Application
    Set myOlExp = Application.ActiveExplorer
    myOlExp.Activate
    Set myOlsel = myOlExp.Selection
        c = myOlsel.Item(1).Categories
      c = Replace(c, Mid(c, InStr(c, "ToDo"), 6), "")
     myOlsel.Item(1).Categories = c
        myOlsel.Item(1).Save
    Set obApp = Nothing
    Set NewEmail = Nothing

End Sub

Sub add_cat()
  Dim objApp As Outlook.Application
    Set oOutlook = GetObject(, "Outlook.Application")
    Set ns = oOutlook.GetNamespace("MAPI")
    
    For Each objCategory In ns.Categories
        UserForm1.lstCategories.AddItem objCategory.Name
    Next
'
'Dim oFolder As MAPIFolder
'Dim oDict As Object
'Dim sStartDate As String
'Dim sEndDate As String
'Dim oItems As Outlook.Items
'Dim sStr As String
'Dim sMsg As String
'
'
' Set myNameSpace = Application.GetNamespace("MAPI")
''Set oFolder = Application.ActiveExplorer.CurrentFolder
'Set oFolder = myNameSpace.GetDefaultFolder(olFolderInbox)
'Set oDict = CreateObject("Scripting.Dictionary")
'Set oDictFull = CreateObject("Scripting.Dictionary")
'
''sStartDate = InputBox("Type the start date (format MM/DD/YYYY)")
'sStartDate = "December2021"
'Set oItems = oFolder.Items.Restrict("[Received] >= '" & sStartDate & "'")
''oItems.SetColumns ("Categories")
'
'For Each aitem In oItems
'    sStr = aitem.Categories
'    cats = Split(sStr, ";")
'    For Each cat In cats
'        cat = Trim(cat)
'        If Not oDict.Exists(cat) Then
'        oDict(cat) = 0
'        End If
'        oDict(cat) = CLng(oDict(cat)) + 1
'    Next cat
'       If Not oDictFull.Exists(sStr) Then
'        oDictFull(sStr) = 0
'        End If
'        oDictFull(sStr) = CLng(oDictFull(sStr)) + 1
'Next aitem
'
'
'sMsg = ""
'For Each aKey In oDict.Keys
'sMsg = sMsg & aKey & ":   " & oDict(aKey) & vbCrLf
'Next
''Debug.Print sMsg
'
'Set oFolder = Nothing
'
'lstCategoriesRecent.Clear
'Dim Arr As Variant
'Arr = SortDict(oDict)
'
'    'Add the sorted keys and items from the array back to the Dictionary
'    For i = LBound(Arr, 1) To UBound(Arr, 1)
'        'Dict.Add Key:=Arr(i, 0), Item:=Arr(i, 1)
'        UserForm1.lstCategoriesRecent.AddItem Arr(i, 0)
'    Next i

    
    UserForm1.Show
        
    Set oOutlook = Nothing
    Set ns = Nothing
End Sub
Public Sub DisplaySenderDetails()
 Dim Explorer As Outlook.Explorer
 Dim CurrentItem As Object
 Dim Sender As Outlook.AddressEntry
 Dim Contact As Outlook.ContactItem
 
 Set Explorer = Application.ActiveExplorer
 
 ' Check whether any item is selected in the current folder.
 If Explorer.Selection.Count Then
 
 ' Get the first selected item.
 Set CurrentItem = Explorer.Selection(1)
 
 ' Check for the type of the selected item as only the
 ' MailItem object has the Sender property.
 If CurrentItem.Class = olMail Then
MsgBox CurrentItem.Sender
 
 ' There is no sender if the item has not been sent yet.
 If Sender Is Nothing Then
 MsgBox "There's no sender for the current email", vbInformation
 Exit Sub
 End If
 
 Set Contact = Sender.GetContact
 
 If Not Contact Is Nothing Then
 ' The sender is stored in the contacts folder,
 ' so the contact item can be displayed.
 Contact.Display
 
 Else
 ' If the contact cannot be found, display the
 ' address entry in the properties dialog box.
 Sender.Details 0
 End If
 End If
 End If
End Sub


Sub apply_cat()

Dim oFolder As MAPIFolder
Dim oDict As Object
Dim oDictNames As Object

Dim sStartDate As String
Dim sEndDate As String
Dim oItems As Outlook.Items

Dim sStr As String

 'Set myNameSpace = Application.GetNamespace("MAPI")
Set oFolder = Application.ActiveExplorer.CurrentFolder
'Set oFolder = myNameSpace.GetDefaultFolder(olFolderInbox)
Set oDict = CreateObject("Scripting.Dictionary")
Set oDictNames = CreateObject("Scripting.Dictionary")

sStartDate = "December2021"
Set oItems = oFolder.Items.Restrict("[Received] >= '" & sStartDate & "'")

 
For Each aitem In oItems
    sStr = aitem.Categories
    cats = Split(sStr, ";")
    For Each cat In cats
        cat = Trim(cat)
        If Not oDict.Exists(cat) Then
        oDict(cat) = 0
        End If
        oDict(cat) = CLng(oDict(cat)) + 1
    Next cat
     If aitem.Class = olMail Then
        naam = Trim(CStr(aitem.Sender))
        If Not oDictNames.Exists(naam) Then
        oDictNames(naam) = 0
        End If
        oDictNames(naam) = CLng(oDictNames(naam)) + 1
End If
Next aitem

 
Set oFolder = Nothing
Dim Arr As Variant
Arr = SortDict(oDict)
Dim ArrNames As Variant
ArrNames = SortDict(oDictNames)


    'Add the sorted keys and items from the array back to the Dictionary
    For i = LBound(Arr, 1) To UBound(Arr, 1)
        'Dict.Add Key:=Arr(i, 0), Item:=Arr(i, 1)
        ApplyCategories.lstCategories.AddItem Arr(i, 0)
    Next i

    For i = LBound(ArrNames, 1) To UBound(ArrNames, 1)
        'Dict.Add Key:=Arr(i, 0), Item:=Arr(i, 1)
        ApplyCategories.lstNames.AddItem ArrNames(i, 0)
    Next i

ApplyCategories.Show

End Sub


Sub outlookweava()
'add category to selected e-mail in inbox
    Set myOlExp = Application.ActiveExplorer
    Set myOlsel = myOlExp.Selection
    c = myOlsel.Item(1).Categories
    myOlsel.Item(1).Categories = c & "; AddedtoWeava"

'Get additional comments for the e-mail
    toelichting = InputBox("Toelichting")

 'Store the selected text in a string and add new lines
    Dim objItem As Object
    Set objItem = Application.ActiveInspector.CurrentItem
    Set objInsp = objItem.GetInspector
    Set objDoc = objInsp.WordEditor
    Dim text As String
    text = ""
    Dim v() As String
    v = Split(objDoc.Windows(1).Selection, Chr(13))
    For i = 0 To UBound(v)
            text = text & v(i) & Chr(13) & Chr(10)
    Next

 'Open excel
    Dim xlApp As Excel.Application
    Dim xlBook As Excel.Workbook
    Dim xlSheet As Object
    Dim fileDoesExist As Boolean
    Dim FileName As String

    ' Create Excel Application, in case its not running yet
    On Error Resume Next
    Set xlApp = GetObject(, "Excel.Application")
        If Err <> 0 Then
            Set xlApp = CreateObject("Excel.Application")
        End If
        
    xlApp.Visible = True
    FileName = "outlook.xlsm"

    ' Check if the workbook is already open
    Set xlBook = xlApp.Workbooks(FileName)
    
    If xlBook Is Nothing Then
        fileDoesExist = Dir("C:\Temp\" & FileName) > ""
        ' Check for existing file
        If fileDoesExist Then
            ' Open Excel file
            Set xlBook = xlApp.Workbooks.Open("C:\Temp\" & FileName)
            Set xlSheet = xlBook.Sheets(1)
        Else
            MsgBox "File not found"
        End If
    End If
    
    ' The table is on the first sheet
    Set xlSheet = xlBook.Sheets(1)
    
    With xlApp
        With xlBook
            ' Get the table and populate with info
            Dim tbl As ListObject
            Set tbl = xlSheet.ListObjects("Tabel1")
            'Set tbl = Range("Tabel1").ListObject
            Set newrow = tbl.ListRows.Add(AlwaysInsert:=True)
            newrow.Range(1, 1).Value = objItem.EntryID
            newrow.Range(1, 2).Value = objItem.Subject
            newrow.Range(1, 3).Value = objItem.Sender
            newrow.Range(1, 4).Value = Format(objItem.SentOn, "MMM d, yyyy")
            newrow.Range(1, 5).Value = text 'objsel
            newrow.Range(1, 6).Value = objsel.Start
            newrow.Range(1, 7).Value = objsel.End
            newrow.Range(1, 8).Value = toelichting
        End With
    End With
    
    
    Set objInsp = Nothing
    Set objDoc = Nothing
    Set objsel = Nothing
    Set xlApp = Nothing
    Set xlSheet = Nothing
    Set xlBook = Nothing
    Set newrow = Nothing
    Set tbl = Nothing
    Set objItem = Nothing
End Sub


Public Sub test()
MsgBox "test"
End Sub
