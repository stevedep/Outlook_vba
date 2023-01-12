VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Organize your email"
   ClientHeight    =   14895
   ClientLeft      =   90
   ClientTop       =   405
   ClientWidth     =   20715
   OleObjectBlob   =   "UserForm1.frx":0000
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub addToMail_Click()
 Dim obApp As Object
    Dim NewEmail As MailItem
    Set obApp = Outlook.Application
    'Set NewEmail = obApp.ActiveInspector.CurrentItem
    'Set NewEmail = obApp.Selection.Item(1)
    Set myOlExp = Application.ActiveExplorer
    myOlExp.Activate
 Set myOlsel = myOlExp.Selection
    'MsgBox myOlsel.Item(1).Subject
    c = myOlsel.Item(1).Categories
    'MsgBox c
    'If you want to set a specific category to the new email manually
    'You can use the following line instead to show the Category dialog
    'NewEmail.ShowCategoriesDialog
  For intCurrentRow = 0 To lstCategories.ListCount - 1
    If lstCategories.Selected(intCurrentRow) Then
        stritems = stritems & lstCategories.Column(0, _
        intCurrentRow) & ";"
    End If
 Next intCurrentRow
    
'Dim individualItem As Object
'MsgBox myOlsel.Count
'For Each individualItem In myOlsel
   myOlsel.Item(1).Categories = c & "; " & stritems
  ' individualItem.Categories = c & "; " & stritems
  'MsgBox individualItem.Subject
'Next
  
   
    'myOlsel.Item(1).Save
    Set obApp = Nothing
    Set NewEmail = Nothing
    UserForm1.Hide
    UserForm1.Show
'    TextBox1.SetFocus
    TextBox1.text = ""
   ' cmdSave.SetFocus
End Sub


Private Sub cmdAddandSaveMultiple_Click()
Call AddSelection("Yes")
 
End Sub

Private Sub cmdAddEmpty_Click()
     Dim obApp As Object
    Dim NewEmail As MailItem
    Set obApp = Outlook.Application
    Set myOlExp = Application.ActiveExplorer
    myOlExp.Activate
   Set myOlsel = myOlExp.Selection
    c = myOlsel.Item(1).Categories
    
   myOlsel.Item(1).Categories = ""
  
    Set obApp = Nothing
    Set NewEmail = Nothing
    UserForm1.Hide
    UserForm1.Show
   TextBox1.text = ""

End Sub


 Sub AddSelection(SaveQuestion As String)

Dim CategoriesString As String
   For ListBoxItem = 0 To lstCategoriesRecent.ListCount - 1
            If lstCategoriesRecent.Selected(ListBoxItem) Then
                CategoriesString = CategoriesString & lstCategoriesRecent.List(ListBoxItem) & "; "
            End If
        Next
    Debug.Print CategoriesString
    
     Dim obApp As Object
    Dim NewEmail As MailItem
    Set obApp = Outlook.Application
    Set myOlExp = Application.ActiveExplorer
    myOlExp.Activate
   Set myOlsel = myOlExp.Selection
    c = myOlsel.Item(1).Categories
    
   myOlsel.Item(1).Categories = c & "; " & CategoriesString
  
    Set obApp = Nothing
    Set NewEmail = Nothing
   
   If SaveQuestion = "Yes" Then
'   MsgBox "save"
    cmdSave_Click
    Else
        UserForm1.Hide
    UserForm1.Show
   TextBox1.text = ""

   End If
   
End Sub

Private Sub cmdAddSelection_Click()
Call AddSelection("No")
End Sub

Private Sub cmdcustomcat_Click()
   Dim obApp As Object
    Dim NewEmail As MailItem
    Set obApp = Outlook.Application
    Set myOlExp = Application.ActiveExplorer
    myOlExp.Activate
    Set myOlsel = myOlExp.Selection
      
    c = myOlsel.Item(1).Categories

 ' For intCurrentRow = 0 To lstCategories.ListCount - 1
  '  If lstCategories.Selected(intCurrentRow) Then
  '      stritems = stritems & lstCategories.Column(0, _
    '    intCurrentRow) & ";"
 '   End If
 'Next intCurrentRow
 cat = InputBox("Set Cats")
  Set objNameSpace = Application.GetNamespace("MAPI")
  
Dim answer As Integer
answer = MsgBox("Add Category", vbQuestion + vbYesNo + vbDefaultButton2, "Category Question")
If answer = vbYes Then
  objNameSpace.Categories.Add (cat)
Else
  'MsgBox "No"
End If
  
   myOlsel.Item(1).Categories = c & "; " & stritems & "; " & cat
        'myOlsel.Item(1).Save
  Set objNameSpace = Nothing
    Set obApp = Nothing
    Set NewEmail = Nothing
    UserForm1.Hide
    UserForm1.Show
    UserForm1.TextBox1.SetFocus
    
    TextBox1.SetFocus
    

  
   TextBox1.text = ""

    
    
End Sub

Private Sub cmdDeselect_Click()
Dim i As Long
For i = 0 To lstCategoriesRecent.ListCount - 1
    lstCategoriesRecent.Selected(i) = False
Next
End Sub

Private Sub cmdDone_Click()
 Dim myNameSpace As Outlook.NameSpace
 Dim myInbox As Outlook.Folder
 Dim myDestFolder As Outlook.Folder
 Dim myItems As Outlook.Items
 Dim myItem As Object
 
 Set myNameSpace = Application.GetNamespace("MAPI")
 Set myInbox = myNameSpace.GetDefaultFolder(olFolderInbox)
Dim obApp As Object
    Dim NewEmail As MailItem

    Set obApp = Outlook.Application
    Set myOlExp = Application.ActiveExplorer
    myOlExp.Activate
    Set myOlsel = myOlExp.Selection
    
    myOlsel.Item(1).Move myInbox.Folders("Done")
    
'    myOlsel.Item(1).Save
    
    Set myInbox = Nothing
    Set myDestFolder = Nothing
    Set myItems = Nothing
    Set myItem = Nothing
    Set myNameSpace = Nothing
    Set myOlExp = Nothing
    Set myOlsel = Nothing
    Set obApp = Nothing
    Set NewEmail = Nothing
'    UserForm1.Hide
 '   UserForm1.Show
 '   UserForm1.TextBox1.SetFocus
 '   UserForm1.TextBox1.text = ""
    UserForm1.Hide
    UserForm1.Show
    UserForm1.TextBox1.SetFocus
    
End Sub

Private Sub KeyHandler_KeyDown(KeyCode As Integer, _
     Shift As Integer)
    Dim intShiftDown As Integer, intAltDown As Integer
    Dim intCtrlDown As Integer
 
If KeyCode = vbKeyReturn Then MsgBox ("enter")
' Use bit masks to determine which key was pressed.
    intShiftDown = (Shift And acShiftMask) > 0
    intAltDown = (Shift And acAltMask) > 0
    intCtrlDown = (Shift And acCtrlMask) > 0
    intEnter = acCtrlMask > 0
    ' Display message telling user which key was pressed.
    If intShiftDown Then MsgBox "You pressed the Shift key."
     If intEnter Then MsgBox "You pressed the enter key."
    If intAltDown Then MsgBox "You pressed the Alt key."
    If intCtrlDown Then MsgBox "You pressed the Ctrl key."
End Sub

Private Sub cmdFilterInbox_Click()
Dim SearchString As String
 For ListBoxItem = 0 To lstCategoriesRecent.ListCount - 1
            If lstCategoriesRecent.Selected(ListBoxItem) Then
                SearchString = SearchString & ", " & lstCategoriesRecent.List(ListBoxItem)
            End If
        Next
With Outlook.ActiveExplorer
    .ClearSearch ' Clear previous search if any
    .Search "Category: (" & SearchString & ")", olSearchScopeAllFolders
    .Display 'Shows search results
End With
End Sub

Private Sub cmdRemoveToDo_Click()
    Dim obApp As Object
    Dim NewEmail As MailItem
    Set obApp = Outlook.Application
    Set myOlExp = Application.ActiveExplorer
    myOlExp.Activate
    Set myOlsel = myOlExp.Selection
        c = myOlsel.Item(1).Categories
        'MsgBox c
        c = Replace(c, ", ToDo", "")
        c = Replace(c, "ToDo,", "")
        'MsgBox c
        myOlsel.Item(1).Categories = c
        myOlsel.Item(1).Save
     
    Set obApp = Nothing
    Set NewEmail = Nothing
    UserForm1.Hide
    UserForm1.Show
    UserForm1.TextBox1.SetFocus
    
    TextBox1.SetFocus
   
End Sub

 

Private Sub cmdRemoveTodoL_Click()

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
    UserForm1.Hide
    UserForm1.Show
    UserForm1.TextBox1.SetFocus

    TextBox1.SetFocus
End Sub

Private Sub cmdReplaceToDoArchive_Click()
   Dim obApp As Object
    Dim NewEmail As MailItem
    Set obApp = Outlook.Application
    Set myOlExp = Application.ActiveExplorer
    myOlExp.Activate
    Set myOlsel = myOlExp.Selection
        c = myOlsel.Item(1).Categories
        c = Replace(c, "ToDo", "Archive")
        myOlsel.Item(1).Categories = c
        myOlsel.Item(1).Save
     
    Set obApp = Nothing
    Set NewEmail = Nothing
    UserForm1.Hide
    UserForm1.Show
    UserForm1.TextBox1.SetFocus
    
    TextBox1.SetFocus
End Sub

Private Sub cmdreplacetodoimportantl2_Click()
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
    UserForm1.Hide
    UserForm1.Show
    UserForm1.TextBox1.SetFocus
    
    TextBox1.SetFocus
End Sub

Private Sub cmdReplaceToDoL1_Click()
    Dim obApp As Object
    Dim NewEmail As MailItem
    Set obApp = Outlook.Application
    Set myOlExp = Application.ActiveExplorer
    myOlExp.Activate
    Set myOlsel = myOlExp.Selection
        c = myOlsel.Item(1).Categories
                  
  For intCurrentRow = 0 To lstCategories.ListCount - 1
    If lstCategories.Selected(intCurrentRow) Then
        stritems = stritems & lstCategories.Column(0, _
        intCurrentRow) & ";"
    End If
 Next intCurrentRow
        
        c = Replace(c, Mid(c, InStr(c, "ToDo"), 6), stritems)
        
        myOlsel.Item(1).Categories = c
     '   myOlsel.Item(1).Save
     
    Set obApp = Nothing
    Set NewEmail = Nothing
    UserForm1.Hide
    UserForm1.Show
    UserForm1.TextBox1.SetFocus
End Sub

Private Sub cmdReplaceToDoL2_Click()
   Dim obApp As Object
    Dim NewEmail As MailItem
    Set obApp = Outlook.Application
    Set myOlExp = Application.ActiveExplorer
    myOlExp.Activate
    Set myOlsel = myOlExp.Selection
        c = myOlsel.Item(1).Categories
        c = Replace(c, Mid(c, InStr(c, "ToDo"), 6), "ToDoL2")
        myOlsel.Item(1).Categories = c
        myOlsel.Item(1).Save
     
    Set obApp = Nothing
    Set NewEmail = Nothing
    UserForm1.Hide
    UserForm1.Show
    UserForm1.TextBox1.SetFocus
    
    TextBox1.SetFocus
End Sub

Private Sub cmdReplToDoImportant_Click()
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
    UserForm1.Hide
    UserForm1.Show
    UserForm1.TextBox1.SetFocus
    
'    TextBox1.SetFocus
End Sub

Private Sub cmdreset_Click()
 Set myOlExp = Application.ActiveExplorer
    myOlExp.Activate
    UserForm1.Hide
    UserForm1.Show
    TextBox1.SetFocus
Set myOlExp = Nothing

End Sub

Private Sub cmdSave_Click()
'MsgBox "start save"
 Dim obApp As Object
    Dim NewEmail As MailItem

    Set obApp = Outlook.Application
    Set myOlExp = Application.ActiveExplorer
    myOlExp.Activate
    Set myOlsel = myOlExp.Selection
    
    myOlsel.Item(1).Save
    Set obApp = Nothing
    Set NewEmail = Nothing
    Set myOlsel = Nothing
    UserForm1.Hide
    UserForm1.Show
    UserForm1.TextBox1.SetFocus
    
End Sub


Private Sub cmdMeeting_Click()
 Dim myNameSpace As Outlook.NameSpace
 Dim myInbox As Outlook.Folder
 Dim myDestFolder As Outlook.Folder
 Dim myItems As Outlook.Items
 Dim myItem As Object
 
 Set myNameSpace = Application.GetNamespace("MAPI")
 Set myInbox = myNameSpace.GetDefaultFolder(olFolderInbox)
 'MsgBox olFolderInbox
Dim obApp As Object
    Dim NewEmail As MailItem

    Set obApp = Outlook.Application
    Set myOlExp = Application.ActiveExplorer
    myOlExp.Activate
    Set myOlsel = myOlExp.Selection
    
    myOlsel.Item(1).Move myInbox.Folders("Meetings")
    
'   myOlsel.Item(1).Save
    Set myOlExp = Nothing
    Set myOlsel = Nothing
    Set obApp = Nothing
    Set NewEmail = Nothing
    Set myNameSpace = Nothing
    Set myInbox = Nothing

    UserForm1.Hide
    UserForm1.Show
    UserForm1.TextBox1.SetFocus
End Sub

Private Sub cmdMultiple_Click()
 Dim obApp As Object
    Dim NewEmail As MailItem
    Set obApp = Outlook.Application
    Set myOlExp = Application.ActiveExplorer
    myOlExp.Activate
 Set myOlsel = myOlExp.Selection
    c = myOlsel.Item(1).Categories

  For intCurrentRow = 0 To lstCategories.ListCount - 1
    If lstCategories.Selected(intCurrentRow) Then
        stritems = stritems & lstCategories.Column(0, _
        intCurrentRow) & ";"
    End If
 Next intCurrentRow
    
Dim individualItem As Object

For Each individualItem In myOlsel
   'myOlsel.Item(1).Categories = c & "; " & stritems
   individualItem.Categories = c & "; " & stritems
    individualItem.Save
Next
    
   
    'myOlsel.Item(1).Save
    Set obApp = Nothing
    Set NewEmail = Nothing
    Set obApp = Nothing
    Set myOlExp = Nothing
    UserForm1.Hide
    UserForm1.Show
    UserForm1.TextBox1.SetFocus
    
End Sub

Private Sub addCategoryAndSave_Click()
 Dim obApp As Object
    Dim NewEmail As MailItem

    Set obApp = Outlook.Application
    Set myOlExp = Application.ActiveExplorer
    myOlExp.Activate
 Set myOlsel = myOlExp.Selection
    c = myOlsel.Item(1).Categories
  For intCurrentRow = 0 To lstCategories.ListCount - 1
    If lstCategories.Selected(intCurrentRow) Then
        stritems = stritems & lstCategories.Column(0, _
        intCurrentRow) & ";"
    End If
 Next intCurrentRow
   
   myOlsel.Item(1).Categories = c & "; " & stritems
        myOlsel.Item(1).Save
    
    Set obApp = Nothing
    Set NewEmail = Nothing
    UserForm1.Hide
    UserForm1.Show
    UserForm1.TextBox1.SetFocus
    TextBox1.text = ""

End Sub

Private Sub CommandButton1_Click()
   Dim obApp As Object
  '  Dim NewEmail As MailItem
    Set obApp = Outlook.Application
    Set myOlExp = Application.ActiveExplorer
    UserForm1.Hide
    myOlExp.Activate
    UserForm1.Show
    UserForm1.TextBox1.SetFocus
    
   ' Set myOlsel = myOlExp.Selection
    '    c = myOlsel.Item(1).Categories
    '    c = Replace(c, Mid(c, InStr(c, "ToDo"), 6), "ImportantL1")
        'MsgBox c
     '   myOlsel.Item(1).Categories = c
    '    myOlsel.Item(1).Save
     
    Set obApp = Nothing
   ' Set NewEmail = Nothing
   ' UserForm1.Hide
  '  UserForm1.Show
  '  UserForm1.TextBox1.SetFocus
    
   ' TextBox1.SetFocus
End Sub


Option Explicit



Private Sub CommandButton2_Click()
Dim oFolder As MAPIFolder
Dim oDict As Object
Dim sStartDate As String
Dim sEndDate As String
Dim oItems As Outlook.Items
Dim sStr As String
Dim sMsg As String
 
 
'On Error Resume Next

 Set myNameSpace = Application.GetNamespace("MAPI")
'Set oFolder = Application.ActiveExplorer.CurrentFolder
Set oFolder = myNameSpace.GetDefaultFolder(olFolderInbox)
Set oDict = CreateObject("Scripting.Dictionary")
Set oDictFull = CreateObject("Scripting.Dictionary")
 
'sStartDate = InputBox("Type the start date (format MM/DD/YYYY)")
sStartDate = "December2021"
Set oItems = oFolder.Items.Restrict("[Received] >= '" & sStartDate & "'")
'oItems.SetColumns ("Categories")
 
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
       If Not oDictFull.Exists(sStr) Then
        oDictFull(sStr) = 0
        End If
        oDictFull(sStr) = CLng(oDictFull(sStr)) + 1
Next aitem
 
 
sMsg = ""
For Each aKey In oDict.Keys
sMsg = sMsg & aKey & ":   " & oDict(aKey) & vbCrLf
Next
'Debug.Print sMsg
 
Set oFolder = Nothing

lstCategoriesRecent.Clear
Dim Arr As Variant
Arr = SortDict(oDict)



    'Add the sorted keys and items from the array back to the Dictionary
    For i = LBound(Arr, 1) To UBound(Arr, 1)
        'Dict.Add Key:=Arr(i, 0), Item:=Arr(i, 1)
        UserForm1.lstCategoriesRecent.AddItem Arr(i, 0)
    Next i
    'Build a list of keys and items from the Dictionary
    'For i = 0 To Dict.Count - 1
    '  UserForm1.lstCategoriesRecent.AddItem Dict(i)
    '  Txt = Txt & Dict.Keys(i) & vbTab & Dict.Items(i) & vbCrLf
   ' Next i
    
   'Display the list in a message box
   'Debug.Print Txt


End Sub


Private Sub CommandButton3_Click()
SortDictionaryByItem
End Sub

Private Sub CommandButton4_Click()

End Sub

Private Sub Image1_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

End Sub

Private Sub Label10_Click()

End Sub

Private Sub Label11_Click()

End Sub

Private Sub Label2_Click()

End Sub

Private Sub Label7_Click()

End Sub

Private Sub Label8_Click()

End Sub

Private Sub lstCategories_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
'MsgBox KeyCode
    'bij enter toevoegen en bewaren cat
    If KeyCode = vbKeyReturn Then
                     addCategoryAndSave_Click
                       myOlExp = Application.ActiveExplorer
                myOlExp.Activate
                Set myOlExp = Nothing
   End If
    'bij ctrl toevoegen
    If KeyCode = 32 Then
                 addToMail_Click
                   myOlExp = Application.ActiveExplorer
                myOlExp.Activate
                Set myOlExp = Nothing
   End If
    If KeyCode = 17 Or KeyCode = 16 Then cmdSave_Click
    If KeyCode > 64 Then
        TextBox1.SetFocus
        TextBox1.text = LCase(ChrW(KeyCode))
    End If
    If KeyCode = 37 Then cmdDone_Click
    If KeyCode = 39 Then
       'MsgBox "meeting"
                    cmdMeeting_Click
    End If
  'TextBox1.SetFocus
  
End Sub

Private Sub TextBox1_Change()
    Set oOutlook = GetObject(, "Outlook.Application")
    Set ns = oOutlook.GetNamespace("MAPI")
    lstCategories.Clear
    For Each objCategory In ns.Categories
         If LCase(objCategory.Name) Like "*" & LCase(TextBox1.text) & "*" Then
             UserForm1.lstCategories.AddItem objCategory.Name
         End If
     Next
    Set oOutlook = Nothing
    Set ns = Nothing
End Sub

Private Sub TextBox1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
Set myOlExp = Application.ActiveExplorer
 'Debug.Print KeyCode
 
 If KeyCode = 13 Then
        lstCategories.SetFocus
        lstCategories.ListIndex = 0
        'SendKeys "{DOWN}", 1
End If
 If KeyCode = 37 Then
    UserForm1.Hide
    myOlExp.Activate
    cmdDone_Click
    UserForm1.Show
 End If
 If KeyCode = 39 Then
    'MsgBox "meeting"
    UserForm1.Hide
    myOlExp.Activate
    cmdMeeting_Click
    UserForm1.Show
 End If
 If KeyCode = 18 Then cmdRemoveToDo_Click
 If KeyCode = 38 Then
 '   myOlExp.Activate
    'SendKeys "{UP}", 1
 End If

 If KeyCode = 16 Then
  'myOlExp.Activate
    lstCategories.SetFocus
    'SendKeys "{DOWN}", 1
End If

If KeyCode = 40 Then
   ' myOlExp.Activate
    'SendKeys "{DOWN}", 1
    lstCategories.SetFocus
 End If

Set myOlExp = Nothing
'TextBox1.SetFocus
End Sub

Private Sub UserForm_Click()

End Sub

