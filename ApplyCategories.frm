VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ApplyCategories 
   Caption         =   "UserForm2"
   ClientHeight    =   10950
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12885
   OleObjectBlob   =   "ApplyCategories.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ApplyCategories"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdFilter_Click()
Dim SearchString As String
 For ListBoxItem = 0 To lstCategories.ListCount - 1
            If lstCategories.Selected(ListBoxItem) Then
                SearchString = SearchString & ", " & lstCategories.List(ListBoxItem)
            End If
        Next

With Outlook.ActiveExplorer
  '  .ClearSearch ' Clear previous search if any
    .Search "Category: (" & SearchString & ")", olSearchScopeCurrentFolder
    .Display 'Shows search results
End With

End Sub

Private Sub cmdFilterName_Click()
Dim SearchString As String
 For ListBoxItem = 0 To lstNames.ListCount - 1
            If lstNames.Selected(ListBoxItem) Then
                SearchString = SearchString & ", " & lstNames.List(ListBoxItem)
            End If
        Next

With Outlook.ActiveExplorer
  '  .ClearSearch ' Clear previous search if any
    .Search "From: (" & SearchString & ")", olSearchScopeCurrentFolder
    .Display 'Shows search results
End With
End Sub
