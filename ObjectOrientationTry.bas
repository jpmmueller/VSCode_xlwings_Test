Attribute VB_Name = "ObjectOrientationTry"
'@Folder("Child")
Option Explicit

Private Sub HowToCreateObjects()

Dim fso As FileSystemObject
Dim FolderPlace As String
Dim Environment As String
Dim FolderName As String
Dim ThePath As String

Environment = Environ$("UserProfile")
FolderPlace = "\Desktop\"
FolderName = "Test"
ThePath = Environment & FolderPlace & FolderName

Set fso = New FileSystemObject
msgbox "Ich wurde mit VS Code erstellt! und bearbeitet"
'fso.CreateFolder ThePath
Set fso = Nothing

End Sub

Private Sub LetsRunIt()

HowToCreateObjects

End Sub

