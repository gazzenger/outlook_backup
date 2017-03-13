VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisOutlookSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private WithEvents Items As Outlook.Items
Attribute Items.VB_VarHelpID = -1

Private Sub Application_Startup()
    Set Items = Session.GetDefaultFolder(olFolderSentMail).Items
 '  ExpandAllFolders
End Sub



Private Sub Items_ItemAdd(ByVal Item As Object)

If Item.Class = olMail Then
    'repeat for each account
    'If (Item.SendUsingAccount = "2H Exchange 2010") Or (Item.SendUsingAccount = "Gary.Namestnik@2hoffshore.com") Then

    ' Get the GetFolderPath function from http://slipstick.me/getfolderpath
    ' Set MovePst = GetFolderPath("Public Folders - Gary.Namestnik@2hoffshore.com\Favorites\301 Buoyancy Tank")

        If MsgBox("Would you like to shift this message to a public folder?", vbYesNo, "Confirm") = vbYes Then
            Set MovePst = FolderPick
            MsgBox MovePst
            If Not (MovePst Is Nothing) Then
                Item.UnRead = False
                Item.Move MovePst
            End If
        End If
    'End If
End If
End Sub

Function GetFolderPath(ByVal FolderPath As String) As Outlook.Folder
    Dim oFolder As Outlook.Folder
    Dim FoldersArray As Variant
    Dim i As Integer

    On Error GoTo GetFolderPath_Error
    If Left(FolderPath, 2) = "\\" Then
        FolderPath = Right(FolderPath, Len(FolderPath) - 2)
    End If
    'Convert folderpath to array
    FoldersArray = Split(FolderPath, "\")
    Set oFolder = Application.Session.Folders.Item(FoldersArray(0))
    If Not oFolder Is Nothing Then
        For i = 1 To UBound(FoldersArray, 1)
            Dim SubFolders As Outlook.Folders
            Set SubFolders = oFolder.Folders
            Set oFolder = SubFolders.Item(FoldersArray(i))
            If oFolder Is Nothing Then
                Set GetFolderPath = Nothing
            End If
        Next
    End If
    'Return the oFolder
    Set GetFolderPath = oFolder
    Exit Function

GetFolderPath_Error:
    Set GetFolderPath = Nothing
    Exit Function
End Function

    Function FolderPick() As Outlook.Folder
    Dim objNS As NameSpace
    Dim objFolder As Folder

    Set objNS = Application.GetNamespace("MAPI")
    Set objFolder = objNS.PickFolder



    If TypeName(objFolder) <> "Nothing" Then

        Debug.Print vbCr & " objFolder: " & objFolder

    Else

        Debug.Print vbCr & "Cancel"

    End If



Set FolderPick = objFolder
'    Set objFolder = Nothing
'    Set objNS = Nothing
    End Function

Private Sub ExpandAllFolders()
  On Error Resume Next
  Dim Ns As Outlook.NameSpace
  Dim Folders As Outlook.Folders
  Dim CurrF As Outlook.MAPIFolder
  Dim F As Outlook.MAPIFolder
  Dim ExpandDefaultStoreOnly As Boolean

  ExpandDefaultStoreOnly = False

  Set Ns = Application.GetNamespace("Mapi")
  Set CurrF = Application.ActiveExplorer.CurrentFolder

  If ExpandDefaultStoreOnly = True Then
    Set F = Ns.GetDefaultFolder(olFolderInbox)
    Set F = F.Parent
    Set Folders = F.Folders
    LoopFolders Folders, False

  Else
    LoopFolders Ns.Folders, True
  End If

  DoEvents
  Set Application.ActiveExplorer.CurrentFolder = CurrF
End Sub

Private Sub LoopFolders(Folders As Outlook.Folders, _
  ByVal bRecursive As Boolean _
)
  Dim F As Outlook.MAPIFolder

  For Each F In Folders
    Set Application.ActiveExplorer.CurrentFolder = F
    DoEvents

    If bRecursive Then
      If F.Name = "Public Folders - Gary.Namestnik@2hoffshore.com" Then
        LoopFolders F.Folders, True
      ElseIf F.Folders.Count Then
        LoopFolders F.Folders, False
      End If
    End If
  Next
End Sub

