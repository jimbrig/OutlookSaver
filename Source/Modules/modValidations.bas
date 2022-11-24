Attribute VB_Name = "modValidations"
Option Explicit

Public Function ValidatePath(Path As String) As Boolean

  Dim ValidPath As Boolean
  Dim PathParent As String
  Dim FSO As New FileSystemObject
  
  Debug.Print "Validating Path: " & Path
  
  PathParent = FSO.GetParentFolderName(Path)
  
  Debug.Print "Parent Folder: " & PathParent
  
  If FSO.FolderExists(PathParent) = False Then
    ValidPath = False
    Debug.Print "Path is Invalid: Parent Folder not found."
    ValidatePath = ValidPath
    Exit Function
  End If
  
  If FSO.FolderExists(Path) = False Then
    Debug.Print "Path provided is valid, but folder does not exist. Create the folder if necessary."
    ValidPath = False
    ValidatePath = ValidPath
    Exit Function
  End If
  
  Set FSO = Nothing
  
  ValidPath = True
  Debug.Print "Path Validation Successful."
  ValidatePath = ValidPath
  Exit Function

End Function
