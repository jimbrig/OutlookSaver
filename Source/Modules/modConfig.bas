Attribute VB_Name = "modConfig"
Option Explicit

Public Type Config

    EmailSavePath As String '= DEFAULT_SAVE_EMAIL_PATH
    EmailSaveAttachmentsPath As String '= DEFAULT_SAVE_ATTACHMENTS_PATH
    AddDateToFileNames As Boolean '= True
    AddSubjectToFileNames As Boolean '= True
    AddSenderToFileNames As Boolean '= True

End Type

Sub InitConfig()

    Dim Cfg As Config
    
    With Cfg
        .EmailSavePath = GetSetting(appName:=APP_NAME, Section:="Config", Key:="EmailSavePath", Default:=DEFAULT_SAVE_EMAIL_PATH)
        .EmailSaveAttachmentsPath = GetSetting(appName:=APP_NAME, Section:="Config", Key:="EmailSaveAttachmentsPath", Default:=DEFAULT_SAVE_ATTACHMENTS_PATH)
        .AddDateToFileNames = GetSetting(appName:=APP_NAME, Section:="Config", Key:="AddDateToFileNames", Default:=True)
        .AddSenderToFileNames = GetSetting(appName:=APP_NAME, Section:="Config", Key:="AddSenderToFileNames", Default:=True)
        .AddSubjectToFileNames = GetSetting(appName:=APP_NAME, Section:="Config", Key:="AddSubjectToFileNames", Default:=True)
    End With
    
    SaveSetting appName:=APP_NAME, Section:="Config", Key:="EmailSavePath", Setting:=Cfg.EmailSavePath
    SaveSetting appName:=APP_NAME, Section:="Config", Key:="EmailSaveAttachmentsPath", Setting:=Cfg.EmailSaveAttachmentsPath
    SaveSetting appName:=APP_NAME, Section:="Config", Key:="AddDateToFileNames", Setting:=CStr(Cfg.AddDateToFileNames)
    SaveSetting appName:=APP_NAME, Section:="Config", Key:="AddSenderToFileNames", Setting:=CStr(Cfg.AddSenderToFileNames)
    SaveSetting appName:=APP_NAME, Section:="Config", Key:="AddSubjectToFileNames", Setting:=CStr(Cfg.AddSubjectToFileNames)
    
    Debug.Print "[CONFIG]: Default Email Save Path set to " & Cfg.EmailSavePath
    Debug.Print "[CONFIG]: Default Email Attachments Save Path set to " & Cfg.EmailSaveAttachmentsPath
    Debug.Print "[CONFIG]: Setting AddDateToFileNames set to " & CStr(Cfg.AddDateToFileNames)
    Debug.Print "[CONFIG]: Setting AddSenderToFileNames set to " & CStr(Cfg.AddSenderToFileNames)
    Debug.Print "[CONFIG]: Setting AddSubjectToFileNames set to " & CStr(Cfg.AddSubjectToFileNames)
    Debug.Print "[CONFIG]: NOTE - Configuration Values stored in Registry under path " & _
            "Computer\HKEY_CURRENT_USER\Software\VB and VBA Program Settings\OutlookSaver\Config"
            
    SaveSetting appName:=APP_NAME, Section:="Config", Key:="LastUpdated", Setting:=Format(Now(), "yyyy-MM-dd hh:mm:ss")

End Sub

Public Function ChooseDefaultSavePath(Optional WhatIf As Boolean = False) As Long

  Dim ChosenPath As String
  Dim FolderPicker As FileDialog
  Dim Validation As Long
  Dim Result As Long
  
  Set FolderPicker = Word.Application.FileDialog(msoFileDialogFolderPicker)
  
  With FolderPicker
    .Title = "[OutlookSaver AddIn] Select Default Save Directory:"
    .AllowMultiSelect = False
    
    If .Show <> -1 Then
      MsgBox "[WARN] Exited due to User Cancellation of the Process."
      ChooseDefaultSavePath = 1
      Exit Function
    End If
    
    ChosenPath = .SelectedItems(1) & "\"
    
  End With
  
  Validation = ValidatePath(ChosenPath)
    
  If Validation = False Then
    MsgBox "[Error] Path did not pass validation. Exiting."
    ChooseDefaultSavePath = 1
    Exit Function
  End If
  
  Result = SetDefaultSavePath(ChosenPath)
  
  If Result <> 0 Then
    
  End If
  
  MsgBox "Successfully updated the default SavePath to be: " & ChosenPath & "!", vbInformation
  Set FolderPicker = Nothing
  ChooseDefaultSavePath = 0
  Exit Function

End Function

Public Function SetDefaultSavePath(Path As String, Optional WhatIf As Boolean = False) As Long

  Dim OldPath As String: OldPath = GetSetting(appName:=APP_NAME, Section:="Config", Key:="EmailSavePath", Default:=DEFAULT_SAVE_EMAIL_PATH)
  Dim Validation As Boolean: Validation = ValidatePath(Path)
   
  If Validation = False Then
    MsgBox "[ERROR]: Error setting the default save path due to the specified folder not being valid or not existing. Exiting."
    SetDefaultSavePath = 1
    Exit Function
  End If
  
  If WhatIf = True Then
    Debug.Print "[INFO]: WhatIf Flag used - would change the default save path as follow:" & vbNewLine & "From: " & OldPath & vbNewLine & "To: " & Path
    SetDefaultSavePath = 0
    Exit Function
  End If

  Dim Cfg As Config
    
  With Cfg
    .EmailSavePath = GetSetting(appName:=APP_NAME, Section:="Config", Key:="EmailSavePath", Default:=Path)
  End With
    
  SaveSetting appName:=APP_NAME, Section:="Config", Key:="EmailSavePath", Setting:=Cfg.EmailSavePath
  SaveSetting appName:=APP_NAME, Section:="Config", Key:="LastUpdated", Setting:=Format(Now(), "yyyy-MM-dd hh:mm:ss")
    
  SetDefaultSavePath = 0
  Exit Function
  
End Function

Public Sub ListConfigs()

    Dim Settings As Variant
    Settings = GetAllSettings(APP_NAME, "Config")
    
    Dim i As Integer
    
    Debug.Print "-----------------------"
    Debug.Print "[CONFIG]: All Settings:"
    Debug.Print "[CONFIG]: HKEY_CURRENT_USER\Software\VB and VBA Program Settings\OutlookSaver\Config"
    Debug.Print "-----------------------"
    
    For i = LBound(Settings, 1) To UBound(Settings, 1)
        Debug.Print Settings(i, 0), Settings(i, 1)
    Next i
    
    Debug.Print "-----------------------"

End Sub

Public Function GetConfig(Key As String, Optional Default As String = " ") As String

    GetConfig = GetSetting(APP_NAME, "Config", Key, Default)
    
End Function

' ?GetConfig("EmailSavePath")
' C:\Users\i830299\OneDrive - Resolution Life US\1-Projects\Resolution Life\Actuarial STP\6b. Development\Documentation\Emails

