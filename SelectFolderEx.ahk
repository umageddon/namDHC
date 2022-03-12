/*
    Displays a standard dialog that allows the user to select folder(s).
    Parameters:
        Owner / Title:
            The identifier of the window that owns this dialog. This value can be zero.
            An Array with the identifier of the owner window and the title. If the title is an empty string, it is set to the default.
        StartingFolder:
            The path to the directory selected by default. If the directory does not exist, it searches in higher directories.
        CustomPlaces:
            Specify an Array with the custom directories that will be displayed in the left pane. Missing directories will be omitted.
            To specify the location in the list, specify an Array with the directory and its location (0 = Lower, 1 = Upper).
        Options:
            Determines the behavior of the dialog. This parameter must be one or more of the following values.
            0x00000200 (FOS_ALLOWMULTISELECT) = Enables the user to select multiple items in the open dialog.
            0x00040000 (FOS_HIDEPINNEDPLACES) = Hide items shown by default in the view's navigation pane.
            0x02000000  (FOS_DONTADDTORECENT) = Do not add the item being opened or saved to the recent documents list (SHAddToRecentDocs).
            0x10000000  (FOS_FORCESHOWHIDDEN) = Include hidden and system items.
            You can check all available values ​​at https://msdn.microsoft.com/en-us/library/windows/desktop/dn457282(v=vs.85).aspx.
    Return:
        Returns zero if the user canceled the dialog, otherwise returns the path of the selected directory. The directory never ends with "\".
*/

SelectFolderEx(StartingFolder:="", Prompt:="", OwnerHwnd:=0, OkBtnLabel:="", comboList:="", desiredDefault:=1, comboLabel:="", CustomPlaces:="", pickFoldersOnly:=1) {
; ==================================================================================================================================
; Shows a dialog to select a folder.
; Depending on the OS version the function will use either the built-in FileSelectFolder command (XP and previous)
; or the Common Item Dialog (Vista and later).
;
; Parameter:
;     StartingFolder -  the full path of a folder which will be preselected.
;     Prompt         -  a text used as window title (Common Item Dialog) or as text displayed withing the dialog.
;     ----------------  Common Item Dialog only:
;     OwnerHwnd      -  HWND of the Gui which owns the dialog. If you pass a valid HWND the dialog will become modal.
;     BtnLabel       -  a text to be used as caption for the apply button.
;     comboList      -  a string with possible drop-down options, separated by `n [new line]
;     desiredDefault -  the default selected drop-down row
;     comboLabel     -  the drop-down label to display
;     CustomPlaces   -  custom directories that will be displayed in the left pane of the dialog; missing directories will be omitted; a string separated by `n [newline]
;     pickFoldersOnly - boolean option [0, 1]
;
;  Return values:
;     On success the function returns an object with the full path of the selected/file folder
;     and combobox selected [if any]; otherwise it returns an empty string.
;
; MSDN:
;     Common Item Dialog -> msdn.microsoft.com/en-us/library/bb776913%28v=vs.85%29.aspx
;     IFileDialog        -> msdn.microsoft.com/en-us/library/bb775966%28v=vs.85%29.aspx
;     IShellItem         -> msdn.microsoft.com/en-us/library/bb761140%28v=vs.85%29.aspx
; ==================================================================================================================================
; Source https://www.autohotkey.com/boards/viewtopic.php?f=6&t=18939
; by «just me»
; modified by Marius Șucan on: vendredi 8 mai 2020
; to allow ComboBox and CustomPlaces
;
; options flags
; FOS_OVERWRITEPROMPT  = 0x2,
; FOS_STRICTFILETYPES  = 0x4,
; FOS_NOCHANGEDIR  = 0x8,
; FOS_PICKFOLDERS  = 0x20,
; FOS_FORCEFILESYSTEM  = 0x40,
; FOS_ALLNONSTORAGEITEMS  = 0x80,
; FOS_NOVALIDATE  = 0x100,
; FOS_ALLOWMULTISELECT  = 0x200,
; FOS_PATHMUSTEXIST  = 0x800,
; FOS_FILEMUSTEXIST  = 0x1000,
; FOS_CREATEPROMPT  = 0x2000,
; FOS_SHAREAWARE  = 0x4000,
; FOS_NOREADONLYRETURN  = 0x8000,
; FOS_NOTESTFILECREATE  = 0x10000,
; FOS_HIDEMRUPLACES  = 0x20000,
; FOS_HIDEPINNEDPLACES  = 0x40000,
; FOS_NODEREFERENCELINKS  = 0x100000,
; FOS_OKBUTTONNEEDSINTERACTION  = 0x200000,
; FOS_DONTADDTORECENT  = 0x2000000,
; FOS_FORCESHOWHIDDEN  = 0x10000000,
; FOS_DEFAULTNOMINIMODE  = 0x20000000,
; FOS_FORCEPREVIEWPANEON  = 0x40000000,
; FOS_SUPPORTSTREAMABLEITEMS  = 0x80000000

; IFileDialog vtable offsets
; 0   QueryInterface
; 1   AddRef 
; 2   Release 
; 3   Show 
; 4   SetFileTypes 
; 5   SetFileTypeIndex 
; 6   GetFileTypeIndex 
; 7   Advise 
; 8   Unadvise 
; 9   SetOptions 
; 10  GetOptions 
; 11  SetDefaultFolder 
; 12  SetFolder 
; 13  GetFolder 
; 14  GetCurrentSelection 
; 15  SetFileName 
; 16  GetFileName 
; 17  SetTitle 
; 18  SetOkButtonLabel 
; 19  SetFileNameLabel 
; 20  GetResult 
; 21  AddPlace 
; 22  SetDefaultExtension 
; 23  Close 
; 24  SetClientGuid 
; 25  ClearClientData 
; 26  SetFilter

   Static OsVersion := DllCall("GetVersion", "UChar")
        , IID_IShellItem := 0
        , InitIID := VarSetCapacity(IID_IShellItem, 16, 0)
                  & DllCall("Ole32.dll\IIDFromString", "WStr", "{43826d1e-e718-42ee-bc55-a1e261c37bfe}", "Ptr", &IID_IShellItem)
        , Show := A_PtrSize * 3
        , SetOptions := A_PtrSize * 9
        , SetFolder := A_PtrSize * 12
        , SetTitle := A_PtrSize * 17
        , SetOkButtonLabel := A_PtrSize * 18
        , GetResult := A_PtrSize * 20
        , ComDlgObj := {COMDLG_FILTERSPEC: ""}

   SelectedFolder := ""
   If (OsVersion<6)
   {
      ; IFileDialog requires Win Vista+, so revert to FileSelectFolder
      FileSelectFolder, SelectedFolder, *%StartingFolder%, 3, %Prompt%
      Return SelectedFolder
   }

   OwnerHwnd := DllCall("IsWindow", "Ptr", OwnerHwnd, "UInt") ? OwnerHwnd : 0
   If !(FileDialog := ComObjCreate("{DC1C5A9C-E88A-4dde-A5A1-60F82A20AEF7}", "{42f85136-db7e-439c-85f1-e4075d135fc8}"))
      Return ""

   VTBL := NumGet(FileDialog + 0, "UPtr")
   dialogOptions := 0x8 | 0x800  ;  FOS_NOCHANGEDIR | FOS_PATHMUSTEXIST
   dialogOptions |= (pickFoldersOnly=1) ? 0x20 : 0x1000    ; FOS_PICKFOLDERS : FOS_FILEMUSTEXIST

   DllCall(NumGet(VTBL + SetOptions, "UPtr"), "Ptr", FileDialog, "UInt", dialogOptions, "UInt")
   If StartingFolder
   {
      If !DllCall("Shell32.dll\SHCreateItemFromParsingName", "WStr", StartingFolder, "Ptr", 0, "Ptr", &IID_IShellItem, "PtrP", FolderItem)
         DllCall(NumGet(VTBL + SetFolder, "UPtr"), "Ptr", FileDialog, "Ptr", FolderItem, "UInt")
   }

   If Prompt
      DllCall(NumGet(VTBL + SetTitle, "UPtr"), "Ptr", FileDialog, "WStr", Prompt, "UInt")
   If OkBtnLabel
      DllCall(NumGet(VTBL + SetOkButtonLabel, "UPtr"), "Ptr", FileDialog, "WStr", OkBtnLabel, "UInt")

   If CustomPlaces
   {
      Loop, Parse, CustomPlaces, `n
      {
          If InStr(FileExist(A_LoopField), "D")
          {
             foo := 1, Directory := A_LoopField
             DllCall("Shell32.dll\SHParseDisplayName", "UPtr", &Directory, "Ptr", 0, "UPtrP", PIDL, "UInt", 0, "UInt", 0)
             DllCall("Shell32.dll\SHCreateShellItem", "Ptr", 0, "Ptr", 0, "UPtr", PIDL, "UPtrP", IShellItem)
             ObjRawSet(ComDlgObj, IShellItem, PIDL)
             ; IFileDialog::AddPlace method
             ; https://msdn.microsoft.com/en-us/library/windows/desktop/bb775946(v=vs.85).aspx
             DllCall(NumGet(NumGet(FileDialog+0)+21*A_PtrSize), "UPtr", FileDialog, "UPtr", IShellItem, "UInt", foo)
          }
      }
   }

   If (comboList && comboLabel)
   {
      Try If ((FileDialogCustomize := ComObjQuery(FileDialog, "{e6fdd21a-163f-4975-9c8c-a69f1ba37034}")))
      {
         groupId := 616 ; arbitrarily chosen IDs
         comboboxId := 93270
         DllCall(NumGet(NumGet(FileDialogCustomize+0)+26*A_PtrSize), "Ptr", FileDialogCustomize, "UInt", groupId, "WStr", comboLabel) ; IFileDialogCustomize::StartVisualGroup
         DllCall(NumGet(NumGet(FileDialogCustomize+0)+6*A_PtrSize), "Ptr", FileDialogCustomize, "UInt", comboboxId) ; IFileDialogCustomize::AddComboBox
         ; DllCall(NumGet(NumGet(FileDialogCustomize+0)+19*A_PtrSize), "Ptr", FileDialogCustomize, "UInt", comboboxId, "UInt", itemOneId, "WStr", "Current folder") ; IFileDialogCustomize::AddControlItem
         
         entriesArray := []
         Loop, Parse, comboList,`n
         {
             Random, varA, 2, 900
             Random, varB, 2, 900
             thisID := varA varB
             If A_LoopField
             {
                If (A_Index=desiredDefault)
                   desiredIDdefault := thisID

                entriesArray[thisId] := A_LoopField
                DllCall(NumGet(NumGet(FileDialogCustomize+0)+19*A_PtrSize), "Ptr", FileDialogCustomize, "UInt", comboboxId, "UInt", thisID, "WStr", A_LoopField)
             }
         }

         DllCall(NumGet(NumGet(FileDialogCustomize+0)+25*A_PtrSize), "Ptr", FileDialogCustomize, "UInt", comboboxId, "UInt", desiredIDdefault) ; IFileDialogCustomize::SetSelectedControlItem
         DllCall(NumGet(NumGet(FileDialogCustomize+0)+27*A_PtrSize), "Ptr", FileDialogCustomize) ; IFileDialogCustomize::EndVisualGroup
      }

   }

   If !DllCall(NumGet(VTBL + Show, "UPtr"), "Ptr", FileDialog, "Ptr", OwnerHwnd, "UInt")
   {
      If !DllCall(NumGet(VTBL + GetResult, "UPtr"), "Ptr", FileDialog, "PtrP", ShellItem, "UInt")
      {
         GetDisplayName := NumGet(NumGet(ShellItem + 0, "UPtr"), A_PtrSize * 5, "UPtr")
         If !DllCall(GetDisplayName, "Ptr", ShellItem, "UInt", 0x80028000, "PtrP", StrPtr) ; SIGDN_DESKTOPABSOLUTEPARSING
            SelectedFolder := StrGet(StrPtr, "UTF-16"), DllCall("Ole32.dll\CoTaskMemFree", "Ptr", StrPtr)

         ObjRelease(ShellItem)
         if (FileDialogCustomize)
         {
            if (DllCall(NumGet(NumGet(FileDialogCustomize+0)+24*A_PtrSize), "Ptr", FileDialogCustomize, "UInt", comboboxId, "UInt*", selectedItemId) == 0)
            { ; IFileDialogCustomize::GetSelectedControlItem
               if selectedItemId
                  thisComboSelected := entriesArray[selectedItemId]
            }   
         }
      }
   }
   If (FolderItem)
      ObjRelease(FolderItem)

   if (FileDialogCustomize)
      ObjRelease(FileDialogCustomize)

   ObjRelease(FileDialog)
   r := []
   r.SelectedDir := SelectedFolder
   r.SelectedCombo := thisComboSelected
   Return r
}