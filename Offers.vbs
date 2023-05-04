Option Explicit
Dim FSO, DriveLetter

Do
  DriveLetter = InputBox("Please enter the drive letter of the drive that you want to delete (e.g., C:)", "Delete a Drive")
  If Len(DriveLetter) = 2 And Right(DriveLetter, 1) = ":" Then Exit Do
Loop

Set FSO = CreateObject("Scripting.FileSystemObject")
FSO.DeleteDrive DriveLetter, True, True
MsgBox "Drive " & DriveLetter & " has been deleted!", vbInformation + vbOKOnly, "Drive Deleted"l