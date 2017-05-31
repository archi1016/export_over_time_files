Attribute VB_Name = "func"
Option Explicit

Public Declare Function SystemTimeToFileTime Lib "Kernel32" _
    (lpSystemTime As SYSTEMTIME, _
     lpFileTime As FILETIME) As Long

Sub Main()
    Call InitCommonControls

    MainForm.Show
End Sub
