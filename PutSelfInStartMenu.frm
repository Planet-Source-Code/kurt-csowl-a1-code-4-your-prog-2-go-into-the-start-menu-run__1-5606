VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

Dim strDir As String
Dim intDir As Integer
Dim strpath As String


intDir = Len(CurDir())  'Gets length of Current directory
'Gets last character from CurDir() and checks if it's a \
strDir = Mid(CurDir(), intDir)
If strDir = "\" Then
    strpath = CurDir() & App.EXEName & ".exe"
    'If is in main drive like C:\ or D:\ it simply
    'puts the file name, "C:\Blah.exe"
Else
    strpath = CurDir() & "\" & App.EXEName & ".exe"
    'If CurDir() returns no \ then its in a folder
    'and will necessitate a \ inserted so that it looks like
    'C:\Folder\Blah.exe and NOT like C:\FolderBlah.exe
End If
On Error GoTo Death
'Error statement allows this code to run if it is already
'in the Start Menu

FileCopy strpath, _
"C:\WINDOWS\Start Menu\Programs\StartUp\" _
& App.EXEName & ".exe"
Death:
Resume Next
On Error GoTo Ender
FileCopy strpath, "C:\Winnt\Start Menu\Programs\StartUp\" & App.EXEName & ".exe"
Ender:
Exit Sub
Resume Next
End Sub

