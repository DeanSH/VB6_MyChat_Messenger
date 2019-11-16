VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MY! Chat - Run Time Files Installer"
   ClientHeight    =   4575
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5295
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   5295
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      BackColor       =   &H00000000&
      Caption         =   "Skip Any Moving File Error's...."
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   6120
      TabIndex        =   2
      Top             =   1680
      Value           =   1  'Checked
      Width           =   5055
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H000000FF&
      Height          =   4350
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5055
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000C000&
      Caption         =   "Move && Register All R4 Needed Files!!"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1200
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Is64 As Boolean
Private LastFile As String

Private Sub Command1_Click()
List1.AddItem "----------------------------------------------------------------------------------------------------"
List1.AddItem "Moving All & Replacing Any Existing Files..."
List1.AddItem "----------------------------------------------------------------------------------------------------"
LastFile = ""
'''''Now Move The Files
Dim Files() As String
Files = Split("AniGIF.ocx,comctl32.ocx,COMDLG32.OCX,mscomctl.ocx,MSWINSCK.OCX,quartz.dll,picclp32.ocx,richtx32.ocx,tsd32.dll,tssoft32.acm,yacscom.dll,yacsui.dll,wmp.dll", ",")
Dim I As Integer
For I = 0 To UBound(Files)
If CopyTheFile(Files(I)) = False Then EncounteredError: Exit Sub
DoEvents
Next I
DoEvents
'''''Now Register The Moved Files
List1.AddItem "----------------------------------------------------------------------------------------------------"
List1.AddItem "Done Moving All Files Successfully!"
List1.AddItem "Attempting To Register All Files..."
List1.AddItem "----------------------------------------------------------------------------------------------------"
List1.ListIndex = List1.ListCount - 1
On Error Resume Next
For I = 0 To UBound(Files)
Shell "regsvr32 /s C:\windows\system32\" & Files(I)
DoEvents
List1.AddItem "Registered File >> " & Files(I)
List1.ListIndex = List1.ListCount - 1
If Is64 = True Then
Shell "regsvr32 /s C:\windows\SysWOW64\" & Files(I)
DoEvents
End If
Next I
DoEvents
List1.AddItem "----------------------------------------------------------------------------------------------------"
List1.AddItem "Done Registering All Files!"
List1.AddItem "Congradulations, Finished!!"
List1.AddItem "----------------------------------------------------------------------------------------------------"
List1.ListIndex = List1.ListCount - 1
If Is64 = True Then
MsgBox "All Files Should Now Be Successfully Registered Correctly For 64 Bits!" & vbCrLf & vbCrLf & "(If They Are Not.. Then You Probably Use Vista/Win7 And Did Not Run This Tool As Administrator!!)", vbExclamation, "Finished!!"
Else
MsgBox "All Files Should Now Be Successfully Registered Correctly For 32 Bits!" & vbCrLf & vbCrLf & "(If They Are Not.. Then You Probably Use Vista/Win7 And Did Not Run This Tool As Administrator!!)", vbExclamation, "Finished!!"
End If
DoEvents
End
End Sub

Public Sub EncounteredError()
On Error Resume Next
List1.AddItem "----------------------------------------------------------------------------------------------------"
List1.AddItem "ERROR!!! Failed To Move >> " & LastFile
List1.AddItem "ERROR!!! Close Any Running Programs And Try Again!"
List1.AddItem "----------------------------------------------------------------------------------------------------"
List1.ListIndex = List1.ListCount - 1
MsgBox "Attention There Was A Error While Moving..." & LastFile & vbCrLf & "It Might Be In Use By Another Program You Have Running!!" & vbCrLf & vbCrLf & "(1) Make Sure You Extracted This Program With All The Needed Files, All In The Same Folder Together!!" & vbCrLf & "(2) Make Sure You Close Any Running Programs, Yahoo Tools, Chat Clients, Skype, Yahoo, MSN, Web Browsers, Shut Them All Down, Or Reboot PC.. And Run This Before Open Anything After Startup!!" & vbCrLf & "(3) If Using Vista/Win7/Win8 You Must Run This Tool With Admin Previledges (Right Click EXE, Select Run As Administrator)" & vbCrLf & "(4) If Your Anti Virus, Blocks Or Makes Any Problems, Then Disable It While Register Files, And While Running Software!", vbCritical, "Error While Moving Files!!"
End Sub

Public Function CopyTheFile(fileName As String) As Boolean
On Error GoTo Error

Dim sourceFile As String
Dim destFile As String

LastFile = fileName

sourceFile = App.Path & "\" & fileName
destFile = "C:\windows\system32\" & fileName
FileCopy sourceFile, destFile
DoEvents

If Is64 = True Then
destFile = "C:\windows\SysWOW64\" & fileName
FileCopy sourceFile, destFile
DoEvents
End If

List1.AddItem "GOOD || File Moved >> " & fileName
List1.ListIndex = List1.ListCount - 1
CopyTheFile = True
Exit Function

Error:
List1.AddItem "FAIL || File In Use >> " & fileName
List1.ListIndex = List1.ListCount - 1
If Check1.Value = 1 Then
CopyTheFile = True
Else
CopyTheFile = False
End If
End Function

Private Sub Form_Load()
On Error Resume Next
Me.Show
DoEvents
'''''1st Check For 64 Bits By Checking Existance Of the SysWOW64 Folder!!
List1.Clear
Dim fs
Set fs = CreateObject("Scripting.FileSystemObject")
If fs.FolderExists("C:\windows\SysWOW64") Then
Is64 = True
List1.AddItem "64 Bits Windows Detected..."
List1.AddItem "Ready To Move & Register Files For 64 Bits!!"
Else
Is64 = False
List1.AddItem "32 Bits Windows Detected..."
List1.AddItem "Ready To Move & Register Files For 32 Bits!!"
End If
DoEvents
MsgBox "Important Before You Start..." & vbCrLf & vbCrLf & "(1) Make Sure You This Open's In Same Folder Where MY!Chat Is With Its Needed Files!!" & vbCrLf & "(2) Make Sure You Close Any Running Programs, Skype, Yahoo, MSN, Web Browsers, Shut Them All Down, Or Reboot PC.. And Run This Before Open Anything After Startup!!" & vbCrLf & "(3) If Using Vista/Win7/Win8 You Must Run This Tool With Admin Previledges (Right Click EXE, Select Run As Administrator)" & vbCrLf & "(4) If Your Anti Virus, Blocks Or Makes Any Problems, Then Disable It While Register Files!", vbInformation, "Important To Read!!"
DoEvents
Command1_Click
End Sub

