VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Windows Informations"
   ClientHeight    =   4935
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7320
   Icon            =   "Main Form.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   7320
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Default         =   -1  'True
      Height          =   495
      Left            =   3053
      TabIndex        =   1
      Top             =   4320
      Width           =   1215
   End
   Begin VB.TextBox txtMain 
      Height          =   4095
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   7095
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'########## Code written by Sebastian Mares (MaresWEB) ##########

'Force variable declaration
Option Explicit

Private Sub cmdExit_Click()

    Unload frmMain

End Sub

'Displays the informations found about the system in a TextBox
Private Sub Form_Load()

    On Error Resume Next
      txtMain.Text = "Windows Version:" & vbTab & vbTab & GetWindowsVersion
      txtMain.Text = txtMain.Text & vbCrLf & "Code Name:" & vbTab & vbTab & GetWindowsCodeName
      txtMain.Text = txtMain.Text & vbCrLf & vbCrLf & "Platform ID:" & vbTab & vbTab & GetPlatformNumber
      txtMain.Text = txtMain.Text & vbCrLf & "Platform Type:" & vbTab & vbTab & GetPlatformType & " Based"
      txtMain.Text = txtMain.Text & vbCrLf & vbCrLf & "Major Number:" & vbTab & vbTab & GetMajorNumber
      txtMain.Text = txtMain.Text & vbCrLf & "Minor Number:" & vbTab & vbTab & GetMinorNumber
      txtMain.Text = txtMain.Text & vbCrLf & "Build Number:" & vbTab & vbTab & GetBuildNumber
      txtMain.Text = txtMain.Text & vbCrLf & "Service Pack:" & vbTab & vbTab & OldGetServicePackNumber
      txtMain.Text = txtMain.Text & vbCrLf & "Extended Informations:" & vbTab & GetExtendedInformations
      txtMain.Text = txtMain.Text & vbCrLf & vbCrLf & "Language (English):" & vbTab & vbTab & GetWindowsLanguage(True)
      txtMain.Text = txtMain.Text & vbCrLf & "Language (Native):" & vbTab & vbTab & GetWindowsLanguage(False)
      txtMain.Text = txtMain.Text & vbCrLf & vbCrLf & "Product Family:" & vbTab & vbTab & GetProductFamily
      txtMain.Text = txtMain.Text & vbCrLf & vbCrLf & "64 Bit System:" & vbTab & vbTab & Is64BitSystem
      txtMain.Text = txtMain.Text & vbCrLf & vbCrLf & "Windows Folder:" & vbTab & vbTab & GetWindowsFolder
    On Error GoTo 0
    'Many of you are asking themselves why I didn't used a variable (it is faster to work with variables than to work with TextBoxes/Labels...) instead of this.
    'Well, under my Windows 2000, everything works great if I put all this stuff in a messaga box or in a variable, but I have sent the code to a friend which uses Windows XP and on his system, only the first 3 lines were displayed.
    'I have no idea why!

End Sub

'Unload the form (and free up space used for it - you HAVE to use this fellow programmers!!!)
Private Sub Form_Unload(Cancel As Integer)

    Set frmMain = Nothing

End Sub
