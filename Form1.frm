VERSION 5.00
Object = "{F5BE8BC2-7DE6-11D0-91FE-00C04FD701A5}#2.0#0"; "AGENTCTL.DLL"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00800000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cartoon Cinema©"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   7080
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   7080
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cmd 
      Left            =   240
      Top             =   3480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   -600
      TabIndex        =   19
      Top             =   0
      Width           =   7935
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFC0C0&
      Caption         =   "New"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   2760
      Width           =   2295
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Stop"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   2760
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Play"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   2760
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Add"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   360
      Width           =   2055
   End
   Begin VB.ListBox List4 
      Height          =   255
      ItemData        =   "Form1.frx":0CCA
      Left            =   3480
      List            =   "Form1.frx":0DA0
      TabIndex        =   10
      Top             =   4440
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ListBox List3 
      Height          =   255
      ItemData        =   "Form1.frx":1096
      Left            =   2400
      List            =   "Form1.frx":119F
      TabIndex        =   9
      Top             =   4440
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ListBox List2 
      Height          =   255
      ItemData        =   "Form1.frx":15A5
      Left            =   1320
      List            =   "Form1.frx":168A
      TabIndex        =   8
      Top             =   4440
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ListBox List1 
      Height          =   255
      ItemData        =   "Form1.frx":19CB
      Left            =   240
      List            =   "Form1.frx":1AB9
      TabIndex        =   7
      Top             =   4440
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ComboBox Combo2 
      BackColor       =   &H00FFC0C0&
      Height          =   315
      Left            =   2520
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   360
      Width           =   2175
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00FFC0C0&
      Height          =   315
      ItemData        =   "Form1.frx":1E15
      Left            =   120
      List            =   "Form1.frx":1E28
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   360
      Width           =   1815
   End
   Begin VB.ListBox ActionList 
      BackColor       =   &H00FFFFFF&
      Height          =   1425
      ItemData        =   "Form1.frx":1E53
      Left            =   4680
      List            =   "Form1.frx":1E55
      TabIndex        =   2
      Top             =   1200
      Width           =   2295
   End
   Begin VB.ListBox EventList 
      BackColor       =   &H00FFFFFF&
      Height          =   1425
      ItemData        =   "Form1.frx":1E57
      Left            =   2400
      List            =   "Form1.frx":1E59
      TabIndex        =   1
      Top             =   1200
      Width           =   2295
   End
   Begin VB.ListBox CharList 
      BackColor       =   &H00FFFFFF&
      Height          =   1425
      ItemData        =   "Form1.frx":1E5B
      Left            =   120
      List            =   "Form1.frx":1E5D
      TabIndex        =   0
      Top             =   1200
      Width           =   2295
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Caption         =   "The Actors"
      Height          =   255
      Left            =   240
      TabIndex        =   21
      Top             =   4920
      Width           =   4335
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "The Entire List Of Commands For Each Character"
      Height          =   255
      Left            =   240
      TabIndex        =   20
      Top             =   4080
      Width           =   4335
   End
   Begin VB.Image DirectorPic 
      Height          =   615
      Left            =   3360
      Picture         =   "Form1.frx":1E5F
      Stretch         =   -1  'True
      Top             =   5280
      Width           =   615
   End
   Begin VB.Image RobbyPic 
      Height          =   615
      Left            =   2520
      Picture         =   "Form1.frx":2B29
      Stretch         =   -1  'True
      Top             =   5280
      Width           =   615
   End
   Begin VB.Image PeedyPic 
      Height          =   615
      Left            =   1680
      Picture         =   "Form1.frx":3211
      Stretch         =   -1  'True
      Top             =   5280
      Width           =   615
   End
   Begin VB.Image MerlinPic 
      Height          =   615
      Left            =   960
      Picture         =   "Form1.frx":3838
      Stretch         =   -1  'True
      Top             =   5280
      Width           =   615
   End
   Begin VB.Image GeniePic 
      Height          =   615
      Left            =   240
      Picture         =   "Form1.frx":3DEA
      Stretch         =   -1  'True
      Top             =   5280
      Width           =   615
   End
   Begin VB.Image Image1 
      Height          =   615
      Left            =   1920
      Stretch         =   -1  'True
      Top             =   80
      Width           =   615
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Note:  Select a character from the list below and press the DEL key to remove the entire action."
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Left            =   120
      TabIndex        =   18
      Top             =   720
      Width           =   6735
   End
   Begin AgentObjectsCtl.Agent MSagent 
      Left            =   3360
      Top             =   1320
      _cx             =   847
      _cy             =   847
   End
   Begin VB.Label Label5 
      BackColor       =   &H0080FF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Action"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4680
      TabIndex        =   14
      Top             =   960
      Width           =   2295
   End
   Begin VB.Label Label4 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Event"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2400
      TabIndex        =   13
      Top             =   960
      Width           =   2295
   End
   Begin VB.Label Label3 
      BackColor       =   &H0080C0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Character"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   960
      Width           =   2295
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select Action"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2520
      TabIndex        =   6
      Top             =   120
      Width           =   1155
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select Character"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   1440
   End
   Begin VB.Menu ScriptMNU 
      Caption         =   "Script"
      Begin VB.Menu NewMNU 
         Caption         =   "Start a New Script"
      End
      Begin VB.Menu OpenMNU 
         Caption         =   "Open a Saved Script"
      End
      Begin VB.Menu SaveMNU 
         Caption         =   "Save this Script"
      End
      Begin VB.Menu linageMNU 
         Caption         =   "-"
      End
      Begin VB.Menu ExitMNU 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'I.     Set Up Characters
Private Genie As IAgentCtlCharacterEx
Private Merlin As IAgentCtlCharacterEx
Private Peedy As IAgentCtlCharacterEx
Private Robby As IAgentCtlCharacterEx

'II.    Set Booleans
Private bContinue As Boolean
Private bReady As Boolean
Private bUnload As Boolean

'III.   Set Constants
Const BalloonOn = 1
Const SizeToText = 2
Const AutoHide = 4
Const AutoPace = 8


Sub DirectorAction()
'Pauses for the specified amount of seconds
Pause (Val(ActionList.Text))
End Sub


Sub GenieAction()
'Control section to figure out what
'the user wants the character(s) to
'do or say.

If ActionList.Text = "Hide" Then
Genie.Hide
Exit Sub
End If

If ActionList.Text = "Show" Then
Genie.Show
Genie.Speak "\Mrk=99\"
bReady = False
Do Until bReady
  DoEvents
Loop
Exit Sub
End If

If EventList.Text = "MoveToPoint" Then
Dim sCoordinates As String
Dim lX As Long
Dim lY As Long
sCoordinates = ActionList.List(ActionList.ListIndex)
lX = CLng(Left$(sCoordinates, InStr(1, sCoordinates, ",") - 1))
lY = CLng(Right$(sCoordinates, InStrRev(sCoordinates, ",")))
Genie.MoveTo lX, lY
Genie.Show
Genie.Speak "\Mrk=99\"
bReady = False
Do Until bReady
  DoEvents
Loop
Exit Sub
End If

If EventList.Text = "Speak" Then
Genie.Stop
Genie.Show
Genie.Speak ActionList.List(ActionList.ListIndex)
bContinue = False
    Do Until bContinue
        DoEvents
    Loop
Exit Sub
End If
Genie.Show
Genie.Play ActionList.List(ActionList.ListIndex)
Genie.Speak "\Mrk=99\"
bReady = False
Do Until bReady
  DoEvents
Loop
End Sub


Sub MerlinAction()
'Control section to figure out what
'the user wants the character(s) to
'do or say.

If ActionList.Text = "Hide" Then
Merlin.Hide
Exit Sub
End If

If ActionList.Text = "Show" Then
Merlin.Show
Merlin.Speak "\Mrk=99\"
bReady = False
Do Until bReady
  DoEvents
Loop
Exit Sub
End If

If EventList.Text = "MoveToPoint" Then
Dim sCoordinates As String
Dim lX As Long
Dim lY As Long
sCoordinates = ActionList.List(ActionList.ListIndex)
lX = CLng(Left$(sCoordinates, InStr(1, sCoordinates, ",") - 1))
lY = CLng(Right$(sCoordinates, InStrRev(sCoordinates, ",")))
Merlin.MoveTo lX, lY
Merlin.Show
Merlin.Speak "\Mrk=99\"
bReady = False
Do Until bReady
  DoEvents
Loop
Exit Sub
End If

If EventList.Text = "Speak" Then
Merlin.Stop
Merlin.Show
Merlin.Speak ActionList.List(ActionList.ListIndex)
bContinue = False
    Do Until bContinue
        DoEvents
    Loop
Exit Sub
End If
Merlin.Show
Merlin.Play ActionList.List(ActionList.ListIndex)
Merlin.Speak "\Mrk=99\"
bReady = False
Do Until bReady
  DoEvents
Loop
End Sub

Sub PeedyAction()
'Control section to figure out what
'the user wants the character(s) to
'do or say.

If ActionList.Text = "Hide" Then
Peedy.Hide
Exit Sub
End If

If ActionList.Text = "Show" Then
Peedy.Show
Peedy.Speak "\Mrk=99\"
bReady = False
Do Until bReady
  DoEvents
Loop
Exit Sub
End If

If EventList.Text = "MoveToPoint" Then
Dim sCoordinates As String
Dim lX As Long
Dim lY As Long
sCoordinates = ActionList.List(ActionList.ListIndex)
lX = CLng(Left$(sCoordinates, InStr(1, sCoordinates, ",") - 1))
lY = CLng(Right$(sCoordinates, InStrRev(sCoordinates, ",")))
Peedy.MoveTo lX, lY
Peedy.Show
Peedy.Speak "\Mrk=99\"
bReady = False
Do Until bReady
  DoEvents
Loop
Exit Sub
End If

If EventList.Text = "Speak" Then
Peedy.Stop
Peedy.Show
Peedy.Speak ActionList.List(ActionList.ListIndex)
bContinue = False
    Do Until bContinue
        DoEvents
    Loop
Exit Sub
End If
Peedy.Show
Peedy.Play ActionList.List(ActionList.ListIndex)
Peedy.Speak "\Mrk=99\"
bReady = False
Do Until bReady
  DoEvents
Loop
End Sub

Sub RobbyAction()
'Control section to figure out what
'the user wants the character(s) to
'do or say.

If ActionList.Text = "Hide" Then
Robby.Hide
Exit Sub
End If

If ActionList.Text = "Show" Then
Robby.Show
Robby.Speak "\Mrk=99\"
bReady = False
Do Until bReady
  DoEvents
Loop
Exit Sub
End If

If EventList.Text = "MoveToPoint" Then
Dim sCoordinates As String
Dim lX As Long
Dim lY As Long
sCoordinates = ActionList.List(ActionList.ListIndex)
lX = CLng(Left$(sCoordinates, InStr(1, sCoordinates, ",") - 1))
lY = CLng(Right$(sCoordinates, InStrRev(sCoordinates, ",")))
Robby.MoveTo lX, lY
Robby.Show
Robby.Speak "\Mrk=99\"
bReady = False
Do Until bReady
  DoEvents
Loop
Exit Sub
End If

If EventList.Text = "Speak" Then
Robby.Stop
Robby.Show
Robby.Speak ActionList.List(ActionList.ListIndex)
bContinue = False
    Do Until bContinue
        DoEvents
    Loop
Exit Sub
End If
Robby.Show
Robby.Play ActionList.List(ActionList.ListIndex)
Robby.Speak "\Mrk=99\"
bReady = False
Do Until bReady
  DoEvents
Loop
End Sub

Private Sub ActionList_Click()
If ActionList.ListCount = 0 Then Exit Sub
If ActionList.Text = "" Then Exit Sub
CharList.ListIndex = ActionList.ListIndex
EventList.ListIndex = ActionList.ListIndex
End Sub

Private Sub CharList_Click()
If CharList.ListCount = 0 Then Exit Sub
If CharList.Text = "" Then Exit Sub
EventList.ListIndex = CharList.ListIndex
ActionList.ListIndex = CharList.ListIndex
End Sub

Private Sub CharList_KeyDown(KeyCode As Integer, Shift As Integer)
If CharList.ListCount = 0 Then Exit Sub
If CharList.Text = "" Then Exit Sub
If KeyCode = vbKeyDelete Then
CharList.RemoveItem CharList.ListIndex
EventList.RemoveItem EventList.ListIndex
ActionList.RemoveItem ActionList.ListIndex
Exit Sub
End If
End Sub


Private Sub Combo1_Click()
Dim X

'Add only the Pause function to the Action list.
If Combo1.Text = "Director" Then
Image1.Picture = DirectorPic.Picture
Combo2.Clear
Combo2.AddItem "Pause"
Exit Sub
End If

'Set the picture of the character
'the user has chosen and flood the
'action list box with the specific
'actions for that character

If Combo1.Text = "Genie" Then
Image1.Picture = GeniePic.Picture
Combo2.Clear
For X = 0 To List1.ListCount - 1
List1.ListIndex = X
Combo2.AddItem List1.Text
Next X
Combo2.ListIndex = 0
Exit Sub
End If

If Combo1.Text = "Merlin" Then
Image1.Picture = MerlinPic.Picture
Combo2.Clear
For X = 0 To List2.ListCount - 1
List2.ListIndex = X
Combo2.AddItem List2.Text
Next X
Combo2.ListIndex = 0
Exit Sub
End If

If Combo1.Text = "Peedy" Then
Image1.Picture = PeedyPic.Picture
Combo2.Clear
For X = 0 To List3.ListCount - 1
List3.ListIndex = X
Combo2.AddItem List3.Text
Next X
Combo2.ListIndex = 0
Exit Sub
End If

If Combo1.Text = "Robby" Then
Image1.Picture = RobbyPic.Picture
Combo2.Clear
For X = 0 To List4.ListCount - 1
List4.ListIndex = X
Combo2.AddItem List4.Text
Next X
Combo2.ListIndex = 0
Exit Sub
End If

End Sub


Private Sub Combo2_Click()
Dim Message, Title, Default, MyValue

'Determine what the user has chosen

'Display input box for amount of seconds
If Combo2.Text = "Pause" Then
Message = "Enter number of seconds to pause."
Title = "INSERT SECONDS"
MyValue = InputBox(Message, Title)
If MyValue = "" Then
Command1.Tag = ""
Exit Sub
End If
Command1.Tag = MyValue
Exit Sub
End If


'Display input box for X and Y coordinates
If Combo2.Text = "MoveToPoint" Then
Message = "Enter points to move to. (Ex.  255, 143)"
Title = "INSERT TEXT"
MyValue = InputBox(Message, Title)
If MyValue = "" Then
Command1.Tag = ""
Exit Sub
End If
Command1.Tag = MyValue
Exit Sub
End If

'Display input box for a sentence to be spoken
If Combo2.Text = "SPEAK" Then
Message = "Enter the text to speak."
Title = "INSERT TEXT"
MyValue = InputBox(Message, Title)
If MyValue = "" Then
Command1.Tag = ""
Exit Sub
End If
Command1.Tag = MyValue
Exit Sub
End If
End Sub


Private Sub Command1_Click()
'Error Handling
If (Combo2.Text = "SPEAK") And (Command1.Tag = "") Then
MsgBox "Cannot use SPEAK event without an entry."
Exit Sub
End If
If (Combo2.Text = "MoveToPoint") And (Command1.Tag = "") Then
MsgBox "Cannot use MoveToPoint event without an entry."
Exit Sub
End If
If (Combo2.Text = "Pause") And (Command1.Tag = "") Then
MsgBox "Cannot use Pause event without an entry."
Exit Sub
End If

'Add the items to their respective list boxes

CharList.AddItem Combo1.Text

If Combo2.Text = "Pause" Then
EventList.AddItem "Pause"
ActionList.AddItem Command1.Tag
Exit Sub
End If


If Combo2.Text = "SPEAK" Then
EventList.AddItem "Speak"
ActionList.AddItem Command1.Tag
Exit Sub
End If

If Combo2.Text = "MoveToPoint" Then
EventList.AddItem "MoveToPoint"
ActionList.AddItem Command1.Tag
Exit Sub
End If

EventList.AddItem "Action"
ActionList.AddItem Combo2.Text
End Sub


Private Sub Command2_Click()
'Play the script

'Error Handling
If CharList.ListCount = 0 Then Exit Sub

'Remove the list boxes from view for
'faster animation roll-over
Label3.Visible = False
Label4.Visible = False
Label5.Visible = False
CharList.Visible = False
EventList.Visible = False
ActionList.Visible = False

'Move the app off the stage
Form1.WindowState = 1

'CHARACTER ENGINE:  Determines who says
'                   or does what
Dim X
For X = 0 To CharList.ListCount - 1
CharList.ListIndex = X
If CharList.Text = "Director" Then
Call DirectorAction
End If
If CharList.Text = "Genie" Then
Genie.Stop
Call GenieAction
End If
If CharList.Text = "Merlin" Then
Merlin.Stop
Call MerlinAction
End If
If CharList.Text = "Peedy" Then
Peedy.Stop
Call PeedyAction
End If
If CharList.Text = "Robby" Then
Robby.Stop
Call RobbyAction
End If
Next X

'Allow time for closure of script
Pause (2)

'Bring everything back on the set
Label3.Visible = True
Label4.Visible = True
Label5.Visible = True
CharList.Visible = True
EventList.Visible = True
ActionList.Visible = True
Form1.WindowState = 0
End Sub

Private Sub Command3_Click()
'Stop everything and yell "CUT!"
If CharList.ListCount = 0 Then
Genie.Stop
Genie.Hide
End If
Label3.Visible = True
Label4.Visible = True
Label5.Visible = True
CharList.Visible = True
EventList.Visible = True
ActionList.Visible = True
Genie.Stop
Merlin.Stop
Peedy.Stop
Robby.Stop
Genie.Hide
Merlin.Hide
Peedy.Hide
Robby.Hide
Do
DoEvents
Loop
End Sub

Private Sub Command4_Click()
'Start A New Project procedure
Dim Answer
Answer = MsgBox("Start new project?", vbQuestion + vbYesNo, "Cartoon Cinema")
If Answer = 6 Then
Genie.Stop
Merlin.Stop
Peedy.Stop
Robby.Stop
Genie.Hide
Merlin.Hide
Peedy.Hide
Robby.Hide
CharList.Clear
EventList.Clear
ActionList.Clear
Command1.Tag = ""
Exit Sub
End If
End Sub


Private Sub EventList_Click()
If EventList.ListCount = 0 Then Exit Sub
If EventList.Text = "" Then Exit Sub
CharList.ListIndex = EventList.ListIndex
ActionList.ListIndex = EventList.ListIndex
End Sub


Private Sub ExitMNU_Click()
Genie.Stop
Merlin.Stop
Peedy.Stop
Robby.Stop
Unload Me
End
End Sub

Private Sub Form_Load()
    bUnload = False
    
    'Load Characters
    MSagent.Characters.Load "Genie", "genie.acs"
    MSagent.Characters.Load "Merlin", "merlin.acs"
    MSagent.Characters.Load "Peedy", "peedy.acs"
    MSagent.Characters.Load "Robby", "Robby.acs"
    
    Set Genie = MSagent.Characters("Genie")
    Set Merlin = MSagent.Characters("Merlin")
    Set Peedy = MSagent.Characters("Peedy")
    Set Robby = MSagent.Characters("Robby")
    
    Genie.LanguageID = &H409
    Merlin.LanguageID = &H409
    Peedy.LanguageID = &H409
    Robby.LanguageID = &H409
    
    Genie.Balloon.Style = BalloonOn Or AutoHide Or AutoPace Or SizeToText
    Merlin.Balloon.Style = BalloonOn Or AutoHide Or AutoPace Or SizeToText
    Peedy.Balloon.Style = BalloonOn Or AutoHide Or AutoPace Or SizeToText
    Robby.Balloon.Style = BalloonOn Or AutoHide Or AutoPace Or SizeToText
Show
'Begin Brief Tutorial
Genie.MoveTo 330, 70
Genie.Show
Genie.Speak "Welcome to Cartoon \emp\Cinema!  In \emp\this brief tutorial, I'll show you how to create your \emp\very \emp\own animations."
bContinue = False
    Do Until bContinue
        DoEvents
    Loop
Genie.Speak "To stop this tutorial at any time, click the \emp\Stop button."
bContinue = False
    Do Until bContinue
        DoEvents
    Loop
Genie.MoveTo 250, 90
Genie.Play "GestureDown"
Genie.Speak "First of all, if you'll look over here, you'll  notice \emp\two lists that will drop when you click on them."
bContinue = False
    Do Until bContinue
        DoEvents
    Loop
Genie.Speak "The one on the left is for choosing a character, like \emp\me.  You'll also notice a \emp\Director character, which lets you add a \emp\pause between animations."
bContinue = False
    Do Until bContinue
        DoEvents
    Loop
Genie.MoveTo 330, 90
Genie.Play "GestureDown"
Genie.Speak "\emp\This list displays all the movements your chosen character can make.  The list may be empty now, because a character might not have been selected yet.  The Director's pause event is \emp\also in here."
bContinue = False
    Do Until bContinue
        DoEvents
    Loop
Genie.MoveTo 500, 90
Genie.Play "GestureDown"
Genie.Speak "After selecting a character and a movement, click the Add button to add them to your \emp\script."
bContinue = False
    Do Until bContinue
        DoEvents
    Loop
Genie.MoveTo 330, 70
Genie.Speak "\emp\If you have chosen the Speak event, a dialog box will be displayed and will allow you enter the sentence you want the selected character to \emp\say.  If you want to put emphasis on a word, put a back slash on either side of the letters, EMP.  The word you want to put emphasis on must immediately follow this insertion."
bContinue = False
    Do Until bContinue
        DoEvents
    Loop
Genie.Speak "The Move To event displays a dialog box.  Here, you will enter two numbers.  The first one will be the X coordinates on the screen, and the second one will be the Y coordinates.  If you don't know exactly where the coordinates are, \emp\experiment a little.  The numbers must be separated with a comma and a space, like \emp\this..."
bContinue = False
    Do Until bContinue
        DoEvents
    Loop
Genie.Speak "255, 255"
bContinue = False
    Do Until bContinue
        DoEvents
    Loop
Genie.Speak "The Director's pause event displays a dialog box, asking you to enter the number of seconds you want the script to pause between animations.  All animations have their own pause event following an action."
bContinue = False
    Do Until bContinue
        DoEvents
    Loop
Genie.Play "Greet"
Genie.Speak "We have come to the end of the tutorial...."
bContinue = False
    Do Until bContinue
        DoEvents
    Loop
Genie.Play "RestPose"
Genie.Play "Explain"
Genie.Speak "Have \emp\Fun!"
bContinue = False
    Do Until bContinue
        DoEvents
    Loop
Genie.Hide
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
End
End Sub


Private Sub Form_Terminate()
End
End Sub


Private Sub Form_Unload(Cancel As Integer)
End
End Sub


Private Sub MSagent_BalloonHide(ByVal CharacterID As String)
'This is one of the primary lines
'of code.  This little boolean
'tells the script that a character
'is finished talking.
    bContinue = True
End Sub

Private Sub MSagent_Bookmark(ByVal BookmarkID As Long)
'This is one of the primary lines
'of code.  This little boolean
'tells the script that a character
'is finished moving.

bReady = True
End Sub


Private Sub MSagent_Hide(ByVal CharacterID As String, ByVal Cause As Integer)
    If bUnload Then bContinue = True
End Sub


Private Sub MSagent_IdleStart(ByVal CharacterID As String)
bContinue = True
End Sub


Private Sub NewMNU_Click()
Command4_Click
End Sub


Private Sub OpenMNU_Click()
'Open a saved script
On Error Resume Next
cmd.DialogTitle = "Open Script..."
cmd.FileName = "*.ccs"
cmd.InitDir = CurDir
cmd.Filter = "Script Files (*.ccs)"
cmd.ShowOpen
If (cmd.FileName = "") Or (cmd.FileName = "*.ccs") Then Exit Sub
Dim sFile As String, sShortFile As String * 67
Dim lRet As Long
sFile = cmd.FileName
lRet = GetShortPathName(sFile, sShortFile, Len(sShortFile))
sFile = Left(sShortFile, lRet)
CharList.Clear
EventList.Clear
ActionList.Clear
Open sFile For Input As 1
Dim a As String
Do Until EOF(1)
Input #1, a
CharList.AddItem a
Input #1, a
EventList.AddItem a
Input #1, a
ActionList.AddItem a
Loop
Close 1
End Sub


Private Sub SaveMNU_Click()
'Save the Current Script
Dim X
Dim G
cmd.DialogTitle = "Save Script..."
cmd.FileName = "Script.ccs"
cmd.InitDir = CurDir
cmd.Filter = "Script Files (*.ccs)"
cmd.ShowOpen
X = cmd.FileName
Open X For Output As #1
CharList.Visible = False
EventList.Visible = False
ActionList.Visible = False
For G = 0 To CharList.ListCount - 1
CharList.ListIndex = G
Print #1, CharList.Text
Print #1, EventList.Text
Print #1, ActionList.Text
Next G
Close #1
CharList.Visible = True
EventList.Visible = True
ActionList.Visible = True
Dim Answer
Answer = MsgBox("Your script has been saved.", vbInformation + vbOKOnly, "Cartoon Cinema")
End Sub


