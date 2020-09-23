VERSION 5.00
Object = "{EEE78583-FE22-11D0-8BEF-0060081841DE}#1.0#0"; "XVOICE.DLL"
Begin VB.Form frmHelper 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "Helper"
   ClientHeight    =   4380
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8820
   ControlBox      =   0   'False
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4380
   ScaleWidth      =   8820
   Begin VB.CheckBox chkAudibleHelp 
      Height          =   195
      Left            =   150
      TabIndex        =   5
      Top             =   1890
      Width           =   195
   End
   Begin ACTIVEVOICEPROJECTLibCtl.DirectSS Voice 
      Height          =   435
      Left            =   6000
      OleObjectBlob   =   "frmHelper.frx":0000
      TabIndex        =   4
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtHelper 
      Height          =   1305
      Left            =   150
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   270
      Width           =   3375
   End
   Begin VB.CheckBox chkAlign 
      Height          =   195
      Left            =   150
      TabIndex        =   1
      Top             =   1650
      Width           =   195
   End
   Begin VB.Label lblAudibleHelp 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Audible Help"
      ForeColor       =   &H00FFFFC0&
      Height          =   195
      Left            =   450
      TabIndex        =   6
      Tag             =   "Label"
      ToolTipText     =   "Makes the help window speak"
      Top             =   1890
      Width           =   900
   End
   Begin VB.Label lblAlign 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Align to Windows"
      ForeColor       =   &H00FFFFC0&
      Height          =   195
      Left            =   450
      TabIndex        =   2
      Tag             =   "Label"
      ToolTipText     =   "Aligns the help window to other windows"
      Top             =   1650
      Width           =   1230
   End
   Begin VB.Image imgButton 
      Height          =   360
      Index           =   1
      Left            =   5970
      Picture         =   "frmHelper.frx":0058
      Top             =   510
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.Image imgButton 
      Height          =   360
      Index           =   0
      Left            =   5970
      Picture         =   "frmHelper.frx":1A7A
      Top             =   90
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.Label lblClose 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Close"
      ForeColor       =   &H00C0FFC0&
      Height          =   195
      Left            =   2850
      TabIndex        =   0
      Tag             =   "ButtonLabel"
      Top             =   1800
      Width           =   390
   End
   Begin VB.Image imgClose 
      Height          =   360
      Left            =   2520
      Picture         =   "frmHelper.frx":349C
      Stretch         =   -1  'True
      Top             =   1710
      Width           =   1035
   End
   Begin VB.Image Image1 
      Height          =   2190
      Left            =   0
      Picture         =   "frmHelper.frx":4EBE
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3705
   End
End
Attribute VB_Name = "frmHelper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub SpeakHelp()

'Audible Help...
If Help.AudibleHelp = True Then
    Call Say(txtHelper.Text, Voice)
End If

End Sub
Private Sub chkAlign_Click()

'Align the help window or not...
Help.HelpIsAligned = chkAlign.Value = 1
Call AlignHelpToForm

End Sub

Private Sub Form_Load()

'Load INI Settings...
Call LoadINISettings

'Program Colors...
Call SetColors(Me)

'Form Coordinates...
Me.Height = 2190
Me.Width = 3705

'Tells other forms whether or not the help form is loaded...
Help.HelpIsLoaded = True

End Sub
Sub LoadINISettings()

'Form Coordinates...
Me.Left = Val(ReadINI(Me.Name, "Left"))
Me.Top = Val(ReadINI(Me.Name, "Top"))

'Align...
chkAlign.Value = Val(ReadINI(Me.Name, "Align"))

'Audible Help...
chkAudibleHelp.Value = Val(ReadINI(Me.Name, "AudibleHelp"))
Help.AudibleHelp = chkAudibleHelp.Value = 1

End Sub

Private Sub Form_Unload(Cancel As Integer)

'Save INI Settings...
Call SaveINISettings

'Tells other forms whether or not the help form is loaded...
Help.HelpIsLoaded = False

End Sub
Sub SaveINISettings()

'Left and top properties...
Call WriteINI(Me.Name, "Left", Me.Left)
Call WriteINI(Me.Name, "Top", Me.Top)

'Alignment...
Call WriteINI(Me.Name, "Align", chkAlign.Value)

'Audible Help...
Call WriteINI(Me.Name, "AudibleHelp", chkAudibleHelp.Value)

End Sub
Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

'Move the form if the user is pressing and holding the mouse button...
If Button = vbLeftButton Then
    Call DragForm(Me)
    Call AlignHelpToForm
End If

End Sub
Private Sub imgClose_Click()

lblClose_Click

End Sub
Private Sub imgClose_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbLeftButton Then
    imgClose.Picture = imgButton(1).Picture
    lblClose.ForeColor = QBColor(0)
End If

End Sub

Private Sub imgClose_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

imgClose.Picture = imgButton(0).Picture
lblClose.ForeColor = lButtonForeColor

End Sub

Private Sub lblAlign_Click()

'Toggle on / off state...
If chkAlign.Value = False Then
    chkAlign.Value = 1
Else
    chkAlign.Value = False
End If

'Align the help window or not...
Help.HelpIsAligned = chkAlign.Value = 1
Call AlignHelpToForm

End Sub

Private Sub lblAlign_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Help.HelpText = "Aligns the help window to other windows."

End Sub
Private Sub lblAudibleHelp_Click()

'Toggle the audible help checkbox...
If chkAudibleHelp.Value = 1 Then
    chkAudibleHelp.Value = 0
Else
    chkAudibleHelp.Value = 1
End If

'Set to true or false...
Help.AudibleHelp = (chkAudibleHelp.Value = 1)
Call SpeakHelp

End Sub

Private Sub lblAudibleHelp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Help.HelpText = "Displays help audibly."

End Sub
Private Sub lblClose_Click()

Unload Me

End Sub
Private Sub lblClose_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbLeftButton Then
    imgClose.Picture = imgButton(1).Picture
    lblClose.ForeColor = QBColor(0)
End If

End Sub

Private Sub lblClose_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

imgClose.Picture = imgButton(0).Picture
lblClose.ForeColor = lButtonForeColor

End Sub

Private Sub txtHelper_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

'Move the form if the user is pressing and holding the mouse button...
If Button = vbLeftButton Then
    Call DragForm(Me)
    Call AlignHelpToForm
End If

End Sub
