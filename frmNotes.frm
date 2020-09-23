VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{EEE78583-FE22-11D0-8BEF-0060081841DE}#1.0#0"; "XVoice.dll"
Begin VB.Form frmNotes 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Notes"
   ClientHeight    =   7005
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11085
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmNotes.frx":0000
   ScaleHeight     =   7005
   ScaleWidth      =   11085
   ShowInTaskbar   =   0   'False
   Begin MSComDlg.CommonDialog Dialog 
      Left            =   9390
      Top             =   1470
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.TextBox txtNotes 
      BackColor       =   &H00C0E0FF&
      Height          =   3045
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   570
      Width           =   6525
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   9390
      Top             =   990
   End
   Begin ACTIVEVOICEPROJECTLibCtl.DirectSS Voice 
      Height          =   525
      Left            =   9390
      OleObjectBlob   =   "frmNotes.frx":62BC2
      TabIndex        =   7
      Top             =   1980
      Width           =   525
   End
   Begin VB.Label lblHelp 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Help..."
      ForeColor       =   &H00C0FFC0&
      Height          =   195
      Left            =   1095
      TabIndex        =   8
      Tag             =   "ButtonLabel"
      Top             =   3840
      Width           =   495
   End
   Begin VB.Label lblSpeak 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Speak"
      ForeColor       =   &H00C0FFC0&
      Height          =   195
      Left            =   2100
      TabIndex        =   6
      Tag             =   "ButtonLabel"
      Top             =   3840
      Width           =   465
   End
   Begin VB.Label lblPrint 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Print"
      ForeColor       =   &H00C0FFC0&
      Height          =   195
      Left            =   3180
      TabIndex        =   5
      Tag             =   "ButtonLabel"
      Top             =   3840
      Width           =   315
   End
   Begin VB.Label lblClear 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Clear"
      ForeColor       =   &H00C0FFC0&
      Height          =   195
      Left            =   4140
      TabIndex        =   4
      Tag             =   "ButtonLabel"
      Top             =   3840
      Width           =   360
   End
   Begin VB.Image imgButton 
      Height          =   375
      Index           =   1
      Left            =   9390
      Picture         =   "frmNotes.frx":62C1A
      Top             =   540
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.Image imgButton 
      Height          =   360
      Index           =   0
      Left            =   9390
      Picture         =   "frmNotes.frx":64304
      Stretch         =   -1  'True
      Top             =   150
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.Label lblSave 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Save"
      Enabled         =   0   'False
      ForeColor       =   &H00C0FFC0&
      Height          =   195
      Left            =   5100
      TabIndex        =   1
      Tag             =   "ButtonLabel"
      Top             =   3840
      Width           =   375
   End
   Begin VB.Label lblExit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Exit"
      ForeColor       =   &H00C0FFC0&
      Height          =   195
      Left            =   6180
      TabIndex        =   0
      Tag             =   "ButtonLabel"
      Top             =   3840
      Width           =   255
   End
   Begin VB.Image imgExit 
      Height          =   375
      Left            =   5790
      Picture         =   "frmNotes.frx":65D26
      Stretch         =   -1  'True
      Top             =   3750
      Width           =   1005
   End
   Begin VB.Image imgSave 
      Enabled         =   0   'False
      Height          =   375
      Left            =   4800
      Picture         =   "frmNotes.frx":67748
      Stretch         =   -1  'True
      Top             =   3750
      Width           =   1005
   End
   Begin VB.Label lblCategories 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Notes"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   1
      Left            =   330
      TabIndex        =   2
      Top             =   45
      Width           =   525
   End
   Begin VB.Image imgClear 
      Height          =   375
      Left            =   3810
      Picture         =   "frmNotes.frx":6916A
      Stretch         =   -1  'True
      Top             =   3750
      Width           =   1005
   End
   Begin VB.Image imgPrint 
      Height          =   375
      Left            =   2820
      Picture         =   "frmNotes.frx":6AB8C
      Stretch         =   -1  'True
      Top             =   3750
      Width           =   1005
   End
   Begin VB.Image imgSpeak 
      Height          =   375
      Left            =   1830
      Picture         =   "frmNotes.frx":6C5AE
      Stretch         =   -1  'True
      Top             =   3750
      Width           =   1005
   End
   Begin VB.Image imgHelp 
      Height          =   375
      Left            =   840
      Picture         =   "frmNotes.frx":6DFD0
      Stretch         =   -1  'True
      Top             =   3750
      Width           =   1005
   End
End
Attribute VB_Name = "frmNotes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim iDirty As Boolean
Private Sub Form_Load()

'Load the main menu's form settings...
Call LoadINISettings

'Set program colors...
Call SetColors(Me)

'Width and Height...
Me.Width = QuickRef.MediumMenuWidth
Me.Height = QuickRef.MediumMenuHeight

'Get Notes...
If QuickRef.PassNotes = False Then
    Call GetNotes(txtNotes)
End If

iDirty = False

End Sub
Sub LoadINISettings()

'Form Coordinates...
Me.Left = Val(ReadINI(Me.Name, "Left"))
Me.Top = Val(ReadINI(Me.Name, "Top"))

End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

Help.HelpText = ""

'Move the form if the user is pressing and holding the mouse button...
If Button = vbLeftButton Then
    Call DragForm(Me)
End If

End Sub
Private Sub Form_Unload(Cancel As Integer)

On Local Error Resume Next

Dim x As Long

'Save changes...
If iDirty Then
    x = MsgBox("Save changes before exiting?", vbYesNoCancel + vbQuestion, "Save Changes...")
    Select Case x
        Case vbYes
            If SaveChanges() = False Then
                If MsgBox("Changes were not saved. Do you still want to exit anyway?", vbYesNo + vbQuestion, "Exit without saving...") = vbNo Then
                    Cancel = True
                    Exit Sub
                End If
            End If
        Case vbCancel
            Cancel = True
            Exit Sub
    End Select
End If

'Save INI Settings...
Call SaveINISettings

End Sub
Sub SaveINISettings()

'Form coordinates...
Call WriteINI(Me.Name, "Left", Me.Left)
Call WriteINI(Me.Name, "Top", Me.Top)

End Sub
Private Sub imgClear_Click()

lblClear_Click

End Sub

Private Sub imgClear_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

If Button = vbLeftButton Then
    imgClear.Picture = imgButton(1).Picture
    lblClear.ForeColor = QBColor(0)
End If

End Sub
Private Sub imgClear_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

imgClear.Picture = imgButton(0).Picture
lblClear.ForeColor = lButtonForeColor

End Sub
Private Sub imgExit_Click()

lblExit_Click

End Sub
Private Sub imgExit_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

If Button = vbLeftButton Then
    imgExit.Picture = imgButton(1).Picture
    lblExit.ForeColor = QBColor(0)
End If

End Sub
Private Sub imgExit_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

imgExit.Picture = imgButton(0).Picture
lblExit.ForeColor = lButtonForeColor

End Sub

Private Sub imgHelp_Click()

lblHelp_Click

End Sub

Private Sub imgHelp_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

If Button = vbLeftButton Then
    imgHelp.Picture = imgButton(1).Picture
    lblHelp.ForeColor = QBColor(0)
End If

End Sub

Private Sub imgHelp_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

imgHelp.Picture = imgButton(0).Picture
lblHelp.ForeColor = lButtonForeColor

End Sub
Private Sub imgPrint_Click()

lblPrint_Click

End Sub
Private Sub imgPrint_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

If Button = vbLeftButton Then
    imgPrint.Picture = imgButton(1).Picture
    lblPrint.ForeColor = QBColor(0)
End If

End Sub
Private Sub imgPrint_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

imgPrint.Picture = imgButton(0).Picture
lblPrint.ForeColor = lButtonForeColor

End Sub
Private Sub imgSave_Click()

lblSave_Click

End Sub
Private Sub imgSave_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

If Button = vbLeftButton Then
    imgSave.Picture = imgButton(1).Picture
    lblSave.ForeColor = QBColor(0)
End If

End Sub
Private Sub imgSave_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

imgSave.Picture = imgButton(0).Picture
lblSave.ForeColor = lButtonForeColor

End Sub

Private Sub imgSpeak_Click()

Call lblSpeak_Click

End Sub
Private Sub imgSpeak_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

If Button = vbLeftButton Then
    imgSpeak.Picture = imgButton(1).Picture
    lblSpeak.ForeColor = QBColor(0)
End If

End Sub

Private Sub imgSpeak_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

imgSpeak.Picture = imgButton(0).Picture
lblSpeak.ForeColor = lButtonForeColor

End Sub
Private Sub lblCategories_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)

'Move the form if the user is pressing and holding the mouse button...
If Button = vbLeftButton Then
    Call DragForm(Me)
End If

End Sub

Private Sub lblClear_Click()

txtNotes = ""

End Sub
Private Sub lblClear_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

If Button = vbLeftButton Then
    imgClear.Picture = imgButton(1).Picture
    lblClear.ForeColor = QBColor(0)
End If

End Sub

Private Sub lblClear_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

Help.HelpText = "Click here to clear the notes."

End Sub
Private Sub lblClear_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

imgClear.Picture = imgButton(0).Picture
lblClear.ForeColor = lButtonForeColor

End Sub
Private Sub lblExit_Click()

'Unload the help window...
If Help.HelpCallingForm = Me.Name Then
    Unload frmHelper
End If

Unload Me

End Sub
Private Sub lblExit_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

If Button = vbLeftButton Then
    imgExit.Picture = imgButton(1).Picture
    lblExit.ForeColor = QBColor(0)
End If

End Sub

Private Sub lblExit_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

Help.HelpText = "Click here to exit the notes window."

End Sub
Private Sub lblExit_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

imgExit.Picture = imgButton(0).Picture
lblExit.ForeColor = lButtonForeColor

End Sub

Private Sub lblHelp_Click()

Help.HelpCallingForm = Me.Name

frmHelper.Show
frmHelper.ZOrder

End Sub
Private Sub lblHelp_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

If Button = vbLeftButton Then
    imgHelp.Picture = imgButton(1).Picture
    lblHelp.ForeColor = QBColor(0)
End If

End Sub

Private Sub lblHelp_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

Help.HelpText = "Click here to show the help window."

End Sub
Private Sub lblHelp_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

imgHelp.Picture = imgButton(0).Picture
lblHelp.ForeColor = lButtonForeColor

End Sub
Private Sub lblPrint_Click()

On Local Error Resume Next

Dialog.ShowPrinter

If Dialog.CancelError Then Exit Sub

'Print the notes to the printer...
Printer.Print txtNotes
Printer.EndDoc

End Sub
Private Sub lblPrint_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

If Button = vbLeftButton Then
    imgPrint.Picture = imgButton(1).Picture
    lblPrint.ForeColor = QBColor(0)
End If

End Sub

Private Sub lblPrint_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

Help.HelpText = "Click here to print the notes to the printer."

End Sub
Private Sub lblPrint_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

imgPrint.Picture = imgButton(0).Picture
lblPrint.ForeColor = lButtonForeColor

End Sub
Private Sub lblSave_Click()

Call SaveChanges

End Sub
Function SaveChanges() As Boolean

On Local Error GoTo SaveChangesError

If OpenDB(DB, QuickRef.DBPassWord) = False Then Exit Function
Set RS = DB.OpenRecordset("SELECT Comments FROM tblContacts WHERE ContactName = '" & QuickRef.ContactName & "' AND ContactID = " & QuickRef.ContactID, dbOpenDynaset)

'This should never happen!...
If RS.RecordCount = 0 Then
    RS.Close
    DB.Close
    Exit Function
End If

'Save the notes...
RS.Edit
RS!Comments = txtNotes
RS.Update
RS.Close
DB.Close

QuickRef.UpdateNotes = QuickRef.NotesHaveChanged
SaveChanges = True
iDirty = False

Exit Function



SaveChangesError:
    DB.Close
    Call WriteToErrorLog(Me.Name, "SaveChangesError", Error, Err, True)
    Exit Function

End Function
Private Sub lblSave_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

If Button = vbLeftButton Then
    imgSave.Picture = imgButton(1).Picture
    lblSave.ForeColor = QBColor(0)
End If

End Sub

Private Sub lblSave_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

Help.HelpText = "Click here to save any changes that you have made."

End Sub
Private Sub lblSave_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

imgSave.Picture = imgButton(0).Picture
lblSave.ForeColor = lButtonForeColor

End Sub

Private Sub lblSpeak_Click()

'Speak the notes...
If txtNotes.SelText <> "" Then
    Call Say(txtNotes.SelText, Voice)
Else
    Call Say(txtNotes, Voice)
End If

End Sub
Private Sub lblSpeak_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

If Button = vbLeftButton Then
    imgSpeak.Picture = imgButton(1).Picture
    lblSpeak.ForeColor = QBColor(0)
End If

End Sub

Private Sub lblSpeak_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

imgSpeak.Picture = imgButton(0).Picture
lblSpeak.ForeColor = lButtonForeColor

End Sub
Private Sub Timer1_Timer()

On Local Error Resume Next

Dim iTempDirty As Boolean

'Notes...
If QuickRef.UpdateNotes And QuickRef.CallingForm <> "Notes" Then
    iTempDirty = iDirty
    Call GetNotes(txtNotes)
    iDirty = iTempDirty
End If

'Save...
If imgSave.Enabled = False And iDirty = True Then
    imgSave.Enabled = True
    lblSave.Enabled = True
ElseIf imgSave.Enabled = True And iDirty = False Then
    imgSave.Enabled = False
    lblSave.Enabled = False
End If

'Print...
If imgPrint.Enabled = False And txtNotes <> "" Then
    imgPrint.Enabled = True
    lblPrint.Enabled = True
ElseIf imgPrint.Enabled = True And txtNotes = "" Then
    imgPrint.Enabled = False
    lblPrint.Enabled = False
End If

'Clear...
If imgClear.Enabled = False And txtNotes <> "" Then
    imgClear.Enabled = True
    lblClear.Enabled = True
ElseIf imgClear.Enabled = True And txtNotes = "" Then
    imgClear.Enabled = False
    lblClear.Enabled = False
End If

End Sub
Private Sub txtNotes_Change()

QuickRef.NotesHaveChanged = True
QuickRef.CallingForm = "Notes"
iDirty = True

End Sub

Private Sub txtNotes_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

Help.HelpText = txtNotes

End Sub
