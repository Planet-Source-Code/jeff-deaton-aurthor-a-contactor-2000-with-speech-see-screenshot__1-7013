VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form frmVoiceMessages 
   BorderStyle     =   0  'None
   Caption         =   "Voice Messages"
   ClientHeight    =   6555
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10845
   Icon            =   "frmVoiceMessages.frx":0000
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmVoiceMessages.frx":08CA
   ScaleHeight     =   6555
   ScaleWidth      =   10845
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrLoadAllWaveFiles 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   9540
      Top             =   330
   End
   Begin VB.FileListBox File1 
      Height          =   285
      Left            =   8880
      Pattern         =   "*.Wav"
      TabIndex        =   9
      Top             =   1620
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.CheckBox chkAutoRewind 
      Height          =   195
      Left            =   4230
      TabIndex        =   7
      Top             =   3855
      Width           =   195
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   8970
      Top             =   330
   End
   Begin MCI.MMControl MM 
      Height          =   330
      Left            =   240
      TabIndex        =   0
      Top             =   3780
      Width           =   2730
      _ExtentX        =   4815
      _ExtentY        =   582
      _Version        =   393216
      BorderStyle     =   0
      PrevEnabled     =   -1  'True
      NextEnabled     =   -1  'True
      PlayEnabled     =   -1  'True
      PauseEnabled    =   -1  'True
      StopEnabled     =   -1  'True
      RecordEnabled   =   -1  'True
      EjectEnabled    =   -1  'True
      BackVisible     =   0   'False
      StepVisible     =   0   'False
      DeviceType      =   """WaveAudio"""
      FileName        =   ""
   End
   Begin VB.ListBox lstSoundFiles 
      Height          =   2790
      ItemData        =   "frmVoiceMessages.frx":6348C
      Left            =   270
      List            =   "frmVoiceMessages.frx":6348E
      Sorted          =   -1  'True
      TabIndex        =   4
      Top             =   810
      Width           =   5115
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Voice Messages"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   390
      TabIndex        =   10
      Top             =   60
      Width           =   1455
   End
   Begin VB.Label lblExistingSoundFiles 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFC0&
      Height          =   195
      Left            =   300
      TabIndex        =   5
      Tag             =   "Label"
      Top             =   555
      Width           =   45
   End
   Begin VB.Label lblAutoRewind 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Auto Rewind"
      ForeColor       =   &H00FFFFC0&
      Height          =   195
      Left            =   4500
      TabIndex        =   8
      Tag             =   "Label"
      Top             =   3855
      Width           =   915
   End
   Begin VB.Label lblExit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Exit"
      ForeColor       =   &H00C0FFC0&
      Height          =   195
      Left            =   6180
      TabIndex        =   6
      Tag             =   "ButtonLabel"
      Top             =   3840
      Width           =   255
   End
   Begin VB.Label lblDelete 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Delete"
      Enabled         =   0   'False
      ForeColor       =   &H00C0FFC0&
      Height          =   195
      Left            =   6075
      TabIndex        =   3
      Tag             =   "ButtonLabel"
      Top             =   3420
      Width           =   465
   End
   Begin VB.Label lblSave 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Save"
      Enabled         =   0   'False
      ForeColor       =   &H00C0FFC0&
      Height          =   195
      Left            =   6120
      TabIndex        =   2
      Tag             =   "ButtonLabel"
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label lblNew 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "New"
      ForeColor       =   &H00C0FFC0&
      Height          =   195
      Left            =   6135
      TabIndex        =   1
      Tag             =   "ButtonLabel"
      Top             =   2580
      Width           =   345
   End
   Begin VB.Image imgOKPicture 
      Height          =   360
      Index           =   0
      Left            =   8910
      Picture         =   "frmVoiceMessages.frx":63490
      Stretch         =   -1  'True
      Top             =   840
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.Image imgOKPicture 
      Height          =   375
      Index           =   1
      Left            =   8910
      Picture         =   "frmVoiceMessages.frx":64EB2
      Top             =   1230
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.Image imgSave 
      Enabled         =   0   'False
      Height          =   375
      Left            =   5760
      Picture         =   "frmVoiceMessages.frx":6659C
      Stretch         =   -1  'True
      Top             =   2910
      Width           =   1065
   End
   Begin VB.Image imgDelete 
      Enabled         =   0   'False
      Height          =   375
      Left            =   5760
      Picture         =   "frmVoiceMessages.frx":67FBE
      Stretch         =   -1  'True
      Top             =   3330
      Width           =   1065
   End
   Begin VB.Image imgNew 
      Height          =   375
      Left            =   5760
      Picture         =   "frmVoiceMessages.frx":699E0
      Stretch         =   -1  'True
      Top             =   2490
      Width           =   1065
   End
   Begin VB.Image imgExit 
      Height          =   375
      Left            =   5760
      Picture         =   "frmVoiceMessages.frx":6B402
      Stretch         =   -1  'True
      Top             =   3750
      Width           =   1065
   End
End
Attribute VB_Name = "frmVoiceMessages"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim iDirty As Boolean
Sub GetNextAvailableMMFileName()

On Local Error Resume Next

Dim x As Long

'Get the latest filename first...
File1.Path = App.Path & "\VM\" & QuickRef.ContactName
File1.Refresh

'Set the new filename for the MCI control...
MM.FileName = App.Path & "\VM\" & QuickRef.ContactName & "\" & QuickRef.ContactName & File1.ListCount + 1 & ".Wav"

End Sub
Function LoadAllWaveFiles() As Boolean

On Local Error Resume Next

Dim x As Long

'Add sound file to listbox...
lstSoundFiles.Clear
File1.Path = App.Path & "\VM\" & QuickRef.ContactName
File1.Refresh
For x = 0 To File1.ListCount - 1
    lstSoundFiles.AddItem File1.List(x)
Next x

LoadAllWaveFiles = File1.ListCount > 0

End Function
Private Sub Form_Load()

On Local Error Resume Next

'Set file1 pattern...
File1.Pattern = "*.Wav"

'Load this forms ini settings...
Call LoadINISettings

'Set program colors...
Call SetColors(Me)

'Setup the multimedia device...
Call SetupMultimediaControl("WaveAudio")

'Load all existing wave files for calling form...
Call LoadAllWaveFiles

'Display contact name...
lblExistingSoundFiles.Caption = QuickRef.ContactName

'Form Coordinates...
Me.Height = QuickRef.MediumMenuHeight
Me.Width = QuickRef.MediumMenuWidth

iDirty = False

End Sub
Sub LoadINISettings()

'Form Coordinates...
Me.Left = Val(ReadINI(Me.Name, "Left"))
Me.Top = Val(ReadINI(Me.Name, "Top"))

'Auto Rewind...
chkAutoRewind.Value = Val(ReadINI(Me.Name, "AutoRewind"))

End Sub
Function SaveChanges() As Boolean

On Local Error GoTo SaveChangesError

'Save the wave sound filename...
SaveSoundFile:
MM.Command = "Save"

'Re-load all wave files...
tmrLoadAllWaveFiles.Enabled = True

SaveChanges = True
iDirty = False
Exit Function



SaveChangesError:
    Call WriteToErrorLog(Me.Name, "SaveChanges", Error, Err, True)
    Exit Function

End Function
Sub SetupMultimediaControl(sDeviceType As String)

On Local Error Resume Next

'Setup the MCI control...
MM.Notify = False
MM.Wait = False
MM.Shareable = False
MM.DeviceType = sDeviceType

'Close any open wave audio devices first...
MM.Command = "Close"

'Make sure the subdirectory for this sound file exists...
If Dir$(App.Path & "\VM", vbDirectory) = "" Then
    MkDir App.Path & "\VM"
End If

'Make sure the subdirectory for this sound file exists...
If Dir$(App.Path & "\VM\" & QuickRef.ContactName, vbDirectory) = "" Then
    MkDir App.Path & "\VM\" & QuickRef.ContactName
End If

'Assign a filename if ones do not already exist...
File1.Path = App.Path & "\VM\" & QuickRef.ContactName
File1.Refresh
If File1.ListCount = 0 Then
    MM.FileName = App.Path & "\VM\" & QuickRef.ContactName & "\" & QuickRef.ContactName & "1.Wav"
End If

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

'Move the form if the user is pressing and holding the mouse button...
If Button = vbLeftButton Then
    Call DragForm(Me)
    If Help.HelpIsAligned = True And Help.HelpIsLoaded = True Then
        Call AlignHelpToForm
    End If
End If

End Sub
Private Sub Form_Unload(Cancel As Integer)

Dim x As Long

'Prompt to save first...
If iDirty Then
    x = MsgBox("Save changes before exiting?", vbYesNoCancel + vbQuestion, "Save Changes...")
    Select Case x
        Case vbYes
            If SaveChanges() = False Then
                If MsgBox("Changes were not saved. Do you still want to exit anyway?", vbYesNo + vbQuestion, "Save Changes...") = vbNo Then
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

'Unload the helper form if it is loaded...
If Help.HelpCallingForm = Me.Name Then
    Unload frmHelper
End If

End Sub
Sub SaveINISettings()

'Form coordinates...
Call WriteINI(Me.Name, "Left", Me.Left)
Call WriteINI(Me.Name, "Top", Me.Top)

'Auto Rewind...
Call WriteINI(Me.Name, "AutoRewind", chkAutoRewind.Value)

End Sub
Private Sub imgDelete_Click()

lblDelete_Click

End Sub
Private Sub imgDelete_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

If Button = vbLeftButton Then
    imgDelete.Picture = imgOKPicture(1).Picture
    lblDelete.ForeColor = QBColor(0)
End If

End Sub
Private Sub imgDelete_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

imgDelete.Picture = imgOKPicture(0).Picture
lblDelete.ForeColor = lButtonForeColor

End Sub
Private Sub imgExit_Click()

lblExit_Click

End Sub
Private Sub imgExit_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

If Button = vbLeftButton Then
    imgExit.Picture = imgOKPicture(1).Picture
    lblExit.ForeColor = QBColor(0)
End If

End Sub
Private Sub imgExit_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

imgExit.Picture = imgOKPicture(0).Picture
lblExit.ForeColor = lButtonForeColor

End Sub
Private Sub imgNew_Click()

lblNew_Click

End Sub
Private Sub imgNew_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

If Button = vbLeftButton Then
    imgNew.Picture = imgOKPicture(1).Picture
    lblNew.ForeColor = QBColor(0)
End If

End Sub
Private Sub imgNew_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

imgNew.Picture = imgOKPicture(0).Picture
lblNew.ForeColor = lButtonForeColor

End Sub

Private Sub imgSave_Click()

lblSave_Click

End Sub
Private Sub imgSave_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

If Button = vbLeftButton Then
    imgSave.Picture = imgOKPicture(1).Picture
    lblSave.ForeColor = QBColor(0)
End If

End Sub

Private Sub imgSave_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

imgSave.Picture = imgOKPicture(0).Picture
lblSave.ForeColor = lButtonForeColor

End Sub
Private Sub lblAutoRewind_Click()

'Toggle the autorewind check box...
If chkAutoRewind.Value = 0 Then
    chkAutoRewind.Value = 1
Else
    chkAutoRewind.Value = 0
End If

End Sub
Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

'Move the form if the user is pressing and holding the mouse button...
If Button = vbLeftButton Then
    Call DragForm(Me)
    If Help.HelpIsAligned = True And Help.HelpIsLoaded = True Then
        Call AlignHelpToForm
    End If
End If

End Sub

Private Sub lblDelete_Click()

On Local Error Resume Next

'Confirm...
If MsgBox("Are you sure you want to delete the Voice Message " & lstSoundFiles.List(lstSoundFiles.ListIndex) & "?", vbYesNo + vbQuestion, "Delete Voice Message...") = vbNo Then
    Exit Sub
End If

'Delete the file...
MM.Command = "Close"
Kill App.Path & "\VM\" & QuickRef.ContactName & "\" & lstSoundFiles.List(lstSoundFiles.ListIndex)
lstSoundFiles.RemoveItem lstSoundFiles.ListIndex

'Remove the directory if empty...
If Dir$(App.Path & "\VM\" & QuickRef.ContactName & "\*.Wav") = "" Then
    RmDir App.Path & "\VM\" & QuickRef.ContactName
End If

End Sub
Private Sub lblDelete_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

If Button = vbLeftButton Then
    imgDelete.Picture = imgOKPicture(1).Picture
    lblDelete.ForeColor = QBColor(0)
End If

End Sub

Private Sub lblDelete_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

imgDelete.Picture = imgOKPicture(0).Picture
lblDelete.ForeColor = lButtonForeColor

End Sub

Private Sub lblExistingSoundFiles_Click()

End Sub

Private Sub lblExistingSoundFiles_DblClick()

End Sub

Private Sub lblExit_Click()

Unload Me

End Sub
Private Sub lblExit_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

If Button = vbLeftButton Then
    imgExit.Picture = imgOKPicture(1).Picture
    lblExit.ForeColor = QBColor(0)
End If

End Sub
Private Sub lblExit_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

imgExit.Picture = imgOKPicture(0).Picture
lblExit.ForeColor = lButtonForeColor

End Sub
Private Sub lblNew_Click()

Dim x As Long

'Prompt to save first...
If iDirty Then
    x = MsgBox("Save changes before creating a new Voice Message?", vbYesNoCancel + vbQuestion, "Save Changes...")
    Select Case x
        Case vbYes
            If SaveChanges() = False Then
                If MsgBox("Changes were not saved. Do you still want to continue with creating a new Voice Message?", vbYesNo + vbQuestion, "Save Changes...") = vbNo Then
                    Exit Sub
                End If
            End If
        Case vbCancel
            Exit Sub
    End Select
End If

'Get next available filename...
Call GetNextAvailableMMFileName

'Start a new recording session...
MM.Command = "Close"
MM.Command = "Open"

'Tell them how to record and save...
If Dir$(App.Path & "\VM\" & QuickRef.ContactName & "\*.Wav") = "" Then
    MsgBox "You may begin recording now. Click Save when you are finished to save the voice message.", vbInformation, "New..."
End If

End Sub
Private Sub lblNew_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

If Button = vbLeftButton Then
    imgNew.Picture = imgOKPicture(1).Picture
    lblNew.ForeColor = QBColor(0)
End If

End Sub
Private Sub lblNew_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

imgNew.Picture = imgOKPicture(0).Picture
lblNew.ForeColor = lButtonForeColor

End Sub
Private Sub lblSave_Click()

Call SaveChanges

End Sub
Private Sub lblSave_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

If Button = vbLeftButton Then
    imgSave.Picture = imgOKPicture(1).Picture
    lblSave.ForeColor = QBColor(0)
End If

End Sub

Private Sub lblSave_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

imgSave.Picture = imgOKPicture(0).Picture
lblSave.ForeColor = lButtonForeColor

End Sub

Private Sub lstSoundFiles_Click()

Dim x As Long

'Prompt to save first...
If iDirty Then
    x = MsgBox("Save changes to your new recorded Voice Message first?", vbYesNoCancel + vbQuestion, "Save Changes...")
    Select Case x
        Case vbYes
            If SaveChanges() = False Then
                If MsgBox("Changes were not saved. Do you still want to continue?", vbYesNo + vbQuestion, "Save Changes...") = vbNo Then
                    Exit Sub
                End If
            End If
        Case vbCancel
            Exit Sub
    End Select
End If

'Retrieve the sound file name...
Call SetupMultimediaControl("WaveAudio")
MM.FileName = App.Path & "\VM\" & QuickRef.ContactName & "\" & lstSoundFiles.List(lstSoundFiles.ListIndex)
MM.Command = "Open"
MM.Command = "Play"

End Sub
Private Sub MM_PlayCompleted(Errorcode As Long)

'Rewind the track...
If chkAutoRewind.Value = 1 Then
    MM.Command = "Prev"
End If

End Sub

Private Sub MM_RecordClick(Cancel As Integer)

iDirty = True

End Sub
Private Sub MM_RecordCompleted(Errorcode As Long)

'Error...
If Errorcode > 0 Then
    MsgBox MM.ErrorMessage, vbInformation, "Recording Error..."
    Exit Sub
End If

iDirty = True

End Sub

Private Sub Timer1_Timer()

On Local Error Resume Next

'New...
If imgNew.Enabled = True And iDirty = True Then
    imgNew.Enabled = False
    lblNew.Enabled = False
ElseIf imgNew.Enabled = False And iDirty = False Then
    imgNew.Enabled = True
    lblNew.Enabled = True
End If

'Save...
If imgSave.Enabled = False And iDirty = True Then
    imgSave.Enabled = True
    lblSave.Enabled = True
ElseIf imgSave.Enabled = True And iDirty = False Then
    imgSave.Enabled = False
    lblSave.Enabled = False
End If

'Delete...
If imgDelete.Enabled = True And iDirty = True Or imgDelete.Enabled = True And lstSoundFiles.ListIndex = -1 Then
    imgDelete.Enabled = Not iDirty And lstSoundFiles.ListIndex > -1
    lblDelete.Enabled = Not iDirty And lstSoundFiles.ListIndex > -1
ElseIf imgDelete.Enabled = False And iDirty = False And lstSoundFiles.ListIndex > -1 Then
    imgDelete.Enabled = True
    lblDelete.Enabled = True
End If

End Sub

Private Sub tmrLoadAllWaveFiles_Timer()

'Re-load all wave files...
Call LoadAllWaveFiles

'Disable this timer control...
tmrLoadAllWaveFiles.Enabled = False

End Sub
