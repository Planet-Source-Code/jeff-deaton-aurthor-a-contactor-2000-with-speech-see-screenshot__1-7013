VERSION 5.00
Object = "{2398E321-5C6E-11D1-8C65-0060081841DE}#1.0#0"; "Vtext.dll"
Begin VB.Form frmAccounts 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "System Accounts"
   ClientHeight    =   5880
   ClientLeft      =   420
   ClientTop       =   0
   ClientWidth     =   11355
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmAccounts.frx":0000
   ScaleHeight     =   5880
   ScaleWidth      =   11355
   ShowInTaskbar   =   0   'False
   Begin HTTSLibCtl.TextToSpeech Voice 
      Height          =   405
      Left            =   9810
      OleObjectBlob   =   "frmAccounts.frx":62BC2
      TabIndex        =   18
      Top             =   1470
      Width           =   495
   End
   Begin VB.TextBox txtNickName 
      Height          =   285
      Left            =   4260
      TabIndex        =   4
      ToolTipText     =   "Users password"
      Top             =   1860
      Width           =   2445
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   10260
      Top             =   1020
   End
   Begin VB.CheckBox chkAdministrator 
      Caption         =   "Yes / No"
      Height          =   315
      Left            =   4260
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Gives the user administration priviledges"
      Top             =   2280
      Width           =   1185
   End
   Begin VB.TextBox txtPassWord 
      Height          =   285
      Left            =   4260
      TabIndex        =   3
      ToolTipText     =   "Users password"
      Top             =   1440
      Width           =   2445
   End
   Begin VB.TextBox txtFullName 
      Height          =   285
      Left            =   4260
      TabIndex        =   2
      ToolTipText     =   "Users full name"
      Top             =   1020
      Width           =   2445
   End
   Begin VB.TextBox txtLoginName 
      Height          =   285
      Left            =   4260
      TabIndex        =   1
      ToolTipText     =   "Users login name"
      Top             =   600
      Width           =   2445
   End
   Begin VB.ListBox lstUsers 
      Height          =   3180
      ItemData        =   "frmAccounts.frx":62BE6
      Left            =   300
      List            =   "frmAccounts.frx":62BE8
      TabIndex        =   0
      Top             =   900
      Width           =   2565
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   9810
      Top             =   1020
   End
   Begin VB.Label lblHelp 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Help..."
      ForeColor       =   &H00C0FFC0&
      Height          =   195
      Left            =   6090
      TabIndex        =   17
      Tag             =   "ButtonLabel"
      Top             =   3480
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nick Name"
      ForeColor       =   &H00FFFFC0&
      Height          =   195
      Index           =   5
      Left            =   3150
      TabIndex        =   16
      Tag             =   "Label"
      Top             =   1890
      Width           =   795
   End
   Begin VB.Label lblDelete 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Delete"
      ForeColor       =   &H00C0FFC0&
      Height          =   195
      Left            =   5130
      TabIndex        =   15
      Tag             =   "ButtonLabel"
      Top             =   3840
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Administrator"
      ForeColor       =   &H00FFFFC0&
      Height          =   195
      Index           =   3
      Left            =   3150
      TabIndex        =   14
      Tag             =   "Label"
      Top             =   2310
      Width           =   915
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      ForeColor       =   &H00FFFFC0&
      Height          =   195
      Index           =   2
      Left            =   3150
      TabIndex        =   13
      Tag             =   "Label"
      Top             =   1470
      Width           =   705
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Full Name"
      ForeColor       =   &H00FFFFC0&
      Height          =   195
      Index           =   1
      Left            =   3150
      TabIndex        =   12
      Tag             =   "Label"
      Top             =   1050
      Width           =   705
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Login Name"
      ForeColor       =   &H00FFFFC0&
      Height          =   195
      Index           =   0
      Left            =   3150
      TabIndex        =   11
      Tag             =   "Label"
      Top             =   630
      Width           =   855
   End
   Begin VB.Label lblUsers 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Users"
      ForeColor       =   &H00FFFFC0&
      Height          =   195
      Left            =   330
      TabIndex        =   10
      Tag             =   "Label"
      Top             =   630
      Width           =   405
   End
   Begin VB.Label lblSystemAccounts 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "System Accounts"
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
      Left            =   390
      TabIndex        =   9
      Top             =   60
      Width           =   1560
   End
   Begin VB.Image imgOKPicture 
      Height          =   375
      Index           =   1
      Left            =   9810
      Picture         =   "frmAccounts.frx":62BEA
      Top             =   630
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.Image imgOKPicture 
      Height          =   360
      Index           =   0
      Left            =   9810
      Picture         =   "frmAccounts.frx":642D4
      Stretch         =   -1  'True
      Top             =   240
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.Label lblNew 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "New"
      ForeColor       =   &H00C0FFC0&
      Height          =   195
      Left            =   3300
      TabIndex        =   8
      Tag             =   "ButtonLabel"
      Top             =   3840
      Width           =   345
   End
   Begin VB.Label lblSave 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Save"
      Enabled         =   0   'False
      ForeColor       =   &H00C0FFC0&
      Height          =   195
      Left            =   4260
      TabIndex        =   7
      Tag             =   "ButtonLabel"
      Top             =   3840
      Width           =   375
   End
   Begin VB.Label lblExit 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Exit"
      ForeColor       =   &H00C0FFC0&
      Height          =   195
      Left            =   6210
      TabIndex        =   6
      Tag             =   "ButtonLabel"
      Top             =   3840
      Width           =   255
   End
   Begin VB.Image imgExit 
      Height          =   375
      Left            =   5850
      Picture         =   "frmAccounts.frx":65CF6
      Stretch         =   -1  'True
      Top             =   3750
      Width           =   975
   End
   Begin VB.Image imgNew 
      Height          =   375
      Left            =   2970
      Picture         =   "frmAccounts.frx":67718
      Stretch         =   -1  'True
      Top             =   3750
      Width           =   975
   End
   Begin VB.Image imgSave 
      Enabled         =   0   'False
      Height          =   375
      Left            =   3930
      Picture         =   "frmAccounts.frx":6913A
      Stretch         =   -1  'True
      Top             =   3750
      Width           =   975
   End
   Begin VB.Image imgDelete 
      Height          =   375
      Left            =   4890
      Picture         =   "frmAccounts.frx":6AB5C
      Stretch         =   -1  'True
      Top             =   3750
      Width           =   975
   End
   Begin VB.Image imgHelp 
      Height          =   375
      Left            =   5850
      Picture         =   "frmAccounts.frx":6C57E
      Stretch         =   -1  'True
      Top             =   3390
      Width           =   975
   End
End
Attribute VB_Name = "frmAccounts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim iDirty As Boolean

Sub ClearAllControls()

On Local Error Resume Next

txtLoginName = ""
txtFullName = ""
txtPassWord = ""
txtNickName = ""
chkAdministrator.Value = 0

iDirty = False

End Sub
Sub DeleteAccount()

On Local Error GoTo DeleteAccountError

'Confirm...
If MsgBox("Are you sure you want to delete this account?", vbYesNo + vbQuestion, "Delete Account...") = vbNo Then
    Exit Sub
End If

If OpenDB(DB, QuickRef.DBPassWord) = False Then Exit Sub
Set RS = DB.OpenRecordset("SELECT * FROM tblLogin WHERE LoginName = '" & txtLoginName & "'", dbOpenDynaset)

'No record was found for username ???...
If RS.RecordCount = 0 Then
    MsgBox "For some reason, I can't find this users account. This should not happen! Just proceed without deleting this account for now.", vbInformation, "Account Not Found..."
    Exit Sub
End If

Do
    RS.Delete
    RS.MoveNext
Loop Until RS.EOF

'Clear all controls...
Call ClearAllControls

RS.Close
DB.Close

'Re-Load all system accounts...
If LoadAllSystemAccounts() = False Then
    MsgBox "An unexpected error occured while trying to load the system accounts. Unable to continue.", vbCritical, "Load System Accounts..."
    Timer2.Enabled = True
    Exit Sub
End If

Exit Sub



DeleteAccountError:
    DB.Close
    Call WriteToErrorLog(Me.Name, "DeleteAccountError", Error, Err, True)
    Exit Sub

End Sub
Function LoadAllSystemAccounts() As Boolean

On Local Error GoTo LoadAllSystemAccountsError

Dim x As Long

'Query the database and see if user exists...
If OpenDB(DB, QuickRef.DBPassWord) = False Then Exit Function
Set RS = DB.OpenRecordset("SELECT * FROM tblLogin", dbOpenSnapshot)

'No info found for contact ???...
If RS.RecordCount = 0 Then
    RS.Close
    DB.Close
    Exit Function
End If

'Load all users...
lstUsers.Clear
Do
    lstUsers.AddItem RS!LoginName
    RS.MoveNext
Loop Until RS.EOF

RS.Close
DB.Close

'Set the listindex...
lstUsers.Enabled = True
If txtLoginName <> "" Then
    For x = 0 To lstUsers.ListCount - 1
        If LCase$(lstUsers.List(x)) = LCase$(txtLoginName) Then
            lstUsers.ListIndex = x
            Exit For
        End If
    Next x
ElseIf lstUsers.ListCount > 0 Then
    Call ClearAllControls
    lstUsers.ListIndex = 0
End If

LoadAllSystemAccounts = True
Exit Function



LoadAllSystemAccountsError:
    DB.Close
    Call WriteToErrorLog(Me.Name, "LoadAllSystemAccountsError", Error, Err, True)
    Exit Function

End Function
Function SaveChanges() As Boolean

On Local Error GoTo SaveChangesError

'Check for errors before saving...
If NoAccountErrors() = False Then
    Exit Function
End If

'Query the database and see if user exists...
If OpenDB(DB, QuickRef.DBPassWord) = False Then Exit Function
Set RS = DB.OpenRecordset("SELECT * FROM tblLogin WHERE LoginName = '" & lstUsers & "'", dbOpenDynaset)

'No info found for contact ???...
If RS.RecordCount > 0 Then
    RS.Edit
Else
    RS.AddNew
End If

'Update the system accounts...
RS!LoginName = txtLoginName
RS!FullName = txtFullName
RS!Password = txtPassWord
RS!LoginDateTime = Now
RS!NickName = txtNickName
RS!Administrator = chkAdministrator.Value = 1

RS.Update
RS.Close
DB.Close
SaveChanges = True
iDirty = False

'Load all system accounts...
If LoadAllSystemAccounts() = False Then
    MsgBox "An unexpected error occured while trying to re-load the system account information. Unable to continue!!!", vbCritical, "Re-Load System Accounts..."
    Timer2.Enabled = True
    Exit Function
End If

Exit Function



SaveChangesError:
    DB.Close
    Call WriteToErrorLog(Me.Name, "SaveChangesError", Error, Err, True)
    Exit Function
    Resume Next

End Function
Function NoAccountErrors() As Boolean

On Local Error GoTo NoAccountErrorsError

'Login Name...
If Trim$(txtLoginName) = "" Then
    MsgBox "No login name was specified.", vbInformation, "Error..."
    txtLoginName.SetFocus
    Exit Function
End If

'Full Name...
If Trim$(txtFullName) = "" Then
    MsgBox "No full name was specified.", vbInformation, "Error..."
    txtFullName.SetFocus
    Exit Function
End If

'PassWord...
If Trim$(txtPassWord) = "" Then
    MsgBox "No password was specified.", vbInformation, "Error..."
    txtPassWord.SetFocus
    Exit Function
End If

'NickName...
If Trim$(txtNickName) = "" Then
    If MsgBox("No Nickname was specified for this account. When in the Chat screen, this user's login name will be used. Would you like to enter a nickname for this user now?", vbYesNo + vbQuestion, "No Nickname...") = vbYes Then
        txtNickName.SetFocus
        Exit Function
    Else
        txtNickName = txtLoginName
    End If
End If

NoAccountErrors = True
Exit Function



NoAccountErrorsError:
    Call WriteToErrorLog(Me.Name, "NoAccountErrorsError", Error, Err, False)
    Exit Function

End Function
Sub UserIsAdminDisableAllControls(iIsAdmin As Boolean)

'Disallow any changes to the admin account except the password...

On Local Error GoTo UserIsAdminDisableAllControlsError

If iIsAdmin = True Then
    txtLoginName.Enabled = False
    txtFullName.Enabled = False
    txtPassWord.Enabled = True
    txtNickName.Enabled = True
    chkAdministrator.Enabled = False
Else
    txtLoginName.Enabled = True
    txtFullName.Enabled = True
    txtPassWord.Enabled = True
    txtNickName.Enabled = True
    chkAdministrator.Enabled = True
End If

Exit Sub



UserIsAdminDisableAllControlsError:
    Call WriteToErrorLog(Me.Name, "UserIsAdminDisableAllControlsError", Error$, Err, False)
    Exit Sub

End Sub
Private Sub chkAdministrator_Click()

iDirty = True

If lstUsers.Enabled And lstUsers.Visible Then
    lstUsers.SetFocus
Else
    If txtLoginName.Visible Then
        txtLoginName.SetFocus
    End If
End If

End Sub

Private Sub chkAdministrator_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

Help.HelpText = "Sets whether or not this user has administrator priviledges. If so, this user can access areas of the program that others can not."

End Sub

Private Sub Form_Load()

'Load the main menu's form settings...
Call LoadINISettings

'Set program colors...
Call SetColors(Me)

'Load all system accounts...
If LoadAllSystemAccounts() = False Then
    MsgBox "An unexpected error occured while trying to load the system accounts. Unable to continue.", vbCritical, "Load System Accounts..."
    Timer2.Enabled = True
    Exit Sub
End If

'Form Coordinates...
Me.Width = QuickRef.MediumMenuWidth
Me.Height = QuickRef.MediumMenuHeight

iDirty = False

End Sub
Sub LoadINISettings()

'Form Coordinates...
Me.Left = Val(ReadINI(Me.Name, "Left"))
Me.Top = Val(ReadINI(Me.Name, "Top"))

End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

'Move the form if the user is pressing and holding the mouse button...
If Button = vbLeftButton Then
    Call DragForm(Me)
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

End Sub
Sub SaveINISettings()

'Form coordinates...
Call WriteINI(Me.Name, "Left", Me.Left)
Call WriteINI(Me.Name, "Top", Me.Top)

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

Private Sub imgHelp_Click()

lblHelp_Click

End Sub
Private Sub imgHelp_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

If Button = vbLeftButton Then
    imgHelp.Picture = imgOKPicture(1).Picture
    lblHelp.ForeColor = QBColor(0)
End If

End Sub
Private Sub imgHelp_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

imgHelp.Picture = imgOKPicture(0).Picture
lblHelp.ForeColor = lButtonForeColor

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

Private Sub lblDelete_Click()

'Delete account...
Call DeleteAccount

End Sub
Private Sub lblDelete_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

If Button = vbLeftButton Then
    imgDelete.Picture = imgOKPicture(1).Picture
    lblDelete.ForeColor = QBColor(0)
End If

End Sub

Private Sub lblDelete_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

Help.HelpText = "Deletes this user from the system."

End Sub
Private Sub lblDelete_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

imgDelete.Picture = imgOKPicture(0).Picture
lblDelete.ForeColor = lButtonForeColor

End Sub

Private Sub lblExit_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

Help.HelpText = "Exit this screen."

End Sub

Private Sub lblHelp_Click()

Help.HelpCallingForm = Me.Name

frmHelper.Show
frmHelper.ZOrder

End Sub
Private Sub lblHelp_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

If Button = vbLeftButton Then
    imgHelp.Picture = imgOKPicture(1).Picture
    lblHelp.ForeColor = QBColor(0)
End If

End Sub

Private Sub lblHelp_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

Help.HelpText = "Shows the Help Window."

End Sub
Private Sub lblHelp_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

imgHelp.Picture = imgOKPicture(0).Picture
lblHelp.ForeColor = lButtonForeColor

End Sub
Private Sub lblNew_Click()

On Local Error GoTo lblNew_ClickError

Dim x As Long
Dim sInput As String

'Enter a new user name...
EnterNewUserName:
sInput = Trim$(InputBox$("Enter the user's login name for this new account.", "New Account..."))
If sInput = "" Then Exit Sub

If OpenDB(DB, QuickRef.DBPassWord) = False Then Exit Sub
Set RS = DB.OpenRecordset("SELECT * FROM tblLogin WHERE LoginName = '" & sInput & "'", dbOpenSnapshot)

'Create a new account...
If RS.RecordCount > 0 Then
    MsgBox "This user name already exists. Please select a unique user name.", vbInformation, "Account Already Exists..."
    GoTo EnterNewUserName
End If

'Select all from tblLogin to create a new account...
Set RS = DB.OpenRecordset("SELECT * FROM tblLogin", dbOpenDynaset)

'Clear out all controls...
Call ClearAllControls
txtLoginName = sInput

'Add the new account...
RS.AddNew
RS!LoginName = sInput
RS!FullName = txtFullName
RS!Password = txtPassWord
RS!NickName = txtNickName
RS!Administrator = chkAdministrator.Value = 1
RS.Update
RS.Close
DB.Close

'Re-Load all system accounts...
If LoadAllSystemAccounts() = False Then
    MsgBox "An unexpected error occured while trying to load the system accounts. Unable to continue.", vbCritical, "Load System Accounts..."
    Timer2.Enabled = True
    Exit Sub
End If

'Find the newly created entry...
For x = 0 To lstUsers.ListCount - 1
    If lstUsers.List(x) = sInput Then
        lstUsers.ListIndex = x
        Exit For
    End If
Next x

txtFullName.SetFocus
iDirty = False
Exit Sub



lblNew_ClickError:
    DB.Close
    Call WriteToErrorLog(Me.Name, "lblNew_ClickError", Error$, Err, True)
    Exit Sub
    Resume Next

End Sub

Private Sub lblNew_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

Help.HelpText = "Click here to create a new account."

End Sub
Private Sub lblSave_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

Help.HelpText = "Saves any changes you have made to this account to the database."

End Sub
Private Sub lblSystemAccounts_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

'Move the form if the user is pressing and holding the mouse button...
If Button = vbLeftButton Then
    Call DragForm(Me)
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
Private Sub lblExit_Click()

'Unload the help window...
If Help.HelpCallingForm = Me.Name Then
    Unload frmHelper
End If

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
Private Sub lstUsers_Click()

On Local Error Resume Next

Dim iLocalIPHasChanged As Boolean

'Query the database and see if user exists...
If OpenDB(DB, QuickRef.DBPassWord) = False Then Exit Sub
Set RS = DB.OpenRecordset("SELECT * FROM tblLogin WHERE LoginName = '" & lstUsers & "'", dbOpenSnapshot)

'No info found for contact ???...
If RS.RecordCount = 0 Then
    RS.Close
    DB.Close
    Exit Sub
End If

'Clear out all of the fields...
Call ClearAllControls

'Populate fields...
txtLoginName = RS!LoginName
txtFullName = RS!FullName
txtPassWord = RS!Password
txtNickName = RS!NickName

If RS!Administrator = True Then
    chkAdministrator.Value = 1
Else
    chkAdministrator.Value = 0
End If

RS.Close
DB.Close
iDirty = (iLocalIPHasChanged = True)

End Sub

Private Sub lstUsers_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

Help.HelpText = "Listing of all users currently set up in the system."

End Sub
Private Sub Timer1_Timer()

On Local Error Resume Next

'Disable if its the administrator account...
If lstUsers.List(lstUsers.ListIndex) = "Administrator" Then
    Call UserIsAdminDisableAllControls(True)
Else
    Call UserIsAdminDisableAllControls(False)
End If

'Users Listbox...
lstUsers.Enabled = iDirty = False

'New...
If imgNew.Enabled = False And iDirty = False Then
    imgNew.Enabled = True
    lblNew.Enabled = True
ElseIf imgNew.Enabled = True And iDirty = True Then
    imgNew.Enabled = False
    lblNew.Enabled = False
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
If imgDelete.Enabled = True And lstUsers.List(lstUsers.ListIndex) = "Administrator" And iDirty = False Then
    imgDelete.Enabled = False
    lblDelete.Enabled = False
ElseIf imgDelete.Enabled = False And lstUsers.List(lstUsers.ListIndex) <> "Administrator" And iDirty = False Then
    imgDelete.Enabled = True
    lblDelete.Enabled = True
End If

End Sub
Private Sub Timer2_Timer()

'Unload the help window...
If Help.HelpCallingForm = Me.Name Then
    Unload frmHelper
End If

Unload Me

End Sub
Private Sub txtFullName_Change()

iDirty = True

End Sub

Private Sub txtFullName_GotFocus()

txtFullName.SelStart = 0
txtFullName.SelLength = Len(txtFullName)

End Sub

Private Sub txtFullName_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

Help.HelpText = "Type in this users full name here."

End Sub
Private Sub txtLoginName_Change()

iDirty = True

End Sub
Private Sub txtLoginName_GotFocus()

txtLoginName.SelStart = 0
txtLoginName.SelLength = Len(txtLoginName)

End Sub

Private Sub txtLoginName_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

Help.HelpText = "Type in the users login name here."

End Sub

Private Sub txtNickName_Change()

iDirty = True

End Sub

Private Sub txtNickName_GotFocus()

txtNickName.SelStart = 0
txtNickName.SelLength = Len(txtNickName)

End Sub

Private Sub txtNickName_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

Help.HelpText = "This field is optional. You can assign a nick name for this user."

End Sub


Private Sub txtPassWord_Change()

iDirty = True

End Sub

Private Sub txtPassWord_GotFocus()

txtPassWord.SelStart = 0
txtPassWord.SelLength = Len(txtPassWord)

End Sub

Private Sub txtPassWord_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

Help.HelpText = "Type in a password for this user to log in with."

End Sub


