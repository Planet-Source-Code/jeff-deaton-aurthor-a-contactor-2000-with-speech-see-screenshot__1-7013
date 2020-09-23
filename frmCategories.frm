VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCategories 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Options"
   ClientHeight    =   7710
   ClientLeft      =   420
   ClientTop       =   0
   ClientWidth     =   13560
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmCategories.frx":0000
   ScaleHeight     =   7710
   ScaleWidth      =   13560
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   11040
      Top             =   990
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCategories.frx":B3412
            Key             =   "RED"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCategories.frx":B3CEE
            Key             =   "BLUE"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvwCategories 
      Height          =   4575
      Left            =   300
      TabIndex        =   9
      Top             =   870
      Width           =   3795
      _ExtentX        =   6694
      _ExtentY        =   8070
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   4
      Sorted          =   -1  'True
      Style           =   7
      FullRowSelect   =   -1  'True
      ImageList       =   "ImageList1"
      Appearance      =   1
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   10560
      Top             =   960
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   10110
      Top             =   960
   End
   Begin MSComctlLib.TreeView tvwApplicants 
      Height          =   4575
      Left            =   4230
      TabIndex        =   10
      Top             =   870
      Width           =   3795
      _ExtentX        =   6694
      _ExtentY        =   8070
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   4
      LabelEdit       =   1
      Sorted          =   -1  'True
      Style           =   7
      FullRowSelect   =   -1  'True
      ImageList       =   "ImageList1"
      Appearance      =   1
   End
   Begin VB.Label lblApplicant 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Applicant:"
      ForeColor       =   &H00FFFFC0&
      Height          =   195
      Left            =   4260
      TabIndex        =   12
      Tag             =   "Label"
      Top             =   600
      Width           =   705
   End
   Begin VB.Label lblCategory 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Category:"
      ForeColor       =   &H00FFFFC0&
      Height          =   195
      Left            =   330
      TabIndex        =   11
      Tag             =   "Label"
      Top             =   600
      Width           =   675
   End
   Begin VB.Label lblExit 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Exit"
      ForeColor       =   &H00C0FFC0&
      Height          =   195
      Left            =   8700
      TabIndex        =   8
      Tag             =   "ButtonLabel"
      ToolTipText     =   "Exit the categories screen"
      Top             =   5190
      Width           =   255
   End
   Begin VB.Label lblHelp 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Help"
      ForeColor       =   &H00C0FFC0&
      Height          =   195
      Left            =   8655
      TabIndex        =   7
      Tag             =   "ButtonLabel"
      Top             =   4800
      Width           =   330
   End
   Begin VB.Label lblMoveTo 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Move"
      ForeColor       =   &H00C0FFC0&
      Height          =   195
      Left            =   8640
      TabIndex        =   6
      Tag             =   "ButtonLabel"
      ToolTipText     =   "Move this applicant to another category"
      Top             =   4410
      Width           =   405
   End
   Begin VB.Image imgButton 
      Height          =   375
      Index           =   1
      Left            =   9930
      Picture         =   "frmCategories.frx":B45CA
      Stretch         =   -1  'True
      Top             =   540
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.Image imgButton 
      Height          =   360
      Index           =   0
      Left            =   9930
      Picture         =   "frmCategories.frx":B5CB4
      Stretch         =   -1  'True
      Top             =   150
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.Label lblReload 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Reload"
      ForeColor       =   &H00C0FFC0&
      Height          =   195
      Left            =   8580
      TabIndex        =   5
      Tag             =   "ButtonLabel"
      ToolTipText     =   "Lose changes and reload all categories"
      Top             =   4020
      Width           =   525
   End
   Begin VB.Label lblEdit 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Rename"
      ForeColor       =   &H00C0FFC0&
      Height          =   195
      Left            =   8520
      TabIndex        =   4
      Tag             =   "ButtonLabel"
      ToolTipText     =   "Edit this category"
      Top             =   3630
      Width           =   600
   End
   Begin VB.Label lblCategories 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Categories"
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
      Index           =   0
      Left            =   420
      TabIndex        =   3
      Top             =   60
      Width           =   945
   End
   Begin VB.Label lblNew 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "New"
      ForeColor       =   &H00C0FFC0&
      Height          =   195
      Left            =   8670
      TabIndex        =   2
      Tag             =   "ButtonLabel"
      ToolTipText     =   "Create a new category"
      Top             =   2460
      Width           =   345
   End
   Begin VB.Label lblSave 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Save"
      ForeColor       =   &H00C0FFC0&
      Height          =   195
      Left            =   8640
      TabIndex        =   1
      Tag             =   "ButtonLabel"
      ToolTipText     =   "Save this category"
      Top             =   2850
      Width           =   375
   End
   Begin VB.Label lblDelete 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Delete"
      ForeColor       =   &H00C0FFC0&
      Height          =   195
      Left            =   8610
      TabIndex        =   0
      Tag             =   "ButtonLabel"
      ToolTipText     =   "Delete this category"
      Top             =   3240
      Width           =   465
   End
   Begin VB.Image imgExit 
      Height          =   375
      Left            =   8190
      Picture         =   "frmCategories.frx":B76D6
      Stretch         =   -1  'True
      Top             =   5100
      Width           =   1245
   End
   Begin VB.Image imgNew 
      Height          =   375
      Left            =   8190
      Picture         =   "frmCategories.frx":B90F8
      Stretch         =   -1  'True
      Top             =   2370
      Width           =   1245
   End
   Begin VB.Image imgDelete 
      Height          =   375
      Left            =   8190
      Picture         =   "frmCategories.frx":BAB1A
      Stretch         =   -1  'True
      Top             =   3150
      Width           =   1245
   End
   Begin VB.Image imgSave 
      Height          =   375
      Left            =   8190
      Picture         =   "frmCategories.frx":BC53C
      Stretch         =   -1  'True
      Top             =   2760
      Width           =   1245
   End
   Begin VB.Image imgEdit 
      Height          =   375
      Left            =   8190
      Picture         =   "frmCategories.frx":BDF5E
      Stretch         =   -1  'True
      Top             =   3540
      Width           =   1245
   End
   Begin VB.Image imgReload 
      Height          =   375
      Left            =   8190
      Picture         =   "frmCategories.frx":BF980
      Stretch         =   -1  'True
      Top             =   3930
      Width           =   1245
   End
   Begin VB.Image imgMoveTo 
      Height          =   375
      Left            =   8190
      Picture         =   "frmCategories.frx":C13A2
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   1245
   End
   Begin VB.Image imgHelp 
      Height          =   375
      Left            =   8190
      Picture         =   "frmCategories.frx":C2DC4
      Stretch         =   -1  'True
      Top             =   4710
      Width           =   1245
   End
End
Attribute VB_Name = "frmCategories"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim iDirty As Boolean
Dim sTreeViewID As String
Sub CreateNewApplicant()

On Local Error GoTo CreateNewApplicantError

Dim x As Long
Dim sInput As String
Dim DB As Database
Dim RS As Recordset

'Get the new name of the Applicant to create...
CreateNewApplicantName:
sInput = Trim$(InputBox$("Enter the name of this new Applicant.", "New Applicant..."))

'Check for invalid characters...
If InStr(sInput, "/") > 0 Or InStr(sInput, "\") > 0 Or InStr(sInput, "#") > 0 Then
    MsgBox "Invalid character detected in the Applicant name. Please use only letters for Applicant names.", vbInformation, "Applicant Name..."
    GoTo CreateNewApplicantName
End If

'Check for duplicates...
For x = 1 To tvwApplicants.Nodes.Count
    If LCase$(tvwApplicants.Nodes(x).Text) = LCase$(sInput) Then
        MsgBox "This Applicant already exists.", vbInformation, "New Applicant..."
        Exit Sub
    End If
Next x

'Check for nothing entered...
If sInput = "" Then Exit Sub

If OpenDB(DB, QuickRef.DBPassWord) = False Then Exit Sub
Set RS = DB.OpenRecordset("SELECT * FROM tblApplicants", dbOpenDynaset)

'Create the new Applicant...
RS.AddNew
RS!ApplicantName = sInput
RS!ApplicantID = GetNextApplicantID()
If Val(tvwCategories.SelectedItem.Tag) > 0 Then
    RS!CategoryID = Val(tvwCategories.SelectedItem.Tag)
Else
    RS!CategoryID = 999
End If
RS.Update
RS.Close
DB.Close

'Reload applicants...
Call tvwCategories_NodeClick(tvwCategories.SelectedItem)

Exit Sub



CreateNewApplicantError:
    Call WriteToErrorLog(Me.Name, "CreateNewApplicantError", Error$, Err, True)
    Exit Sub

End Sub
Sub CreateNewCategory()

On Local Error GoTo CreateNewCategoryError

Dim x As Long
Dim sInput As String
Dim DB As Database
Dim RS As Recordset

'Get the new name of the category to create...
CreateNewCategoryName:
sInput = Trim$(InputBox$("Enter the name of this new Category.", "New Category..."))

'Check for invalid characters...
If InStr(sInput, "/") > 0 Or InStr(sInput, "\") > 0 Or InStr(sInput, "#") > 0 Then
    MsgBox "Invalid character detected in the Category name. Please use only letters and/or numbers for Category names.", vbInformation, "Category Name..."
    GoTo CreateNewCategoryName
End If

'Check for duplicates...
For x = 1 To tvwCategories.Nodes.Count
    If LCase$(tvwCategories.Nodes(x).Text) = LCase$(sInput) Then
        MsgBox "This category already exists.", vbInformation, "Categories..."
        Exit Sub
    End If
Next x

'Check for nothing entered...
If sInput = "" Then Exit Sub

If OpenDB(DB, QuickRef.DBPassWord) = False Then Exit Sub
Set RS = DB.OpenRecordset("SELECT * FROM tblCategories", dbOpenDynaset)

'Create the new category...
RS.AddNew
RS!CategoryName = sInput
RS!CategoryID = 999
RS.Update
RS.Close
DB.Close

'Load all existing categories...
If LoadAllCategories(tvwCategories) = False Then
    MsgBox "Unable to load the Categories. This is an unexpected error and also critical. After restarting the program, and you keep getting this error, please contact " & Help.TechnicalSupportCompany & " at " & Help.TechnicalSupportPhone & " to help you resolve the problem.", vbCritical, "Categories Error..."
    Timer2.Enabled = True
End If

Exit Sub



CreateNewCategoryError:
    Call WriteToErrorLog(Me.Name, "CreateNewCategoryError", Error$, Err, True)
    Exit Sub

End Sub
Sub DeleteApplicant()

On Local Error GoTo DeleteApplicantError

Dim DB As Database
Dim RS As Recordset

'Nothing to delete...
If tvwApplicants.SelectedItem = "Applicants" Then Exit Sub

'Confirm...
If MsgBox("Are you sure you want to delete this Applicant?", vbYesNo + vbQuestion, "Delete Applicant...") = vbNo Then
    Exit Sub
End If

If OpenDB(DB, QuickRef.DBPassWord) = False Then Exit Sub
Set RS = DB.OpenRecordset("SELECT * FROM tblApplicants WHERE ApplicantName = '" & tvwApplicants.SelectedItem & "'", dbOpenDynaset)

'Delete the Applicant...
If RS.RecordCount > 0 Then
    Do
        RS.Delete
        RS.MoveNext
    Loop Until RS.EOF
End If

RS.Close
DB.Close
Call tvwCategories_NodeClick(tvwCategories.SelectedItem)
Exit Sub



DeleteApplicantError:
    Call WriteToErrorLog(Me.Name, "DeleteApplicantError", Error$, Err, True)
    Exit Sub

End Sub
Sub DeleteCategory()

On Local Error GoTo DeleteCategoryError

Dim DB As Database
Dim RS As Recordset

'Nothing to delete...
If tvwCategories.SelectedItem = "Categories" Then Exit Sub

'Confirm...
If MsgBox("Are you sure you want to delete this category?", vbYesNo + vbQuestion, "Delete Category...") = vbNo Then
    Exit Sub
End If

If OpenDB(DB, QuickRef.DBPassWord) = False Then Exit Sub
Set RS = DB.OpenRecordset("SELECT * FROM tblCategories WHERE CategoryName = '" & tvwCategories.SelectedItem & "' AND CategoryID = " & Val(tvwCategories.SelectedItem.Tag), dbOpenDynaset)

'Delete the category...
If RS.RecordCount > 0 Then
    Do
        RS.Delete
        RS.MoveNext
    Loop Until RS.EOF
End If

RS.Close
DB.Close

'Load all existing categories...
If LoadAllCategories(tvwCategories) = False Then
    MsgBox "Unable to load the Categories. This is an unexpected error and also critical. After restarting the program, and you keep getting this error, please contact " & Help.TechnicalSupportCompany & " at " & Help.TechnicalSupportPhone & " to help you resolve the problem.", vbCritical, "Categories Error..."
    Timer2.Enabled = True
End If

Exit Sub



DeleteCategoryError:
    Call WriteToErrorLog(Me.Name, "DeleteCategoryError", Error$, Err, True)
    Exit Sub

End Sub
Sub RenameApplicant()

On Local Error GoTo RenameApplicantError



















RenameApplicantError:
    Call WriteToErrorLog(Me.Name, "RenameApplicantError", Error$, Err, True)
    Exit Sub

End Sub
Sub RenameCategory()

On Local Error GoTo RenameCategoryError



















RenameCategoryError:
    Call WriteToErrorLog(Me.Name, "RenameCategoryError", Error$, Err, True)
    Exit Sub

End Sub
Function SaveChanges() As Boolean

On Local Error GoTo SaveChangesError

Dim x As Long
Dim DB As Database
Dim RS As Recordset

'Save the categories...
If OpenDB(DB, QuickRef.DBPassWord) = False Then Exit Function
Set RS = DB.OpenRecordset("SELECT * FROM tblCategories", dbOpenDynaset)
If RS.RecordCount > 0 Then
    Do
        RS.Delete
        RS.MoveNext
    Loop Until RS.EOF
End If
For x = 2 To tvwCategories.Nodes.Count
    RS.AddNew
    RS!CategoryName = tvwCategories.Nodes(x).Text
    RS!CategoryID = Val(tvwCategories.Nodes(x).Tag)
    RS.Update
Next x

iDirty = False
SaveChanges = True
Exit Function



SaveChangesError:
    DB.Close
    Call WriteToErrorLog(Me.Name, "SaveChangesError", Error, Err, True)
    Exit Function

End Function
Private Sub Form_Load()

'Load the main menu's form settings...
Call LoadINISettings

'Set program colors...
Call SetColors(Me)

'Load all existing categories...
If LoadAllCategories(tvwCategories) = False Then
    MsgBox "Unable to load the Categories. This is an unexpected error and also critical. After restarting the program, and you keep getting this error, please contact " & Help.TechnicalSupportCompany & " at " & Help.TechnicalSupportPhone & " to help you resolve the problem.", vbCritical, "Categories Error..."
    Timer2.Enabled = True
End If

'Set Width and Height...
Me.Width = 9650
Me.Height = 5700

'Set user menu permissions...
Call SetUserMenuPermissions(Me)

'Set the dirty flag to false...
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
    If Help.HelpIsAligned = True And Help.HelpIsLoaded = True Then
        Call AlignHelpToForm
    End If
End If

End Sub
Private Sub Form_Unload(Cancel As Integer)

'Save INI Settings...
Call SaveINISettings

'Unload the helper form if it is loaded...
If Help.HelpCallingForm = Me.Name Then
    Unload frmHelper
End If

End Sub
Sub SaveINISettings()

'Left and top properties...
Call WriteINI(Me.Name, "Left", Me.Left)
Call WriteINI(Me.Name, "Top", Me.Top)

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

Private Sub imgMoveTo_Click()

lblMoveTo_Click

End Sub
Private Sub imgMoveTo_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

If Button = vbLeftButton Then
    imgMoveTo.Picture = imgButton(1).Picture
    lblMoveTo.ForeColor = QBColor(0)
End If

End Sub
Private Sub imgMoveTo_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

imgMoveTo.Picture = imgButton(0).Picture
lblMoveTo.ForeColor = lButtonForeColor

End Sub
Private Sub imgDelete_Click()

lblDelete_Click

End Sub
Private Sub imgDelete_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

If Button = vbLeftButton Then
    imgDelete.Picture = imgButton(1).Picture
    lblDelete.ForeColor = QBColor(0)
End If

End Sub

Private Sub imgDelete_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

imgDelete.Picture = imgButton(0).Picture
lblDelete.ForeColor = lButtonForeColor

End Sub

Private Sub imgEdit_Click()

lblEdit_Click

End Sub
Private Sub imgEdit_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

If Button = vbLeftButton Then
    imgEdit.Picture = imgButton(1).Picture
    lblEdit.ForeColor = QBColor(0)
End If

End Sub

Private Sub imgEdit_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

imgEdit.Picture = imgButton(0).Picture
lblEdit.ForeColor = lButtonForeColor

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
Private Sub imgNew_Click()

lblNew_Click

End Sub
Private Sub imgNew_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

If Button = vbLeftButton Then
    imgNew.Picture = imgButton(1).Picture
    lblNew.ForeColor = QBColor(0)
End If

End Sub

Private Sub imgNew_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

imgNew.Picture = imgButton(0).Picture
lblNew.ForeColor = lButtonForeColor

End Sub

Private Sub imgReload_Click()

lblReload_Click

End Sub
Private Sub imgReload_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

If Button = vbLeftButton Then
    imgReload.Picture = imgButton(1).Picture
    lblReload.ForeColor = QBColor(0)
End If

End Sub

Private Sub imgReload_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

imgReload.Picture = imgButton(0).Picture
lblReload.ForeColor = lButtonForeColor

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

Private Sub lblDelete_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

If sTreeViewID = "Categories" Then
    Help.HelpText = "Deletes the selected category."
Else
    Help.HelpText = "Deletes the selected applicant."
End If

End Sub

Private Sub lblEdit_Click()

'Rename...
Select Case sTreeViewID

    'Rename Category...
    Case "Categories"
        Call RenameCategory

    'Rename Applicant...
    Case "Applicants"
        Call RenameApplicant

End Select

End Sub
Private Sub lblEdit_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

If sTreeViewID = "Categories" Then
    Help.HelpText = "Edits the selected category."
Else
    Help.HelpText = "Edits the selected applicant."
End If

End Sub
Private Sub lblExit_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

Help.HelpText = "Closes the category window."

End Sub

Private Sub lblHelp_Click()

'Show the help window...
Call ShowHelp(Help.HelpText, Me)

End Sub
Private Sub lblHelp_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

If Button = vbLeftButton Then
    imgHelp.Picture = imgButton(1).Picture
    lblHelp.ForeColor = QBColor(0)
End If

End Sub
Private Sub lblHelp_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

Help.HelpText = "Shows the help window."

End Sub
Private Sub lblHelp_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

imgHelp.Picture = imgButton(0).Picture
lblHelp.ForeColor = lButtonForeColor

End Sub

Private Sub lblMoveTo_Click()

'Move applicants to another category...
'frmMoveApplicant.Show
'frmMoveApplicant.ZOrder

End Sub
Private Sub lblMoveTo_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

If Button = vbLeftButton Then
    imgMoveTo.Picture = imgButton(1).Picture
    lblMoveTo.ForeColor = QBColor(0)
End If

End Sub
Private Sub lblMoveTo_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

Help.HelpText = "Moves a company or applicant to another category."

End Sub
Private Sub lblMoveTo_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

imgMoveTo.Picture = imgButton(0).Picture
lblMoveTo.ForeColor = lButtonForeColor

End Sub
Private Sub lblCategories_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)

'Move the form if the user is pressing and holding the mouse button...
If Button = vbLeftButton Then
    Call DragForm(Me)
    If Help.HelpIsAligned = True And Help.HelpIsLoaded = True Then
        Call AlignHelpToForm
    End If
End If

End Sub
Private Sub lblDelete_Click()

'Delete...
Select Case sTreeViewID

    'Delete Category...
    Case "Categories"
        Call DeleteCategory

    'Delete Applicant...
    Case "Applicants"
        Call DeleteApplicant

End Select

End Sub
Private Sub lblDelete_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

If Button = vbLeftButton Then
    imgDelete.Picture = imgButton(1).Picture
    lblDelete.ForeColor = QBColor(0)
End If

End Sub

Private Sub lblDelete_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

imgDelete.Picture = imgButton(0).Picture
lblDelete.ForeColor = lButtonForeColor

End Sub

Private Sub lblEdit_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

If Button = vbLeftButton Then
    imgEdit.Picture = imgButton(1).Picture
    lblEdit.ForeColor = QBColor(0)
End If

End Sub
Private Sub lblEdit_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

imgEdit.Picture = imgButton(0).Picture
lblEdit.ForeColor = lButtonForeColor

End Sub
Private Sub lblExit_Click()

Unload Me

End Sub
Private Sub lblExit_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

If Button = vbLeftButton Then
    imgExit.Picture = imgButton(1).Picture
    lblExit.ForeColor = QBColor(0)
End If

End Sub
Private Sub lblExit_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

imgExit.Picture = imgButton(0).Picture
lblExit.ForeColor = lButtonForeColor

End Sub
Private Sub lblNew_Click()

'New...
Select Case sTreeViewID

    'New Category...
    Case "Categories"
        Call CreateNewCategory

    'New Applicant...
    Case "Applicants"
        Call CreateNewApplicant

End Select

End Sub
Private Sub lblNew_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

If Button = vbLeftButton Then
    imgNew.Picture = imgButton(1).Picture
    lblNew.ForeColor = QBColor(0)
End If

End Sub

Private Sub lblNew_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

If sTreeViewID = "Categories" Then
    Help.HelpText = "Creates a new category."
Else
    Help.HelpText = "Creates a new applicant."
End If

End Sub
Private Sub lblNew_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

imgNew.Picture = imgButton(0).Picture
lblNew.ForeColor = lButtonForeColor

End Sub

Private Sub lblReload_Click()

'Load all existing categories...
If LoadAllCategories(tvwCategories) = False Then
    MsgBox "Unable to load the Categories. This is an unexpected error and also critical. After restarting the program, and you keep getting this error, please contact " & Help.TechnicalSupportCompany & " at " & Help.TechnicalSupportPhone & " to help you resolve the problem.", vbCritical, "Categories Error..."
    Timer2.Enabled = True
End If

End Sub
Private Sub lblReload_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

If Button = vbLeftButton Then
    imgReload.Picture = imgButton(1).Picture
    lblReload.ForeColor = QBColor(0)
End If

End Sub

Private Sub lblReload_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

If sTreeViewID = "Categories" Then
    Help.HelpText = "Reloads all categories."
Else
    Help.HelpText = "Reloads all applicants."
End If

End Sub
Private Sub lblReload_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

imgReload.Picture = imgButton(0).Picture
lblReload.ForeColor = lButtonForeColor

End Sub

Private Sub lblSave_Click()

Call SaveChanges

End Sub
Private Sub lblSave_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

If Button = vbLeftButton Then
    imgSave.Picture = imgButton(1).Picture
    lblSave.ForeColor = QBColor(0)
End If

End Sub

Private Sub lblSave_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

Help.HelpText = "Saves any changes you have made."

End Sub
Private Sub lblSave_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

imgSave.Picture = imgButton(0).Picture
lblSave.ForeColor = lButtonForeColor

End Sub

Private Sub tvwApplicants_AfterLabelEdit(Cancel As Integer, NewString As String)

iDirty = True

End Sub

Private Sub tvwApplicants_DblClick()

On Local Error Resume Next

Dim x As Long

frmApplicants.Show

If frmApplicants.lstContacts.Enabled = True Then
    QuickRef.Name = tvwApplicants.SelectedItem
    For x = 0 To frmApplicants.lstContacts.ListCount - 1
        If frmApplicants.lstContacts.List(x) = QuickRef.Name Then
            frmApplicants.lstContacts.ListIndex = x
            Exit For
        End If
    Next x
End If

frmApplicants.ZOrder

End Sub
Private Sub tvwApplicants_GotFocus()

sTreeViewID = "Applicants"

End Sub
Private Sub tvwApplicants_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

Help.HelpText = "List of all applicants in the selected category."

End Sub

Private Sub tvwApplicants_NodeClick(ByVal Node As MSComctlLib.Node)

'Applicant...
lblApplicant.Caption = "Applicant: " & tvwApplicants.SelectedItem

End Sub
Private Sub tvwCategories_AfterLabelEdit(Cancel As Integer, NewString As String)

iDirty = True

End Sub
Private Sub tvwCategories_GotFocus()

sTreeViewID = "Categories"

End Sub
Private Sub tvwCategories_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

Help.HelpText = "List of all categories currently installed on the system."

End Sub
Private Sub Timer1_Timer()

On Local Error Resume Next

Dim x As Long

'New...
If sTreeViewID = "Categories" Then
    If lblNew.Enabled = True And iDirty = True Then
        imgNew.Enabled = False
        lblNew.Enabled = False
    ElseIf lblNew.Enabled = False And iDirty = False Then
        imgNew.Enabled = True
        lblNew.Enabled = True
    End If
ElseIf sTreeViewID = "Applicants" Then
    If lblNew.Enabled = True And iDirty = True Then
        imgNew.Enabled = False
        lblNew.Enabled = False
    ElseIf lblNew.Enabled = False And iDirty = False Then
        imgNew.Enabled = True
        lblNew.Enabled = True
    End If
End If

'Edit...
If lblEdit.Enabled = True And iDirty = True Then
    imgEdit.Enabled = False
    lblEdit.Enabled = False
ElseIf lblEdit.Enabled = False And iDirty = False Then
    imgEdit.Enabled = True
    lblEdit.Enabled = True
End If

'Save and Reload...
If lblSave.Enabled = True And iDirty = False Then
    imgSave.Enabled = False
    lblSave.Enabled = False
    imgReload.Enabled = False
    lblReload.Enabled = False
ElseIf lblSave.Enabled = False And iDirty = True Then
    imgSave.Enabled = True
    lblSave.Enabled = True
    imgReload.Enabled = True
    lblReload.Enabled = True
End If

'Delete...
If lblDelete.Enabled = True And iDirty = True Then
    imgDelete.Enabled = False
    lblDelete.Enabled = False
ElseIf lblDelete.Enabled = False And iDirty = False Then
    imgDelete.Enabled = True
    lblDelete.Enabled = True
End If

'Move To and Copy To...
If lblMoveTo.Enabled = True And iDirty = True Then
    imgMoveTo.Enabled = False
    lblMoveTo.Enabled = False
ElseIf lblMoveTo.Enabled = False And iDirty = False Then
    imgMoveTo.Enabled = True
    lblMoveTo.Enabled = True
End If

End Sub
Private Sub Timer2_Timer()

Timer2.Enabled = False
iDirty = False
Unload Me

End Sub

Private Sub tvwCategories_NodeClick(ByVal Node As MSComctlLib.Node)

'No category selected...
If tvwCategories.SelectedItem.Tag = "" Then
    tvwApplicants.Nodes.Clear
    Exit Sub
End If

'Get all Applicants for this Category...
If LoadAllCategorizedApplicants(tvwApplicants, Val(tvwCategories.SelectedItem.Tag)) = False Then
    Exit Sub
End If

'Category...
lblCategory.Caption = "Category: " & tvwCategories.SelectedItem

End Sub
