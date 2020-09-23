VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmInternetSites 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Internet Sites"
   ClientHeight    =   8775
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14550
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmInternetSites.frx":0000
   ScaleHeight     =   8775
   ScaleWidth      =   14550
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkSingleClickLaunchesWebSite 
      Height          =   195
      Left            =   6570
      TabIndex        =   11
      Top             =   4830
      Width           =   195
   End
   Begin VB.FileListBox File1 
      Height          =   285
      Left            =   13110
      Pattern         =   "*.URL;*.HTM;*.HTML;*.LNK"
      TabIndex        =   9
      Top             =   2100
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.DirListBox Dir1 
      Height          =   315
      Left            =   13110
      TabIndex        =   8
      Top             =   1680
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.CheckBox chkUseContactor2000Browser 
      Height          =   195
      Left            =   6570
      TabIndex        =   3
      Top             =   4560
      Width           =   195
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   13140
      Top             =   240
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   13620
      Top             =   180
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
            Picture         =   "frmInternetSites.frx":B3412
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInternetSites.frx":B3CEE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvwFavoriteDirs 
      Height          =   3855
      Left            =   270
      TabIndex        =   5
      Top             =   570
      Width           =   4275
      _ExtentX        =   7541
      _ExtentY        =   6800
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   4
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      FullRowSelect   =   -1  'True
      ImageList       =   "ImageList1"
      Appearance      =   1
   End
   Begin MSComctlLib.TreeView tvwFavoriteSites 
      Height          =   3885
      Left            =   4590
      TabIndex        =   10
      Top             =   540
      Width           =   4785
      _ExtentX        =   8440
      _ExtentY        =   6853
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   4
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      ImageList       =   "ImageList1"
      Appearance      =   1
   End
   Begin VB.Image Image1 
      Height          =   3855
      Left            =   4560
      Picture         =   "frmInternetSites.frx":B45CA
      Stretch         =   -1  'True
      Top             =   570
      Width           =   45
   End
   Begin VB.Label lblSingleClickLaunchesWebSite 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Single Click Launches Website"
      ForeColor       =   &H00FFFFC0&
      Height          =   195
      Left            =   6840
      TabIndex        =   12
      Tag             =   "Label"
      Top             =   4830
      Width           =   2205
   End
   Begin VB.Label lblDelete 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Delete"
      ForeColor       =   &H00C0FFC0&
      Height          =   195
      Left            =   7560
      TabIndex        =   7
      Tag             =   "ButtonLabel"
      Top             =   5250
      Width           =   495
   End
   Begin VB.Label lblRefresh 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Refresh"
      ForeColor       =   &H00C0FFC0&
      Height          =   195
      Left            =   6375
      TabIndex        =   6
      Tag             =   "ButtonLabel"
      Top             =   5250
      Width           =   585
   End
   Begin VB.Label lblHelp 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Help..."
      ForeColor       =   &H00C0FFC0&
      Height          =   195
      Left            =   5280
      TabIndex        =   4
      Tag             =   "ButtonLabel"
      Top             =   5250
      Width           =   495
   End
   Begin VB.Label lblInternetSites 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Internet Sites / Favorites"
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
      Left            =   420
      TabIndex        =   2
      Top             =   60
      Width           =   2130
   End
   Begin VB.Label lblUseContactor2000Browser 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Use Contactor 2000's Web Browser"
      ForeColor       =   &H00FFFFC0&
      Height          =   195
      Left            =   6840
      TabIndex        =   1
      Tag             =   "Label"
      Top             =   4560
      Width           =   2535
   End
   Begin VB.Image imgOKPicture 
      Height          =   375
      Index           =   1
      Left            =   13080
      Picture         =   "frmInternetSites.frx":B662C
      Top             =   1230
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.Image imgOKPicture 
      Height          =   360
      Index           =   0
      Left            =   13080
      Picture         =   "frmInternetSites.frx":B7D16
      Stretch         =   -1  'True
      Top             =   810
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.Label lblExit 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Exit"
      ForeColor       =   &H00C0FFC0&
      Height          =   195
      Left            =   8820
      TabIndex        =   0
      Tag             =   "ButtonLabel"
      Top             =   5250
      Width           =   255
   End
   Begin VB.Image imgExit 
      Height          =   375
      Left            =   8340
      Picture         =   "frmInternetSites.frx":B9738
      Stretch         =   -1  'True
      Top             =   5160
      Width           =   1155
   End
   Begin VB.Image imgHelp 
      Height          =   375
      Left            =   4920
      Picture         =   "frmInternetSites.frx":BB15A
      Stretch         =   -1  'True
      Top             =   5160
      Width           =   1155
   End
   Begin VB.Image imgRefresh 
      Height          =   375
      Left            =   6060
      Picture         =   "frmInternetSites.frx":BCB7C
      Stretch         =   -1  'True
      Top             =   5160
      Width           =   1155
   End
   Begin VB.Image imgDelete 
      Height          =   375
      Left            =   7200
      Picture         =   "frmInternetSites.frx":BE59E
      Stretch         =   -1  'True
      Top             =   5160
      Width           =   1155
   End
End
Attribute VB_Name = "frmInternetSites"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub ClearAllFields()

Dim TempNode As Node

'Clear the treeview control and then add the default root item...
tvwFavoriteDirs.Nodes.Clear
Set TempNode = tvwFavoriteDirs.Nodes.Add(, , "R", "Favorites", 1)
TempNode.EnsureVisible

End Sub
Function LoadFavoriteDirs() As Boolean

On Local Error GoTo LoadFavoriteDirsError

Dim x As Long
Dim TempNode As Node

'Clear the treeview control...
tvwFavoriteDirs.Nodes.Clear

'Look for the favorites folder...
If Dir$("C:\WINDOWS\FAVORITES\", vbDirectory) <> "" Then
    Dir1.Path = "C:\WINDOWS\FAVORITES\"
ElseIf Dir$("C:\WINNT\FAVORITES\", vbDirectory) <> "" Then
    Dir1.Path = "C:\WINNT\FAVORITES\"
Else
    MsgBox "I can't find your favorites folder. Attempting to load sites from the database.", vbInformation, "Favorites..."
    Call LoadAllInternetSites(tvwFavoriteSites)
    Exit Function
End If

'Add Favorite Dirs...
Dir1.Refresh
If Dir1.ListCount > 0 Then
    Set TempNode = tvwFavoriteDirs.Nodes.Add(, , Dir1.Path, UCase$(Dir1.Path), 1)
    TempNode.Tag = Dir1.Path
    For x = 0 To Dir1.ListCount - 1
        File1.Path = Dir1.List(x)
        File1.Refresh
        If File1.ListCount > 0 Then
            Set TempNode = tvwFavoriteDirs.Nodes.Add(Dir1.Path, tvwChild, Dir1.List(x), UCase$(Dir1.List(x)), 2)
            TempNode.Tag = Dir1.List(x)
            TempNode.EnsureVisible
        End If
    Next x
End If

LoadFavoriteDirs = True
Exit Function



LoadFavoriteDirsError:
    Call WriteToErrorLog("GLOBAL", "LoadFavoriteDirsError", Error, Err, False)
    Exit Function
    Resume Next

End Function
Private Sub chkUseContactor2000Browser_Click()

WebBrowser.UserContactorsWebBrowser = chkUseContactor2000Browser.Value = 1

End Sub

Private Sub Form_Load()

'Load the main menu's form settings...
Call LoadINISettings

'Set program colors...
Call SetColors(Me)

'Load keywords for this applicant, company...
Call LoadFavoriteDirs

'Form Coordinates...
Me.Height = QuickRef.LargeMenuHeight
Me.Width = QuickRef.LargeMenuWidth

End Sub
Function LoadFavoriteSites() As Boolean

On Local Error GoTo LoadFavoriteSitesError

Dim x As Long
Dim TempNode As Node

'Clear the treeview control...
File1.Path = tvwFavoriteDirs.SelectedItem.Tag
File1.Refresh
tvwFavoriteSites.Nodes.Clear

'Add Favorite Sites...
Set TempNode = tvwFavoriteSites.Nodes.Add(, , File1.Path, "Internet Sites", 1)
TempNode.Tag = "Internet Sites"
For x = 0 To File1.ListCount - 1
    Set TempNode = tvwFavoriteSites.Nodes.Add(File1.Path, tvwChild, File1.List(x), File1.List(x), 2)
    TempNode.Tag = File1.List(x)
    TempNode.EnsureVisible
Next x

LoadFavoriteSites = True
Exit Function



LoadFavoriteSitesError:
    Call WriteToErrorLog("GLOBAL", "LoadFavoriteSitesError", Error, Err, False)
    Exit Function
    Resume Next

End Function
Sub LoadINISettings()

'Form Coordinates...
Me.Left = Val(ReadINI(Me.Name, "Left"))
Me.Top = Val(ReadINI(Me.Name, "Top"))

'Web Browser...
chkUseContactor2000Browser.Value = Val(ReadINI(Me.Name, "chkUseContactor2000Browser"))

'Single Click for Browsing...
chkSingleClickLaunchesWebSite.Value = Val(ReadINI(Me.Name, "chkSingleClickLaunchesWebSite"))

End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

'Move the form if the user is pressing and holding the mouse button...
If Button = vbLeftButton Then
    Call DragForm(Me)
End If

Help.HelpText = ""

End Sub
Private Sub Form_Unload(Cancel As Integer)

'Save INI Settings...
Call SaveINISettings

End Sub
Sub SaveINISettings()

'Form coordinates...
Call WriteINI(Me.Name, "Left", Me.Left)
Call WriteINI(Me.Name, "Top", Me.Top)

'Web Browser...
Call WriteINI(Me.Name, "chkUseContactor2000Browser", chkUseContactor2000Browser.Value)

'Single Click for Web Browsing...
Call WriteINI(Me.Name, "chkSingleClickLaunchesWebSite", chkSingleClickLaunchesWebSite.Value)

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

Private Sub imgRefresh_Click()

lblRefresh_Click

End Sub
Private Sub imgRefresh_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

If Button = vbLeftButton Then
    imgRefresh.Picture = imgOKPicture(1).Picture
    lblRefresh.ForeColor = QBColor(0)
End If

End Sub

Private Sub imgRefresh_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

imgRefresh.Picture = imgOKPicture(0).Picture
lblRefresh.ForeColor = lButtonForeColor

End Sub

Private Sub lblDelete_Click()

On Local Error GoTo lblDelete_ClickError

Dim x As Long

'Disallow deletion of folders...
If tvwFavoriteSites.SelectedItem.Image = 1 Then Exit Sub

'Confirm to delete the selected site...
If MsgBox("Are you sure you want to Delete " & UCase$(tvwFavoriteSites.SelectedItem) & "?", vbYesNo + vbQuestion + vbDefaultButton2, "Delete Site...") = vbNo Then
    Exit Sub
End If

'Delete the site...
Kill tvwFavoriteSites.Nodes(1).Key & "\" & tvwFavoriteSites.SelectedItem
tvwFavoriteSites.Nodes.Remove tvwFavoriteSites.SelectedItem.Index

'Set The Dirty Flag...
Exit Sub



lblDelete_ClickError:
    Call WriteToErrorLog(Me.Name, "lblDelete_ClickError", Error$, Err, True)
    Exit Sub

End Sub
Private Sub lblDelete_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

If Button = vbLeftButton Then
    imgDelete.Picture = imgOKPicture(1).Picture
    lblDelete.ForeColor = QBColor(0)
End If

End Sub

Private Sub lblDelete_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

Help.HelpText = "Deletes the selected site from disk."

End Sub
Private Sub lblDelete_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

imgDelete.Picture = imgOKPicture(0).Picture
lblDelete.ForeColor = lButtonForeColor

End Sub

Private Sub lblExit_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

Help.HelpText = "Exits this screen."

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
Private Sub lblInternetSites_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

'Move the form if the user is pressing and holding the mouse button...
If Button = vbLeftButton Then
    Call DragForm(Me)
End If

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

Private Sub lblRefresh_Click()

'Refresh the sites...
Call LoadFavoriteSites

End Sub

Private Sub lblRefresh_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

If Button = vbLeftButton Then
    imgRefresh.Picture = imgOKPicture(1).Picture
    lblRefresh.ForeColor = QBColor(0)
End If

End Sub

Private Sub lblRefresh_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

Help.HelpText = "Refreshes the internet sites listbox."

End Sub
Private Sub lblRefresh_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

imgRefresh.Picture = imgOKPicture(0).Picture
lblRefresh.ForeColor = lButtonForeColor

End Sub
Private Sub lblSingleClickLaunchesWebSite_Click()

'Toggle on / off state...
If chkSingleClickLaunchesWebSite.Value = False Then
    chkSingleClickLaunchesWebSite.Value = 1
Else
    chkSingleClickLaunchesWebSite.Value = False
End If

End Sub

Private Sub lblSingleClickLaunchesWebSite_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

Help.HelpText = "Determines whether or not a single mouse click or a double mouse click will launch the selected web site."

End Sub
Private Sub lblUseContactor2000Browser_Click()

'Toggle on / off state...
If chkUseContactor2000Browser.Value = False Then
    chkUseContactor2000Browser.Value = 1
Else
    chkUseContactor2000Browser.Value = False
End If

'Update the reference immediately...
WebBrowser.UserContactorsWebBrowser = chkUseContactor2000Browser.Value = 1

End Sub

Private Sub lblUseContactor2000Browser_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

Help.HelpText = "Determines whether or not to use Contactor 2000's web browser or the windows default web browser."

End Sub
Private Sub Timer1_Timer()

On Local Error Resume Next

'Delete...
If imgDelete.Enabled = True And tvwFavoriteSites.SelectedItem = "Internet Sites" Or imgDelete.Enabled = True And tvwFavoriteSites.SelectedItem = "" Then
    imgDelete.Enabled = False
    lblDelete.Enabled = False
ElseIf imgDelete.Enabled = False And tvwFavoriteSites.SelectedItem <> "Internet Sites" And tvwFavoriteSites.SelectedItem <> "" Then
    imgDelete.Enabled = True
    lblDelete.Enabled = True
End If

End Sub
Private Sub tvwFavoriteDirs_BeforeLabelEdit(Cancel As Integer)

Cancel = True

End Sub

Private Sub tvwFavoriteDirs_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

tvwFavoriteDirs.ToolTipText = ""
Help.HelpText = "Listing of all directories under the favorites folder on your hard drive."

End Sub
Private Sub tvwFavoriteDirs_NodeClick(ByVal Node As MSComctlLib.Node)

'Get Sub Folders...
Call LoadFavoriteSites

End Sub

Private Sub tvwFavoriteSites_BeforeLabelEdit(Cancel As Integer)

Cancel = True

End Sub
Private Sub tvwFavoriteSites_Click()

On Local Error GoTo tvwFavoriteSites_ClickError

'Check the single click status...
If chkSingleClickLaunchesWebSite.Value = 0 Then Exit Sub

'Exit if it's a top level item...
If tvwFavoriteSites.SelectedItem = "Internet Sites" Then Exit Sub

'Start the internet site...
If WebBrowser.UserContactorsWebBrowser = False Then
    Call Shell("Start " & ParseURL(tvwFavoriteSites.Nodes(1).Key & "\" & tvwFavoriteSites.SelectedItem.Tag), vbHide)
Else
    If WebBrowser.IsLoaded = False Then
        Load frmWebBrowser
    End If
    frmWebBrowser.Web.Navigate ParseURL(tvwFavoriteSites.Nodes(1).Key & "\" & tvwFavoriteSites.SelectedItem.Tag)
    frmWebBrowser.Show
    frmWebBrowser.ZOrder
End If

Exit Sub



tvwFavoriteSites_ClickError:
    Call WriteToErrorLog(Me.Name, "tvwFavoriteSites_ClickError", Error$, Err, True)
    Exit Sub

End Sub

Private Sub tvwFavoriteSites_DblClick()

On Local Error GoTo tvwFavoriteSites_DblClickError

'Check the single click status...
If chkSingleClickLaunchesWebSite.Value = 1 Then Exit Sub

'Exit if it's a top level item...
If tvwFavoriteSites.SelectedItem = "Internet Sites" Then Exit Sub

'Start the internet site...
If WebBrowser.UserContactorsWebBrowser = False Then
    Call Shell("Start " & ParseURL(tvwFavoriteSites.Nodes(1).Key & "\" & tvwFavoriteSites.SelectedItem.Tag), vbHide)
Else
    If WebBrowser.IsLoaded = False Then
        Load frmWebBrowser
    End If
    frmWebBrowser.Web.Navigate ParseURL(tvwFavoriteSites.Nodes(1).Key & "\" & tvwFavoriteSites.SelectedItem.Tag)
    frmWebBrowser.Show
    frmWebBrowser.ZOrder
End If

Exit Sub



tvwFavoriteSites_DblClickError:
    Call WriteToErrorLog(Me.Name, "tvwFavoriteSites_DblClickError", Error$, Err, True)
    Exit Sub

End Sub

Private Sub tvwFavoriteSites_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

Help.HelpText = "Listing of all internet sites currently on the hard drive."

End Sub
