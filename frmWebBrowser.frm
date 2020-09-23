VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form frmWebBrowser 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Web Browser"
   ClientHeight    =   4425
   ClientLeft      =   0
   ClientTop       =   285
   ClientWidth     =   9915
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   9915
   ShowInTaskbar   =   0   'False
   Begin MSComDlg.CommonDialog Dialog 
      Left            =   1800
      Top             =   2550
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.ComboBox cboAddressBar 
      Height          =   315
      Left            =   960
      TabIndex        =   1
      Top             =   900
      Width           =   1365
   End
   Begin SHDocVwCtl.WebBrowser Web 
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1260
      Width           =   1395
      ExtentX         =   2461
      ExtentY         =   2355
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   870
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   9915
      _ExtentX        =   17489
      _ExtentY        =   1535
      ButtonWidth     =   1746
      ButtonHeight    =   1376
      Appearance      =   1
      ImageList       =   "ImageList"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Back"
            Key             =   "BACK"
            Object.ToolTipText     =   "Go Back (Ctrl + B)"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Forward"
            Key             =   "FORWARD"
            Object.ToolTipText     =   "Go Forward (Ctrl + F)"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Stop"
            Key             =   "STOP"
            Object.ToolTipText     =   "Stop (Ctrl + S)"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Refresh"
            Key             =   "REFRESH"
            Object.ToolTipText     =   "Refresh Page (F5)"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Home"
            Key             =   "HOME"
            Object.ToolTipText     =   "Home Page (Ctrl + H)"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Favorites..."
            Key             =   "FAVORITES"
            Object.ToolTipText     =   "Manage Favorite Sites...(Ctrl + F)"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Help..."
            Key             =   "HELP"
            Object.ToolTipText     =   "Help..."
            ImageIndex      =   7
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "       Exit       "
            Key             =   "EXIT"
            Object.ToolTipText     =   "Exit Program (Ctrl + X)"
            ImageIndex      =   8
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   1740
      Top             =   1320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWebBrowser.frx":0000
            Key             =   "BACK"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWebBrowser.frx":0454
            Key             =   "FORWARD"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWebBrowser.frx":08A8
            Key             =   "STOP"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWebBrowser.frx":0CFC
            Key             =   "REFRESH"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWebBrowser.frx":1150
            Key             =   "HOME"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWebBrowser.frx":15A4
            Key             =   "FAVORITES"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWebBrowser.frx":1E80
            Key             =   "HELP"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWebBrowser.frx":275C
            Key             =   "EXIT"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblAddressBar 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Address Bar"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   60
      TabIndex        =   2
      Top             =   945
      Width           =   855
   End
   Begin VB.Image imgButton 
      Height          =   300
      Index           =   0
      Left            =   1740
      Picture         =   "frmWebBrowser.frx":3054
      Stretch         =   -1  'True
      Top             =   1890
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgButton 
      Height          =   330
      Index           =   1
      Left            =   1740
      Picture         =   "frmWebBrowser.frx":4A76
      Stretch         =   -1  'True
      Top             =   2190
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuBack 
         Caption         =   "&Back"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuForward 
         Caption         =   "&Forward"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuStop 
         Caption         =   "&Stop"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuRefresh 
         Caption         =   "&Refresh"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuHomePage 
         Caption         =   "H&ome Page"
         Shortcut        =   ^H
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "&Help..."
      End
      Begin VB.Menu h87625 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrinterSetup 
         Caption         =   "P&rinter Setup..."
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "&Print"
         Shortcut        =   ^P
      End
      Begin VB.Menu h8276354 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuFavoriteS 
      Caption         =   "&Favorites"
      Begin VB.Menu mnuAddToFavorites 
         Caption         =   "&Add To Favorites..."
         Shortcut        =   {F7}
      End
      Begin VB.Menu mnuOrganizeFavorites 
         Caption         =   "&Organize Favorites..."
         Shortcut        =   {F4}
      End
   End
End
Attribute VB_Name = "frmWebBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboAddressBar_Click()

On Local Error Resume Next

Web.Navigate cboAddressBar.Text
Web.Refresh

End Sub
Private Sub cboAddressBar_KeyPress(KeyAscii As Integer)

On Local Error Resume Next

Dim x As Long

'Enter Key...
If KeyAscii = vbKeyReturn Then
    KeyAscii = 0
    Web.Navigate cboAddressBar.Text

    'Add new website address to combobox...
    Call AddWebEntry(cboAddressBar.Text)

    'Add site to address bar...
    For x = 0 To cboAddressBar.ListCount - 1
        If cboAddressBar.List(x) = cboAddressBar.Text Then
            x = True
            Exit For
        End If
    Next x
    If x <> True Then
        cboAddressBar.AddItem cboAddressBar.Text
    End If

End If

End Sub

Private Sub Form_Load()

On Local Error Resume Next

'Hide the main menu toolbar...
mdiMainMenu.Toolbar1.Visible = False

'Tell other programs that the web browser is loaded...
WebBrowser.IsLoaded = True

'Load Internet Sites...
Call LoadInternetSites

'Start with the home page...
Web.GoHome

End Sub
Sub LoadInternetSites()

On Local Error Resume Next

'Query the database and see if user exists...
If OpenDB(DB, QuickRef.DBPassWord) = False Then Exit Sub
Set RS = DB.OpenRecordset("SELECT * FROM tblInternetSites", dbOpenSnapshot)

'No info found for contact ???...
If RS.RecordCount = 0 Then
    RS.Close
    DB.Close
    Exit Sub
End If

'Load all internet sites into the combo box...
cboAddressBar.Clear
Do
    cboAddressBar.AddItem RS!Address
    RS.MoveNext
Loop Until RS.EOF

'Close the db...
RS.Close
DB.Close

End Sub
Private Sub Form_Resize()

'Resize the web browser form and all controls on it...
Call ResizeWebBrowserForm

End Sub

Private Sub Form_Unload(Cancel As Integer)

On Local Error Resume Next

'Tell other programs that the web browser is unloaded...
WebBrowser.IsLoaded = False

End Sub

Private Sub mnuBack_Click()

'Back...
On Local Error Resume Next
Web.GoBack
Web.Refresh

End Sub

Private Sub mnuExit_Click()

'Unload the help form...
If Help.HelpCallingForm = Me.Name Then
    Unload frmHelper
End If

Unload Me

End Sub
Private Sub mnuForward_Click()

'Forward...
On Local Error Resume Next
Web.GoForward
Web.Refresh

End Sub

Private Sub mnuHomePage_Click()

On Local Error Resume Next

'Home...
Web.GoHome

End Sub

Private Sub mnuPrint_Click()

On Local Error Resume Next

Printer.Print Web.Document

End Sub
Private Sub mnuPrinterSetup_Click()

On Local Error Resume Next

Dialog.ShowPrinter

'Ask to print now...
If MsgBox("Ok to print now?", vbYesNo + vbQuestion, "Print...") = vbYes Then
    Printer.Print Web.Document
End If

End Sub
Private Sub mnuRefresh_Click()

On Local Error Resume Next

'Refresh...
Web.Refresh

End Sub
Private Sub mnuStop_Click()

'Stop...
Web.Stop

End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

On Local Error GoTo ToolBar1_ButtonClickError

'Toolbar Buttons...
Select Case Button.Key

    'Back...
    Case "BACK"
        On Local Error Resume Next
        Web.GoBack
        Web.Refresh

    'Forward...
    Case "FORWARD"
        On Local Error Resume Next
        Web.GoForward
        Web.Refresh

    'Stop...
    Case "STOP"
        Web.Stop

    'Refresh...
    Case "REFRESH"
        Web.Refresh

    'Home...
    Case "HOME"
        On Local Error Resume Next
        Web.GoHome

    'Favorites...
    Case "FAVORITES"
        frmInternetSites.Show
        frmInternetSites.ZOrder

    'Help...
    Case "HELP"
        Help.HelpCallingForm = Me.Name
        frmHelper.Show
        frmHelper.ZOrder

    'Exit...
    Case "EXIT"
        If Help.HelpCallingForm = Me.Name Then
            Unload frmHelper
        End If
        Unload Me

End Select

Exit Sub



ToolBar1_ButtonClickError:
    On Local Error Resume Next
    Exit Sub

End Sub
Private Sub Web_NavigateComplete2(ByVal pDisp As Object, URL As Variant)

cboAddressBar.Text = Web.LocationURL

End Sub

