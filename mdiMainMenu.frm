VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm mdiMainMenu 
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000C&
   Caption         =   "Contactor 2000"
   ClientHeight    =   8190
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   11880
   Icon            =   "mdiMainMenu.frx":0000
   Picture         =   "mdiMainMenu.frx":08CA
   ScrollBars      =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   1290
      Top             =   1110
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   870
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   1535
      ButtonWidth     =   2117
      ButtonHeight    =   1376
      Appearance      =   1
      ImageList       =   "ImageList"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Contacts..."
            Key             =   "CONTACTS"
            Object.ToolTipText     =   "Contacts Screen..."
            ImageIndex      =   2
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Internet Sites..."
            Key             =   "INTERNETSITES"
            Object.ToolTipText     =   "Manage Internet Site Favorites..."
            ImageIndex      =   4
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Web Browser..."
            Key             =   "WEBBROWSER"
            Object.ToolTipText     =   "Connect to the Internet..."
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Re-Login..."
            Key             =   "LOGIN"
            Object.ToolTipText     =   "Login as another user..."
            ImageIndex      =   5
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Accounts..."
            Key             =   "ACCOUNTS"
            Object.ToolTipText     =   "Modify the system accounts..."
            ImageIndex      =   6
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Colors..."
            Key             =   "COLORS"
            Object.ToolTipText     =   "Set Program Colors..."
            ImageIndex      =   7
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Exit"
            Key             =   "EXIT"
            Object.ToolTipText     =   "Exit Program"
            ImageIndex      =   8
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.PictureBox picUserInfo 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   780
         Left            =   8640
         Picture         =   "mdiMainMenu.frx":24090C
         ScaleHeight     =   780
         ScaleWidth      =   3285
         TabIndex        =   2
         Top             =   30
         Width           =   3285
         Begin VB.Label lblLoginTime 
            BackStyle       =   0  'Transparent
            ForeColor       =   &H00C0FFC0&
            Height          =   195
            Left            =   1110
            TabIndex        =   6
            Tag             =   "ButtonLabel"
            Top             =   420
            Width           =   2085
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Login Time:"
            ForeColor       =   &H00C0E0FF&
            Height          =   195
            Left            =   120
            TabIndex        =   5
            Top             =   420
            Width           =   825
         End
         Begin VB.Label lblUser 
            BackStyle       =   0  'Transparent
            ForeColor       =   &H00C0FFC0&
            Height          =   195
            Left            =   1110
            TabIndex        =   4
            Tag             =   "ButtonLabel"
            Top             =   120
            Width           =   2085
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "User:"
            ForeColor       =   &H00C0E0FF&
            Height          =   195
            Left            =   120
            TabIndex        =   3
            Top             =   120
            Width           =   375
         End
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   7905
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   503
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   180
      Top             =   990
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
            Picture         =   "mdiMainMenu.frx":248F5E
            Key             =   "CONSULTANTS"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMainMenu.frx":24983A
            Key             =   "CLIENTS"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMainMenu.frx":24A116
            Key             =   "JOBORDERS"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMainMenu.frx":24A9F2
            Key             =   "ENGAGEMENTS"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMainMenu.frx":24B2CE
            Key             =   "LOGIN"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMainMenu.frx":24BBB2
            Key             =   "ACCOUNTS"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMainMenu.frx":24C496
            Key             =   "COLORS"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMainMenu.frx":24C7B2
            Key             =   "EXIT"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog Dialog 
      Left            =   780
      Top             =   1050
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuContacts 
         Caption         =   "&Contacts..."
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuWebBrowser 
         Caption         =   "&Web Browser..."
         Shortcut        =   ^W
      End
      Begin VB.Menu mnuReLogin 
         Caption         =   "&Re-Login..."
      End
      Begin VB.Menu H837625 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInternetSites 
         Caption         =   "&Internet Sites..."
         Begin VB.Menu mnuInternetLinks 
            Caption         =   ""
            Index           =   0
         End
      End
      Begin VB.Menu h981826 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrinterSetup 
         Caption         =   "Printer &Setup..."
         Shortcut        =   ^P
      End
      Begin VB.Menu H83762 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuFind 
         Caption         =   "&Find..."
         Shortcut        =   ^F
      End
      Begin VB.Menu h837652 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAccounts 
         Caption         =   "&Accounts..."
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuColorSettings 
         Caption         =   "&Colors..."
         Shortcut        =   ^L
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      WindowList      =   -1  'True
      Begin VB.Menu mnuCascade 
         Caption         =   "&Cascade"
      End
      Begin VB.Menu mnuTileHorizontally 
         Caption         =   "Tile &Horizontally"
      End
      Begin VB.Menu mnuTileVertically 
         Caption         =   "Tile &Vertically"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuIndex 
         Caption         =   "&Index..."
      End
      Begin VB.Menu mnuSearchForHelpOn 
         Caption         =   "&Search for help on..."
      End
      Begin VB.Menu h9287653 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About..."
      End
   End
End
Attribute VB_Name = "mdiMainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Sub LoadInternetSites()

On Local Error Resume Next

Dim iCount As Integer

'Query the database and see if user exists...
If OpenDB(DB, QuickRef.DBPassWord) = False Then Exit Sub
Set RS = DB.OpenRecordset("SELECT * FROM tblInternetSites", dbOpenSnapshot)

'No info found for contact ???...
If RS.RecordCount = 0 Then
    RS.Close
    DB.Close
    Exit Sub
End If

'First two items...
mnuInternetLinks(0).Caption = "Manage Sites..."
Load mnuInternetLinks(1)
mnuInternetLinks(1).Caption = "-"

'Load all internet sites...
iCount = 2
Do While Not RS.EOF
    Load mnuInternetLinks(iCount)
    mnuInternetLinks(iCount).Caption = RS!Address
    iCount = iCount + 1
    RS.MoveNext
Loop

'Close the db...
RS.Close
DB.Close

End Sub

Private Sub MDIForm_Load()

On Local Error Resume Next

'Set program colors...
Call SetColors(Me)

'Load the main menu's form settings...
Call LoadINISettings

'Set visible to true so that the main menu will be visible with the login dialog box in front of it...
Me.Visible = True
DoEvents

'Show the login screen...
QuickRef.ReLoggingIn = False
frmLogin.Show vbModal

'Load Internet Sites...
Call LoadInternetSites

Timer1.Enabled = True

End Sub
Sub LoadINISettings()

'Form properties...
If Trim$(ReadINI(Me.Name, "Caption")) <> "" Then
    Me.Caption = ReadINI(Me.Name, "Caption")
End If

'Form Coordinates...
Me.WindowState = Val(ReadINI(Me.Name, "WindowState"))
If Me.WindowState = vbMaximized Then Exit Sub
Me.Left = Val(ReadINI(Me.Name, "Left"))
Me.Top = Val(ReadINI(Me.Name, "Top"))
Me.Height = Val(ReadINI(Me.Name, "Height"))
Me.Width = Val(ReadINI(Me.Name, "Width"))

End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)

'Don't unload the entire program if the web browser is currently loaded, just unload the web browser...
If WebBrowser.IsLoaded Then
    Unload frmWebBrowser
    Cancel = True
End If

End Sub

Private Sub MDIForm_Resize()

'Resize the web browser form and all controls on it...
Call ResizeWebBrowserForm

End Sub
Private Sub MDIForm_Unload(Cancel As Integer)

'Save this form's settings...
Call SaveINISettings

End Sub
Sub SaveINISettings()

'WindowState...
Call WriteINI(Me.Name, "WindowState", Me.WindowState)

'If windowstate is maximized, exit...
If Me.WindowState = vbMaximized Then Exit Sub

'Form coordinates...
Call WriteINI(Me.Name, "Left", Me.Left)
Call WriteINI(Me.Name, "Top", Me.Top)
Call WriteINI(Me.Name, "Height", Me.Height)
Call WriteINI(Me.Name, "Width", Me.Width)

End Sub

Private Sub mnuAccounts_Click()

frmAccounts.Show
frmAccounts.ZOrder

End Sub
Private Sub mnuCascade_Click()

Call ArrangeIcons(vbCascade)

End Sub

Private Sub mnuColorSettings_Click()

frmColors.Show
frmColors.ZOrder

End Sub
Private Sub mnuContacts_Click()

frmContactor.Show
frmContactor.ZOrder

End Sub
Private Sub mnuExit_Click()

'Unload the help window...
If Help.HelpCallingForm = Me.Name Then
    Unload frmHelper
End If

Unload Me

End Sub
Private Sub mnuFind_Click()

MsgBox "To be implemented.", vbInformation, "Find..."

End Sub

Private Sub mnuInternetLinks_Click(Index As Integer)

On Local Error Resume Next

Dim x As Long

'Internet Links...
If mnuInternetLinks(Index).Caption = "Manage Sites..." Then
    frmInternetSites.Show
    frmInternetSites.ZOrder
    Exit Sub
End If

'Start the internet site...
If WebBrowser.UserContactorsWebBrowser = False Then
    Shell ("Start " & mnuInternetLinks(Index).Caption), vbHide
Else
    If WebBrowser.IsLoaded = False Then
        Load frmWebBrowser
    End If
    frmWebBrowser.Web.Navigate mnuInternetLinks(Index).Caption
    frmWebBrowser.Show
    frmWebBrowser.ZOrder
End If

End Sub
Private Sub mnuPrinterSetup_Click()

On Local Error Resume Next

Dialog.ShowPrinter

End Sub

Private Sub mnuReLogin_Click()

On Local Error Resume Next

'Confirm Re-Login...
If Forms.Count > 1 Then
    If MsgBox("Note: This will close all open windows. Do you want to continue?", vbYesNo + vbQuestion, "Re-Login...") = vbNo Then
        Exit Sub
    End If
End If

'Close all open windows...
Call CloseAllOpenWindows

'User clicked cancel on one of the open forms, so exit...
If Forms.Count > 1 Then Exit Sub

'Show the login window...
QuickRef.ReLoggingIn = True
frmLogin.Show vbModal
QuickRef.ReLoggingIn = False

End Sub

Private Sub mnuTileHorizontally_Click()

Call ArrangeIcons(vbHorizontal)

End Sub
Private Sub mnuTileVertically_Click()

Call ArrangeIcons(vbVertical)

End Sub

Private Sub mnuWebBrowser_Click()

frmWebBrowser.Show
frmWebBrowser.ZOrder

End Sub
Private Sub Timer1_Timer()

On Local Error Resume Next

Dim x As Long

'Show the main menu toolbar...
mdiMainMenu.Toolbar1.Visible = WebBrowser.IsLoaded = False

'Set Colors...
If QuickRef.UpdateColors = True Then
    Call LoadProgramColors
    QuickRef.UpdateColors = False
    For x = 0 To Forms.Count - 1
        Call SetColors(Forms(x))
    Next x
End If

'Accounts Toolbar Button...
If WebBrowser.IsLoaded = False Then
    Toolbar1.Buttons(5).Enabled = Login.IsAdmin
End If

'Audible Help...
If Help.HelpIsLoaded = True Then
    If frmHelper.txtHelper.Text <> Help.HelpText Then
        frmHelper.txtHelper.Text = Help.HelpText
        Call frmHelper.SpeakHelp
    End If
End If

End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

On Local Error GoTo ToolBar1_ButtonClickError

'Toolbar Buttons...
Select Case Button.Key

    'Contacts...
    Case "CONTACTS"
        mnuContacts_Click

    'Internet Sites...
    Case "INTERNETSITES"
        frmInternetSites.Show
        frmInternetSites.ZOrder

    'Web Browser...
    Case "WEBBROWSER"
        mnuWebBrowser_Click

    'Re-Login...
    Case "LOGIN"
        mnuReLogin_Click

    'Accounts...
    Case "ACCOUNTS"
        mnuAccounts_Click

    'Colors...
    Case "COLORS"
        mnuColorSettings_Click

    'Exit...
    Case "EXIT"
        Unload frmHelper
        Unload Me

End Select

Exit Sub



ToolBar1_ButtonClickError:
    Call WriteToErrorLog(Me.Name, "ToolBar1_ButtonClickError", Error, Err, False)
    Exit Sub

End Sub

