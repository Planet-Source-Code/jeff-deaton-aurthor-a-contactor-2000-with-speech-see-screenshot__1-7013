Attribute VB_Name = "Globals"
Option Explicit

'Misc...
Global DB As Database
Global RS As Recordset
Global SQL As String
Global Help As tHelp
Global Login As tLogin
Global QuickRef As tQuickRef
Global WebBrowser As tWeb
Global iStartDragAndDrop As Boolean
Global iActiveStatus As Long

'Color variables...
Global lButtonForeColor As Long
Global lLabelForeColor As Long
Global lTextBoxBackColor As Long
Global lTextBoxForeColor As Long
Global lListBoxBackColor As Long

'For Dragging Borderless Forms...
Private Declare Function SendMessage Lib "USER32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function ReleaseCapture Lib "USER32" () As Long
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2

'INI File Functions...
Public Declare Function WritePrivateProfileString Lib "KERNEL32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileString Lib "KERNEL32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

'Always on top...
Declare Function SetWindowPos Lib "USER32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Global Const HWND_TOPMOST = -1
Global Const HWND_NOTOPMOST = -2
Global Const SWP_NOACTIVATE = &H10
Global Const SWP_SHOWWINDOW = &H40

'Quick Reference Array...
Type tQuickRef
    CallingForm As String
    ContactName As String
    ContactID As Long
    DBFileName As String
    DBPassWord As String
    DBTimeOut As Long
    INIFileName As String
    LargeMenuHeight As Long
    LargeMenuWidth As Long
    MediumMenuHeight As Long
    MediumMenuWidth As Long
    NotesHaveChanged As Boolean
    PassNotes As Boolean
    ReLoggingIn As Boolean
    UpdateColors As Boolean
    UpdateInternetSites As Boolean
    UpdateNotes As Boolean
End Type

'Login...
Type tLogin
    FullName As String
    IsAdmin As Boolean
    LoginDateTime As String
    LoginName As String
End Type

'Web Browser Type...
Type tWeb
    HomePage As String
    IsLoaded As Boolean
    UserContactorsWebBrowser As Boolean
End Type

'Technical Support Type...
Type tHelp
    TechnicalSupportCompany As String
    TechnicalSupportPhone As String
    AudibleHelp As Boolean
    HelpText As String
    HelpIsLoaded As Boolean             'Tells all forms that the help form is loaded...
    HelpIsAligned As Boolean            'Tells all forms that the help form is loaded...
    HelpCallingForm As String           'Tells the helper form what form loaded it...
End Type
Public Sub AlwaysOnTop(Who As Form, iPosition As Boolean)

Dim lFlag As Long

'On top or not on top...
If iPosition Then
    lFlag = -1
Else
    lFlag = -2
End If

'Call the API to make the form on or not on top...
Call SetWindowPos(Who.hwnd, lFlag, Who.Left / Screen.TwipsPerPixelX, Who.Top / Screen.TwipsPerPixelY, Who.Width / Screen.TwipsPerPixelX, Who.Height / Screen.TwipsPerPixelY, SWP_NOACTIVATE Or SWP_SHOWWINDOW)

End Sub
Function AddNewInternetSite(sSiteName As String, tvw As TreeView) As Boolean

On Local Error GoTo AddNewInternetSiteError

Dim x As Long
Dim TempNode As Node

'Determine which to save, Company or Applicant...
If OpenDB(DB, QuickRef.DBPassWord) = False Then Exit Function
Set RS = DB.OpenRecordset("SELECT * FROM tblInternetSites", dbOpenDynaset)

'Add the new internet address...
RS.AddNew
RS!Address = sSiteName
RS.Update

'Load all Internet Sites...
Set RS = DB.OpenRecordset("SELECT * FROM tblInternetSites", dbOpenDynaset)
tvw.Nodes.Clear
Set TempNode = tvw.Nodes.Add(, , "R", "Internet Sites", 1)
Do
    Set TempNode = tvw.Nodes.Add("R", tvwChild, "", RS!Address, 2)
    TempNode.EnsureVisible
    RS.MoveNext
Loop Until RS.EOF

RS.Close
DB.Close
AddNewInternetSite = True
Exit Function



AddNewInternetSiteError:
    DB.Close
    Call WriteToErrorLog("GLOBAL", "AddNewInternetSiteError", Error, Err, True)
    Exit Function

End Function
Function LoadAllInternetSites(tvw As TreeView) As Boolean

On Local Error GoTo LoadAllInternetSitesError

Dim x As Long
Dim TempNode As Node

'Determine which to save, Company or Applicant...
If OpenDB(DB, QuickRef.DBPassWord) = False Then Exit Function
ReLoadInternetSites:
Set RS = DB.OpenRecordset("SELECT * FROM tblInternetSites", dbOpenDynaset)

'No Internet Sites to load. Create a default category and exit...
If RS.RecordCount = 0 Then
    RS.AddNew
    RS!Address = "[None]"
    RS.Close
    GoTo ReLoadInternetSites
End If

'Load all Internet Sites...
tvw.Nodes.Clear
Set TempNode = tvw.Nodes.Add(, , "R", "Internet Sites", 1)
Do
    Set TempNode = tvw.Nodes.Add("R", tvwChild, "", RS!Address, 2)
    TempNode.EnsureVisible
    RS.MoveNext
Loop Until RS.EOF

RS.Close
DB.Close
LoadAllInternetSites = True
Exit Function



LoadAllInternetSitesError:
    DB.Close
    Call WriteToErrorLog("GLOBAL", "LoadAllInternetSitesError", Error, Err, True)
    Exit Function

End Function
Sub AlignHelpToForm()

On Local Error Resume Next

Dim x As Long

'Align the helper form to the form it is associated with...
If Help.HelpIsAligned = True Then
    For x = 0 To Forms.Count - 1
        If Forms(x).Name = Help.HelpCallingForm Then
            frmHelper.Left = Forms(x).Left + Forms(x).Width + 20
            frmHelper.Top = Forms(x).Top
            Exit For
        End If
    Next x
End If

'Make sure the helper form is visible...
If frmHelper.Left >= mdiMainMenu.Width - 600 Then
    frmHelper.Left = mdiMainMenu.Width - frmHelper.Width - 180
    frmHelper.ZOrder
End If

End Sub
Sub AddWebEntry(sAddress As String)

On Local Error GoTo AddWebEntryError

If OpenDB(DB, QuickRef.DBPassWord) = False Then Exit Sub
Set RS = DB.OpenRecordset("SELECT * FROM tblInternetSites WHERE Address = '" & sAddress & "'", dbOpenDynaset)

'Search for this site already in the list...
If RS.RecordCount > 0 Then
    RS.Close
    DB.Close
    Exit Sub
Else
    RS.AddNew
    RS!Address = sAddress
    RS.Update
End If

RS.Close
DB.Close
Exit Sub



AddWebEntryError:
    DB.Close
    Call WriteToErrorLog("Globals", "AddWebEntryError", Error$, Err, False)
    Exit Sub

End Sub
Sub CloseAllOpenWindows()

On Local Error Resume Next

Dim x As Long

'Close all currently open forms...
For x = 1 To Forms.Count - 1
    If Forms(x).Name <> "mdiMainMenu" Then
        Unload Forms(x)
    End If
Next x

End Sub
Sub ArrangeIcons(iArrangeType As Integer)

mdiMainMenu.Arrange iArrangeType

End Sub
Function DeleteContact(sContact As String) As Boolean

On Local Error GoTo DeleteContactError

'Confirm...
If MsgBox("Are you sure you want to delete the contact " & UCase$(sContact) & "?", vbYesNo + vbQuestion + vbDefaultButton2, "Delete...") = vbNo Then
    Exit Function
End If

If OpenDB(DB, QuickRef.DBPassWord) = False Then Exit Function
Set RS = DB.OpenRecordset("SELECT * FROM tblContacts WHERE ContactName = '" & sContact & "' AND ContactID = " & QuickRef.ContactID, dbOpenDynaset)

'Mark all contacts with that name as inactive...(keep them in the db)...
If RS.RecordCount > 0 Then
    Do
        RS.Edit
        RS!Active = False
        RS.Update
        RS.MoveNext
    Loop Until RS.EOF
End If

RS.Close
DB.Close
DeleteContact = True
Exit Function



DeleteContactError:
    DB.Close
    Call WriteToErrorLog("Global", "DeleteContactError", Error$, Err, True)
    Exit Function

End Function
Function GetNotes(ThisTextBox As TextBox) As String

On Local Error GoTo GetNotesError

'Get the notes for this contact...
If OpenDB(DB, QuickRef.DBPassWord) = False Then Exit Function
Set RS = DB.OpenRecordset("SELECT Comments FROM tblContacts WHERE ContactName = '" & QuickRef.ContactName & "' AND ContactID = " & QuickRef.ContactID, dbOpenSnapshot)

'Should never happen...???
If RS.RecordCount = 0 Then
    RS.Close
    DB.Close
    ThisTextBox = ""
Else
    ThisTextBox = RS!Comments
    RS.Close
    DB.Close
End If

QuickRef.NotesHaveChanged = False
QuickRef.UpdateNotes = False

Exit Function



GetNotesError:
    DB.Close
    Call WriteToErrorLog("Global", "GetNotesError", Error$, Err, True)
    Exit Function

End Function

Sub LoadProgramColors()

On Local Error Resume Next

lLabelForeColor = Val(ReadINI("Colors", "lLabelForeColor"))
lButtonForeColor = Val(ReadINI("Colors", "lButtonForeColor"))
lTextBoxBackColor = Val(ReadINI("Colors", "lTextBoxBackColor"))
lTextBoxForeColor = Val(ReadINI("Colors", "lTextBoxForeColor"))
lListBoxBackColor = Val(ReadINI("Colors", "lListBoxBackColor"))

End Sub
Function OpenFile(cControl As TextBox, sFileName As String) As Boolean

On Local Error GoTo OpenFileError

Dim sNam As String
Dim FileFree As Long

'Read in the file...
FileFree = FreeFile
Open sFileName For Input As #FileFree
    Do
        Line Input #FileFree, sNam
        cControl = cControl & sNam
    Loop Until EOF(FileFree)
Close #FileFree

Exit Function



OpenFileError:
    Close
    Call WriteToErrorLog("Global", "OpenFileError", Error$, Err, True)
    Exit Function
    Resume Next

End Function
Function ParseURL(sFileName As String) As String

On Local Error GoTo ParseURLError

Dim sNam As String
Dim FileFree As Long

FileFree = FreeFile
Open sFileName For Input As #FileFree
    Do
        Line Input #FileFree, sNam
        If InStr(sNam, "[InternetShortcut]") > 0 Then
            Line Input #FileFree, sNam
            If Right$(sFileName, 3) = "url" Then
                ParseURL = Mid$(sNam, 5)
            End If
            Exit Do
        End If
    Loop Until EOF(FileFree)
Close #FileFree

Exit Function



ParseURLError:
    Close
    ParseURL = "Error"
    Call WriteToErrorLog("Globals", "ParseURLError", Error$, Err, False)
    Exit Function

End Function
Sub Say(sSayWhat As String, vVoice As DirectSS)

On Local Error Resume Next

'Speak...
vVoice.AudioReset
vVoice.Speak sSayWhat

End Sub
Sub SetColors(Who As Form)

On Local Error Resume Next

Dim x As Long

'Set label and button label fore colors...
For x = 0 To Who.Controls.Count - 1
    If InStr(LCase$(Who.Controls(x).Tag), "nocolorchange") = 0 Then
        'Label Colors...
        If Who.Controls(x).Tag = "Label" Then
            Who.Controls(x).ForeColor = lLabelForeColor
        'Button Label Colors...
        ElseIf Who.Controls(x).Tag = "ButtonLabel" Then
            Who.Controls(x).ForeColor = lButtonForeColor
        'Textbox ForeGround and BackGround Colors...
        ElseIf TypeOf Who.Controls(x) Is TextBox Then
            Who.Controls(x).ForeColor = lTextBoxForeColor
            Who.Controls(x).BackColor = lTextBoxBackColor
        'List and combo box BackGround Colors...
        ElseIf TypeOf Who.Controls(x) Is ListBox Or TypeOf Who.Controls(x) Is ComboBox Then
            Who.Controls(x).BackColor = lListBoxBackColor
            Who.Controls(x).ForeColor = lTextBoxForeColor
        End If
    End If
Next x

End Sub
Function OpenDB(DB As Database, sPassWord As String) As Boolean

'This routine waits until the db is free and then opens it for the user.
'It times out when QuickRef.DBTimeOut reaches it's limit and returns an open error...

On Local Error Resume Next

Dim iOpenCount As Long

Do
    Err = 0
    Set DB = OpenDatabase(QuickRef.DBFileName, True, False, ";pwd=" & sPassWord)
    iOpenCount = iOpenCount + 1
Loop Until Err = 0 Or iOpenCount >= QuickRef.DBTimeOut

'Set function to true or false...
If iOpenCount >= QuickRef.DBTimeOut Then
    OpenDB = False
Else
    OpenDB = True
End If

End Function
Public Sub DragForm(Frm As Form)

On Local Error Resume Next

'Move the borderless form...
Call ReleaseCapture
Call SendMessage(Frm.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)

'Align the help window to the form that loaded it...
If Help.HelpIsLoaded And Help.HelpIsAligned Then
    Call AlignHelpToForm
End If

End Sub
Sub Main()

On Local Error Resume Next

'DB Password...
QuickRef.DBPassWord = "jkddjkdd"

'Window Heights and Widths...
QuickRef.MediumMenuHeight = 4320
QuickRef.MediumMenuWidth = 7020
QuickRef.LargeMenuHeight = 5705
QuickRef.LargeMenuWidth = 9665

'INI Filename...
If Dir$(App.Path & "\Contactor.Ini") <> "" Then
    QuickRef.INIFileName = App.Path & "\Contactor.Ini"
Else
    MsgBox "I can't find the Contactor.Ini file. This file is mandatory for this program to run correctly. If you can find this file yourself elsewhere on your computer, you can safely copy it to " & UCase$(App.Path) & " yourself. If this does not work, you will have to reinstall the program.", vbCritical, "File not found..."
    End
End If

'Database Filename...
If Trim(ReadINI("Database", "DatabaseLocation")) <> "" Then
    QuickRef.DBFileName = ReadINI("Database", "DatabaseLocation") & "Contactor.Mdb"
Else
    QuickRef.DBFileName = App.Path & "\Contactor.Mdb"
End If
If Dir$(QuickRef.DBFileName) = "" Then
    MsgBox "I can't find the program's Database. The Database's filename is " & Chr$(34) & "Contactor.Mdb" & Chr$(34) & ". If you can find this file somewhere else on your computer, you should copy or move it to this location: " & UCase$(App.Path) & ". If all else fails, you will have to re-install the program.", vbCritical, "Lost Database..."
    End
End If
QuickRef.DBTimeOut = 500

'Load color settings...
Call LoadProgramColors

'HomePage...
WebBrowser.HomePage = ReadINI("frmInternetSites", "HomePage")
WebBrowser.UserContactorsWebBrowser = Val(ReadINI("frmInternetSites", "chkUseContactor2000Browser")) = 1

'Voice Messages...
If Dir$(App.Path & "\" & "VM", vbDirectory) = "" Then
    MkDir App.Path & "\" & "VM"
End If

'Main Menu...
mdiMainMenu.Show

End Sub
Sub ResizeWebBrowserForm()

On Local Error Resume Next

'Exit if the web browser form is not loaded...
If WebBrowser.IsLoaded = False Then Exit Sub

'Resize the form...
frmWebBrowser.Left = 0
frmWebBrowser.Top = 0
frmWebBrowser.Width = mdiMainMenu.Width - 160
frmWebBrowser.Height = mdiMainMenu.Height - 1040

'Web Browser Control...
frmWebBrowser.Web.Left = 0
frmWebBrowser.Web.Top = frmWebBrowser.Toolbar1.Height + 370
frmWebBrowser.Web.Height = mdiMainMenu.Height - 2240
frmWebBrowser.Web.Width = frmWebBrowser.Width

'Address Bar...
frmWebBrowser.cboAddressBar.Top = frmWebBrowser.Toolbar1.Height + 30
frmWebBrowser.cboAddressBar.Left = 980
frmWebBrowser.cboAddressBar.Width = frmWebBrowser.Width - 980

End Sub
Sub ResizeToMdiMain(Who As Form)

End Sub
Function StartWord(sFileName As String) As Boolean

On Local Error GoTo StartWordError

Dim x As Long

'Shell the application...
Call Shell("Start " & sFileName, vbHide)

Exit Function



StartWordError:
    Call WriteToErrorLog("Global", "StartWordError", Error$, Err, True)
    Exit Function

End Function
Function WriteINI(sSection As String, sKeyName As String, sNewString As String) As Boolean

On Local Error Resume Next

Call WritePrivateProfileString(sSection, sKeyName, sNewString, QuickRef.INIFileName)

WriteINI = (Err = 0)

End Function
Function ReadINI(sSection As String, sKeyName As String) As String

On Local Error Resume Next

Dim sRet As String

sRet = String(255, Chr(0))

ReadINI = Left(sRet, GetPrivateProfileString(sSection, ByVal sKeyName, "", sRet, Len(sRet), QuickRef.INIFileName))

End Function
Sub WriteToErrorLog(sFormName As String, sRoutineName As String, sError As String, iErrorNumber As Integer, iDisplayMsgBox As Boolean)

On Local Error Resume Next

Dim FileFree As Integer

FileFree = FreeFile
Open App.Path & "\ErrorLog.Txt" For Append As #FileFree
    Print #FileFree, sFormName, sRoutineName, sError, iErrorNumber
Close #FileFree

'Display the error that occured...
If iDisplayMsgBox = True Then
    MsgBox "The following error has occured in your program: " & vbCrLf & vbCrLf & sError & vbCrLf & vbCrLf & "Error Number: " & iErrorNumber, vbInformation, "Error..."
End If

End Sub
