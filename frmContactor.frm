VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{2398E321-5C6E-11D1-8C65-0060081841DE}#1.0#0"; "Vtext.dll"
Begin VB.Form frmContactor 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Address Book"
   ClientHeight    =   6300
   ClientLeft      =   60
   ClientTop       =   -225
   ClientWidth     =   11295
   ControlBox      =   0   'False
   Icon            =   "frmContactor.frx":0000
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   Picture         =   "frmContactor.frx":08CA
   ScaleHeight     =   6300
   ScaleWidth      =   11295
   ShowInTaskbar   =   0   'False
   Begin HTTSLibCtl.TextToSpeech Voice 
      Height          =   555
      Left            =   10500
      OleObjectBlob   =   "frmContactor.frx":B3CDC
      TabIndex        =   43
      Top             =   1950
      Width           =   435
   End
   Begin VB.CheckBox chkActiveStatus 
      Height          =   195
      Left            =   270
      TabIndex        =   40
      Top             =   4800
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.CommandButton btnExit 
      Caption         =   "E&xit"
      Height          =   285
      Left            =   9870
      TabIndex        =   39
      Top             =   4380
      Width           =   1065
   End
   Begin VB.CommandButton btnPrint 
      Caption         =   "&Print"
      Height          =   285
      Left            =   9870
      TabIndex        =   38
      Top             =   4080
      Width           =   1065
   End
   Begin VB.CommandButton btnColors 
      Caption         =   "&Colors"
      Height          =   285
      Left            =   9870
      TabIndex        =   37
      Top             =   3780
      Width           =   1065
   End
   Begin VB.CommandButton btnNew 
      Caption         =   "&New"
      Height          =   285
      Left            =   9870
      TabIndex        =   36
      Top             =   3480
      Width           =   1065
   End
   Begin VB.CommandButton btnDelete 
      Caption         =   "&Delete"
      Height          =   285
      Left            =   9870
      TabIndex        =   35
      Top             =   3180
      Width           =   1065
   End
   Begin VB.CommandButton btnSave 
      Caption         =   "&Save"
      Height          =   285
      Left            =   9870
      TabIndex        =   34
      Top             =   2880
      Width           =   1065
   End
   Begin VB.CommandButton btnReload 
      Caption         =   "&Reload"
      Height          =   285
      Left            =   9870
      TabIndex        =   33
      Top             =   2580
      Width           =   1065
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   9900
      Top             =   1410
   End
   Begin MSComDlg.CommonDialog Dialog 
      Left            =   9900
      Top             =   1890
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   9900
      Top             =   960
   End
   Begin VB.TextBox txtComments 
      BackColor       =   &H00C0E0FF&
      DragIcon        =   "frmContactor.frx":B3D00
      Height          =   1095
      Left            =   5520
      MultiLine       =   -1  'True
      OLEDropMode     =   1  'Manual
      ScrollBars      =   2  'Vertical
      TabIndex        =   23
      Top             =   3510
      Width           =   3855
   End
   Begin VB.TextBox txtJobEndTime 
      BackColor       =   &H00C0E0FF&
      DragIcon        =   "frmContactor.frx":B45CA
      Height          =   285
      Left            =   7950
      TabIndex        =   21
      Top             =   3150
      Width           =   1425
   End
   Begin VB.TextBox txtJobStartTime 
      BackColor       =   &H00C0E0FF&
      DragIcon        =   "frmContactor.frx":B4E94
      Height          =   285
      Left            =   5520
      TabIndex        =   19
      Top             =   3150
      Width           =   1095
   End
   Begin VB.TextBox txtConsultingFee 
      BackColor       =   &H00C0E0FF&
      DragIcon        =   "frmContactor.frx":B575E
      Height          =   285
      Left            =   5520
      TabIndex        =   17
      Top             =   2790
      Width           =   3855
   End
   Begin VB.TextBox txtIPAddress 
      BackColor       =   &H00C0E0FF&
      DragIcon        =   "frmContactor.frx":B6028
      Height          =   285
      Left            =   8160
      TabIndex        =   15
      Top             =   2430
      Width           =   1215
   End
   Begin VB.TextBox txtCellPhone 
      BackColor       =   &H00C0E0FF&
      DragIcon        =   "frmContactor.frx":B68F2
      Height          =   285
      Left            =   5520
      TabIndex        =   13
      Top             =   2430
      Width           =   1545
   End
   Begin VB.TextBox txtFax 
      BackColor       =   &H00C0E0FF&
      DragIcon        =   "frmContactor.frx":B71BC
      Height          =   285
      Left            =   7650
      TabIndex        =   11
      Top             =   2070
      Width           =   1725
   End
   Begin VB.TextBox txtWorkPhone 
      BackColor       =   &H00C0E0FF&
      DragIcon        =   "frmContactor.frx":B7A86
      Height          =   285
      Left            =   5520
      TabIndex        =   9
      Top             =   2070
      Width           =   1545
   End
   Begin VB.TextBox txtEMailAddress 
      BackColor       =   &H00C0E0FF&
      DragIcon        =   "frmContactor.frx":B8350
      Height          =   285
      Left            =   5520
      TabIndex        =   7
      Top             =   1710
      Width           =   3855
   End
   Begin VB.TextBox txtCityStateZip 
      BackColor       =   &H00C0E0FF&
      DragIcon        =   "frmContactor.frx":B8C1A
      Height          =   285
      Left            =   5520
      TabIndex        =   5
      Top             =   1350
      Width           =   3855
   End
   Begin VB.TextBox txtAddress 
      BackColor       =   &H00C0E0FF&
      DragIcon        =   "frmContactor.frx":B94E4
      Height          =   285
      Left            =   5520
      TabIndex        =   3
      Top             =   990
      Width           =   3855
   End
   Begin VB.TextBox txtContactName 
      BackColor       =   &H00C0E0FF&
      DragIcon        =   "frmContactor.frx":B9DAE
      Height          =   285
      Left            =   5520
      TabIndex        =   1
      Top             =   630
      Width           =   3855
   End
   Begin VB.ListBox lstContacts 
      BackColor       =   &H00C0E0FF&
      DragIcon        =   "frmContactor.frx":BA678
      Height          =   3960
      ItemData        =   "frmContactor.frx":BAF42
      Left            =   270
      List            =   "frmContactor.frx":BAF49
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   630
      Width           =   3855
   End
   Begin VB.Image imgThereAreVoiceMessages 
      Height          =   210
      Left            =   6990
      Picture         =   "frmContactor.frx":BAF5A
      Top             =   4800
      Width           =   180
   End
   Begin VB.Image imgRedLight 
      Height          =   210
      Index           =   1
      Left            =   9870
      Picture         =   "frmContactor.frx":BB194
      Top             =   4740
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Image imgRedLight 
      Height          =   210
      Index           =   0
      Left            =   10110
      Picture         =   "frmContactor.frx":BB3CE
      Top             =   4740
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Label lblVoiceMessages 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Voice Msgs..."
      ForeColor       =   &H00C0FFC0&
      Height          =   195
      Left            =   7260
      TabIndex        =   46
      Tag             =   "ButtonLabel"
      Top             =   4800
      UseMnemonic     =   0   'False
      Width           =   975
   End
   Begin VB.Label lblNotes 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Notes..."
      ForeColor       =   &H00C0FFC0&
      Height          =   195
      Left            =   8580
      TabIndex        =   45
      Tag             =   "ButtonLabel"
      Top             =   4800
      UseMnemonic     =   0   'False
      Width           =   585
   End
   Begin VB.Label lblHelp 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Help..."
      ForeColor       =   &H00C0FFC0&
      Height          =   195
      Left            =   1080
      TabIndex        =   44
      Tag             =   "ButtonLabel"
      Top             =   5250
      Width           =   495
   End
   Begin VB.Label lblMarkActive 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mark Active"
      ForeColor       =   &H00C0FFC0&
      Height          =   195
      Left            =   3090
      TabIndex        =   42
      Tag             =   "ButtonLabel"
      Top             =   4800
      UseMnemonic     =   0   'False
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.Label lblActiveStatus 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Viewing Active Contacts"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   225
      Left            =   540
      TabIndex        =   41
      Top             =   4800
      Width           =   1815
   End
   Begin VB.Image imgDragDrop 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   1
      Left            =   10500
      Picture         =   "frmContactor.frx":BB608
      Top             =   1530
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgDragDrop 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   0
      Left            =   10500
      Picture         =   "frmContactor.frx":BBA4A
      Top             =   960
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgTrashCan 
      Height          =   480
      Left            =   4620
      Picture         =   "frmContactor.frx":BC314
      Top             =   4230
      Width           =   480
   End
   Begin VB.Label lblReload 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Reload"
      Enabled         =   0   'False
      ForeColor       =   &H00C0FFC0&
      Height          =   195
      Left            =   2130
      TabIndex        =   24
      Tag             =   "ButtonLabel"
      Top             =   5250
      UseMnemonic     =   0   'False
      Width           =   525
   End
   Begin VB.Label lblSave 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Save"
      Enabled         =   0   'False
      ForeColor       =   &H00C0FFC0&
      Height          =   195
      Left            =   3270
      TabIndex        =   25
      Tag             =   "ButtonLabel"
      Top             =   5250
      UseMnemonic     =   0   'False
      Width           =   405
   End
   Begin VB.Label lblDelete 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Delete"
      Enabled         =   0   'False
      ForeColor       =   &H00C0FFC0&
      Height          =   195
      Left            =   4320
      TabIndex        =   26
      Tag             =   "ButtonLabel"
      Top             =   5250
      UseMnemonic     =   0   'False
      Width           =   495
   End
   Begin VB.Label lblNew 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "New"
      ForeColor       =   &H00C0FFC0&
      Height          =   195
      Left            =   5460
      TabIndex        =   27
      Tag             =   "ButtonLabel"
      Top             =   5250
      UseMnemonic     =   0   'False
      Width           =   345
   End
   Begin VB.Label lblColors 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Colors"
      ForeColor       =   &H00C0FFC0&
      Height          =   195
      Left            =   6480
      TabIndex        =   28
      Tag             =   "ButtonLabel"
      Top             =   5250
      UseMnemonic     =   0   'False
      Width           =   465
   End
   Begin VB.Label lblPrint 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Print"
      Enabled         =   0   'False
      ForeColor       =   &H00C0FFC0&
      Height          =   195
      Left            =   7620
      TabIndex        =   29
      Tag             =   "ButtonLabel"
      Top             =   5250
      UseMnemonic     =   0   'False
      Width           =   345
   End
   Begin VB.Image imgButton 
      Height          =   360
      Index           =   0
      Left            =   9870
      Picture         =   "frmContactor.frx":BCBDE
      Stretch         =   -1  'True
      Top             =   120
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.Image imgButton 
      Height          =   375
      Index           =   1
      Left            =   9870
      Picture         =   "frmContactor.frx":BE600
      Top             =   510
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
      Left            =   8730
      TabIndex        =   30
      Tag             =   "ButtonLabel"
      Top             =   5250
      UseMnemonic     =   0   'False
      Width           =   285
   End
   Begin VB.Label lblCaptions 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contacts"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   13
      Left            =   360
      TabIndex        =   32
      Top             =   60
      Width           =   795
   End
   Begin VB.Label lblCaptions 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Notes"
      ForeColor       =   &H00FFFFC0&
      Height          =   195
      Index           =   11
      Left            =   4410
      TabIndex        =   31
      Tag             =   "Label"
      Top             =   3540
      Width           =   435
   End
   Begin VB.Label lblCaptions 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Job End Time"
      ForeColor       =   &H00FFFFC0&
      Height          =   195
      Index           =   10
      Left            =   6900
      TabIndex        =   22
      Tag             =   "Label"
      Top             =   3180
      Width           =   975
   End
   Begin VB.Label lblCaptions 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Job Start Time"
      ForeColor       =   &H00FFFFC0&
      Height          =   195
      Index           =   9
      Left            =   4380
      TabIndex        =   20
      Tag             =   "Label"
      Top             =   3180
      Width           =   1035
   End
   Begin VB.Label lblCaptions 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Consulting Fee"
      ForeColor       =   &H00FFFFC0&
      Height          =   195
      Index           =   8
      Left            =   4380
      TabIndex        =   18
      Tag             =   "Label"
      Top             =   2820
      Width           =   1065
   End
   Begin VB.Label lblCaptions 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "IP Address"
      ForeColor       =   &H00FFFFC0&
      Height          =   195
      Index           =   7
      Left            =   7320
      TabIndex        =   16
      Tag             =   "Label"
      Top             =   2460
      Width           =   765
   End
   Begin VB.Label lblCaptions 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cell Phone"
      ForeColor       =   &H00FFFFC0&
      Height          =   195
      Index           =   6
      Left            =   4380
      TabIndex        =   14
      Tag             =   "Label"
      Top             =   2460
      Width           =   765
   End
   Begin VB.Label lblCaptions 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fax"
      ForeColor       =   &H00FFFFC0&
      Height          =   195
      Index           =   5
      Left            =   7320
      TabIndex        =   12
      Tag             =   "Label"
      Top             =   2100
      Width           =   255
   End
   Begin VB.Label lblCaptions 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Work Phone"
      ForeColor       =   &H00FFFFC0&
      Height          =   195
      Index           =   4
      Left            =   4380
      TabIndex        =   10
      Tag             =   "Label"
      Top             =   2100
      Width           =   915
   End
   Begin VB.Label lblCaptions 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "E-Mail Address"
      ForeColor       =   &H00FFFFC0&
      Height          =   195
      Index           =   3
      Left            =   4380
      TabIndex        =   8
      Tag             =   "Label"
      Top             =   1740
      Width           =   1065
   End
   Begin VB.Label lblCaptions 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "City, State, Zip"
      ForeColor       =   &H00FFFFC0&
      Height          =   195
      Index           =   2
      Left            =   4380
      TabIndex        =   6
      Tag             =   "Label"
      Top             =   1380
      Width           =   1035
   End
   Begin VB.Label lblCaptions 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      ForeColor       =   &H00FFFFC0&
      Height          =   195
      Index           =   1
      Left            =   4380
      TabIndex        =   4
      Tag             =   "Label"
      Top             =   1020
      Width           =   585
   End
   Begin VB.Label lblCaptions 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contact Name"
      ForeColor       =   &H00FFFFC0&
      Height          =   195
      Index           =   0
      Left            =   4380
      TabIndex        =   2
      Tag             =   "Label"
      Top             =   660
      Width           =   1035
   End
   Begin VB.Image imgExit 
      Height          =   375
      Left            =   8310
      Picture         =   "frmContactor.frx":BFCEA
      Stretch         =   -1  'True
      Top             =   5160
      Width           =   1095
   End
   Begin VB.Image imgLabelHolder 
      Height          =   315
      Index           =   7
      Left            =   4260
      Picture         =   "frmContactor.frx":C170C
      Stretch         =   -1  'True
      Top             =   630
      Width           =   1245
   End
   Begin VB.Image imgLabelHolder 
      Height          =   315
      Index           =   8
      Left            =   4260
      Picture         =   "frmContactor.frx":C2D3E
      Stretch         =   -1  'True
      Top             =   990
      Width           =   1245
   End
   Begin VB.Image imgLabelHolder 
      Height          =   315
      Index           =   9
      Left            =   4260
      Picture         =   "frmContactor.frx":C4370
      Stretch         =   -1  'True
      Top             =   1350
      Width           =   1245
   End
   Begin VB.Image imgLabelHolder 
      Height          =   315
      Index           =   10
      Left            =   4260
      Picture         =   "frmContactor.frx":C59A2
      Stretch         =   -1  'True
      Top             =   1710
      Width           =   1245
   End
   Begin VB.Image imgLabelHolder 
      Height          =   315
      Index           =   11
      Left            =   4260
      Picture         =   "frmContactor.frx":C6FD4
      Stretch         =   -1  'True
      Top             =   2070
      Width           =   1245
   End
   Begin VB.Image imgLabelHolder 
      Height          =   315
      Index           =   12
      Left            =   7230
      Picture         =   "frmContactor.frx":C8606
      Stretch         =   -1  'True
      Top             =   2070
      Width           =   405
   End
   Begin VB.Image imgLabelHolder 
      Height          =   315
      Index           =   13
      Left            =   4260
      Picture         =   "frmContactor.frx":C9C38
      Stretch         =   -1  'True
      Top             =   2430
      Width           =   1245
   End
   Begin VB.Image imgLabelHolder 
      Height          =   315
      Index           =   4
      Left            =   7200
      Picture         =   "frmContactor.frx":CB26A
      Stretch         =   -1  'True
      Top             =   2430
      Width           =   945
   End
   Begin VB.Image imgLabelHolder 
      Height          =   315
      Index           =   3
      Left            =   4260
      Picture         =   "frmContactor.frx":CC89C
      Stretch         =   -1  'True
      Top             =   2790
      Width           =   1245
   End
   Begin VB.Image imgLabelHolder 
      Height          =   315
      Index           =   2
      Left            =   4260
      Picture         =   "frmContactor.frx":CDECE
      Stretch         =   -1  'True
      Top             =   3150
      Width           =   1245
   End
   Begin VB.Image imgLabelHolder 
      Height          =   315
      Index           =   1
      Left            =   6780
      Picture         =   "frmContactor.frx":CF500
      Stretch         =   -1  'True
      Top             =   3150
      Width           =   1155
   End
   Begin VB.Image imgLabelHolder 
      Height          =   315
      Index           =   0
      Left            =   4260
      Picture         =   "frmContactor.frx":D0B32
      Stretch         =   -1  'True
      Top             =   3510
      Width           =   1245
   End
   Begin VB.Image imgPrint 
      Height          =   375
      Left            =   7230
      Picture         =   "frmContactor.frx":D2164
      Stretch         =   -1  'True
      Top             =   5160
      Width           =   1095
   End
   Begin VB.Image imgColors 
      Height          =   375
      Left            =   6150
      Picture         =   "frmContactor.frx":D3B86
      Stretch         =   -1  'True
      Top             =   5160
      Width           =   1095
   End
   Begin VB.Image imgNew 
      Height          =   375
      Left            =   5070
      Picture         =   "frmContactor.frx":D55A8
      Stretch         =   -1  'True
      Top             =   5160
      Width           =   1095
   End
   Begin VB.Image imgDelete 
      Height          =   375
      Left            =   3990
      Picture         =   "frmContactor.frx":D6FCA
      Stretch         =   -1  'True
      Top             =   5160
      Width           =   1095
   End
   Begin VB.Image imgSave 
      Height          =   375
      Left            =   2910
      Picture         =   "frmContactor.frx":D89EC
      Stretch         =   -1  'True
      Top             =   5160
      Width           =   1095
   End
   Begin VB.Image imgReload 
      Height          =   375
      Left            =   1830
      Picture         =   "frmContactor.frx":DA40E
      Stretch         =   -1  'True
      Top             =   5160
      Width           =   1095
   End
   Begin VB.Image imgMarkActive 
      Height          =   375
      Left            =   2940
      Picture         =   "frmContactor.frx":DBE30
      Stretch         =   -1  'True
      Top             =   4710
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.Image imgHelp 
      Height          =   375
      Left            =   750
      Picture         =   "frmContactor.frx":DD852
      Stretch         =   -1  'True
      Top             =   5160
      Width           =   1095
   End
   Begin VB.Image imgNotes 
      Height          =   375
      Left            =   8310
      Picture         =   "frmContactor.frx":DF274
      Stretch         =   -1  'True
      Top             =   4710
      Width           =   1095
   End
   Begin VB.Image imgVoiceMessages 
      Height          =   375
      Left            =   6900
      Picture         =   "frmContactor.frx":E0C96
      Stretch         =   -1  'True
      Top             =   4710
      Width           =   1425
   End
End
Attribute VB_Name = "frmContactor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim iDirty As Boolean
Dim iContactNameHasChanged As Boolean
Sub ClearAllFields()

txtContactName = ""
txtAddress = ""
txtCityStateZip = ""
txtEMailAddress = ""
txtWorkPhone = ""
txtFax = ""
txtCellPhone = ""
txtIPAddress = ""
txtConsultingFee = ""
txtJobStartTime = ""
txtJobEndTime = ""
txtComments = ""

iDirty = False
iContactNameHasChanged = False

End Sub
Function LoadAllContacts(cControl As ListBox) As Boolean

On Local Error GoTo LoadAllContactsError

If OpenDB(DB, QuickRef.DBPassWord) = False Then Exit Function
Set RS = DB.OpenRecordset("SELECT * FROM tblContacts WHERE Active = " & iActiveStatus, dbOpenSnapshot)

'Clear the screen of all data including the contacts listbox...
Call ClearAllFields
cControl.Clear

'No records there...
If RS.RecordCount = 0 Then
    RS.Close
    DB.Close
    Exit Function
End If

cControl.Clear
Do
    cControl.AddItem RS!ContactName
    cControl.ItemData(cControl.NewIndex) = RS!ContactID
    RS.MoveNext
Loop Until RS.EOF

RS.Close
DB.Close

'Set to first item...
If cControl.ListCount > 0 Then
    QuickRef.ContactName = cControl.List(0)
    cControl.ListIndex = 0
End If

iContactNameHasChanged = False
Exit Function



LoadAllContactsError:
    DB.Close
    MsgBox "Error: " & Error$, vbInformation, "Error..."
    Exit Function
    Resume Next

End Function
Function MarkActive() As Boolean

On Local Error GoTo MarkActiveError

'Open the database...
If OpenDB(DB, QuickRef.DBPassWord) = False Then Exit Function
Set RS = DB.OpenRecordset("SELECT * FROM tblContacts WHERE ContactName = '" & QuickRef.ContactName & "' AND ContactID = " & QuickRef.ContactID, dbOpenDynaset)
If RS.RecordCount = 0 Then
    RS.AddNew
Else
    RS.Edit
End If

'Populate the fields on the form with the Applicant information...
RS!ContactName = txtContactName
RS!Address = txtAddress
RS!CityStateZip = txtCityStateZip
RS!EMailAddress = txtEMailAddress
RS!WorkPhone = txtWorkPhone
RS!Fax = txtFax
RS!CellPhone = txtCellPhone
RS!IPAddress = txtIPAddress
RS!ConsultingFee = txtConsultingFee
RS!Active = True
If Trim$(txtJobStartTime) <> "" Then
    If IsDate(Format$(txtJobStartTime, "MM/DD/YYYY H:MM AMPM")) Then
        RS!JobStartTime = Format$(txtJobStartTime, "MM/DD/YYYY H:MM AMPM")
    End If
End If
If Trim$(txtJobEndTime) <> "" Then
    If IsDate(Format$(txtJobEndTime, "MM/DD/YYYY H:MM AMPM")) Then
        RS!JobEndTime = txtJobEndTime
    End If
End If
RS!Comments = txtComments

RS.Update
RS.Close
DB.Close

MarkActive = True
iDirty = False
QuickRef.UpdateNotes = QuickRef.NotesHaveChanged

'Load all contacts based upon active status...
Call LoadAllContacts(lstContacts)

Exit Function



MarkActiveError:
    DB.Close
    Call WriteToErrorLog(Me.Name, "MarkActive", Error, Err, True)
    Exit Function

End Function
Function SaveChanges() As Boolean

On Local Error GoTo SaveChangesError

'Open the database...
If OpenDB(DB, QuickRef.DBPassWord) = False Then Exit Function
Set RS = DB.OpenRecordset("SELECT * FROM tblContacts WHERE ContactName = '" & QuickRef.ContactName & "' AND ContactID = " & QuickRef.ContactID, dbOpenDynaset)
If RS.RecordCount = 0 Then
    RS.AddNew
Else
    RS.Edit
End If

'Populate the fields on the form with the Applicant information...
RS!ContactName = txtContactName
RS!Address = txtAddress
RS!CityStateZip = txtCityStateZip
RS!EMailAddress = txtEMailAddress
RS!WorkPhone = txtWorkPhone
RS!Fax = txtFax
RS!CellPhone = txtCellPhone
RS!IPAddress = txtIPAddress
RS!ConsultingFee = txtConsultingFee
RS!Active = iActiveStatus

'Dates...
If Trim$(txtJobStartTime) <> "" Then
    If IsDate(Format$(txtJobStartTime, "MM/DD/YYYY H:MM AMPM")) Then
        RS!JobStartTime = Format$(txtJobStartTime, "MM/DD/YYYY H:MM AMPM")
    End If
End If
If Trim$(txtJobEndTime) <> "" Then
    If IsDate(Format$(txtJobEndTime, "MM/DD/YYYY H:MM AMPM")) Then
        RS!JobEndTime = txtJobEndTime
    End If
End If
RS!Comments = txtComments

RS.Update
RS.Close
DB.Close

'Reload the contacts listbox if the contact name has changed...
If iContactNameHasChanged Then
    Call LoadAllContacts(lstContacts)
End If

SaveChanges = True
iDirty = False
QuickRef.UpdateNotes = QuickRef.NotesHaveChanged

Exit Function


SaveChangesError:
    DB.Close
    Call WriteToErrorLog(Me.Name, "SaveChanges", Error, Err, True)
    Exit Function
    Resume Next

End Function
Private Sub btnColors_Click()

'These buttons are hidden. I am using them simply for the hot keys (ALT + Key)...

'Colors...
If lblColors.Enabled Then
    lblColors_Click
End If

End Sub
Private Sub btnDelete_Click()

'These buttons are hidden. I am using them simply for the hot keys (ALT + Key)...

'Delete...
If lblDelete.Enabled Then
    lblDelete_Click
End If

End Sub

Private Sub btnExit_Click()

'These buttons are hidden. I am using them simply for the hot keys (ALT + Key)...

'Exit...
If lblExit.Enabled Then
    lblExit_Click
End If

End Sub
Private Sub btnNew_Click()

'These buttons are hidden. I am using them simply for the hot keys (ALT + Key)...

'New...
If lblNew.Enabled Then
    lblNew_Click
End If

End Sub

Private Sub btnPrint_Click()

'These buttons are hidden. I am using them simply for the hot keys (ALT + Key)...

'Print...
If lblPrint.Enabled Then
    lblPrint_Click
End If

End Sub
Private Sub btnReload_Click()

'These buttons are hidden. I am using them simply for the hot keys (ALT + Key)...

'Reload...
If lblReload.Enabled Then
    lblReload_Click
End If

End Sub
Private Sub btnSave_Click()

'These buttons are hidden. I am using them simply for the hot keys (ALT + Key)...

'Save...
If lblSave.Enabled Then
    lblSave_Click
End If

End Sub

Private Sub chkActiveStatus_Click()

'Change the caption of the label control...
If chkActiveStatus.Value = 0 Then
    lblActiveStatus.Caption = "Viewing Inactive Contacts"
Else
    lblActiveStatus.Caption = "Viewing Active Contacts"
End If

'Set active status variable...
iActiveStatus = (chkActiveStatus.Value = 1)

'Load all contacts based upon active status...
Call LoadAllContacts(lstContacts)

End Sub
Private Sub Form_DblClick()

On Local Error Resume Next

'Change the windowstate...
If Me.WindowState = vbNormal Then
    Me.WindowState = vbMinimized
ElseIf Me.WindowState = vbMinimized Then
    Me.WindowState = vbNormal
End If

End Sub
Private Sub Form_Load()

On Local Error Resume Next

'Load INI Settings...
Call LoadINISettings

'Set Colors...
Call SetColors(Me)

'Load all doctors...
iActiveStatus = True
Call LoadAllContacts(lstContacts)

'Set form width and height...
Me.Height = QuickRef.LargeMenuHeight
Me.Width = QuickRef.LargeMenuWidth

iDirty = False
iContactNameHasChanged = False
QuickRef.NotesHaveChanged = False
QuickRef.UpdateNotes = False

End Sub
Sub LoadINISettings()

'Form Coordinates...
Me.Left = Val(ReadINI(Me.Name, "Left"))
Me.Top = Val(ReadINI(Me.Name, "Top"))

End Sub
Sub SaveINISettings()

'Form coordinates...
Call WriteINI(Me.Name, "Left", Me.Left)
Call WriteINI(Me.Name, "Top", Me.Top)

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

Help.HelpText = ""

'Move the form...
If Button = vbLeftButton And Me.WindowState = vbNormal Then
    Call DragForm(Me)
End If

End Sub
Private Sub Form_Unload(Cancel As Integer)

'Save INI Settings...
Call SaveINISettings

End Sub

Private Sub imgColors_Click()

lblColors_Click

End Sub
Private Sub imgColors_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

If Button = vbLeftButton Then
    imgColors.Picture = imgButton(1).Picture
    lblColors.ForeColor = QBColor(0)
End If

End Sub
Private Sub imgColors_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

imgColors.Picture = imgButton(0).Picture
lblColors.ForeColor = lButtonForeColor

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
Private Sub imgLabelHolder_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)

'Move the form...
If Button = vbLeftButton And Me.WindowState = vbNormal Then
    Call DragForm(Me)
End If

End Sub

Private Sub imgMarkActive_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

If Button = vbLeftButton Then
    imgMarkActive.Picture = imgButton(1).Picture
    lblMarkActive.ForeColor = QBColor(0)
End If

End Sub

Private Sub imgMarkActive_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

imgMarkActive.Picture = imgButton(0).Picture
lblMarkActive.ForeColor = lButtonForeColor

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

Private Sub imgNotes_Click()

lblNotes_Click

End Sub
Private Sub imgNotes_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

If Button = vbLeftButton Then
    imgNotes.Picture = imgButton(1).Picture
    lblNotes.ForeColor = QBColor(0)
End If

End Sub
Private Sub imgNotes_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

imgNotes.Picture = imgButton(0).Picture
lblNotes.ForeColor = lButtonForeColor

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

Private Sub imgTrashCan_DragDrop(Source As Control, x As Single, Y As Single)

On Local Error Resume Next

'Drag and Drop...
If TypeOf Source Is TextBox Then
    Source.Text = ""
ElseIf TypeOf Source Is ListBox Then
    Call lblDelete_Click
End If

End Sub

Private Sub imgTrashCan_DragOver(Source As Control, x As Single, Y As Single, State As Integer)

On Local Error Resume Next

'Drag picture...
Select Case State
    Case vbEnter
        Source.DragIcon = imgDragDrop(1).Picture
    Case vbLeave
        Source.DragIcon = imgDragDrop(0).Picture
End Select

End Sub

Private Sub imgTrashCan_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

Help.HelpText = "You can click and drag any field to this trash can for quick delete."

End Sub

Private Sub imgVoiceMessages_Click()

lblVoiceMessages_Click

End Sub
Private Sub imgVoiceMessages_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

If Button = vbLeftButton Then
    imgVoiceMessages.Picture = imgButton(1).Picture
    lblVoiceMessages.ForeColor = QBColor(0)
End If

End Sub

Private Sub imgVoiceMessages_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

imgVoiceMessages.Picture = imgButton(0).Picture
lblVoiceMessages.ForeColor = lButtonForeColor

End Sub
Private Sub lblActiveStatus_Click()

'Toggle Active Status...
If chkActiveStatus.Value = 1 Then
    chkActiveStatus.Value = 0
Else
    chkActiveStatus.Value = 1
End If

'Change the caption of the label control...
If chkActiveStatus.Value = 0 Then
    lblActiveStatus.Caption = "Viewing Inactive Contacts"
Else
    lblActiveStatus.Caption = "Viewing Active Contacts"
End If

'Set active status variable...
iActiveStatus = (chkActiveStatus.Value = 1)

End Sub

Private Sub lblActiveStatus_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

Help.HelpText = "Views active and inactive contacts."

End Sub
Private Sub lblCaptions_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)

'Move the form...
If Button = vbLeftButton And Me.WindowState = vbNormal Then
    Call DragForm(Me)
End If

End Sub

Private Sub lblColors_Click()

frmColors.Show
frmColors.ZOrder

End Sub
Private Sub lblColors_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

If Button = vbLeftButton Then
    imgColors.Picture = imgButton(1).Picture
    lblColors.ForeColor = QBColor(0)
End If

End Sub

Private Sub lblColors_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

Help.HelpText = "Click here to go to the colors window."

End Sub
Private Sub lblColors_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

imgColors.Picture = imgButton(0).Picture
lblColors.ForeColor = lButtonForeColor

End Sub

Private Sub lblDelete_Click()

'Delete this contact...
If DeleteContact(lstContacts.List(lstContacts.ListIndex)) = True Then
    Call LoadAllContacts(lstContacts)
End If

End Sub
Private Sub lblDelete_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

If Button = vbLeftButton Then
    imgDelete.Picture = imgButton(1).Picture
    lblDelete.ForeColor = QBColor(0)
End If

End Sub

Private Sub lblDelete_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

Help.HelpText = "Click here to delete this contact."

End Sub
Private Sub lblDelete_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

imgDelete.Picture = imgButton(0).Picture
lblDelete.ForeColor = lButtonForeColor

End Sub
Private Sub lblExit_Click()

Unload frmNotes
Unload frmContactor

Set frmNotes = Nothing
Set frmContactor = Nothing

End Sub
Private Sub lblExit_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

If Button = vbLeftButton Then
    imgExit.Picture = imgButton(1).Picture
    lblExit.ForeColor = QBColor(0)
End If

End Sub

Private Sub lblExit_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

Help.HelpText = "Click here to exit this window."

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
Private Sub lblMarkActive_Click()

Call MarkActive

End Sub
Private Sub lblMarkActive_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

If Button = vbLeftButton Then
    imgMarkActive.Picture = imgButton(1).Picture
    lblMarkActive.ForeColor = QBColor(0)
End If

End Sub

Private Sub lblMarkActive_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

Help.HelpText = "Click here to mark this contact as active again."

End Sub
Private Sub lblMarkActive_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

imgMarkActive.Picture = imgButton(0).Picture
lblMarkActive.ForeColor = lButtonForeColor

End Sub
Private Sub lblNew_Click()

On Local Error Resume Next

Dim x As Long
Dim sInput As String

'Clear all fields first...
Call ClearAllFields

'Get new contact name...
sInput = Trim$(InputBox$("Enter the name of this new Company or Contact.", "New..."))

'Nothing entered...
If sInput = "" Then Exit Sub

'Strip out invalid characters in the contact name...
If InStr(sInput, "'") > 0 Then
    Mid$(sInput, InStr(sInput, "'"), 1) = Chr$(32)
ElseIf InStr(sInput, "#") > 0 Then
    Mid$(sInput, InStr(sInput, "#"), 1) = Chr$(32)
ElseIf InStr(sInput, "$") > 0 Then
    Mid$(sInput, InStr(sInput, "$"), 1) = Chr$(32)
ElseIf InStr(sInput, "%") > 0 Then
    Mid$(sInput, InStr(sInput, "%"), 1) = Chr$(32)
ElseIf InStr(sInput, "&") > 0 Then
    Mid$(sInput, InStr(sInput, "&"), 1) = Chr$(32)
ElseIf InStr(sInput, "*") > 0 Then
    Mid$(sInput, InStr(sInput, "*"), 1) = Chr$(32)
ElseIf InStr(sInput, "(") > 0 Then
    Mid$(sInput, InStr(sInput, "("), 1) = Chr$(32)
ElseIf InStr(sInput, ")") > 0 Then
    Mid$(sInput, InStr(sInput, ")"), 1) = Chr$(32)
End If

'Create the new contact...
If OpenDB(DB, QuickRef.DBPassWord) = False Then Exit Sub
Set RS = DB.OpenRecordset("SELECT * FROM tblContacts WHERE Active = " & iActiveStatus, dbOpenDynaset)

RS.AddNew
RS!ContactName = sInput
RS!Active = iActiveStatus
RS.Update
RS.Close
DB.Close

'Load all contacts...
Call LoadAllContacts(lstContacts)

'Find this new contact and set the listindex...
For x = 0 To lstContacts.ListCount - 1
    If lstContacts.List(x) = sInput Then
        lstContacts.ListIndex = x
        If QuickRef.ContactName <> sInput Then
            QuickRef.ContactName = sInput
        End If
        txtAddress.SetFocus
        Exit For
    End If
Next x

iDirty = False
iContactNameHasChanged = False

End Sub
Private Sub lblNew_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

If Button = vbLeftButton Then
    imgNew.Picture = imgButton(1).Picture
    lblNew.ForeColor = QBColor(0)
End If

End Sub

Private Sub lblNew_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

Help.HelpText = "Click here to create a new contact."

End Sub
Private Sub lblNew_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

imgNew.Picture = imgButton(0).Picture
lblNew.ForeColor = lButtonForeColor

End Sub

Private Sub lblNotes_Click()

txtComments_DblClick

End Sub
Private Sub lblNotes_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

If Button = vbLeftButton Then
    imgNotes.Picture = imgButton(1).Picture
    lblNotes.ForeColor = QBColor(0)
End If

End Sub
Private Sub lblNotes_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

Help.HelpText = "Click here to view the notes in a larger window."

End Sub
Private Sub lblNotes_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

imgNotes.Picture = imgButton(0).Picture
lblNotes.ForeColor = lButtonForeColor

End Sub
Private Sub lblPrint_Click()

On Local Error Resume Next

Dim lFontSize As Single
Dim sFontName As String

'Exit if nothing to print...
If lstContacts.ListIndex = -1 Or lstContacts.ListCount = 0 Then
    MsgBox "First, select a contact to print.", vbInformation, "Print..."
    Exit Sub
End If

'Setup the printer...
lFontSize = Printer.FontSize
sFontName = Printer.Font.Name
Printer.Font.Name = "Courier New"
Printer.FontSize = 10

'Show the printer setup dialog box...
Dialog.ShowPrinter
If Err Then Exit Sub

'Print this record to the printer...
Printer.Print vbCrLf
Printer.Print Chr$(9); "Contact:              " & txtContactName
Printer.Print Chr$(9); "Address:              " & txtAddress
Printer.Print Chr$(9); "City, State, Zip:     " & txtCityStateZip
Printer.Print Chr$(9); "E-Mail Address:       " & txtEMailAddress
Printer.Print Chr$(9); "Work Phone:           " & txtWorkPhone
Printer.Print Chr$(9); "Fax:                  " & txtFax
Printer.Print Chr$(9); "Cell Phone:           " & txtCellPhone
Printer.Print Chr$(9); "IP Address:           " & txtIPAddress
Printer.Print Chr$(9); "Consulting Fee:       " & txtConsultingFee
Printer.Print Chr$(9); "Job Start Time:       " & txtJobStartTime
Printer.Print Chr$(9); "Job End Time:         " & txtJobEndTime
Printer.Print Chr$(9); "Comments:             " & txtComments
Printer.Print Chr$(9); String$(110, "-") & vbCrLf & vbCrLf

'Eject the page...
Printer.EndDoc

'Restore the printer...
Printer.FontSize = lFontSize
Printer.Font.Name = sFontName

'Msg...
MsgBox "Printing is complete...", vbInformation, "Print..."

End Sub
Private Sub lblPrint_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

If Button = vbLeftButton Then
    imgPrint.Picture = imgButton(1).Picture
    lblPrint.ForeColor = QBColor(0)
End If

End Sub

Private Sub lblPrint_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

Help.HelpText = "Click here to print this contact's information to the printer."

End Sub
Private Sub lblPrint_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

imgPrint.Picture = imgButton(0).Picture
lblPrint.ForeColor = lButtonForeColor

End Sub

Private Sub lblReload_Click()

On Local Error Resume Next

'Reload contact...
iDirty = False
lstContacts.Enabled = True
Call GetContactInfo

'Set focus...
Screen.ActiveControl.SelStart = 0
Screen.ActiveControl.SelLength = Len(Screen.ActiveControl)

End Sub
Private Sub lblReload_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

If Button = vbLeftButton Then
    imgReload.Picture = imgButton(1).Picture
    lblReload.ForeColor = QBColor(0)
End If

End Sub

Private Sub lblReload_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

Help.HelpText = "Click here to reload this contact, losing any changes made."

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

Help.HelpText = "Click here to save any changes you have made to this contact."

End Sub
Private Sub lblSave_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

imgSave.Picture = imgButton(0).Picture
lblSave.ForeColor = lButtonForeColor

End Sub

Private Sub lblVoiceMessages_Click()

frmVoiceMessages.Show
frmVoiceMessages.ZOrder

End Sub
Private Sub lblVoiceMessages_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

If Button = vbLeftButton Then
    imgVoiceMessages.Picture = imgButton(1).Picture
    lblVoiceMessages.ForeColor = QBColor(0)
End If

End Sub

Private Sub lblVoiceMessages_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

Help.HelpText = "Click here to record a voice message for this contact."

End Sub

Private Sub lblVoiceMessages_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

imgVoiceMessages.Picture = imgButton(0).Picture
lblVoiceMessages.ForeColor = lButtonForeColor

End Sub
Private Sub lstContacts_Click()

QuickRef.ContactName = lstContacts.List(lstContacts.ListIndex)
QuickRef.ContactID = lstContacts.ItemData(lstContacts.ListIndex)

Call GetContactInfo

End Sub
Function GetContactInfo() As Boolean

On Local Error GoTo GetContactInfoError

If OpenDB(DB, QuickRef.DBPassWord) = False Then Exit Function
Set RS = DB.OpenRecordset("SELECT * FROM tblContacts WHERE ContactName = '" & QuickRef.ContactName & "' AND Active = " & iActiveStatus & " AND ContactID = " & QuickRef.ContactID, dbOpenSnapshot)

'No records found...
If RS.RecordCount = 0 Then
    RS.Close
    DB.Close
    Exit Function
End If

'Clear all fields first...
Call ClearAllFields

'Records found...
If Not IsNull(RS!ContactName) Then
    txtContactName = RS!ContactName
    QuickRef.ContactName = RS!ContactName
End If
If Not IsNull(RS!Address) Then
    txtAddress = RS!Address
End If
If Not IsNull(RS!CityStateZip) Then
    txtCityStateZip = RS!CityStateZip
End If
If Not IsNull(RS!EMailAddress) Then
    txtEMailAddress = RS!EMailAddress
End If
If Not IsNull(RS!WorkPhone) Then
    txtWorkPhone = RS!WorkPhone
End If
If Not IsNull(RS!Fax) Then
    txtFax = RS!Fax
End If
If Not IsNull(RS!CellPhone) Then
    txtCellPhone = RS!CellPhone
End If
If Not IsNull(RS!IPAddress) Then
    txtIPAddress = RS!IPAddress
End If
If Not IsNull(RS!ConsultingFee) Then
    txtConsultingFee = RS!ConsultingFee
End If
If Not IsNull(RS!JobStartTime) Then
    txtJobStartTime = RS!JobStartTime
End If
If Not IsNull(RS!JobEndTime) Then
    txtJobEndTime = RS!JobEndTime
End If
If Not IsNull(RS!Comments) Then
    txtComments = RS!Comments
End If

iDirty = False
iContactNameHasChanged = False
RS.Close
DB.Close
Exit Function



GetContactInfoError:
    Call WriteToErrorLog(Me.Name, "GetContactInfoError", Error$, Err, True)
    Exit Function
    Resume Next

End Function
Private Sub lstContacts_KeyDown(KeyCode As Integer, Shift As Integer)

'Delete key...
If KeyCode = vbKeyDelete And lblDelete.Enabled = True Then
    lblDelete_Click
End If

End Sub
Private Sub lstContacts_KeyPress(KeyAscii As Integer)

'Enter Key...
If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}", True
    KeyAscii = 0
End If

End Sub
Private Sub lstContacts_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

'Drag...
If Button = vbLeftButton Then
    Timer2.Enabled = True
    If iStartDragAndDrop = True Then
        Timer2.Enabled = False
        iStartDragAndDrop = False
        lstContacts.DragIcon = imgDragDrop(0).Picture
        lstContacts.Drag
    End If
End If

End Sub

Private Sub lstContacts_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

Help.HelpText = "Click here to view information on a contact."

End Sub
Private Sub Timer1_Timer()

On Local Error Resume Next

Dim x As Long
Dim iTempDirty As Boolean

'Set Colors...
If QuickRef.UpdateColors = True Then
    Call LoadProgramColors
    QuickRef.UpdateColors = False
    For x = 0 To Forms.Count - 1
        Call SetColors(Forms(x))
    Next x
End If

'Notes...
If QuickRef.UpdateNotes And QuickRef.CallingForm <> "Contactor" Then
    iTempDirty = iDirty
    Call GetNotes(txtComments)
    iDirty = iTempDirty
End If

'Contacts Listbox...
If lstContacts.Enabled = True And iDirty = True Then
    lstContacts.Enabled = False
    chkActiveStatus.Enabled = False
    lblActiveStatus.Enabled = False
ElseIf lstContacts.Enabled = False And iDirty = False Then
    lstContacts.Enabled = True
    chkActiveStatus.Enabled = True
    lblActiveStatus.Enabled = True
End If

'Delete...
If lblDelete.Enabled = True And lstContacts.ListIndex = -1 Then
    lblDelete.Enabled = False
    imgDelete.Enabled = False
ElseIf lblDelete.Enabled = False And lstContacts.ListIndex > -1 Then
    lblDelete.Enabled = True
    imgDelete.Enabled = True
End If

'New...
If lblNew.Enabled = True And iDirty = True Then
    lblNew.Enabled = False
    imgNew.Enabled = False
ElseIf lblNew.Enabled = False And iDirty = False Then
    lblNew.Enabled = True
    imgNew.Enabled = True
End If

'Save...
If lblSave.Enabled = True And iDirty = False Then
    lblSave.Enabled = False
    imgSave.Enabled = False
ElseIf lblSave.Enabled = False And iDirty = True Then
    lblSave.Enabled = True
    imgSave.Enabled = True
End If

'Reload...
If lblReload.Enabled = True And iDirty = False Then
    lblReload.Enabled = False
    imgReload.Enabled = False
ElseIf lblReload.Enabled = False And iDirty = True Then
    lblReload.Enabled = True
    imgReload.Enabled = True
End If

'Print...
If lblPrint.Enabled = True And lstContacts.ListIndex = -1 Then
    lblPrint.Enabled = False
    imgPrint.Enabled = False
ElseIf lblPrint.Enabled = False And lstContacts.ListIndex > -1 Then
    lblPrint.Enabled = True
    imgPrint.Enabled = True
End If

'Mark Active Button...
If lblMarkActive.Visible = True And iActiveStatus = True Then
    lblMarkActive.Visible = False
    imgMarkActive.Visible = False
ElseIf lblMarkActive.Visible = False And iActiveStatus = False Then
    lblMarkActive.Visible = True
    imgMarkActive.Visible = True
End If
If lblMarkActive.Visible = True Then
    If lstContacts.ListIndex = -1 And lblMarkActive.Enabled = True Then
        lblMarkActive.Enabled = False
        imgMarkActive.Enabled = False
    ElseIf lstContacts.ListIndex = 0 And lblMarkActive.Enabled = False Then
        lblMarkActive.Enabled = True
        imgMarkActive.Enabled = True
    End If
End If

'Voice Messages...
If Dir$(App.Path & "\VM\" & lstContacts.List(lstContacts.ListIndex) & "\*.Wav") = "" Then
    imgThereAreVoiceMessages.Picture = imgRedLight(0).Picture
ElseIf Dir$(App.Path & "\VM\" & lstContacts.List(lstContacts.ListIndex) & "\*.Wav") <> "" Then
    imgThereAreVoiceMessages.Picture = imgRedLight(1).Picture
End If

End Sub
Private Sub Timer2_Timer()

iStartDragAndDrop = True

End Sub
Private Sub txtAddress_Change()

iDirty = True

End Sub

Private Sub txtAddress_GotFocus()

txtAddress.SelStart = 0
txtAddress.SelLength = Len(txtAddress)

End Sub

Private Sub txtAddress_KeyDown(KeyCode As Integer, Shift As Integer)

'Down key...
If KeyCode = vbKeyDown Then
    SendKeys "{TAB}", True
    KeyCode = 0
End If

'Up key...
If KeyCode = vbKeyUp Then
    SendKeys "+{TAB}", True
    KeyCode = 0
End If

End Sub
Private Sub txtAddress_KeyPress(KeyAscii As Integer)

'Enter Key...
If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}", True
    KeyAscii = 0
End If

End Sub

Private Sub txtAddress_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

'Drag...
If Button = vbLeftButton Then
    Timer2.Enabled = True
    If iStartDragAndDrop = True Then
        Timer2.Enabled = False
        iStartDragAndDrop = False
        txtAddress.DragIcon = imgDragDrop(0).Picture
        txtAddress.Drag
    End If
End If

End Sub

Private Sub txtAddress_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

Help.HelpText = txtAddress

End Sub
Private Sub txtCellPhone_Change()

iDirty = True

End Sub

Private Sub txtCellPhone_GotFocus()

txtCellPhone.SelStart = 0
txtCellPhone.SelLength = Len(txtCellPhone)

End Sub

Private Sub txtCellPhone_KeyDown(KeyCode As Integer, Shift As Integer)

'Down key...
If KeyCode = vbKeyDown Then
    SendKeys "{TAB}", True
    KeyCode = 0
End If

'Up key...
If KeyCode = vbKeyUp Then
    SendKeys "+{TAB}", True
    KeyCode = 0
End If

End Sub
Private Sub txtCellPhone_KeyPress(KeyAscii As Integer)

'Enter Key...
If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}", True
    KeyAscii = 0
End If

End Sub

Private Sub txtCellPhone_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

'Drag...
If Button = vbLeftButton Then
    Timer2.Enabled = True
    If iStartDragAndDrop = True Then
        Timer2.Enabled = False
        iStartDragAndDrop = False
        txtCellPhone.DragIcon = imgDragDrop(0).Picture
        txtCellPhone.Drag
    End If
End If

End Sub

Private Sub txtCellPhone_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

Help.HelpText = txtCellPhone

End Sub
Private Sub txtCityStateZip_Change()

iDirty = True

End Sub

Private Sub txtCityStateZip_GotFocus()

txtCityStateZip.SelStart = 0
txtCityStateZip.SelLength = Len(txtCityStateZip)

End Sub

Private Sub txtCityStateZip_KeyDown(KeyCode As Integer, Shift As Integer)

'Down key...
If KeyCode = vbKeyDown Then
    SendKeys "{TAB}", True
    KeyCode = 0
End If

'Up key...
If KeyCode = vbKeyUp Then
    SendKeys "+{TAB}", True
    KeyCode = 0
End If

End Sub
Private Sub txtCityStateZip_KeyPress(KeyAscii As Integer)

'Enter Key...
If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}", True
    KeyAscii = 0
End If

End Sub

Private Sub txtCityStateZip_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

'Drag...
If Button = vbLeftButton Then
    Timer2.Enabled = True
    If iStartDragAndDrop = True Then
        Timer2.Enabled = False
        iStartDragAndDrop = False
        txtCityStateZip.DragIcon = imgDragDrop(0).Picture
        txtCityStateZip.Drag
    End If
End If

End Sub

Private Sub txtCityStateZip_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

Help.HelpText = txtCityStateZip

End Sub
Private Sub txtComments_Change()

QuickRef.CallingForm = "Contactor"
QuickRef.NotesHaveChanged = True
iDirty = True

End Sub

Private Sub txtComments_DblClick()

'Pass in the notes if not saved...
If iDirty Then
    QuickRef.PassNotes = True
    frmNotes.txtNotes = txtComments
    frmNotes.Show
    frmNotes.ZOrder
Else
    QuickRef.PassNotes = False
    frmNotes.Show
    frmNotes.ZOrder
End If

End Sub
Private Sub txtComments_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

'Drag...
If Button = vbLeftButton Then
    Timer2.Enabled = True
    If iStartDragAndDrop = True Then
        Timer2.Enabled = False
        iStartDragAndDrop = False
        txtComments.DragIcon = imgDragDrop(0).Picture
        txtComments.Drag
    End If
End If

End Sub

Private Sub txtComments_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

Help.HelpText = txtComments

End Sub
Private Sub txtComments_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, Y As Single)

'File is a word file, so start word with this file...
If LCase$(Right$(Data.Files(1), 3)) = "doc" Then
    Call StartWord(Data.Files(1))
Else
    Call OpenFile(txtComments, Data.Files(1))
End If

End Sub

Private Sub txtConsultingFee_GotFocus()

txtConsultingFee.SelStart = 0
txtConsultingFee.SelLength = Len(txtConsultingFee)

End Sub

Private Sub txtConsultingFee_KeyDown(KeyCode As Integer, Shift As Integer)

'Down key...
If KeyCode = vbKeyDown Then
    SendKeys "{TAB}", True
    KeyCode = 0
End If

'Up key...
If KeyCode = vbKeyUp Then
    SendKeys "+{TAB}", True
    KeyCode = 0
End If

End Sub
Private Sub txtConsultingFee_KeyPress(KeyAscii As Integer)

'Enter Key...
If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}", True
    KeyAscii = 0
End If

End Sub

Private Sub txtConsultingFee_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

'Drag...
If Button = vbLeftButton Then
    Timer2.Enabled = True
    If iStartDragAndDrop = True Then
        Timer2.Enabled = False
        iStartDragAndDrop = False
        txtConsultingFee.DragIcon = imgDragDrop(0).Picture
        txtConsultingFee.Drag
    End If
End If

End Sub

Private Sub txtConsultingFee_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

Help.HelpText = txtConsultingFee

End Sub
Private Sub txtContactName_Change()

iDirty = True
iContactNameHasChanged = True

End Sub
Private Sub txtConsultingFee_Change()

iDirty = True

End Sub

Private Sub txtContactName_GotFocus()

txtContactName.SelStart = 0
txtContactName.SelLength = Len(txtContactName)

End Sub

Private Sub txtContactName_KeyDown(KeyCode As Integer, Shift As Integer)

'Down key...
If KeyCode = vbKeyDown Then
    SendKeys "{TAB}", True
    KeyCode = 0
End If

'Up key...
If KeyCode = vbKeyUp Then
    lstContacts.SetFocus
    KeyCode = 0
End If

End Sub
Private Sub txtContactName_KeyPress(KeyAscii As Integer)

'Enter Key...
If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}", True
    KeyAscii = 0
End If

End Sub

Private Sub txtContactName_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

'Drag...
If Button = vbLeftButton Then
    Timer2.Enabled = True
    If iStartDragAndDrop = True Then
        Timer2.Enabled = False
        iStartDragAndDrop = False
        txtContactName.DragIcon = imgDragDrop(0).Picture
        txtContactName.Drag
    End If
End If

End Sub

Private Sub txtContactName_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

Help.HelpText = txtContactName

End Sub
Private Sub txtEMailAddress_GotFocus()

txtEMailAddress.SelStart = 0
txtEMailAddress.SelLength = Len(txtEMailAddress)

End Sub

Private Sub txtEMailAddress_KeyDown(KeyCode As Integer, Shift As Integer)

'Down key...
If KeyCode = vbKeyDown Then
    SendKeys "{TAB}", True
    KeyCode = 0
End If

'Up key...
If KeyCode = vbKeyUp Then
    SendKeys "+{TAB}", True
    KeyCode = 0
End If

End Sub
Private Sub txtEMailAddress_KeyPress(KeyAscii As Integer)

'Enter Key...
If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}", True
    KeyAscii = 0
End If

End Sub

Private Sub txtEMailAddress_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

'Drag...
If Button = vbLeftButton Then
    Timer2.Enabled = True
    If iStartDragAndDrop = True Then
        Timer2.Enabled = False
        iStartDragAndDrop = False
        txtEMailAddress.DragIcon = imgDragDrop(0).Picture
        txtEMailAddress.Drag
    End If
End If

End Sub

Private Sub txtEMailAddress_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

Help.HelpText = txtEMailAddress & ". You can double click in this field to send an email to this contact."

End Sub
Private Sub txtFax_Change()

iDirty = True

End Sub
Private Sub txtEMailAddress_Change()

iDirty = True

End Sub

Private Sub txtFax_GotFocus()

txtFax.SelStart = 0
txtFax.SelLength = Len(txtFax)

End Sub

Private Sub txtFax_KeyDown(KeyCode As Integer, Shift As Integer)

'Down key...
If KeyCode = vbKeyDown Then
    SendKeys "{TAB}", True
    KeyCode = 0
End If

'Up key...
If KeyCode = vbKeyUp Then
    SendKeys "+{TAB}", True
    KeyCode = 0
End If

End Sub
Private Sub txtFax_KeyPress(KeyAscii As Integer)

'Enter Key...
If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}", True
    KeyAscii = 0
End If

End Sub

Private Sub txtFax_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

'Drag...
If Button = vbLeftButton Then
    Timer2.Enabled = True
    If iStartDragAndDrop = True Then
        Timer2.Enabled = False
        iStartDragAndDrop = False
        txtFax.DragIcon = imgDragDrop(0).Picture
        txtFax.Drag
    End If
End If

End Sub

Private Sub txtFax_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

Help.HelpText = txtFax

End Sub
Private Sub txtIPAddress_Change()

iDirty = True

End Sub

Private Sub txtIPAddress_GotFocus()

txtIPAddress.SelStart = 0
txtIPAddress.SelLength = Len(txtIPAddress)

End Sub

Private Sub txtIPAddress_KeyDown(KeyCode As Integer, Shift As Integer)

'Down key...
If KeyCode = vbKeyDown Then
    SendKeys "{TAB}", True
    KeyCode = 0
End If

'Up key...
If KeyCode = vbKeyUp Then
    SendKeys "+{TAB}", True
    KeyCode = 0
End If

End Sub
Private Sub txtIPAddress_KeyPress(KeyAscii As Integer)

'Enter Key...
If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}", True
    KeyAscii = 0
End If

End Sub

Private Sub txtIPAddress_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

'Drag...
If Button = vbLeftButton Then
    Timer2.Enabled = True
    If iStartDragAndDrop = True Then
        Timer2.Enabled = False
        iStartDragAndDrop = False
        txtIPAddress.DragIcon = imgDragDrop(0).Picture
        txtIPAddress.Drag
    End If
End If

End Sub

Private Sub txtIPAddress_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

Help.HelpText = txtIPAddress

End Sub
Private Sub txtJobEndTime_Change()

iDirty = True

End Sub

Private Sub txtJobEndTime_GotFocus()

txtJobEndTime.SelStart = 0
txtJobEndTime.SelLength = Len(txtJobEndTime)

End Sub

Private Sub txtJobEndTime_KeyDown(KeyCode As Integer, Shift As Integer)

'Down key...
If KeyCode = vbKeyDown Then
    SendKeys "{TAB}", True
    KeyCode = 0
End If

'Up key...
If KeyCode = vbKeyUp Then
    SendKeys "+{TAB}", True
    KeyCode = 0
End If

End Sub
Private Sub txtJobEndTime_KeyPress(KeyAscii As Integer)

'Enter Key...
If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}", True
    KeyAscii = 0
End If

End Sub

Private Sub txtJobEndTime_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

'Drag...
If Button = vbLeftButton Then
    Timer2.Enabled = True
    If iStartDragAndDrop = True Then
        Timer2.Enabled = False
        iStartDragAndDrop = False
        txtJobEndTime.DragIcon = imgDragDrop(0).Picture
        txtJobEndTime.Drag
    End If
End If

End Sub

Private Sub txtJobEndTime_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

Help.HelpText = "The time a job ends for this contact."

End Sub
Private Sub txtJobStartTime_Change()

iDirty = True

End Sub

Private Sub txtJobStartTime_GotFocus()

txtJobStartTime.SelStart = 0
txtJobStartTime.SelLength = Len(txtJobStartTime)

End Sub

Private Sub txtJobStartTime_KeyDown(KeyCode As Integer, Shift As Integer)

'Down key...
If KeyCode = vbKeyDown Then
    SendKeys "{TAB}", True
    KeyCode = 0
End If

'Up key...
If KeyCode = vbKeyUp Then
    SendKeys "+{TAB}", True
    KeyCode = 0
End If

End Sub
Private Sub txtJobStartTime_KeyPress(KeyAscii As Integer)

'Enter Key...
If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}", True
    KeyAscii = 0
End If

End Sub

Private Sub txtJobStartTime_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

'Drag...
If Button = vbLeftButton Then
    Timer2.Enabled = True
    If iStartDragAndDrop = True Then
        Timer2.Enabled = False
        iStartDragAndDrop = False
        txtJobStartTime.DragIcon = imgDragDrop(0).Picture
        txtJobStartTime.Drag
    End If
End If

End Sub

Private Sub txtJobStartTime_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

Help.HelpText = "The time a job starts for this contact."

End Sub
Private Sub txtWorkPhone_Change()

iDirty = True

End Sub

Private Sub txtWorkPhone_GotFocus()

txtWorkPhone.SelStart = 0
txtWorkPhone.SelLength = Len(txtWorkPhone)

End Sub

Private Sub txtWorkPhone_KeyDown(KeyCode As Integer, Shift As Integer)

'Down key...
If KeyCode = vbKeyDown Then
    SendKeys "{TAB}", True
    KeyCode = 0
End If

'Up key...
If KeyCode = vbKeyUp Then
    SendKeys "+{TAB}", True
    KeyCode = 0
End If

End Sub
Private Sub txtWorkPhone_KeyPress(KeyAscii As Integer)

'Enter Key...
If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}", True
    KeyAscii = 0
End If

End Sub

Private Sub txtWorkPhone_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

'Drag...
If Button = vbLeftButton Then
    Timer2.Enabled = True
    If iStartDragAndDrop = True Then
        Timer2.Enabled = False
        iStartDragAndDrop = False
        txtWorkPhone.DragIcon = imgDragDrop(0).Picture
        txtWorkPhone.Drag
    End If
End If

End Sub

Private Sub txtWorkPhone_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

Help.HelpText = txtWorkPhone

End Sub
