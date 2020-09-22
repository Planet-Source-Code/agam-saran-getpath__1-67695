VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmGetPath 
   BackColor       =   &H00B24801&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "GetPath"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6825
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   6825
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog Dialogs 
      Left            =   5760
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ComboBox cmbDemo 
      Height          =   315
      ItemData        =   "Form1.frx":000C
      Left            =   360
      List            =   "Form1.frx":001C
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   1800
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   5520
      TabIndex        =   14
      Top             =   5520
      Width           =   1215
   End
   Begin VB.CommandButton cmdEvaluate 
      Caption         =   "&Evaluate"
      Height          =   375
      Left            =   4200
      TabIndex        =   13
      Top             =   5520
      Width           =   1215
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "&Browse..."
      Height          =   375
      Left            =   5040
      TabIndex        =   3
      Top             =   1800
      Width           =   1335
   End
   Begin VB.TextBox txtPath 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000007&
      Height          =   285
      Left            =   360
      TabIndex        =   2
      Text            =   "C:\This\Is\To\Test\GetPath\File.exe"
      Top             =   1440
      Width           =   6015
   End
   Begin VB.Image imgBullet 
      Height          =   120
      Left            =   240
      Picture         =   "Form1.frx":0040
      Top             =   5640
      Width           =   120
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "Created by Agam Saran"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   480
      TabIndex        =   26
      Top             =   5595
      Width           =   2310
   End
   Begin VB.Shape Shape8 
      BackColor       =   &H00EEEEEE&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00BFBFBF&
      Height          =   375
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   5520
      Width           =   2895
   End
   Begin VB.Label lblPath 
      AutoSize        =   -1  'True
      BackColor       =   &H00EEEEEE&
      Caption         =   "Value"
      Height          =   195
      Index           =   9
      Left            =   2160
      TabIndex        =   25
      Top             =   4680
      Width           =   405
   End
   Begin VB.Label lblPath 
      AutoSize        =   -1  'True
      BackColor       =   &H00EEEEEE&
      Caption         =   "Value"
      Height          =   195
      Index           =   8
      Left            =   2520
      TabIndex        =   24
      Top             =   4440
      Width           =   405
   End
   Begin VB.Label lblPath 
      AutoSize        =   -1  'True
      BackColor       =   &H00EEEEEE&
      Caption         =   "Value"
      Height          =   195
      Index           =   7
      Left            =   1440
      TabIndex        =   23
      Top             =   4200
      Width           =   405
   End
   Begin VB.Label lblPath 
      AutoSize        =   -1  'True
      BackColor       =   &H00EEEEEE&
      Caption         =   "Value"
      Height          =   195
      Index           =   6
      Left            =   1440
      TabIndex        =   22
      Top             =   3960
      Width           =   405
   End
   Begin VB.Label lblPath 
      AutoSize        =   -1  'True
      BackColor       =   &H00EEEEEE&
      Caption         =   "Value"
      Height          =   195
      Index           =   5
      Left            =   960
      TabIndex        =   21
      Top             =   3720
      Width           =   405
   End
   Begin VB.Label lblPath 
      AutoSize        =   -1  'True
      BackColor       =   &H00EEEEEE&
      Caption         =   "Value"
      Height          =   195
      Index           =   4
      Left            =   1200
      TabIndex        =   20
      Top             =   3480
      Width           =   405
   End
   Begin VB.Label lblPath 
      AutoSize        =   -1  'True
      BackColor       =   &H00EEEEEE&
      Caption         =   "Value"
      Height          =   195
      Index           =   3
      Left            =   1560
      TabIndex        =   19
      Top             =   3240
      Width           =   405
   End
   Begin VB.Label lblPath 
      AutoSize        =   -1  'True
      BackColor       =   &H00EEEEEE&
      Caption         =   "Value"
      Height          =   195
      Index           =   2
      Left            =   1320
      TabIndex        =   18
      Top             =   3000
      Width           =   405
   End
   Begin VB.Label lblPath 
      AutoSize        =   -1  'True
      BackColor       =   &H00EEEEEE&
      Caption         =   "Value"
      Height          =   195
      Index           =   1
      Left            =   1320
      TabIndex        =   17
      Top             =   2760
      Width           =   405
   End
   Begin VB.Label lblPath 
      AutoSize        =   -1  'True
      BackColor       =   &H00EEEEEE&
      Caption         =   "Extractions"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   0
      Left            =   2760
      TabIndex        =   16
      Top             =   2580
      Width           =   1110
   End
   Begin VB.Line lneSep 
      BorderColor     =   &H00CC8859&
      X1              =   -840
      X2              =   7920
      Y1              =   5400
      Y2              =   5400
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackColor       =   &H00EEEEEE&
      Caption         =   "DriveAndFirstFolder:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   360
      TabIndex        =   12
      Top             =   4680
      Width           =   1740
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00EEEEEE&
      Caption         =   "LastFolderAndFileName:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   360
      TabIndex        =   11
      Top             =   4440
      Width           =   2070
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00EEEEEE&
      Caption         =   "FirstFolder:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   360
      TabIndex        =   10
      Top             =   4200
      Width           =   960
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00EEEEEE&
      Caption         =   "LastFolder:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   360
      TabIndex        =   9
      Top             =   3960
      Width           =   960
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00EEEEEE&
      Caption         =   "Drive:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   360
      TabIndex        =   8
      Top             =   3720
      Width           =   525
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00EEEEEE&
      Caption         =   "FilePath:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   360
      TabIndex        =   7
      Top             =   3480
      Width           =   765
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00EEEEEE&
      Caption         =   "FileExtension:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   360
      TabIndex        =   6
      Top             =   3240
      Width           =   1200
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00EEEEEE&
      Caption         =   "JustName:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   360
      TabIndex        =   5
      Top             =   3000
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00EEEEEE&
      Caption         =   "FileName:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   360
      TabIndex        =   4
      Top             =   2760
      Width           =   855
   End
   Begin VB.Shape Shape7 
      BackColor       =   &H00EEEEEE&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00EEEEEE&
      Height          =   855
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   4200
      Width           =   6255
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00EEEEEE&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00EEEEEE&
      Height          =   855
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   3720
      Width           =   6255
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00EEEEEE&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00EEEEEE&
      Height          =   855
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   3120
      Width           =   6255
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H00EEEEEE&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00EEEEEE&
      Height          =   855
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   2520
      Width           =   6255
   End
   Begin VB.Label lblApp 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "GetPath"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   525
      Left            =   1320
      TabIndex        =   0
      Top             =   315
      Width           =   1950
   End
   Begin VB.Image imgHistory 
      Height          =   735
      Left            =   360
      Picture         =   "Form1.frx":0142
      ToolTipText     =   "About"
      Top             =   200
      Width           =   765
   End
   Begin VB.Label lblShadow 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "GetPath"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00BD6234&
      Height          =   525
      Left            =   1350
      TabIndex        =   1
      Top             =   345
      Width           =   1950
   End
   Begin VB.Shape shpName 
      BackColor       =   &H00CC8859&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00CC8859&
      FillColor       =   &H00CC8859&
      Height          =   855
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   6495
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00CC8859&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00CC8859&
      FillColor       =   &H00FFFFFF&
      Height          =   4215
      Left            =   960
      Top             =   1080
      Width           =   4575
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00CC8859&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00CC8859&
      FillColor       =   &H00FFFFFF&
      Height          =   4215
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00CC8859&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00CC8859&
      FillColor       =   &H00FFFFFF&
      Height          =   4215
      Left            =   5280
      Shape           =   4  'Rounded Rectangle
      Top             =   1080
      Width           =   1335
   End
End
Attribute VB_Name = "frmGetPath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Created by Agam Saran

Private Declare Function SHAutoComplete Lib "shlwapi.dll" (ByVal hWndEdit As Long, ByVal dwFlags As AutoCompleteFlags) As Integer
Private Enum AutoCompleteFlags
    SHACF_DEFAULT = 0
    SHACF_FILESYSTEM = 1
    SHACF_URLHISTORY = 2
    SHACF_URLMRU = 4
    SHACF_USETAB = 8
    SHACF_URLALL = (SHACF_URLHISTORY Or SHACF_URLMRU)
    SHACF_FILESYS_ONLY = 10
    SHACF_FILESYS_DIRS = 20
    SHACF_AUTOSUGGEST_FORCE_ON = 10000000
    SHACF_AUTOSUGGEST_FORCE_OFF = 20000000
    SHACF_AUTOAPPEND_FORCE_ON = 40000000
    SHACF_AUTOAPPEND_FORCE_OFF = 80000000
End Enum



Private Sub cmbDemo_Click()
Select Case cmbDemo.ListIndex
    Case 0
        txtPath.Text = "C:\This\Is\To\Test\GetPath\File.exe"
    Case 1
        txtPath.Text = "\Relative\Paths\Are\Supported\Too\File.exe"
    Case 2
        txtPath.Text = "C:\Look\At\The\FileExtension\Of\This\Path.exe\"
    Case 3
        txtPath.Text = "C:\Temp\Is\A\File\Not\A\Folder\Temp"
End Select
End Sub

Private Sub cmdBrowse_Click()
On Error GoTo ErrHandler
Dialogs.DialogTitle = "Browse..."
Dialogs.Flags = &H4
Dialogs.Filter = "All Files (*.*)|*.*"
Dialogs.ShowOpen
txtPath.Text = Dialogs.FileName
Dialogs.FileName = ""

ErrHandler:
    Exit Sub
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdEvaluate_Click()
Dim i As Integer
For i = 1 To 9
    lblPath(i).Caption = GetPath(txtPath.Text, i)
Next
End Sub

Private Sub Form_Load()
    SHAutoComplete txtPath.hWnd, _
    SHACF_FILESYS_DIRS Or _
    SHACF_AUTOSUGGEST_FORCE_ON Or _
    SHACF_AUTOAPPEND_FORCE_ON Or _
    SHACF_FILESYSTEM
    cmbDemo.ListIndex = 0
    
    cmdEvaluate_Click
End Sub

