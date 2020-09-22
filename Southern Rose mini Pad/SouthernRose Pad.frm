VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Rose 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6255
   ClientLeft      =   3270
   ClientTop       =   1815
   ClientWidth     =   6525
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   6525
   Begin VB.CheckBox chkUnderline 
      Caption         =   "Underline"
      Height          =   195
      Left            =   2040
      TabIndex        =   8
      Top             =   5640
      Width           =   975
   End
   Begin VB.Timer TimeDate 
      Interval        =   1
      Left            =   5880
      Top             =   4440
   End
   Begin VB.Timer FlashWin 
      Interval        =   1000
      Left            =   5760
      Top             =   4440
   End
   Begin MSComDlg.CommonDialog Srdialog 
      Left            =   5640
      Top             =   4440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CheckBox chkItalic 
      Caption         =   "Italic"
      Height          =   255
      Left            =   2040
      TabIndex        =   5
      Top             =   5400
      Width           =   975
   End
   Begin VB.CheckBox chkBold 
      Caption         =   "Bold"
      Height          =   255
      Left            =   2040
      TabIndex        =   4
      Top             =   5160
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   3735
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   5895
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "<===another way to to fonts"
      Height          =   390
      Left            =   4440
      TabIndex        =   9
      Top             =   5280
      Width           =   1935
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblDate 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   5640
      TabIndex        =   7
      Top             =   6000
      Width           =   105
   End
   Begin VB.Label lblTime 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   6000
      Width           =   105
   End
   Begin VB.Label lblMail 
      Caption         =   "Cpvio@netscape.net"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   2520
      MousePointer    =   10  'Up Arrow
      TabIndex        =   3
      ToolTipText     =   "E-Mail To: Cpvio"
      Top             =   5880
      Width           =   1935
   End
   Begin VB.Label lblOk 
      Height          =   375
      Left            =   2880
      TabIndex        =   2
      Top             =   5160
      Width           =   1215
   End
   Begin VB.Label lblTitle 
      Height          =   975
      Left            =   2040
      TabIndex        =   1
      Top             =   4080
      Width           =   3495
   End
   Begin VB.Image Image1 
      Height          =   1665
      Left            =   240
      Picture         =   "SouthernRose Pad.frx":0000
      Top             =   4080
      Width           =   1695
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuSpace 
         Caption         =   "-"
      End
      Begin VB.Menu mnuNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSpace2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrintset 
         Caption         =   "P&rinter Setup"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "&Print"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuSpace3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuQuit 
         Caption         =   "&Quit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuFont 
      Caption         =   "FontSiz&e"
      Begin VB.Menu mnu1 
         Caption         =   "&Text Fonts Size"
      End
      Begin VB.Menu mnu2 
         Caption         =   "&Printer Fonts Size"
      End
   End
End
Attribute VB_Name = "Rose"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function FlashWindow Lib "user32" (ByVal HWND As Long, ByVal dumb As Long) As Long
Dim fnum

Private Sub chkUnderline_Click()
'If The Check Box Is Not Checked
If chkUnderline.Value = 0 Then
'Then The Text1.Text Is Not Underlined
Text1.FontUnderline = False
'If The Check Box Is Checked
ElseIf chkUnderline.Value = 1 Then
'Then The Text1.Text Is Underlined
Text1.FontUnderline = True
End If
End Sub

Private Sub chkBold_Click()
'If The Check Box Is Not Checked
If chkBold.Value = 0 Then
'Then The Text1.Text Is Not Bold
Text1.FontBold = False
'If The Check Box Is checked
ElseIf chkBold.Value = 1 Then
'Then The Text1.Text Is Bold
Text1.FontBold = True
End If
End Sub

Private Sub chkItalic_Click()
'If The Check Box Is Not Checked
If chkItalic.Value = 0 Then
'Then The Text1.Text Is Not Italic
Text1.FontItalic = False
'If The Check Box Is Checked
ElseIf chkItalic.Value = 1 Then
'Then The Text1.Text Is Italic
Text1.FontItalic = True
End If
End Sub

Private Sub Form_Load()
Rose.Caption = "Southern Rose Pad by Cpvio"
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim X As String
'Calls The vbQuestion MessageBox with vbYesNo Buttons And Forms Caption When You Click The X Button
X = MsgBox("Do You Wish To Exit?", vbQuestion + vbYesNo, "Southern Rose Mini Pad")
'If Yes Is Clicked On The MessageBox
If X = vbYes Then
'It Will End The Program Or Form
End
Else
'If No Is Clicked Program Or Form Stays Running
Cancel = 1
End If
End Sub

Private Sub lblMail_Click()
'Form Will Minimize To Taskbar
Me.WindowState = vbMinimized
'Calls The Mail Open and Insert
 ShellExecute HWND, "open", "mailto:Cpvio@netscape.net?body=About Your Southern Rose Mini Notepad", vbNullString, vbNullString, SW_SHOW
End Sub

Private Sub lblOk_Click()
'Clears lblTitles Caption
lblTitle.Caption = ""
'Clears lblOk Caption
lblOk.Caption = ""
End Sub

Private Sub mnu1_Click()
'Sets The Text1 Fonts
On Error Resume Next
    Srdialog.Flags = 1
    Srdialog.FontName = Rose.Text1.FontName
    Srdialog.FontSize = Rose.Text1.FontSize
    Srdialog.FontBold = Rose.Text1.FontBold
    Srdialog.FontItalic = Rose.Text1.FontItalic
    Srdialog.ShowFont
    If Err = 0 Then
        Rose.Text1.FontName = Srdialog.FontName
        Rose.Text1.FontSize = Srdialog.FontSize
        Rose.Text1.FontBold = Srdialog.FontBold
        Rose.Text1.FontItalic = Srdialog.FontItalic
    End If
End Sub

Private Sub mnu2_Click()
' Set The Printer Fonts.
    On Error Resume Next
    Srdialog.Flags = 2
    Srdialog.FontName = Printer.FontName
    Srdialog.FontSize = Printer.FontSize
    Srdialog.FontBold = Printer.FontBold
    Srdialog.FontItalic = Printer.FontItalic
    Srdialog.ShowFont
    If Err = 0 Then
        Printer.FontName = Srdialog.FontName
        Printer.FontSize = Srdialog.FontSize
        Printer.FontBold = Srdialog.FontBold
        Printer.FontItalic = Srdialog.FontItalic
    End If

End Sub

Private Sub mnuAbout_Click()
'Makes lblTitle Font Bold
lblTitle.FontBold = True
'Sets The Caption in lblTitle
lblTitle.Caption = "Southern Rose Mini Pad Version 1.0.0 Coded By Cpvio From House of Evil`97-2003 Trademark All Rights Reserved"
'Makes lblOk Font Bold
lblOk.FontBold = True
'Sets lblOk Caption To Ok
lblOk.Caption = "OK"
End Sub
Private Sub mnuNew_Click()
'Clears the textbox
Text1.Text = ""
End Sub

Private Sub mnuOpen_Click()
On Error Resume Next
'Calls the Open DialogBox
Srdialog.Filter = "Text Open (*.txt)|*.txt|Rich Text (*.rtf)|*.rtf|All Files (*.*)|*.*"
    Srdialog.ShowOpen
    fnum = FreeFile

    Open Srdialog.filename For Input As #1
    Text1 = Input(LOF(fnum), #fnum)
    Close #1
    
    Text1.Text = Text1
End Sub

Private Sub mnuPrint_Click()
'Prints The Textbox Text
Printer.Print
    Printer.Print Text1.Text
    Printer.EndDoc
End Sub

Private Sub mnuPrintset_Click()
'Calls The Printer setup
On Error Resume Next
Srdialog.Flags = &H40
Srdialog.ShowPrinter

End Sub

Private Sub mnuQuit_Click()
'Terminates current app.
End
End Sub

Private Sub mnuSave_Click()
On Error Resume Next
'Calls the Saveas DialogBox
Srdialog.Filter = "Text Save (*.txt)|*.txt|Rich Text (*.rtf)|*.rtf|All Files (*.*)|*.*"
    Srdialog.ShowSave

    fnum = FreeFile

    Open Srdialog.filename For Output As #1
    Print #1, Text1.Text
    Close #1
End Sub

Private Sub FlashWin_Timer()
'This Will Flash The CaptionBar When Minimized and Stop Flashing When Restored
Dim vio
    'Set Timer Interval To 1000 Or Whatever You Want
    'Remember Lower The Number Faster The Timer Fires/Higher The Number Slower The Timer Fires
    If Rose.WindowState = 1 Then
    vio = FlashWindow(Rose.HWND, True)
    ElseIf Rose.WindowState = 0 Then
    
    End If
End Sub

Private Sub TimeDate_Timer()
'This Calls Your Computers Current Time And Date From The System
'In Seperate Labels
'But If You Want Them In The Same Label Call This
'Dim Today
'Today = Now
'lblTimeDate = Today
Dim MyTime
Dim MyDate
MyTime = Time
MyDate = Date
lblTime.Caption = MyTime
lblDate.Caption = MyDate

End Sub
