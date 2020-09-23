VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmmain 
   BackColor       =   &H80000004&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Compare Utility"
   ClientHeight    =   6390
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7650
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6390
   ScaleWidth      =   7650
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7200
      Top             =   5760
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.CommandButton cmdnc 
      BackColor       =   &H8000000B&
      Caption         =   "NewCompare"
      Height          =   255
      Left            =   2640
      TabIndex        =   17
      Top             =   5160
      Width           =   1095
   End
   Begin VB.CommandButton cmdopenfile2 
      BackColor       =   &H000080FF&
      Caption         =   "File&2"
      Height          =   135
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Open File2"
      Top             =   2640
      Width           =   135
   End
   Begin VB.CommandButton cmdopenfile1 
      BackColor       =   &H000080FF&
      Caption         =   "File&1"
      Height          =   135
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Open File1"
      Top             =   120
      Width           =   135
   End
   Begin VB.Frame Frame1 
      Caption         =   "Find Text"
      ForeColor       =   &H00FF0000&
      Height          =   735
      Left            =   2400
      TabIndex        =   10
      Top             =   5520
      Width           =   4815
      Begin VB.TextBox txtreplace 
         Height          =   285
         Left            =   3000
         TabIndex        =   12
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox txtfind 
         Height          =   285
         Left            =   600
         TabIndex        =   11
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Replace"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   2280
         TabIndex        =   16
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Find"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.Frame CtlParentSelectionFrame 
      BackColor       =   &H8000000A&
      Caption         =   "Select File"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   240
      TabIndex        =   7
      Top             =   5520
      Width           =   2055
      Begin VB.OptionButton Option1 
         BackColor       =   &H8000000A&
         Caption         =   "File1"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H8000000A&
         Caption         =   "File2"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   840
         TabIndex        =   8
         Top             =   360
         Width           =   975
      End
   End
   Begin MSComDlg.CommonDialog CDlg 
      Left            =   5520
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdcompare 
      BackColor       =   &H8000000B&
      Caption         =   "&Compare"
      Height          =   255
      Left            =   1440
      TabIndex        =   6
      Top             =   5160
      Width           =   1095
   End
   Begin VB.CommandButton cmdreplace 
      BackColor       =   &H8000000B&
      Caption         =   "&Replace"
      Height          =   255
      Left            =   3840
      TabIndex        =   5
      Top             =   5160
      Width           =   1095
   End
   Begin VB.CommandButton cmdignore 
      BackColor       =   &H8000000B&
      Caption         =   "&Ignore"
      Height          =   255
      Left            =   5040
      TabIndex        =   4
      Top             =   5160
      Width           =   1095
   End
   Begin VB.CommandButton cmdexit 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      Caption         =   "E&xit"
      Height          =   255
      Left            =   6240
      TabIndex        =   3
      Top             =   5160
      Width           =   1095
   End
   Begin VB.CommandButton cmdfind 
      BackColor       =   &H8000000B&
      Caption         =   "&Find"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   5160
      Width           =   1095
   End
   Begin RichTextLib.RichTextBox File1 
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "File1"
      Top             =   120
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   4260
      _Version        =   393217
      BackColor       =   16777215
      Enabled         =   -1  'True
      ScrollBars      =   3
      Appearance      =   0
      TextRTF         =   $"frmmain.frx":0000
   End
   Begin RichTextLib.RichTextBox File2 
      Height          =   2415
      Left            =   120
      TabIndex        =   1
      ToolTipText     =   "File2"
      Top             =   2640
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   4260
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      Appearance      =   0
      TextRTF         =   $"frmmain.frx":0082
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdcompare_Click()
Me.MousePointer = vbArrowHourglass
'File1.Enabled = False
File2.SelStart = 0
File2.SelLength = Len(File2.Text)
File2.SelBold = False
File2.SelColor = &H80000012

File1.SelStart = 0
File1.SelLength = Len(File1.Text)
File1.SelBold = False
File1.SelColor = &H80000012

    Dim StrFile1 As String, StrFile2 As String
    Dim IntCompStart As Integer
    Dim IntIndex As Integer
    If flag = True Then
    IntIndex = compstart
    Else
    IntIndex = 1
    End If
    Do While IntIndex <= Len(File1.Text) Or IntIndex <= Len(File2.Text)
        StrFile1 = Mid(File1.Text, IntIndex, 1)
        StrFile2 = Mid(File2.Text, IntIndex, 1)
        If StrComp(StrFile1, StrFile2, vbTextCompare) <> 0 Then
       ' MsgBox "Donnot Match", vbInformation, "Compare Utility"
        File1.SelStart = IntIndex
        File1.SelLength = 1
        File1.SelBold = True
        File1.SelColor = &H80FF&
        File2.SetFocus
        
        File2.SelStart = IntIndex
        File2.SelLength = 1
        File2.SelBold = True
        File2.SelColor = &H80FF&
        compstart = IntIndex
        MsgBox "MisMatch"
        Me.MousePointer = vbArrow
        Exit Sub
        End If
        
        IntIndex = IntIndex + 1
    Loop
    Me.MousePointer = vbArrow
    MsgBox "Successfull match", vbExclamation, "Compare Utility"
End Sub

Private Sub cmdexit_Click()
Dim IntBye As Integer
IntBye = MsgBox("Do you really want to quit the application", vbYesNo)
If IntBye = vbYes Then
End
Else
Exit Sub
End If
End Sub

Private Sub cmdfind_Click()
    File1.Enabled = True
    If Option1.Value = True Then
    Start = finder(File1, Trim(txtfind.Text), Start)
    Else
    Start = finder(File2, Trim(txtfind.Text), Start)
    End If
End Sub

Private Sub cmdignore_Click()
Start = Start + 1
End Sub

Private Sub cmdnc_Click()
compstart = 1
End Sub

Private Sub cmdopenfile1_Click()
    Start = 0
    Dim Filename As String
    CDlg.ShowOpen
    Filename = CDlg.Filename
    File1.LoadFile (Filename)
End Sub

Private Sub cmdopenfile2_Click()
    Dim Filename As String
    CDlg.ShowOpen
    Filename = CDlg.Filename
    File2.LoadFile (Filename)
End Sub

Private Sub cmdreplace_Click()
    Dim sub1 As String, sub2 As String
    sub1 = ""
    sub2 = ""
      If Start = 0 Then
        If Option1.Value = True Then
        Start = finder(File1, Trim(txtfind.Text), Start)
        Else
        Start = finder(File2, Trim(txtfind.Text), Start)
        End If
        Exit Sub
      End If
      
 If Start <> 0 Then
        If Option1.Value = True Then
              sub1 = Left(File1.Text, Start - 1)
              sub2 = Replace(File1.Text, txtfind.Text, txtreplace.Text, Start, Len(txtfind.Text))
              File1.Text = sub1 & sub2
        Else
              sub1 = Left(File2.Text, Start - 1)
              sub2 = Replace(File2.Text, txtfind.Text, txtreplace.Text, Start, Len(txtfind.Text))
              File2.Text = sub1 & sub2
        End If
        
        File1.Refresh
        If Option1.Value = True Then
        Start = finder(File1, Trim(txtfind.Text), Start)
        Else
        Start = finder(File2, Trim(txtfind.Text), Start)
        End If
     
 End If
           
End Sub


Private Sub Form_Unload(Cancel As Integer)
Dim IntBye As Integer
IntBye = MsgBox("Do you really want to quit the application", vbYesNo)
If IntBye = vbYes Then
Cancel = False
Else
Cancel = True
End If
End Sub
