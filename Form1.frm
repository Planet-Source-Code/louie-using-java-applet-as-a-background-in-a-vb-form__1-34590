VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form Form1 
   BackColor       =   &H8000000E&
   Caption         =   "Using Java applet as a background in a VB form"
   ClientHeight    =   6600
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8880
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   8880
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame17 
      Appearance      =   0  'Flat
      Caption         =   "Select a background to change it."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   1560
      TabIndex        =   2
      Top             =   240
      Visible         =   0   'False
      Width           =   5535
      Begin VB.OptionButton opt2 
         Caption         =   "Iceberg background"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   360
         TabIndex        =   5
         Top             =   360
         Width           =   2295
      End
      Begin VB.OptionButton opt1 
         Caption         =   "Cliff background"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2880
         TabIndex        =   4
         Top             =   360
         Width           =   2295
      End
      Begin VB.CommandButton Command1 
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5160
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   120
         Width           =   375
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   315
         Left            =   1920
         TabIndex        =   7
         Top             =   840
         Width           =   2655
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   315
         Left            =   360
         TabIndex        =   6
         Top             =   840
         Width           =   1455
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   3120
      Top             =   8880
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   7000
      Left            =   -120
      TabIndex        =   1
      Top             =   0
      Width           =   9000
      ExtentX         =   15875
      ExtentY         =   12347
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.PictureBox Picture1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7000
      Left            =   0
      ScaleHeight     =   3590.473
      ScaleMode       =   0  'User
      ScaleWidth      =   8940
      TabIndex        =   0
      Top             =   0
      Width           =   9000
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Using Java applet as a background in a VB form              '
'                                                            '
'                                                            '
'----------------------------------------------------------- '
'Ever wondered how a java applet would look as a _           '
'background in a VB form?, well 'check out this example.     '
'This example uses Anfy lake class to demonstrate this effect'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
Public g As String


Private Sub Command1_Click()
Form1.WindowState = 0
Dim X As Long
Dim inc As Long
inc = 50


For X = Me.Height To 300 Step -inc


    DoEvents
        Me.Move Me.Left, Me.Top + (inc \ 2), Me.Width, X
    Next X


    'This is the width part of the same sequence above


    For X = Me.Width To 2000 Step -inc


        DoEvents
            Me.Move Me.Left + (inc \ 2), Me.Top, X, Me.Height
        Next X
Me.Hide
Unload Me
End Sub

Private Sub Command13_Click()
If opt1.Value = True Then
WebBrowser1.Visible = True
Picture1.Visible = True
WebBrowser1.Height = Image.Height
WebBrowser1.Width = Image.Width
g = App.Path & "\Cliff.htm"
WebBrowser1.Navigate g
End If
'
If opt2.Value = True Then
WebBrowser1.Visible = True
Picture1.Visible = True
WebBrowser1.Height = Image.Height
WebBrowser1.Width = Image.Width
g = App.Path & "\Ice.htm"
WebBrowser1.Navigate g
End If
End Sub
Private Sub Form_Load()
Frame17.Visible = True
opt1.Value = True
WebBrowser1.Height = Image.Height
WebBrowser1.Width = Image.Width
g = App.Path & "\Cliff.htm"
WebBrowser1.Navigate g
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Form1.WindowState = 0
Dim X As Long
Dim inc As Long
inc = 50


For X = Me.Height To 300 Step -inc


    DoEvents
        Me.Move Me.Left, Me.Top + (inc \ 2), Me.Width, X
    Next X


    'This is the width part of the same sequence above


    For X = Me.Width To 2000 Step -inc


        DoEvents
            Me.Move Me.Left + (inc \ 2), Me.Top, X, Me.Height
        Next X
Me.Hide
Unload Me
End Sub

Private Sub opt1_Click()
WebBrowser1.Visible = True
Picture1.Visible = True
WebBrowser1.Height = Image.Height
WebBrowser1.Width = Image.Width
g = App.Path & "\Cliff.htm"
WebBrowser1.Navigate g
End Sub

Private Sub opt2_Click()
WebBrowser1.Visible = True
Picture1.Visible = True
WebBrowser1.Height = Image.Height
WebBrowser1.Width = Image.Width
g = App.Path & "\Ice.htm"
WebBrowser1.Navigate g
End Sub

Private Sub Timer1_Timer()
Label1.Caption = Format((Now), "h:mm:ss AM/PM")
Label2.Caption = Format((Now), "dddd  mmmm dd, yyyy")
End Sub
