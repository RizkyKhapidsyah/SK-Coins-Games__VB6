VERSION 5.00
Object = "{28D47522-CF84-11D1-834C-00A0249F0C28}#1.0#0"; "GIF89.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCoins 
   BackColor       =   &H00000000&
   Caption         =   "Coin Tossing"
   ClientHeight    =   8745
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   9675
   LinkTopic       =   "Form1"
   ScaleHeight     =   8745
   ScaleWidth      =   9675
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   4200
      Top             =   4080
   End
   Begin VB.CommandButton cmdEnd 
      Caption         =   "Exit Game"
      Height          =   495
      Left            =   6840
      TabIndex        =   10
      Top             =   8880
      Width           =   1215
   End
   Begin VB.CommandButton cmdFlip 
      Caption         =   "Flip Coin"
      Height          =   495
      Left            =   9000
      TabIndex        =   9
      Top             =   8880
      Width           =   1215
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2040
      Top             =   5640
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   36
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":09FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0FE8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1914
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   930
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   9675
      _ExtentX        =   17066
      _ExtentY        =   1640
      ButtonWidth     =   2143
      ButtonHeight    =   1482
      Appearance      =   1
      ImageList       =   "ImageList1"
      HotImageList    =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Hide Status Bar"
            Key             =   "Bar"
            Object.ToolTipText     =   "Hides the Status Bar"
            ImageIndex      =   1
            Style           =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Start Over"
            Key             =   "Start"
            Object.ToolTipText     =   "Starts the Game Over"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Flip Coin"
            Key             =   "Flip Coin"
            Object.ToolTipText     =   "Flips the Coin"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Exit"
            Key             =   "Exit"
            Object.ToolTipText     =   "Ends the Program"
            ImageIndex      =   4
         EndProperty
      EndProperty
      MousePointer    =   10
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   8370
      Width           =   9675
      _ExtentX        =   17066
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   0
            Object.Width           =   4101
            Text            =   "Enjoy Playing Coin Toss!"
            TextSave        =   "Enjoy Playing Coin Toss!"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            Bevel           =   0
            TextSave        =   "2/12/02"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            Bevel           =   0
            TextSave        =   "6:21 PM"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   2760
      Top             =   2400
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Start Over"
      Height          =   495
      Left            =   4680
      TabIndex        =   4
      Top             =   8880
      Width           =   1215
   End
   Begin VB.OptionButton optTails 
      BackColor       =   &H00000000&
      Caption         =   "Tails"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   735
      Left            =   5640
      TabIndex        =   1
      Top             =   6600
      Width           =   1575
   End
   Begin VB.OptionButton optHeads 
      BackColor       =   &H00000000&
      Caption         =   "Heads"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   615
      Left            =   5640
      TabIndex        =   0
      Top             =   5880
      Width           =   1815
   End
   Begin GIF89LibCtl.Gif89a Gif89a1 
      Height          =   3495
      Left            =   5760
      OleObjectBlob   =   "Form1.frx":2080
      TabIndex        =   6
      Top             =   1680
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.Image imgCoin 
      Height          =   3495
      Left            =   5760
      Picture         =   "Form1.frx":20C2
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   3735
   End
   Begin VB.Label lblCorrect 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "0 % Correct"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   495
      Left            =   6360
      TabIndex        =   5
      Top             =   7920
      Width           =   2775
   End
   Begin VB.Label lblTails 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   495
      Left            =   8280
      TabIndex        =   3
      Top             =   6720
      Width           =   1215
   End
   Begin VB.Label lblHeads 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   495
      Left            =   8280
      TabIndex        =   2
      Top             =   5880
      Width           =   1215
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmCoins"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CorrectAmt As Integer
Dim TailAmt As Integer
Dim HeadAmt As Integer
Dim GotoVal
Dim Gointo
Dim coinvalue


Private Sub cmdClear_Click()
    imgCoin.Visible = True
    Gif89a1.Visible = False
    Timer1.Enabled = False
    optHeads.Value = False
    optTails.Value = False
    HeadAmt = 0
    TailAmt = 0
    CorrectAmt = 0
    lblTails.Caption = "0"
    lblHeads.Caption = "0"
    lblCorrect.Caption = "0% Correct"
End Sub

Private Sub cmdEnd_Click()
    If frmCoins.WindowState = 2 Then
        End
    End If
GotoVal = Me.Height / 2


    For Gointo = 1 To GotoVal
    'NEW ADDITION NEXT LINE


    DoEvents
        Me.Height = Me.Height - 10
        'Me.Top = (Screen.Height - Me.Height) \ 2
        If Me.Height <= 11 Then GoTo horiz
    Next Gointo


    'This is the width part of the same sequence above
horiz:
    Me.Height = 30
    GotoVal = Me.Width / 2


    For Gointo = 1 To GotoVal
        'NEW ADDITION NEXT LINE


        DoEvents
            Me.Width = Me.Width - 10
            'Me.Left = (Screen.Width - Me.Width) \ 2
            If Me.Width <= 11 Then End
        Next Gointo
        
End
End Sub

Private Sub cmdFlip_Click()
If optHeads.Value = False And optTails.Value = False Then
    MsgBox ("Please choose heads or tails.")
Else:
    imgCoin.Visible = False
    Gif89a1.Visible = True
    StatusBar1.Panels(1).Text = "Please Wait...Coin is Flipping"
    Timer1.Enabled = True
    Gif89a1.FileName = ("D:\My Documents\My Pictures\coins\coin flipping.gif")
End If
End Sub

Private Sub Form_Load()
Randomize
    imgCoin.Visible = True
    Gif89a1.Visible = False
    Timer1.Enabled = False
    optHeads.Value = False
    optTails.Value = False
    TailAmt = 0
    HeadAmt = 0
    CorrectAmt = 0
    coinvalue = Rand(0, 1)
End Sub

Public Function Rand(ByVal Low As Long, _
                     ByVal High As Long) As Long
  Rand = Int((1 - 0 + 1) * Rnd + 0)

End Function

Private Sub mnuFileExit_Click()
If frmCoins.WindowState = 2 Then
    End
End If
GotoVal = Me.Height / 2


    For Gointo = 1 To GotoVal
    'NEW ADDITION NEXT LINE


    DoEvents
        Me.Height = Me.Height - 10
        'Me.Top = (Screen.Height - Me.Height) \ 2
        If Me.Height <= 11 Then GoTo horiz
    Next Gointo


    'This is the width part of the same sequence above
horiz:
    Me.Height = 30
    GotoVal = Me.Width / 2


    For Gointo = 1 To GotoVal
        'NEW ADDITION NEXT LINE


        DoEvents
            Me.Width = Me.Width - 10
            'Me.Left = (Screen.Width - Me.Width) \ 2
            If Me.Width <= 11 Then End
        Next Gointo
        
End
End Sub

Private Sub mnuHelpAbout_Click()
frmAbout.Show
End Sub

Private Sub optHeads_Click()
    StatusBar1.Panels(1).Text = "Heads has been chosen"
End Sub

Private Sub optTails_Click()
    StatusBar1.Panels(1).Text = "Tails has been chosen"
End Sub

Private Sub Timer1_Timer()
    
Randomize
coinvalue = Rand(0, 1)
    If coinvalue = 0 And (optTails.Value = True Or optHeads.Value = True) Then
       HeadAmt = HeadAmt + 1
       lblHeads.Caption = HeadAmt
       imgCoin.Picture = LoadPicture("D:\My Documents\My Pictures\coins\fishheads.jpg")
    ElseIf coinvalue = 1 And (optTails.Value = True Or optHeads.Value = True) Then
           TailAmt = TailAmt + 1
           imgCoin.Picture = LoadPicture("D:\My Documents\My Pictures\coins\fishtails.jpg")
           lblTails.Caption = TailAmt
    End If
    
    If optHeads.Value = True And coinvalue = 0 Then
        CorrectAmt = CorrectAmt + 1
    ElseIf optTails.Value = True And coinvalue = 1 Then
        CorrectAmt = CorrectAmt + 1
    End If
    
    Gif89a1.FileName = ("D:\My Documents\My Pictures\coins\coin flipping.gif")
    Gif89a1.Visible = False
    StatusBar1.Panels(1).Text = "Enjoy Playing Coin Toss!"
    imgCoin.Visible = True
    
    Timer1.Enabled = False
    Timer2.Enabled = True
End Sub

Private Sub Timer2_Timer()
    Dim percent As Integer
    percent = ((CorrectAmt) / (TailAmt + HeadAmt) * 100)
    lblCorrect.Caption = percent & "% Correct"
    Timer2.Enabled = False
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
    Case "Bar"
        If Toolbar1.Buttons(1).Value = 1 Then
        StatusBar1.Visible = False
        Toolbar1.Buttons(1).Caption = "Show Status Bar"
        Else:
        StatusBar1.Visible = True
        Toolbar1.Buttons(1).Caption = "Hide Status Bar"
        End If
    
    Case "Start"
        imgCoin.Visible = True
        Gif89a1.Visible = False
        Timer1.Enabled = False
        optHeads.Value = False
        optTails.Value = False
        HeadAmt = 0
        TailAmt = 0
        CorrectAmt = 0
        lblTails.Caption = "0"
        lblHeads.Caption = "0"
        lblCorrect.Caption = "0% Correct"
        
    Case "Flip Coin"
    If optHeads.Value = False And optTails.Value = False Then
        MsgBox ("Please choose heads or tails.")
    Else:
        imgCoin.Visible = False
        Gif89a1.Visible = True
        StatusBar1.Panels(1).Text = "Please Wait...Coin is Flipping"
        Timer1.Enabled = True
        Gif89a1.FileName = ("D:\My Documents\My Pictures\coins\coin flipping.gif")
    End If
    
    Case "Exit"
        If frmCoins.WindowState = 2 Then
           End
        End If

GotoVal = Me.Height / 2


    For Gointo = 1 To GotoVal
    'NEW ADDITION NEXT LINE


    DoEvents
        Me.Height = Me.Height - 10
        'Me.Top = (Screen.Height - Me.Height) \ 2
        If Me.Height <= 11 Then GoTo horiz
    Next Gointo


    'This is the width part of the same sequence above
horiz:
    Me.Height = 30
    GotoVal = Me.Width / 2


    For Gointo = 1 To GotoVal
        'NEW ADDITION NEXT LINE


        DoEvents
            Me.Width = Me.Width - 10
            'Me.Left = (Screen.Width - Me.Width) \ 2
            If Me.Width <= 11 Then End
        Next Gointo
        
End
End Select
End Sub
