VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bitmap to GIF Conversion"
   ClientHeight    =   5280
   ClientLeft      =   2355
   ClientTop       =   1890
   ClientWidth     =   7125
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   2  'Custom
   ScaleHeight     =   5280
   ScaleWidth      =   7125
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Options"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1335
      Left            =   120
      TabIndex        =   4
      Top             =   3840
      Width           =   6855
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "&Convet Now.."
         Height          =   495
         Left            =   5280
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   720
         Width           =   1515
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00000000&
         Caption         =   "For &Web [Normal Quality]"
         ForeColor       =   &H00C0E0FF&
         Height          =   255
         Left            =   2280
         TabIndex        =   7
         Top             =   360
         Width           =   2175
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00000000&
         Caption         =   "&Pallate [Best Quality]"
         ForeColor       =   &H00C0E0FF&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   1935
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Save &transparent"
         ForeColor       =   &H00C0C0FF&
         Height          =   255
         Left            =   4800
         TabIndex        =   5
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "If u like this then plz mail me at gaurangvyas@hotmail.com"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   840
         Width           =   5055
      End
   End
   Begin VB.PictureBox Picture3 
      AutoSize        =   -1  'True
      Height          =   3315
      Left            =   3645
      ScaleHeight     =   3255
      ScaleWidth      =   3300
      TabIndex        =   3
      Top             =   60
      Width           =   3360
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Left            =   180
      ScaleHeight     =   315
      ScaleWidth      =   6735
      TabIndex        =   0
      Top             =   3420
      Width           =   6795
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
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
         Left            =   180
         TabIndex        =   1
         Top             =   60
         Width           =   675
      End
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      FillStyle       =   0  'Solid
      Height          =   3315
      Left            =   180
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   3255
      ScaleWidth      =   3300
      TabIndex        =   2
      Top             =   60
      Width           =   3360
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetTickCount Lib "kernel32" () As Long
Dim WithEvents cGif As GIF
Attribute cGif.VB_VarHelpID = -1

Private Sub cGif_Progress(ByVal Percents As Integer)
   Dim lEnd As Long
   lEnd = Picture2.Width * Percents / 100
   Picture2.Line (0, 0)-(lEnd, Picture1.Height), vbBlue, BF
   Picture2.CurrentX = lPos
   If lEnd >= Label1.Left Then Label1.ForeColor = vbWhite
   Label1 = Percents & "%"
End Sub

Private Sub Command1_Click()
Dim start As Long
Dim ending As Long
    
   start = GetTickCount
   Set cGif = New GIF
   Picture2.Cls
   Label1.ForeColor = vbBlack
   Picture2.Visible = True
   Form1.MousePointer = 11
   Command1.Enabled = False
   Picture1.Picture = Picture1.Image
   Picture1.Refresh
   cGif.SaveGIF Picture1.Picture, App.Path & "\test.gif", Picture1.hdc, CBool(Check1.Value), Picture1.Point(0, 0)
   Form1.MousePointer = 0
   Caption = "Convert File To Gif " & " (output file is test.gif and size is " & CInt(FileLen(App.Path & "\test.gif") / 1000) & "K)"
   Command1.Enabled = True
   Picture2.Visible = False
   Picture3.Picture = LoadPicture(App.Path & "\test.gif")
   Set cGif = Nothing
   ending = GetTickCount
   Debug.Print "total time taken = " & ((ending - start) / 1000)
End Sub

Private Sub Form_Load()

   Dim s As String, sFile As String
   sFile = "logo.bmp"
   With Picture1
       .AutoRedraw = True
       .FontBold = True
       .Picture = LoadPicture(sFile)
   End With
   Option1.Value = True
End Sub

Private Sub Frame1_DragDrop(Source As Control, X As Single, Y As Single)

End Sub
