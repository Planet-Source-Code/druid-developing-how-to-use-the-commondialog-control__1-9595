VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Commondialog Example"
   ClientHeight    =   7065
   ClientLeft      =   3435
   ClientTop       =   1905
   ClientWidth     =   8040
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7065
   ScaleWidth      =   8040
   Begin VB.CommandButton Command7 
      Caption         =   "Exit"
      Height          =   975
      Left            =   120
      Picture         =   "Form1.frx":0000
      Style           =   1  'Grafisch
      TabIndex        =   9
      Top             =   6000
      Width           =   7815
   End
   Begin VB.Frame Frame4 
      Caption         =   "Font :"
      Height          =   615
      Left            =   120
      TabIndex        =   7
      Top             =   3000
      Width           =   3855
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   120
         TabIndex        =   8
         Text            =   "Text Text Text"
         Top             =   240
         Width           =   3615
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Color :"
      Height          =   615
      Left            =   4080
      TabIndex        =   5
      Top             =   3000
      Width           =   3855
      Begin VB.PictureBox Picture1 
         Height          =   255
         Left            =   120
         ScaleHeight     =   195
         ScaleWidth      =   3555
         TabIndex        =   6
         Top             =   240
         Width           =   3615
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Picture :"
      Height          =   2775
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   7815
      Begin VB.PictureBox Picture2 
         Height          =   2415
         Left            =   120
         ScaleHeight     =   2355
         ScaleWidth      =   7515
         TabIndex        =   10
         Top             =   240
         Width           =   7575
      End
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Open"
      Height          =   975
      Left            =   120
      Picture         =   "Form1.frx":0442
      Style           =   1  'Grafisch
      TabIndex        =   3
      Top             =   3840
      Width           =   3855
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Color"
      Height          =   975
      Left            =   4080
      Picture         =   "Form1.frx":0884
      Style           =   1  'Grafisch
      TabIndex        =   2
      Top             =   4920
      Width           =   3855
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Printer"
      Height          =   975
      Left            =   120
      Picture         =   "Form1.frx":0CC6
      Style           =   1  'Grafisch
      TabIndex        =   1
      Top             =   4920
      Width           =   3855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Font"
      Height          =   975
      Left            =   4080
      Picture         =   "Form1.frx":1108
      Style           =   1  'Grafisch
      TabIndex        =   0
      Top             =   3840
      Width           =   3855
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   3480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command2_Click()
    'Show the FontSelect Common Dialog
    CommonDialog1.ShowFont
    'Set the font of the TextBox to the Selected Font
    Text2.Font = CommonDialog1.FontName
    'Set the FontSize of the TextBox to the selected FontSize
    Text2.FontSize = CommonDialog1.FontSize
End Sub

Private Sub Command4_Click()
    'Open the PrinterSelect Common Dialog
    CommonDialog1.ShowPrinter
End Sub

Private Sub Command5_Click()
    'Open the ColorSelect Common Dialog
    CommonDialog1.ShowColor
    'Set the BackColor of the Picture to the selected Color
    Picture1.BackColor = CommonDialog1.Color
End Sub

Private Sub Command6_Click()
    'Show the FileSelect Common Dialog
    CommonDialog1.ShowOpen
    'Load the selected file into the PictureBox
    Picture2.Picture = LoadPicture(CommonDialog1.filename)
End Sub

Private Sub Command7_Click()
    'Exit the program
    End
End Sub

Private Sub Form_Load()
    'Only show Bitmap files in the FileSelect Common Dialog
    CommonDialog1.Filter = "Bitmaps (*.bmp)|*.bmp"
End Sub

