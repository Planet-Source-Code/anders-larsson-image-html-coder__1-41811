VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Image Coder By Rynch"
   ClientHeight    =   3660
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7005
   Icon            =   "Image Coder.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3660
   ScaleWidth      =   7005
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   22
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox Text7 
      Height          =   195
      Left            =   0
      TabIndex        =   9
      Text            =   """"
      Top             =   3600
      Width           =   150
      Visible         =   0   'False
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   0
      TabIndex        =   3
      Top             =   960
      Width           =   1695
   End
   Begin VB.DirListBox Dir1 
      Height          =   1215
      Left            =   0
      TabIndex        =   2
      Top             =   1320
      Width           =   1695
   End
   Begin VB.FileListBox File1 
      Height          =   1065
      Left            =   0
      TabIndex        =   1
      Top             =   2520
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   3375
      Left            =   3600
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   240
      Width           =   2835
   End
   Begin VB.Frame Frame1 
      Height          =   3285
      Left            =   1800
      TabIndex        =   5
      Top             =   360
      Width           =   1695
      Begin VB.CommandButton Command3 
         Caption         =   "Clear"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   2280
         Width           =   1455
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1155
         TabIndex        =   19
         Text            =   "3"
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Quit"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   2880
         Width           =   1455
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Help/About"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   2640
         Width           =   1455
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Full Path"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         ToolTipText     =   "While Write the full path in the code instead of just the filename (is used if the html file is not in the same dir as the images)"
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox txthei 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1155
         TabIndex        =   12
         Text            =   "50"
         Top             =   1440
         Width           =   375
         Visible         =   0   'False
      End
      Begin VB.TextBox txtwid 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1155
         TabIndex        =   11
         Text            =   "50"
         Top             =   1200
         Width           =   375
         Visible         =   0   'False
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Miniatures/link"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         ToolTipText     =   "Makes the images to small miniatures that is links..to show the whole picture you just have to click on the miniature"
         Top             =   960
         Width           =   1455
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Filenames"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         ToolTipText     =   "Will Show the Filename after the file"
         Top             =   480
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Save HTML"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   2040
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Write Code"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "imgs on lines"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "img height"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1440
         Width           =   1335
         Visible         =   0   'False
      End
      Begin VB.Label Label3 
         Caption         =   "img width"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1200
         Width           =   735
         Visible         =   0   'False
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Files in Folder"
      Height          =   255
      Left            =   1800
      TabIndex        =   21
      Top             =   120
      Width           =   975
   End
   Begin VB.Image Image2 
      Height          =   3360
      Left            =   6480
      Picture         =   "Image Coder.frx":0442
      Top             =   240
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   960
      Left            =   720
      Picture         =   "Image Coder.frx":2225
      Top             =   0
      Width           =   960
   End
   Begin VB.Label Label1 
      Caption         =   "Code"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3600
      TabIndex        =   4
      Top             =   0
      Width           =   1935
   End
   Begin VB.Label Label5 
      Caption         =   "Image Coder 1.0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   17
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strmin As String
Dim strlink As String
Dim strlinklast As String
Dim strfilie As String
Dim intcolumns As Integer
Dim AntalNow As Integer
Dim strfilename As String
Dim Antal As Integer
Dim strimglast As String
Dim strimg As String




Private Sub Check2_Click()
If Check2.Value = 1 Then
Label3.Visible = True
Label4.Visible = True
txtwid.Visible = True
txthei.Visible = True
Else
txtwid.Visible = False
txthei.Visible = False
Label3.Visible = False
Label4.Visible = False
End If
End Sub

Private Sub Command1_Click()
Dim intnewline As Integer
Dim strnewline As String
Dim intuntilnewline As Integer
Text1.Text = "<!-- Rynch Image Coder For Lazzy Men -->"
Text1.Text = Text1.Text & vbCrLf & "<html>" & vbCrLf
intuntilnewline = Text3.Text - 1














strimg = "<img src=" & Text7.Text
strimglast = Text7.Text & ">"
For X = 1 To Antal


AntalNow = AntalNow + 1

If intnewline = intuntilnewline Then
strnewline = "<br>"
intnewline = 0
Else
strnewline = ""
intnewline = intnewline + 1
End If

If Check3.Value = 1 Then
    If File1.Path = "C:\" Or File1.Path = "D:\" Or File1.Path = "E:\" Or File1.Path = "F:\" Or File1.Path = "G:\" Or File1.Path = "H:\" Then
        strfilename = File1.Path & File1.List(AntalNow - 1)
        Else
        strfilename = File1.Path & "\" & File1.List(AntalNow - 1)
    End If
Else
strfilename = File1.List(AntalNow - 1)
End If

If Check1.Value = 1 Then
strfilie = strfilename
Else
strfilie = ""
End If
If Check2.Value = 1 Then
strlink = "<a href=" & Text7.Text & strfilename & Text7.Text & ">"
strlinklast = "</a>"
strmin = Text7.Text & " width=" & Text7.Text & txtwid.Text & Text7.Text & " height=" & Text7.Text & txthei & Text7.Text

Else
strlink = ""
strlinklast = ""
strmin = ""
End If







Text1.Text = Text1.Text & strlink & strimg & strfilename & strmin & strimglast & strlinklast & strfilie & strnewline & vbCrLf

Next
Text1.Text = Text1.Text & "</html>"
AntalNow = 0


End Sub

Private Sub Command3_Click()
Text1.Text = ""

End Sub

Private Sub Command2_Click()
Form2.Show

End Sub

Private Sub Command4_Click()
End
End Sub

Private Sub Command5_Click()
Form3.Show

End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
Antal = File1.ListCount
Text2.Text = Antal


End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
Antal = File1.ListCount
Text2.Text = Antal

End Sub

Private Sub Form_Load()
Antal = File1.ListCount
Text2.Text = Antal

End Sub

