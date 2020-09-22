VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Help / About"
   ClientHeight    =   3780
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4275
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3780
   ScaleWidth      =   4275
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Ok, Got It!"
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
      Left            =   120
      TabIndex        =   6
      Top             =   3480
      Width           =   4095
   End
   Begin VB.Label Label6 
      Caption         =   $"Image Coder03.frx":0000
      Height          =   735
      Left            =   120
      TabIndex        =   5
      Top             =   2760
      Width           =   4335
   End
   Begin VB.Label Label5 
      Caption         =   "Miniatures/link - This makes small images of choosen height/width...to see the full image just press on the little miniaturelink."
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   2040
      Width           =   3975
   End
   Begin VB.Label Label4 
      Caption         =   "Full Path - This prints the whole path to the file so you dont need to have the html file in the same directory as the images."
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   3855
   End
   Begin VB.Label Label3 
      Caption         =   "Filenames - This feature writes the filename(and path if Full Path is choosen) after the image"
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   3975
   End
   Begin VB.Label Label2 
      Caption         =   "Imgs on lines - This means how many images there shall be on each line in the html code"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   3855
   End
   Begin VB.Label Label1 
      Caption         =   "Explanations:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me

End Sub
