VERSION 5.00
Begin VB.Form Menu_f 
   BackColor       =   &H00404000&
   Caption         =   "„—ò“ ﬁ—¬‰ Ê ÕœÌÀ ò—Ì„Â «Â· »Ì  ⁄·ÌÂ« «·”·«„"
   ClientHeight    =   3225
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6435
   Icon            =   "Menu_F.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   6435
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00004080&
      Caption         =   "‰„«Ì‘ Ê÷⁄Ì  ò· ò«·«Â«Ì „ÊÃÊœ"
      BeginProperty Font 
         Name            =   "B Titr"
         Size            =   15.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   6135
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      X1              =   240
      X2              =   6240
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "œ—»«—Â »—‰«„Â"
      BeginProperty Font 
         Name            =   "B Titr"
         Size            =   15.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   615
      Left            =   3240
      TabIndex        =   3
      Top             =   2520
      Width           =   3015
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00000080&
      Caption         =   "Œ—ÊÃ"
      BeginProperty Font 
         Name            =   "B Titr"
         Size            =   15.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   2520
      Width           =   3015
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00004080&
      Caption         =   "›—Ê‘ ò«·«"
      BeginProperty Font 
         Name            =   "B Titr"
         Size            =   15.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   6135
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00004080&
      Caption         =   "(«÷«›Â ò—œ‰ ò«·« ( ›«ò Ê— ÃœÌœ"
      BeginProperty Font 
         Name            =   "B Titr"
         Size            =   15.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   6135
   End
End
Attribute VB_Name = "Menu_f"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Label1_Click()
show_book.Show

End Sub

Private Sub Label2_Click()
Forosh.Show

End Sub

Private Sub Label3_Click()
End
End Sub

Private Sub Label4_Click()
WE.Show

End Sub

Private Sub Label7_Click()
Add_kl.Show

End Sub
