VERSION 5.00
Begin VB.Form frmHelpAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Help About"
   ClientHeight    =   3690
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4500
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   4500
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   915
      Left            =   120
      TabIndex        =   8
      Top             =   1740
      Width           =   1185
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "neenojee@hotmail.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   2025
      MouseIcon       =   "Form2.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   9
      Top             =   1980
      Width           =   1665
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "By: Naveed ur Rahman"
      Height          =   195
      Left            =   2025
      TabIndex        =   7
      Top             =   1755
      Width           =   1665
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "Date: 11 Feb 2003"
      Height          =   195
      Left            =   2025
      TabIndex        =   6
      Top             =   2340
      Width           =   1335
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "Dedicated to all VB Programmers"
      Height          =   195
      Left            =   2025
      TabIndex        =   5
      Top             =   2580
      Width           =   2325
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "- Naveed's Software"
      Height          =   195
      Left            =   2910
      TabIndex        =   4
      Top             =   2835
      Width           =   1440
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      Caption         =   "Copyrights(c) 2002-2003 Naveed's Software. All Rights Are Reserved."
      Height          =   480
      Left            =   630
      TabIndex        =   3
      Top             =   3180
      Width           =   3180
   End
   Begin VB.Line Line1 
      X1              =   90
      X2              =   4410
      Y1              =   1545
      Y2              =   1545
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Just Vote For Me"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   1208
      TabIndex        =   2
      Top             =   1110
      Width           =   2085
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "If you like my codes"
      Height          =   195
      Left            =   1553
      TabIndex        =   1
      Top             =   750
      Width           =   1395
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Naveed Animation Files (*.ani) Display"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4020
   End
End
Attribute VB_Name = "frmHelpAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Me.Hide
End Sub

Private Sub Label7_Click()
'Email Address: neenojee@hotmail.com
'Subject: Naveed Animation Files Display

'I know this is not a fair method :)
Shell "start mailto:neenojee@hotmail.com?subject=Naveed%20Animation%20Files%20Display", vbHide

End Sub
