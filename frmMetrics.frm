VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00E0E0E0&
   Caption         =   "The Unanswered Questions"
   ClientHeight    =   2160
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5220
   LinkTopic       =   "Form1"
   ScaleHeight     =   2160
   ScaleWidth      =   5220
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command6 
      Caption         =   "How Tall in pixels are my caption bitmaps?"
      Height          =   675
      Left            =   3060
      TabIndex        =   5
      Top             =   840
      Width           =   2055
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Pixels between Desktop Icons (Horiz)"
      Height          =   675
      Left            =   1560
      TabIndex        =   4
      Top             =   840
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Pixels between Desktop Icons (Vert)"
      Height          =   675
      Left            =   60
      TabIndex        =   3
      Top             =   840
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "What Boot Mode am I in?"
      Height          =   735
      Left            =   60
      TabIndex        =   2
      Top             =   60
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Is my computer too slow for Win95?"
      Height          =   735
      Left            =   3360
      TabIndex        =   1
      Top             =   60
      Width           =   1755
   End
   Begin VB.CommandButton Command1 
      Caption         =   "How Many Buttons does my mouse have?"
      Height          =   735
      Left            =   1560
      TabIndex        =   0
      Top             =   60
      Width           =   1755
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "There are MANY more things you are able to find out with this code. This is just a sample consisting of only 6."
      Height          =   495
      Left            =   60
      TabIndex        =   6
      Top             =   1620
      Width           =   5055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    MsgBox GetSystemMetrics(SM_CMOUSEBUTTONS)
End Sub

Private Sub Command2_Click()
    If GetSystemMetrics(SM_SLOWMACHINE) = 0 Then
        MsgBox "No."
    Else
        MsgBox "Yes."
    End If
End Sub

Private Sub Command3_Click()
    If GetSystemMetrics(SM_CLEANBOOT) = 0 Then
        MsgBox "Normal Mode."
    ElseIf GetSystemMetrics(SM_CLEANBOOT) = 1 Then
        MsgBox "Safe Mode"
    Else
        MsgBox "Safe Mode +Network ability"
    End If
End Sub

Private Sub Command4_Click()
    MsgBox GetSystemMetrics(SM_CYICONSPACING)
End Sub

Private Sub Command5_Click()
    MsgBox GetSystemMetrics(SM_CXICONSPACING)
End Sub

Private Sub Command6_Click()
    MsgBox GetSystemMetrics(SM_CYSIZE)
End Sub
