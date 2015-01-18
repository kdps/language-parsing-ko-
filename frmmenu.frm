VERSION 5.00
Begin VB.Form frmmenu 
   Appearance      =   0  '평면
   BackColor       =   &H80000005&
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows 기본값
   Begin VB.Menu mnuControl 
      Caption         =   "메뉴(&M)"
      Begin VB.Menu mnuIn 
         Caption         =   "최소화(&I)"
      End
      Begin VB.Menu mnuLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "닫기(&C)"
      End
   End
End
Attribute VB_Name = "frmmenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnuClose_Click()
End
End Sub

Private Sub mnuIn_Click()
frmmain.WindowState = 1
End Sub
