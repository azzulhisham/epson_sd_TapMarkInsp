VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  '固定(実線)
   Caption         =   "DIO Input Sample for Visual Basic"
   ClientHeight    =   345
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   4830
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   345
   ScaleWidth      =   4830
   StartUpPosition =   3  'Windows の既定値
   Begin VB.Menu mnuFile 
      Caption         =   "ファイル"
      Begin VB.Menu mnuExit 
         Caption         =   "終了"
      End
   End
   Begin VB.Menu mnuBoard 
      Caption         =   "ボード"
      Begin VB.Menu mnuOpen 
         Caption         =   "オープン"
      End
   End
   Begin VB.Menu mnuInput 
      Caption         =   "入力"
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "ヘルプ"
      Begin VB.Menu mnuAbout 
         Caption         =   "バージョン情報"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Unload(Cancel As Integer)
    If g_intOpenFlag = DEVICE_OPENED Then
        If YdciClose(g_intID) = False Then
            MsgBox "クローズ失敗", vbExclamation
        End If
    End If
End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show vbModal, Me
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuInput_Click()
    frmInput.Show vbModal, Me
End Sub

Private Sub mnuOpen_Click()
    frmOpen.Show vbModal, Me
End Sub
