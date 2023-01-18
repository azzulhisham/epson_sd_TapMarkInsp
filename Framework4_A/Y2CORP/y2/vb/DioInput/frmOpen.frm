VERSION 5.00
Begin VB.Form frmOpen 
   BorderStyle     =   3  '固定ﾀﾞｲｱﾛｸﾞ
   Caption         =   "オープン"
   ClientHeight    =   1245
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4650
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1245
   ScaleWidth      =   4650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'ｵｰﾅｰ ﾌｫｰﾑの中央
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3240
      TabIndex        =   3
      Top             =   600
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3240
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox txtModelName 
      Height          =   300
      Left            =   1080
      TabIndex        =   1
      Top             =   720
      Width           =   1935
   End
   Begin VB.TextBox txtBoardSw 
      Height          =   300
      Left            =   1080
      TabIndex        =   0
      Text            =   "0"
      Top             =   240
      Width           =   1935
   End
   Begin VB.Label lblModelName 
      AutoSize        =   -1  'True
      Caption         =   "型名"
      Height          =   180
      Left            =   240
      TabIndex        =   5
      Top             =   780
      Width           =   360
   End
   Begin VB.Label lblBoardSw 
      AutoSize        =   -1  'True
      Caption         =   "ボードSW"
      Height          =   180
      Left            =   240
      TabIndex        =   4
      Top             =   300
      Width           =   705
   End
End
Attribute VB_Name = "frmOpen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim intBoardSw As Integer
    Dim strModelName As String
    Dim lngResult As Long
    
    intBoardSw = Val(txtBoardSw.Text)
    strModelName = txtModelName.Text & Chr(0)
    
    lngResult = YdciOpen(intBoardSw, strModelName, g_intID)
    If lngResult = YDCI_RESULT_SUCCESS Then
        MsgBox "オープン成功", vbInformation
        g_intOpenFlag = DEVICE_OPENED
        Unload Me
    Else
        MsgBox "オープン失敗", vbExclamation
    End If
End Sub
