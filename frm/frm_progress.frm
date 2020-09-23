VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_progress 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Please wait..."
   ClientHeight    =   1335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1335
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar ProgressBar 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label lbl_Action_caption 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "frm_progress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Action_Caption As String

Private bHideProgress As Boolean
Public Function Ini_Progress(ByVal s_Action As String, ByVal lMax As Long) As Boolean
    ProgressBar.Max = lMax
    ProgressBar.value = 0
    lbl_Action_caption.Caption = s_Action
    lbl_Action_caption.Refresh
    DoEvents
End Function

Public Function Update_Progresss(ByVal s_Action As String, ByVal lProgress As Long)
    On Error Resume Next
    If Not Me.HideProgress Then
        If ProgressBar.value <> lProgress Then
            ProgressBar.value = lProgress
        End If
    End If
    If s_Action <> lbl_Action_caption.Caption Then
        lbl_Action_caption.Caption = s_Action
    End If

End Function

Public Property Get HideProgress() As Boolean
    HideProgress = bHideProgress
End Property

Public Property Let HideProgress(ByVal bNewValue As Boolean)
    bHideProgress = bNewValue
    If bHideProgress Then
        ProgressBar.Visible = False
    Else
        ProgressBar.Visible = False
    End If
End Property
