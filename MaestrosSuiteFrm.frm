VERSION 5.00
Begin VB.Form MaestrosSuiteFrm 
   BorderStyle     =   0  'None
   Caption         =   "Maestro's Suite"
   ClientHeight    =   3090
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "MaestrosSuiteFrm.frx":0000
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   8880
      Top             =   7440
   End
   Begin VB.PictureBox Picture1 
      Height          =   8200
      Left            =   0
      Picture         =   "MaestrosSuiteFrm.frx":C3EA
      ScaleHeight     =   8145
      ScaleWidth      =   7935
      TabIndex        =   0
      Top             =   0
      Width           =   8000
   End
End
Attribute VB_Name = "MaestrosSuiteFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cnt As Integer


Private Sub Form_Load()

    Me.Height = 8200

    Me.Width = 8000

End Sub

Private Sub Timer1_Timer()

    cnt = cnt + 1
    
    If cnt = 5 Then
        
        Unload Me
    
        Main
        
    End If

End Sub
