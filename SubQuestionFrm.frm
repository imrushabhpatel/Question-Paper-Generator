VERSION 5.00
Begin VB.Form SubQuestionFrm 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10860
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   10860
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton NextCmd 
      Caption         =   "Next"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   3000
      TabIndex        =   1
      Top             =   0
      Width           =   2000
   End
   Begin VB.CommandButton ResetCmd 
      Caption         =   "Reset"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   1000
      TabIndex        =   0
      Top             =   0
      Width           =   2000
   End
   Begin VB.Label ChapterLbl 
      Alignment       =   2  'Center
      Caption         =   "Chapter No."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   3000
      TabIndex        =   3
      Top             =   1000
      Width           =   2000
   End
   Begin VB.Label SubQuestionLbl 
      Alignment       =   2  'Center
      Caption         =   "Sub Question"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   1000
      TabIndex        =   2
      Top             =   1000
      Width           =   2000
   End
End
Attribute VB_Name = "SubQuestionFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SubQSLbl As Label
Dim temp As TextBox
Dim ChapterTxt As TextBox

Private Sub Form_Load()

    Me.Caption = MainModule.Subject

    MainModule.RNo = MainModule.RNo + 1

    Me.Width = 7000

    Set temp = BitPatternForm.Controls("SQOText" & MainModule.RNo)

    For i = 1 To temp.Text Step 1
    
        Set SubQSLbl = Controls.Add("Vb.Label", "Label" & i)
        
        Set ChapterTxt = Controls.Add("Vb.TextBox", "Text" & i)
        
        With SubQSLbl
        
            .Alignment = 2
            .Caption = i
            .FontName = "Times New Roman"
            .FontSize = 12
            .Height = 500
            .Left = 1000
            .Top = i * 1000 + 1000
            .Visible = True
            .Width = 2000
        
        End With
        
        With ChapterTxt
        
            .FontName = "Times New Roman"
            .FontSize = 12
            .Height = 500
            .Left = 3000
            .Top = i * 1000 + 1000
            .Visible = True
            .Width = 2000
        
        End With
    
    Next

    ResetCmd.Top = ChapterTxt.Top + 1000
    NextCmd.Top = ChapterTxt.Top + 1000
    
    Me.Height = ChapterTxt.Top + 3000

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    f1

End Sub

Private Sub NextCmd_Click()

    For i = 1 To temp.Text Step 1
    
        Set ChapterTxt = Me.Controls("Text" & i)
    
        If Not IsNumeric(ChapterTxt.Text) Then
        
            MsgBox "Enter numeric values only.", , "Error"
            
            ChapterTxt.Text = ""
            
            ChapterTxt.SetFocus
            
            Exit Sub
            
        Else
        
            MainModule.Subjects.Open "SELECT NoofChapters FROM Subjects WHERE SubjectCode=" & MainModule.Subject, MainModule.con
            
            If Val(ChapterTxt.Text) < 1 Or Val(ChapterTxt.Text) > Subjects!NoofChapters Then
                
                    MsgBox "Please, enter valid chapter no."
                    
                    ChapterTxt.Text = ""
            
                    ChapterTxt.SetFocus
                
                    Subjects.Close
                
                    Exit Sub
            
            End If
        
            Subjects.Close
        
        End If
    
    Next

    Dim m, o, opt As Integer

    m = BitPatternForm.Controls("MarksText" & MainModule.RNo).Text / BitPatternForm.Controls("SQWOText" & MainModule.RNo).Text

    o = Val(BitPatternForm.Controls("SQWOText" & MainModule.RNo).Text)

    Set TitleTxt = BitPatternForm.Controls("TitleText" & MainModule.RNo)

    For i = 1 To temp.Text Step 1
    
        If o >= i Then
        
            opt = 1
            
        Else
            
            opt = 0
        
        End If

        Set ChapterTxt = Me.Controls("Text" & i)

        BitPatternRS.Open "INSERT INTO " & MainModule.Subject & "BitPattern VALUES(" & MainModule.RNo & "," & i & "," & ChapterTxt.Text & "," & m & "," & opt & ",'" & TitleTxt.Text & "')", MainModule.con
    
    Next
    
    Unload Me
    
    If Not MainModule.RNo = BitPatternForm.R Then
    
        Me.Show
    
    End If

End Sub


