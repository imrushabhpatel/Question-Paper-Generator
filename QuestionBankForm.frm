VERSION 5.00
Begin VB.Form QuestionBankForm 
   Caption         =   "Question Bank"
   ClientHeight    =   9345
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   11715
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   15
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   9345
   ScaleWidth      =   11715
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame RepetitionFrame 
      Caption         =   "Repetition"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2000
      Left            =   4200
      TabIndex        =   14
      Top             =   6355
      Width           =   3200
      Begin VB.TextBox RepetitionText 
         Height          =   450
         Left            =   500
         TabIndex        =   15
         Top             =   775
         Width           =   2292
      End
   End
   Begin VB.CommandButton DeleteCmd 
      Caption         =   "Delete"
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
      Left            =   6692
      TabIndex        =   11
      Top             =   9355
      Width           =   2000
   End
   Begin VB.CommandButton LockCmd 
      Caption         =   "Lock"
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
      Left            =   14520
      Picture         =   "QuestionBankForm.frx":0000
      TabIndex        =   10
      Top             =   960
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.CommandButton EditCmd 
      Caption         =   "Edit"
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
      Left            =   6692
      TabIndex        =   9
      Top             =   9355
      Width           =   2000
   End
   Begin VB.CommandButton AddCmd 
      Caption         =   "Add"
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
      Left            =   2846
      TabIndex        =   8
      Top             =   9355
      Width           =   2000
   End
   Begin VB.CommandButton DoneCmd 
      Caption         =   "Exit"
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
      Left            =   10538
      TabIndex        =   7
      Top             =   9355
      Width           =   2000
   End
   Begin VB.Frame MarksFrame 
      Caption         =   "Marks"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2000
      Left            =   1000
      TabIndex        =   6
      Top             =   6355
      Width           =   3200
      Begin VB.ComboBox MarksCombo 
         Height          =   450
         Left            =   500
         TabIndex        =   13
         Top             =   775
         Width           =   2292
      End
   End
   Begin VB.Frame TitleFrame 
      Caption         =   "Title"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2000
      Left            =   7400
      TabIndex        =   5
      Top             =   6360
      Width           =   6985
      Begin VB.ComboBox TitleCombo 
         Height          =   450
         Left            =   500
         TabIndex        =   12
         Top             =   775
         Width           =   5985
      End
   End
   Begin VB.ComboBox QuestionCombo 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3000
      TabIndex        =   4
      Text            =   "Question"
      Top             =   1000
      Width           =   11385
   End
   Begin VB.ComboBox ChapterCombo 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1000
      Sorted          =   -1  'True
      TabIndex        =   3
      Text            =   "Chapter No."
      Top             =   1000
      Width           =   2000
   End
   Begin VB.CommandButton SubmitCmd 
      Caption         =   "Submit"
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
      Left            =   2846
      TabIndex        =   1
      Top             =   9355
      Width           =   2000
   End
   Begin VB.Frame QuestionFrame 
      Caption         =   "Question"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4500
      Left            =   1000
      TabIndex        =   0
      Top             =   2000
      Width           =   13385
      Begin VB.TextBox QuestionText 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3500
         IMEMode         =   3  'DISABLE
         Left            =   500
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   600
         Width           =   12385
      End
   End
End
Attribute VB_Name = "QuestionBankForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Question As String
Dim MarkOption As Integer

Private Sub AddCmd_Click()
    
    reset
    
    ChapterCombo.Enabled = True
    
    QuestionFrame.Enabled = True
    
    MarksFrame.Enabled = True
    
    RepetitionFrame.Enabled = True
    
    SubmitCmd.Caption = "Insert question"
    
    AddCmd.Visible = False

    EditCmd.Visible = False

    QuestionText.Text = ""
    
    RepetitionText.Text = ""
    
'    Mark2.Value = False

'    Mark4.Value = False

'    Mark6.Value = False

'    Mark8.Value = False

    SubmitCmd.Visible = True

    SubmitCmd.Enabled = True
    
    EditCmd.Enabled = False
    
    QuestionCombo.Enabled = False

    'DoneCmd.Caption = "Exit"

End Sub

Private Sub Combo1_Change()

End Sub

Private Sub Combo1_GotFocus()

End Sub

Private Sub DeleteCmd_Click()

    MainModule.SelSubject.Open "DELETE FROM " & MainModule.Subject & " WHERE Question='" & QuestionCombo.Text & "'", MainModule.con

    MsgBox "Question deleted.", , "Error"

    reset

End Sub

Private Sub DoneCmd_Click()

    QuestionCombo.Enabled = True

    DeleteCmd.Visible = False

    SubmitCmd.Visible = False
    
    EditCmd.Visible = True
    
    AddCmd.Visible = True
    
    If (AddCmd.Enabled = True) And (EditCmd.Enabled = True) Then
    
        Unload Me
        
        ChoiceForm.Show
    
        Exit Sub
    
    End If
        
    If AddCmd.Enabled = False Then
    
        LockCmd.Visible = False
    
        AddCmd.Enabled = True
    
    End If
    
    If EditCmd.Enabled = False Then
    
        EditCmd.Enabled = True
        
    End If
    
    'ChapterCombo.Enabled = False
    
    'QuestionCombo.Enabled = False
    
    'QuestionFrame.Enabled = False
    
    'MarksFrame.Enabled = False
    
    'RepetitionFrame.Enabled = False

End Sub

Private Sub EditCmd_Click()

    DeleteCmd.Visible = True

    ChapterCombo.Enabled = True

    QuestionCombo.Enabled = True
    
    QuestionFrame.Enabled = True
    
    MarksFrame.Enabled = True
    
    RepetitionFrame.Enabled = True

    SubmitCmd.Caption = "Update question"

    AddCmd.Visible = False

    EditCmd.Visible = False

    LockCmd.Visible = True

    SubmitCmd.Visible = True

    SubmitCmd.Enabled = True
    
    AddCmd.Enabled = False

End Sub

Sub reset()

    ChapterCombo.Text = "Chapter No."

    QuestionCombo = "Question"

    QuestionText.Text = ""
    
    MarksCombo.Text = ""
    
    RepetitionText.Text = ""
    
    TitleCombo.Text = ""
    
End Sub

Private Sub Form_Load()

    Me.Caption = MainModule.Subject & "Question Bank"

    DeleteCmd.Visible = False

    'ChapterCombo.Enabled = False
    
    'QuestionCombo.Enabled = False

    SubmitCmd.Visible = False

    'QuestionFrame.Enabled = False
    
    'RepetitionFrame.Enabled = False

    'MarksFrame.Enabled = False

    MainModule.Subjects.Open "SELECT NoofChapters FROM Subjects WHERE SubjectCode=" & MainModule.Subject, MainModule.con

    For i = 1 To MainModule.Subjects!NoofChapters Step 1
    
        ChapterCombo.AddItem i
    
    Next
 
    MainModule.Subjects.Close
    
End Sub

Public Sub AddQuestionstoQuestionCombo()

    rc = MainModule.SelSubjectRecordCount

    If rc = 0 Then
    
        QuestionCombo.Enabled = False
    
    Else
    
        SelSubject.Open "SELECT * FROM " & MainModule.Subject & " WHERE Chapter=" & ChapterCombo.Text, MainModule.con

            For i = 1 To SelSubject.RecordCount Step 1
        
                QuestionCombo.AddItem SelSubject!Question
            
                SelSubject.MoveNext
        
            Next
    
        SelSubject.Close
    
    End If

End Sub

Private Sub LockCmd_Click()

    QuestionCombo.Enabled = False
    
    LockCmd.Enabled = False

End Sub


Private Sub MarksCombo_GotFocus()

    MarksCombo.Clear
    
    MainModule.BitPatternRS.Open "SELECT DISTINCT(Marks) FROM " & MainModule.Subject & "BitPattern ORDER BY Marks", MainModule.con
    
    For i = 1 To BitPatternRS.RecordCount Step 1
    
        MarksCombo.AddItem BitPatternRS!Marks
        
        BitPatternRS.MoveNext
    
    Next
    
    MainModule.BitPatternRS.Close

End Sub

Private Sub QuestionCombo_GotFocus()
    
    QuestionCombo.Clear
    
    MainModule.Subjects.Open "SELECT NoofChapters FROM Subjects WHERE SubjectCode=" & MainModule.Subject, MainModule.con

    For i = 1 To MainModule.Subjects!NoofChapters Step 1
    
        If ChapterCombo.Text = i Then
        
            Exit For
        
        End If
    
    Next
    
    If (i - 1) = MainModule.Subjects!NoofChapters Then
        
        MsgBox "Please select proper chapter"
            
        MainModule.Subjects.Close
            
        Exit Sub
            
    End If

    MainModule.Subjects.Close
    
    AddQuestionstoQuestionCombo

End Sub

Private Sub MarksCombo_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If Not EditCmd.Enabled = False And Not AddCmd.Enabled = False Then
    
        KeyCode = 0
        
        Shift = 0
    
    End If

End Sub

Private Sub MarksCombo_KeyPress(KeyAscii As Integer)

    If Not EditCmd.Enabled = False And Not AddCmd.Enabled = False Then
    
        KeyAscii = 0
     
    End If

End Sub

Private Sub TitleCombo_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If Not EditCmd.Enabled = False And Not AddCmd.Enabled = False Then
    
        KeyCode = 0
        
        Shift = 0
    
    End If

End Sub

Private Sub TitleCombo_KeyPress(KeyAscii As Integer)

    If Not EditCmd.Enabled = False And Not AddCmd.Enabled = False Then
    
        KeyAscii = 0
     
    End If

End Sub

Private Sub RepetitionText_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If Not EditCmd.Enabled = False And Not AddCmd.Enabled = False Then
    
        KeyCode = 0
        
        Shift = 0
    
    End If

End Sub

Private Sub RepetitionText_KeyPress(KeyAscii As Integer)

    If Not EditCmd.Enabled = False And Not AddCmd.Enabled = False Then
    
        KeyAscii = 0
     
    End If

End Sub

Private Sub QuestionCombo_LostFocus()

    MainModule.SelSubject.Open "SELECT Question FROM " & MainModule.Subject, MainModule.con
    
    For i = 1 To SelSubject.RecordCount Step 1
    
        If QuestionCombo.Text = SelSubject!Question Then
        
            Exit For
        
        End If
        
        SelSubject.MoveNext
    
    Next
    
    If i > SelSubject.RecordCount Then
    
        MsgBox "Please, select proper question.", , "Error"
        
        MainModule.SelSubject.Close
        
        Exit Sub
    
    End If
    
    MainModule.SelSubject.Close

    MainModule.SelSubject.Open "SELECT * FROM " & MainModule.Subject & " WHERE Question='" & QuestionCombo.Text & "'", MainModule.con
        
    QuestionText.Text = MainModule.SelSubject!Question
        
    RepetitionText.Text = MainModule.SelSubject!Repetition

    MarksCombo.Text = MainModule.SelSubject!Marks
        
    TitleCombo.Text = MainModule.SelSubject!Title
        
    MainModule.SelSubject.Close

End Sub

Private Sub QuestionText_KeyDown(KeyCode As Integer, Shift As Integer)

    If Not EditCmd.Enabled = False And Not AddCmd.Enabled = False Then
    
        KeyCode = 0
        
        Shift = 0
    
    End If

End Sub

Private Sub QuestionText_KeyPress(KeyAscii As Integer)
    
    If Not EditCmd.Enabled = False And Not AddCmd.Enabled = False Then
    
        KeyAscii = 0
    
    End If

End Sub

Private Sub SubmitCmd_Click()

   MainModule.Subjects.Open "SELECT NoofChapters FROM Subjects WHERE SubjectCode=" & MainModule.Subject, MainModule.con

    For i = 1 To MainModule.Subjects!NoofChapters Step 1
    
        If ChapterCombo.Text = i Then
        
            Exit For
        
        End If
    
    Next
    
    If (i - 1) = MainModule.Subjects!NoofChapters Then
        
        MsgBox "Please select proper chapter"
            
        MainModule.Subjects.Close
            
        Exit Sub
            
    End If

    MainModule.Subjects.Close
    
    If QuestionText.Text = "" Or RepetitionText.Text = "" Or MarksCombo.Text = "" Then
    
        MsgBox "One or more fields empty"
    
        Exit Sub
    
    End If
    
    If EditCmd.Enabled = False Then
    
        MarkOption = MarksCombo.Text
        
        SelSubject.Open "INSERT INTO " & MainModule.Subject & " VALUES(" & ChapterCombo.Text & ",'" & QuestionText.Text & "'," & RepetitionText.Text & "," & MarkOption & ",0,'" & TitleCombo.Text & "')", MainModule.con
    
        MsgBox "Question inserted."
    
    End If
    
    If AddCmd.Enabled = False Then
    
        MarkOption = MarksCombo.Text
        
        SelSubject.Open "UPDATE " & MainModule.Subject & " SET Chapter=" & ChapterCombo.Text & ",Question='" & QuestionText.Text & "',Repetition=" & RepetitionText.Text & ",Marks=" & MarkOption & ",Flag=0 WHERE Question='" & QuestionCombo.Text & "'", MainModule.con
    
        MsgBox "Question edited."
        
        QuestionCombo.Enabled = True
        
        LockCmd.Enabled = True
        
    End If
    
    reset

End Sub

Private Sub TitleCombo_GotFocus()

    TitleCombo.Clear
    
    MainModule.BitPatternRS.Open "SELECT DISTINCT(Title) FROM " & MainModule.Subject & "BitPattern", MainModule.con

    For i = 1 To BitPatternRS.RecordCount Step 1
    
        TitleCombo.AddItem BitPatternRS!Title
        
        BitPatternRS.MoveNext
    
    Next

    MainModule.BitPatternRS.Close

End Sub
