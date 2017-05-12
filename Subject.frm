VERSION 5.00
Begin VB.Form SubjectForm 
   Caption         =   "Subject"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
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
      Left            =   2500
      TabIndex        =   3
      Top             =   2000
      Width           =   1500
   End
   Begin VB.CommandButton Next 
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
      Left            =   4000
      TabIndex        =   2
      Top             =   2000
      Width           =   1500
   End
   Begin VB.CommandButton AddSubject 
      Caption         =   "Add Subject"
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
      TabIndex        =   1
      Top             =   2000
      Width           =   1500
   End
   Begin VB.ComboBox SubjectsCombo 
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
      Left            =   1500
      TabIndex        =   0
      Text            =   "Subjects"
      Top             =   1000
      Width           =   3500
   End
End
Attribute VB_Name = "SubjectForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub AddSubject_Click()

    Unload SubjectForm
    
    Load AddSubjects

    AddSubjects.Show

End Sub

Private Sub DeleteCmd_Click()

Subjects.Open "SELECT * FROM Subjects", MainModule.con
    
    For rc = Subjects.RecordCount To 1 Step -1
            
        If Not SubjectsCombo.Text = Subjects!subjectcode & " " & Subjects!subjectname Then
            
            Subjects.MoveNext
            
        Else
            
            Subjects.Close
            
            Exit For
            
        End If
            
    Next
    
    If rc = 0 Then
    
        Subjects.Close
    
        MsgBox "Select or enter proper subject.", , "Error"
    
        Exit Sub
    
    End If
    
    MainModule.Subjects.Open "DROP TABLE " & MainModule.Subject, MainModule.con

    MainModule.Subjects.Open "DROP TABLE " & MainModule.Subject & "BitPattern", MainModule.con

    MainModule.Subjects.Open "DELETE * FROM Subjects WHERE SubjectCode=" & MainModule.Subject, MainModule.con

    MsgBox "Subject deleted."

End Sub

Private Sub Form_Load()

    SubjectForm.Width = 6500
    
    SubjectForm.Height = 4000
    
    AddSubjectstoSubjectsCombo

End Sub

Private Sub Next_Click()
    
    Subjects.Open "SELECT * FROM Subjects", MainModule.con
    
    For rc = Subjects.RecordCount To 1 Step -1
            
        If Not SubjectsCombo.Text = Subjects!subjectcode & " " & Subjects!subjectname Then
            
            Subjects.MoveNext
            
        Else
            
            Subjects.Close
            
            Exit For
            
        End If
            
    Next
    
    If rc = 0 Then
    
        Subjects.Close
    
        MsgBox "Select or enter proper subject.", , "Error"
    
        Exit Sub
    
    End If
       
    MainModule.Subject = Mid(SubjectsCombo.Text, 1, 5)
    
    Unload SubjectForm
    
    BitPatternRS.Open "SELECT * FROM " & Subject & "BitPattern", con
    
    If BitPatternRS.RecordCount = 0 Then
    
        BitPatternRS.Close
        
        BitPatternForm.Show
    
    Else
    
        BitPatternRS.Close
    
        ChoiceForm.Show
    
    End If
    
End Sub

Public Sub AddSubjectstoSubjectsCombo()
    
    rc = MainModule.SubjectsRecordCount
    
    If rc = 0 Then
       
       MsgBox "There are no subjects"
       
    Else
        
        Subjects.Open "SELECT * FROM Subjects", MainModule.con
        
        For i = 1 To rc Step 1
            
            SubjectsCombo.AddItem Subjects!subjectcode & " " & Subjects!subjectname
            
            Subjects.MoveNext
            
        Next
        
        Subjects.Close
         
    End If
    
End Sub

Private Sub SubjectsCombo_GotFocus()
    
    rc = MainModule.SubjectsRecordCount
    
    If rc = 0 Then
       
       MsgBox "There are no subjects"
       
    Else
        
        SubjectsCombo.Clear
        
        Subjects.Open "SELECT * FROM Subjects", MainModule.con
        
        For i = 1 To rc Step 1
            
            SubjectsCombo.AddItem Subjects!subjectcode & " " & Subjects!subjectname
            
            Subjects.MoveNext
            
        Next
        
        Subjects.Close
         
    End If
End Sub

Private Sub SubjectsCombo_LostFocus()

    MainModule.Subject = Mid(SubjectsCombo.Text, 1, 5)
    
End Sub
