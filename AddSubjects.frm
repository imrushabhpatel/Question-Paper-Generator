VERSION 5.00
Begin VB.Form AddSubjects 
   Caption         =   "Add Subjects"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton BackCmd 
      Caption         =   "Done"
      Height          =   500
      Left            =   6000
      TabIndex        =   8
      Top             =   4500
      Width           =   2000
   End
   Begin VB.CommandButton ResetCmd 
      Caption         =   "Reset"
      Height          =   500
      Left            =   3500
      TabIndex        =   7
      Top             =   4500
      Width           =   2000
   End
   Begin VB.CommandButton AddCmd 
      Caption         =   "Add  Subject"
      Height          =   500
      Left            =   1000
      TabIndex        =   6
      Top             =   4500
      Width           =   2000
   End
   Begin VB.TextBox NoofChaptersText 
      Height          =   500
      Left            =   4500
      TabIndex        =   2
      Top             =   3000
      Width           =   3000
   End
   Begin VB.TextBox SubjectNameText 
      Height          =   500
      Left            =   4500
      TabIndex        =   1
      Top             =   2000
      Width           =   3000
   End
   Begin VB.TextBox SubjectCodeText 
      Height          =   500
      Left            =   4500
      MaxLength       =   5
      TabIndex        =   0
      Top             =   1000
      Width           =   3000
   End
   Begin VB.Label NoofChaptersLabel 
      Caption         =   "Number of Chapters"
      Height          =   500
      Left            =   1500
      TabIndex        =   5
      Top             =   3000
      Width           =   2000
   End
   Begin VB.Label SubjectNameLabel 
      Caption         =   "Subject Name"
      Height          =   500
      Left            =   1500
      TabIndex        =   4
      Top             =   2000
      Width           =   2000
   End
   Begin VB.Label SubjectCodeLabel 
      Caption         =   "Subject Code"
      Height          =   500
      Left            =   1500
      TabIndex        =   3
      Top             =   1000
      Width           =   2000
   End
End
Attribute VB_Name = "AddSubjects"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim AddSubjectRS As New ADODB.Recordset
Dim rs1, rs2 As New ADODB.Recordset

Private Sub AddCmd_Click()

    If SubjectCodeText.Text = "" Or SubjectNameText.Text = "" Or NoofChaptersText.Text = "" Then
    
        MsgBox "One or more fields are empty", , "Error"
    
        Exit Sub
    
    End If
    
    If IsNumeric(NoofChaptersText.Text) = False Then
        
        MsgBox "Please enter valid Number of Chapters", , "Error"
        
        NoofChaptersText.Text = ""
        
        NoofChaptersText.SetFocus
        
        Exit Sub
        
    Else
    
        If NoofChaptersText.Text <= 0 Then
        
            MsgBox "Please enter valid Number of Chapters", , "Error"
        
            NoofChaptersText.Text = ""
            
            NoofChaptersText.SetFocus
        
            Exit Sub
        
        End If
    
    End If

    AddSubjectRS.Open "SELECT SubjectCode FROM Subjects", MainModule.con
    
    For i = 1 To AddSubjectRS.RecordCount Step 1
    
        If SubjectCodeText.Text = AddSubjectRS!subjectcode Then
        
            MsgBox "Subject already present", , "Error"
            
            SubjectCodeText.Text = ""
            
            SubjectCodeText.SetFocus
            
            AddSubjectRS.Close
            
            Exit Sub
        
        End If
    
        AddSubjectRS.MoveNext
    
    Next
    
    AddSubjectRS.Close
    
    If IsNumeric(SubjectCodeText.Text) = False Then
    
        MsgBox "Please, enter valid Subject Code.", , "Error"
        
        SubjectCodeText.Text = ""
        
        SubjectCodeText.SetFocus
        
        Exit Sub
        
    Else
    
        If Len(SubjectCodeText.Text) < 5 Then
        
            MsgBox "Please, enter 5 digit Subject Code", , "Error"
            
            SubjectCodeText.SetFocus
            
            Exit Sub
        
        End If
    
    End If
    
    SQL = "INSERT INTO Subjects VALUES(" & SubjectCodeText.Text & ",'" & SubjectNameText.Text & "'," & NoofChaptersText.Text & ")"

    AddSubjectRS.Open SQL, MainModule.con, adOpenDynamic, adLockOptimistic
    
    SQL = "CREATE TABLE " & SubjectCodeText.Text & "(Chapter Number,Question Varchar(255),Repetition Number,Marks Number,Flag YESNO,Title Varchar(255))"
    
    AddSubjectRS.Open SQL, MainModule.con
    
    AddSubjectRS.Open "CREATE TABLE " & SubjectCodeText.Text & "BitPattern" & "(Roman Number,Subquestion Number,Chapter Number,Marks Number,Opt YESNO,Title Varchar(255))", MainModule.con
    
    MsgBox SubjectNameText.Text & " added."
    
    ResetCmd_Click

End Sub

Private Sub BackCmd_Click()
   
    Unload Me
    
    SubjectForm.Show

End Sub

Private Sub Form_Load()

    AddSubjects.Width = 9000
    
    AddSubjects.Height = 6500

End Sub

Private Sub ResetCmd_Click()

    SubjectCodeText.Text = ""
    
    SubjectNameText.Text = ""
    
    NoofChaptersText = ""

End Sub

Private Sub SubjectCodeText_KeyPress(KeyAscii As Integer)

    If Not (KeyAscii > 47 And KeyAscii < 57) And Not KeyAscii = 8 Then

        KeyAscii = 0

    End If

End Sub
