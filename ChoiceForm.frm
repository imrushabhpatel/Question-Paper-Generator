VERSION 5.00
Begin VB.Form ChoiceForm 
   Caption         =   "Choice"
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
   LinkTopic       =   "Form2"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton PrintPaperCmd 
      Caption         =   "Print Paper"
      Height          =   1000
      Left            =   5000
      TabIndex        =   2
      Top             =   1000
      Width           =   2000
   End
   Begin VB.CommandButton BitPatternCmd 
      Caption         =   "Bit Pattern"
      Height          =   1000
      Left            =   3000
      TabIndex        =   1
      Top             =   1000
      Width           =   2000
   End
   Begin VB.CommandButton QuestionBankCmd 
      Caption         =   "Question Bank"
      Height          =   1000
      Left            =   1000
      TabIndex        =   0
      Top             =   1000
      Width           =   2000
   End
End
Attribute VB_Name = "ChoiceForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim QuestionBank As New ADODB.Recordset
Dim QuestionBank1 As New ADODB.Recordset
Dim BitPattern As New ADODB.Recordset

Dim word_app As Word.Application
Dim word_doc As Word.Document

Dim chno, m, R As Integer
Dim Question As String

Private Sub BitPatternCmd_Click()

    BitPatternForm.Show

    If BitPatternForm.R = 0 Or BitPatternForm.X = vbNo Then
    
        Unload BitPatternForm
    
    End If

End Sub

Private Sub Form_Load()

    Me.Caption = Subject & " Choice"

    ChoiceForm.Height = 3500
    
    ChoiceForm.Width = 8000

End Sub

Private Sub PrintPaperCmd_Click()

    MainModule.SelSubject.Open "SELECT * FROM " & MainModule.Subject, MainModule.con

    MainModule.BitPatternRS.Open "SELECT * FROM " & MainModule.Subject & "BitPattern", MainModule.con

    If SelSubject.RecordCount = 0 Or BitPatternRS.RecordCount = 0 Then
    
        MsgBox "Question bank or Bitpattern not inserted.", , "Error"
    
        SelSubject.Close
        BitPatternRS.Close
        
        Exit Sub
    
    End If

    SelSubject.Close
    BitPatternRS.Close

    Dim TRNOvar As Integer

    Set word_app = New Word.Application
    
    Set word_doc = word_app.Documents.Add(Documenttype:=wdNewBlankDocument)

    SQL = "SELECT COUNT(*) AS TRNO FROM (SELECT DISTINCT Roman FROM " & MainModule.Subject & "BitPattern)"

    BitPattern.Open SQL, MainModule.con
    
    TRNOvar = BitPattern!TRNO
    
    If BitPattern.State = 1 Then BitPattern.Close
    
    word_app.Selection.Font.Name = "Times New Roman"
    
    word_doc.PageSetup.PaperSize = wdPaperA4
    
    For i = 1 To TRNOvar Step 1
    
        SQL = "SELECT Title FROM " & MainModule.Subject & "BitPattern WHERE Opt=-1 AND Roman=" & i
    
        BitPattern.Open SQL, MainModule.con
    
        If BitPatternRS.State = 1 Then BitPatternRS.Close
    
        MainModule.BitPatternRS.Open "SELECT * FROM " & MainModule.Subject & "BitPattern WHERE Roman=" & i, MainModule.con
        
        word_app.Selection.Font.Bold = True
    
        word_app.Selection.Font.AllCaps = True
    
        word_app.Selection.TypeText "Q" & i & ". " & BitPattern!Title
        
        X = BitPatternRS.RecordCount
    
        BitPatternRS.Close
    
        MainModule.BitPatternRS.Open "SELECT * FROM " & MainModule.Subject & "BitPattern WHERE Opt=-1 AND Roman=" & i, MainModule.con
        
        If Not BitPatternRS.RecordCount = X Then
        
            word_app.Selection.TypeText " (Any " & BitPattern.RecordCount & ")"
        
        End If
        
        word_app.Selection.TypeText "           " & ((BitPatternRS!Marks) * (BitPatternRS.RecordCount)) & " mks" & vbCrLf
    
        BitPatternRS.Close
    
        word_app.Selection.Font.AllCaps = False
    
        word_app.Selection.Font.Bold = False
    
        BitPattern.Close
        
        BitPattern.Open "SELECT COUNT(*) AS SQNO FROM " & MainModule.Subject & "BitPattern WHERE Roman =" & i, MainModule.con
    
        j = BitPattern!SQNO
    
        BitPattern.Close
    
        BitPattern.Open "SELECT Chapter,Marks,Title FROM " & MainModule.Subject & "BitPattern WHERE Roman=" & i, MainModule.con
    
        For k = 1 To j Step 1
        
            SQL = "SELECT Question FROM " & MainModule.Subject & " WHERE Chapter=" & BitPattern!Chapter & " AND Marks=" & BitPattern!Marks & " AND Flag=0 AND Title='" & BitPattern!Title & "'"
        
            QuestionBank.Open SQL, MainModule.con
            
            If QuestionBank.RecordCount = 0 Then
            
                MsgBox "Not enough questions Error", , "Error"
            
                QuestionBank.Close
            
                BitPattern.Close
            
                word_doc.Close (0)
                
                word_app.Quit
                
                Exit Sub
            
            End If
            
            Randomize
            
            Num = Int((QuestionBank.RecordCount * Rnd) + 1)
        
            For l = 2 To Num Step 1
            
                QuestionBank.MoveNext
            
            Next
        
            word_app.Selection.TypeText k & ") " & QuestionBank!Question & vbCrLf
            
            SQL = "UPDATE " & MainModule.Subject & " SET Flag=-1 WHERE Question='" & QuestionBank!Question & "'"
            
            QuestionBank1.Open SQL, MainModule.con
            
            QuestionBank.Close
            
            BitPattern.MoveNext
        
        Next
        
        BitPattern.Close
    
    Next
    
    QuestionBank.Open "UPDATE " & MainModule.Subject & " SET Flag=0", MainModule.con
    
    Do
    
        pass = InputBox("enter password")
    
        pass1 = InputBox("confirm password")
        
        If Not pass = pass1 Then
        
            MsgBox "Verify password.", , "Error"
        
        End If
        
    Loop Until (pass = pass1)
    
    word_doc.SaveAs FileName:=App.Path & "\" & MainModule.Subject & " QP.doc", Password:=pass
    
    MsgBox "Paper printed."

    word_doc.Close
    
    word_app.Quit

End Sub

Private Sub QuestionBankCmd_Click()

    Unload Me
    
    QuestionBankForm.Show
    
End Sub
