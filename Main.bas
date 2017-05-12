Attribute VB_Name = "MainModule"
Public con As New ADODB.Connection
Public Subjects As New ADODB.Recordset
Public SelSubject As New ADODB.Recordset
Public BitPatternRS As New ADODB.Recordset
Public RNo As Integer
Public Subject As String

Sub Main()

    con.CursorLocation = adUseClient

    con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\QPG.mdb"
    
    If SubjectsRecordCount = 0 Then
    
        AddSubjects.Show
    
    Else
    
        SubjectForm.Show
    
    End If
    
End Sub

Function SubjectsRecordCount() As Integer

    Subjects.Open "SELECT * FROM Subjects", MainModule.con
    
        SubjectsRecordCount = Subjects.RecordCount

    Subjects.Close
    
End Function

Function SelSubjectRecordCount() As Integer

    SelSubject.Open "SELECT * FROM " & Subject, con

        SelSubjectRecordCount = SelSubject.RecordCount

    SelSubject.Close

End Function

Public Sub BitPatternModule()
    
    Dim TRNO As Integer
    
    TRNO = InputBox("Please, enter number of Roman numbers:")
    
    Dim MSQ, TSQNO As Integer
    
    For R = 0 To TRNO Step 1
    
        MSQ = InputBox("Enter marks for each subquestion for Roman No. " & (R + 1))

        TSQNO = InputBox("Enter total number of subquestions (with option) for Roman No. " & (R + 1))
    
        For s = 0 To TSQNO - 1 Step 1
        
            ip = InputBox("Enter Chapter No. for Question " & R + 1 & "." & s + 1)
            
            SQL = "INSERT INTO " & Subject & "BitPattern VALUES(" & (R + 1) & "," & (s + 1) & "," & ip & "," & MSQ & ")"

            BitPattern.Open SQL, con
        
        Next
    
    Next

End Sub

Sub f1()

    If RNo < BitPatternForm.R Then
        
        Exit Sub
    
    End If

    If RNo = BitPatternForm.R Then

        MsgBox "Bit pattern created"

    End If
    
    Unload SubQuestionFrm
    
    Unload BitPatternForm
    
    ChoiceForm.Show

End Sub
