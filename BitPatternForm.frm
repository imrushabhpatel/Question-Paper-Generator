VERSION 5.00
Begin VB.Form BitPatternForm 
   BorderStyle     =   0  'None
   Caption         =   "Bit Pattern"
   ClientHeight    =   7815
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14250
   LinkTopic       =   "Form1"
   ScaleHeight     =   7815
   ScaleWidth      =   14250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
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
      Left            =   7000
      TabIndex        =   4
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
      Left            =   4000
      TabIndex        =   3
      Top             =   0
      Width           =   2000
   End
   Begin VB.Label TitleLbl 
      Alignment       =   2  'Center
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
      Height          =   600
      Left            =   12000
      TabIndex        =   6
      Top             =   900
      Width           =   2000
   End
   Begin VB.Label MarksLbl 
      Alignment       =   2  'Center
      Caption         =   "Total Marks"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   9000
      TabIndex        =   5
      Top             =   900
      Width           =   2000
   End
   Begin VB.Label SQWOLbl 
      Alignment       =   2  'Center
      Caption         =   "Compulsory questions"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   6000
      TabIndex        =   2
      Top             =   900
      Width           =   2000
   End
   Begin VB.Label SQOLbl 
      Alignment       =   2  'Center
      Caption         =   "Total questions"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   3000
      TabIndex        =   1
      Top             =   900
      Width           =   2000
   End
   Begin VB.Label RomanLbl 
      Alignment       =   2  'Center
      Caption         =   "Roman"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   1000
      TabIndex        =   0
      Top             =   900
      Width           =   2000
   End
End
Attribute VB_Name = "BitPatternForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public R As Integer
Dim RString As String
Public RomanLbl1 As Label
Public SQOTxt, SQWOTxt, MarksTxt, TitleTxt As TextBox
Public X As Integer

Private Sub Form_Load()

    Me.Caption = MainModule.Subject & " Bit Pattern"
    
    MainModule.BitPatternRS.Open "SELECT * FROM " & MainModule.Subject & "BitPattern", MainModule.con
        
    If Not MainModule.BitPatternRS.RecordCount = 0 Then
    
        X = MsgBox("Do you want to delete existing bit pattern?", vbYesNo)
    
        If X = vbNo Then
        
            MainModule.BitPatternRS.Close
    
            Exit Sub
    
        End If
    
    End If
    
    MainModule.BitPatternRS.Close

    MainModule.SelSubject.Open "DELETE FROM " & MainModule.Subject & "BitPattern", MainModule.con

    Do

        RString = InputBox("Enter number of roman numbers:")
    
        If IsNumeric(RString) = False Then
    
            MsgBox "Please, enter proper roman numbers.", , "Error"
    
        End If
        
    Loop Until (IsNumeric(RString) = True)

    R = Int(RString)

    For i = 1 To R Step 1
    
        Set RomanLbl1 = Controls.Add("Vb.Label", "Label" & i)
    
        Set SQOTxt = Controls.Add("Vb.TextBox", "SQOText" & i)
        
        Set SQWOTxt = Controls.Add("Vb.TextBox", "SQWOText" & i)
        
        Set MarksTxt = Controls.Add("Vb.TextBox", "MarksText" & i)
    
        Set TitleTxt = Controls.Add("Vb.TextBox", "TitleText" & i)
    
        With RomanLbl1
        
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
        
        With SQOTxt
        
            .FontName = "Times New Roman"
            .FontSize = 12
            .Height = 500
            .Left = 3000
            .Top = i * 1000 + 1000
            .Visible = True
            .Width = 2000
        
        End With
        
        With SQWOTxt
        
            .FontName = "Times New Roman"
            .FontSize = 12
            .Height = 500
            .Left = 6000
            .Top = i * 1000 + 1000
            .Visible = True
            .Width = 2000
        
        End With

        With MarksTxt
        
            .FontName = "Times New Roman"
            .FontSize = 12
            .Height = 500
            .Left = 9000
            .Top = i * 1000 + 1000
            .Visible = True
            .Width = 2000
        
        End With

        With TitleTxt
        
            .FontName = "Times New Roman"
            .FontSize = 12
            .Height = 500
            .Left = 12000
            .Top = i * 1000 + 1000
            .Visible = True
            .Width = 2000
        
        End With

    Next
    
    ResetCmd.Top = MarksTxt.Top + 1000
    NextCmd.Top = MarksTxt.Top + 1000

    Me.Width = 16000
    Me.Height = MarksTxt.Top + 3000
    
End Sub

Private Sub NextCmd_Click()

    For i = 1 To R Step 1
    
        Set SQOTxt = Controls("SQOText" & i)
        
        Set SQWOTxt = Controls("SQWOText" & i)
        
        Set MarksTxt = Controls("MarksText" & i)
    
        If Not IsNumeric(SQOTxt.Text) Or Not IsNumeric(SQWOTxt.Text) Or Not IsNumeric(MarksTxt.Text) Then
        
            MsgBox "Enter numeric values only.", , "Error"
            
            Exit Sub
            
        Else
            
            If Val(SQOTxt.Text) < 1 Or Val(SQWOTxt.Text) < 1 Or Val(MarksTxt.Text) < 1 Then
            
                MsgBox "Please, enter values greater than zero.", , "Error"
            
                Exit Sub
            
            End If
            
            If SQOTxt.Text < SQWOTxt.Text Then
                
                MsgBox "Please, enter greater no of options.", , "Error"
                
                Exit Sub
            
            End If
        
        End If
        
    Next

    Me.Visible = False

    SubQuestionFrm.Show

End Sub
