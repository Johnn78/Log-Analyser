VERSION 5.00
Begin VB.Form Log_Analyzer 
   Caption         =   "Log Analyzer"
   ClientHeight    =   7080
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   11505
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   161
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   7080
   ScaleWidth      =   11505
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Temp 
      Caption         =   "Temp"
      Height          =   375
      Left            =   10440
      TabIndex        =   15
      Top             =   6480
      Width           =   735
   End
   Begin VB.CommandButton cmdDetuct 
      Caption         =   "Detuct"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8880
      TabIndex        =   14
      Top             =   4200
      Width           =   1935
   End
   Begin VB.CommandButton cmdExtract 
      Caption         =   "Extract"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7800
      TabIndex        =   13
      Top             =   5040
      Width           =   1935
   End
   Begin VB.TextBox txtEnd 
      Height          =   615
      Left            =   5640
      TabIndex        =   12
      Top             =   5040
      Width           =   1500
   End
   Begin VB.TextBox txtBegin 
      Height          =   615
      Left            =   3240
      TabIndex        =   10
      Top             =   5040
      Width           =   1500
   End
   Begin VB.CommandButton cmdAnalyze 
      Caption         =   "Analyze"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6720
      TabIndex        =   8
      Top             =   4200
      Width           =   1935
   End
   Begin VB.TextBox txtKeyword 
      Height          =   615
      Left            =   3240
      TabIndex        =   7
      Top             =   4200
      Width           =   3015
   End
   Begin VB.TextBox txtLogPath 
      Height          =   495
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   3360
      Width           =   7695
   End
   Begin VB.FileListBox File 
      Height          =   2625
      Left            =   5280
      Pattern         =   "*.log;*.txt"
      TabIndex        =   3
      Top             =   600
      Width           =   5655
   End
   Begin VB.DirListBox Dir 
      Height          =   2115
      Left            =   2640
      TabIndex        =   2
      Top             =   600
      Width           =   2175
   End
   Begin VB.DriveListBox Drive 
      Height          =   315
      Left            =   480
      TabIndex        =   1
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label lblto 
      Caption         =   "to"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5040
      TabIndex        =   11
      Top             =   5160
      Width           =   255
   End
   Begin VB.Label lblextraact 
      Caption         =   "Extract From :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   9
      Top             =   5160
      Width           =   1935
   End
   Begin VB.Label lblKeyword 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Write the Keywords :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   480
      TabIndex        =   6
      Top             =   4320
      Width           =   2640
   End
   Begin VB.Label lblLogFile 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Log File Selected :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   480
      TabIndex        =   4
      Top             =   3360
      Width           =   2310
   End
   Begin VB.Label lblSelectLog 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select Log File"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   480
      TabIndex        =   0
      Top             =   600
      Width           =   1800
   End
End
Attribute VB_Name = "Log_Analyzer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim InputData As String
Dim Ext As String
Dim arrSplitStrings() As Variant
Dim i As Integer
Dim arrSplitKeyword As Variant
Dim detucted As Integer

Private Sub cmdAnalyze_Click()



Ext = Mid(txtLogPath.Text, InStrRev(txtLogPath.Text, "."), Len(txtLogPath.Text))
Open Replace(txtLogPath.Text, Ext, "_Analyzed" + Ext) For Output As #1
Open txtLogPath.Text For Input As #2

'Dim Count As Double

arrSplitKeyword = Split(txtKeyword.Text, ", ")

Do While Not EOF(2)
Line Input #2, InputData


'Count = Count + 1

For i = LBound(arrSplitKeyword) To UBound(arrSplitKeyword) Step 1

    If InStrRev(InputData, arrSplitKeyword(i)) <> 0 Then
    Print #1, InputData
    'Count = Count + 1
    End If
    
Next i
    
'Debug.Print Count
Loop

Debug.Print Count



Close #1
Close #2


MsgBox "The Analyzation has been completed Successfully ! ! !", 64, "Message"


End Sub



Private Sub cmdDetuct_Click()

Ext = Mid(txtLogPath.Text, InStrRev(txtLogPath.Text, "."), Len(txtLogPath.Text))
Open Replace(txtLogPath.Text, Ext, "_Detucted" + Ext) For Output As #1
Open txtLogPath.Text For Input As #2

arrSplitKeyword = Split(txtKeyword.Text, ", ")

Do While Not EOF(2)
Line Input #2, InputData

detucted = 0

For i = LBound(arrSplitKeyword) To UBound(arrSplitKeyword) Step 1

    If InStrRev(InputData, arrSplitKeyword(i)) = 0 Then
    detucted = detucted + 1
    End If
    
Next i

If detucted = UBound(arrSplitKeyword) + 1 Then
Print #1, InputData
End If

    
Loop

Close #1
Close #2

For i = LBound(arrSplitKeyword) To UBound(arrSplitKeyword) Step 1
Debug.Print arrSplitKeyword(i) & vbCrLf
Next i

MsgBox "The Detuction has been completed Successfully ! ! !", 64, "Message"

End Sub

Private Sub cmdExtract_Click()

'On Error GoTo HandleErrors


Dim Count As Double

Ext = Mid(txtLogPath.Text, InStrRev(txtLogPath.Text, "."), Len(txtLogPath.Text))
Open Replace(txtLogPath.Text, Ext, "_Extracted" + Ext) For Output As #1

Open txtLogPath.Text For Input As #2

Do While Not EOF(2)
Line Input #2, InputData



    If InStrRev(InputData, txtBegin.Text) <> 0 And InStrRev(InputData, txtEnd.Text) <> 0 And InStrRev(InputData, txtBegin.Text) <= InStrRev(InputData, txtEnd.Text) Then
    Print #1, Mid(InputData, InStrRev(InputData, txtBegin.Text) - 34, InStrRev(InputData, txtEnd.Text) - InStrRev(InputData, txtBegin.Text) + 39)
    
    'We give as many char we can because it finds the last char which exists & not the first and by writing as many chars then we have more accuracy to extract the text we want
    'Mid("abcd", 2, 2) -> bc
    'InStrRev("Hello", "l") -> 4
    
    'Debug.Print InputData
    'Debug.Print InStrRev(InputData, txtBegin.Text)
    'Debug.Print InStrRev(InputData, txtEnd.Text)
    


    
    End If
    


Loop

Close #1
Close #2

MsgBox "The Exctraction has been completed Successfully ! ! !", 64, "Message"

Exit Sub


HandleErrors:
Print #1, Mid(InputData, InStrRev(InputData, txtBegin.Text) + 13, InStrRev(InputData, txtEnd.Text) - InStrRev(InputData, txtBegin.Text) + 15)
Resume Next



End Sub


Private Sub cmdTemp_Click()


Ext = Mid(txtLogPath.Text, InStrRev(txtLogPath.Text, "."), Len(txtLogPath.Text))
Open Replace(txtLogPath.Text, Ext, "_deleted" + Ext) For Output As #1

Open txtLogPath.Text For Input As #2

Do While Not EOF(2)
Line Input #2, InputData



    If InStrRev(InputData, txtKeyword.Text) = 0 Then
    Print #1, InputData
    End If
    

Loop

Close #1
Close #2


End Sub

Private Sub Temp_Click()
    Dim S As String
    Dim arrSplitStrings1 As Variant
    
    S = "Hello World wHAT"
    arrSplitStrings1 = Split(S, " ")
    
    Debug.Print arrSplitStrings1(1)
    'Debug.Print UBound(arrSplitStrings1)
End Sub

Private Sub Dir_Change()
File.Path = Dir.Path
End Sub

Private Sub Drive_Change()
Dir.Path = Drive.Drive
End Sub

Private Sub File_Click()
txtLogPath.Text = File.Path + "\" + File.FileName
End Sub

Private Sub Form_Load()
Dir.Path = "C:\Users\" + Environ("USERNAME") + "\Desktop"
txtLogPath.Locked = True


End Sub

Private Sub Label1_Click()

End Sub


