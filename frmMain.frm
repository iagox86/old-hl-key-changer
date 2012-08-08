VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "HL Key Changer"
   ClientHeight    =   1410
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2445
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1410
   ScaleWidth      =   2445
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Default         =   -1  'True
      Height          =   375
      Left            =   1320
      TabIndex        =   5
      Top             =   960
      Width           =   975
   End
   Begin VB.CommandButton cmdResetKey 
      Cancel          =   -1  'True
      Caption         =   "&Reset"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   960
      Width           =   975
   End
   Begin VB.TextBox txtCode3 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1680
      MaxLength       =   4
      TabIndex        =   3
      Top             =   480
      Width           =   615
   End
   Begin VB.TextBox txtCode2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   840
      MaxLength       =   5
      TabIndex        =   2
      Top             =   480
      Width           =   735
   End
   Begin VB.TextBox txtCode1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      MaxLength       =   4
      TabIndex        =   1
      Top             =   480
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Half-life Key:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit




Private Sub cmdResetKey_Click()
    Call Form_Load

    txtCode1.SetFocus
End Sub

Private Sub cmdSave_Click()
    Reg.regwrite strRegistry, txtCode1.Text & txtCode2.Text & txtCode3.Text
    'MsgBox txtCode1.Text & " " & txtCode2.Text & " " & txtCode3.Text
End Sub

'Reg.regwrite "HKEY_CURRENT_USER\Software\Ron\WSPS\History\history" & iindex, gHistoryList(iindex)


Private Sub Form_Load()
    Dim strTemp As String
    
    On Error GoTo handler
    
    Set Reg = CreateObject("wscript.shell")
    strTemp = Reg.regread(strRegistry)
    
    txtCode1.Text = Mid(strTemp, 1, 4)
    txtCode2.Text = Mid(strTemp, 5, 5)
    txtCode3.Text = Mid(strTemp, 10, 4)

    Exit Sub
    
handler:
    MsgBox "There was an error in reading the key.  It probably doesn't exist.", vbCritical, "Error!"

End Sub

Private Sub txtCode1_Change()
    On Error Resume Next
    If Len(txtCode1.Text) = 4 Then
        txtCode2.SetFocus
    End If
End Sub

Private Sub txtCode1_GotFocus()
    txtCode1.SelStart = 0
    txtCode1.SelLength = Len(txtCode1.Text)
End Sub

Private Sub txtCode1_KeyPress(KeyAscii As Integer)
    If KeyAscii < &H30 Or KeyAscii > &H39 Then
        KeyAscii = 0
    End If
End Sub
Private Sub txtCode2_KeyPress(KeyAscii As Integer)
    If KeyAscii < &H30 Or KeyAscii > &H39 Then
        KeyAscii = 0
    End If
End Sub
Private Sub txtCode3_KeyPress(KeyAscii As Integer)
    If KeyAscii < &H30 Or KeyAscii > &H39 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtCode2_GotFocus()
    txtCode2.SelStart = 0
    txtCode2.SelLength = Len(txtCode2.Text)
End Sub
Private Sub txtCode3_GotFocus()
    txtCode3.SelStart = 0
    txtCode3.SelLength = Len(txtCode3.Text)
End Sub

Private Sub txtCode2_Change()
    On Error Resume Next
    If Len(txtCode2.Text) = 5 Then
        txtCode3.SetFocus
    End If
End Sub

Private Sub txtCode3_Change()
    On Error Resume Next
    If Len(txtCode3.Text) = 4 Then
        cmdSave.SetFocus
    End If
End Sub
