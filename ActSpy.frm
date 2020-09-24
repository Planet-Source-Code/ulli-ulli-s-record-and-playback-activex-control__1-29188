VERSION 5.00
Object = "{C9E395D1-208D-11D4-921C-525400E3EBE8}#15.0#0"; "Recorder.ocx"
Begin VB.Form fActivitySpy 
   Caption         =   "Stopped"
   ClientHeight    =   2385
   ClientLeft      =   1545
   ClientTop       =   1935
   ClientWidth     =   4710
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Moveable        =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   4710
   StartUpPosition =   2  'Bildschirmmitte
   Begin Recorder.RecPlay RecPlay1 
      Left            =   120
      Top             =   1815
      _ExtentX        =   1296
      _ExtentY        =   820
   End
   Begin VB.CheckBox ckStealth 
      Caption         =   "Stealth &Mode"
      Height          =   375
      Left            =   1710
      Style           =   1  'Grafisch
      TabIndex        =   4
      ToolTipText     =   "Run recording session in stealth mode"
      Top             =   660
      Width           =   1245
   End
   Begin VB.CheckBox ckKbd 
      Caption         =   "&Keyboard only"
      Height          =   375
      Left            =   300
      Style           =   1  'Grafisch
      TabIndex        =   3
      ToolTipText     =   "Record keyboard activity only"
      Top             =   660
      Width           =   1245
   End
   Begin VB.TextBox txTest 
      BackColor       =   &H00C0C0C0&
      Height          =   1110
      Left            =   1020
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   6
      ToolTipText     =   "Use this box (or any other app) for testing"
      Top             =   1170
      Width           =   3360
   End
   Begin VB.CommandButton btPlay 
      Caption         =   "Start &Play"
      Height          =   495
      Left            =   1710
      TabIndex        =   1
      ToolTipText     =   "Start playback"
      Top             =   105
      Width           =   1260
   End
   Begin VB.CommandButton btStop 
      Caption         =   "&Stop Rec/Play"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3135
      TabIndex        =   2
      ToolTipText     =   "Stop recordig or playback"
      Top             =   105
      Width           =   1260
   End
   Begin VB.CommandButton btRec 
      Caption         =   "Start &Rec"
      Height          =   495
      Left            =   285
      TabIndex        =   0
      ToolTipText     =   "Start recording session"
      Top             =   105
      Width           =   1260
   End
   Begin VB.Label lb 
      Caption         =   "Msgs"
      Height          =   195
      Index           =   1
      Left            =   4125
      TabIndex        =   8
      Top             =   705
      Width           =   375
   End
   Begin VB.Label lbMsgs 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fest Einfach
      Height          =   270
      Left            =   3135
      TabIndex        =   7
      ToolTipText     =   "Number of recorded events"
      Top             =   675
      Width           =   915
   End
   Begin VB.Label lb 
      Caption         =   "&Testbox and Replay"
      Height          =   405
      Index           =   0
      Left            =   90
      TabIndex        =   5
      Top             =   1200
      Width           =   855
   End
End
Attribute VB_Name = "fActivitySpy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Private Const RecorderFile As String = "c:\test.rcd"

Private Sub btPlay_Click()
 
    If ckKbd = vbChecked Then   'expecting a kbd only playback
        txTest.SetFocus         'so that keyboard playback is put in txTest
    End If
    Caption = "Playback"
    btRec.Enabled = False
    btStop.Enabled = True
    btPlay.Enabled = False
    ckKbd.Enabled = False
    ckStealth.Enabled = False
    txTest = ""
    lbMsgs = ""
    With RecPlay1
        .AppInstance = App.hInstance
        '.FileName = RecorderFile
        .StartPlayback
    End With 'RECPLAY1

End Sub

Private Sub btRec_Click()
    
    Caption = "Recording..."
    btRec.Enabled = False
    btStop.Enabled = True
    btPlay.Enabled = False
    ckKbd.Enabled = False
    ckStealth.Enabled = False
    txTest = ""
    lbMsgs = ""
    With RecPlay1
        .AppInstance = App.hInstance
        .KbdOnly = (ckKbd = vbChecked)
        '.FileName = RecorderFile
        If ckStealth = vbChecked Then
            fActivitySpy.Visible = False
            .Stealth = True
          Else 'NOT CKSTEALTH...
            .Stealth = False
        End If
        .StartRecord
    End With 'RECPLAY1

End Sub

Private Sub btstop_Click()

    RecPlay1.Halt
    
End Sub

Private Sub ckKbd_Click()

    ckKbd.BackColor = IIf(ckKbd = vbChecked, vbHighlight, vb3DFace)

End Sub

Private Sub ckStealth_Click()

    ckStealth.BackColor = IIf(ckStealth = vbChecked, vbHighlight, vb3DFace)

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    If UnloadMode = vbFormControlMenu Then
        Cancel = RecPlay1.IsBusy
    End If
    
End Sub

Private Sub RecPlay1_Halted(ByVal NumMessages As Long)

    lbMsgs = NumMessages
    Caption = "Stopped"
    btRec.Enabled = True
    btStop.Enabled = False
    btPlay.Enabled = True
    ckKbd.Enabled = True
    ckStealth.Enabled = True
    btRec.SetFocus
    Select Case NumMessages
      Case -1
        MsgBox "No Recorder-File specified."
      Case -2
        MsgBox "Recorder is disabled."
    End Select
    
End Sub

Private Sub RecPlay1_ShiftF12()
        
    fActivitySpy.Visible = True
        
End Sub

Private Sub txTest_GotFocus()

    txTest.BackColor = vbWhite
    
End Sub

Private Sub txTest_LostFocus()

    txTest.BackColor = BackColor
    
End Sub

':) Ulli's VB Code Formatter V2.5.12 (24.11.2001 18:30:27) 2 + 111 = 113 Lines
