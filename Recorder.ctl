VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.UserControl RecPlay 
   Appearance      =   0  '2D
   BackColor       =   &H80000005&
   CanGetFocus     =   0   'False
   ClientHeight    =   450
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1455
   ForwardFocus    =   -1  'True
   HasDC           =   0   'False
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   450
   ScaleWidth      =   1455
   ToolboxBitmap   =   "Recorder.ctx":0000
   Begin MSComDlg.CommonDialog dlgFile 
      Left            =   885
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   ".rcd"
      DialogTitle     =   "Select Recorder File"
      Filter          =   "Recorder Files (*.rcd)|*.rcd|All Files (*.*)|*.*"
   End
   Begin VB.Image imRec 
      BorderStyle     =   1  'Fest Einfach
      Height          =   465
      Left            =   0
      Picture         =   "Recorder.ctx":00FA
      Top             =   0
      Width           =   735
   End
End
Attribute VB_Name = "RecPlay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function RegisterServiceProcess Lib "Kernel32.dll" (ByVal dwProcessId As Long, ByVal dwType As Long) As Long

Private myEnabled     As Boolean
Private myAppInstance As Long
Private myFileName    As String
Private myIsRecording As Boolean
Private myIsPlaying   As Boolean
Private mySloMoFactor As Single
Private myKbdOnly     As Boolean
Private myStealth     As Boolean

Private Disabled      As Boolean

Public Event Halted(ByVal NumMessages As Long)
Public Event ShiftF12()

Friend Property Get AppInstance() As Long

    AppInstance = myAppInstance

End Property

Public Property Let AppInstance(ByVal nuAppInstance As Long)

    myAppInstance = nuAppInstance
    CheckAppInstance

End Property

Private Sub CheckAppInstance()

    If myAppInstance = 0 Then
        Err.Raise 380, , "Application Instance missing."
    End If

End Sub

Public Property Get Enabled() As Boolean

    Enabled = myEnabled

End Property

Public Property Let Enabled(ByVal nuEnabled As Boolean)

    myEnabled = (nuEnabled <> False) And Not Disabled
    PropertyChanged "Enabled"
    If myEnabled = False Then
        Halt
    End If

End Property

Public Property Get FileName() As String

    FileName = myFileName

End Property

Public Property Let FileName(ByVal nuFileName As String)

    myFileName = nuFileName
    PropertyChanged "FileName"

End Property

Friend Sub FireHalt(NumMessages As Long)

    RaiseEvent Halted(NumMessages)

End Sub

Friend Sub FireShiftF12()

    RaiseEvent ShiftF12
    If myStealth Then
        RegisterServiceProcess 0, 0
    End If

End Sub

Private Sub GetFilename()

    If Len(myFileName) = 0 Then
        dlgFile.Flags = dlgFile.Flags Or _
                        cdlOFNExplorer Or _
                        cdlOFNHelpButton Or _
                        cdlOFNLongNames Or _
                        cdlOFNPathMustExist
        On Error Resume Next
          If dlgFile.Flags And cdlOFNOverwritePrompt Then
              dlgFile.ShowSave
            Else 'NOT DLGFILE.FLAGS...
              dlgFile.ShowOpen
          End If
          If Err = 0 Then
              myFileName = dlgFile.FileName
          End If
          DoEvents
        On Error GoTo 0
    End If

End Sub

Public Sub Halt()

    If myIsRecording Then
        Hooks.RecordStop
    End If
    If myIsPlaying Then
        Hooks.PlaybackStop
    End If
    myAppInstance = 0

End Sub

Public Property Get IsBusy() As Variant

    IsBusy = myIsPlaying Or myIsRecording

End Property

Public Property Get IsPlaying() As Boolean

    IsPlaying = myIsPlaying

End Property

Friend Property Let IsPlaying(nuIsPlaying As Boolean)

    myIsPlaying = nuIsPlaying

End Property

Public Property Get IsRecording() As Boolean

    IsRecording = myIsRecording

End Property

Friend Property Let IsRecording(nuIsRecording As Boolean)

    myIsRecording = nuIsRecording

End Property

Public Property Get KbdOnly() As Boolean

    KbdOnly = myKbdOnly

End Property

Public Property Let KbdOnly(ByVal nuKbdOnly As Boolean)

    myKbdOnly = (nuKbdOnly = True)
    PropertyChanged "KbdOnly"

End Property

Public Property Get SloMoFactor() As Single

    SloMoFactor = mySloMoFactor

End Property

Public Property Let SloMoFactor(ByVal nuSloMoFactor As Single)

    mySloMoFactor = IIf(nuSloMoFactor < 0, 0, IIf(nuSloMoFactor > 10, 10, nuSloMoFactor))
    PropertyChanged "SloMoFactor"

End Property

Public Function StartPlayback() As Boolean

    If myEnabled And myAppInstance Then
        dlgFile.Flags = dlgFile.Flags And _
                        Not cdlOFNOverwritePrompt Or _
                        cdlOFNFileMustExist
        GetFilename
        If Len(FileName) Then
            Hooks.PlaybackStart Me
          Else 'LEN(FILENAME) = 0
            RaiseEvent Halted(-1&)
            CheckAppInstance
        End If
      Else 'NOT MYENABLED...
        RaiseEvent Halted(-2&)
        CheckAppInstance
    End If

End Function

Public Function StartRecord() As Boolean

    If myEnabled And myAppInstance Then
        dlgFile.Flags = dlgFile.Flags And _
                        Not cdlOFNFileMustExist Or _
                        cdlOFNOverwritePrompt
        GetFilename
        If Len(FileName) Then
            If myStealth Then
                MsgBox "To cancel Stealth Mode press Shift+F12 while recording.", vbInformation
                RegisterServiceProcess 0, 1
            End If
            Hooks.RecordStart Me, myKbdOnly
          Else 'NOT MYENABLED...'LEN(FILENAME) = 0
            RaiseEvent Halted(-1&)
            CheckAppInstance
        End If
      Else 'NOT ...'NOT MYENABLED...
        RaiseEvent Halted(-2&)
        CheckAppInstance
    End If

End Function

Public Property Get Stealth() As Boolean

    Stealth = myStealth

End Property

Public Property Let Stealth(ByVal Hide As Boolean)

    myStealth = Hide

End Property

Private Sub UserControl_InitProperties()
 
    myEnabled = True 'i am new here and i'm pretty much enabled
    mySloMoFactor = 1

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    With PropBag
        Enabled = PropBag.ReadProperty("Enabled", True)
        myFileName = PropBag.ReadProperty("FileName", "")
        mySloMoFactor = .ReadProperty("SloMoFactor", 1)
        myKbdOnly = .ReadProperty("KbdOnly", False)
        Stealth = .ReadProperty("Stealth", False)
    End With 'PROPBAG

End Sub

Private Sub UserControl_Resize()

    Size imRec.Width, imRec.Height 'you try it but i won't let you

End Sub

Private Sub UserControl_Show()

  Dim Control As Control
  Dim Cnt As Long

    'are any colleagues around
    For Each Control In Parent.Controls
        If TypeOf Control Is RecPlay Then
            Cnt = Cnt + 1
        End If
    Next Control
    If Cnt > 1 Then     'if only 1 then that's myself...
        Enabled = False '...but if i'm not the first then i won't do any work (i have my pride)
        Disabled = True
        MsgBox "One Recorder Control only, please.", vbCritical
    End If
 
End Sub

Private Sub UserControl_Terminate()
     
    Halt
     
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    With PropBag
        .WriteProperty "Enabled", myEnabled, True
        .WriteProperty "FileName", myFileName, ""
        .WriteProperty "SloMoFactor", mySloMoFactor, 1
        .WriteProperty "KbdOnly", myKbdOnly, False
        .WriteProperty "Stealth", myStealth, False
    End With 'PROPBAG

End Sub

':) Ulli's VB Code Formatter V2.5.12 (24.11.2001 18:30:23) 17 + 276 = 293 Lines
