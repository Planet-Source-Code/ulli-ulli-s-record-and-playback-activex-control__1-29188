Attribute VB_Name = "Hooks"
Option Explicit
#Const Test = False
Private Const HC_ACTION As Long = 0
Private Const HC_GETNEXT As Long = 1
Private Const HC_SKIP As Long = 2
Private Const WH_JOURNALRECORD As Long = 0
Private Const WH_JOURNALPLAYBACK As Long = 1

Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long 'Caveat: this rolls over after about 50 days of
'                                                                Windows Up-Time and the time calculations will
Private Type typEventMsg                                        'fail if that happens while we're busily doin'it.
    Message           As Long                                   'but then - who has Windows up for 50 days in a row?
    ParamL            As Long
    ParamH            As Long
    Time              As Long
    hWnd              As Long
End Type

Private objRecPlay    As RecPlay
Private JournalMsg    As typEventMsg
Private Instance      As Long 'application instance of the hosting form
Private hHook         As Long 'the Current Hook
Private Delay         As Long
Private StartTime     As Long
Private NumMsgs       As Long
Private MsgNumber     As Long
Private hFile         As Long
Private ShiftF12Flag  As Long
Private MsgDelivered  As Boolean
Private KbdFilter     As Boolean
Private RecFileName   As String
Private SlowMoFact    As Single

Public Function PlaybackMessages(ByVal nCode As Long, ByVal wParam As Long, TheMessage As typEventMsg) As Long

  'called back by windows message system

    PlaybackMessages = 0                          'will return zero if nothing happens
    If nCode < 0 Then                             'message is not for me
        PlaybackMessages = CallNextHookEx(hHook, nCode, wParam, ByVal TheMessage)  'let somebody else do the work
      Else 'NOT NCODE...
        With JournalMsg
            Select Case nCode                     'what should i do ?
              Case HC_SKIP                        'get next message to replay
                If MsgDelivered Then
                    MsgDelivered = False
                    If MsgNumber = NumMsgs Then   'all done
                        PlaybackStop
                      Else 'NOT MSGNUMBER...
                        MsgNumber = MsgNumber + 1 'fetch next message from recorder file
                        Get hFile, , JournalMsg
                        .Time = .Time * SlowMoFact + StartTime 'relative time back to absolute time
                    End If
                End If
              Case HC_GETNEXT                     'deliver current message for replay
                TheMessage = JournalMsg           'this call can occur many times in a row so we have to
                Delay = .Time - GetTickCount()    'recalulate the time to wait before this msg is due
#If Test Then
                ''''''''''''''''''''''''''''''''''''
                If Delay > 10000 Then              'who wants to wait for more than 10 secs during tests ?
                    Delay = 10000                  'speed things up a bit
                    .Time = GetTickCount() + Delay 'adjust time in msg
                End If                             '
                ''''''''''''''''''''''''''''''''''''
#End If
                If Delay < 0 Then                 'we're in a hurry now
                    Delay = 0                     'replay immediately
                End If
                MsgDelivered = True
                PlaybackMessages = Delay          'tell windows how long to wait before processing this message
            End Select
        End With 'JOURNALMSG
    End If

End Function

Public Sub PlaybackStart(Cntl As RecPlay)

    Set objRecPlay = Cntl
    With objRecPlay
        If Not .IsBusy Then                        'don't disturb me while i'm busy
            .IsPlaying = True
            Instance = .AppInstance                'get data from control
            RecFileName = .FileName
            SlowMoFact = .SloMoFactor
            hFile = FreeFile()
            NumMsgs = 0
            Open RecFileName For Binary As hFile   'open recorder file
            Get hFile, , JournalMsg                'get marker
            With JournalMsg
                If .Message = &HAAAAAAAA And .ParamL = &HBBBBBBBB And .ParamH = &HCCCCCCCC And .Time = &HDDDDDDDD And .hWnd = &HEEEEEEEE Then
                    NumMsgs = LOF(hFile) / Len(JournalMsg) - 1   'how many messages are there in the file
                    If NumMsgs > 0 Then            'there are some - so go ahead
                        MsgNumber = 1
                        Get hFile, , JournalMsg    'get first message
                        MsgDelivered = False
                        StartTime = GetTickCount() 'what's the time now?
                        hHook = SetWindowsHookEx(WH_JOURNALPLAYBACK, AddressOf PlaybackMessages, Instance, 0) 'activate playback hook
                      Else                         'no messages there 'NOT NUMMSGS...
                        PlaybackStop
                    End If
                  Else                             'not a recorder file 'NOT .MESSAGE...
                    PlaybackStop
                End If
            End With 'JOURNALMSG
        End If
    End With 'OBJRECPLAY

End Sub

Public Sub PlaybackStop()

  'it may be a bit difficult to stop playing manually because the hardware is disabled during play.
  'you will either have to wait until replay stops by itself or press ctl alt del - the book says
  'that will stop play (can you believe it - Microsoft said so and it really works; well... sometimes)

    With objRecPlay
        If .IsPlaying Then
            If NumMsgs > 0 Then                   'replayed any messages
                UnhookWindowsHookEx hHook         'unhook the journal hook
            End If
            Close hFile
            If FileLen(RecFileName) = 0 Then         'if the recorder file is empty then kill it
                Kill RecFileName
            End If
            .FileName = ""
            .FireHalt MsgNumber                   'inform client
            .IsPlaying = False                    'phooh - done
        End If
    End With 'OBJRECPLAY
    Set objRecPlay = Nothing

End Sub

Public Function RecordMessages(ByVal nCode As Long, ByVal wParam As Long, TheMessage As typEventMsg) As Long

  'called back by windows message system

    If nCode = HC_ACTION Then                      'what should i do ?
        JournalMsg = TheMessage                    'save this message
        With JournalMsg
            If KbdFilter Then
                .Time = 0                          'hi speed playback
              Else 'KBDFILTER = 0
                .Time = .Time - StartTime          'relative time since start
            End If
            If .Message = 256 Then
                Select Case .ParamL
                  Case 10768                       'shift
                    ShiftF12Flag = 1
                  Case 22651                       'f12
                    ShiftF12Flag = ShiftF12Flag Or 2
                    If ShiftF12Flag = 3 Then
                        objRecPlay.FireShiftF12
                    End If
                  Case Else
                    ShiftF12Flag = 0
                End Select
            End If
            If (Not KbdFilter) Or (KbdFilter And (.Message = 256 Or .Message = 257)) Then
                MsgNumber = MsgNumber + 1          'count message
                Put hFile, , JournalMsg            'record message
            End If
        End With 'JOURNALMSG
        RecordMessages = 0                         'return zero
      Else                                         'this message is not for me 'NOT NCODE...
        RecordMessages = CallNextHookEx(hHook, nCode, wParam, ByVal TheMessage) 'let somebody else do the work
    End If

End Function

Public Sub RecordStart(Cntl As RecPlay, KbdOnly As Boolean)

    KbdFilter = KbdOnly
    Set objRecPlay = Cntl
    ShiftF12Flag = 0
    With objRecPlay
        If Not .IsBusy Then                        'don't disturb me while i'm busy
            .IsRecording = True
            Instance = .AppInstance                'get data from control
            RecFileName = .FileName
            On Error Resume Next
              Kill RecFileName                        'kill recorder file if it exists
            On Error GoTo 0
            hFile = FreeFile()
            Open RecFileName For Binary As hFile      'open new recorder file
            With JournalMsg                        'prepare header
                .Message = &HAAAAAAAA
                .ParamL = &HBBBBBBBB
                .ParamH = &HCCCCCCCC
                .Time = &HDDDDDDDD
                .hWnd = &HEEEEEEEE
            End With 'JOURNALMSG
            Put hFile, , JournalMsg                'put header
            MsgNumber = 0
            StartTime = GetTickCount()             'what's the time now ?
            hHook = SetWindowsHookEx(WH_JOURNALRECORD, AddressOf RecordMessages, Instance, 0) 'activate journal hook
        End If
    End With 'OBJRECPLAY

End Sub

Public Sub RecordStop()

    With objRecPlay
        If .IsRecording Then
            UnhookWindowsHookEx hHook              'unhook the journal hook
            Close hFile
            .FileName = ""                         'kill file name
            .FireHalt MsgNumber                    'inform client
            .IsRecording = False                   'phooh - done
        End If
    End With 'OBJRECPLAY
    Set objRecPlay = Nothing

End Sub

Public Sub Terminate()                             'the client is disappearing so close down

    With objRecPlay
        If .IsPlaying Then
            PlaybackStop
        End If
        If .IsRecording Then
            RecordStop
        End If
    End With 'OBJRECPLAY

End Sub

':) Ulli's VB Code Formatter V2.5.12 (24.11.2001 18:30:25) 35 + 199 = 234 Lines
