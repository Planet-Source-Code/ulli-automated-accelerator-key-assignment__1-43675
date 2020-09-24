VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form fHotkey 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fester Dialog
   ClientHeight    =   4410
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   6510
   ForeColor       =   &H00E0E0E0&
   Icon            =   "fHotkey.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   6510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CommandButton btBreak 
      BackColor       =   &H00C0C0FF&
      Caption         =   "&Break"
      Height          =   420
      Left            =   2280
      MousePointer    =   1  'Pfeil
      Style           =   1  'Grafisch
      TabIndex        =   8
      TabStop         =   0   'False
      ToolTipText     =   "Use this button to interrupt."
      Top             =   3090
      Width           =   675
   End
   Begin VB.Timer tmrBreak 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   2955
      Top             =   3090
   End
   Begin VB.Frame fr 
      BackColor       =   &H00E0E0E0&
      Caption         =   " Action "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00606060&
      Height          =   1440
      Index           =   2
      Left            =   3795
      TabIndex        =   9
      ToolTipText     =   "Select an action."
      Top             =   2775
      Width           =   2490
      Begin VB.CommandButton btApCa 
         BackColor       =   &H00C0C0FF&
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         Height          =   420
         Index           =   1
         Left            =   630
         Style           =   1  'Grafisch
         TabIndex        =   11
         ToolTipText     =   "Cancel and exit."
         Top             =   870
         Width           =   1260
      End
      Begin VB.CommandButton btApCa 
         BackColor       =   &H00C0FFC0&
         Caption         =   "&Apply"
         Default         =   -1  'True
         Height          =   420
         Index           =   0
         Left            =   630
         Style           =   1  'Grafisch
         TabIndex        =   10
         ToolTipText     =   "Apply and exit."
         Top             =   315
         Width           =   1260
      End
   End
   Begin VB.Frame fr 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   2520
      Index           =   1
      Left            =   225
      TabIndex        =   0
      Top             =   120
      Width           =   6060
      Begin MSFlexGridLib.MSFlexGrid fgdCaptions 
         Height          =   1935
         Left            =   225
         TabIndex        =   1
         ToolTipText     =   "Control names and their captions"
         Top             =   345
         Width           =   5625
         _ExtentX        =   9922
         _ExtentY        =   3413
         _Version        =   393216
         Rows            =   9
         FixedCols       =   0
         BackColor       =   15794175
         ForeColor       =   6316128
         BackColorFixed  =   13693183
         ForeColorFixed  =   96
         BackColorSel    =   15794175
         ForeColorSel    =   6316128
         BackColorBkg    =   15794175
         GridColorFixed  =   12632256
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         HighLight       =   0
         GridLinesFixed  =   1
         ScrollBars      =   2
         AllowUserResizing=   1
         BorderStyle     =   0
         Appearance      =   0
         FormatString    =   " Name                                     |Caption                                                              "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame fr 
      BackColor       =   &H00E0E0E0&
      Caption         =   " Hotkey Algorithm "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00006000&
      Height          =   1440
      Index           =   0
      Left            =   225
      TabIndex        =   2
      ToolTipText     =   "Select and run algorithm."
      Top             =   2775
      Width           =   3090
      Begin VB.OptionButton opAlgo 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&None"
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   3
         Left            =   195
         TabIndex        =   3
         ToolTipText     =   "This removes all hotkeys."
         Top             =   315
         Width           =   705
      End
      Begin VB.CommandButton btGo 
         BackColor       =   &H00C0FFC0&
         Caption         =   "&Go"
         Height          =   420
         Left            =   1185
         Style           =   1  'Grafisch
         TabIndex        =   7
         ToolTipText     =   "Run the selected algorithm."
         Top             =   315
         Width           =   675
      End
      Begin VB.OptionButton opAlgo 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Optimum   (may take very long)"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   2
         Left            =   195
         TabIndex        =   6
         ToolTipText     =   "This examines all permutations to find the best solution."
         Top             =   1080
         Width           =   2490
      End
      Begin VB.OptionButton opAlgo 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Smart   (usually quite fast)"
         ForeColor       =   &H00008000&
         Height          =   195
         Index           =   1
         Left            =   195
         TabIndex        =   5
         ToolTipText     =   "This makes as many passes as are necessary to resolve all clashes."
         Top             =   825
         Value           =   -1  'True
         Width           =   2115
      End
      Begin VB.OptionButton opAlgo 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Dumb"
         ForeColor       =   &H00008080&
         Height          =   195
         Index           =   0
         Left            =   195
         TabIndex        =   4
         ToolTipText     =   "This makes one pass over all captions. Clashes are not resolved."
         Top             =   570
         Width           =   735
      End
   End
   Begin VB.Menu mnuSend 
      Caption         =   "&Mail"
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&Info"
   End
End
Attribute VB_Name = "fHotkey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Tooltips        As New Collection

Private Type Element
    Name                As String
    OrgCaption          As String
    BckCaption          As String
    ModCaption          As String
    Permutation         As Long
    MissCount           As Long
End Type

Private Enum Modes
    Dumb = 0
    Smart = 1
    Optimum = 2
    None = 3
End Enum

Private Elements()      As Element
Private i               As Long
Private NumEle          As Long
Private StartTime       As Long
Private Busy            As Boolean
Private btBreakCapt     As String
Private Const Infinity  As Long = 2 ^ 31 - 1
Private Const MissScore As Long = 9999 'score for MissScore
Private Const Nulls     As String = vbNullChar & vbNullChar

Private Sub btBreak_Click()

    Busy = False

End Sub

Private Sub btGo_Click()

  Dim Idx               As Long
  Dim CurrScore         As Long
  Dim BestScore         As Long
  Dim PermCount         As Long
  Dim Caption           As String
  Dim ThisChar          As String
  Dim ThisCombi         As String
  Dim ThisCombiLwr      As String
  Dim BestCombi         As String
  Dim AlgoSmart         As Boolean
  Dim AlgoDumb          As Boolean

    MousePointer = vbHourglass
    StartTime = GetTickCount
    AlgoDumb = opAlgo(Dumb) 'storing them locally improves by about 15%
    AlgoSmart = opAlgo(Smart)
    fr(0).Enabled = False
    fr(1).Enabled = False
    fr(2).Enabled = False
    If opAlgo(None) = False Then
        Busy = True
        tmrBreak.Enabled = AlgoSmart
        For i = 1 To NumEle
            With Elements(i)
                .Permutation = i 'initial permutation is simply the normal sequence
                .ModCaption = Replace$(Replace$(.BckCaption, "&&", ""), " ", "")
            End With 'ELEMENTS(I)
        Next i

        BestScore = Infinity
        'loop thru all permutations
        'note that the number of permutations is the factorial of UBound(Elements) which
        'in turn is the same as the number of controls with a caption in the current form
        Do
            Inc PermCount
            Select Case PermCount Mod 65536
              Case 0
                btBreak.Caption = ""
              Case 32768
                btBreak.Caption = btBreakCapt
            End Select
            ThisCombi = ""
            CurrScore = 0
            'extract char
            For Idx = 1 To NumEle
                Caption = Elements(Elements(Idx).Permutation).ModCaption
                ThisCombiLwr = LCase$(ThisCombi) 'saves about 8%
                For i = 1 To Len(Caption)
                    ThisChar = Mid$(Caption, i, 1)
                    If InStr(ThisCombiLwr, LCase$(ThisChar)) = 0 Then 'ThisChar is not yet used
                        ThisCombi = ThisCombi & ThisChar
                        Inc CurrScore, i
                        Exit For '>---> Next
                    End If
                Next i
                If i > Len(Caption) Then 'this try was a miss
                    ThisCombi = ThisCombi & vbNullChar
                    Inc CurrScore, MissScore
                    If AlgoSmart Then 'smart mode
                        With Elements(Idx)
                            Inc .MissCount
                            If .MissCount > 99999 Then 'stop smart mode after 100000 misses
                                Busy = False
                            End If
                        End With 'ELEMENTS(Idx)
                    End If
                End If
                If CurrScore >= BestScore Then 'no use trying any more, saves another 10%
                    Exit For '>---> Next
                End If
            Next Idx

            'is this better?
            If CurrScore < BestScore Then
                BestScore = CurrScore
                're-arrange
                BestCombi = AllocString(Len(ThisCombi))
                For i = 1 To Len(ThisCombi)
                    Mid$(BestCombi, Elements(i).Permutation, 1) = Mid$(ThisCombi, i, 1)
                Next i
                If AlgoSmart And BestScore < MissScore Then 'smart and no miss - exit
                    Busy = False
                End If
            End If
            If AlgoDumb Then 'dumb - exit immediately
                Busy = False
            End If
            DoEvents 'to be able to react to btBreak
        Loop While MakeNextPermutation And Busy 'more permutations? still busy?

        tmrBreak.Enabled = False
        Busy = False
        btBreak.Caption = btBreakCapt
    End If
    For Idx = 1 To NumEle
        ThisChar = Mid$(BestCombi, Idx, 1)
        With Elements(Idx)
            .OrgCaption = Replace$(.BckCaption, ThisChar, "&" & ThisChar, , 1)
        End With 'ELEMENTS(IDX)
    Next Idx

    FillTheGrid (Not opAlgo(None))
    fr(0).Enabled = True
    fr(1).Enabled = True
    fr(2).Enabled = True
    MousePointer = vbDefault
    If GetTickCount - StartTime > 15000 Then 'took longer than 15 secs - wake up the user
        Beep
    End If
    ThisChar = ConvertToText(PermCount) & " permutation" & IIf(PermCount = 1, "", "s") & " examined."
    If BestScore < MissScore Then
        If PermCount > 9999 Then
            MsgBox ThisChar, vbInformation, App.ProductName
        End If
      Else 'NOT BESTSCORE...
        MsgBox ThisChar & vbCrLf & "Could not resolve all ambiguities.", vbExclamation, App.ProductName
    End If

End Sub

Private Sub btApCa_Click(Index As Integer)

    Tag = Index '0 = Apply, 1 = Cancel
    Hide

End Sub

Private Function ConvertToText(Num As Long) As String

  Dim NumText   As Variant

    NumText = Array("No", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine", "Ten")
    If Num <= UBound(NumText) Then
        ConvertToText = NumText(Num)
      Else 'NOT NUM...
        ConvertToText = Format$(Num, "#,0")
    End If

End Function

Private Sub fgdCaptions_GotFocus()

    fgdCaptions.TabStop = False

End Sub

Public Sub FillTheGrid(EnableHilite As Boolean)

  Dim CurrRow     As Long

    With fgdCaptions

        'prepare
        CurrRow = .TopRow
        .ColAlignment(-1) = flexAlignLeftCenter
        For i = 1 To NumEle
            .Rows = i + 1 'needs more rows
            .Row = i
            .Col = 0
            .Text = Elements(i).Name
            .Col = 1
            .Text = Elements(i).OrgCaption
            If EnableHilite And Elements(i).OrgCaption = Elements(i).BckCaption Then
                .CellBackColor = &HD0D0FF
              Else 'NOT ENABLEHILITE...
                .CellBackColor = &HF0FFFF
            End If
        Next i

        'sort it
        .Col = 0
        .ColSel = 0
        .Sort = flexSortStringNoCaseAscending
        .TopRow = CurrRow

        'adjust to fit
        .Height = IIf(.Rows > 8, 8, .Rows) * .RowHeight(1) + 15
        fr(1).Height = .Height + 600
        fr(0).Top = fr(1).Top + fr(1).Height + 150
        fr(2).Top = fr(0).Top
        btBreak.Top = fr(0).Top + 315
        Height = fr(0).Top + fr(0).Height + 975
    End With 'FGDCAPTIONS

    opAlgo(Optimum).Enabled = (NumEle < 11) 'else it would produce too many permutations

End Sub

Private Sub Form_Load()

    NumEle = 0
    Set Tooltips = CreateTooltips(Me)
    btBreakCapt = btBreak.Caption

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If Busy Then 'still busy
        Cancel = True
      Else 'BUSY = FALSE/0
        If UnloadMode <> vbFormCode Then
            Tag = "1" 'same as cancel
        End If
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set Tooltips = Nothing

End Sub

Public Function GetCapt(Posn As Long) As String

    GetCapt = Elements(Posn).OrgCaption

End Function

Public Property Let LetCapt(Capt As String, Name As String)

    Inc NumEle
    ReDim Preserve Elements(1 To NumEle)
    With Elements(NumEle)
        .Name = Name
        'save with removed ampersands (but leave double ampersands alone)
        .OrgCaption = Replace$(Replace$(Replace$(Capt, "&&", Nulls), "&", ""), Nulls, "&&")
        .BckCaption = .OrgCaption
    End With 'ELEMENTS(NUMELE)
    fr(1).Caption = " " & ConvertToText(NumEle) & " Control" & IIf(NumEle > 1, "s ", " ")

End Property

Private Function MakeNextPermutation() As Boolean

  Dim IdxLeft   As Long
  Dim IdxRight  As Long
  Dim Done      As Boolean
  Dim Temp      As Long

  'Note that the number of permutations is the factorial of UBound(Elements) which
  'in turn is the same as the number of controls with a caption in the current form

    IdxLeft = NumEle
    IdxRight = NumEle
    Do
        With Elements(IdxLeft)
            Temp = Elements(IdxRight).Permutation
            Select Case .Permutation
              Case Is = Temp 'start again
                Dec IdxLeft
                IdxRight = NumEle
                Done = (IdxLeft = 0) 'thats all
              Case Is > Temp 'retry
                Dec IdxRight
              Case Else 'propagate
                Elements(IdxRight).Permutation = .Permutation
                .Permutation = Temp
                Inc IdxLeft        'the elements right of IdxLeft (if any) are now in reverse order
                IdxRight = NumEle  'so we will correct that in the followning loop
                Do While IdxLeft < IdxRight
                    With Elements(IdxLeft)
                        Temp = .Permutation
                        .Permutation = Elements(IdxRight).Permutation
                        Elements(IdxRight).Permutation = Temp
                    End With 'ELEMENTS(IDXLEFT)
                    Inc IdxLeft
                    Dec IdxRight
                Loop
                Done = True 'we have the next permutation
            End Select
        End With 'ELEMENTS(IDXLEFT)
    Loop Until Done
    MakeNextPermutation = CBool(IdxLeft) 'returns True with legal permutations and False when done

End Function

Private Sub mnuAbout_Click()

    If Not Busy Then
        fgdCaptions.TabStop = True
        With App
            ShowAbout Me.hWnd, "About " & .ProductName & "#Operating System:", AppDetails & vbCrLf & .LegalCopyright, Me.Icon.Handle
        End With 'APP
    End If

End Sub

Private Sub mnuSend_Click()

    If Not Busy Then
        fgdCaptions.TabStop = True
        SendMeMail hWnd, AppDetails
    End If

End Sub

Private Sub tmrBreak_Timer()

  'stops smart mode after one minute

    Busy = False

End Sub

':) Ulli's VB Code Formatter V2.16.11 (2003-Mrz-04 23:24) 30 + 316 = 346 Lines
