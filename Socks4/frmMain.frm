VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Zyfir's Socks4 Proxy Server [OFF]"
   ClientHeight    =   3300
   ClientLeft      =   75
   ClientTop       =   360
   ClientWidth     =   6630
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   3300
   ScaleWidth      =   6630
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer timer_CheckClose 
      Left            =   960
      Top             =   0
   End
   Begin VB.CommandButton Cmd 
      Caption         =   "&command"
      Height          =   280
      Left            =   5520
      TabIndex        =   2
      Top             =   3000
      Width           =   1095
   End
   Begin MSComctlLib.ListView Log 
      Height          =   3015
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   5318
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Time"
         Object.Width           =   2205
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Source IP"
         Object.Width           =   2558
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "S.P."
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Destination IP"
         Object.Width           =   2558
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "D.P."
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Index"
         Object.Width           =   1235
      EndProperty
   End
   Begin MSWinsockLib.Winsock sockOut 
      Index           =   0
      Left            =   480
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock sockIn 
      Index           =   0
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label status 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Status: Ready"
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   3020
      Width           =   5535
   End
   Begin VB.Menu Hidden_Menu 
      Caption         =   "&Hidden_Menu"
      Visible         =   0   'False
      Begin VB.Menu MnuProxyStats 
         Caption         =   "Turn Proxy &On"
      End
      Begin VB.Menu MnuBreak1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuKillConnect 
         Caption         =   "&Terminate Connection"
      End
      Begin VB.Menu MnuSort 
         Caption         =   "&Sort Display"
         Begin VB.Menu MnuSortBy 
            Caption         =   "By &Time"
            Index           =   0
         End
         Begin VB.Menu MnuSortBy 
            Caption         =   "By &Source IP"
            Index           =   1
         End
         Begin VB.Menu MnuSortBy 
            Caption         =   "By Source &Port"
            Index           =   2
         End
         Begin VB.Menu MnuSortBy 
            Caption         =   "By &Destination IP"
            Index           =   3
         End
         Begin VB.Menu MnuSortBy 
            Caption         =   "By Destination p&ort"
            Index           =   4
         End
         Begin VB.Menu MnuSortBy 
            Caption         =   "By Socket &Index"
            Index           =   5
         End
      End
      Begin VB.Menu MnuBreak2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'#########################################################################
'## sorry about the no comments...                                      ##
'## this program wasnt meant for viewing source..                       ##
'## i just thought i would share it since no one has it yet..           ##
'## the coding isnt 100% neat and some of the stuff can be taken out... ##
'## like the extra .Refresh and some small stuff...                     ##
'## If you have any question... just ask or email me at:                ##
'## Zyfir@Myrealbox.com                                                 ##
'## Give me some credit for using this code at least :oþ                ##
'## ------------------------------------------------------------------- ##
'## THIS WAS MEANT FOR PERSONAL USAGE.. DONT MINE THE WEIRD CODING      ##
'#########################################################################
Private Sub Cmd_Click()
PopupMenu Hidden_Menu                                   '## Displays the Menu
End Sub

Private Sub Form_Load()
ReDim Preserve People(0)    '## Must Cover for the Index 0 for SockIN and SockOUT
End Sub

Private Sub Form_Unload(Cancel As Integer)
If MnuProxyStats.Caption <> "Turn Proxy &On" Then       '## Notify that the Server is on before exit
    If MsgBox("Proxy is still on. Are you sure you want to Exit?", vbYesNo + vbQuestion, "Exit") = vbYes Then End
Else                                                    '## All the other programs have it..why not this one?
    If MsgBox("Are you sure you want to Exit?", vbYesNo + vbQuestion, "Exit") = vbYes Then End
End If
Cancel = 1                                              '## If it reaches this...then user clicked NO
End Sub

Private Sub Log_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Call SortList(Log, ColumnHeader.Index - 1)          '## Sort the list
End Sub

Private Sub Log_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu Hidden_Menu            '## Displays the Menu
End Sub

Private Sub mnuExit_Click()
Unload Me                                               '## Activates the Form_Unload
End Sub

Private Sub MnuKillConnect_Click()
On Error Resume Next                                    '## Close all the Checked
Dim i As Integer                                        '## Checking Backward is a
    For i = Log.ListItems.Count To 1 Step -1            '## Safeguard...Forward will cause errors :ox
        If Log.ListItems(i).Checked Then CloseSocket (Log.ListItems(i).SubItems(5))
    Next i
    frmMain.status = "Status: Checked Connection Terminated"
End Sub

Private Sub MnuProxyStats_Click()
Dim i As Integer
    If MnuProxyStats.Caption = "Turn Proxy &On" Then
        MnuProxyStats.Caption = "Turn Proxy &Off"
        frmMain.Caption = "Zyfir's Socks4 Proxy Server [ON]" '## User Friendly
        Log.ListItems.Clear
        sockIn(0).Close
        timer_CheckClose.Interval = 100                 '## Check the Loop for Details
        sockIn(0).LocalPort = 1080                      '## Change this if desired.
        sockIn(0).Listen                                '## Ill Make an Option Window Later
        status = "Status: " & "Server On"
    Else
        MnuProxyStats.Enabled = False
        MnuProxyStats.Caption = "Closing..."
        status = "Status: " & "Server Closing..."
        For i = 0 To sockIn.UBound                      '## Close All The Socket
            CloseSocket (i)                             '##
        Next i                                          '##
        timer_CheckClose.Interval = 0
        MnuProxyStats.Caption = "Turn Proxy &On"
        frmMain.Caption = "Zyfir's Socks4 Proxy Server [OFF]"
        MnuProxyStats.Enabled = True
        status = "Status: " & "Server Off"
    End If
End Sub

Private Sub MnuSortBy_Click(Index As Integer)
Call SortList(Log, Index)                               '## Sort by what they Selected (I made the menu into an Array)
End Sub

Private Sub sockIn_DataArrival(Index As Integer, ByVal bytesTotal As Long)
On Error Resume Next
Dim strIncome As String, ByteIncome() As Byte
sockIn(Index).GetData strIncome
ByteIncome = strIncome
If People(Index).First Then
    People(Index).First = False
    '## Read the Protocol if you want to know the new few lines
    If Mid(strIncome, 1, 1) <> Chr(4) Or Mid(strIncome, 2, 1) <> Chr(1) Then sockIn(Index).Close: People(Index).Closed = True: Exit Sub
    People(Index).Port = 256 * ByteIncome(4) + ByteIncome(6)
    People(Index).IP = ByteIncome(8) & "." & ByteIncome(10) & "." & ByteIncome(12) & "." & ByteIncome(14)
    Call Add2List(Index, sockIn(Index).RemoteHostIP, sockIn(Index).RemotePort, People(Index).IP, People(Index).Port)
    '## Cheap way of finding what to send to the Source when the other side is connected
    People(Index).toSendBack = Mid(strIncome, 3, 6)
    sockOut(Index).Connect People(Index).IP, People(Index).Port '## Connects to the Destination
Else
    While People(Index).SendToOK = False    '## i dont know if this is just me or not...
        DoEvents                            '## but i do this in all my program...
    Wend                                    '## just to make sure everything is sent correctly
    DoEvents
    People(Index).SendToOK = False          '## to use with ^^
    sockOut(Index).SendData strIncome       '## Redirect to Destination
End If
End Sub

Private Sub sockOut_DataArrival(Index As Integer, ByVal bytesTotal As Long)
On Error Resume Next
Dim strIN As String
sockOut(Index).GetData strIN
While People(Index).SendFromOK = False      '## Same as for the SockIN
    DoEvents                                '##
Wend                                        '##
DoEvents
People(Index).SendFromOK = False
sockIn(Index).SendData strIN                '## Redirect to Source
End Sub

Private Sub sockIn_Close(Index As Integer)
If People(Index).Closed <> True Then        '## dont take this If out..
    People(Index).Closed = True             '## if you do.. youll have a loop
    sockOut(Index).Close                    '## that goes on forever.
    DeleteLog (Index)                       '## This Deletes the Item from the ListView
End If
End Sub

Private Sub sockIn_ConnectionRequest(Index As Integer, ByVal requestID As Long)
On Error Resume Next
Dim i As Integer
status = "Status: " & "Incoming Attempt From: " & sockIn(Index).RemoteHostIP
i = CreateNew                               '## Gets the Next Free Socket
sockIn(i).Close                             '## Close it to Accept
FixNew (i)                                  '## Must to Make sure everything is set right
sockIn(i).Accept requestID                  '## Accepts the connection
status = "Status: " & "Accepted Connection From: " & sockIn(Index).RemoteHostIP
End Sub

Private Sub sockIn_SendComplete(Index As Integer)
People(Index).SendFromOK = True             '## Triggers OK for Next Send to Source
End Sub


Private Sub sockOut_Close(Index As Integer)
If People(Index).Closed <> True Then        '## Again.. dont take out this IF..
    People(Index).Closed = True             '## Same reason as for the other one
    sockIn(Index).Close                     '##
End If                                      '##
End Sub

Private Sub sockOut_Connect(Index As Integer)
On Error Resume Next                        '## This tells the Source that the redirect was successful
sockIn(Index).SendData Chr(4) & Chr(90) & People(Index).toSendBack
End Sub


Private Sub sockOut_SendComplete(Index As Integer)
People(Index).SendToOK = True               '## Triggers OK for Next Send to Destination
End Sub

Private Sub timer_CheckClose_Timer()        '## This function is here just to make sure someone closes the connect.
    For i = 0 To UBound(People)             '## It Disconnects both end and deletes it from the ListView
        If People(i).Closed = False And (sockIn(i).State = sckClosed Or sockIn(i).State = sckClosed) Then
            People(i).Closed = True
            DeleteLog (i)
        End If
    Next i
End Sub
