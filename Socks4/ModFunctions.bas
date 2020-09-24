Attribute VB_Name = "ModFunctions"
Public Type eachPerson                          '## Good thing for (Structs)
    Port As Integer                             '## Or else there would have been
    IP As String                                '## 7 Different Array..for each Items
    Closed As Boolean
    SendFromOK As Boolean
    SendToOK As Boolean
    First As Boolean
    toSendBack As String
End Type
Public People() As eachPerson                   '## Define it to use it

Public Function CreateNew() As Integer
For i = 1 To UBound(People)
    If People(i).Closed = True Then             '## Checks if theres any closed Socket
        CreateNew = i                           '## Gets the First One to use.
        Exit Function                           '## Saves Alot of Socket :o)
    End If
Next i
ReDim Preserve People(0 To UBound(People) + 1)  '## Create 1 more if no Free Socket
CreateNew = UBound(People)                      '## Set the Index to return
Load frmMain.sockIn(CreateNew)                  '## Create another SockIN to use
Load frmMain.sockOut(CreateNew)                 '## Create another SockOUT to use
End Function
Public Function FixNew(Index As Integer)        '## Function used to set all the values
People(Index).SendFromOK = True                 '## To Default when created
People(Index).SendToOK = True
People(Index).Closed = False
People(Index).First = True
End Function

Public Sub Add2List(Index As Integer, strFrom As String, iFromPort As Integer, StrTo As String, iToPort As Integer)
Dim xitem As ListItem                                   '## Function used to Add
    Set xitem = frmMain.Log.ListItems.Add(Text:=Time)   '## To the ListView
        xitem.ListSubItems.Add Text:=strFrom
        xitem.ListSubItems.Add Text:=CStr(iFromPort)
        xitem.ListSubItems.Add Text:=StrTo
        xitem.ListSubItems.Add Text:=CStr(iToPort)
        xitem.ListSubItems.Add Text:=CStr(Index)
End Sub

Public Sub DeleteLog(Index As Integer)                  '## Find the Index in the ListView..
Dim i As Integer                                        '## Its the 5 subitem..and remove it..
    For i = 1 To frmMain.Log.ListItems.Count
        If frmMain.Log.ListItems(i).SubItems(5) = CStr(Index) Then
            frmMain.Log.ListItems.Remove i
            frmMain.Log.Refresh
            Exit For
        End If
    Next i
End Sub

Public Sub SortList(ctlListView As ListView, intColulunHeaderIndex As Integer)
ctlListView.Refresh    'Just Incase :o)                 '## Sort the ListView According
DoEvents: DoEvents     'Just Incase :ox                 '## To What they want to Sort By.
    ctlListView.Sorted = True
    ctlListView.SortKey = intColulunHeaderIndex
    If ctlListView.SortOrder = lvwAscending Then
        ctlListView.SortOrder = lvwDescending
    Else
        ctlListView.SortOrder = lvwAscending
    End If
    DoEvents: DoEvents  'Just Incase :ox
    ctlListView.Refresh                                 '## Updates.. Just Incase it doesnt redraw
End Sub
Public Sub CloseSocket(Index As Integer)                '## This closes the Socket
frmMain.status = "Status: Closing Socket: " & Index     '## That you Delete it to
frmMain.sockIn(Index).Close
frmMain.sockOut(Index).Close
People(Index).Closed = True                             '## Deletes it From the ListView
DeleteLog (Index)                                       '## If it Exist
frmMain.status = "Status: Closed Socket: " & Index
End Sub
