VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "The Cachinator"
   ClientHeight    =   6135
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9165
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   9165
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDestroyAll 
      Caption         =   "Destroy All!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2460
      TabIndex        =   5
      Top             =   5610
      Width           =   1065
   End
   Begin VB.CommandButton cmdFindIt 
      Caption         =   "Find Them!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   180
      TabIndex        =   4
      Top             =   5610
      Width           =   1065
   End
   Begin VB.CommandButton cmdDestroyIt 
      Caption         =   "Destroy It!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1320
      TabIndex        =   3
      Top             =   5610
      Width           =   1065
   End
   Begin MSComctlLib.ListView lvCache 
      Height          =   3915
      Left            =   3660
      TabIndex        =   2
      Top             =   1590
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   6906
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   12648447
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "FILENAME"
         Text            =   "File Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "URL"
         Text            =   "Source URL"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "EXPIRES"
         Text            =   "Expires"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Key             =   "LASTACCESS"
         Text            =   "Last Accessed"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      Height          =   5325
      Left            =   180
      Picture         =   "frmMain.frx":0000
      ScaleHeight     =   5265
      ScaleWidth      =   3315
      TabIndex        =   0
      Top             =   180
      Width           =   3375
      Begin VB.Label lblCaption 
         BackStyle       =   0  'Transparent
         Caption         =   "Vote"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   495
         Left            =   60
         TabIndex        =   1
         Top             =   4650
         Width           =   4095
      End
   End
   Begin VB.Label lblStatus 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3660
      TabIndex        =   14
      Top             =   5655
      Width           =   5415
   End
   Begin VB.Label lblLastAccessed 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   5280
      TabIndex        =   13
      Top             =   1170
      Width           =   3675
   End
   Begin VB.Label Label6 
      Caption         =   "Last Accessed:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3750
      TabIndex        =   12
      Top             =   1170
      Width           =   1335
   End
   Begin VB.Label lblExpires 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   5280
      TabIndex        =   11
      Top             =   850
      Width           =   3675
   End
   Begin VB.Label Label4 
      Caption         =   "Expires:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3750
      TabIndex        =   10
      Top             =   850
      Width           =   1335
   End
   Begin VB.Label lblSourceUrl 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   5280
      TabIndex        =   9
      Top             =   530
      Width           =   3675
   End
   Begin VB.Label Label2 
      Caption         =   "Source URL:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3750
      TabIndex        =   8
      Top             =   530
      Width           =   1335
   End
   Begin VB.Label lblFileName 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   5280
      TabIndex        =   7
      Top             =   210
      Width           =   3675
   End
   Begin VB.Label Label1 
      Caption         =   "File Name:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3750
      TabIndex        =   6
      Top             =   210
      Width           =   1335
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub FindCacheEntries()
    
    '' First clear the cache list view
    lvCache.ListItems.Clear
    
    ''disable our buttons
    cmdFindIt.Enabled = False
    cmdDestroyIt.Enabled = False
    cmdDestroyAll.Enabled = False
    
    ''and put up an hourglass mousepointer
    MousePointer = vbHourglass
    
    '' Next we want to enumerate all the things in the cache
    '' we have to call the FindFirstCacheEntry function once and then
    '' the FindNextCacheEntry until it returns false
    If FindFirstCacheEntry() Then
        '' add our item the first column is also the text field in a listview
        lvCache.ListItems.Add , , Cache.CachedEntryFileName
        
        '' now add the subitems to the listview
        With lvCache.ListItems(lvCache.ListItems.Count)
            .SubItems(1) = Cache.CachedEntrySourceURL
            .SubItems(2) = Cache.CachedEntryExpireTime
            .SubItems(3) = Cache.CachedEntryLastAccessTime
            
        End With
                
        '' now loop through the rest of the cache
        Do While Cache.FindNextCacheEntry
            '' add our item the first column is also the text field in a listview
            '' only add if the filename is valid
            If Cache.CachedEntryCacheType And &H1 Then
                lvCache.ListItems.Add , , IIf(Cache.CachedEntryFileName = vbNullString, Cache.CachedEntrySourceURL, Cache.CachedEntryFileName)
                
                '' now add the subitems to the listview
                With lvCache.ListItems(lvCache.ListItems.Count)
                    .SubItems(1) = Cache.CachedEntrySourceURL
                    .SubItems(2) = Cache.CachedEntryExpireTime
                    .SubItems(3) = Cache.CachedEntryLastAccessTime
                    
                End With
            End If
        Loop
        
    End If
    
    '' always remember to release the cache if you have used the findfirst / findnext
    '' functions
    Cache.ReleaseCache
    
    '' enable our buttons
    cmdFindIt.Enabled = True
    cmdDestroyIt.Enabled = True
    cmdDestroyAll.Enabled = True
    
    '' and show our mousepointer again
    MousePointer = vbArrow
    
    lblStatus = "Found " & lvCache.ListItems.Count & " cache entries"
    
    '' highlight first item
    If lvCache.ListItems.Count > 0 Then
        lvCache.ListItems(1).Selected = True
        Call RefreshCacheList
    End If
    
End Sub

Private Sub RefreshCacheList()
    '' set our tooltip text and labels

    lvCache.ToolTipText = lvCache.SelectedItem.Text
    lblFileName = lvCache.SelectedItem.Text
    lblSourceUrl = lvCache.SelectedItem.SubItems(1)
    lblExpires = lvCache.SelectedItem.SubItems(2)
    lblLastAccessed = lvCache.SelectedItem.SubItems(3)

End Sub

Private Sub cmdDestroyAll_Click()
    '' first prompt to make sure that the user wants to do this
    
    If MsgBox("Are you sure you want to delete all these cache items?", vbQuestion Or vbYesNoCancel, "Delete Entire Cache?") = vbYes Then
        cmdFindIt.Enabled = False
        cmdDestroyAll.Enabled = False
        cmdDestroyIt.Enabled = False
        MousePointer = vbHourglass
        
        Do While lvCache.ListItems.Count > 0
            Cache.DeleteCacheEntry lvCache.ListItems(1).SubItems(1)
            lvCache.ListItems.Remove 1
        Loop
        
        lblCaption = "Terminated"
        cmdFindIt.Enabled = True
        cmdDestroyAll.Enabled = True
        cmdDestroyIt.Enabled = True
        MousePointer = vbHourglass
    End If
    
End Sub

Private Sub cmdDestroyIt_Click()
    '' if there is something selected then
    If Not lvCache.SelectedItem Is Nothing Then
        Cache.DeleteCacheEntry lvCache.SelectedItem.SubItems(1)
        lvCache.ListItems.Remove lvCache.SelectedItem.Index
        lblStatus = "Cache entry removed"
        lblCaption = "Terminated"
    End If
    
    If lvCache.ListItems.Count > 0 Then
        lvCache.ListItems(1).Selected = True
        Call RefreshCacheList
    End If
    
End Sub

Private Sub cmdFindIt_Click()
    Call FindCacheEntries
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim lWait As Long
    
    '' give a goodbye treat
    lblCaption = "I'll be back"
    lWait = Timer + 1
    Do While Timer < lWait
        DoEvents
    Loop
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ''don't forget to release our data allocations and stuff
    Call Cache.ReleaseCache
    
End Sub

Private Sub lblFileName_Change()
    lblFileName.ToolTipText = lblFileName.Caption
    
End Sub

Private Sub lvCache_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Call RefreshCacheList
    
End Sub


