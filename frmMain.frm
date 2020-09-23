VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmMain 
   Caption         =   "CPU Usage - MadMatt"
   ClientHeight    =   8925
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   7095
   LinkTopic       =   "Form1"
   ScaleHeight     =   8925
   ScaleWidth      =   7095
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.ListView lstProcess 
      Height          =   8295
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   14631
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Processus"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Utilisation CPU"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Timer tmrRefresh 
      Interval        =   500
      Left            =   4440
      Top             =   120
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "List of the processes running and their CPU usage :"
      Height          =   195
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   3630
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'                                        Matthieu Napoli
'      /\   ^   .^. +---._  /\   ^   .^. .-----..-----.
'     /  \ / \ /   \|     \/  \ / \ /   \|__  __\__  __|
'    /    M   |  A  |  D  /    M   |  A  \ T  T   |  |
'   /  /\  /\ | |-\ |    /  /\  /\ | |-\  \|  |   |  |
'  /__/  ><  \|_|  \|__./__/  ><  \|_|  \__|__|   |__|
'                                   madmatt_12@msn.com

' Program made by MadMatt
' clsPHDQuery class made by ShareVB (i lightly modified it)
' Thanks to ShareVB and the MSDN

' I'm french, so sorry if my english is not perfect ;-)


Dim PDHQuery As clsPDHQuery

Private Sub Form_Load()
    ' Create the query
    Set PDHQuery = New clsPDHQuery
    ' Make the list of the processes
    Dim tabID() As Long
    Dim i As Long
    Dim Counter As String
    Dim indexCounter As String
    Dim pName As String
    Dim Key As String
    Dim Ret As Long
    ' Get the processes list
    mpListProcess tabID()
    lstProcess.ListItems.Clear
    For i = 0 To UBound(tabID)
        ' Get the process name
        pName = mpGetProcessName(tabID(i))
        ' Create the counter
        Counter = PDHQuery.GetFormatedCounter(COUNTERPERF_PROCESS, GetFileNameWithoutExtension(pName), COUNTERPERF_PERCENTPROCESSORTIME)
        ' Add the counter
        Ret = PDHQuery.AddCounter(Counter)
        ' Get his index
        indexCounter = PDHQuery.Count - 1
        If indexCounter < 0 Then indexCounter = 0
        If Ret = -1 Then indexCounter = -1   ' in case of error
        ' We save the process ID and his counter index in the item key
        Key = str(tabID(i)) + "//" + str(indexCounter)
        ' Add the item to the listview
        lstProcess.ListItems.Add , Key, pName
    Next i
    ' Ask a first collect
    PDHQuery.Collect
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set PDHQuery = Nothing
End Sub

' Refresh the list
Private Sub tmrRefresh_Timer()
    ' Listitem
    Dim Item As ListItem
    Dim i As Long
    Dim str() As String
    Dim pName As String
    Dim Index As Long
    Dim Counter As String
    ' Get the CPU usage
    PDHQuery.Collect
    For i = 1 To lstProcess.ListItems.Count
        Set Item = lstProcess.ListItems.Item(i)
        ' Get the counter index
        str() = Split(Item.Key, "//")
        Index = Val(str(1))
        ' Fill the subitem
        If Index > -1 Then Item.SubItems(1) = PDHQuery.GetCounterData(Index)
    Next i
End Sub
