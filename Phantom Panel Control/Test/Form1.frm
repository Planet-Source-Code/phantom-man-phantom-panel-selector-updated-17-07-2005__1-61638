VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "*\A..\pPanelSelector\Phantom.vbp"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form1 
   Caption         =   "Phantom Panel Selector Test"
   ClientHeight    =   4980
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6135
   FillColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   ScaleHeight     =   4980
   ScaleWidth      =   6135
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   5040
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   21
      Top             =   1320
      Width           =   975
      Begin MSComctlLib.TreeView TreeView1 
         Height          =   855
         Left            =   120
         TabIndex        =   9
         Top             =   120
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1508
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   353
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         HotTracking     =   -1  'True
         Appearance      =   0
      End
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   5040
      ScaleHeight     =   1455
      ScaleWidth      =   975
      TabIndex        =   14
      Top             =   2400
      Width           =   975
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Email: gwnoble@msn.com"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   19
         Top             =   1560
         Width           =   4215
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Version: 1.01"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   18
         Top             =   1080
         Width           =   4215
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Copyright: (c) 2005 "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   17
         Top             =   840
         Width           =   4215
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Author: Gary Noble"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   16
         Top             =   600
         Width           =   4215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "About Phantom Panel Selector..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   120
         Width           =   4095
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   5040
      ScaleHeight     =   1095
      ScaleWidth      =   1215
      TabIndex        =   13
      Top             =   120
      Width           =   1215
      Begin MSComCtl2.MonthView MonthView1 
         Height          =   4455
         Left            =   0
         TabIndex        =   20
         Top             =   120
         Width           =   2640
         _ExtentX        =   4657
         _ExtentY        =   7858
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   -2147483633
         Appearance      =   0
         MonthRows       =   2
         MultiSelect     =   -1  'True
         StartOfWeek     =   20578306
         CurrentDate     =   38543
      End
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Hide Tips"
      Height          =   375
      Left            =   3480
      TabIndex        =   12
      Top             =   4080
      Width           =   1300
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Change Style"
      Height          =   495
      Left            =   3480
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   10
      Top             =   3600
      Width           =   1300
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Use Custom Color"
      Height          =   495
      Left            =   3480
      TabIndex        =   8
      Top             =   3120
      Width           =   1300
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Panel Left/Right"
      Height          =   495
      Left            =   3480
      TabIndex        =   7
      Top             =   2640
      Width           =   1300
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Picture Size"
      Height          =   495
      Left            =   3480
      TabIndex        =   5
      Top             =   2160
      Width           =   1300
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Change Panel Caption"
      Height          =   495
      Left            =   3480
      TabIndex        =   6
      Top             =   1680
      Width           =   1300
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Add New Panel"
      Height          =   375
      Left            =   3480
      TabIndex        =   2
      Top             =   1320
      Width           =   1300
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Reload"
      Height          =   375
      Left            =   3480
      TabIndex        =   0
      Top             =   960
      Width           =   1300
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Clear"
      Height          =   375
      Left            =   3480
      TabIndex        =   1
      Top             =   600
      Width           =   1300
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   5040
      ScaleHeight     =   615
      ScaleWidth      =   1095
      TabIndex        =   4
      Top             =   4200
      Width           =   1095
      Begin VB.Image Image1 
         Height          =   2895
         Left            =   0
         Picture         =   "Form1.frx":0000
         Top             =   0
         Width           =   2880
      End
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   0
      Top             =   4440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1B282
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1B3DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1B536
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1B690
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1B7EA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   795
      Left            =   5040
      ScaleHeight     =   795
      ScaleWidth      =   1035
      TabIndex        =   3
      Top             =   3360
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00400040&
      Caption         =   "Disable"
      Height          =   375
      Left            =   3480
      TabIndex        =   11
      Top             =   240
      Width           =   1300
   End
   Begin PhantomPanel.PanelControl PanelControl1 
      Height          =   4530
      Left            =   120
      TabIndex        =   22
      Top             =   120
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   7990
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PanelStyle      =   2
      ToolTipStyle    =   1
      CustomColor     =   33023
      SelectedItemColor=   192
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Command1_Click()
    Dim c As pTab

    Dim i As Long

    On Error GoTo Command1_Click_Error

    Me.PanelControl1.Clear

    Me.PanelControl1.LockUpdate = True

    Set c = Me.PanelControl1.Panels.Add("View Your Mail Items", Me.ImageList2.ListImages(2).Picture, "Mail", 2, Picture4, True)
    Set c = Me.PanelControl1.Panels.Add("Add/Edit Delete Appointments", Me.ImageList2.ListImages(1).Picture, "Calendar", 1, Picture1)
    Set c = Me.PanelControl1.Panels.Add("View Your Contacts", Me.ImageList2.ListImages(4).Picture, "Contacts", 4, Picture3)
    Set c = Me.PanelControl1.Panels.Add("Information Regarding The Panel Selector Control" & vbCrLf & "Copyright Noticies etc...", Me.ImageList2.ListImages(3).Picture, "About...", 3, Picture5)

    Me.PanelControl1.LockUpdate = False


    Set c = Nothing

    On Error GoTo 0
    Exit Sub

Command1_Click_Error:

    MsgBox "The Following Error occured: " & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & _
           "Source: " & Err.Source
    Resume Next

End Sub

Private Sub Command10_Click()


    On Error GoTo Command2_Click_Error

    Me.PanelControl1.ShowTips = Not Me.PanelControl1.ShowTips
    
    If Me.PanelControl1.ShowTips Then
        Me.Command10.Caption = "Hide Tips"
    Else
        Me.Command10.Caption = "Show Tips"
    End If
    

    On Error GoTo 0
    Exit Sub

Command2_Click_Error:

    MsgBox "The Following Error occured: " & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & _
           "Source: " & Err.Source

End Sub

Private Sub Command2_Click()


    On Error GoTo Command2_Click_Error

    Me.PanelControl1.Clear


    On Error GoTo 0
    Exit Sub

Command2_Click_Error:

    MsgBox "The Following Error occured: " & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & _
           "Source: " & Err.Source

End Sub

Private Sub Command3_Click()

    Dim c As pTab
    On Error GoTo Command3_Click_Error

    Set c = Me.PanelControl1.Panels.Add("New Item Added", Me.ImageList2.ListImages(5).Picture, "New Item ( Panel Count: " & Me.PanelControl1.Panels.Count + 1 & ")", Me.PanelControl1.Panels.Count + 1, Me.Picture1)
    Set c = Nothing

    On Error GoTo 0
    Exit Sub

Command3_Click_Error:

    MsgBox "The Following Error occured: " & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & _
           "Source: " & Err.Source

End Sub

Private Sub Command4_Click()

    If Me.PanelControl1.PanelIconPictureSize = EPanelSize48x48_ Then
        Me.PanelControl1.PanelIconPictureSize = EPanelSize16x16_
    ElseIf Me.PanelControl1.PanelIconPictureSize = EPanelSize16x16_ Then
        Me.PanelControl1.PanelIconPictureSize = EPanelSize32x32_
    ElseIf Me.PanelControl1.PanelIconPictureSize = EPanelSize32x32_ Then
        Me.PanelControl1.PanelIconPictureSize = EPanelSize48x48_
    ElseIf Me.PanelControl1.PanelIconPictureSize = EPanelSize16x16_ Then
        Me.PanelControl1.PanelIconPictureSize = EPanelSize32x32_
    End If

End Sub

Private Sub Command5_Click()

    On Error GoTo Command5_Click_Error

    Me.PanelControl1.SelectedItem.Caption = "Changed!"


    On Error GoTo 0
    Exit Sub

Command5_Click_Error:

    MsgBox "The Following Error occured: " & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & _
           "Source: " & Err.Source

End Sub

Private Sub Command6_Click()

    If Me.PanelControl1.PanelAlignment = EDPA_Left Then
        Me.PanelControl1.PanelAlignment = EDPA_Right
    Else
        Me.PanelControl1.PanelAlignment = EDPA_Left
    End If

End Sub

Private Sub Command7_Click()

    Me.PanelControl1.UseCustomColor = Not Me.PanelControl1.UseCustomColor
    If Me.PanelControl1.UseCustomColor Then
        Me.Command7.Caption = "Use Windows Colors"
    Else
        Me.Command7.Caption = "Use Custom Color"
    End If
    

End Sub

Private Sub Command8_Click()

    If Me.PanelControl1.PanelStyle = ETDS_ClassicX Then
            Me.PanelControl1.SelectedItemColor = &HC0&
            Me.PanelControl1.PanelStyle = ETDS_MSMessenger
    
    ElseIf Me.PanelControl1.PanelStyle = ETDS_MSMessenger Then
            Me.PanelControl1.SelectedItemColor = &HC0&
            Me.PanelControl1.PanelStyle = ETDS_Classic
    
    Else
        Me.PanelControl1.SelectedItemColor = vbBlack
        Me.PanelControl1.PanelStyle = ETDS_ClassicX
    End If
        
End Sub

Private Sub Command9_Click()
    
    Me.PanelControl1.Enabled = Not Me.PanelControl1.Enabled
    If Me.PanelControl1.Enabled Then
        Command9.Caption = "Disable"
    Else
        Command9.Caption = "Enable"
    End If
    
    
End Sub

Private Sub Form_DblClick()


    On Error GoTo Form_DblClick_Error

    'Me.PanelControl1.Panels.Remove Me.PanelControl1.SelectedItem.Key

    On Error GoTo 0
    Exit Sub

Form_DblClick_Error:

    MsgBox "The Following Error occured: " & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & _
           "Source: " & Err.Source

End Sub

Private Sub Form_Load()
        
    Dim oNode As Node
    Dim oNodeChild As Node
    Dim oNodeChildx As Node
    
    Set oNode = Me.TreeView1.Nodes.Add(, , , "Inbox")
        oNode.Bold = True
    
    Set oNodeChild = Me.TreeView1.Nodes.Add(oNode, tvwChild, , "Unread Mail (3)")
    oNodeChild.Bold = True
    
    Set oNodeChild = Me.TreeView1.Nodes.Add(oNode, tvwChild, , "My Folders")
        Set oNodeChildx = Me.TreeView1.Nodes.Add(oNodeChild, tvwChild, , "Programming (127)")
        Set oNodeChildx = Me.TreeView1.Nodes.Add(oNodeChild, tvwChild, , "Info (152)")
    
    Set oNodeChild = Me.TreeView1.Nodes.Add(oNode, tvwChild, , "Drafts")
    Set oNodeChild = Me.TreeView1.Nodes.Add(oNode, tvwChild, , "Outbox")
    Set oNodeChild = Me.TreeView1.Nodes.Add(oNode, tvwChild, , "Sent Items (2136)")
    Set oNodeChild = Me.TreeView1.Nodes.Add(oNode, tvwChild, , "Deleted Items")
    
    oNode.Expanded = True
    
    
    Command1_Click
    
End Sub

Private Sub PanelControl1_PanelHovering(oTab As PhantomPanel.pTab)
    Debug.Print "Hovering Panel: " & oTab.Caption
End Sub

Private Sub PanelControl1_PanelSelected(oTab As PhantomPanel.pTab)
    Debug.Print "Panel Selected: " & oTab.Caption
End Sub

Private Sub Picture1_Resize()
MonthView1.Move 100, 0, Picture1.Top, Picture1.Height
End Sub

Private Sub Picture4_Resize()
TreeView1.Move 0, 0, Picture4.Width, Picture4.Height
End Sub
