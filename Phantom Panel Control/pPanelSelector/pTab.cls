VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pTab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"pTabs"
'---------------------------------------------------------------------------------------
' Name      : pTab
' DateTime  : 8 July 2005 10:41
' Author    : Gary Noble
' Purpose   : Tab Properties Handler
'---------------------------------------------------------------------------------------
' History   : 12 July 2005 - Added Maskcolor For Icon/Picture Backgrounds
'---------------------------------------------------------------------------------------

Option Explicit

'API's Function Declarations
Private Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
Private mTabeventhandler As New ITabEventHandler
Private mvarTooltipText As String
Private mvarpicIcon As StdPicture
Private mvarCaption As String
Private mvarKey As Variant
Private mvarEmbededControl As Object
Private mvarRecLeft As Long
Private mvarRecTop As Long
Private mvarRecBottom As Long
Private mvarRecRight As Long
Private mvarPicMaskColor As Long

Public Property Let PicMaskcolor(ByVal vData As Long)
    mvarPicMaskColor = vData
    mTabeventhandler.RaiseChangeEvent ("PicIcon Mask Color Changed/Set")
End Property

Public Property Get PicMaskcolor() As Long
    Let PicMaskcolor = mvarPicMaskColor
End Property

Friend Property Let Eventhandler(ByVal IEventhandler As ITabEventHandler)
    Set mTabeventhandler = IEventhandler
End Property

Friend Property Let RecRight(ByVal vData As Long)
    mvarRecRight = vData
End Property


Friend Property Get RecRight() As Long
    RecRight = mvarRecRight
End Property

Friend Property Let RecBottom(ByVal vData As Long)
    mvarRecBottom = vData
End Property


Friend Property Get RecBottom() As Long
    RecBottom = mvarRecBottom
End Property

Friend Property Let RecTop(ByVal vData As Long)
    mvarRecTop = vData
End Property


Friend Property Get RecTop() As Long
    Let RecTop = mvarRecTop
End Property

Friend Property Let RecLeft(ByVal vData As Long)
    mvarRecLeft = vData
End Property

Friend Property Get RecLeft() As Long
    RecLeft = mvarRecLeft
End Property


Private Sub Class_Terminate()

    If Not mvarEmbededControl Is Nothing Then
        ' If Not mvarEmbededControl Is Nothing Then mvarEmbededControl.Visible = False
    End If

    
End Sub


Public Property Set EmbededControl(ByVal vData As Object)
    On Error GoTo EmbededControl_Error

    Dim hWndControl As Long

    If Not vData Is Nothing Then

        hWndControl = vData.hwnd

        Set mvarEmbededControl = vData
        mvarEmbededControl.Visible = False
        mTabeventhandler.RaiseChangeEvent ("Embebbed Control Added Or Changed")

    End If

    On Error GoTo 0
    Exit Property

EmbededControl_Error:

    Err.Raise vbObjectError + Err.Number, "Embedded Control", " ( Embedded Control Must Have A Valid Window Handle ) "

End Property


Public Property Get EmbededControl() As Object
    Set EmbededControl = mvarEmbededControl
End Property


Friend Property Let Key(ByVal vData As Variant)
    mvarKey = vData
End Property


Friend Property Set Key(ByVal vData As Variant)
    Set mvarKey = vData
End Property


Public Property Get Key() As Variant
    If IsObject(mvarKey) Then
        Set Key = mvarKey
    Else
        Key = mvarKey
    End If
End Property



Public Property Let Caption(ByVal vData As String)
    mvarCaption = vData
    mTabeventhandler.RaiseChangeEvent ("Caption Set Or Changed")
End Property


Public Property Get Caption() As String
    Let Caption = mvarCaption
End Property



Public Property Set picIcon(ByVal vData As StdPicture)
    Set mvarpicIcon = vData
    mTabeventhandler.RaiseChangeEvent ("Picture Added Or Changed")
End Property


Public Property Get picIcon() As StdPicture
    Set picIcon = mvarpicIcon
End Property



Public Property Let ToolTipText(ByVal vData As String)
    mvarTooltipText = vData
    mTabeventhandler.RaiseChangeEvent ("TooltipText Added Or Changed")
End Property


Public Property Get ToolTipText() As String
    Let ToolTipText = mvarTooltipText
End Property



