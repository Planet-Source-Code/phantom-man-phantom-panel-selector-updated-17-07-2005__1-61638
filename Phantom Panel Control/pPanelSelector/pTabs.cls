VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pTabs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"pTab"
Attribute VB_Ext_KEY = "Member0" ,"pTab"
'---------------------------------------------------------------------------------------
' Name      : pTabs
' DateTime  : 08 July 2005 10:36
' Author    : Gary Noble
' Purpose   : Tabs collection Handler
'---------------------------------------------------------------------------------------
' History   :
'---------------------------------------------------------------------------------------

Option Explicit

Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long

'-- Implements Event Handler
'-- Enables the Tabcontrol To Redraw with the User Saying something Like (Ummmm) Redraw.
Implements ITabEventHandler

Private mCol As Collection

Private mvarParentControl As PhantomPanel.PanelControl

'-- Calling TabControl
Friend Property Set ParentControl(ByVal vData As PhantomPanel.PanelControl)
    Set mvarParentControl = vData
End Property


Friend Property Get ParentControl() As PhantomPanel.PanelControl
    Set ParentControl = mvarParentControl
End Property

'---------------------------------------------------------------------------------------
' Name      : Add
' DateTime  : 08 July 2005 10:34
' Author    : Gary Noble
' Purpose   : Adds A Tab To The Collection
'---------------------------------------------------------------------------------------
'
Public Function Add(ToolTipText As String, picIcon As Picture, Caption As String, Key As Variant, Optional EmbededControl As Object = Nothing, Optional Selected As Boolean = False, Optional PicMaskcolor As Long = -1) As pTab
'create a new object
    Dim objNewMember As pTab


    On Error GoTo Add_Error

    Set objNewMember = New pTab


    'set the properties passed into the method
    objNewMember.ToolTipText = ToolTipText
    If IsObject(picIcon) Then
        Set objNewMember.picIcon = picIcon
    Else
        Set objNewMember.picIcon = picIcon
    End If
    objNewMember.Caption = Caption
    If IsObject(Key) Then
        Set objNewMember.Key = Key
    Else
        objNewMember.Key = Key
    End If
    If Not EmbededControl Is Nothing Then

        If IsWindow(EmbededControl.hwnd) Then
            If IsObject(EmbededControl) Then
                Set objNewMember.EmbededControl = EmbededControl
            Else
                objNewMember.EmbededControl = EmbededControl
            End If
            SetParent EmbededControl.hwnd, Me.ParentControl.hwnd

        Else
            Err.Raise 12400, "Add Tab", "( Embedded Control Must Have A Valid Window Handle ) "
            Set objNewMember = Nothing
            Exit Function
        End If

    End If

    objNewMember.PicMaskcolor = PicMaskcolor
    
    objNewMember.Eventhandler = Me

    mCol.Add objNewMember, "V" & Key


    'return the object created
    Set Add = objNewMember

    If Selected Then
        Set mvarParentControl.SelectedItem = objNewMember
    Else
        If mvarParentControl.SelectedItem Is Nothing Then
            Set mvarParentControl.SelectedItem = objNewMember
            mvarParentControl.Redraw

        End If
    End If

    Set objNewMember = Nothing

    Me.ParentControl.Eventhandler "Panel Added"

    On Error GoTo 0
    Exit Function

Add_Error:

    If Err.Number = 438 Then
        Err.Raise Err.Number, App.ProductName & " _pTab Add", " ( Embedded Control Must Have A Valid Window Handle ) "
    Else
        Err.Raise Err.Number, App.ProductName & " _pTab Add", Err.Description
    End If
    Set Add = Nothing

End Function

Public Property Get Item(vntIndexKey As Variant) As pTab
Attribute Item.VB_UserMemId = 0
    Set Item = mCol(vntIndexKey)
End Property

Public Property Get Count() As Long
    Count = mCol.Count
End Property


Public Sub Remove(vntIndexKey As Variant)
    On Error Resume Next

    Dim oBJ As Object

    If Not mCol("V" & vntIndexKey).EmbededControl Is Nothing Then
        Set oBJ = mCol("V" & vntIndexKey).EmbededControl
    End If

    mCol.Remove "V" & vntIndexKey

    '-- Tell The TabControl That The Tab Has Been Removed
    mvarParentControl.ResetSelected

    '-- Hide The Embebbed control If any
    If Not oBJ Is Nothing Then oBJ.Visible = False

End Sub


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mCol.[_NewEnum]
End Property


Private Sub Class_Initialize()
    Set mCol = New Collection
End Sub


Private Sub Class_Terminate()
    Set mCol = Nothing

End Sub

Private Function ITabEventHandler_RaiseChangeEvent(sData As String) As Variant
    Me.ParentControl.Eventhandler sData

End Function
