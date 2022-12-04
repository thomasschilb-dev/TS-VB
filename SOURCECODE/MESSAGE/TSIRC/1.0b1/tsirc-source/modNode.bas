Attribute VB_Name = "modNode"
'API Call for Sending the messages

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long

Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long

Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

'Constants for changing the treeview
Public Const GWL_STYLE = -16&
Public Const TVM_SETBKCOLOR = 4381&
Public Const TVM_GETBKCOLOR = 4383&
Public Const TVS_HASLINES = 2&
Public Const TV_FIRST As Long = &H1100
Public Const TVM_GETTEXTCOLOR As Long = (TV_FIRST + 32)
Public Const TVM_SETTEXTCOLOR As Long = (TV_FIRST + 30)
Public Const TVS_CHECKBOXES = &H100
Public Const TVS_TRACKSELECT = &H200
Public Function AddParentNode(tvlTemp As TreeView, strDescription) As Node
    Set AddParentNode = tvlTemp.Nodes.Add()
    AddParentNode.Text = strDescription
End Function
Public Function AddChildNode(tvlTemp As TreeView, strDescription As String, intPosition As Integer) As Node
    Set AddChildNode = tvlTemp.Nodes.Add(intPosition, tvwChild)
    AddChildNode.Text = strDescription
End Function

Public Sub ChangeColor(tvlTemp As TreeView, iRed As Integer, iGreen As Integer, iBlue As Integer)
    Call SendMessage(tvlTemp.hWnd, TVM_SETBKCOLOR, 0, ByVal RGB(iRed, iGreen, iBlue))
End Sub

Public Sub ChangeForeColor(tvlTemp As TreeView, iRed As Integer, iGreen As Integer, iBlue As Integer)
    Call SendMessage(tvlTemp.hWnd, TVM_SETTEXTCOLOR, 0, ByVal RGB(iRed, iGreen, iBlue))
End Sub

Private Sub SetTreeViewAttrib(tvlTemp As TreeView, ByVal Attrib As Long)
    Const GWL_STYLE As Long = -16
    Dim rStyle As Long
    rStyle = GetWindowLong(tvlTemp.hWnd, GWL_STYLE)
    rStyle = rStyle Or Attrib
    Call SetWindowLong(tvlTemp.hWnd, GWL_STYLE, rStyle)
End Sub

