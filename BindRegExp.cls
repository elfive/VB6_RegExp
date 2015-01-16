VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BindRegExp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Private mRegExp                As New RegExp

Private WithEvents mIgnoreCase As CheckBox
Attribute mIgnoreCase.VB_VarHelpID = -1

Private WithEvents mMultiLine  As CheckBox
Attribute mMultiLine.VB_VarHelpID = -1

Private WithEvents mGlobal     As CheckBox
Attribute mGlobal.VB_VarHelpID = -1

'Event Bindings
Public Sub Bind(ByRef chkIgnoreCase As CheckBox, _
                ByRef chkMultiLine As CheckBox, _
                ByRef chkGlobal As CheckBox)
    Set mIgnoreCase = chkIgnoreCase
    Set mMultiLine = chkMultiLine
    Set mGlobal = chkGlobal

    With mRegExp
        .IgnoreCase = IIf(mIgnoreCase.Value = 1, True, False)
        .Global = IIf(mGlobal.Value = 1, True, False)
        .MultiLine = IIf(mMultiLine.Value = 1, True, False)
    End With

End Sub

Private Sub mGlobal_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If mGlobal.Value = 1 Then
        mRegExp.Global = True
    Else
        mRegExp.Global = False
    End If

End Sub

Private Sub mIgnoreCase_MouseUp(Button As Integer, _
                                Shift As Integer, _
                                X As Single, _
                                Y As Single)

    If mIgnoreCase.Value = 1 Then
        mRegExp.IgnoreCase = True
    Else
        mRegExp.IgnoreCase = False
    End If

End Sub

Private Sub mMultiLine_MouseUp(Button As Integer, _
                               Shift As Integer, _
                               X As Single, _
                               Y As Single)

    If mMultiLine.Value = 1 Then
        mRegExp.MultiLine = True
    Else
        mRegExp.MultiLine = False
    End If

End Sub

Public Function Match(ByVal Patten As String, ByVal Sample As String) As MatchCollection

    With mRegExp
        .Pattern = Patten
        Set Match = .Execute(Sample)
    End With

End Function
