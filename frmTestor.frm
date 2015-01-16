VERSION 5.00
Begin VB.Form frmTestor 
   Caption         =   "Reg Expression Testor"
   ClientHeight    =   5640
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10695
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTestor.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5640
   ScaleWidth      =   10695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox txtAutoCode 
      Height          =   2775
      Left            =   5760
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   9
      Top             =   2760
      Width           =   4815
   End
   Begin VB.CheckBox chkGlobal 
      Caption         =   "Global"
      Height          =   195
      Left            =   2400
      TabIndex        =   6
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox txtPreView 
      Height          =   1935
      Left            =   5760
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   5
      Top             =   480
      Width           =   4815
   End
   Begin VB.ListBox lstFited 
      Height          =   2010
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   5535
   End
   Begin VB.TextBox txtArticle 
      Height          =   1575
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Text            =   "frmTestor.frx":08CA
      Top             =   3960
      Width           =   5535
   End
   Begin VB.TextBox txtPatten 
      Height          =   735
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Text            =   "frmTestor.frx":08EC
      Top             =   2880
      Width           =   5535
   End
   Begin VB.CheckBox chkMulitiline 
      Caption         =   "Mulitiline"
      Height          =   195
      Left            =   1320
      TabIndex        =   1
      Top             =   120
      Value           =   1  'Checked
      Width           =   975
   End
   Begin VB.CheckBox chkIgnoreCase 
      Caption         =   "IgnoreCase"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.Label lblFitDetail 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "匹配项详情↓"
      Height          =   195
      Left            =   5760
      TabIndex        =   12
      Top             =   240
      Width           =   1080
   End
   Begin VB.Label lblFitList 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "↓匹配项列表"
      Height          =   195
      Left            =   4560
      TabIndex        =   11
      Top             =   240
      Width           =   1080
   End
   Begin VB.Label lblVB 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "自动生成的VB代码"
      Height          =   195
      Left            =   5760
      TabIndex        =   10
      Top             =   2520
      Width           =   1440
   End
   Begin VB.Label lblMatchPatten 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "匹配模式字符串"
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   2640
      Width           =   1260
   End
   Begin VB.Label lblTestedStr 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "被处理字符串"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   3720
      Width           =   1080
   End
End
Attribute VB_Name = "frmTestor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub chkGlobal_MouseUp(Button As Integer, _
                              Shift As Integer, _
                              X As Single, _
                              Y As Single)
    Call ReMatch
End Sub

Private Sub chkIgnoreCase_MouseUp(Button As Integer, _
                                  Shift As Integer, _
                                  X As Single, _
                                  Y As Single)
    Call ReMatch
End Sub

Private Sub chkMulitiline_MouseUp(Button As Integer, _
                                  Shift As Integer, _
                                  X As Single, _
                                  Y As Single)
    Call ReMatch
End Sub

Private Sub Form_Load()
    Me.Show
    'init textbox bindings
    tbArticle.InitTextBox txtArticle
    tbPatten.InitTextBox txtPatten
    tbPreView.InitTextBox txtPreView
    tbVbCode.InitTextBox txtAutoCode
    tbReg.Bind chkIgnoreCase, chkMulitiline, chkGlobal
    Call ReMatch
End Sub

Private Sub Form_Unload(Cancel As Integer)

    End

End Sub

Private Sub lstFited_Click()

    If lstFited.ListCount > 0 Then
        If lstFited.ListIndex >= 0 Then
            txtPreView.Text = lstFited.List(lstFited.ListIndex)
        End If
    End If

End Sub

Private Sub txtArticle_KeyUp(KeyCode As Integer, Shift As Integer)
    Call ReMatch
End Sub

Private Sub txtPatten_KeyUp(KeyCode As Integer, Shift As Integer)
    Call ReMatch
End Sub

Public Sub ReMatch()

    '如果遇到错误自动继续
    On Error Resume Next

    '匹配集合
    Dim tmpMC As MatchCollection

    '循环变量
    Dim i     As Long

    '获取查询到的匹配集合
    Set tmpMC = tbReg.Match(txtPatten.Text, txtArticle.Text)
    '清空列表框
    lstFited.Clear
    '清空文本显示
    txtPreView.Text = ""
    '创建VB信息
    tbVbCode.SetText tbReg.ToVBCode

    '如果查询没有遇到错误
    If Err.Number = 0 Then

        '开始填充新内容
        If tmpMC.Count > 0 Then

            For i = 0 To tmpMC.Count - 1
                lstFited.AddItem tmpMC.Item(i)
            Next i

        End If

        Me.Caption = "Reg Expression Testor"
        
        Exit Sub

    Else
        '如果有错误，则在标题栏显示刚才的错误。
        Me.Caption = Err.Description
        Err.Clear
    End If

End Sub
