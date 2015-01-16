VERSION 5.00
Begin VB.Form frmTestor 
   Caption         =   "Reg Expression Testor"
   ClientHeight    =   4785
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10755
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
   ScaleHeight     =   4785
   ScaleWidth      =   10755
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CheckBox chkGlobal 
      Caption         =   "Global"
      Height          =   195
      Left            =   2400
      TabIndex        =   6
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox txtPreView 
      Height          =   1215
      Left            =   6600
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   5
      Top             =   480
      Width           =   3975
   End
   Begin VB.ListBox lstFited 
      Height          =   2790
      Left            =   6600
      TabIndex        =   4
      Top             =   1800
      Width           =   3975
   End
   Begin VB.TextBox txtArticle 
      Height          =   2415
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Top             =   2160
      Width           =   6255
   End
   Begin VB.TextBox txtPatten 
      Height          =   1455
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   480
      Width           =   6255
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
End
Attribute VB_Name = "frmTestor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub chkGlobal_Click()
    Call ReMatch
End Sub

Private Sub chkIgnoreCase_Click()
    Call ReMatch
End Sub

Private Sub chkMulitiline_Click()
    Call ReMatch
End Sub

Private Sub Form_Load()
    Me.Show
    'init textbox bindings
    tbArticle.InitTextBox txtArticle
    tbPatten.InitTextBox txtPatten
    tbPreView.InitTextBox txtPreView
    tbReg.Bind chkIgnoreCase, chkMulitiline, chkGlobal
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

    '������������Զ�����
    On Error Resume Next

    'ƥ�伯��
    Dim tmpMC As MatchCollection

    'ѭ������
    Dim i     As Long

    '��ȡ��ѯ����ƥ�伯��
    Set tmpMC = tbReg.Match(txtPatten.Text, txtArticle.Text)
    '����б��
    lstFited.Clear
    '����ı���ʾ
    txtPreView.Text = ""

    '�����ѯû����������
    If Err.Number = 0 Then

        '��ʼ���������
        If tmpMC.Count > 0 Then

            For i = 0 To tmpMC.Count - 1
                lstFited.AddItem tmpMC.Item(i)
            Next i

        End If

        Me.Caption = "Reg Expression Testor"

        Exit Sub

    Else
    '����д������ڱ�������ʾ�ղŵĴ���
        Me.Caption = Err.Description
        Err.Clear
    End If

End Sub
