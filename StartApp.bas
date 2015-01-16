Attribute VB_Name = "StartApp"

Public tbPatten  As New TextBoxHelper

Public tbPreView As New TextBoxHelper

Public tbArticle As New TextBoxHelper

Public tbReg     As New BindRegExp

Sub Main()
    Load frmTestor
    frmTestor.Show
End Sub
