VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm 
   Caption         =   "系統查詢"
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4710
   OleObjectBlob   =   "UserForm.frx":0000
   StartUpPosition =   1  '所屬視窗中央
End
Attribute VB_Name = "UserForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnsearch_Click()
Dim name As String
Dim rownum As Integer
Dim content As String
Dim satus As Boolean
name = txbname.Text

For rownum = 2 To 7
If (Cells(rownum, "A").Value = name) Then
    lblresult.Caption = Cells(rownum, 4).Value
    If (Cells(rownum, "C").Value = "Y") Then
        satus = True
    Else
    satus = False
    End If
Else
End If
Next
content = "運作狀態" & satus
MsgBox (content)

End Sub
