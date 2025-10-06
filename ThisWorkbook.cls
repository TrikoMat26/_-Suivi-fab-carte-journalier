VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()
    Call AjouterMenuMacro
End Sub




'activation de la boite par double click dans la colonne observation
Private Sub Workbook_SheetBeforeDoubleClick(ByVal Sh As Object, ByVal Target As Range, Cancel As Boolean)

Cancel = True   'désactive le curseur dans la cellule

Dim Plage As Range
Dim Article As String



If Sh.Name = "CMS" Or Sh.Name = "Planning" Then

    Set Plage = Range("C:C")

    If Not Application.Intersect(Target, Plage) Is Nothing Then


            If ActiveCell.Row > 1 Then
                Article = ActiveCell.Value
                RechIPR Article
                
            End If

    Cancel = True
    End If

End If

End Sub

