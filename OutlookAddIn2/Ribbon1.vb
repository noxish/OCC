Imports Microsoft.Office.Tools.Ribbon
Imports OutlookAddIn2.Form1
Public Class Ribbon1

    Private Sub Ribbon1_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button1_Click(sender As Object, e As RibbonControlEventArgs) Handles Button1.Click
        Dim frm As New Form1
        frm.Show()
    End Sub
End Class
