Imports Microsoft.Office.Interop
Public Class Form1
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'TODO: Diese Codezeile lädt Daten in die Tabelle "TestDataSet3.Contacts". Sie können sie bei Bedarf verschieben oder entfernen.
        Me.ContactsTableAdapter.Fill(Me.TestDataSet3.Contacts)


    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        NewMail()
    End Sub

    Private Sub NewMail()

        Dim Outl As Object
        Outl = CreateObject("Outlook.Application")
        If Outl IsNot Nothing Then
            Dim omsg As Object
            omsg = Outl.CreateItem(0) '=Outlook.OlItemType.olMailItem'
            'set message properties here...'
            omsg.Display(True) 'will display message to user
        End If
    End Sub
End Class