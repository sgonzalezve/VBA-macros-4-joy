Option Explicit

Private WithEvents oItems As Outlook.Items

Private Sub Application_Startup()

Dim OutlookApp As Outlook.Application
Dim oNameSpace As Outlook.NameSpace

Set OutlookApp = Outlook.Application
Set oNameSpace = Outlook.GetNamespace("MAPI")
Set oItems = oNameSpace.GetDefaultFolder(olFolderInbox).Items

Debug.Print "Desencadenador Iniciado" & VBA.Now


End Sub


Private Sub oItems_ItemAdd(ByVal Item As Object)

Dim myMail As Outlook.MailItem
Dim oAtt As Outlook.Attachment

If VBA.TypeName(Item) = "MailItem" Then

    Set myMail = Item
    
    If myMail.Subject Like "*" & "Name or pattern" & "*" Then
    
        For Each oAtt In myMail.Attachments
            oAtt.SaveAsFile "file/Path" & oAtt.FileName
        Next oAtt
    End If
    
End If


    
If VBA.TypeName(Item) = "MailItem" Then

    Set myMail = Item
    
    If myMail.Subject Like "*" & "Name or pattern" & "*" Then
    
        For Each oAtt In myMail.Attachments
            oAtt.SaveAsFile "file/path" & oAtt.FileName
        Next oAtt
    End If
    
End If

If VBA.TypeName(Item) = "MailItem" Then

    Set myMail = Item
    
    If myMail.Subject Like "*" & "New name or pattern to change name" & "*" Then
    
        For Each oAtt In myMail.Attachments
            oAtt.SaveAsFile "Path" & "_2024_&" & oAtt.FileName
        Next oAtt
    End If
    
End If

    Set myMail = Nothing
    

End Sub
