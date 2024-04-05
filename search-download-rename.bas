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
    
    If myMail.Subject Like "*" & "Reporte Integrado Operador CCE Teca" & "*" Then
    
        For Each oAtt In myMail.Attachments
            oAtt.SaveAsFile "S:\Sp_Project\39. KPIs_Meeting\2022\Weekly Follow Up\Supports and Working Files\Daily Reports\TECA\" & oAtt.FileName
        Next oAtt
    End If
    
End If


    
If VBA.TypeName(Item) = "MailItem" Then

    Set myMail = Item
    
    If myMail.Subject Like "*" & "DAILY REPORT" & "*" Then
    
        For Each oAtt In myMail.Attachments
            oAtt.SaveAsFile "S:\Sp_Project\39. KPIs_Meeting\2022\Weekly Follow Up\Supports and Working Files\Daily Reports\LLC\" & oAtt.FileName
        Next oAtt
    End If
    
End If

If VBA.TypeName(Item) = "MailItem" Then

    Set myMail = Item
    
    If myMail.Subject Like "*" & "REPORTE DIARIO DE PRODUCCIÓN Y OPERACIONES-GCT" & "*" Then
    
        For Each oAtt In myMail.Attachments
            oAtt.SaveAsFile "S:\Sp_Project\39. KPIs_Meeting\2023\Supports and Working Files\Daily Reports\LCI\" & "_2024_&" & oAtt.FileName
        Next oAtt
    End If
    
End If

    Set myMail = Nothing
    

End Sub
