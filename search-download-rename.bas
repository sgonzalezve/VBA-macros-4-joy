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



















######################
Option Explicit

Private WithEvents oItems As Outlook.Items

Private Sub Application_Startup()

    Dim OutlookApp As Outlook.Application
    Dim oNameSpace As Outlook.NameSpace
    Dim userId As String
    Dim inboxFolder As Outlook.Folder
    Dim subFolder As Outlook.Folder

    Set OutlookApp = Outlook.Application
    Set oNameSpace = Outlook.GetNamespace("MAPI")
    
    ' Suponiendo que tienes el ID del perfil del usuario
    userId = "100320039612BFBE" ' Reemplaza con el ID real del usuario
    
    ' Obtener la carpeta de inbox del perfil del usuario
    Set inboxFolder = oNameSpace.GetDefaultFolder(olFolderInbox)
    Set inboxFolder = inboxFolder.Folders(userId).Folders("Inbox")
    
    ' Obtener las subcarpetas específicas
    
    Set subFolder = subFolder.Folders("TECA")
    Set subFolder = subFolder.Folders("LLC")
    Set subFolder = subFolder.Folders("LCI")
    Set subFolder = subFolder.Folders("Caracara")
    Set subFolder = subFolder.Folders("LLN22")

    Set oItems = subFolder.Items

    Debug.Print "Desencadenador Iniciado" & VBA.Now

End Sub

Private Sub oItems_ItemAdd(ByVal Item As Object)

    Dim myMail As Outlook.MailItem
    Dim oAtt As Outlook.Attachment
    Dim folderPath As String

    If VBA.TypeName(Item) = "MailItem" Then

        Set myMail = Item
        
        ' Verifica el nombre de la carpeta y establece el path correspondiente
        Select Case myMail.Subject
            Case "*Reporte Integrado Operador CCE Teca*"
                folderPath = "S:\Sp_Project\39. KPIs_Meeting\2022\Weekly Follow Up\Supports and Working Files\Daily Reports\TECA\"
            Case "*DAILY REPORT*"
                folderPath = "S:\Sp_Project\39. KPIs_Meeting\2022\Weekly Follow Up\Supports and Working Files\Daily Reports\LLC\"
            Case "Reporte Diario de Producción EPF*"
                folderPath = "S:\Sp_Project\39. KPIs_Meeting\2024\Support and Working Files\Daily Reports\JAGUAR - CC\"
            Case "Reporte de Producción Caracara*"
                folderPath = "S:\Sp_Project\39. KPIs_Meeting\2024\Support and Working Files\Daily Reports\JAGUAR - LLN22\"
            Case "*REPORTE DIARIO DE PRODUCCIÓN Y OPERACIONES-GCT*"
                folderPath = "S:\Sp_Project\39. KPIs_Meeting\2023\Supports and Working Files\Daily Reports\LCI\_2024_&"
            Case Else
                ' Si no coincide con ninguna de las carpetas, no hagas nada
                Exit Sub
        End Select
        
        ' Guarda el archivo adjunto en el path especificado
        For Each oAtt In myMail.Attachments
            oAtt.SaveAsFile folderPath & oAtt.FileName
        Next oAtt
        
    End If
    
End Sub


