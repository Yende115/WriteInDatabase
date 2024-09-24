Sub WriteInDatabase
	Dim cn As ADODB.Connection, rs As ADODB.Recordset, r As Long
    
    Set cn = New ADODB.Connection
    cn.Open "Provider=Microsoft.ACE.OLEDB.12.0;" & "Data Source=C:\path\to\database\database.accdb;"
    
    Set rs = New ADODB.Recordset
    rs.Open "tblLog", cn, adOpenKeyset, adLockOptimistic, adCmdTable
    With rs
        .AddNew
            .Fields("1eColonne") = Now 'Date et heure
            .Fields("2eColonne") = Application.UserName 'Nom de l utilisateur qui a ouvert le classeur
            .Fields("3eColonne") = Range(Prenom) & " " & Range(Nom) 'Champ "Prenom" et "Nom" avec un espace
            .Fields("4eColonne") = "Ceci est un texte" 'Texte
        .Update
    End With
    rs.Close
    Set rs = Nothing
    cn.Close
    Set cn = Nothing
End Sub
