x = MsgBox("Ceci est un exemple de bo�tes de dialogue en visual basic (vba) Le format est vbs pour vba script. Le bouton par d�faut est 1. Ce message est critique.", vbOKCancel + vbCritical + vbDefaultButton1, "Exemple")

If x = vbOK Then
	x = MsgBox("Vous avez r�pondu OK.", vbInformation, "Exemple")
Else
	x = MsgBox("Vous avez r�pondu Annuler.", vbInformation, "Exemple")
End If

x = MsgBox("Est-ce que vous aimez Windows ?", vbQuestion + vbDefaultButton2 + vbYesNo, "Sondage de satisfaction")

If x = vbYes Then
	x = MsgBox("OK d'ac.", vbInformation, "Exemple")
Else
	x = MsgBox("Ah, bon...", vbInformation, "Exemple")
End If

x = MsgBox("Ceci est une bo�te avec de l'aide o� le texte est align� � droite.", vbMsgBoxHelpButton + vbMsgBoxRight, "M�me le titre ^^ (sous win7)")

count = 0
Do While count < 5
	x = MsgBox("Ce message appar�tra 5 fois.", vbInfo + vbOKOnly, "Exemple")
	count = count + 1
Loop