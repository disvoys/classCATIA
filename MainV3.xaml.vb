Public Class MainV3



    Private Sub test_Click(sender As Object, e As RoutedEventArgs)

        'Initialize nouvelle instance CATIA
        Dim CCatia As New CatiaClass

        'Retourne le nom du document actif
        MsgBox(CCatia.getActiveDoc, Title:="Nom du Document actif")

        'Retourne le status de CATIA
        MsgBox(CCatia.getStatus, Title:="Status de CATIA (1, 2 ou 3)")

        'Retourne la langue CATIA
        MsgBox(CCatia.checkLangue, Title:="Langue de CATIA")

        'Générer la nommenclature
        CCatia.GetBOM(CATIA.ActiveDocument.product)  'Choisir n'importe quel product du CATIA (pas nécéssairement l'activeDocument
        Process.Start(My.Computer.FileSystem.SpecialDirectories.Temp & "\BOM_.txt")

        'Recupère le type de l'élément ouvert
        MsgBox(CCatia.getTypeElement(CATIA.ActiveDocument), Title:="Type de l'élément actif") 'Choisir n'importe quel document CATIA (pas nécéssairement l'activeDocument

        'Open a file and focus on windowCATIA
        CCatia.OpenFile(CATIA.ActiveDocument.FullName) 'Choisir n'importe quel fichier CATIA

        'Créer à partir d'un fichier (copie)
        CCatia.CreateFrom(CATIA.ActiveDocument.FullName) 'Choisir n'importe quel fichier CATIA

        'Focus et selectionne à partir d'une recherche
        CCatia.SelectCATIA("123D021234567899991") 'Choisir n'importe quel PartNumber à rechercher

        'fixer l'assemblage
        CCatia.fixAll(CATIA.ActiveDocument) 'Choisir n'importe quel ProductDocument CATIA (pas nécéssairement l'activeDocument


    End Sub







End Class