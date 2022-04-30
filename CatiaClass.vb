Imports System.ComponentModel
Imports System.IO
Imports System.Text
Imports CATMat
Imports DRAFTINGITF
Imports HybridShapeTypeLib
Imports INFITF
Imports KnowledgewareTypeLib
Imports MECMOD
Imports PARTITF
Imports ProductStructureTypeLib

Public Class CatiaClass


    Dim MaLangue As String = "Français"

    'Initialize nouvelle instance
    Sub New()
        CATIA = GetObject(, "CATIA.Application")
    End Sub

    'Get nom du document actif
    Function getActiveDoc() As String
        Dim s As String = CATIA.ActiveDocument.Name
        If s Is Nothing Then Return Nothing Else Return s
    End Function

    'Retourne le status de CATIA
    Function getStatus() As Integer
        Dim s As String
        If CATIA Is Nothing Then
            Return 0 'CATIA n'est pas ouvert
        ElseIf CATIA IsNot Nothing And getActiveDoc() Is Nothing Then
            Return 1 'Aucun document ouvert dans le CATIA actif
        Else
            Return 2 'CATIA est actif et document ouvert
        End If
    End Function

    'Recupère la langue utilisée dans CATIA
    Function checkLangue() As String

        Dim p As Product = CATIA.ActiveDocument.Product
        Dim NomFichier As String = My.Computer.FileSystem.SpecialDirectories.Temp & "\BOM.txt"
        Dim AssConvertor As AssemblyConvertor
        AssConvertor = p.GetItem("BillOfMaterial")
        Dim nullstr(0)
        AssConvertor.SetCurrentFormat(nullstr)
        Dim VarMaListNom(0)
        AssConvertor.SetSecondaryFormat(VarMaListNom)
        AssConvertor.Print("HTML", NomFichier, p)

        Dim fs As FileStream = Nothing
        If IO.File.Exists(NomFichier) Then
            Using sr As StreamReader = New StreamReader(NomFichier, Encoding.GetEncoding("iso-8859-1"))

                While Not sr.EndOfStream
                    Dim line As String = sr.ReadLine
                    If line Like "<b>Pièces différentes :*<br*" Then
                        MaLangue = "Français"
                        Return "Français"
                    ElseIf line Like "<b>Different parts:*<br*" Then
                        MaLangue = "Anglais"
                        Return "Anglais"
                    End If
                End While
                sr.Close()
            End Using
        End If

        Return "Langue non trouvée. A compléter."
    End Function

    'Genere la BOM
    Public Sub GetBOM(p As Product)

        Dim NomFichier As String = My.Computer.FileSystem.SpecialDirectories.Temp & "\BOM.txt"
        Dim AssConvertor As AssemblyConvertor
        AssConvertor = p.GetItem("BillOfMaterial")
        Dim nullstr(2)
        If MaLangue = "Anglais" Then
            nullstr(0) = "Part Number"
            nullstr(1) = "Quantity"
            nullstr(2) = "Type"
        ElseIf MaLangue = "Francais" Then
            nullstr(0) = "Référence"
            nullstr(1) = "Quantité"
            nullstr(2) = "Type"
        End If

        AssConvertor.SetCurrentFormat(nullstr)

        Dim VarMaListNom(1)
        If MaLangue = "Anglais" Then
            VarMaListNom(0) = "Part Number"
            VarMaListNom(1) = "Quantity"
        ElseIf MaLangue = "Français" Then
            VarMaListNom(0) = "Référence"
            VarMaListNom(1) = "Quantité"
        End If

        AssConvertor.SetSecondaryFormat(VarMaListNom)
        AssConvertor.Print("HTML", NomFichier, p)

        ModifFichierNomenclature(My.Computer.FileSystem.SpecialDirectories.Temp & "\BOM.txt")


    End Sub
    Sub ModifFichierNomenclature(txt As String)

        Dim strtocheck As String = ""
        If MaLangue = "Francais" Then
            strtocheck = "<b>Total des p"
        Else
            strtocheck = "<b>Total parts"
        End If

        Dim FichierNomenclature As String = My.Computer.FileSystem.SpecialDirectories.Temp & "\BOM_.txt"
        If IO.File.Exists(FichierNomenclature) Then
            IO.File.Delete(FichierNomenclature)
        End If
        Dim fs As FileStream = Nothing
        fs = New FileStream(FichierNomenclature, FileMode.CreateNew)
        Using sw As StreamWriter = New StreamWriter(fs, Encoding.GetEncoding("iso-8859-1"))
            If IO.File.Exists(txt) Then
                Using sr As StreamReader = New StreamReader(txt, Encoding.GetEncoding("iso-8859-1"))
                    Dim BoolStart As Boolean = False
                    While Not sr.EndOfStream
                        Dim line As String = sr.ReadLine
                        If Left(line, 8) = "<a name=" Then
                            If MaLangue = "Français" Then
                                line = "[" & Right(line, line.Length - 24)
                                line = Left(line, line.Length - 8)
                                line = line & "]"
                                sw.WriteLine(line)
                            Else
                                line = "[" & Right(line, line.Length - 27)
                                line = Left(line, line.Length - 8)
                                line = line & "]"
                                sw.WriteLine(line)
                            End If
                        ElseIf line Like "  <tr><td><A HREF=*</td> </tr>*" Then
                            line = Replace(line, "</td><td>Assembly</td> </tr>", "") 'pas fait
                            line = Replace(line, "</td><td>Assemblage</td> </tr> ", "")
                            line = Replace(line, "  <tr><td><A HREF=", "")
                            line = Replace(line, "</A></td><td>", ControlChars.Tab)
                            line = Replace(line, "#Bill of Material: ", "")
                            line = Replace(line, "#Nomenclature : ", "")
                            If line.Contains(">") Then
                                Dim lines() = Strings.Split(line, ">")
                                line = lines(1)
                            End If
                            Dim lines_() = Strings.Split(line, ControlChars.Tab)
                            line = lines_(0) & ControlChars.Tab & lines_(1)
                            If Strings.Left(line, 2) = "  " Then line = Strings.Right(line, line.Length - 2)
                            sw.WriteLine(line)
                        ElseIf Left(line, 14) = strtocheck Then
                            sw.WriteLine("[ALL-BOM-APPKD]")
                        ElseIf line Like "*<tr><td>*</td> </tr>*" Then
                            line = Replace(line, "<tr><td>", "")
                            line = Replace(line, "</td> </tr> ", "")
                            line = Replace(line, "</td><td>", ControlChars.Tab)
                            Dim lines_() = Strings.Split(line, ControlChars.Tab)
                            line = lines_(0) & ControlChars.Tab & lines_(1)
                            If Strings.Left(line, 2) = "  " Then line = Strings.Right(line, line.Length - 2)
                            sw.WriteLine(line)
                        Else
                            'nothing
                        End If

                    End While
                    sr.Close()
                End Using
            End If
            sw.Close()
        End Using

    End Sub

    'getType de l'élément
    Function getTypeElement(d As Document) As String

        Dim s() As String = Strings.Split(d.Name, ".")
        If s.Count > 0 Then
            Return s(UBound(s))
        Else
            Return Nothing
        End If

    End Function

    'Open and focus
    Sub OpenFile(file As String)
        Try
            AppActivate("CATIA V5 - [" & CATIA.ActiveDocument.Name & "]")
            CATIA.Documents.Open(file)
        Catch ex As Exception
            'erreur, vérifier le nom complet du fichier
        End Try
    End Sub

    'Creer nouvelle instance à partir de
    Sub CreateFrom(file As String)
        Try
            AppActivate("CATIA V5 - [" & CATIA.ActiveDocument.Name & "]")
            Dim d As Document = CATIA.Documents.NewFrom(file)
        Catch ex As Exception
            'erreur
        End Try
    End Sub


    'Selectionne et focus CATIA à partir d'une recherche
    Sub SelectCATIA(partNumber As String)
        Try
            AppActivate("CATIA V5 - [" & CATIA.ActiveDocument.Name & "]")
            CATIA.ActiveDocument.Selection.Clear()
            If MaLangue = "Anglais" Then
                CATIA.ActiveDocument.Selection.Search("Name='" & partNumber & "';all")
                CATIA.StartCommand("Reframe On")
                CATIA.StartCommand("Center graph")
            ElseIf MaLangue = "Français" Then
                CATIA.ActiveDocument.Selection.Search("Nom='" & partNumber & "';tout")
                CATIA.StartCommand("Centrer sur")
                CATIA.StartCommand("Centrer le graphe")
            End If
        Catch ex As Exception
            'erreur
        End Try
    End Sub



    'Fixer un assemblage
    Dim oList
    Dim oSelection As INFITF.Selection
    Dim oVisProp As VisPropertySet
    Dim Pint As Integer = 0
    Sub fixAll(oTopDoc As Document)

        'Declarations
        Dim oTopProd As ProductDocument = Nothing
        Dim oCurrentProd As Object

        'Check si c'est un assemblage
        If Strings.Right(oTopDoc.Name, 7) <> "Product" Then
            Exit Sub
        End If

        oSelection = oTopDoc.Selection
        oVisProp = oSelection.VisProperties
        oCurrentProd = oTopDoc.Product
        oList = CreateObject("Scripting.dictionary")

        oSelection.Clear()



        FixSingleLevel(oCurrentProd)

    End Sub
    Private Sub FixSingleLevel(ByRef oCurrentProd As Object)

        On Error Resume Next


        Dim ItemToFix As Product
        Dim iProdCount As Integer
        Dim i As Integer
        Dim j As Integer
        Dim oConstraints As Constraints
        Dim oReference As Reference
        Dim sItemName As String
        Dim constraint1 As MECMOD.Constraint
        Dim pActivation As KnowledgewareTypeLib.Parameter
        Dim N, m As Integer
        Dim sActivationName As String

        Err.Clear()
        oCurrentProd = oCurrentProd.ReferenceProduct
        iProdCount = oCurrentProd.Products.Count
        oConstraints = oCurrentProd.Connections("CATIAConstraints")

        N = oConstraints.Count
        m = N
        For i = 1 To m
            oConstraints.Remove(N)
            N = N - 1
        Next


        For i = 1 To iProdCount
            Pint += 1
            ItemToFix = oCurrentProd.Products.Item(i)

CreateReference:

            sItemName = ItemToFix.Name

            oReference = oCurrentProd.CreateReferenceFromName(sItemName & "/!" & "/")

            constraint1 = oConstraints.AddMonoEltCst(CatConstraintType.catCstTypeReference, oReference)
            constraint1.ReferenceType = CatConstraintRefType.catCstRefTypeFixInSpace

            oSelection.Add(constraint1)
            oVisProp.SetShow(CatVisPropertyShow.catVisPropertyNoShowAttr)
            oSelection.Clear()

RecursionCall:
            If ItemToFix.Products.Count <> 0 Then
                If oList.exists(ItemToFix.PartNumber) Then GoTo Finish

                If ItemToFix.PartNumber = ItemToFix.ReferenceProduct.Parent.Product.PartNumber Then oList.Add(ItemToFix.PartNumber, 1)
                Call FixSingleLevel(ItemToFix)
            End If
Finish:
        Next

    End Sub

















End Class

Public Module VariablesPublic

    Public CATIA As INFITF.Application

End Module


