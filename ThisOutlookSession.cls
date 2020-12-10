VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisOutlookSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
' // Luodaan MFilesin globaalit objektit
Dim oClientApp As New MFilesAPI.MFilesClientApplication
Dim oVaultConnection As MFilesAPI.VaultConnection
Dim oVault As MFilesAPI.Vault
Dim oFile As New MFilesAPI.SourceObjectFile
Dim oFiles As New MFilesAPI.SourceObjectFiles
Dim oObjectCreationInfo As New MFilesAPI.ObjectCreationInfo

Private WithEvents olItems As Outlook.Items
Attribute olItems.VB_VarHelpID = -1

Private Sub Application_Startup()
' // Määritellään Outlook muuttujat
Dim olApp As Outlook.Application
Dim olNS As Outlook.NameSpace
Set olApp = Outlook.Application
Set olNS = olApp.GetNamespace("MAPI")
Set olItems = olNS.GetDefaultFolder(olFolderInbox).Items
Debug.Print "Application_Startup triggeröitynyt " & Now()
' // Tällä voit zekkaa Classien ID:n: Debug.Print oVault.ClassOperations.GetAllObjectClasses.item(1).Name
' // Tällä saat etsitty itsellesi Property:n ID:n
'    Dim MF_FindPropertyDefID As Integer
'    Dim i As Integer
'    Dim gPropertyDefs
'    Set gPropertyDefs = oVault.PropertyDefOperations.GetPropertyDefs
'    Dim oPropDef As MFilesAPI.PropertyDef
'    For i = 1 To gPropertyDefs.Count
'        Set oPropDef = gPropertyDefs.item(i)
'        If oPropDef.Name = "Tiimi" Then
'            MF_FindPropertyDefID = oPropDef.ID
'        End If
'    Next
'    Debug.Print MF_FindPropertyDefID
' // ===============
End Sub


Private Sub OlItems_ItemAdd(ByVal item As Object)

' // Muodostetaan yhteys M-files rajapintaan
Dim my_olMail As Outlook.MailItem
Dim my_olAtt As Outlook.Attachment
Set oVaultConnection = oClientApp.GetVaultConnection("Masinodok")
Set oVault = oClientApp.BindToVault("Masinodok", 0, True, False)
oObjectCreationInfo.SetObjectType MFilesAPI.MFBuiltInObjectType.MFBuiltInObjectTypeDocument, False


If TypeName(item) = "MailItem" Then
    Set my_olMail = item

        Debug.Print ("Toimiiko tänne?")
        '// Jos liitteitä ja s-postin otsikko "certificate" tai "packing list"
        If my_olMail.Attachments.Count > 0 And (InStr(LCase(my_olMail.Subject), "certifikat") <> 0 Or InStr(LCase(my_olMail.Subject), "certificate") <> 0 Or InStr(LCase(Left$(my_olMail.Subject, 12)), "packing list") <> 0) Then
            
                '// Käydään liitteet läpi
                For Each my_olAtt In my_olMail.Attachments
                    Dim olFileType As String
                    olFileType = LCase$(Right$(my_olAtt.FileName, 3))
                    
                        '// Jos liite on .pdf (Masino Welding Oy:n sertit poikkeuksetta .pdf)
                        Select Case olFileType
                        Case "pdf"
                            oFile.Title = my_olAtt.FileName
                            oFile.Extension = olFileType
                            Debug.Print ("Liite saapunut: " & my_olAtt.FileName)
                            Debug.Print ("oFiles collectioni: " & oFiles.Count)
                           ' Debug.Print ("oPropsVals collectioni: " & oPropVals.Count)
       
                            
                            '// Talleta liite ensin kansioon ja sieltä M-filesiin
                            my_olAtt.SaveAsFile "C:\Users\tommib\Desktop\Sertifikaatit\" & my_olAtt.DisplayName
                            oFile.SourceFilePath = "C:\Users\tommib\Desktop\Sertifikaatit\" & my_olAtt.DisplayName
                    
                            '// Construct parameters for document card.
                            oFiles.Add -1, oFile
                            oObjectCreationInfo.SetSourceFiles oFiles
                            oObjectCreationInfo.SetDisableObjectCreation (False)
                            oObjectCreationInfo.SetSingleFileDocument True, False
                    
                            Dim oPropVals As New MFilesAPI.PropertyValues
                                            
                            ' // Add 'Nimi ja otsikko' property by creating a new PropertyValue object.
                            Dim oPropsVal0 As New MFilesAPI.PropertyValue
                            oPropsVal0.PropertyDef = MFilesAPI.MFBuiltInPropertyDef.MFBuiltInPropertyDefNameOrTitle
                            oPropsVal0.TypedValue.SetValue MFDatatypeText, my_olAtt.FileName
                            oPropVals.Add 0, oPropsVal0
                    
                            ' // Add 'Luokka' property
                            Dim oPropsVal1 As New MFilesAPI.PropertyValue
                            oPropsVal1.PropertyDef = MFilesAPI.MFBuiltInPropertyDef.MFBuiltInPropertyDefClass
                            oPropsVal1.TypedValue.SetValue MFilesAPI.MFDataType.MFDatatypeLookup, 1 '// 1 on "Muut dokumentit" id
                            oPropVals.Add 1, oPropsVal1
                    
                            ' // Add 'Dokumentin päiväys' property
                            Dim oPropsVal2 As New MFilesAPI.PropertyValue
                            oPropsVal2.PropertyDef = 1035 '// "Dokumentin päiväys" property-ID: 1035
                            oPropsVal2.TypedValue.SetValue MFDataType.MFDatatypeDate, Now()
                            oPropVals.Add 2, oPropsVal2
                    
                            ' // Add document 'vastuuhenkilö' property ("Vastuuhenkilö" property-ID: 1116)
                            Dim oPropsVal3 As New MFilesAPI.PropertyValue
                            oPropsVal3.PropertyDef = 1116
                            oPropsVal3.TypedValue.SetValue MFDataType.MFDatatypeMultiSelectLookup, 197
                            oPropVals.Add 3, oPropsVal3
                    
                            ' // Add document 'Tiimi' 'property id: 1114
                            Dim oPropsVal4 As New MFilesAPI.PropertyValue
                            oPropsVal4.PropertyDef = 1114
                            oPropsVal4.TypedValue.SetValue MFDataType.MFDatatypeMultiSelectLookup, 5
                            oPropVals.Add 4, oPropsVal4
                    
                            ' // Add document 'Asiakokonaisuus' 'property id:
                            Dim oPropsVal5 As New MFilesAPI.PropertyValue
                            oPropsVal5.PropertyDef = 1025
                            oPropsVal5.TypedValue.SetValue MFDataType.MFDatatypeMultiSelectLookup, 11
                            oPropVals.Add 5, oPropsVal5
                    
                            Dim oObjectVersionAndProperties As MFilesAPI.ObjectVersionAndProperties
                            Set oObjectVersionAndProperties = oVault.ObjectOperations.CreateNewObjectEx(MFilesAPI.MFBuiltInObjectType.MFBuiltInObjectTypeDocument, oPropVals, oFiles, True, True)
                            oFiles.Remove 1
                            Set oPropVals = Nothing
                                                        
                        End Select
                Next
        End If
    Set my_olMail = Nothing
End If

End Sub
