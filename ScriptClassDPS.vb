Imports System
Imports SolidWorks.Interop.sldworks
Imports SolidWorks.Interop.swconst
Imports JobBox.Interfaces

Public Class ScriptClassDps

    Public Sub CheckAddin(ByVal context As IRunContext, ByVal options As IRunOptions)

        Dim swApp As SldWorks = context.Variable.Get("%SOLIDWORKS%")
        Dim addin As String = context.Variable.Get("%AddIn%")
        Dim isactive As Boolean = IsAddinActive(addin, swApp)

        context.Variable.Set("%AddInIsActive%", isactive)

    End Sub

    Public Sub CloseProgramm(variables As JobBox.Interfaces.IVariable)
        System.Environment.Exit(0)
    End Sub

    Public Sub WriteProperty(variables As JobBox.Interfaces.IVariable)

        Dim swApp As SldWorks = variables.Get("%SOLIDWORKS%")

        Dim file As String = variables.Resolve("%WriteProperty_File%")
        Dim name As String = variables.Resolve("%WriteProperty_Name%")
        Dim config As String = variables.Resolve("%WriteProperty_Config%")
        Dim value As String = variables.Resolve("%WriteProperty_Value%")


        Dim swModelActiveBefore As ModelDoc2 = swApp.ActiveDoc

        Dim documentType As swDocumentTypes_e
        Select Case IO.Path.GetExtension(file).ToUpper()
            Case ".SLDPRT" : documentType = swDocumentTypes_e.swDocPART
            Case ".SLDASM" : documentType = swDocumentTypes_e.swDocASSEMBLY
            Case ".SLDDRW" : documentType = swDocumentTypes_e.swDocDRAWING
            Case ".SLDLFP" : documentType = swDocumentTypes_e.swDocPART
            Case Else
        End Select

        Dim docSpec As DocumentSpecification = swApp.GetOpenDocSpec(file)
        docSpec.Silent = True

        Dim swModel As ModelDoc2 = swApp.OpenDoc7(docSpec)
        Dim swCustomPropMgr As CustomPropertyManager = Nothing
        If Not swModel Is Nothing Then
            Select Case documentType
                Case swDocumentTypes_e.swDocPART, swDocumentTypes_e.swDocASSEMBLY
                    swCustomPropMgr = swModel.Extension.CustomPropertyManager(config)
                Case swDocumentTypes_e.swDocDRAWING
                    swCustomPropMgr = swModel.Extension.CustomPropertyManager("")
            End Select
            If swCustomPropMgr Is Nothing Then Throw New ApplicationException(String.Format("Die angegebene Konfiguration {0} existiert nicht.", config))
        Else
            Throw New ApplicationException("Dokument '" + file + "' nicht gefunden.")
        End If

        If TryCast(swCustomPropMgr.GetNames(), String()) IsNot Nothing AndAlso TryCast(swCustomPropMgr.GetNames(), String()).Contains(name, StringComparer.OrdinalIgnoreCase) Then
            swCustomPropMgr.Set(name, value)
        Else
            swCustomPropMgr.Add2(name, swCustomInfoType_e.swCustomInfoText, value)
        End If

        Dim swModelActiveAfter As ModelDoc2 = swApp.ActiveDoc

        If swModelActiveBefore IsNot Nothing AndAlso Not swModelActiveBefore.GetPathName.Equals(swModelActiveAfter.GetPathName, StringComparison.OrdinalIgnoreCase) Then
            Dim iErr As Integer
            swApp.ActivateDoc3(swModelActiveBefore.GetPathName, False, swRebuildOnActivation_e.swUserDecision, iErr)
        End If
    End Sub

    Public Sub WriteProperties(variables As JobBox.Interfaces.IVariable)
        Const cDataTable As String = "%WriteProperties_Ergebnistabelle%"
        Dim swApp As SldWorks = variables.Get("%SOLIDWORKS%")

        Dim dt As System.Data.DataTable = TryCast(variables.Get(cDataTable), System.Data.DataTable)
        If dt Is Nothing Then
            Throw New ApplicationException("Keine Ergebnistabelle in der Variable " + cDataTable)
        End If

        Dim swModelActiveBefore As ModelDoc2 = swApp.ActiveDoc

        Dim lstModelWithFingerprint As New List(Of String)
        For Each row As System.Data.DataRow In dt.Rows
            Dim file As String = row("File").ToString

            Dim documentType As swDocumentTypes_e
            Select Case IO.Path.GetExtension(file).ToUpper()
                Case ".SLDPRT" : documentType = swDocumentTypes_e.swDocPART
                Case ".SLDASM" : documentType = swDocumentTypes_e.swDocASSEMBLY
                Case ".SLDDRW" : documentType = swDocumentTypes_e.swDocDRAWING
                Case ".SLDLFP" : documentType = swDocumentTypes_e.swDocPART
                Case Else
            End Select


            Dim docSpec As DocumentSpecification = swApp.GetOpenDocSpec(file)
            docSpec.Silent = True

            Dim swModel As ModelDoc2 = swApp.OpenDoc7(docSpec)
            Dim swCustomPropMgr As CustomPropertyManager = Nothing
            If swModel Is Nothing Then
                Throw New ApplicationException("Dokument '" + file + "' nicht gefunden.")
            End If

            Dim numbers As List(Of String) = (From c As System.Data.DataColumn In dt.Columns Where c.ColumnName.StartsWith("Name") Select nr = c.ColumnName.Substring(4)).ToList
            Dim activeConfig As String = "~"
            For Each nr As String In numbers
                Dim name As String = row("Name" + nr).ToString
                If Not String.IsNullOrEmpty(name) Then
                    Dim config As String = row("Config" + nr).ToString
                    Dim value As String = row("Value" + nr).ToString
                    If Not config.Equals(activeConfig) Then
                        activeConfig = config
                        swCustomPropMgr = Nothing
                        Select Case documentType
                            Case swDocumentTypes_e.swDocPART, swDocumentTypes_e.swDocASSEMBLY
                                swCustomPropMgr = swModel.Extension.CustomPropertyManager(config)
                            Case swDocumentTypes_e.swDocDRAWING
                                swCustomPropMgr = swModel.Extension.CustomPropertyManager("")
                        End Select
                        If swCustomPropMgr Is Nothing Then Throw New ApplicationException(String.Format("Die angegebene Konfiguration {0} existiert nicht.", config))
                    End If

                    If TryCast(swCustomPropMgr.GetNames(), String()) IsNot Nothing AndAlso TryCast(swCustomPropMgr.GetNames(), String()).Contains(name, StringComparer.OrdinalIgnoreCase) Then
                        swCustomPropMgr.Set(name, value)
                    Else
                        swCustomPropMgr.Add2(name, swCustomInfoType_e.swCustomInfoText, value)
                    End If
                End If
            Next

            'Dokument Speichern
            Dim errors As Integer
            Dim warnings As Integer
            swModel.Save3(swSaveAsOptions_e.swSaveAsOptions_Silent + swSaveAsOptions_e.swSaveAsOptions_AvoidRebuildOnSave, errors, warnings)
        Next

        Dim swModelActiveAfter As ModelDoc2 = swApp.ActiveDoc

        If swModelActiveBefore IsNot Nothing AndAlso Not swModelActiveBefore.GetPathName.Equals(swModelActiveAfter.GetPathName, StringComparison.OrdinalIgnoreCase) Then
            Dim iErr As Integer
            swApp.ActivateDoc3(swModelActiveBefore.GetPathName, False, swRebuildOnActivation_e.swUserDecision, iErr)
        End If
    End Sub

    Public Sub ReplaceText(variables As JobBox.Interfaces.IVariable)
        Const cDataTable As String = "%ReplaceText_Ergebnistabelle%"

        Dim dt As System.Data.DataTable = TryCast(variables.Get(cDataTable), System.Data.DataTable)
        If dt Is Nothing Then
            Throw New ApplicationException("Keine Ergebnistabelle in der Variable " + cDataTable)
        End If


        For Each row As System.Data.DataRow In dt.Rows
            Dim file As String = row("File").ToString

            'Datei vorhanden?
            If Not IO.File.Exists(file) Then
                Throw New ApplicationException("Datei '" + file + "' nicht gefunden.")
            End If

            'Dateiinhalt laden
            Dim text As String = IO.File.ReadAllText(file, System.Text.Encoding.Default)

            'Ersetzungen durchgehen
            Dim numbers As List(Of String) = (From c As System.Data.DataColumn In dt.Columns Where c.ColumnName.StartsWith("Find") Select nr = c.ColumnName.Substring(4)).ToList
            For Each nr As String In numbers
                Dim textFind As String = row("Find" + nr).ToString
                Dim textReplace As String = row("Replace" + nr).ToString
                If Not String.IsNullOrEmpty(textFind) Then
                    text = Replace(text, textFind, textReplace,,, CompareMethod.Binary)
                End If
            Next

            'Datei speichern
            IO.File.WriteAllText(file, text, System.Text.Encoding.Default)
        Next

    End Sub

    Public Sub FindText(variables As JobBox.Interfaces.IVariable)
        Dim file As String = variables.Get("%DateiListe_DateiPfad%")
        Dim find As String = variables.Get("%Find%")
        Dim text As String = IO.File.ReadAllText(file, System.Text.Encoding.Default)
        'Datei vorhanden?
        If Not IO.File.Exists(file) Then
            Throw New ApplicationException("Datei '" + file + "' nicht gefunden.")
        End If

        variables.Set("%Exists%", IIf(text.Contains(find), "True", "False"))
    End Sub

    Public Sub FindTextValueAsList(variables As JobBox.Interfaces.IVariable)
        Dim file As String = variables.Get("%DateiListe_DateiPfad%")
        Dim find As String = variables.Get("%Find%")
        Dim text As String = IO.File.ReadAllText(file, System.Text.Encoding.Default)
        'Datei vorhanden?
        If Not IO.File.Exists(file) Then
            Throw New ApplicationException("Datei '" + file + "' nicht gefunden.")
        End If

        Dim lst As New List(Of String)

        For Each row As String In Split(text, vbNewLine)
            If row.Trim.StartsWith(find) Then
                'Zeile gefunden
                Dim i As Integer = InStr(row, "=")
                If i > 0 Then
                    lst.Add(row.Substring(i).Trim)
                End If
            End If
        Next

        variables.Set("%List%", lst)
    End Sub

    Public Sub MoveToEnd(variables As JobBox.Interfaces.IVariable)
        Dim file As String = variables.Get("%DateiListe_DateiPfad%")
        Dim textStart As String = variables.Get("%TextStart%") '"<Anfang>"
        Dim textEnd As String = variables.Get("%TextEnde%") '"<Ende>"
        Dim textFileEndChar As String = variables.Get("%TextEndZeichen%") '"!"
        Dim text As String = IO.File.ReadAllText(file, System.Text.Encoding.Default)

        'Datei vorhanden?
        If Not IO.File.Exists(file) Then
            Throw New ApplicationException("Datei '" + file + "' nicht gefunden.")
        End If

        Dim newText As String = String.Empty
        Dim found As Integer = 0
        Dim lst As New List(Of String)
        For Each line As String In Split(text, System.Environment.NewLine)
            Select Case line
                Case textStart
                    found = 1
                Case textEnd
                    If found = 1 Then found = 2
                Case textFileEndChar
                    For Each l As String In lst
                        newText += l + System.Environment.NewLine
                    Next
                    lst.Clear()
            End Select

            If found > 0 Then
                lst.Add(line)
                If found = 2 Then found = 0
            Else
                newText += line + System.Environment.NewLine
            End If
        Next

        If newText.EndsWith(System.Environment.NewLine + System.Environment.NewLine) Then
            newText = Left(newText, Len(newText) - Len(System.Environment.NewLine))
        End If

        IO.File.WriteAllText(file, newText, System.Text.Encoding.Default)
    End Sub

    Public Sub FingerprintAssembly(variables As JobBox.Interfaces.IVariable)
        Dim swApp As SldWorks = variables.Get("%SOLIDWORKS%")


        Dim commandInProgress As Boolean = swApp.CommandInProgress
        swApp.CommandInProgress = True
        Try

            Dim swAssy As AssemblyDoc = swApp.ActiveDoc
            Dim comparr As Object() = swAssy.GetComponents(False)

            If comparr IsNot Nothing Then
                Dim lst As New List(Of String)
                For Each swComp As Component2 In comparr
                    Dim swPart As PartDoc = TryCast(swComp.GetModelDoc2, PartDoc)
                    If swPart IsNot Nothing Then
                        Dim swModel As ModelDoc2 = swPart
                        If Not swModel.GetPathName.Contains("\Beschlaege\") AndAlso swModel.Extension.ToolboxPartType = swToolBoxPartType_e.swNotAToolboxPart Then
                            Dim config As String = TryCast(swModel.GetActiveConfiguration, Configuration).Name
                            Dim id As String = swModel.GetPathName + "_" + config
                            If Not lst.Contains(id) Then
                                lst.Add(id)
                                Dim fp As New DPS.FrameWork.SolidWorks.FingerprintCalculator(swApp, swPart)
                                Dim s As String = fp.Fingerprint
                                Dim swCustomPropMgr As CustomPropertyManager = swModel.Extension.CustomPropertyManager(config)
                                Dim name As String = "Fingerprint"
                                If TryCast(swCustomPropMgr.GetNames(), String()) IsNot Nothing AndAlso TryCast(swCustomPropMgr.GetNames(), String()).Contains(name, StringComparer.OrdinalIgnoreCase) Then
                                    swCustomPropMgr.Set(name, s)
                                Else
                                    swCustomPropMgr.Add2(name, swCustomInfoType_e.swCustomInfoText, s)
                                End If
                            End If
                        End If
                    End If

                Next
            End If

        Catch ex As Exception
            Throw ex
        Finally
            swApp.CommandInProgress = commandInProgress
        End Try

    End Sub


    ''' <summary>
    ''' Gets the addin is active matching the passed GUID
    ''' </summary>
    ''' <param name="addinGuidOrProgId">Addin GUID in curly brackets -> {GUID} or ProgId</param>
    ''' <param name="swApp">Active SOLIDWORKS Object</param>
    ''' <returns>Addin active (True/False)</returns>
    Private Function IsAddinActive(addinGuidOrProgId As String, ByVal swApp As Object) As Boolean
        Try
            Dim addinGuid As String = GetGuidFromProgId(addinGuidOrProgId)
            Dim addin As Object = Nothing
            Try
                If addinGuid IsNot Nothing Then
                    addin = swApp.GetAddInObject(addinGuid)
                End If
            Catch ex As Exception
            End Try
            Return (addin IsNot Nothing)
        Catch
        End Try

        Return False
    End Function

    Private Function GetGuidFromProgId(addinGuidOrProgId As String) As String

        'temporäre Lösung nur für GUIDs
        Return addinGuidOrProgId

        'If addinGuidOrProgId.StartsWith("{") AndAlso addinGuidOrProgId.EndsWith("}") Then
        '    Return addinGuidOrProgId
        'Else
        '    Try
        '        Dim regBasePath As String = $"{addinGuidOrProgId}\CLSID"
        '        Dim hkcr As Microsoft.Win32.RegistryKey = Microsoft.Win32.RegistryKey.OpenBaseKey(Microsoft.Win32.RegistryHive.ClassesRoot, Microsoft.Win32.RegistryView.Registry64)
        '        Dim rk As Microsoft.Win32.RegistryKey = hkcr.OpenSubKey(regBasePath, False)
        '        If rk IsNot Nothing Then
        '            Return rk.GetValue("")
        '            'Else
        '            '    Dim clsid As Microsoft.Win32.RegistryKey = hkcr.OpenSubKey("CLSID", False)
        '            '    For Each id As String In clsid.GetSubKeyNames
        '            '        If addinGuidOrProgId.StartsWith("{") AndAlso addinGuidOrProgId.EndsWith("}") Then

        '            '        End If
        '            '    Next
        '        End If
        '    Catch ex As Exception

        '    End Try
        'End If
        Return String.Empty
    End Function

End Class
