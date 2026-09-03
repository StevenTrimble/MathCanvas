Imports System.Globalization
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Text
Imports Microsoft.Web.WebView2.Core
Imports netDxf
Imports netDxf.Entities
Imports netDxf.Header
Imports netDxf.Tables
Imports netDxf.Units
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports Excel = Microsoft.Office.Interop.Excel

Public Class frmMath

    ' =========================================================
    ' EXCEL INTEGRATION OPTIONS
    ' =========================================================
    ' (1) Work in background (Default Ticked - ON)
    Public Property ExcelWorkInBackground As Boolean = True

    ' (2) Auto save and close Excel when closing file (Default Ticked - ON)
    Public Property AutoSaveAndCloseExcel As Boolean = True
    ' =========================================================

    Private _webViewReady As Boolean
    Private _presentationFullscreen As Boolean
    Private _windowStateBeforePresentation As FormWindowState
    Private _formBorderStyleBeforePresentation As FormBorderStyle
    Private _boundsBeforePresentation As Rectangle
    Private _topMostBeforePresentation As Boolean

    Private _excelApp As Excel.Application
    Private _linkedTables As New Dictionary(Of String, String)() ' Canvas Block ID -> Excel Range Address

    <DllImport("user32.dll")>
    Private Shared Function SetForegroundWindow(hWnd As IntPtr) As Boolean
    End Function
    Private Async Sub frmMath_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            Await wbMath.EnsureCoreWebView2Async(Nothing)
            AddHandler wbMath.CoreWebView2.WebMessageReceived, AddressOf OnWebMessageReceived
            wbMath.CoreWebView2.Settings.AreDefaultContextMenusEnabled = False

            Dim htmlPath = FindMathCanvasPath()
            If File.Exists(htmlPath) Then
                wbMath.CoreWebView2.Navigate(New Uri(htmlPath).AbsoluteUri)
            Else
                wbMath.NavigateToString("<html><body><h2>MathCanvas.html was not deployed.</h2></body></html>")
            End If
        Catch ex As Exception
            MessageBox.Show("The math canvas could not be initialized." & Environment.NewLine & ex.Message,
                            "Math Canvas", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Shared Function FindMathCanvasPath() As String
        Dim folder = New DirectoryInfo(Application.StartupPath)
        For i As Integer = 0 To 6
            If folder Is Nothing Then Exit For

            Dim candidate = Path.Combine(folder.FullName, "Math", "MathCanvas.html")
            If File.Exists(candidate) Then
                CheckCssDeployed(Path.GetDirectoryName(candidate))
                Return candidate
            End If

            candidate = Path.Combine(folder.FullName, "MathCanvas.html")
            If File.Exists(candidate) Then
                CheckCssDeployed(Path.GetDirectoryName(candidate))
                Return candidate
            End If

            folder = folder.Parent
        Next
        Return Path.Combine(Application.StartupPath, "Math", "MathCanvas.html")
    End Function

    Private Shared Sub CheckCssDeployed(htmlDir As String)
        Dim cssPath = Path.Combine(htmlDir, "MathCanvas.css")
        If Not File.Exists(cssPath) Then
            System.Diagnostics.Debug.WriteLine("WARNING: MathCanvas.css was not found in " & htmlDir)
        End If
    End Sub

    Private Sub OnWebMessageReceived(sender As Object, args As CoreWebView2WebMessageReceivedEventArgs)
        Try
            Dim message = JObject.Parse(args.TryGetWebMessageAsString())
            Select Case message.Value(Of String)("type")
                Case "ready"
                    _webViewReady = True
                Case "open"
                    LoadCanvas()
                Case "save"
                    Dim doc = message("document")
                    Dim isSaveAs As Boolean = If(message("isSaveAs") IsNot Nothing, message.Value(Of Boolean)("isSaveAs"), False)
                    Dim currentPath As String = If(message("currentPath") IsNot Nothing, message("currentPath").ToString(), "")
                    SaveCanvas(doc, currentPath, isSaveAs)
                Case "exportDxf"
                    SaveDxf(message.Value(Of String)("fileName"), message("drawing"))
                Case "presentationFullscreen"
                    SetPresentationFullscreen(message.Value(Of Boolean)("enabled"))

               ' --- EXCEL INTEGRATION CASES ---
                Case "excelOpen"
                    HandleExcelOpen(message.Value(Of String)("path"))
                Case "excelLink"
                    HandleExcelLink(message.Value(Of Integer)("id"), "A")
                Case "excelLinkB"
                    HandleExcelLink(message.Value(Of Integer)("id"), "B")

                ' FIX: Safely parse the path even if it's missing from the JSON payload
                Case "excelUpdate"
                    Dim pathStr As String = If(message("path") IsNot Nothing, message("path").ToString(), "")
                    HandleExcelUpdate(message.Value(Of Integer)("id"), "A", message.Value(Of String)("range"), pathStr, message.Value(Of String)("data"))
                Case "excelUpdateB"
                    Dim pathStr As String = If(message("path") IsNot Nothing, message("path").ToString(), "")
                    HandleExcelUpdate(message.Value(Of Integer)("id"), "B", message.Value(Of String)("range"), pathStr, message.Value(Of String)("data"))

                Case "excelUnlink"
                    HandleExcelUnlink(message.Value(Of Integer)("id"), "A")
                Case "excelUnlinkB"
                    HandleExcelUnlink(message.Value(Of Integer)("id"), "B")
                Case "excelLinkResult"
                    HandleExcelLinkResult(message.Value(Of Integer)("id"))

                ' FIX: Safely parse the path for results too
                Case "excelUpdateResult"
                    Dim pathStr As String = If(message("path") IsNot Nothing, message("path").ToString(), "")
                    HandleExcelUpdateResult(message.Value(Of Integer)("id"), message.Value(Of String)("range"), pathStr, message.Value(Of String)("data"))

                Case "excelUnlinkResult"
                    HandleExcelUnlinkResult(message.Value(Of Integer)("id"))
                Case "excelRelink"
                    Dim rangeStr As String = If(message("range") IsNot Nothing, message("range").ToString(), "")
                    HandleExcelRelink(message.Value(Of Integer)("id"), message.Value(Of String)("slot"), rangeStr)
                Case "excelForceRead"
                    Dim pathStr As String = If(message("path") IsNot Nothing, message("path").ToString(), "")
                    HandleExcelForceRead(message.Value(Of Integer)("id"), message.Value(Of String)("slot"), message.Value(Of String)("range"), pathStr)

                    ' --- TRIGGERS WHEN A FILE TAB IS CLOSED ---
                'Case "fileClosed"
                '    Dim pathsArray = TryCast(message("excelPaths"), JArray)
                '    HandleFileClosed(pathsArray)

                    ' --- PROJECT MANAGER CASES ---
                Case "selectFolder"
                    HandleSelectFolder()
                Case "openFolder"
                    Dim pathStr As String = If(message("path") IsNot Nothing, message("path").ToString(), "")
                    HandleOpenFolder(pathStr)
                    ' --- NEW PROJECT MANAGER CASES ---
                Case "openProjectFile"
                    Dim filePath As String = If(message("path") IsNot Nothing, message("path").ToString(), "")
                    If File.Exists(filePath) Then
                        Dim json = File.ReadAllText(filePath, Encoding.UTF8)
                        Dim jDoc = JToken.Parse(json)

                        jDoc("title") = Path.GetFileName(filePath)
                        jDoc("filePath") = filePath ' <--- ADD THIS LINE

                        Dim script = "window.loadDocument(" & jDoc.ToString(Formatting.None) & ");"
                        wbMath.CoreWebView2.ExecuteScriptAsync(script)
                    Else
                        MessageBox.Show("File not found: " & filePath, "Open File", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    End If

                Case "renameProjectFile"
                    Dim oldPath As String = If(message("path") IsNot Nothing, message("path").ToString(), "")
                    Dim newName As String = If(message("newName") IsNot Nothing, message("newName").ToString(), "")
                    HandleRenameProjectFile(oldPath, newName)

                Case "duplicateProjectFile"
                    Dim srcPath As String = If(message("path") IsNot Nothing, message("path").ToString(), "")
                    HandleDuplicateProjectFile(srcPath)

                Case "refreshFolders"
                    Dim pathsArray = TryCast(message("paths"), JArray)
                    If pathsArray IsNot Nothing Then
                        For Each pathToken In pathsArray
                            Dim folderPath = pathToken.ToString()
                            If Directory.Exists(folderPath) Then
                                Dim folderName = Path.GetFileName(folderPath)
                                If String.IsNullOrEmpty(folderName) Then folderName = folderPath

                                Dim files As New JArray()
                                Dim dirInfo As New DirectoryInfo(folderPath)
                                Dim allowedExtensions = {".mc.json", ".mathcanvas.json", ".json", ".mc"}

                                For Each fileInfo In dirInfo.GetFiles()
                                    If allowedExtensions.Contains(fileInfo.Extension.ToLowerInvariant()) Then
                                        Dim fileObj As New JObject()
                                        fileObj("id") = Guid.NewGuid().ToString("N")
                                        fileObj("name") = fileInfo.Name
                                        fileObj("path") = fileInfo.FullName
                                        files.Add(fileObj)

                                        ' Publish the document contents so the canvas can expose its named
                                        ' values as "File.variable" globals in autocomplete and calculations.
                                        Try
                                            Dim contents = File.ReadAllText(fileInfo.FullName, Encoding.UTF8)
                                            Dim jsFilePath = JsonConvert.SerializeObject(fileInfo.FullName)
                                            Dim jsFileName = JsonConvert.SerializeObject(fileInfo.Name)
                                            Dim jsContents = JsonConvert.SerializeObject(contents)
                                            wbMath.CoreWebView2.ExecuteScriptAsync($"if(typeof window.setProjectFileDocument === 'function') window.setProjectFileDocument({jsFilePath}, {jsFileName}, {jsContents});")
                                        Catch
                                            ' A file that cannot be read is simply skipped as a variable source.
                                        End Try
                                    End If
                                Next

                                Dim jsFolder = JsonConvert.SerializeObject(folderName)
                                Dim jsPath = JsonConvert.SerializeObject(folderPath)
                                Dim jsFiles = JsonConvert.SerializeObject(files.ToString(Formatting.None))
                                wbMath.CoreWebView2.ExecuteScriptAsync($"if(typeof window.addProjectFolderFromHost === 'function') window.addProjectFolderFromHost({jsFolder}, {jsPath}, {jsFiles});")
                            End If
                        Next
                        wbMath.CoreWebView2.ExecuteScriptAsync("if(typeof window.projectFilesLoaded === 'function') window.projectFilesLoaded();")
                    End If

                    ' --- TEKLA STRUCTURAL DESIGNER MODEL BRIDGE ---
                Case "tsdConnect"
                    HandleTsdConnect()
                Case "tsdFetchModel"
                    Dim includeRebar As Boolean = If(message("includeRebar") IsNot Nothing, message.Value(Of Boolean)("includeRebar"), False)
                    HandleTsdFetchModel(includeRebar)
                Case "tsdSelectInstance"
                    ' User picked a TSD instance from the multi-instance chooser.
                    TsdInstanceIndex = If(message("index") IsNot Nothing, message.Value(Of Integer)("index"), 0)
                    _tsdInstanceChosen = True
                    _tsdModel = Nothing
                    _tsdDocument = Nothing
                    _tsdApp = Nothing
                    HandleTsdFetchModel(_tsdPendingIncludeRebar)
                Case "tsdFetchResult"
                    Dim resKey As String = If(message("key") IsNot Nothing, message("key").ToString(), "")
                    Dim resName As String = If(message("name") IsNot Nothing, message("name").ToString(), "")
                    Dim resType As String = If(message("resultType") IsNot Nothing, message("resultType").ToString(), "")
                    Dim resSpan As Integer = If(message("span") IsNot Nothing, message.Value(Of Integer)("span"), 0)
                    Dim resPos As String = If(message("position") IsNot Nothing, message("position").ToString(), "MAX")
                    HandleTsdFetchResult(resKey, resName, resType, resSpan, resPos)

            End Select
        Catch ex As Exception
            MessageBox.Show("The math canvas message was invalid." & Environment.NewLine & ex.Message,
                            "Math Canvas", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try
    End Sub


    ' =========================================================
    ' FILE CLOSE & APP CLOSE LOGIC (NEW Auto Save/Close Routine)
    ' =========================================================
    Private Sub HandleFileClosed(pathsArray As JArray)
        If _excelApp Is Nothing OrElse pathsArray Is Nothing Then Return

        Try
            If AutoSaveAndCloseExcel Then
                Dim wbs As Excel.Workbooks = ComRetry(Function() _excelApp.Workbooks)
                For Each pathToken In pathsArray
                    Dim filePath = pathToken.ToString()
                    If Not String.IsNullOrWhiteSpace(filePath) Then
                        Dim count As Integer = ComRetry(Function() wbs.Count)
                        For i As Integer = count To 1 Step -1
                            Dim wb As Excel.Workbook = ComRetry(Function() wbs.Item(i))
                            If wb IsNot Nothing AndAlso ComRetry(Function() wb.FullName).Equals(filePath, StringComparison.OrdinalIgnoreCase) Then
                                ComRetryAction(Sub() wb.Close(SaveChanges:=True))
                                Marshal.ReleaseComObject(wb)
                                Exit For
                            End If
                        Next
                    End If
                Next

                If ComRetry(Function() wbs.Count) = 0 Then
                    ComRetryAction(Sub() _excelApp.Quit())
                    Marshal.ReleaseComObject(_excelApp)
                    _excelApp = Nothing
                End If
                Marshal.ReleaseComObject(wbs)
            Else
                ComRetryAction(Sub() _excelApp.Visible = True)
            End If
        Catch ex As Exception
        End Try
    End Sub

    ' =========================================================
    ' PROJECT MANAGER LOGIC
    ' =========================================================
    Private Sub HandleSelectFolder()
        Using dialog As New FolderBrowserDialog()
            dialog.Description = "Select a Project Folder"
            dialog.ShowNewFolderButton = True

            If dialog.ShowDialog(Me) = DialogResult.OK Then
                Dim folderPath As String = dialog.SelectedPath
                Dim folderName As String = Path.GetFileName(folderPath)
                If String.IsNullOrEmpty(folderName) Then folderName = folderPath ' Fallback for root drives (e.g., "C:\")

                ' Scan for canvas files inside the selected folder
                Dim files As New JArray()
                Try
                    Dim dirInfo As New DirectoryInfo(folderPath)
                    ' Look for standard math canvas files
                    Dim allowedExtensions = {".mc.json", ".mathcanvas.json", ".json", ".mc"}

                    For Each fileInfo In dirInfo.GetFiles()
                        If allowedExtensions.Contains(fileInfo.Extension.ToLowerInvariant()) Then
                            Dim fileObj As New JObject()
                            fileObj("id") = Guid.NewGuid().ToString("N")
                            fileObj("name") = fileInfo.Name
                            fileObj("path") = fileInfo.FullName
                            files.Add(fileObj)
                        End If
                    Next
                Catch ex As Exception
                    ' Ignore folder access exceptions
                End Try

                ' Safely serialize strings to prevent injection/syntax errors in JavaScript
                Dim jsFolder As String = JsonConvert.SerializeObject(folderName)
                Dim jsPath As String = JsonConvert.SerializeObject(folderPath)
                Dim jsFiles As String = JsonConvert.SerializeObject(files.ToString(Formatting.None))

                Dim script As String = $"if(typeof window.addProjectFolderFromHost === 'function') window.addProjectFolderFromHost({jsFolder}, {jsPath}, {jsFiles});"
                wbMath.CoreWebView2.ExecuteScriptAsync(script)
            Else
                ' If canceled, clear the "Waiting..." status in JS
                wbMath.CoreWebView2.ExecuteScriptAsync("document.getElementById('status').textContent = 'Folder selection canceled.';")
            End If
        End Using
    End Sub

    Private Sub HandleOpenFolder(folderPath As String)
        Try
            If Not String.IsNullOrWhiteSpace(folderPath) AndAlso Directory.Exists(folderPath) Then
                System.Diagnostics.Process.Start("explorer.exe", folderPath)
            End If
        Catch ex As Exception
            MessageBox.Show("Could not open the folder: " & ex.Message, "Open Folder", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try
    End Sub

    ' Renames a project file on disk and tells the canvas so that the file tab,
    ' the cached document and every "File.variable" reference are updated.
    Private Sub HandleRenameProjectFile(oldPath As String, newName As String)
        Try
            If String.IsNullOrWhiteSpace(oldPath) OrElse String.IsNullOrWhiteSpace(newName) Then Return
            If Not File.Exists(oldPath) Then
                MessageBox.Show("File not found: " & oldPath, "Rename File", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            If newName.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0 Then
                MessageBox.Show("The name contains characters that are not allowed in a file name.", "Rename File", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            Dim oldName = Path.GetFileName(oldPath)
            Dim newPath = Path.Combine(Path.GetDirectoryName(oldPath), newName)
            If String.Equals(oldPath, newPath, StringComparison.OrdinalIgnoreCase) Then Return

            If File.Exists(newPath) Then
                MessageBox.Show("A file named """ & newName & """ already exists in that folder.", "Rename File", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            File.Move(oldPath, newPath)

            Dim jsOldPath = JsonConvert.SerializeObject(oldPath)
            Dim jsNewPath = JsonConvert.SerializeObject(newPath)
            Dim jsOldName = JsonConvert.SerializeObject(oldName)
            Dim jsNewName = JsonConvert.SerializeObject(newName)
            wbMath.CoreWebView2.ExecuteScriptAsync($"if(typeof window.projectFileRenamed === 'function') window.projectFileRenamed({jsOldPath}, {jsNewPath}, {jsOldName}, {jsNewName});")
        Catch ex As Exception
            MessageBox.Show("Could not rename the file: " & ex.Message, "Rename File", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try
    End Sub

    ' Copies a project file on disk with a unique "(Copy)" name and tells the canvas
    ' so the duplicate is listed and opened under its own file name.
    Private Sub HandleDuplicateProjectFile(sourcePath As String)
        Try
            If String.IsNullOrWhiteSpace(sourcePath) OrElse Not File.Exists(sourcePath) Then
                MessageBox.Show("File not found: " & sourcePath, "Duplicate File", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            Dim folder = Path.GetDirectoryName(sourcePath)
            Dim fileName = Path.GetFileName(sourcePath)
            Dim knownExtensions = {".mc.json", ".mathcanvas.json"}
            Dim ext As String = Path.GetExtension(fileName)
            For Each known In knownExtensions
                If fileName.EndsWith(known, StringComparison.OrdinalIgnoreCase) Then
                    ext = fileName.Substring(fileName.Length - known.Length)
                    Exit For
                End If
            Next
            Dim baseName = fileName.Substring(0, fileName.Length - ext.Length)

            Dim newName = baseName & " (Copy)" & ext
            Dim newPath = Path.Combine(folder, newName)
            Dim counter As Integer = 2
            While File.Exists(newPath)
                newName = baseName & " (Copy " & counter & ")" & ext
                newPath = Path.Combine(folder, newName)
                counter += 1
            End While

            File.Copy(sourcePath, newPath)

            Dim jsSource = JsonConvert.SerializeObject(sourcePath)
            Dim jsNewPath = JsonConvert.SerializeObject(newPath)
            Dim jsNewName = JsonConvert.SerializeObject(newName)

            ' Publish the copy's contents so its named values immediately become
            ' "File.variable" globals, exactly as refreshFolders does for existing files.
            Try
                Dim contents = File.ReadAllText(newPath, Encoding.UTF8)
                Dim jsContents = JsonConvert.SerializeObject(contents)
                wbMath.CoreWebView2.ExecuteScriptAsync($"if(typeof window.setProjectFileDocument === 'function') window.setProjectFileDocument({jsNewPath}, {jsNewName}, {jsContents});")
            Catch
            End Try

            wbMath.CoreWebView2.ExecuteScriptAsync($"if(typeof window.projectFileDuplicated === 'function') window.projectFileDuplicated({jsSource}, {jsNewPath}, {jsNewName});")
        Catch ex As Exception
            MessageBox.Show("Could not duplicate the file: " & ex.Message, "Duplicate File", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try
    End Sub

    ' =========================================================
    ' EXCEL BIDIRECTIONAL INTEGRATION LOGIC
    ' =========================================================

    ' COM Retry Helper: Waits if Excel is busy instead of crashing/skipping
    Private Function ComRetry(Of T)(action As Func(Of T)) As T
        Dim retries As Integer = 20
        While True
            Try
                Return action()
            Catch ex As COMException When ex.ErrorCode = &H800AC472 OrElse ex.ErrorCode = &H80010001 OrElse ex.ErrorCode = &H8001010A
                If retries <= 0 Then Throw
                retries -= 1
                System.Threading.Thread.Sleep(150)
                Application.DoEvents()
            End Try
        End While
        Return Nothing
    End Function

    Private Sub ComRetryAction(action As Action)
        Dim retries As Integer = 20
        While True
            Try
                action()
                Return
            Catch ex As COMException When ex.ErrorCode = &H800AC472 OrElse ex.ErrorCode = &H80010001 OrElse ex.ErrorCode = &H8001010A
                If retries <= 0 Then Throw
                retries -= 1
                System.Threading.Thread.Sleep(150)
                Application.DoEvents()
            End Try
        End While
    End Sub

    Private Sub SetConnectionStatus(id As Integer, slot As String, status As String)
        wbMath.CoreWebView2.ExecuteScriptAsync($"if(typeof window.updateExcelConnectionStatus === 'function') window.updateExcelConnectionStatus({id}, '{slot}', '{status}');")
    End Sub

    Private Sub HandleExcelForceRead(id As Integer, slot As String, address As String, filePath As String)
        Try
            If String.IsNullOrWhiteSpace(address) Then Return
            EnsureExcelConnected()
            If _excelApp Is Nothing Then Return

            Dim wb As Excel.Workbook = Nothing
            If Not String.IsNullOrWhiteSpace(filePath) Then
                wb = EnsureWorkbookOpen(filePath)
            End If

            If wb Is Nothing Then
                SetConnectionStatus(id, slot, "broken")
                Return
            End If

            Dim r As Excel.Range = ComRetry(Function() _excelApp.Range(address))
            Dim tsv = ComRetry(Function() ConvertRangeToTsv(r))

            Dim jsTsv = JsonConvert.SerializeObject(tsv)
            Dim jsAddress = JsonConvert.SerializeObject(address)
            Dim jsPath = JsonConvert.SerializeObject(If(filePath, ""))

            If slot = "A" Then
                wbMath.CoreWebView2.ExecuteScriptAsync($"window.updateTableFromExcel({id}, {jsTsv}, {jsAddress}, {jsPath});")
            ElseIf slot = "B" Then
                wbMath.CoreWebView2.ExecuteScriptAsync($"window.updateTableBFromExcel({id}, {jsTsv}, {jsAddress}, {jsPath});")
            End If
            SetConnectionStatus(id, slot, "connected")
        Catch ex As Exception
            SetConnectionStatus(id, slot, "broken")
        End Try
    End Sub

    Private Sub HandleExcelRelink(id As Integer, slot As String, existingRange As String)
        Using dialog As New OpenFileDialog With {
            .Filter = "Excel Workbooks (*.xlsx;*.xls;*.xlsm)|*.xlsx;*.xls;*.xlsm|All files (*.*)|*.*",
            .Title = "Select Replacement Excel File"
        }
            If dialog.ShowDialog(Me) <> DialogResult.OK Then Return

            EnsureExcelConnected()
            Dim wb As Excel.Workbook = EnsureWorkbookOpen(dialog.FileName)
            If wb Is Nothing Then
                MessageBox.Show("Could not open the selected Excel file.", "Excel Relink", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            If String.IsNullOrWhiteSpace(existingRange) Then
                MessageBox.Show("No existing range was found. Please use the 'Link' button to create a new link.", "Excel Relink", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            Try
                ' 1. Parse the old external address (e.g., '[Old.xlsx]Sheet 1'!$A$1:$B$2)
                Dim bangIdx As Integer = existingRange.LastIndexOf("!")
                If bangIdx = -1 Then Throw New Exception("Invalid existing range format.")

                Dim cellsPart As String = existingRange.Substring(bangIdx + 1)
                Dim sheetPart As String = existingRange.Substring(0, bangIdx)

                ' Strip out the old workbook name
                Dim bracketIdx As Integer = sheetPart.LastIndexOf("]")
                If bracketIdx > -1 Then
                    sheetPart = sheetPart.Substring(bracketIdx + 1)
                End If
                ' Strip trailing single quote if present
                If sheetPart.EndsWith("'") Then
                    sheetPart = sheetPart.Substring(0, sheetPart.Length - 1)
                End If

                ' 2. Locate the matching sheet in the NEW workbook
                Dim sheet As Excel.Worksheet = ComRetry(Function() TryCast(wb.Worksheets(sheetPart), Excel.Worksheet))
                If sheet Is Nothing Then
                    Throw New Exception($"Could not find a sheet named '{sheetPart}' in the new workbook.")
                End If

                ' 3. Get the range and generate the brand new external address
                Dim r As Excel.Range = ComRetry(Function() sheet.Range(cellsPart))
                Dim newExternalAddress As String = ComRetry(Function() r.Address(External:=True))

                Dim jsAddress = JsonConvert.SerializeObject(newExternalAddress)
                Dim jsPath = JsonConvert.SerializeObject(dialog.FileName)

                If slot = "A" OrElse slot = "B" Then
                    Dim tsvData = ComRetry(Function() ConvertRangeToTsv(r))
                    Dim jsTsv = JsonConvert.SerializeObject(tsvData)
                    Dim fn = If(slot = "A", "updateTableFromExcel", "updateTableBFromExcel")
                    wbMath.CoreWebView2.ExecuteScriptAsync($"window.{fn}({id}, {jsTsv}, {jsAddress}, {jsPath});")

                    ' Crucial: Update internal VB.NET tracking dictionary so future typing pushes to the new file
                    Dim key As String = id.ToString() & "_" & slot
                    _linkedTables(key) = newExternalAddress
                Else
                    wbMath.CoreWebView2.ExecuteScriptAsync($"window.setExcelResultRange({id}, {jsAddress}, {jsPath});")
                End If

                SetConnectionStatus(id, slot, "connected")

            Catch ex As Exception
                MessageBox.Show($"Failed to relink: {ex.Message}" & vbCrLf & "Ensure the new file has a sheet with the exact same name.", "Excel Relink", MessageBoxButtons.OK, MessageBoxIcon.Error)
                SetConnectionStatus(id, slot, "broken")
            End Try
        End Using
    End Sub
    Private Sub EnsureExcelConnected()
        Try
            If _excelApp IsNot Nothing Then
                Dim test As String = ComRetry(Function() _excelApp.Name)

                ' FIX: If we are holding an empty instance, check if the user manually opened a real one
                If ComRetry(Function() _excelApp.Workbooks.Count) = 0 Then
                    Try
                        Dim activeApp = DirectCast(Marshal.GetActiveObject("Excel.Application"), Excel.Application)
                        If activeApp IsNot Nothing AndAlso activeApp.Workbooks.Count > 0 Then
                            ' Switch to the user's active instance
                            Try
                                RemoveHandler _excelApp.SheetChange, AddressOf OnExcelSheetChange
                            Catch
                            End Try
                            _excelApp = activeApp
                            AddHandler _excelApp.SheetChange, AddressOf OnExcelSheetChange
                        End If
                    Catch
                        ' No other active instance found, continue using the empty one
                    End Try
                End If
            End If
        Catch ex As Exception
            _excelApp = Nothing
        End Try

        If _excelApp Is Nothing Then
            Try
                _excelApp = DirectCast(Marshal.GetActiveObject("Excel.Application"), Excel.Application)
                AddHandler _excelApp.SheetChange, AddressOf OnExcelSheetChange
            Catch ex As Exception
                _excelApp = New Excel.Application()
                ComRetryAction(Sub() _excelApp.Visible = True)
                AddHandler _excelApp.SheetChange, AddressOf OnExcelSheetChange
            End Try
        End If
    End Sub

    Private Function EnsureWorkbookOpen(filePath As String) As Excel.Workbook
        If String.IsNullOrWhiteSpace(filePath) Then Return Nothing

        Dim wbs As Excel.Workbooks = ComRetry(Function() _excelApp.Workbooks)
        Dim count As Integer = ComRetry(Function() wbs.Count)

        If Not filePath.Contains("\") AndAlso Not filePath.Contains("/") Then
            ' Safely iterate using a For loop instead of For Each
            For i As Integer = 1 To count
                Dim b As Excel.Workbook = ComRetry(Function() wbs.Item(i))
                If ComRetry(Function() b.Name).Equals(filePath, StringComparison.OrdinalIgnoreCase) Then Return b
            Next
            Return Nothing
        End If

        ' Safely iterate using a For loop instead of For Each
        For i As Integer = 1 To count
            Dim b As Excel.Workbook = ComRetry(Function() wbs.Item(i))
            If ComRetry(Function() b.FullName).Equals(filePath, StringComparison.OrdinalIgnoreCase) Then Return b
        Next

        If File.Exists(filePath) Then
            Return ComRetry(Function() wbs.Open(filePath, UpdateLinks:=0, ReadOnly:=False, IgnoreReadOnlyRecommended:=True))
        End If

        Return Nothing
    End Function

    Private Sub HandleExcelOpen(filePath As String)
        Try
            EnsureExcelConnected()
            ComRetryAction(Sub() _excelApp.Visible = True)

            Dim wb As Excel.Workbook = EnsureWorkbookOpen(filePath)
            If wb IsNot Nothing Then
                ' ComRetryAction(Sub() wb.Activate())
                Try
                    SetForegroundWindow(New IntPtr(_excelApp.Hwnd))
                Catch
                End Try
            Else
                If Not filePath.Contains("\") AndAlso Not filePath.Contains("/") Then
                    MessageBox.Show($"The workbook '{filePath}' was never saved to your computer." & vbCrLf &
                                    "You must save the file in Excel before linking if you want to reopen it later.",
                                    "Excel Link", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Else
                    MessageBox.Show($"Could not find the Excel file at:" & vbCrLf & filePath & vbCrLf &
                                    "It may have been moved or deleted.",
                                    "Excel Link", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                End If
            End If
        Catch ex As Exception
            MessageBox.Show("Could not open the Excel workbook." & Environment.NewLine & ex.Message,
                            "Excel Link", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try
    End Sub

    Private Sub HandleExcelLink(id As Integer, tablePrefix As String)
        Try
            EnsureExcelConnected()
            If _excelApp Is Nothing Then Throw New Exception("Excel is not running.")

            Dim wbCount As Integer = ComRetry(Function() _excelApp.Workbooks.Count)
            If wbCount = 0 Then
                ' KILL THE GHOST: Destroy the empty background instance so it drops out of the Windows ROT
                Try
                    RemoveHandler _excelApp.SheetChange, AddressOf OnExcelSheetChange
                    _excelApp.Quit()
                    Marshal.ReleaseComObject(_excelApp)
                Catch
                End Try
                _excelApp = Nothing
                Throw New Exception("Connected to an empty Excel process. We have cleared it. Please ensure your file is open and click Link again.")
            End If

            Dim selection As Excel.Range = ComRetry(Function() _excelApp.ActiveWindow.RangeSelection)
            If selection Is Nothing Then Throw New Exception("Please select a range of cells in Excel first.")

            Dim wb As Excel.Workbook = ComRetry(Function() selection.Worksheet.Parent)
            Dim filePath As String = ComRetry(Function() wb.FullName)
            Dim address As String = ComRetry(Function() selection.Address(External:=True))

            Dim key As String = id.ToString() & "_" & tablePrefix
            _linkedTables(key) = address

            Dim tsvData = ComRetry(Function() ConvertRangeToTsv(selection))

            Dim jsTsv = JsonConvert.SerializeObject(tsvData)
            Dim jsAddress = JsonConvert.SerializeObject(address)
            Dim jsPath = JsonConvert.SerializeObject(filePath)

            Dim script As String
            If tablePrefix = "A" Then
                script = $"window.updateTableFromExcel({id}, {jsTsv}, {jsAddress}, {jsPath});"
            Else
                script = $"window.updateTableBFromExcel({id}, {jsTsv}, {jsAddress}, {jsPath});"
            End If

            wbMath.CoreWebView2.ExecuteScriptAsync(script)
            SetConnectionStatus(id, tablePrefix, "connected")

        Catch ex As COMException When ex.ErrorCode = &H800AC472
            MessageBox.Show("Excel is currently busy. Please press Enter in Excel to finish editing the cell, then try linking again.", "Excel Busy", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            HandleExcelUnlink(id, tablePrefix)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Excel Link", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            HandleExcelUnlink(id, tablePrefix)
        End Try
    End Sub

    Private Sub HandleExcelLinkResult(id As Integer)
        Try
            EnsureExcelConnected()
            If _excelApp Is Nothing Then Throw New Exception("Excel is not running.")

            Dim wbCount As Integer = ComRetry(Function() _excelApp.Workbooks.Count)
            If wbCount = 0 Then
                ' KILL THE GHOST
                Try
                    RemoveHandler _excelApp.SheetChange, AddressOf OnExcelSheetChange
                    _excelApp.Quit()
                    Marshal.ReleaseComObject(_excelApp)
                Catch
                End Try
                _excelApp = Nothing
                Throw New Exception("Connected to an empty Excel process. We have cleared it. Please ensure your file is open and click Link again.")
            End If

            Dim selection As Excel.Range = ComRetry(Function() _excelApp.ActiveWindow.RangeSelection)
            If selection Is Nothing Then Throw New Exception("Please select a target cell/range in Excel first.")

            Dim wb As Excel.Workbook = ComRetry(Function() selection.Worksheet.Parent)
            Dim filePath As String = ComRetry(Function() wb.FullName)
            Dim address As String = ComRetry(Function() selection.Address(External:=True))

            Dim jsAddress = JsonConvert.SerializeObject(address)
            Dim jsPath = JsonConvert.SerializeObject(filePath)

            Dim script = $"window.setExcelResultRange({id}, {jsAddress}, {jsPath});"
            wbMath.CoreWebView2.ExecuteScriptAsync(script)
            SetConnectionStatus(id, "Result", "connected")

        Catch ex As COMException When ex.ErrorCode = &H800AC472
            MessageBox.Show("Excel is currently busy. Please press Enter in Excel, then try linking again.", "Excel Busy", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            HandleExcelUnlinkResult(id)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Excel Output Link", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            HandleExcelUnlinkResult(id)
        End Try
    End Sub

    Private Sub HandleExcelUpdate(id As Integer, tablePrefix As String, address As String, filePath As String, tsvData As String)
        Try
            ' Validate file exists before attempting connection
            If Not String.IsNullOrWhiteSpace(filePath) AndAlso Not File.Exists(filePath) Then
                SetConnectionStatus(id, tablePrefix, "broken")
                Return
            End If

            EnsureExcelConnected()
            If _excelApp Is Nothing Then Return

            Dim wb As Excel.Workbook = EnsureWorkbookOpen(filePath)
            If wb Is Nothing Then
                SetConnectionStatus(id, tablePrefix, "broken")
                Return
            End If

            RemoveHandler _excelApp.SheetChange, AddressOf OnExcelSheetChange
            'ComRetryAction(Sub() wb.Activate())

            Dim range As Excel.Range = ComRetry(Function() _excelApp.Range(address))
            Dim rows() As String = tsvData.Split({vbLf, vbCr}, StringSplitOptions.RemoveEmptyEntries)

            If rows.Length > 0 Then
                Dim cols() As String = rows(0).Split(vbTab)
                Dim data(rows.Length - 1, cols.Length - 1) As Object

                For r As Integer = 0 To rows.Length - 1
                    Dim cellValues() As String = rows(r).Split(vbTab)
                    For c As Integer = 0 To Math.Min(cols.Length - 1, cellValues.Length - 1)
                        data(r, c) = cellValues(c)
                    Next
                Next

                Dim startCell As Excel.Range = ComRetry(Function() range.Cells(1, 1))
                Dim targetRange As Excel.Range = ComRetry(Function() startCell.Resize(rows.Length, cols.Length))
                ComRetryAction(Sub() targetRange.Value = data)

                Dim key As String = id.ToString() & "_" & tablePrefix
                _linkedTables(key) = ComRetry(Function() targetRange.Address(External:=True))
                SetConnectionStatus(id, tablePrefix, "connected")
            End If

            AddHandler _excelApp.SheetChange, AddressOf OnExcelSheetChange
        Catch ex As Exception
            SetConnectionStatus(id, tablePrefix, "broken")
        End Try
    End Sub

    Private Sub HandleExcelUpdateResult(id As Integer, address As String, filePath As String, tsvData As String)
        Try
            ' Validate file exists before attempting connection
            If Not String.IsNullOrWhiteSpace(filePath) AndAlso Not File.Exists(filePath) Then
                SetConnectionStatus(id, "Result", "broken")
                Return
            End If

            EnsureExcelConnected()
            If _excelApp Is Nothing Then Return

            Dim wb As Excel.Workbook = Nothing
            If Not String.IsNullOrWhiteSpace(filePath) Then
                wb = EnsureWorkbookOpen(filePath)
            End If

            If wb Is Nothing Then
                SetConnectionStatus(id, "Result", "broken")
                Return
            End If

            RemoveHandler _excelApp.SheetChange, AddressOf OnExcelSheetChange

            'ComRetryAction(Sub() wb.Activate())

            Dim range As Excel.Range = ComRetry(Function() _excelApp.Range(address))
            Dim rows() As String = tsvData.Split({vbLf, vbCr}, StringSplitOptions.RemoveEmptyEntries)

            If rows.Length > 0 Then
                Dim cols() As String = rows(0).Split(vbTab)
                Dim data(rows.Length - 1, cols.Length - 1) As Object

                For r As Integer = 0 To rows.Length - 1
                    Dim cellValues() As String = rows(r).Split(vbTab)
                    For c As Integer = 0 To Math.Min(cols.Length - 1, cellValues.Length - 1)
                        Dim numVal As Double
                        ' Try to parse numbers natively, otherwise pass as text
                        If Double.TryParse(cellValues(c), NumberStyles.Float, CultureInfo.InvariantCulture, numVal) Then
                            data(r, c) = numVal
                        Else
                            data(r, c) = cellValues(c)
                        End If
                    Next
                Next

                Dim startCell As Excel.Range = ComRetry(Function() range.Cells(1, 1))
                Dim targetRange As Excel.Range = ComRetry(Function() startCell.Resize(rows.Length, cols.Length))
                ComRetryAction(Sub() targetRange.Value = data)

                SetConnectionStatus(id, "Result", "connected")
            End If

            AddHandler _excelApp.SheetChange, AddressOf OnExcelSheetChange
        Catch ex As Exception
            SetConnectionStatus(id, "Result", "broken")
        End Try
    End Sub

    Private Sub HandleExcelUnlink(id As Integer, tablePrefix As String)
        Dim key As String = id.ToString() & "_" & tablePrefix
        If _linkedTables.ContainsKey(key) Then
            _linkedTables.Remove(key)
        End If
        Dim fn As String = If(tablePrefix = "A", "unlinkExcelRange", "unlinkExcelRangeB")
        wbMath.CoreWebView2.ExecuteScriptAsync($"if(typeof window.{fn} === 'function') window.{fn}({id});")
    End Sub

    Private Sub HandleExcelUnlinkResult(id As Integer)
        wbMath.CoreWebView2.ExecuteScriptAsync($"if(typeof window.unlinkExcelResultRange === 'function') window.unlinkExcelResultRange({id});")
    End Sub

    Private Sub OnExcelSheetChange(sh As Object, target As Excel.Range)
        If _linkedTables.Count = 0 Then Return

        For Each kvp In _linkedTables.ToList()
            Dim keyParts = kvp.Key.Split("_"c)
            Dim id As Integer = Integer.Parse(keyParts(0))
            Dim tablePrefix As String = keyParts(1)
            Dim address As String = kvp.Value

            Try
                Dim linkedRange As Excel.Range = ComRetry(Function() _excelApp.Range(address))
                Dim sourceSheet As Excel.Worksheet = ComRetry(Function() linkedRange.Worksheet)

                If sh Is sourceSheet Then
                    Dim intersect As Excel.Range = ComRetry(Function() _excelApp.Intersect(target, linkedRange))
                    If intersect IsNot Nothing Then
                        Dim tsvData = ComRetry(Function() ConvertRangeToTsv(linkedRange))
                        Dim wb As Excel.Workbook = ComRetry(Function() sourceSheet.Parent)
                        Dim filePath As String = ComRetry(Function() wb.FullName)

                        Dim jsTsv = JsonConvert.SerializeObject(tsvData)
                        Dim jsAddress = JsonConvert.SerializeObject(address)
                        Dim jsPath = JsonConvert.SerializeObject(filePath)

                        Me.Invoke(Sub()
                                      Dim script As String
                                      If tablePrefix = "A" Then
                                          script = $"window.updateTableFromExcel({id}, {jsTsv}, {jsAddress}, {jsPath});"
                                      Else
                                          script = $"window.updateTableBFromExcel({id}, {jsTsv}, {jsAddress}, {jsPath});"
                                      End If
                                      wbMath.CoreWebView2.ExecuteScriptAsync(script)
                                  End Sub)
                    End If
                End If
            Catch ex As Exception
            End Try
        Next
    End Sub

    Private Function ConvertRangeToTsv(range As Excel.Range) As String
        If range Is Nothing Then Return ""

        Dim cellCount As Integer = ComRetry(Function() range.Cells.Count)
        If cellCount = 1 Then
            Dim val As Object = ComRetry(Function() range.Value)
            Return If(val Is Nothing, "0", val.ToString())
        End If

        Dim values As Object(,) = TryCast(ComRetry(Function() range.Value), Object(,))
        If values Is Nothing Then Return ""

        Dim sb As New StringBuilder()
        Dim rowLower = values.GetLowerBound(0)
        Dim rowUpper = values.GetUpperBound(0)
        Dim colLower = values.GetLowerBound(1)
        Dim colUpper = values.GetUpperBound(1)

        For r As Integer = rowLower To rowUpper
            For c As Integer = colLower To colUpper
                Dim val = values(r, c)
                sb.Append(If(val Is Nothing, "0", val.ToString()))
                If c < colUpper Then sb.Append(vbTab)
            Next
            If r < rowUpper Then sb.Append(vbLf)
        Next

        Return sb.ToString()
    End Function

    Private Sub SyncBlockFromExcel(block As JObject, rangeKey As String, pathKey As String, dataKey As String, rowKey As String, colKey As String, slotSuffix As String)
        Dim rangeAddr = block.Value(Of String)(rangeKey)
        Dim pathStr = block.Value(Of String)(pathKey)

        If Not String.IsNullOrWhiteSpace(rangeAddr) Then
            If Not String.IsNullOrWhiteSpace(pathStr) AndAlso Not File.Exists(pathStr) Then
                block("excelStatus" & slotSuffix) = "broken"
                Return
            End If

            Try
                Dim wb As Excel.Workbook = EnsureWorkbookOpen(pathStr)
                If wb IsNot Nothing Then
                    ' ComRetryAction(Sub() wb.Activate())
                    Dim r As Excel.Range = ComRetry(Function() _excelApp.Range(rangeAddr))
                    Dim tsv = ComRetry(Function() ConvertRangeToTsv(r))

                    block(dataKey) = tsv
                    Dim rows = tsv.Split({vbLf, vbCr}, StringSplitOptions.RemoveEmptyEntries)
                    If rows.Length > 0 Then
                        block(rowKey) = rows.Length
                        block(colKey) = rows(0).Split(vbTab).Length
                    End If
                    block("excelStatus" & slotSuffix) = "connected"
                Else
                    block("excelStatus" & slotSuffix) = "broken"
                End If
            Catch ex As Exception
                block("excelStatus" & slotSuffix) = "broken"
            End Try
        End If
    End Sub

    Private Sub SetPresentationFullscreen(enabled As Boolean)
        If enabled = _presentationFullscreen Then Return

        If enabled Then
            _windowStateBeforePresentation = WindowState
            _formBorderStyleBeforePresentation = FormBorderStyle
            _boundsBeforePresentation = Bounds
            _topMostBeforePresentation = TopMost
            WindowState = FormWindowState.Normal
            FormBorderStyle = FormBorderStyle.None
            Bounds = Screen.FromControl(Me).Bounds
            TopMost = True
        Else
            TopMost = _topMostBeforePresentation
            FormBorderStyle = _formBorderStyleBeforePresentation
            Bounds = _boundsBeforePresentation
            WindowState = _windowStateBeforePresentation
        End If

        _presentationFullscreen = enabled
    End Sub




    'Private Sub SetPresentationFullscreen(enabled As Boolean)
    '    ' Ensure we are safely updating the UI on the main thread
    '    If Me.InvokeRequired Then
    '        Me.Invoke(Sub() SetPresentationFullscreen(enabled))
    '        Return
    '    End If

    '    If enabled = _presentationFullscreen Then Return

    '    If enabled Then
    '        ' Save the current state so we can restore it later
    '        _windowStateBeforePresentation = Me.WindowState
    '        _formBorderStyleBeforePresentation = Me.FormBorderStyle
    '        _boundsBeforePresentation = Me.Bounds
    '        _topMostBeforePresentation = Me.TopMost

    '        ' Suspend layout to prevent flickering
    '        Me.SuspendLayout()

    '        ' The bulletproof WinForms Fullscreen combo: No Border + Maximized
    '        Me.FormBorderStyle = FormBorderStyle.None
    '        Me.WindowState = FormWindowState.Maximized
    '        Me.TopMost = True

    '        Me.ResumeLayout()
    '    Else
    '        ' Restore the original state
    '        Me.SuspendLayout()

    '        Me.TopMost = _topMostBeforePresentation
    '        Me.FormBorderStyle = _formBorderStyleBeforePresentation
    '        Me.WindowState = _windowStateBeforePresentation

    '        ' Only restore bounds if the window was previously in a Normal (windowed) state
    '        If Me.WindowState = FormWindowState.Normal Then
    '            Me.Bounds = _boundsBeforePresentation
    '        End If

    '        Me.ResumeLayout()
    '    End If

    '    _presentationFullscreen = enabled
    'End Sub

    Private Sub SaveCanvas(document As JToken, currentPath As String, isSaveAs As Boolean)
        Dim targetPath As String = currentPath

        ' If it's Save As, or if the file has never been saved (no path / doesn't exist), prompt the user
        If isSaveAs OrElse String.IsNullOrWhiteSpace(targetPath) OrElse Not File.Exists(targetPath) Then
            Dim suggestedName = Path.GetFileName(targetPath)
            If String.IsNullOrWhiteSpace(suggestedName) Then suggestedName = "Untitled.mc.json"

            Using dialog As New SaveFileDialog With {
                .Filter = "Math canvas (*.mc.json)|*.mc.json|Legacy (*.mathcanvas.json)|*.mathcanvas.json|JSON (*.json)|*.json",
                .DefaultExt = "mc.json",
                .AddExtension = True,
                .FileName = suggestedName,
                .Title = If(isSaveAs, "Save As", "Save Math Canvas")}
                If dialog.ShowDialog(Me) <> DialogResult.OK Then Return
                targetPath = dialog.FileName
            End Using
        End If

        Try
            ' Save the file
            File.WriteAllText(targetPath, document.ToString(Formatting.Indented), Encoding.UTF8)

            ' Callback to the UI to update the tab title and clear the dirty flag
            Dim newTitle = Path.GetFileName(targetPath)
            Dim jsTitle = JsonConvert.SerializeObject(newTitle)
            Dim jsPath = JsonConvert.SerializeObject(targetPath)
            wbMath.CoreWebView2.ExecuteScriptAsync($"if(typeof window.documentSaved === 'function') window.documentSaved({jsTitle}, {jsPath});")
        Catch ex As Exception
            MessageBox.Show("Could not save the file: " & ex.Message, "Save Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub SaveDxf(fileName As String, drawing As JToken)
        Dim suggestedName = Path.GetFileName(fileName)
        If String.IsNullOrWhiteSpace(suggestedName) Then suggestedName = "engineering-detail.dxf"

        Using dialog As New SaveFileDialog With {
            .Filter = "AutoCAD DXF (*.dxf)|*.dxf|All files (*.*)|*.*",
            .DefaultExt = "dxf",
            .AddExtension = True,
            .FileName = suggestedName,
            .Title = "Export 2D CAD Drawing"}
            If dialog.ShowDialog(Me) <> DialogResult.OK Then Return

            Dim document = CreateDxfDocument(drawing)
            If Not document.Save(dialog.FileName) Then
                Throw New IOException("netDxf could not save the drawing.")
            End If
        End Using
    End Sub

    Private Shared Function CreateDxfDocument(drawing As JToken) As DxfDocument
        If drawing Is Nothing OrElse drawing.Type <> JTokenType.Object Then
            Throw New InvalidDataException("The CAD drawing data is missing.")
        End If

        Dim header As New HeaderVariables With {
            .InsUnits = GetDrawingUnits(drawing.Value(Of String)("drawingUnit"))}
        Dim document As New DxfDocument(header)
        Dim layers As New Dictionary(Of String, Layer)(StringComparer.OrdinalIgnoreCase) From {
            {"0", document.Layers("0")}}

        Dim layerData = TryCast(drawing("layers"), JArray)
        If layerData IsNot Nothing Then
            For Each item In layerData.OfType(Of JObject)()
                AddDxfLayer(document, layers, item)
            Next
        End If

        Dim entities = TryCast(drawing("entities"), JArray)
        If entities Is Nothing Then Throw New InvalidDataException("The CAD entity list is missing.")

        For Each entity In entities.OfType(Of JObject)()
            AddDxfEntity(document, layers, entity)
        Next

        Return document
    End Function

    Private Shared Sub AddDxfLayer(document As DxfDocument,
                                   layers As Dictionary(Of String, Layer),
                                   data As JObject)
        Dim name = GetLayerName(data.Value(Of String)("name"))
        Dim layer As Layer
        If layers.TryGetValue(name, layer) Then
            ApplyLayerStyle(layer, data)
            Return
        End If

        layer = New Layer(name)
        ApplyLayerStyle(layer, data)
        document.Layers.Add(layer)
        layers.Add(name, layer)
    End Sub

    Private Shared Sub ApplyLayerStyle(layer As Layer, data As JObject)
        layer.IsVisible = data.Value(Of Boolean?)("visible").GetValueOrDefault(True)
        layer.Color = GetColor(data.Value(Of String)("color"), AciColor.FromCadIndex(7))
        layer.Lineweight = GetLineweight(data("lineWeight"), Lineweight.Default)
    End Sub

    Private Shared Sub AddDxfEntity(document As DxfDocument,
                                    layers As Dictionary(Of String, Layer),
                                    data As JObject)
        Dim layer = GetOrCreateLayer(document, layers, data.Value(Of String)("layer"))

        Select Case data.Value(Of String)("type")
            Case "line"
                Dim p1 As New Vector3(GetNumber(data, "x1", 0.0), GetNumber(data, "y1", 0.0), GetNumber(data, "z1", 0.0))
                Dim p2 As New Vector3(GetNumber(data, "x2", 0.0), GetNumber(data, "y2", 0.0), GetNumber(data, "z2", 0.0))
                Dim lineEntity As New Line(p1, p2)
                ApplyEntityStyle(lineEntity, layer, data)
                document.AddEntity(lineEntity)

            Case "polyline"
                Dim points = TryCast(data("points"), JArray)
                If points Is Nothing OrElse points.Count < 2 Then Return

                Dim vertices3D As New List(Of Vector3)()
                For Each point In points
                    Dim coords = TryCast(point, JArray)
                    If coords IsNot Nothing AndAlso coords.Count >= 2 Then
                        Dim xVal As Double = GetNumber(coords(0), "X", 0.0)
                        Dim yVal As Double = GetNumber(coords(1), "Y", 0.0)
                        Dim zVal As Double = If(coords.Count >= 3, GetNumber(coords(2), "Z", 0.0), 0.0)
                        vertices3D.Add(New Vector3(xVal, yVal, zVal))
                    End If
                Next

                Dim isClosed As Boolean = data.Value(Of Boolean?)("closed").GetValueOrDefault()
                Dim poly3D As New Polyline(vertices3D, isClosed)
                ApplyEntityStyle(poly3D, layer, data)
                document.AddEntity(poly3D)

            Case "circle"
                Dim center As New Vector3(GetNumber(data, "cx", 0.0), GetNumber(data, "cy", 0.0), GetNumber(data, "cz", 0.0))
                Dim radius = Math.Abs(GetNumber(data, "r", 1.0))
                Dim circleEntity As New Circle(center, radius)
                ApplyEntityStyle(circleEntity, layer, data)
                document.AddEntity(circleEntity)

            Case "arc"
                Dim center As New Vector3(GetNumber(data, "cx", 0.0), GetNumber(data, "cy", 0.0), GetNumber(data, "cz", 0.0))
                Dim radius = Math.Abs(GetNumber(data, "r", 1.0))
                Dim startAngle = GetNumber(data, "start", 0.0)
                Dim endAngle = GetNumber(data, "end", 360.0)
                Dim arcEntity As New Arc(center, radius, startAngle, endAngle)
                ApplyEntityStyle(arcEntity, layer, data)
                document.AddEntity(arcEntity)

            Case "text"
                Dim txtVal As String = CleanDxfText(data.Value(Of String)("text"))
                If Not String.IsNullOrWhiteSpace(txtVal) Then
                    Dim pos As New Vector3(GetNumber(data, "x", 0.0), GetNumber(data, "y", 0.0), GetNumber(data, "z", 0.0))
                    Dim height As Double = Math.Max(0.1, Math.Abs(GetNumber(data, "height", 10.0)))
                    Dim textEntity As New netDxf.Entities.Text(txtVal, pos, height)

                    ' Set Normal Vector: Default to (0, -1, 0) so 3D text stands upright facing front (-Y axis)
                    Dim nx As Double = GetNumber(data, "nx", 0.0)
                    Dim ny As Double = GetNumber(data, "ny", -1.0)
                    Dim nz As Double = GetNumber(data, "nz", 0.0)
                    textEntity.Normal = New Vector3(nx, ny, nz)

                    ' Optional text rotation angle around the normal
                    Dim rotation As Double = GetNumber(data, "rotation", 0.0)
                    textEntity.Rotation = rotation

                    ApplyEntityStyle(textEntity, layer, data)
                    document.AddEntity(textEntity)
                End If

            Case "dimension"
                AddDimension(document, layer, data)
        End Select
    End Sub

    Private Shared Sub AddDimension(document As DxfDocument, layer As Layer, data As JObject)
        Try
            Dim x1 = GetNumber(data, "x1", 0.0)
            Dim y1 = GetNumber(data, "y1", 0.0)
            Dim x2 = GetNumber(data, "x2", 0.0)
            Dim y2 = GetNumber(data, "y2", 0.0)
            Dim dx = x2 - x1
            Dim dy = y2 - y1
            Dim length = Math.Sqrt(dx * dx + dy * dy)
            If length <= 0 Then Return

            Dim offset = GetNumber(data, "offset", 0.0)
            Dim textHeight = Math.Max(0.1, Math.Abs(GetNumber(data, "height", GetNumber(data, "textHeight", 10.0))))
            Dim dimensionStyle As New DimensionStyle("MathCanvas_" & Guid.NewGuid().ToString("N")) With {
                .ArrowSize = textHeight,
                .TextHeight = textHeight,
                .TextOffset = textHeight * 0.25,
                .TextVerticalPlacement = DimensionStyleTextVerticalPlacement.Above,
                .ExtLineOffset = textHeight * 0.25,
                .ExtLineExtend = textHeight * 0.5
            }
            Dim definitionLine = If(offset < 0,
                                    New Line(New Vector2(x2, y2), New Vector2(x1, y1)),
                                    New Line(New Vector2(x1, y1), New Vector2(x2, y2)))
            Dim entity As New AlignedDimension(definitionLine, Math.Abs(offset), dimensionStyle)

            ' Set Facing Normal Vector (Defaults to Front Facing: nx=0, ny=-1, nz=0)
            Dim nx As Double = GetNumber(data, "nx", 0.0)
            Dim ny As Double = GetNumber(data, "ny", -1.0)
            Dim nz As Double = GetNumber(data, "nz", 0.0)
            entity.Normal = New Vector3(nx, ny, nz)

            Dim label = data.Value(Of String)("text")
            If Not String.IsNullOrWhiteSpace(label) Then entity.UserText = CleanDxfText(label)
            ApplyEntityStyle(entity, layer, data)
            document.AddEntity(entity)
        Catch ex As Exception
            ' Skip unparseable dimension objects
        End Try
    End Sub

    Private Shared Sub AddLine(document As DxfDocument, layer As Layer, data As JObject,
                               x1 As Double, y1 As Double, x2 As Double, y2 As Double)
        Dim entity As New Line(New Vector2(x1, y1), New Vector2(x2, y2))
        ApplyEntityStyle(entity, layer, data)
        document.AddEntity(entity)
    End Sub

    Private Shared Sub AddText(document As DxfDocument, layer As Layer, data As JObject,
                               x As Double, y As Double, height As Double, value As String, rotation As Double)
        Dim entity As New netDxf.Entities.Text(CleanDxfText(value), New Vector2(x, y), Math.Max(0.1, Math.Abs(height))) With {
            .Rotation = rotation}
        ApplyEntityStyle(entity, layer, data)
        document.AddEntity(entity)
    End Sub

    Private Shared Sub ApplyEntityStyle(entity As EntityObject, layer As Layer, data As JObject)
        entity.Layer = layer
        entity.Linetype = GetLinetype(data.Value(Of String)("lineType"))
        entity.Lineweight = GetLineweight(data("lineWeight"), Lineweight.ByLayer)
        Dim color = data.Value(Of String)("color")
        If Not String.IsNullOrWhiteSpace(color) Then entity.Color = GetColor(color, AciColor.ByLayer)
    End Sub

    Private Shared Function GetOrCreateLayer(document As DxfDocument,
                                             layers As Dictionary(Of String, Layer),
                                             name As String) As Layer
        name = GetLayerName(name)
        Dim layer As Layer
        If layers.TryGetValue(name, layer) Then Return layer

        layer = New Layer(name)
        document.Layers.Add(layer)
        layers.Add(name, layer)
        Return layer
    End Function

    Private Shared Function GetDrawingUnits(value As String) As DrawingUnits
        Select Case If(value, String.Empty).ToLowerInvariant()
            Case "in" : Return DrawingUnits.Inches
            Case "ft" : Return DrawingUnits.Feet
            Case "mm" : Return DrawingUnits.Millimeters
            Case "cm" : Return DrawingUnits.Centimeters
            Case "m" : Return DrawingUnits.Meters
            Case Else : Return DrawingUnits.Unitless
        End Select
    End Function

    Private Shared Function GetLinetype(value As String) As Linetype
        Select Case If(value, String.Empty).ToUpperInvariant()
            Case "CENTER" : Return Linetype.Center
            Case "DASHED" : Return Linetype.Dashed
            Case "CONTINUOUS" : Return Linetype.Continuous
            Case Else : Return Linetype.ByLayer
        End Select
    End Function

    Private Shared Function GetColor(value As String, fallback As AciColor) As AciColor
        If String.IsNullOrWhiteSpace(value) OrElse Not value.StartsWith("#") OrElse value.Length <> 7 Then Return fallback
        Dim red As Byte
        Dim green As Byte
        Dim blue As Byte
        If Byte.TryParse(value.Substring(1, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, red) AndAlso
           Byte.TryParse(value.Substring(3, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, green) AndAlso
           Byte.TryParse(value.Substring(5, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, blue) Then
            Return New AciColor(red, green, blue)
        End If
        Return fallback
    End Function

    Private Shared Function GetLineweight(value As JToken, fallback As Lineweight) As Lineweight
        If value Is Nothing OrElse value.Type = JTokenType.Null Then Return fallback
        Dim millimeters As Double
        If Not Double.TryParse(value.ToString(), NumberStyles.Float, CultureInfo.InvariantCulture, millimeters) Then Return fallback

        Dim supported = {0, 5, 9, 13, 15, 18, 20, 25, 30, 35, 40, 50, 53, 60, 70, 80, 90, 100, 106, 120, 140, 158, 200, 211}
        Dim requested = Math.Max(0, Math.Min(211, CInt(Math.Round(millimeters * 100))))
        Dim nearest = supported.OrderBy(Function(candidate) Math.Abs(candidate - requested)).First()
        Return DirectCast([Enum].ToObject(GetType(Lineweight), nearest), Lineweight)
    End Function

    ' =========================================================
    ' UNIFIED JSON NUMBER PARSER (Handles JObject and JValue)
    ' =========================================================
    Private Shared Function GetNumber(data As JToken, Optional propertyName As String = "", Optional defaultValue As Double? = Nothing) As Double
        If data Is Nothing OrElse data.Type = JTokenType.Null Then
            If defaultValue.HasValue Then Return defaultValue.Value
            Throw New InvalidDataException(If(String.IsNullOrEmpty(propertyName), "Value", propertyName) & " is required.")
        End If

        ' Extract property from JObject, or use data directly if it's already a scalar JValue
        Dim targetToken As JToken = data
        If data.Type = JTokenType.Object AndAlso Not String.IsNullOrEmpty(propertyName) Then
            targetToken = data(propertyName)
        End If

        If targetToken Is Nothing OrElse targetToken.Type = JTokenType.Null OrElse String.IsNullOrWhiteSpace(targetToken.ToString()) Then
            If defaultValue.HasValue Then Return defaultValue.Value
            Throw New InvalidDataException(If(String.IsNullOrEmpty(propertyName), "Value", propertyName) & " is required.")
        End If

        Dim number As Double
        If Double.TryParse(targetToken.ToString(), NumberStyles.Float, CultureInfo.InvariantCulture, number) AndAlso
           Not Double.IsNaN(number) AndAlso Not Double.IsInfinity(number) Then
            Return number
        End If

        If defaultValue.HasValue Then Return defaultValue.Value
        Throw New InvalidDataException(If(String.IsNullOrEmpty(propertyName), "Value", propertyName) & " must be a finite number.")
    End Function

    'Private Shared Function GetNumber(data As JToken, propertyName As String, Optional defaultValue As Double? = Nothing) As Double
    '    Dim value = If(data.Type = JTokenType.Object, data(propertyName), data)
    '    If value Is Nothing OrElse value.Type = JTokenType.Null OrElse String.IsNullOrWhiteSpace(value.ToString()) Then
    '        If defaultValue.HasValue Then Return defaultValue.Value
    '        Throw New InvalidDataException(propertyName & " is required.")
    '    End If

    '    Dim number As Double
    '    If Not Double.TryParse(value.ToString(), NumberStyles.Float, CultureInfo.InvariantCulture, number) OrElse
    '       Double.IsNaN(number) OrElse Double.IsInfinity(number) Then
    '        Throw New InvalidDataException(propertyName & " must be a finite number.")
    '    End If
    '    Return number
    'End Function

    '' Helper overload for optional Z defaults
    'Private Shared Function GetNumber(data As JToken, propertyName As String, defaultValue As Double) As Double
    '    Dim val = data(propertyName)
    '    If val Is Nothing OrElse String.IsNullOrWhiteSpace(val.ToString()) Then Return defaultValue
    '    Dim result As Double
    '    If Double.TryParse(val.ToString(), NumberStyles.Float, CultureInfo.InvariantCulture, result) Then Return result
    '    Return defaultValue
    'End Function

    Private Shared Function GetLayerName(value As String) As String
        Dim name = If(value, "0").Trim()
        Dim invalidCharacters = New Char() {"<"c, ">"c, "/"c, "\"c, ControlChars.Quote,
                                            ":"c, ";"c, "?"c, "*"c, "|"c, "="c, ","c}
        For Each invalidCharacter In invalidCharacters
            name = name.Replace(invalidCharacter, "_"c)
        Next
        If name.Length = 0 Then name = "0"
        Return If(name.Length <= 255, name, name.Substring(0, 255))
    End Function

    Private Shared Function CleanDxfText(value As String) As String
        Dim text = If(value, String.Empty).
            Replace(vbCr, " ").
            Replace(vbLf, " ").
            Replace(vbTab, " ").
            Replace(ChrW(&H2010), "-").
            Replace(ChrW(&H2011), "-").
            Replace(ChrW(&H2013), "-").
            Replace(ChrW(&H2014), "-").
            Replace(ChrW(&H2212), "-").
            Replace(ChrW(&HB7), ".")
        Return If(text.Length <= 255, text, text.Substring(0, 255))
    End Function


    Private Sub LoadCanvas()
        If Not _webViewReady Then Return

        Using dialog As New OpenFileDialog With {
            .Filter = "Math canvas (*.mc.json;*.mathcanvas.json;*.json;*.mc)|*.mc.json;*.mathcanvas.json;*.json;*.mc|All files (*.*)|*.*",
            .Title = "Open Math Canvas"}
            If dialog.ShowDialog(Me) <> DialogResult.OK Then Return
            Try
                Dim document = JToken.Parse(File.ReadAllText(dialog.FileName, Encoding.UTF8))

                document("title") = Path.GetFileName(dialog.FileName)
                document("filePath") = dialog.FileName ' <--- ADD THIS LINE

                _linkedTables.Clear()
                Dim pagesArray = TryCast(document("pages"), JArray)

                ' 1. Gather all unique Excel paths used in this canvas
                Dim uniquePaths As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

                If pagesArray IsNot Nothing Then
                    For Each page In pagesArray.OfType(Of JObject)()
                        Dim blocksArray = TryCast(page("blocks"), JArray)
                        If blocksArray IsNot Nothing Then
                            For Each block In blocksArray.OfType(Of JObject)()
                                If block.Value(Of String)("type") = "table" Then
                                    ' Link dictionary setup
                                    Dim id = block.Value(Of Integer)("id")
                                    Dim rangeA = block.Value(Of String)("excelRange")
                                    Dim rangeB = block.Value(Of String)("excelRangeB")

                                    If Not String.IsNullOrWhiteSpace(rangeA) Then _linkedTables(id.ToString() & "_A") = rangeA
                                    If Not String.IsNullOrWhiteSpace(rangeB) Then _linkedTables(id.ToString() & "_B") = rangeB

                                    ' Collect unique paths (only add to list if they actually exist on disk)
                                    Dim pathA = block.Value(Of String)("excelPath")
                                    Dim pathB = block.Value(Of String)("excelPathB")
                                    Dim pathRes = block.Value(Of String)("excelResultPath")

                                    If Not String.IsNullOrWhiteSpace(pathA) AndAlso File.Exists(pathA) Then uniquePaths.Add(pathA)
                                    If Not String.IsNullOrWhiteSpace(pathB) AndAlso File.Exists(pathB) Then uniquePaths.Add(pathB)
                                    If Not String.IsNullOrWhiteSpace(pathRes) AndAlso File.Exists(pathRes) Then uniquePaths.Add(pathRes)
                                End If
                            Next
                        End If
                    Next
                End If

                ' 2. Pre-open all valid unique Excel files safely before syncing
                If uniquePaths.Count > 0 Then
                    EnsureExcelConnected()
                    If _excelApp IsNot Nothing Then
                        'ComRetryAction(Sub() _excelApp.Visible = True)

                        For Each p In uniquePaths
                            ' Simulate the user clicking the "Open" button for each unique file
                            HandleExcelOpen(p)
                            ' Give Excel half a second to process the file open before moving to the next
                            ' System.Threading.Thread.Sleep(500)
                            ' Application.DoEvents()
                        Next

                        ' Disable Excel Events temporarily to rapidly inject data
                        Dim prevAlerts As Boolean = ComRetry(Function() _excelApp.DisplayAlerts)
                        Dim prevEvents As Boolean = ComRetry(Function() _excelApp.EnableEvents)
                        ComRetryAction(Sub() _excelApp.DisplayAlerts = False)
                        ComRetryAction(Sub() _excelApp.EnableEvents = False)

                        If pagesArray IsNot Nothing Then
                            For Each page In pagesArray.OfType(Of JObject)()
                                Dim blocksArray = TryCast(page("blocks"), JArray)
                                If blocksArray IsNot Nothing Then
                                    For Each block In blocksArray.OfType(Of JObject)()
                                        If block.Value(Of String)("type") = "table" Then
                                            SyncBlockFromExcel(block, "excelRange", "excelPath", "data", "dataRows", "dataCols", "A")
                                            SyncBlockFromExcel(block, "excelRangeB", "excelPathB", "other", "otherRows", "otherCols", "B")

                                            ' Status verification for the Output/Result Link
                                            Dim resRange = block.Value(Of String)("excelResultRange")
                                            Dim resPath = block.Value(Of String)("excelResultPath")
                                            If Not String.IsNullOrWhiteSpace(resRange) Then
                                                If Not String.IsNullOrWhiteSpace(resPath) AndAlso Not File.Exists(resPath) Then
                                                    block("excelResultStatus") = "broken"
                                                Else
                                                    block("excelResultStatus") = "connected"
                                                End If
                                            End If
                                        End If
                                    Next
                                End If
                            Next
                        End If

                        ComRetryAction(Sub() _excelApp.DisplayAlerts = prevAlerts)
                        ComRetryAction(Sub() _excelApp.EnableEvents = prevEvents)
                    End If
                Else
                    ' If there are no valid unique paths (e.g. they are ALL broken), we still need to run through and flag them
                    If pagesArray IsNot Nothing Then
                        For Each page In pagesArray.OfType(Of JObject)()
                            Dim blocksArray = TryCast(page("blocks"), JArray)
                            If blocksArray IsNot Nothing Then
                                For Each block In blocksArray.OfType(Of JObject)()
                                    If block.Value(Of String)("type") = "table" Then
                                        SyncBlockFromExcel(block, "excelRange", "excelPath", "data", "dataRows", "dataCols", "A")
                                        SyncBlockFromExcel(block, "excelRangeB", "excelPathB", "other", "otherRows", "otherCols", "B")

                                        Dim resRange = block.Value(Of String)("excelResultRange")
                                        Dim resPath = block.Value(Of String)("excelResultPath")
                                        If Not String.IsNullOrWhiteSpace(resRange) Then
                                            If Not String.IsNullOrWhiteSpace(resPath) AndAlso Not File.Exists(resPath) Then
                                                block("excelResultStatus") = "broken"
                                            Else
                                                block("excelResultStatus") = "connected"
                                            End If
                                        End If
                                    End If
                                Next
                            End If
                        Next
                    End If
                End If

                ' 4. Load the updated document into the web view
                Dim script = "window.loadDocument(" & document.ToString(Formatting.None) & ");"
                wbMath.CoreWebView2.ExecuteScriptAsync(script)
            Catch ex As Exception
                MessageBox.Show("The selected canvas could not be loaded." & Environment.NewLine & ex.Message,
                                "Math Canvas", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Using
    End Sub

    ' =========================================================
    ' TEKLA STRUCTURAL DESIGNER (TSD) MODEL BRIDGE
    ' Exposes IMember / IMemberSpan / slab / wall geometry,
    ' properties, reinforcement and analysis results to the
    ' Math Canvas so they can be used inside expressions.
    ' =========================================================
    Private _tsdApp As TSD.API.Remoting.IApplication
    Private _tsdDocument As TSD.API.Remoting.Document.IDocument
    Private _tsdModel As TSD.API.Remoting.Structure.IModel

    ''' <summary>Index of the running TSD instance the canvas talks to.</summary>
    Public Property TsdInstanceIndex As Integer = 0

    ''' <summary>True once the user has explicitly chosen an instance (multi-instance case).</summary>
    Private _tsdInstanceChosen As Boolean = False

    ''' <summary>Remembers the includeRebar flag while a fetch waits for instance selection.</summary>
    Private _tsdPendingIncludeRebar As Boolean = False

    Private Sub TsdStatus(text As String)
        If wbMath Is Nothing OrElse wbMath.CoreWebView2 Is Nothing Then Return
        Dim jsText = JsonConvert.SerializeObject(text)
        wbMath.CoreWebView2.ExecuteScriptAsync($"if(typeof window.tsdStatus === 'function') window.tsdStatus({jsText});")
    End Sub

    ''' <summary>Sends the list of running TSD instances to the canvas so the user can pick one.</summary>
    Private Async Function SendTsdInstanceChooserAsync(instances As IReadOnlyList(Of TSD.API.Remoting.IApplication)) As Task
        Dim arr As New JArray()
        For i As Integer = 0 To instances.Count - 1
            Dim o As New JObject()
            o("index") = i
            Dim title As String
            Try
                title = Await instances(i).GetApplicationTitleAsync()
            Catch
                title = "Tekla Structural Designer instance " & (i + 1)
            End Try
            o("title") = title
            arr.Add(o)
        Next
        If wbMath Is Nothing OrElse wbMath.CoreWebView2 Is Nothing Then Return
        Dim js = JsonConvert.SerializeObject(arr.ToString(Formatting.None))
        wbMath.CoreWebView2.ExecuteScriptAsync($"if(typeof window.tsdChooseInstance === 'function') window.tsdChooseInstance({js});")
    End Function

    ''' <summary>
    ''' Acquires (or re-acquires) the TSD application, document and model.
    ''' When multiple instances are running and none has been chosen yet, prompts the user
    ''' and returns False (the fetch resumes from the tsdSelectInstance message).
    ''' </summary>
    Private Async Function EnsureTsdModelAsync(Optional forceReconnect As Boolean = False) As Task(Of Boolean)
        If forceReconnect Then
            _tsdModel = Nothing
            _tsdDocument = Nothing
            _tsdApp = Nothing
        End If

        If _tsdModel IsNot Nothing Then Return True

        Dim instances = Await TSD.API.Remoting.ApplicationFactory.GetRunningApplicationsAsync()
        If instances Is Nothing OrElse instances.Count = 0 Then
            TsdStatus("No running Tekla Structural Designer instance was found.")
            Return False
        End If

        ' More than one instance and the user has not chosen yet: ask them.
        If instances.Count > 1 AndAlso Not _tsdInstanceChosen Then
            Await SendTsdInstanceChooserAsync(instances)
            Return False
        End If

        Dim index As Integer = TsdInstanceIndex
        If index < 0 Then index = 0
        If index > instances.Count - 1 Then index = instances.Count - 1

        _tsdApp = instances(index)
        _tsdDocument = Await _tsdApp.GetDocumentAsync()
        _tsdModel = Await _tsdDocument.GetModelAsync()
        Return _tsdModel IsNot Nothing
    End Function

    Private Async Sub HandleTsdConnect()
        Try
            _tsdModel = Nothing
            _tsdDocument = Nothing
            _tsdApp = Nothing
            If Not Await EnsureTsdModelAsync() Then Return
            Dim title = Await _tsdApp.GetApplicationTitleAsync()
            TsdStatus("Connected to " & title & ". Loading model data...")
            HandleTsdFetchModel(False)
        Catch ex As Exception
            TsdStatus("TSD connection failed: " & ex.Message)
        End Try
    End Sub

    Private Async Sub HandleTsdFetchModel(includeRebar As Boolean)
        _tsdPendingIncludeRebar = includeRebar
        Dim initialError As Exception = Nothing

        Try
            If Not Await EnsureTsdModelAsync() Then Return
            Await SendTsdModelAsync(includeRebar)
        Catch ex As Exception
            initialError = ex
        End Try

        If initialError Is Nothing Then Return

        ' The cached document/model is most likely stale because the user reopened a
        ' different TSD model. Drop the references and reconnect once before failing.
        Try
            If Not Await EnsureTsdModelAsync(forceReconnect:=True) Then Return
            Await SendTsdModelAsync(includeRebar)
        Catch retryError As Exception
            TsdStatus("TSD model could not be read: " & retryError.Message)
        End Try
    End Sub

    Private Async Function SendTsdModelAsync(includeRebar As Boolean) As Task
        Dim model = Await BuildTsdModelJsonAsync(includeRebar)
        If wbMath Is Nothing OrElse wbMath.CoreWebView2 Is Nothing Then Return
        Dim payload = model.ToString(Formatting.None)
        wbMath.CoreWebView2.ExecuteScriptAsync($"if(typeof window.tsdReceiveModel === 'function') window.tsdReceiveModel({payload});")
    End Function

    Private Async Sub HandleTsdFetchResult(key As String, name As String, resultType As String, spanIndex As Integer, position As String)
        Dim payload As New JObject()
        payload("key") = key
        payload("name") = name
        payload("resultType") = resultType
        payload("span") = spanIndex
        payload("position") = position

        Try
            If Not Await EnsureTsdModelAsync() Then Return

            Dim CF As New CallFunctions()
            Dim resultOutput As List(Of List(Of DesignForces)) = Nothing
            Dim pos = If(position, "MAX").Trim().ToUpperInvariant()

            Dim allMembers = (Await _tsdModel.GetMembersAsync()).ToList()
            Dim targetMember = allMembers.FirstOrDefault(Function(m) String.Equals(m.Name, name, StringComparison.OrdinalIgnoreCase))

            If targetMember IsNot Nothing Then
                If pos = "MAX" Then
                    resultOutput = Await CF.GetMemberResultsAtMaxPoint(_tsdModel, name, resultType, spanIndex)
                ElseIf pos = "MIN" Then
                    resultOutput = Await CF.GetMemberResultsAtMinPoint(_tsdModel, name, resultType, spanIndex)
                ElseIf IsNumeric(pos) Then
                    resultOutput = Await CF.GetMemberResults(_tsdModel, name, resultType, spanIndex, CDbl(pos))
                End If
            Else
                Dim allWalls = (Await _tsdModel.GetStructuralWallsAsync()).ToList()
                Dim targetWall = allWalls.FirstOrDefault(Function(w) String.Equals(w.Name, name, StringComparison.OrdinalIgnoreCase))
                If targetWall IsNot Nothing Then
                    resultOutput = Await CF.GetWallResults(_tsdModel, name, resultType, spanIndex, pos)
                End If
            End If

            If resultOutput IsNot Nothing AndAlso resultOutput.Count >= 2 AndAlso
               resultOutput(0).Count > 0 AndAlso resultOutput(1).Count > 0 Then
                Dim minResult = resultOutput(0)(0)
                Dim maxResult = resultOutput(1)(0)
                payload("min") = minResult.Value
                payload("max") = maxResult.Value
                payload("value") = If(System.Math.Abs(maxResult.Value) >= System.Math.Abs(minResult.Value), maxResult.Value, minResult.Value)
                payload("minCombination") = minResult.FinalCombinationName
                payload("maxCombination") = maxResult.FinalCombinationName
                payload("minSolverModel") = minResult.FinalSolverModelName
                payload("maxSolverModel") = maxResult.FinalSolverModelName
            Else
                payload("error") = "No results available for '" & name & "'."
            End If
        Catch ex As Exception
            payload("error") = ex.Message
        End Try

        If wbMath Is Nothing OrElse wbMath.CoreWebView2 Is Nothing Then Return
        wbMath.CoreWebView2.ExecuteScriptAsync($"if(typeof window.tsdReceiveResult === 'function') window.tsdReceiveResult({payload.ToString(Formatting.None)});")
    End Sub

    Private Async Function BuildTsdModelJsonAsync(includeRebar As Boolean) As Task(Of JObject)
        Dim root As New JObject()
        root("generated") = DateTime.Now.ToString("s")
        root("includeRebar") = includeRebar

        Try
            root("title") = Await _tsdApp.GetApplicationTitleAsync()
        Catch
        End Try

        Dim constructionPoints = (Await _tsdModel.GetConstructionPointsAsync()).ToList()
        Dim pointByIndex As New Dictionary(Of Integer, TSD.API.Remoting.Structure.IConstructionPoint)()
        For Each cp In constructionPoints
            If Not pointByIndex.ContainsKey(cp.Index) Then pointByIndex.Add(cp.Index, cp)
        Next

        root("members") = Await BuildTsdMembersAsync(pointByIndex, includeRebar)
        root("slabs") = Await BuildTsdSlabsAsync(pointByIndex, includeRebar)
        root("walls") = Await BuildTsdWallsAsync(includeRebar)
        root("levels") = Await BuildTsdLevelsAsync()
        root("loadcases") = Await BuildTsdLoadcasesAsync()
        root("combinations") = Await BuildTsdCombinationsAsync()
        root("grids") = Await BuildTsdGridsAsync()
        root("foundations") = Await BuildTsdFoundationsAsync(includeRebar)
        root("piles") = Await BuildTsdPilesAsync()
        Return root
    End Function

    ''' <summary>
    ''' Pad bases and pile caps published as one collection. Both are isolated
    ''' foundations, so they share the geometry, cover and material properties of
    ''' IIsolatedFoundationData; the pile cap rows add the pile arrangement, and the
    ''' design side is reported through the foundation check results (status and
    ''' utilization ratio) exactly as the connector foundation forms read them.
    ''' </summary>
    Private Async Function BuildTsdFoundationsAsync(includeRebar As Boolean) As Task(Of JArray)
        Dim foundations As New JArray()

        Try
            Dim padBases = (Await _tsdModel.GetPadBasesAsync()).ToList()
            For index = 0 To padBases.Count - 1
                foundations.Add(Await TsdFoundationObjectAsync(padBases(index), "PadBase", index + 1, includeRebar))
            Next
        Catch ex As Exception
            foundations.Add(New JObject(New JProperty("error", "Pad bases could not be read: " & ex.Message)))
        End Try

        Try
            Dim pileCaps = (Await _tsdModel.GetPileCapsAsync()).ToList()
            For index = 0 To pileCaps.Count - 1
                foundations.Add(Await TsdFoundationObjectAsync(pileCaps(index), "PileCap", index + 1, includeRebar))
            Next
        Catch ex As Exception
            foundations.Add(New JObject(New JProperty("error", "Pile caps could not be read: " & ex.Message)))
        End Try

        Return foundations
    End Function

    Private Async Function TsdFoundationObjectAsync(foundation As TSD.API.Remoting.Foundations.IIsolatedFoundation,
                                                    kind As String,
                                                    ordinal As Integer,
                                                    includeRebar As Boolean) As Task(Of JObject)
        Dim fo As New JObject()
        fo("kind") = kind
        ' Isolated foundations do not publish a name, so the schedule uses the same
        ' "PadBase 1" / "PileCap 1" labels the connector foundation forms display.
        fo("name") = kind & " " & ordinal.ToString()
        fo("index") = ordinal
        ' SupportedMemberType is published on the foundation and again on its data, so
        ' whichever of the two the model exposes ends up in the schedule.
        TsdTrySet(fo, "supportedMemberType", Function() TsdReadPropertyText(foundation, "SupportedMemberType"))

        Dim data = TsdUnwrap(TsdReadProperty(foundation, "IsolatedFoundationData"))
        ' The plan size, depth and concrete grade are read through the strongly typed
        ' IIsolatedFoundationData contract, exactly as the connector foundation export
        ' does (Footings(i).IsolatedFoundationData.Value.LengthDir1.Value). Reflection is
        ' only used for the remaining optional members, which are not published by every
        ' foundation kind.
        Try
            Dim typedData As TSD.API.Remoting.Foundations.IIsolatedFoundationData = foundation.IsolatedFoundationData.Value
            If typedData IsNot Nothing Then
                TsdTrySet(fo, "lengthDir1", Function() CObj(Convert.ToDouble(typedData.LengthDir1.Value)))
                TsdTrySet(fo, "lengthDir2", Function() CObj(Convert.ToDouble(typedData.LengthDir2.Value)))
                TsdTrySet(fo, "depth", Function() CObj(Convert.ToDouble(typedData.Depth.Value)))
                TsdTrySet(fo, "concrete", Function() typedData.Concrete.Value.Name.ToString())
                If data Is Nothing Then data = typedData
            End If
        Catch ex As Exception
            fo("geometryError") = ex.Message
        End Try

        If data IsNot Nothing Then
            ' Geometry.
            TsdTrySet(fo, "lengthDir1", Function() TsdReadProperty(data, "LengthDir1"))
            TsdTrySet(fo, "lengthDir2", Function() TsdReadProperty(data, "LengthDir2"))
            TsdTrySet(fo, "depth", Function() TsdReadProperty(data, "Depth"))
            TsdTrySet(fo, "rotation", Function() TsdReadProperty(data, "Rotation"))
            TsdTrySet(fo, "eccentricityDir1", Function() TsdReadProperty(data, "EccentricityDir1"))
            TsdTrySet(fo, "eccentricityDir2", Function() TsdReadProperty(data, "EccentricityDir2"))
            TsdTrySet(fo, "pedestalHeight", Function() TsdReadProperty(data, "PedestalHeight"))
            TsdTrySet(fo, "pedestalLengthDir1", Function() TsdReadProperty(data, "PedestalLengthDir1"))
            TsdTrySet(fo, "pedestalLengthDir2", Function() TsdReadProperty(data, "PedestalLengthDir2"))
            TsdTrySet(fo, "surfaceArea", Function() TsdReadProperty(data, "SurfaceArea"))
            TsdTrySet(fo, "volume", Function() TsdReadProperty(data, "Volume"))
            TsdTrySet(fo, "topCover", Function() TsdReadProperty(data, "TopCover"))
            TsdTrySet(fo, "bottomCover", Function() TsdReadProperty(data, "BottomCover"))
            TsdTrySet(fo, "sideCover", Function() TsdReadProperty(data, "SideCover"))
            TsdTrySet(fo, "concrete", Function() TsdMaterialName(TsdReadProperty(data, "Concrete")))
            TsdTrySet(fo, "foundationType", Function() TsdReadPropertyText(data, "FoundationType"))
            TsdTrySet(fo, "shape", Function() TsdReadPropertyText(data, "IsolatedFoundationShape"))
            TsdTrySet(fo, "geometry", Function() TsdReadPropertyText(data, "IsolatedFoundationGeometry"))
            If fo("supportedMemberType") Is Nothing Then
                TsdTrySet(fo, "supportedMemberType", Function() TsdReadPropertyText(data, "SupportedMemberType"))
            End If

            ' Soil / bearing design data used by the geotechnical checks.
            TsdTrySet(fo, "soilUnitWeight", Function() TsdReadProperty(data, "SoilUnitWeight"))
            TsdTrySet(fo, "frictionAngle", Function() TsdReadProperty(data, "FrictionAngle"))
            TsdTrySet(fo, "surchargeDepth", Function() TsdReadProperty(data, "SurchargeDepth"))
            TsdTrySet(fo, "surchargeLoadPermanent", Function() TsdReadProperty(data, "SurchargeLoadPermanent"))
            TsdTrySet(fo, "surchargeLoadVariable", Function() TsdReadProperty(data, "SurchargeLoadVariable"))
            TsdTrySet(fo, "utilizationRatioLimit", Function() TsdReadProperty(data, "UtilizationRatioLimit"))
            TsdTrySet(fo, "bearingCapacity1", Function() TsdReadProperty(data, "AllowableBearingCapacity1"))
            TsdTrySet(fo, "bearingCapacity2", Function() TsdReadProperty(data, "AllowableBearingCapacity2"))

            ' Pile cap arrangement.
            TsdTrySet(fo, "pileCount", Function() TsdReadProperty(data, "NumberOfPiles"))
            TsdTrySet(fo, "pileRows", Function() TsdReadProperty(data, "PileRows"))
            TsdTrySet(fo, "pilesPerRow", Function() TsdReadProperty(data, "PilesPerRow"))
            TsdTrySet(fo, "pileSpacing", Function() TsdReadProperty(data, "PileSpacing"))
            TsdTrySet(fo, "pileArrangement", Function() TsdReadPropertyText(data, "PileArrangementDescription"))
            TsdTrySet(fo, "pileType", Function() TsdMaterialName(TsdReadProperty(data, "PileType")))
        End If

        ' Design side: the check results carry the status and the governing utilization
        ' ratio. CheckResults is a dictionary keyed by the check type, and every entry is
        ' itself a property wrapper, so one result is reached as
        ' CheckResults.Value(kind).Value and its ratio as .UtilizationRatio.Value.
        Dim checks As New JArray()
        Try
            Dim checkResults = foundation.CheckResults.Value
            If checkResults IsNot Nothing Then
                Dim worst As Double = 0
                Dim worstStatus As String = Nothing
                Dim any As Boolean = False
                For Each pair In checkResults
                    Dim check As TSD.API.Remoting.Common.ICheckResult = Nothing
                    Try
                        If pair.Value IsNot Nothing Then check = pair.Value.Value
                    Catch
                    End Try
                    If check Is Nothing Then Continue For
                    Dim co As New JObject()
                    co("check") = pair.Key.ToString()
                    TsdTrySet(co, "status", Function() TsdReadPropertyText(check, "CheckStatus"))
                    Try
                        Dim ratio As Double = Convert.ToDouble(TsdUnwrap(check.UtilizationRatio))
                        co("utilizationRatio") = ratio
                        any = True
                        If ratio > worst Then
                            worst = ratio
                            worstStatus = If(co("status") Is Nothing, Nothing, co("status").ToString())
                        End If
                    Catch
                    End Try
                    checks.Add(co)
                Next
                If any Then
                    fo("utilizationRatio") = worst
                    If worstStatus IsNot Nothing Then fo("checkStatus") = worstStatus
                End If
            End If
        Catch ex As Exception
            fo("checkError") = ex.Message
        End Try
        fo("checks") = checks

        If includeRebar Then
            Try
                fo("rebarTop") = TsdSlabRebarRows(Await foundation.GetTopReinforcementAsync())
            Catch ex As Exception
                fo("rebarTop") = JValue.CreateString("Top reinforcement unavailable: " & ex.Message)
            End Try
            Try
                fo("rebarBottom") = TsdSlabRebarRows(Await foundation.GetBottomReinforcementAsync())
            Catch ex As Exception
                fo("rebarBottom") = JValue.CreateString("Bottom reinforcement unavailable: " & ex.Message)
            End Try
        End If

        Return fo
    End Function

    ''' <summary>
    ''' Piles with their plan position, pile type geometry and the axial / lateral
    ''' resistances and spring stiffnesses used for the pile working load checks.
    ''' </summary>
    Private Async Function BuildTsdPilesAsync() As Task(Of JArray)
        Dim piles As New JArray()
        Try
            Dim allPiles = (Await _tsdModel.GetPilesAsync()).ToList()
            For index = 0 To allPiles.Count - 1
                Dim pile = allPiles(index)
                Dim po As New JObject()
                ' Piles are not named in the model, so they are labelled by position
                ' in the pile collection, as the pile working load form does.
                po("index") = index + 1
                po("name") = "P " & (index + 1).ToString()
                TsdTrySet(po, "isPileCapPile", Function() TsdReadProperty(pile, "IsPileCapPile"))
                TsdTrySet(po, "isMatPile", Function() TsdReadProperty(pile, "IsMatPile"))
                TsdTrySet(po, "indexInPileCap", Function() TsdReadProperty(pile, "IndexInPileCap"))

                Try
                    Dim position = pile.GlobalPosition.Value
                    po("x") = position.X
                    po("y") = position.Y
                    po("z") = position.Z
                Catch
                End Try

                Dim pileType = TsdUnwrap(TsdReadProperty(pile, "PileType"))
                If pileType IsNot Nothing Then
                    ' IPileType labels itself through Description (shape and size); Name is
                    ' only present when the pile type comes from the pile type library.
                    TsdTrySet(po, "pileType", Function() TsdMaterialName(pileType))
                    TsdTrySet(po, "description", Function() TsdReadProperty(pileType, "Description"))
                    TsdTrySet(po, "shape", Function() TsdReadPropertyText(pileType, "PileTypeShape"))
                    TsdTrySet(po, "installation", Function() TsdReadPropertyText(pileType, "PileTypeInstallationType"))
                    TsdTrySet(po, "loadTransfer", Function() TsdReadPropertyText(pileType, "LoadTransferType"))
                    TsdTrySet(po, "linearity", Function() TsdReadPropertyText(pileType, "Linearity"))
                    TsdTrySet(po, "dimension", Function() TsdReadProperty(pileType, "Dimension"))
                    TsdTrySet(po, "length", Function() TsdReadProperty(pileType, "Length"))
                    TsdTrySet(po, "embedment", Function() TsdReadProperty(pileType, "Embedment"))
                    TsdTrySet(po, "area", Function() TsdReadProperty(pileType, "CrossSectionalArea"))
                    TsdTrySet(po, "material", Function() TsdMaterialName(TsdReadProperty(pileType, "Material")))
                    TsdTrySet(po, "materialType", Function() TsdReadPropertyText(pileType, "MaterialType"))
                    ' TSD publishes the pile resistances and stiffness limits in newtons,
                    ' whereas the rest of the bridge (and the design forms) work in kN, so
                    ' the force values are scaled once here.
                    TsdTrySet(po, "compressiveResistance", Function() TsdForceKn(TsdReadProperty(pileType, "AxialCompressiveResistance")))
                    TsdTrySet(po, "compressiveResistanceStr", Function() TsdForceKn(TsdReadProperty(pileType, "AxialCompressiveResistanceEcStr")))
                    TsdTrySet(po, "tensileResistance", Function() TsdForceKn(TsdReadProperty(pileType, "AxialTensileResistance")))
                    TsdTrySet(po, "tensileResistanceStr", Function() TsdForceKn(TsdReadProperty(pileType, "AxialTensileResistanceEcStr")))
                    TsdTrySet(po, "lateralResistance", Function() TsdForceKn(TsdReadProperty(pileType, "LateralResistance")))
                    TsdTrySet(po, "lateralResistanceStr", Function() TsdForceKn(TsdReadProperty(pileType, "LateralResistanceStr")))
                    TsdTrySet(po, "compressionStiffness", Function() TsdReadProperty(pileType, "CompressionStiffnessVertical"))
                    TsdTrySet(po, "tensionStiffness", Function() TsdReadProperty(pileType, "TensionStiffnessVertical"))
                    TsdTrySet(po, "compressionLimit", Function() TsdForceKn(TsdReadProperty(pileType, "CompressionLimitVertical")))
                    TsdTrySet(po, "tensionLimit", Function() TsdForceKn(TsdReadProperty(pileType, "TensionLimitVertical")))
                    TsdTrySet(po, "horizontalStiffness", Function() TsdReadProperty(pileType, "HorizontalStiffness"))
                    TsdTrySet(po, "horizontalRestraint", Function() TsdReadPropertyText(pileType, "HorizontalRestraint"))
                End If

                piles.Add(po)
            Next
        Catch ex As Exception
            piles.Add(New JObject(New JProperty("error", ex.Message)))
        End Try
        Return piles
    End Function

    Private Async Function BuildTsdMembersAsync(pointByIndex As Dictionary(Of Integer, TSD.API.Remoting.Structure.IConstructionPoint),
                                                includeRebar As Boolean) As Task(Of JArray)
        Dim members As New JArray()
        Dim allMembers = (Await _tsdModel.GetMembersAsync()).ToList()

        For Each member In allMembers
            Dim mo As New JObject()
            mo("name") = member.Name
            TsdTrySet(mo, "type", Function() member.Data.Value.MemberType.Value.ToString())
            TsdTrySet(mo, "construction", Function() member.Data.Value.Construction.Value.ToString())
            TsdTrySet(mo, "fabrication", Function() member.Data.Value.Fabrication.Value.ToString())
            TsdTrySet(mo, "spanCount", Function() CObj(member.SpanCount.Value))

            Dim rebarToken As JToken = Nothing
            If includeRebar Then
                Try
                    Dim reinforcement = Await member.GetReinforcementAsync()
                    rebarToken = TsdMemberRebarRows(reinforcement)
                Catch ex As Exception
                    rebarToken = JValue.CreateString("Reinforcement unavailable: " & ex.Message)
                End Try
            End If

            Dim spans As New JArray()
            Dim totalLength As Double = 0
            Try
                Dim memberSpans = (Await member.GetSpanAsync()).ToList()
                For i As Integer = 0 To memberSpans.Count - 1
                    Dim span = memberSpans(i)
                    Dim so As New JObject()
                    so("index") = i
                    so("name") = span.Name

                    Dim startPoint As TSD.API.Remoting.Structure.IConstructionPoint = Nothing
                    Dim endPoint As TSD.API.Remoting.Structure.IConstructionPoint = Nothing
                    Try
                        pointByIndex.TryGetValue(span.StartMemberNode.ConstructionPointIndex.Value, startPoint)
                        pointByIndex.TryGetValue(span.EndMemberNode.ConstructionPointIndex.Value, endPoint)
                    Catch
                    End Try

                    If startPoint IsNot Nothing AndAlso endPoint IsNot Nothing Then
                        Dim sc = startPoint.Coordinates.Value
                        Dim ec = endPoint.Coordinates.Value
                        so("start") = TsdPoint(sc.X, sc.Y, sc.Z)
                        so("end") = TsdPoint(ec.X, ec.Y, ec.Z)
                        Dim dx = ec.X - sc.X, dy = ec.Y - sc.Y, dz = ec.Z - sc.Z
                        Dim length = System.Math.Sqrt(dx * dx + dy * dy + dz * dz)
                        so("length") = length
                        totalLength += length
                    End If

                    Try
                        Dim elementSection = span.ElementSection.Value
                        so("section") = elementSection.ToString()
                        Dim memberSection = TryCast(elementSection, TSD.API.Remoting.Sections.IMemberSection)
                        If memberSection IsNot Nothing Then
                            Dim physical = memberSection.PhysicalSection.Value
                            so("sectionName") = physical.LongName
                            Dim rect = TryCast(physical, TSD.API.Remoting.Sections.IParametricRectangularSection)
                            If rect IsNot Nothing Then
                                so("breadth") = Convert.ToDouble(rect.Breadth)
                                so("depth") = Convert.ToDouble(rect.Depth)
                                so("area") = Convert.ToDouble(rect.Breadth) * Convert.ToDouble(rect.Depth)
                            End If
                            ' Steel (and other non-rectangular) sections publish their rolled
                            ' section dimensions and analysis properties on the section object,
                            ' so they are merged in without overwriting the concrete values above.
                            Dim sectionProperties = TsdSectionProperties(physical)
                            For Each sectionProperty In sectionProperties.Properties()
                                If so(sectionProperty.Name) Is Nothing Then so(sectionProperty.Name) = sectionProperty.Value
                            Next
                        End If
                    Catch
                    End Try

                    ' End fixity (pinned / fixed / spring) of both ends of the span. The
                    ' releases are dereferenced here, as ISpanReleases lives behind the
                    ' IProperty wrapper: span.StartReleases.Value.DegreeOfFreedom.Value.
                    Try
                        so("startFixityData") = TsdReleaseObject(span.StartReleases.Value, "Start")
                        so("endFixityData") = TsdReleaseObject(span.EndReleases.Value, "End")
                        so("startFixity") = so("startFixityData")("fixity")
                        so("endFixity") = so("endFixityData")("fixity")
                    Catch ex As Exception
                        so("fixityError") = ex.Message
                    End Try

                    spans.Add(so)
                Next
            Catch ex As Exception
                mo("spanError") = ex.Message
            End Try

            mo("spans") = spans
            mo("length") = totalLength
            If rebarToken IsNot Nothing Then mo("rebar") = rebarToken
            members.Add(mo)
        Next

        Return members
    End Function

    Private Async Function BuildTsdSlabsAsync(pointByIndex As Dictionary(Of Integer, TSD.API.Remoting.Structure.IConstructionPoint),
                                              includeRebar As Boolean) As Task(Of JArray)
        Dim slabs As New JArray()
        Try
            Dim slabItems = (Await _tsdModel.GetSlabItemsAsync()).ToList()
            For Each item In slabItems
                Dim so As New JObject()
                so("name") = item.Name
                TsdTrySet(so, "depth", Function() CObj(Convert.ToDouble(item.SlabItemData.Value.Depth.Value)))

                If includeRebar Then
                    Try
                        so("rebarTop") = TsdSlabRebarRows(Await item.GetTopReinforcementAsync())
                    Catch ex As Exception
                        so("rebarTop") = JValue.CreateString("Top reinforcement unavailable: " & ex.Message)
                    End Try
                    Try
                        so("rebarBottom") = TsdSlabRebarRows(Await item.GetBottomReinforcementAsync())
                    Catch ex As Exception
                        so("rebarBottom") = JValue.CreateString("Bottom reinforcement unavailable: " & ex.Message)
                    End Try
                End If

                Dim boundary As New JArray()
                Try
                    For Each pointIndex In item.ConstructionPointIndices
                        Dim cp As TSD.API.Remoting.Structure.IConstructionPoint = Nothing
                        If pointByIndex.TryGetValue(pointIndex, cp) Then
                            Dim c = cp.Coordinates.Value
                            boundary.Add(TsdPoint(c.X, c.Y, c.Z))
                        End If
                    Next
                Catch
                End Try
                so("boundary") = boundary
                so("pointCount") = boundary.Count
                so("area") = TsdPolygonArea(boundary)
                slabs.Add(so)
            Next
        Catch ex As Exception
            slabs.Add(New JObject(New JProperty("error", ex.Message)))
        End Try
        Return slabs
    End Function

    ''' <summary>
    ''' Plan area of a slab boundary using the shoelace formula on the X/Y coordinates.
    ''' </summary>
    Private Function TsdPolygonArea(boundary As JArray) As Double
        If boundary Is Nothing OrElse boundary.Count < 3 Then Return 0.0
        Dim total As Double = 0.0
        For i As Integer = 0 To boundary.Count - 1
            Dim a = TryCast(boundary(i), JObject)
            Dim b = TryCast(boundary((i + 1) Mod boundary.Count), JObject)
            If a Is Nothing OrElse b Is Nothing Then Return 0.0
            Dim ax = a.Value(Of Double)("x"), ay = a.Value(Of Double)("y")
            Dim bx = b.Value(Of Double)("x"), by = b.Value(Of Double)("y")
            total += (ax * by) - (bx * ay)
        Next
        Return System.Math.Abs(total) / 2.0
    End Function

    Private Async Function BuildTsdWallsAsync(includeRebar As Boolean) As Task(Of JArray)
        Dim walls As New JArray()
        Try
            Dim structuralWalls = (Await _tsdModel.GetStructuralWallsAsync()).ToList()
            For Each wall In structuralWalls
                Dim wo As New JObject()
                wo("name") = wall.Name
                wo("kind") = "Structural"

                If includeRebar Then
                    Try
                        wo("rebar") = TsdWallRebarRows(Await wall.GetReinforcementAsync())
                    Catch ex As Exception
                        wo("rebar") = JValue.CreateString("Reinforcement unavailable: " & ex.Message)
                    End Try
                End If

                Dim panels As New JArray()
                Try
                    Dim wallPanels = (Await wall.GetSpanAsync()).ToList()
                    For i As Integer = 0 To wallPanels.Count - 1
                        Dim panel = wallPanels(i)
                        Dim po As New JObject()
                        po("index") = i
                        TsdTrySet(po, "thickness", Function() CObj(Convert.ToDouble(panel.WallPanelData.Value.Thickness.Value)))
                        Try
                            Dim sp = panel.BottomSegment.Value.GetPoint(TSD.API.Remoting.Geometry.Location.Start)
                            Dim ep = panel.BottomSegment.Value.GetPoint(TSD.API.Remoting.Geometry.Location.End)
                            po("start") = TsdPoint(sp.X, sp.Y, sp.Z)
                            po("end") = TsdPoint(ep.X, ep.Y, ep.Z)
                            Dim dx = ep.X - sp.X, dy = ep.Y - sp.Y, dz = ep.Z - sp.Z
                            po("length") = System.Math.Sqrt(dx * dx + dy * dy + dz * dz)
                        Catch
                        End Try
                        TsdTrySet(po, "height", Function()
                                                    Dim topZ = panel.TopSegment.Value.GetPoint(TSD.API.Remoting.Geometry.Location.Start).Z
                                                    Dim botZ = panel.BottomSegment.Value.GetPoint(TSD.API.Remoting.Geometry.Location.Start).Z
                                                    Return CObj(Convert.ToDouble(topZ - botZ))
                                                End Function)
                        panels.Add(po)
                    Next
                Catch ex As Exception
                    wo("panelError") = ex.Message
                End Try

                wo("panels") = panels
                wo("panelCount") = panels.Count
                walls.Add(wo)
            Next
        Catch ex As Exception
            walls.Add(New JObject(New JProperty("error", ex.Message)))
        End Try
        Return walls
    End Function

    Private Async Function BuildTsdLevelsAsync() As Task(Of JArray)
        Dim levels As New JArray()
        Try
            Dim allLevels = (Await _tsdModel.GetLevelsAsync()).ToList()
            For i As Integer = 0 To allLevels.Count - 1
                Dim level = allLevels(i)
                Dim lo As New JObject()
                lo("index") = i
                lo("name") = level.Name
                TsdTrySet(lo, "z", Function() CObj(Convert.ToDouble(level.Level.Value)))
                levels.Add(lo)
            Next
        Catch ex As Exception
            levels.Add(New JObject(New JProperty("error", ex.Message)))
        End Try
        Return levels
    End Function

    Private Async Function BuildTsdLoadcasesAsync() As Task(Of JArray)
        Dim loadcases As New JArray()
        Try
            Dim allCases = (Await _tsdModel.GetLoadcasesAsync()).ToList()
            For i As Integer = 0 To allCases.Count - 1
                Dim lo As New JObject()
                lo("index") = i
                lo("name") = allCases(i).Name
                TsdTrySet(lo, "type", Function() allCases(i).Type.Value.ToString())
                loadcases.Add(lo)
            Next
        Catch ex As Exception
            loadcases.Add(New JObject(New JProperty("error", ex.Message)))
        End Try
        Return loadcases
    End Function

    Private Async Function BuildTsdGridsAsync() As Task(Of JArray)
        Dim grids As New JArray()
        Try
            For Each grid In (Await _tsdModel.GetArchitecturalGridsAsync()).ToList()
                grids.Add(New JObject(New JProperty("name", grid.Name)))
            Next
        Catch ex As Exception
            grids.Add(New JObject(New JProperty("error", ex.Message)))
        End Try
        Return grids
    End Function

    ''' <summary>
    ''' Load combinations with the factors applied to each load case, mirroring the
    ''' Combination.LoadcaseFactors pattern used elsewhere in the connector.
    ''' </summary>
    Private Async Function BuildTsdCombinationsAsync() As Task(Of JArray)
        Dim combinations As New JArray()
        Try
            Dim allCases = (Await _tsdModel.GetLoadcasesAsync()).ToList()
            Dim caseNameById As New Dictionary(Of Guid, String)()
            For Each lc In allCases
                If Not caseNameById.ContainsKey(lc.Id) Then caseNameById.Add(lc.Id, lc.Name)
            Next

            Dim allCombinations = (Await _tsdModel.GetCombinationsAsync()).ToList()
            For i As Integer = 0 To allCombinations.Count - 1
                Dim comb = allCombinations(i)
                Dim co As New JObject()
                co("index") = i
                co("name") = comb.Name
                TsdTrySet(co, "class", Function() comb.CombinationClass.Value.ToString())
                TsdTrySet(co, "isStrength", Function() CObj(comb.IsStrength.Value))
                TsdTrySet(co, "isActive", Function() CObj(comb.IsActive.Value))
                TsdTrySet(co, "factoringType", Function() comb.FactoringType.Value.ToString())

                Dim factors As New JArray()
                Try
                    For Each factorProperty In comb.LoadcaseFactors
                        Dim factor = factorProperty.Value
                        Dim fo As New JObject()
                        Dim caseName As String = Nothing
                        Try
                            caseNameById.TryGetValue(factor.LoadcaseId.Value, caseName)
                        Catch
                        End Try
                        fo("combination") = comb.Name
                        fo("loadcase") = If(caseName, "")
                        TsdTrySet(fo, "strengthFactor", Function() CObj(Convert.ToDouble(factor.StrengthFactor.Value)))
                        TsdTrySet(fo, "serviceFactor", Function() CObj(Convert.ToDouble(factor.ServiceFactor.Value)))
                        TsdTrySet(fo, "serviceQuasiFactor", Function() CObj(Convert.ToDouble(factor.ServiceQuasiFactor.Value)))
                        factors.Add(fo)
                    Next
                Catch ex As Exception
                    co("factorError") = ex.Message
                End Try

                co("factors") = factors
                co("factorCount") = factors.Count
                combinations.Add(co)
            Next
        Catch ex As Exception
            combinations.Add(New JObject(New JProperty("error", ex.Message)))
        End Try
        Return combinations
    End Function

    ''' <summary>
    ''' Flattens beam / column reinforcement into "region / bars" rows using the
    ''' scheduling and detailing groups published by the reinforcement collection.
    ''' Beam longitudinal bars are reported in the six DCM regions (top/bottom by
    ''' left/mid/right), beam links in left/middle/right and column links in the
    ''' bottom/mid/top regions implied by the link detailing group type.
    ''' </summary>
    Private Shared Function TsdMemberRebarRows(reinforcement As TSD.API.Remoting.Reinforcement.IReinforcementCollection) As JArray
        Dim rows As New JArray()
        If reinforcement Is Nothing Then Return rows

        ' Beam style reinforcement: longitudinal bar scheduling groups reported as
        ' Top / Bottom x Left / Mid / Right (six regions).
        Try
            For Each groupProperty In reinforcement.LongitudinalBarSchedulingGroups.Value
                Dim group = groupProperty.Value
                Dim ro As New JObject()
                Dim zone As String = TsdReadText(Function() group.Zone.Value.ToString())
                Dim position As String = TsdBeamPositionText(TsdReadText(Function() group.Position.Value.ToString()))
                ro("kind") = "Longitudinal"
                ro("zone") = zone
                ro("position") = position
                ro("region") = TsdRegionText(zone, position)
                TsdTrySet(ro, "spanIndex", Function() CObj(group.EdgeIndex.Value))
                TsdTrySet(ro, "count", Function() CObj(group.Count))
                TsdTrySet(ro, "prefix", Function() group.DetailingPrefix.Value.ToString())
                TsdTrySet(ro, "size", Function() group.Size.Value.Size.ToString())
                TsdTrySet(ro, "grade", Function() group.Size.Value.Grade.Name)
                TsdSetZoneLengths(ro, group)
                ro("bars") = TsdBarText(ro)
                rows.Add(ro)
            Next
        Catch
        End Try

        ' Beam links: reported as Left / Middle / Right with the leg count in front.
        Try
            For Each groupProperty In reinforcement.LinkSchedulingGroups.Value
                Dim group = groupProperty.Value
                Dim ro As New JObject()
                Dim position As String = TsdBeamPositionText(TsdReadText(Function() group.Position.Value.ToString()))
                If position = "Mid" Then position = "Middle"
                ro("kind") = "Link"
                ro("position") = position
                ro("region") = position
                TsdTrySet(ro, "spanIndex", Function() CObj(group.EdgeIndex.Value))
                TsdTrySet(ro, "legs", Function() CObj(group.LegCount.Value))
                TsdTrySet(ro, "prefix", Function() group.DetailingPrefix.Value.ToString())
                TsdTrySet(ro, "size", Function() group.Size.Value.Size.ToString())
                TsdTrySet(ro, "spacing", Function() CObj(Convert.ToDouble(group.CentreSpacing.Value)))
                TsdTrySet(ro, "grade", Function() group.Size.Value.Grade.Name)
                TsdSetZoneLengths(ro, group)
                ro("bars") = TsdBarText(ro)
                rows.Add(ro)
            Next
        Catch
        End Try

        ' Column style reinforcement
        ' longitudinal bars are not spaced bars, so no "-spacing" suffix is shown.
        Try
            For Each groupProperty In reinforcement.LongitudinalBarGroups
                Dim group = groupProperty.Value
                Dim ro As New JObject()
                ro("kind") = "Longitudinal"
                ro("region") = "Stack"
                TsdTrySet(ro, "count", Function() CObj(group.Count))
                TsdTrySet(ro, "prefix", Function() group.DetailingPrefix.Value.ToString())
                TsdTrySet(ro, "size", Function() group.Size.Value.Size.ToString())
                TsdTrySet(ro, "grade", Function() group.Size.Value.Grade.Name)
                TsdTrySet(ro, "spanIndex", Function() CObj(group.Value(0).Value.StartSpanIndex.Value))
                TsdSetZoneLengths(ro, group)
                ro("bars") = TsdBarText(ro)
                rows.Add(ro)
            Next
        Catch
        End Try

        ' Column links
        ' detail both the bottom and the top region of the stack (as in the DCM
        ' column stirrup extraction). Bars are reported by link leg count.
        Try
            For Each groupProperty In reinforcement.LinkGroups
                Dim group = groupProperty.Value
                Dim groupType As String = TsdReadText(Function() group.Type.Value.ToString())
                Dim linkRegions As String() = If(groupType.IndexOf("Support", StringComparison.OrdinalIgnoreCase) >= 0,
                                                 New String() {"Bottom", "Top"},
                                                 New String() {"Mid"})
                For Each linkRegion As String In linkRegions
                    Dim ro As New JObject()
                    ro("kind") = "Link"
                    ro("zone") = groupType
                    ro("region") = linkRegion
                    ro("position") = linkRegion
                    TsdTrySet(ro, "legs", Function() CObj(TsdColumnLinkLegs(group)))
                    TsdTrySet(ro, "prefix", Function() group.DetailingPrefix.Value.ToString())
                    TsdTrySet(ro, "size", Function() group.Size.Value.Size.ToString())
                    TsdTrySet(ro, "spacing", Function() CObj(Convert.ToDouble(group.CentreSpacing.Value)))
                    TsdTrySet(ro, "grade", Function() group.Size.Value.Grade.Name)
                    TsdTrySet(ro, "spanIndex", Function() CObj(group.Value(0).Value.StartSpanIndex.Value))
                    TsdSetZoneLengths(ro, group)
                    ro("bars") = TsdBarText(ro)
                    rows.Add(ro)
                Next
            Next
        Catch
        End Try

        Return rows
    End Function

    ''' <summary>
    ''' Zone (region) extent of a reinforcement group along the member. TSD reports
    ''' the extent of a detailing group by the position of its first and last bar,
    ''' so the start / end positions are read from the group and the zone length is
    ''' the distance between them. Groups that do not expose a position (for example
    ''' slab layers) simply leave the values out.
    ''' </summary>
    Private Shared Sub TsdSetZoneLengths(row As JObject, group As Object)
        Dim readLength = Function(names As String()) As Object
                             For Each name As String In names
                                 Dim value = TsdUnwrap(TsdReadProperty(group, name))
                                 If value IsNot Nothing AndAlso IsNumeric(value) Then Return Convert.ToDouble(value)
                             Next
                             Return Nothing
                         End Function

        Dim startPosition = readLength(New String() {"StartPosition", "Start", "StartOffset", "ZoneStart"})
        Dim endPosition = readLength(New String() {"EndPosition", "EndValue", "EndOffset", "ZoneEnd"})
        Dim zoneLength = readLength(New String() {"ZoneLength", "Length", "BarLength", "TotalLength"})

        If startPosition Is Nothing OrElse endPosition Is Nothing Then
            ' Fall back to the extent of the individual bars of the group.
            Try
                Dim bars = TryCast(TsdUnwrap(TsdReadProperty(group, "Value")), System.Collections.IEnumerable)
                If bars IsNot Nothing Then
                    Dim minStart As Double = Double.MaxValue
                    Dim maxEnd As Double = Double.MinValue
                    For Each barProperty In bars
                        Dim bar = TsdUnwrap(barProperty)
                        Dim barStart = TsdUnwrap(TsdReadProperty(bar, "StartPosition"))
                        Dim barEnd = TsdUnwrap(TsdReadProperty(bar, "EndPosition"))
                        If barStart IsNot Nothing AndAlso IsNumeric(barStart) Then minStart = Math.Min(minStart, Convert.ToDouble(barStart))
                        If barEnd IsNot Nothing AndAlso IsNumeric(barEnd) Then maxEnd = Math.Max(maxEnd, Convert.ToDouble(barEnd))
                    Next
                    If minStart < Double.MaxValue AndAlso maxEnd > Double.MinValue Then
                        startPosition = minStart
                        endPosition = maxEnd
                    End If
                End If
            Catch
            End Try
        End If

        If startPosition IsNot Nothing Then row("zoneStart") = Convert.ToDouble(startPosition)
        If endPosition IsNot Nothing Then row("zoneEnd") = Convert.ToDouble(endPosition)
        If zoneLength Is Nothing AndAlso startPosition IsNot Nothing AndAlso endPosition IsNot Nothing Then
            zoneLength = Math.Abs(Convert.ToDouble(endPosition) - Convert.ToDouble(startPosition))
        End If

        If zoneLength Is Nothing Then
            ' Detailing groups do not publish a zone extent directly. The DCM column and
            ' wall stirrup extraction derives it from the bar layout instead: a group of
            ' n link positions at the centre spacing s covers s * (n - 1). Longitudinal
            ' groups repeat the bars around the section, so the number of positions is
            ' the group count divided by the bars per position.
            Dim spacing = readLength(New String() {"CentreSpacing", "Spacing"})
            If spacing IsNot Nothing AndAlso Convert.ToDouble(spacing) > 0 Then
                Dim positions As Double = 0
                Try
                    positions = Convert.ToDouble(TsdUnwrap(TsdReadProperty(group, "Count")))
                Catch
                End Try
                Dim perPosition = TsdUnwrap(TsdReadProperty(group, "LinksPerPosition"))
                If perPosition IsNot Nothing AndAlso IsNumeric(perPosition) AndAlso Convert.ToDouble(perPosition) > 0 Then
                    positions /= Convert.ToDouble(perPosition)
                End If
                If positions > 1 Then zoneLength = Convert.ToDouble(spacing) * (positions - 1)
            End If
        End If

        If zoneLength IsNot Nothing Then row("zoneLength") = Convert.ToDouble(zoneLength)
    End Sub

    ''' <summary>
    ''' Number of link legs of a column link detailing group. Following the DCM
    ''' column stirrup logic, a closed stirrup contributes two legs per link on the
    ''' longer face; a link whose larger geometry leg is shorter than the column
    ''' face (a C link) contributes a single leg on the shorter face.
    ''' </summary>
    Private Shared Function TsdColumnLinkLegs(group As Object) As Integer
        Dim linksPerPosition As Integer = 1
        Try
            linksPerPosition = Convert.ToInt32(CallByName(group, "LinksPerPosition", CallType.Get).Value)
        Catch
        End Try
        If linksPerPosition < 1 Then linksPerPosition = 1
        Return 2 * linksPerPosition
    End Function

    ''' <summary>Normalises the TSD beam group position to Left / Mid / Right.</summary>
    Private Shared Function TsdBeamPositionText(position As String) As String
        If String.IsNullOrEmpty(position) Then Return ""
        If position.IndexOf("Left", StringComparison.OrdinalIgnoreCase) >= 0 Then Return "Left"
        If position.IndexOf("Right", StringComparison.OrdinalIgnoreCase) >= 0 Then Return "Right"
        If position.IndexOf("Centre", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
           position.IndexOf("Center", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
           position.IndexOf("Middle", StringComparison.OrdinalIgnoreCase) >= 0 Then Return "Mid"
        Return position
    End Function

    ''' <summary>Combines a beam zone and position into "Top left", "Bottom mid", etc.</summary>
    Private Shared Function TsdRegionText(zone As String, position As String) As String
        If String.IsNullOrEmpty(zone) Then Return position
        If String.IsNullOrEmpty(position) Then Return zone
        Return zone & " " & position.ToLowerInvariant()
    End Function

    ''' <summary>Reads a string valued property, returning an empty string on failure.</summary>
    Private Shared Function TsdReadText(reader As Func(Of String)) As String
        Try
            Return If(reader(), "")
        Catch
            Return ""
        End Try
    End Function

    ''' <summary>
    ''' Vertical / horizontal / link groups of a structural wall reinforcement
    ''' collection, reported per wall stack (panel / span number).
    ''' </summary>
    Private Shared Function TsdWallRebarRows(reinforcement As TSD.API.Remoting.Reinforcement.IReinforcementCollection) As JArray
        Dim rows As New JArray()
        If reinforcement Is Nothing Then Return rows

        Try
            For Each groupProperty In reinforcement.LongitudinalBarGroups
                Dim group = groupProperty.Value
                Dim ro As New JObject()
                ro("kind") = "Vertical"
                ro("region") = "Vertical"
                TsdTrySet(ro, "stack", Function() CObj(TsdWallStackNumber(group)))
                TsdTrySet(ro, "spanIndex", Function() CObj(TsdWallStackNumber(group) - 1))
                TsdTrySet(ro, "count", Function() CObj(group.Count))
                TsdTrySet(ro, "prefix", Function() group.Size.Value.Grade.DetailingPrefix.ToString())
                TsdTrySet(ro, "size", Function() CObj(group.Size.Value.Diameter))
                TsdTrySet(ro, "spacing", Function() CObj(Convert.ToDouble(group.DesignSpacing.Value)))
                TsdTrySet(ro, "grade", Function() group.Size.Value.Grade.Name)
                TsdSetZoneLengths(ro, group)
                ro("bars") = TsdBarText(ro)
                rows.Add(ro)
            Next
        Catch
        End Try

        Try
            For Each groupProperty In reinforcement.HorizontalBarGroups
                Dim group = groupProperty.Value
                Dim ro As New JObject()
                ro("kind") = "Horizontal"
                ro("region") = "Horizontal"
                TsdTrySet(ro, "stack", Function() CObj(TsdWallStackNumber(group)))
                TsdTrySet(ro, "spanIndex", Function() CObj(TsdWallStackNumber(group) - 1))
                TsdTrySet(ro, "count", Function() CObj(group.Count))
                TsdTrySet(ro, "prefix", Function() group.Size.Value.Grade.DetailingPrefix.ToString())
                TsdTrySet(ro, "size", Function() CObj(group.Size.Value.Diameter))
                TsdTrySet(ro, "spacing", Function() CObj(Convert.ToDouble(group.CentreSpacing.Value)))
                TsdTrySet(ro, "grade", Function() group.Size.Value.Grade.Name)
                TsdSetZoneLengths(ro, group)
                ro("bars") = TsdBarText(ro)
                rows.Add(ro)
            Next
        Catch
        End Try

        Try
            For Each groupProperty In reinforcement.LinkGroups
                Dim group = groupProperty.Value
                Dim ro As New JObject()
                ro("kind") = "Link"
                ro("region") = "Link"
                TsdTrySet(ro, "stack", Function() CObj(TsdWallStackNumber(group)))
                TsdTrySet(ro, "spanIndex", Function() CObj(TsdWallStackNumber(group) - 1))
                TsdTrySet(ro, "legs", Function() CObj(TsdColumnLinkLegs(group)))
                TsdTrySet(ro, "prefix", Function() group.Size.Value.Grade.DetailingPrefix.ToString())
                TsdTrySet(ro, "size", Function() CObj(group.Size.Value.Diameter))
                TsdTrySet(ro, "spacing", Function() CObj(Convert.ToDouble(group.CentreSpacing.Value)))
                TsdTrySet(ro, "grade", Function() group.Size.Value.Grade.Name)
                TsdSetZoneLengths(ro, group)
                ro("bars") = TsdBarText(ro)
                rows.Add(ro)
            Next
        Catch
        End Try

        Return rows
    End Function

    ''' <summary>
    ''' One based stack (panel / span) number of a wall reinforcement group, taken
    ''' from the StartSpanIndex of the first bar in the group.
    ''' </summary>
    Private Shared Function TsdWallStackNumber(group As Object) As Integer
        Try
            Return Convert.ToInt32(CallByName(CallByName(group, "Value", CallType.Get)(0), "Value", CallType.Get).StartSpanIndex.Value) + 1
        Catch
            Return 1
        End Try
    End Function

    ''' <summary>
    ''' Slab / mat reinforcement layers, following the ISlabReinforcement outside and
    ''' inside layer properties used by the data extractor.
    ''' </summary>
    Private Shared Function TsdSlabRebarRows(reinforcement As TSD.API.Remoting.Reinforcement.ISlabReinforcement) As JArray
        Dim rows As New JArray()
        If reinforcement Is Nothing Then Return rows

        Dim addLayer = Sub(kind As String, sizeReader As Func(Of Object), spacingReader As Func(Of Object), gradeReader As Func(Of Object), prefixReader As Func(Of Object))
                           Dim ro As New JObject()
                           ro("kind") = kind
                           TsdTrySet(ro, "size", sizeReader)
                           TsdTrySet(ro, "spacing", spacingReader)
                           TsdTrySet(ro, "grade", gradeReader)
                           TsdTrySet(ro, "prefix", prefixReader)
                           If ro("size") IsNot Nothing OrElse ro("spacing") IsNot Nothing Then
                               ro("bars") = TsdBarText(ro)
                               rows.Add(ro)
                           End If
                       End Sub

        addLayer("Outside layer",
                 Function() reinforcement.SizeOutside.Value.Size,
                 Function() CObj(Convert.ToDouble(reinforcement.BarDistanceOutside.Value)),
                 Function() reinforcement.SizeOutside.Value.Grade.Name,
                 Function() reinforcement.SizeOutside.Value.Grade.DetailingPrefix)

        addLayer("Inside layer",
                 Function() reinforcement.SizeInside.Value.Size,
                 Function() CObj(Convert.ToDouble(reinforcement.BarDistanceInside.Value)),
                 Function() reinforcement.SizeInside.Value.Grade.Name,
                 Function() reinforcement.SizeInside.Value.Grade.DetailingPrefix)

        Try
            Dim xo As New JObject()
            xo("kind") = "Layer direction"
            xo("bars") = If(reinforcement.HasOutsideLayerInXDirection.Value, "Outside layer runs in X", "Outside layer runs in Y")
            rows.Add(xo)
        Catch
        End Try

        Return rows
    End Function

    ''' <summary>
    ''' Builds the usual detailing string from a rebar row. Links are prefixed with
    ''' the link leg count (2H8-125), longitudinal bars with the bar count and no
    ''' spacing suffix when the group has no design spacing (12H25).
    ''' </summary>
    Private Shared Function TsdBarText(row As JObject) As String
        Dim part = Function(key As String) If(row(key) Is Nothing, "", row(key).ToString())
        Dim quantity As String = If(row("legs") IsNot Nothing, part("legs"), part("count"))
        Dim text As String = quantity & part("prefix") & part("size")
        If row("spacing") IsNot Nothing AndAlso part("spacing") <> "" AndAlso part("spacing") <> "0" Then
            text &= "-" & part("spacing")
        End If
        Return text.Trim()
    End Function

    Private Shared Function TsdPoint(x As Double, y As Double, z As Double) As JObject
        Return New JObject(New JProperty("x", x), New JProperty("y", y), New JProperty("z", z))
    End Function

    ''' <summary>
    ''' Unwraps the TSD optional-style wrappers (IProperty(Of T) publishes HasValue / Value)
    ''' so a property can be read without knowing whether it is wrapped.
    ''' </summary>
    Private Shared Function TsdUnwrap(value As Object) As Object
        Dim current As Object = value
        For attempt As Integer = 0 To 5
            If current Is Nothing Then Return Nothing
            Dim type = current.GetType()
            If type.IsPrimitive OrElse type.IsEnum OrElse TypeOf current Is String OrElse TypeOf current Is Decimal Then Return current
            If Not TsdIsPropertyWrapper(type) Then Return current
            Dim valueProperty = TsdFindProperty(type, "Value")
            If valueProperty Is Nothing Then Return current
            Try
                current = valueProperty.GetValue(current, Nothing)
            Catch
                Return Nothing
            End Try
        Next
        Return current
    End Function

    ''' <summary>
    ''' Finds a readable property by name on a TSD object type. The remoting objects are
    ''' proxies that implement their API interfaces explicitly, so a property such as
    ''' DegreeOfFreedom or LengthDir1 is not visible on the concrete type and is only
    ''' found by walking the implemented interfaces.
    ''' </summary>
    Private Shared Function TsdFindProperty(type As Type, name As String) As System.Reflection.PropertyInfo
        If type Is Nothing Then Return Nothing
        Dim info As System.Reflection.PropertyInfo = Nothing
        Try
            info = type.GetProperty(name)
        Catch ex As System.Reflection.AmbiguousMatchException
            info = Nothing
        End Try
        If info IsNot Nothing AndAlso info.CanRead Then Return info

        For Each contract In type.GetInterfaces()
            Try
                Dim candidate = contract.GetProperty(name)
                If candidate IsNot Nothing AndAlso candidate.CanRead Then Return candidate
            Catch
            End Try
        Next
        Return Nothing
    End Function

    ''' <summary>
    ''' True when the type is an optional-style container that carries the real value in
    ''' its Value member: either a Nullable / HasValue pair or one of the TSD
    ''' IProperty(Of T) wrappers (EnumProperty, MaterialProperty, PileTypeProperty and
    ''' the rest), which publish Value without a HasValue companion.
    ''' </summary>
    Private Shared Function TsdIsPropertyWrapper(type As Type) As Boolean
        If TsdFindProperty(type, "HasValue") IsNot Nothing Then Return True
        For Each contract In type.GetInterfaces()
            If contract.IsGenericType AndAlso
               contract.GetGenericTypeDefinition().FullName = "TSD.API.Remoting.Common.Properties.IProperty`1" Then Return True
        Next
        Return False
    End Function

    ''' <summary>
    ''' Reads a named property from a TSD object by reflection, unwrapping optional
    ''' wrappers and returning Nothing when the property does not exist on this type.
    ''' </summary>
    Private Shared Function TsdReadProperty(source As Object, name As String) As Object
        If source Is Nothing Then Return Nothing
        Try
            Dim info = TsdFindProperty(source.GetType(), name)
            If info Is Nothing Then Return Nothing
            Return TsdUnwrap(info.GetValue(source, Nothing))
        Catch
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' Reads a named property as text. Enum and object valued properties are read
    ''' through ToString, but a property that is missing on this type returns Nothing
    ''' instead of throwing, so an optional API member never raises a
    ''' NullReferenceException while the model is being serialized.
    ''' </summary>
    Private Shared Function TsdReadPropertyText(source As Object, name As String) As Object
        Dim value = TsdReadProperty(source, name)
        If value Is Nothing Then Return Nothing
        Return value.ToString()
    End Function

    ''' <summary>
    ''' Converts a force published by the TSD API into kN. The remoting API reports pile
    ''' resistances and limits in newtons, while the Math canvas works in kN throughout.
    ''' </summary>
    Private Shared Function TsdForceKn(value As Object) As Object
        If value Is Nothing OrElse Not IsNumeric(value) Then Return Nothing
        Return Convert.ToDouble(value) / 1000.0
    End Function

    ''' <summary>
    ''' Converts a rotational stiffness published by the TSD API into kNm/rad. The
    ''' remoting API reports rotational springs in Nmm/rad, so 1.0E9 Nmm/rad is
    ''' 1000 kNm/rad: newtons to kilonewtons and millimetres to metres are both
    ''' factors of 1000.
    ''' </summary>
    Private Shared Function TsdRotationalStiffnessKnm(value As Object) As Object
        If value Is Nothing OrElse Not IsNumeric(value) Then Return Nothing
        Return Convert.ToDouble(value) / 1000000.0
    End Function

    ''' <summary>
    ''' Readable label of a material, concrete grade or pile type. IHaveName publishes
    ''' Name, IPileType labels itself through Description, and anything else falls back
    ''' to its text form so a grade never appears as a wrapper type name.
    ''' </summary>
    Private Shared Function TsdMaterialName(source As Object) As Object
        Dim item = TsdUnwrap(source)
        If item Is Nothing Then Return Nothing
        For Each propertyName As String In {"Name", "Description", "ShortName", "LongName"}
            Dim value = TsdReadProperty(item, propertyName)
            If value IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(value.ToString()) Then Return value.ToString()
        Next
        Return item.ToString()
    End Function

    ''' <summary>
    ''' Property names published by the steel / concrete section objects, mapped from
    ''' the TSD API name to the key the Math canvas reads.
    ''' </summary>
    Private Shared ReadOnly TsdSectionPropertyNames As String()() = {
        New String() {"Mass", "massPerLength"},
        New String() {"Depth", "depth"},
        New String() {"Breadth", "breadth"},
        New String() {"TopFlangeBreadth", "topFlangeBreadth"},
        New String() {"BottomFlangeBreadth", "bottomFlangeBreadth"},
        New String() {"WebThickness", "webThickness"},
        New String() {"FlangeThickness", "flangeThickness"},
        New String() {"RootRadius", "rootRadius"},
        New String() {"OuterDiameter", "outerDiameter"},
        New String() {"Thickness", "wallThickness"},
        New String() {"MajorAxisElasticSectionModulus", "majorElasticModulus"},
        New String() {"MinorAxisElasticSectionModulus", "minorElasticModulus"},
        New String() {"MajorAxisPlasticSectionModulus", "majorPlasticModulus"},
        New String() {"MinorAxisPlasticSectionModulus", "minorPlasticModulus"},
        New String() {"MajorAxisRadiusOfGyration", "majorRadiusOfGyration"},
        New String() {"MinorAxisRadiusOfGyration", "minorRadiusOfGyration"},
        New String() {"WarpingConstant", "warpingConstant"},
        New String() {"SurfaceAreaPerUnitLength", "surfaceAreaPerLength"},
        New String() {"SurfaceAreaPerUnitMass", "surfaceAreaPerMass"}}

    ''' <summary>
    ''' Section properties of a physical section. ISection publishes the analysis
    ''' properties (area, second moments, shear areas, torsion constant) for every
    ''' section, and the concrete steel section objects add the rolled-section
    ''' dimensions and moduli read here by name.
    ''' </summary>
    Private Shared Function TsdSectionProperties(physical As Object) As JObject
        Dim so As New JObject()
        If physical Is Nothing Then Return so

        TsdTrySet(so, "sectionName", Function() TsdReadProperty(physical, "LongName"))
        TsdTrySet(so, "sectionShortName", Function() TsdReadProperty(physical, "ShortName"))
        TsdTrySet(so, "materialType", Function() TsdReadPropertyText(physical, "MaterialType"))
        TsdTrySet(so, "material", Function() TsdReadPropertyText(physical, "MaterialType"))
        TsdTrySet(so, "sectionGeometry", Function() TsdReadPropertyText(physical, "SectionGeometry"))
        TsdTrySet(so, "area", Function() TsdReadProperty(physical, "CrossSectionalArea"))
        TsdTrySet(so, "majorInertia", Function() TsdReadProperty(physical, "MajorAxisSecondMomentOfArea"))
        TsdTrySet(so, "minorInertia", Function() TsdReadProperty(physical, "MinorAxisSecondMomentOfArea"))
        TsdTrySet(so, "torsionConstant", Function() TsdReadProperty(physical, "TorsionConstant"))
        TsdTrySet(so, "majorShearArea", Function() TsdReadProperty(physical, "ShearAreaLoadedParallelToMajorAxis"))
        TsdTrySet(so, "minorShearArea", Function() TsdReadProperty(physical, "ShearAreaLoadedParallelToMinorAxis"))

        For Each entry As String() In TsdSectionPropertyNames
            Dim propertyName As String = entry(0)
            Dim key As String = entry(1)
            TsdTrySet(so, key, Function() TsdReadProperty(physical, propertyName))
        Next

        Return so
    End Function

    ''' <summary>
    ''' End fixity of a member span end. The degree of freedom string reports the
    ''' released directions, the rotational stiffness data reports whether the end is
    ''' fixed, pinned or a spring, and the spring stiffness is published when the
    ''' stiffness type is a spring (as used by the DCM design forms).
    ''' </summary>
    Private Shared Function TsdReleaseObject(data As TSD.API.Remoting.Structure.ISpanReleases,
                                             endName As String) As JObject
        Dim ro As New JObject()
        ro("end") = endName
        If data Is Nothing Then
            ro("fixity") = "Not published"
            Return ro
        End If

        ' The degree of freedom is read through the strongly typed contract, the way the
        ' steel member export does (BeamSpans(j).StartReleases.Value.DegreeOfFreedom.Value),
        ' because the remoting proxies implement ISpanReleases explicitly.
        Dim dofText As String = Nothing
        Try
            Dim dof As TSD.API.Remoting.Solver.DegreeOfFreedom = data.DegreeOfFreedom.Value
            dofText = dof.ToString()
            ro("degreeOfFreedom") = dofText
        Catch ex As Exception
            ro("degreeOfFreedomError") = ex.Message
        End Try

        ' Every read below is a remoting round trip, and this runs for both ends of every
        ' span of every member, so only the members the canvas actually shows are fetched
        ' and each is read through the strongly typed contract rather than by reflection.
        TsdTrySet(ro, "cantilever", Function() CObj(data.Cantilever.Value))

        ' The rotational stiffness is an IStiffnessData behind a property wrapper, so both
        ' levels are dereferenced: StartReleases.Value.MajorRotationalStiffness.Value.Type
        ' for the Solver.SpringStiffness kind and .Value.Stiffness.Value for each number.
        Try
            TsdSetStiffness(ro, "major", data.MajorRotationalStiffness.Value)
        Catch ex As Exception
            ro("majorStiffnessError") = ex.Message
        End Try
        Try
            TsdSetStiffness(ro, "minor", data.MinorRotationalStiffness.Value)
        Catch ex As Exception
            ro("minorStiffnessError") = ex.Message
        End Try

        ' Direction by direction release flags, so a partially released end can be read
        ' without decoding the degree of freedom text on the canvas side.
        If Not String.IsNullOrEmpty(dofText) Then
            Dim flags = TsdDegreeOfFreedomFlags(dofText)
            ro("shearMajorReleased") = Not flags.Contains("FX")
            ro("shearMinorReleased") = Not flags.Contains("FY")
            ro("axialReleased") = Not flags.Contains("FZ")
            ro("torsionReleased") = Not flags.Contains("MX")
            ro("momentMajorReleased") = Not flags.Contains("MY")
            ro("momentMinorReleased") = Not flags.Contains("MZ")
        End If

        ro("fixity") = TsdFixityText(ro)
        Return ro
    End Function

    ''' <summary>
    ''' Publishes one rotational stiffness of a span end. The kind is the
    ''' Solver.SpringStiffness value of IStiffnessData.Type and every numeric member sits
    ''' behind its own property wrapper, so each is read as .Value. Prefix is "major" or
    ''' "minor", giving the majorType / majorRotationalStiffness names the canvas reads.
    ''' Each read is a remoting round trip, so only the members the stiffness kind
    ''' actually defines are requested: a released or fully fixed end carries no spring
    ''' numbers at all and reading them would stall the connection on every span.
    ''' </summary>
    Private Shared Sub TsdSetStiffness(ro As JObject, prefix As String,
                                       stiffness As TSD.API.Remoting.Solver.IStiffnessData)
        If stiffness Is Nothing Then Return

        Dim kind As TSD.API.Remoting.Solver.SpringStiffness
        Try
            kind = stiffness.Type.Value
        Catch ex As Exception
            ro(prefix & "StiffnessError") = ex.Message
            Return
        End Try
        ro(prefix & "Type") = kind.ToString()

        Select Case kind
            Case TSD.API.Remoting.Solver.SpringStiffness.SpringLinear
                ' The canvas fixity table reads majorRotationalStiffness /
                ' minorRotationalStiffness, so the linear stiffness keeps that name. The
                ' API publishes Nmm/rad and the canvas works in kNm/rad throughout.
                TsdTrySet(ro, prefix & "RotationalStiffness", Function() TsdRotationalStiffnessKnm(stiffness.Stiffness.Value))

            Case TSD.API.Remoting.Solver.SpringStiffness.SpringNonLinear
                TsdTrySet(ro, prefix & "StiffnessCompression", Function() TsdRotationalStiffnessKnm(stiffness.StiffnessCompression.Value))
                TsdTrySet(ro, prefix & "StiffnessTension", Function() TsdRotationalStiffnessKnm(stiffness.StiffnessTension.Value))
                TsdTrySet(ro, prefix & "MaxForceCompression", Function() TsdForceKn(stiffness.MaxForceCompression.Value))
                TsdTrySet(ro, prefix & "MaxForceTension", Function() TsdForceKn(stiffness.MaxForceTension.Value))

            Case TSD.API.Remoting.Solver.SpringStiffness.NominallyPinned
                TsdTrySet(ro, prefix & "NominalPinPercentage", Function() stiffness.NominallyPinnedPercentage.Value)

            Case TSD.API.Remoting.Solver.SpringStiffness.NominallyFixed
                TsdTrySet(ro, prefix & "NominalFixPercentage", Function() stiffness.NominallyFixedPercentage.Value)

            Case TSD.API.Remoting.Solver.SpringStiffness.PartiallyFixed
                TsdTrySet(ro, prefix & "PartialFixPercentage", Function() stiffness.PartiallyFixedPercentage.Value)
        End Select
    End Sub

    ''' <summary>
    ''' Restrained direction names of a solver degree of freedom value. The flags print
    ''' as "Fx, Fy, Fz, Mx" and a combined value may add an " Or " alternative, of which
    ''' only the first is meaningful, as in the steel member export.
    ''' </summary>
    Private Shared Function TsdDegreeOfFreedomFlags(degreeOfFreedom As String) As HashSet(Of String)
        Dim head As String = Split(degreeOfFreedom.ToUpperInvariant(), " OR ")(0)
        Return New HashSet(Of String)(head.Split(","c).Select(Function(part) part.Trim()).Where(Function(part) part.Length > 0))
    End Function

    ''' <summary>
    ''' Condenses the release data into the familiar Pinned / Fixed / Spring wording.
    ''' A cantilever end is reported as such, a rotational stiffness type that is not
    ''' fully fixed is a spring, a released rotation (DegreeOfFreedom Mx / My / Mz or
    ''' Free) is a pinned end, and anything else is a fixed (continuous) end. The text
    ''' is always populated so the fixity column is never blank.
    ''' </summary>
    Private Shared Function TsdFixityText(release As JObject) As String
        Dim majorType As String = If(release("majorType") Is Nothing, "", release("majorType").ToString())
        Dim minorType As String = If(release("minorType") Is Nothing, "", release("minorType").ToString())
        Dim dof As String = If(release("degreeOfFreedom") Is Nothing, "", release("degreeOfFreedom").ToString())
        Dim cantilever As Boolean = release("cantilever") IsNot Nothing AndAlso
                                    release("cantilever").Type = JTokenType.Boolean AndAlso
                                    release("cantilever").Value(Of Boolean)()

        If cantilever Then Return "Cantilever"

        ' The degree of freedom flags are the authoritative source, exactly as in the
        ' steel member export, so they are classified before any stiffness wording.
        If Not String.IsNullOrEmpty(dof) Then Return TsdConnectionType(dof)

        For Each stiffnessType As String In {majorType, minorType}
            If stiffnessType.IndexOf("Spring", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
               stiffnessType.IndexOf("Partial", StringComparison.OrdinalIgnoreCase) >= 0 Then Return "Spring"
        Next
        For Each stiffnessType As String In {majorType, minorType}
            If stiffnessType.IndexOf("Pinned", StringComparison.OrdinalIgnoreCase) >= 0 Then Return "Pinned"
        Next
        If String.IsNullOrEmpty(majorType) AndAlso String.IsNullOrEmpty(minorType) Then Return "Not published"
        Return "Fixed"
    End Function

    ''' <summary>
    ''' Free / Pinned / Fixed classification of a span end from the solver degree of
    ''' freedom flags, using the same rule as the steel member export: the flag list
    ''' names the directions that are NOT released, so an end that names none of Fx / Fy
    ''' / Fz is free, an end that restrains all three translations while releasing both
    ''' bending rotations is pinned, and anything else is a fixed (moment) connection.
    ''' </summary>
    Private Shared Function TsdConnectionType(degreeOfFreedom As String) As String
        Dim flags = TsdDegreeOfFreedomFlags(degreeOfFreedom)

        ' True = released (the direction is absent from the flag list).
        Dim shearMajorReleased As Boolean = Not flags.Contains("FX")
        Dim shearMinorReleased As Boolean = Not flags.Contains("FY")
        Dim axialReleased As Boolean = Not flags.Contains("FZ")
        Dim momentMajorReleased As Boolean = Not flags.Contains("MY")
        Dim momentMinorReleased As Boolean = Not flags.Contains("MZ")

        If shearMajorReleased AndAlso shearMinorReleased AndAlso axialReleased Then Return "Free"
        If Not shearMajorReleased AndAlso Not shearMinorReleased AndAlso Not axialReleased AndAlso
           momentMajorReleased AndAlso momentMinorReleased Then Return "Pinned"
        Return "Fixed"
    End Function

    ''' <summary>
    ''' Converts a value read from the API into a JSON value. Enums and any other
    ''' object the serializer does not understand are published as their text form so
    ''' the canvas never receives a raw wrapper type name in a table cell.
    ''' </summary>
    Private Shared Function TsdToJValue(value As Object) As JValue
        If value Is Nothing Then Return Nothing
        Dim type = value.GetType()
        If type.IsEnum Then Return New JValue(value.ToString())
        If type.IsPrimitive OrElse TypeOf value Is String OrElse TypeOf value Is Decimal Then Return New JValue(value)
        Try
            Return New JValue(value)
        Catch
            Return New JValue(value.ToString())
        End Try
    End Function

    Private Shared Sub TsdTrySet(target As JObject, key As String, reader As Func(Of Object))
        Try
            Dim token = TsdToJValue(reader())
            If token IsNot Nothing Then target(key) = token
        Catch
        End Try
    End Sub

    ''' <summary>
    ''' Reflection based dump used for reinforcement data so that every published
    ''' property is forwarded to the canvas without hard coding the API shape.
    ''' </summary>
    Private Shared Function TsdDump(value As Object, depth As Integer) As JToken
        If value Is Nothing Then Return JValue.CreateNull()

        Dim type = value.GetType()
        If TypeOf value Is String Then Return New JValue(DirectCast(value, String))
        If type.IsEnum Then Return New JValue(value.ToString())
        If type.IsPrimitive OrElse TypeOf value Is Decimal OrElse TypeOf value Is DateTime OrElse TypeOf value Is Guid Then
            Return New JValue(If(TypeOf value Is Guid OrElse TypeOf value Is DateTime, CObj(value.ToString()), value))
        End If
        If depth <= 0 Then Return New JValue(value.ToString())

        ' Unwrap TSD optional-style wrappers that publish a single Value property.
        Dim valueProperty = type.GetProperty("Value")
        If valueProperty IsNot Nothing AndAlso type.GetProperty("HasValue") IsNot Nothing Then
            Try
                Return TsdDump(valueProperty.GetValue(value, Nothing), depth - 1)
            Catch ex As Exception
                Return JValue.CreateString("<" & ex.Message & ">")
            End Try
        End If

        Dim sequence = TryCast(value, System.Collections.IEnumerable)
        If sequence IsNot Nothing Then
            Dim array As New JArray()
            Dim count As Integer = 0
            For Each element In sequence
                array.Add(TsdDump(element, depth - 1))
                count += 1
                If count >= 200 Then Exit For
            Next
            Return array
        End If

        Dim result As New JObject()
        For Each info In type.GetProperties()
            If info.GetIndexParameters().Length > 0 OrElse Not info.CanRead Then Continue For
            Try
                result(info.Name) = TsdDump(info.GetValue(value, Nothing), depth - 1)
            Catch ex As Exception
                result(info.Name) = JValue.CreateString("<" & ex.Message & ">")
            End Try
        Next
        If result.Count = 0 Then Return New JValue(value.ToString())
        Return result
    End Function

    Protected Overrides Sub OnFormClosed(e As FormClosedEventArgs)
        If _presentationFullscreen Then SetPresentationFullscreen(False)
        If _excelApp IsNot Nothing Then
            Try
                RemoveHandler _excelApp.SheetChange, AddressOf OnExcelSheetChange
                Marshal.ReleaseComObject(_excelApp)
                _excelApp = Nothing
            Catch
            End Try
        End If
        If wbMath IsNot Nothing AndAlso wbMath.CoreWebView2 IsNot Nothing Then
            RemoveHandler wbMath.CoreWebView2.WebMessageReceived, AddressOf OnWebMessageReceived
        End If
        MyBase.OnFormClosed(e)
    End Sub

    Public Sub OpenCanvasFile()
        LoadCanvas()
    End Sub
End Class