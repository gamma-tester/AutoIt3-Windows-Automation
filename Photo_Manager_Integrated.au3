;==============================================================
; Photo Manager Integrated - AutoIt3 version
; Combines Copy & Resize with Face Detection & Crop
; Uploads to SFTP, Integrates with Cardpresso
;==============================================================
#include <ButtonConstants.au3>
#include <EditConstants.au3>
#include <GUIConstantsEx.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#include <FileConstants.au3>
#include <MsgBoxConstants.au3>
#include <File.au3>
#include <GDIPlus.au3>
#include <INet.au3>

; Global variables
Global $hGui, $srcInput, $destInput, $nameInput, $statusLabel, $srcBrowseBtn, $destBrowseBtn, $processBtn, $progressBar, $previewPic
Global $faceDetectBtn, $modeCombo, $cardpressoBtn, $multiCardpressoBtn
Global $hSrcLabel, $hDestLabel
; Cardpresso preview variables
Global $sCardpressoPhotoPath = ""
; Cardpresso configuration variables
Global $sCardpressoMasterFile = ""
Global $sCardpressoCSVFile = ""
Global $sCardpressoExePath = ""
Global $sCardpressoSheetName = ""
Global $sCardpressoPhotoBaseDir = ""
Global $sScriptDir = StringRegExpReplace(@ScriptDir, "\\$", "") ; strip trailing back-slash
Global $sIrfanViewPath = ""
Global $sFaceAPIKey = ""
Global $sFaceAPIEndpoint = ""
Global $sFolder2Path = ""
Global $sFolder3Path = ""
Global $hPreviewBitmap = 0

; SFTP Upload variables
Global $sSFTPHost = ""
Global $sSFTPUsername = ""
Global $sSFTPRemotePath = ""
; Password will be stored securely using Windows Credential Manager

; Microsoft Face API variables
Global $hPreviewGUI, $hPreviewImage, $aDetectedFaces, $sCurrentImagePath, $sCurrentOutputPath
Global $iSelectedFace = -1  ; -1 means no face selected
Global $aFaceRegions[0]  ; Array to store face regions for click detection
Global $hCropButton  ; Global variable for the crop button
Global $manualSelectBtn, $bManualSelectionMode = False
Global $aManualSelection[4] = [-1, -1, -1, -1] ; [x1, y1, x2, y2]

; UI Guidance variables
Global $iCurrentStep = 1
Global $hStepLabel

;-------------------------------------------------------------- GUI
$hGui = GUICreate("Photo Manager Integrated", 800, 450, -1, -1, $WS_CAPTION + $WS_SYSMENU)
GUISetFont(9, 400, 0, "Segoe UI")

; Left side - Controls
GUICtrlCreateLabel("Operation Mode:", 20, 20, 350, 20)
$modeCombo = GUICtrlCreateCombo("", 20, 40, 350, 24)
GUICtrlSetData($modeCombo, "Resize_and_Copy|Detect_Face_and_Crop|Upload_to_SFTP|Cardpresso", "Resize_and_Copy")

$hSrcLabel = GUICtrlCreateLabel("Source Image File:", 20, 80, 350, 20)
$srcInput = GUICtrlCreateInput("", 20, 100, 350, 24)
GUICtrlSetState(-1, $GUI_DISABLE)
$srcBrowseBtn = GUICtrlCreateButton("Browse", 380, 100, 80, 24)

$hDestLabel = GUICtrlCreateLabel("Destination Folder:", 20, 140, 350, 20)
$destInput = GUICtrlCreateInput("", 20, 160, 350, 24)
$destBrowseBtn = GUICtrlCreateButton("Browse", 380, 160, 80, 24)

GUICtrlCreateLabel("Client Number:", 20, 200, 350, 20)
$nameInput = GUICtrlCreateInput("", 20, 220, 80, 24)
GUICtrlSetLimit(-1, 5)
GUICtrlSetStyle(-1, $ES_NUMBER) ; Only allow numbers

$processBtn = GUICtrlCreateButton("Copy & Resize", 20, 260, 120, 30)
$cropBtn = GUICtrlCreateButton("Crop Face", 150, 260, 120, 30)
GUICtrlSetState($cropBtn, $GUI_HIDE) ; Hidden initially, shown only in Face Detect mode
$manualSelectBtn = GUICtrlCreateButton("Manual Face Select", 20, 330, 120, 30)
GUICtrlSetState($manualSelectBtn, $GUI_HIDE) ; Hidden initially, shown only in Face Detect mode
$sftpTestBtn = GUICtrlCreateButton("Test SFTP", 20, 300, 120, 30)
GUICtrlSetState($sftpTestBtn, $GUI_HIDE) ; Hidden initially, shown only in SFTP mode
$cardpressoBtn = GUICtrlCreateButton("Open Single-ID-Card", 20, 300, 120, 30)
GUICtrlSetState($cardpressoBtn, $GUI_HIDE) ; Hidden initially, shown only in Cardpresso mode
$multiCardpressoBtn = GUICtrlCreateButton("Open Multi-ID-Card", 150, 300, 120, 30)
GUICtrlSetState($multiCardpressoBtn, $GUI_HIDE) ; Hidden initially, shown only in Cardpresso mode
$progressBar = GUICtrlCreateProgress(20, 370, 440, 20)
GUICtrlSetState($progressBar, $GUI_HIDE)

; Right side - Preview
GUICtrlCreateLabel("Image Preview:", 480, 20, 300, 20)
$previewPic = GUICtrlCreatePic("", 480, 40, 300, 300)
GUICtrlSetState($previewPic, $GUI_HIDE)

$statusLabel = GUICtrlCreateLabel("", 480, 350, 300, 80)
GUICtrlSetColor(-1, 0x008000)

; Step-by-step guidance label
$hStepLabel = GUICtrlCreateLabel("Step 1: Click 'Browse' to select source image", 20, 400, 440, 20)
GUICtrlSetColor(-1, 0x008000)
GUICtrlSetFont(-1, 9, 600) ; Bold font

;-------------------------------------------------------------- Load config
_LoadConfig()
_LoadIrfanViewPath()
_LoadFaceAPIConfig()
_LoadCardpressoConfig()
_GDIPlus_Startup()
GUISetState()

;-------------------------------------------------------------- Message loop
; Initialize step guidance
_UpdateStepGuidance()

; Global variable for button state tracking
Global $bLastButtonState = False

While 1
    Switch GUIGetMsg()
        Case $GUI_EVENT_CLOSE
            ExitLoop
        Case $srcBrowseBtn
            _BrowseFile()
            _HighlightNextStep()
        Case $destBrowseBtn
            _BrowseFolder()
            _HighlightNextStep()
        Case $processBtn
            _ProcessImageBasedOnMode()
        Case $cropBtn
            _CropSelectedFace()
        Case $manualSelectBtn
            _ToggleManualSelection()
        Case $sftpTestBtn
            _TestSFTPConnection()
        Case $cardpressoBtn
            _OpenCardpresso("Single-ID")
        Case $multiCardpressoBtn
            _OpenCardpresso("Multi-ID")
        Case $srcInput
            ; Update preview when source file changes
            _UpdatePreview()
            _HighlightNextStep()
        Case $destInput
            _HighlightNextStep()
        Case $nameInput
            _HighlightNextStep()
        Case $modeCombo
            _UpdateModeUI() ; Default parameter (False) clears client number for manual changes
            _HighlightNextStep()
        Case $previewPic
            ; Handle click on preview image for face selection
            If $bManualSelectionMode Then
                _HandleManualSelection()
            Else
                _HandlePreviewClick()
            EndIf
    EndSwitch
    
    ; Real-time monitoring for client number input (runs every loop iteration)
    ; Only update button color when actually needed to prevent blinking
    Local $sName = StringStripWS(GUICtrlRead($nameInput), 3)
    Local $sMode = GUICtrlRead($modeCombo)
    Local $sSrc = StringStripWS(GUICtrlRead($srcInput), 3)
    Local $sDest = StringStripWS(GUICtrlRead($destInput), 3)
    
    If $sMode = "Copy & Resize" Then
        Local $bCurrentState = (StringLen($sName) = 5 And StringIsDigit($sName) And $sSrc <> "" And $sDest <> "")
        
        ; Only update color if state actually changed
        If $bCurrentState <> $bLastButtonState Then
            If $bCurrentState Then
                GUICtrlSetBkColor($processBtn, 0x90EE90) ; Light green when ready
            Else
                GUICtrlSetBkColor($processBtn, 0xF0F0F0) ; Default color when not ready
            EndIf
            $bLastButtonState = $bCurrentState
        EndIf
    EndIf
WEnd

; Cleanup
_GDIPlus_Shutdown()
Exit

;==============================================================
; Functions
;==============================================================
Func _BrowseFile()
    Local $sMode = GUICtrlRead($modeCombo)
    
    ; For SFTP mode, allow multiple file selection
    If $sMode = "Upload_to_SFTP" Then
        Local $sFiles = FileOpenDialog("Select images to upload", @ScriptDir, "Images (*.jpg;*.jpeg;*.png)", 4 + 1) ; 4 = multiple files, 1 = file must exist
        If @error Then Return
        
        ; Process multiple files
        Local $aFiles = StringSplit($sFiles, "|", 2) ; No count flag
        If UBound($aFiles) > 1 Then
            ; First element is the folder path, rest are filenames
            Local $sFolder = $aFiles[0]
            Local $sFileList = ""
            For $i = 1 To UBound($aFiles) - 1
                Local $sFullPath = $sFolder & "\" & $aFiles[$i]
                ; Validate file type
                Local $sExt = StringLower(StringTrimLeft($sFullPath, StringInStr($sFullPath, ".", 0, -1)))
                If $sExt = "jpg" Or $sExt = "jpeg" Or $sExt = "png" Then
                    If $sFileList <> "" Then $sFileList &= "|"
                    $sFileList &= $sFullPath
                EndIf
            Next
            
            If $sFileList <> "" Then
                GUICtrlSetData($srcInput, $sFileList)
                _ShowStatus("Selected " & (StringSplit($sFileList, "|", 2)[0]) & " files for upload", 1)
                ; Update preview with first image
                _UpdatePreview()
                
                ; Extract client number from first filename for SFTP mode
                If $sMode = "Upload_to_SFTP" Then
                    Local $aFiles = StringSplit($sFileList, "|", 2)
                    If UBound($aFiles) > 0 Then
                        Local $sFirstFile = $aFiles[0]
                        Local $sFilename = StringRegExpReplace($sFirstFile, "^.*\\", "")
                        Local $sClientNumber = StringRegExpReplace($sFilename, "^(\d{5})\..*$", "$1")
                        If StringLen($sClientNumber) = 5 And StringIsDigit($sClientNumber) Then
                            GUICtrlSetData($nameInput, $sClientNumber)
                            _ShowStatus("Client number auto-extracted from filename: " & $sClientNumber, 1)
                        Else
                            GUICtrlSetData($nameInput, "")
                            _ShowStatus("Could not extract valid 5-digit client number from filename: " & $sFilename, 0)
                        EndIf
                    EndIf
                EndIf
            EndIf
        Else
            ; Single file selection fallback
            Local $sExt = StringLower(StringTrimLeft($sFiles, StringInStr($sFiles, ".", 0, -1)))
            If $sExt = "jpg" Or $sExt = "jpeg" Or $sExt = "png" Then
                GUICtrlSetData($srcInput, $sFiles)
                ; Update preview for single file selection
                _UpdatePreview()
                
                ; Extract client number from filename for SFTP mode
                If $sMode = "Upload_to_SFTP" Then
                    Local $sFilename = StringRegExpReplace($sFiles, "^.*\\", "")
                    Local $sClientNumber = StringRegExpReplace($sFilename, "^(\d{5})\..*$", "$1")
                    If StringLen($sClientNumber) = 5 And StringIsDigit($sClientNumber) Then
                        GUICtrlSetData($nameInput, $sClientNumber)
                        _ShowStatus("Client number auto-extracted from filename: " & $sClientNumber, 1)
                    Else
                        GUICtrlSetData($nameInput, "")
                        _ShowStatus("Could not extract valid 5-digit client number from filename: " & $sFilename, 0)
                    EndIf
                EndIf
            Else
                MsgBox(16, "Invalid File", "Please select valid image files (JPG, JPEG, or PNG).")
                Return
            EndIf
        EndIf
    Else
        ; For other modes, use single file selection
        Local $sFile = FileOpenDialog("Select an image", @ScriptDir, "Images (*.jpg;*.jpeg;*.png)", 1)
        If @error Then Return
        
        ; Validate file type
        Local $sExt = StringLower(StringTrimLeft($sFile, StringInStr($sFile, ".", 0, -1)))
        If $sExt <> "jpg" And $sExt <> "jpeg" And $sExt <> "png" Then
            MsgBox(16, "Invalid File", "Please select a valid image file (JPG, JPEG, or PNG).")
            Return
        EndIf
        
        GUICtrlSetData($srcInput, $sFile)
        ; Force update preview immediately
        _UpdatePreview()
    EndIf
    
    ; Update step guidance
    _HighlightNextStep()
EndFunc

Func _BrowseFolder()
    Local $sFolder = FileSelectFolder("Select destination folder", "")
    If @error Then Return
    GUICtrlSetData($destInput, $sFolder)
    ; Update step guidance
    _HighlightNextStep()
EndFunc

Func _UpdateModeUI($bPreserveClientNumber = False)
    Local $sMode = GUICtrlRead($modeCombo)
    If $sMode = "Detect_Face_and_Crop" Then
        ; Clear previous source file
        GUICtrlSetData($srcInput, "")
        ; Only reset client number if not preserving it (manual mode change)
        If Not $bPreserveClientNumber Then
            GUICtrlSetData($nameInput, "") ; Reset client number
        EndIf
        GUICtrlSetData($processBtn, "Detect Face")
        GUICtrlSetState($nameInput, $GUI_ENABLE) ; Enable client number for face detection too
        GUICtrlSetState($cropBtn, $GUI_DISABLE) ; Disabled initially until face is selected
        GUICtrlSetState($cropBtn, $GUI_SHOW) ; Show in Face Detect mode
        GUICtrlSetState($manualSelectBtn, $GUI_SHOW) ; Show in Face Detect mode
        GUICtrlSetState($manualSelectBtn, $GUI_ENABLE) ; Enable in Face Detect mode
        GUICtrlSetState($sftpTestBtn, $GUI_HIDE) ; Hide in Face Detect mode
        GUICtrlSetState($cardpressoBtn, $GUI_HIDE) ; Hide in Face Detect mode
        GUICtrlSetState($multiCardpressoBtn, $GUI_HIDE) ; Hide in Face Detect mode
        GUICtrlSetState($processBtn, $GUI_ENABLE) ; Enable Detect Face button
        ; Switch to Folder-2 for face detection
        If $sFolder2Path <> "" Then
            GUICtrlSetData($destInput, $sFolder2Path)
        EndIf
        ; Reset face detection state
        $aDetectedFaces = 0
        $iSelectedFace = -1
        $bManualSelectionMode = False
        _ClearPreview()
        If $bPreserveClientNumber Then
            _ShowStatus("Switched to Detect_Face_and_Crop mode. Client number preserved. Click 'Detect Face' to continue.", 1)
        Else
            _ShowStatus("Select Detect_Face_and_Crop mode - Fill client number and click 'Detect Face'", 1)
        EndIf
    ElseIf $sMode = "Cardpresso" Then
        ; Clear previous source file and reset client number
        GUICtrlSetData($srcInput, "")
        GUICtrlSetData($nameInput, "") ; Reset client number for Cardpresso mode
        GUICtrlSetData($processBtn, "Lookup & Load")
        GUICtrlSetState($nameInput, $GUI_ENABLE) ; Enable client number for Cardpresso
        GUICtrlSetState($cropBtn, $GUI_HIDE) ; Hide in Cardpresso mode
        GUICtrlSetState($manualSelectBtn, $GUI_HIDE) ; Hide in Cardpresso mode
        GUICtrlSetState($sftpTestBtn, $GUI_HIDE) ; Hide in Cardpresso mode
        GUICtrlSetState($cardpressoBtn, $GUI_SHOW) ; Show in Cardpresso mode
        GUICtrlSetState($multiCardpressoBtn, $GUI_SHOW) ; Show in Cardpresso mode
        GUICtrlSetState($processBtn, $GUI_ENABLE) ; Enable Lookup & Load button
        ; Hide source image file elements in Cardpresso mode
        GUICtrlSetState($srcInput, $GUI_HIDE)
        GUICtrlSetState($srcBrowseBtn, $GUI_HIDE)
        ; Hide destination folder elements in Cardpresso mode
        GUICtrlSetState($destInput, $GUI_HIDE)
        GUICtrlSetState($destBrowseBtn, $GUI_HIDE)
        ; Hide labels in Cardpresso mode
        GUICtrlSetState($hSrcLabel, $GUI_HIDE)
        GUICtrlSetState($hDestLabel, $GUI_HIDE)
        ; Reset face detection state
        $aDetectedFaces = 0
        $iSelectedFace = -1
        $bManualSelectionMode = False
        _ClearPreview()
        _ShowStatus("Select Cardpresso mode - Enter client number to lookup and export to CSV", 1)
    ElseIf $sMode = "Upload_to_SFTP" Then
        ; Clear previous source file and reset client number
        GUICtrlSetData($srcInput, "")
        GUICtrlSetData($nameInput, "") ; Reset client number for SFTP mode
        GUICtrlSetData($processBtn, "Upload to SFTP")
        GUICtrlSetState($nameInput, $GUI_DISABLE) ; Disable client number for SFTP - auto-extracted from filename
        GUICtrlSetState($cropBtn, $GUI_HIDE) ; Hide in SFTP mode
        GUICtrlSetState($manualSelectBtn, $GUI_HIDE) ; Hide in SFTP mode
        GUICtrlSetState($sftpTestBtn, $GUI_SHOW) ; Show in SFTP mode
        GUICtrlSetState($cardpressoBtn, $GUI_HIDE) ; Hide in SFTP mode
        GUICtrlSetState($multiCardpressoBtn, $GUI_HIDE) ; Hide in SFTP mode
        GUICtrlSetState($processBtn, $GUI_ENABLE) ; Enable Upload button
        ; Set Folder-3 for SFTP mode and make it read-only
        If $sFolder3Path <> "" Then
            GUICtrlSetData($destInput, $sFolder3Path)
        EndIf
        GUICtrlSetState($destInput, $GUI_DISABLE) ; Make destination folder read-only
        GUICtrlSetState($destBrowseBtn, $GUI_HIDE) ; Hide browse button in SFTP mode
        ; Reset face detection state
        $aDetectedFaces = 0
        $iSelectedFace = -1
        $bManualSelectionMode = False
        _ClearPreview()
        _ShowStatus("Select Upload_to_SFTP mode - Select files to upload (Client number auto-extracted from filename)", 1)
    Else
        ; Clear previous source file
        GUICtrlSetData($srcInput, "")
        ; Only reset client number if not preserving it (manual mode change)
        If Not $bPreserveClientNumber Then
            GUICtrlSetData($nameInput, "") ; Reset client number
        EndIf
        GUICtrlSetData($processBtn, "Copy & Resize")
        GUICtrlSetState($nameInput, $GUI_ENABLE)
        GUICtrlSetState($cropBtn, $GUI_HIDE) ; Hide in Copy & Resize mode
        GUICtrlSetState($manualSelectBtn, $GUI_HIDE) ; Hide in Copy & Resize mode
        GUICtrlSetState($sftpTestBtn, $GUI_HIDE) ; Hide in Copy & Resize mode
        GUICtrlSetState($cardpressoBtn, $GUI_HIDE) ; Hide in Copy & Resize mode
        GUICtrlSetState($multiCardpressoBtn, $GUI_HIDE) ; Hide in Copy & Resize mode
        GUICtrlSetState($processBtn, $GUI_ENABLE) ; Always enable the Copy & Resize button
        ; Re-enable source and destination elements for other modes
        GUICtrlSetState($srcInput, $GUI_SHOW)
        GUICtrlSetState($srcBrowseBtn, $GUI_SHOW)
        GUICtrlSetState($destInput, $GUI_ENABLE)
        GUICtrlSetState($destBrowseBtn, $GUI_SHOW)
        ; Show labels for other modes
        GUICtrlSetState($hSrcLabel, $GUI_SHOW)
        GUICtrlSetState($hDestLabel, $GUI_SHOW)
        
        ; Switch back to Folder-1 for resize operations
        _LoadConfig() ; This will reload Folder-1
        ; Reset face detection state
        $aDetectedFaces = 0
        $iSelectedFace = -1
        $bManualSelectionMode = False
        _ClearPreview()
        If $bPreserveClientNumber Then
            _ShowStatus("Switched to Resize_and_Copy mode. Client number preserved. Fill remaining fields and click 'Copy & Resize'.", 1)
        Else
            _ShowStatus("Select Resize_and_Copy mode - Fill all fields and click 'Copy & Resize'", 1)
        EndIf
    EndIf
    ; Update step guidance
    _HighlightNextStep()
EndFunc

Func _ProcessImageBasedOnMode()
    Local $sMode = GUICtrlRead($modeCombo)
    If $sMode = "Detect_Face_and_Crop" Then
        _RunFaceDetection()
    ElseIf $sMode = "Upload_to_SFTP" Then
        _RunSFTPUpload()
    ElseIf $sMode = "Cardpresso" Then
        _RunCardpresso()
    Else
        _RunBatch()
    EndIf
EndFunc

Func _RunBatch()
    Local $sSrc  = StringStripWS(GUICtrlRead($srcInput), 3)
    Local $sDest = StringStripWS(GUICtrlRead($destInput), 3)
    Local $sName = StringStripWS(GUICtrlRead($nameInput), 3)

    GUICtrlSetData($statusLabel, "")

    If $sSrc = "" Or $sDest = "" Or $sName = "" Then
        _ShowStatus("Please fill in all fields.", 0)
        Return
    EndIf
    If StringLen($sName) <> 5 Or Not StringIsDigit($sName) Then
        _ShowStatus("Name must be exactly 5 digits.", 0)
        Return
    EndIf

    If StringRight($sDest, 1) = "\" Then $sDest = StringTrimRight($sDest, 1)

    ; Convert to absolute paths
    $sSrc = _GetAbsolutePath($sSrc)
    $sDest = _GetAbsolutePath($sDest)
    
    If Not FileExists($sSrc) Then
        _ShowStatus("Source file not found: " & $sSrc, 0)
        Return
    EndIf
    If Not FileExists($sDest) Then
        _ShowStatus("Destination folder not found: " & $sDest, 0)
        Return
    EndIf

    ; Process the image directly in AutoIt3
    _ProcessImage($sSrc, $sDest, $sName)
EndFunc

Func _RunFaceDetection()
    Local $sSrc  = StringStripWS(GUICtrlRead($srcInput), 3)
    Local $sDest = StringStripWS(GUICtrlRead($destInput), 3)
    Local $sName = StringStripWS(GUICtrlRead($nameInput), 3)

    GUICtrlSetData($statusLabel, "")

    If $sSrc = "" Or $sDest = "" Or $sName = "" Then
        _ShowStatus("Please select source image, destination folder, and enter client number.", 0)
        Return
    EndIf
    If StringLen($sName) <> 5 Or Not StringIsDigit($sName) Then
        _ShowStatus("Client number must be exactly 5 digits.", 0)
        Return
    EndIf

    If StringRight($sDest, 1) = "\" Then $sDest = StringTrimRight($sDest, 1)

    ; Convert to absolute paths
    $sSrc = _GetAbsolutePath($sSrc)
    $sDest = _GetAbsolutePath($sDest)
    
    If Not FileExists($sSrc) Then
        _ShowStatus("Source file not found: " & $sSrc, 0)
        Return
    EndIf
    If Not FileExists($sDest) Then
        _ShowStatus("Destination folder not found: " & $sDest, 0)
        Return
    EndIf

    ; Build output filename using client number only
    $sCurrentOutputPath = $sDest & "\" & $sName & ".jpg"
    $sCurrentImagePath = $sSrc

    ; Call Face Detection
    _DetectFacesInPreview($sSrc)
EndFunc

Func _DetectFacesInPreview($sImagePath)
    ConsoleWrite("Starting Microsoft Face API detection..." & @CRLF)
    ConsoleWrite("Using Endpoint: " & $sFaceAPIEndpoint & @CRLF)

    ; Skip SSL testing and proceed directly with the API call
    ; SSL issues will be handled by the main HTTP function with fallback mechanisms
    ConsoleWrite("Skipping SSL test - proceeding with API call..." & @CRLF)

    ; ---- call API ----
    Local $aFaces = _DetectFacesWithMicrosoft($sImagePath)

    If @error Then
        Local $iError = @error
        ConsoleWrite("Face API Error: " & $iError & @CRLF)

        ; Unified error handling for all FaceAPI issues
        If $iError >= 1 And $iError <= 7 Then
            ; All FaceAPI related errors (network, API key, endpoint, etc.)
            MsgBox(16, "Face API Error", "Face API service unavailable." & @CRLF & @CRLF & _
                                  "Possible causes:" & @CRLF & _
                                  "• No internet connection" & @CRLF & _
                                  "• Wrong API key" & @CRLF & _
                                  "• Wrong endpoint URL" & @CRLF & _
                                  "• API service down" & @CRLF & @CRLF & _
                                  "Please verify:" & @CRLF & _
                                  "1. Your internet connection" & @CRLF & _
                                  "2. API key and endpoint in config.file" & @CRLF & _
                                  "3. Azure Face API service status" & @CRLF & @CRLF & _
                                  "Using manual face selection instead.")
        Else
            ; Other non-API errors
            MsgBox(16, "Error", "Failed to call Face API. Error code: " & $iError & @CRLF & @CRLF & "Using manual face selection instead.")
        EndIf

        _ShowStatus("No internet connection - use Manual Face Select", 0)
        
        ; Enable manual selection as fallback
        GUICtrlSetState($manualSelectBtn, $GUI_ENABLE)
        _ShowStatus("No Auto Face detection available - use Manual Face Select to choose face area", 0)
        Return
    EndIf

    ; ---- no faces found ----
    If UBound($aFaces) = 0 Then
        MsgBox(48, "No Faces", "Microsoft AI did not detect any faces in the image." & @CRLF & "Please use Manual Face Select to choose the face area manually.")
        _ShowStatus("No faces detected - use Manual Face Select", 0)
        
        ; Enable manual selection as fallback
        GUICtrlSetState($manualSelectBtn, $GUI_ENABLE)
        Return
    EndIf

    ; Store detected faces
    $aDetectedFaces = $aFaces
    $iSelectedFace = -1  ; Reset selection

    ; Disable Detect Face button and enable Crop Face button (Manual Select is always enabled)
    GUICtrlSetState($processBtn, $GUI_DISABLE)
    GUICtrlSetState($cropBtn, $GUI_ENABLE)

    ; Update preview with face detection
    _UpdatePreviewWithFaces($sImagePath, $aFaces)
    _ShowStatus("Faces detected: " & UBound($aFaces) & " - Click on a face to select it or use Manual Face Select", 1)
EndFunc

; SSL/TLS connectivity testing removed due to system SSL issues
; The main HTTP functions now handle SSL errors with fallback mechanisms

Func _LoadIrfanViewPath()
    Local $sConfig = $sScriptDir & "\config.file"
    If Not FileExists($sConfig) Then Return
    
    Local $aLines = FileReadToArray($sConfig)
    For $i = 0 To UBound($aLines) - 1
        Local $sLine = StringStripWS($aLines[$i], 3)
        If StringInStr($sLine, "IrfanViewPath") Then
            Local $sPath = StringRegExpReplace($sLine, '(?i)^.*=\s*"?([^"]+)"?\s*$', "$1")
            ; Expand environment variables in the path
            $sPath = _ExpandEnvironmentVariables($sPath)
            $sIrfanViewPath = $sPath
            ExitLoop
        EndIf
    Next
EndFunc

Func _LoadFaceAPIConfig()
    Local $sConfig = $sScriptDir & "\config.file"
    If Not FileExists($sConfig) Then Return
    
    Local $aLines = FileReadToArray($sConfig)
    For $i = 0 To UBound($aLines) - 1
        Local $sLine = StringStripWS($aLines[$i], 3)
        If StringInStr($sLine, "FaceAPIKey") Then
            $sFaceAPIKey = StringRegExpReplace($sLine, '(?i)^.*=\s*"?([^"]+)"?\s*$', "$1")
        ElseIf StringInStr($sLine, "FaceAPIEndpoint") Then
            $sFaceAPIEndpoint = StringRegExpReplace($sLine, '(?i)^.*=\s*"?([^"]+)"?\s*$', "$1")
        EndIf
    Next
EndFunc

Func _LoadSFTPConfig()
    Local $sConfig = $sScriptDir & "\config.file"
    If Not FileExists($sConfig) Then Return
    
    Local $aLines = FileReadToArray($sConfig)
    For $i = 0 To UBound($aLines) - 1
        Local $sLine = StringStripWS($aLines[$i], 3)
        If StringInStr($sLine, "SFTPHost") Then
            $sSFTPHost = StringRegExpReplace($sLine, '(?i)^.*=\s*"?([^"]+)"?\s*$', "$1")
        ElseIf StringInStr($sLine, "SFTPUsername") Then
            $sSFTPUsername = StringRegExpReplace($sLine, '(?i)^.*=\s*"?([^"]+)"?\s*$', "$1")
        ElseIf StringInStr($sLine, "SFTPRemotePath") Then
            $sSFTPRemotePath = StringRegExpReplace($sLine, '(?i)^.*=\s*"?([^"]+)"?\s*$', "$1")
        EndIf
    Next
EndFunc

Func _LoadCardpressoConfig()
    Local $sConfig = $sScriptDir & "\config.file"
    If Not FileExists($sConfig) Then Return
    
    Local $aLines = FileReadToArray($sConfig)
    For $i = 0 To UBound($aLines) - 1
        Local $sLine = StringStripWS($aLines[$i], 3)
        If StringInStr($sLine, "CardpressoMasterFile") Then
            $sCardpressoMasterFile = StringRegExpReplace($sLine, '(?i)^.*=\s*"?([^"]+)"?\s*$', "$1")
            ; Expand environment variables in the path
            $sCardpressoMasterFile = _ExpandEnvironmentVariables($sCardpressoMasterFile)
        ElseIf StringInStr($sLine, "CardpressoCSVFile") Then
            $sCardpressoCSVFile = StringRegExpReplace($sLine, '(?i)^.*=\s*"?([^"]+)"?\s*$', "$1")
            ; Expand environment variables in the path
            $sCardpressoCSVFile = _ExpandEnvironmentVariables($sCardpressoCSVFile)
        ElseIf StringInStr($sLine, "CardpressoExePath") Then
            $sCardpressoExePath = StringRegExpReplace($sLine, '(?i)^.*=\s*"?([^"]+)"?\s*$', "$1")
            ; Expand environment variables in the path
            $sCardpressoExePath = _ExpandEnvironmentVariables($sCardpressoExePath)
        ElseIf StringInStr($sLine, "CardpressoSheetName") Then
            $sCardpressoSheetName = StringRegExpReplace($sLine, '(?i)^.*=\s*"?([^"]+)"?\s*$', "$1")
        ElseIf StringInStr($sLine, "CardpressoPhotoBaseDir") Then
            $sCardpressoPhotoBaseDir = StringRegExpReplace($sLine, '(?i)^.*=\s*"?([^"]+)"?\s*$', "$1")
        EndIf
    Next
EndFunc

Func _GetAbsolutePath($sPath)
    ; Convert relative path to absolute path
    If StringLeft($sPath, 1) = "." Then
        ; Relative path starting with .\
        Return $sScriptDir & "\" & StringTrimLeft($sPath, 2)
    ElseIf Not StringInStr($sPath, ":") And StringLeft($sPath, 2) <> "\\" Then
        ; Relative path without drive letter or UNC
        Return $sScriptDir & "\" & $sPath
    Else
        ; Already absolute path
        Return $sPath
    EndIf
EndFunc

Func _ProcessImage($sSrc, $sDest, $sName)
    ; Build output filename
    Local $sOutFile = $sDest & "\" & $sName & ".jpg"
    
    ; Check if output file already exists
    If FileExists($sOutFile) Then
        ; Ask user for overwrite confirmation (modal to main window)
        Local $iResponse = MsgBox(4 + 32 + 262144, "File Exists", "File " & $sName & ".jpg already exists." & @CRLF & "Do you want to overwrite it?")
        
        If $iResponse <> 6 Then ; 6 = Yes
            _ShowStatus("Operation cancelled by user.", 0)
            Return
        EndIf
    EndIf
    
    ; Get IrfanView path from config
    If $sIrfanViewPath = "" Then
        _LoadIrfanViewPath()
    EndIf
    
    ; Check if IrfanView is available
    If Not FileExists($sIrfanViewPath) Then
        _ShowStatus("IrfanView not found at: " & $sIrfanViewPath, 0)
        Return
    EndIf
    
    ; Build IrfanView command for resize and conversion
    Local $sCmd = '"' & $sIrfanViewPath & '" "' & $sSrc & '" /resize_long=800 /aspectratio /resample /convert="' & $sOutFile & '" /jpgq=85'
    
    ; Show progress bar
    GUICtrlSetState($progressBar, $GUI_SHOW)
    GUICtrlSetData($progressBar, 0)
    _ShowStatus("Starting image processing...", 1)
    
    ; Run IrfanView
    Local $iPID = Run($sCmd, "", @SW_HIDE)
    If @error Then
        GUICtrlSetState($progressBar, $GUI_HIDE)
        _ShowStatus("Failed to start IrfanView.", 0)
        Return
    EndIf
    
    ; Simulate progress for longer operations
    For $i = 10 To 90 Step 20
        GUICtrlSetData($progressBar, $i)
        Sleep(100)
        If Not ProcessExists($iPID) Then ExitLoop
    Next
    
    ; Wait for process to complete
    ProcessWaitClose($iPID)
    Local $iExitCode = @extended
    
    ; Complete progress
    GUICtrlSetData($progressBar, 100)
    Sleep(500)
    GUICtrlSetState($progressBar, $GUI_HIDE)
    
    ; Check if output file was created
    If FileExists($sOutFile) Then
        _ShowStatus("Image processed successfully: " & $sName & ".jpg", 1)
        ; Log the resize action
        _LogAction("Resized", $sOutFile)
        
        ; Ask user if they want to proceed to face detection
        Local $iResponse = MsgBox(4 + 32 + 262144, "Copy & Resize Completed", "Copy & Resize has been completed successfully." & @CRLF & @CRLF & "Do you want to move to 'Face Detect & Crop' task?")
        
        If $iResponse = 6 Then ; 6 = Yes
            ; Switch to Detect_Face_and_Crop mode
            GUICtrlSetData($modeCombo, "Detect_Face_and_Crop")
            _UpdateModeUI(True) ; Preserve client number when automatically switching
            
            ; Set the resized image as source for face detection
            GUICtrlSetData($srcInput, $sOutFile)
            _UpdatePreview()
            
            ; Client number is preserved from Copy & Resize mode
            _ShowStatus("Switched to Detect_Face_and_Crop mode. Client number preserved. Click 'Detect Face' to continue.", 1)
        EndIf
    Else
        _ShowStatus("Failed to process image. Check if IrfanView is working properly.", 0)
    EndIf
EndFunc

; ==============================================================
; Face Detection Functions (from Face_Detect_and_Crop.au3)
; ==============================================================
Func _DetectAndCropFace($sImagePath, $sOutputPath)
    ConsoleWrite("Starting Microsoft Face API detection..." & @CRLF)
    ConsoleWrite("Using Endpoint: " & $sFaceAPIEndpoint & @CRLF)

    ; ---- call API ----
    Local $aFaces = _DetectFacesWithMicrosoft($sImagePath)

    If @error Then
        Local $iError = @error
        ConsoleWrite("Face API Error: " & $iError & @CRLF)

        ; Unified error handling for all FaceAPI issues
        If $iError >= 1 And $iError <= 7 Then
            ; All FaceAPI related errors (network, API key, endpoint, etc.)
            MsgBox(16, "Face API Error", "Face API service unavailable." & @CRLF & @CRLF & _
                                  "Possible causes:" & @CRLF & _
                                  "• No internet connection" & @CRLF & _
                                  "• Wrong API key" & @CRLF & _
                                  "• Wrong endpoint URL" & @CRLF & _
                                  "• API service down" & @CRLF & @CRLF & _
                                  "Please verify:" & @CRLF & _
                                  "1. Your internet connection" & @CRLF & _
                                  "2. API key and endpoint in config.file" & @CRLF & _
                                  "3. Azure Face API service status" & @CRLF & @CRLF & _
                                  "Using center crop instead.")
        Else
            ; Other non-API errors
            MsgBox(16, "Error", "Failed to call Face API. Error code: " & $iError & @CRLF & "Using center crop instead.")
        EndIf

        ; ---- fallback to center crop ----
        If _CenterCrop($sImagePath, $sOutputPath) Then
            _ShowStatus("Center crop successful: " & $sOutputPath, 1)
            _AddToRecentFiles($sImagePath)
        Else
            _ShowStatus("Failed to crop image", 0)
        EndIf
        Return
    EndIf

    ; ---- no faces found ----
    If UBound($aFaces) = 0 Then
        MsgBox(48, "No Faces", "Microsoft AI did not detect any faces in the image." & @CRLF & "Using center crop instead.")
        If _CenterCrop($sImagePath, $sOutputPath) Then
            _ShowStatus("Center crop successful: " & $sOutputPath, 1)
            _LogAction("Cropped", $sOutputPath)
        EndIf
        Return
    EndIf

    ; Store global variables for preview
    $aDetectedFaces = $aFaces
    $sCurrentImagePath = $sImagePath
    $sCurrentOutputPath = $sOutputPath

    ; ---- Show interactive preview with face selection ----
    _ShowInteractivePreview($sImagePath, $aFaces)
EndFunc

Func _ShowInteractivePreview($sImagePath, $aFaces)
    ; Load original image
    Local $hOriginalImage = _GDIPlus_ImageLoadFromFile($sImagePath)
    Local $iImgWidth = _GDIPlus_ImageGetWidth($hOriginalImage)
    Local $iImgHeight = _GDIPlus_ImageGetHeight($hOriginalImage)
    
    ; Calculate preview size (max 800x600 while maintaining aspect ratio)
    Local $iPreviewWidth, $iPreviewHeight
    Local $fAspectRatio = $iImgWidth / $iImgHeight
    
    If $iImgWidth > 800 Or $iImgHeight > 600 Then
        If $fAspectRatio > (800/600) Then
            $iPreviewWidth = 800
            $iPreviewHeight = 800 / $fAspectRatio
        Else
            $iPreviewHeight = 600
            $iPreviewWidth = 600 * $fAspectRatio
        EndIf
    Else
        $iPreviewWidth = $iImgWidth
        $iPreviewHeight = $iImgHeight
    EndIf
    
    ; Calculate scale factors
    Local $fScaleX = $iPreviewWidth / $iImgWidth
    Local $fScaleY = $iPreviewHeight / $iImgHeight
    
    ; Create preview GUI
    $hPreviewGUI = GUICreate("Face Detection - Click Face to Select", $iPreviewWidth + 20, $iPreviewHeight + 120)
    GUISetBkColor(0xF0F0F0)
    
    ; Create image control (make it clickable)
    $hPreviewImage = GUICtrlCreatePic("", 10, 10, $iPreviewWidth, $iPreviewHeight)
    GUICtrlSetCursor(-1, 2) ; Set cursor to hand when hovering over image
    
    ; Create buttons
    $hCropButton = GUICtrlCreateButton("Crop Face", 10, $iPreviewHeight + 20, 150, 30)
    GUICtrlSetState($hCropButton, $GUI_DISABLE) ; Initially disabled until face is selected
    Local $hCancelButton = GUICtrlCreateButton("Cancel", 170, $iPreviewHeight + 20, 80, 30)
    
    ; Create status label
    Local $hStatusLabel = GUICtrlCreateLabel("Faces detected: " & UBound($aFaces) & " - Click the Face you want to crop", 10, $iPreviewHeight + 60, $iPreviewWidth, 20)
    Local $hSelectionLabel = GUICtrlCreateLabel("No face selected", 10, $iPreviewHeight + 85, $iPreviewWidth, 20)
    
    ; Create preview image with face rectangles and store face regions
    _CreateInteractivePreviewImage($sImagePath, $aFaces, $iPreviewWidth, $iPreviewHeight, $fScaleX, $fScaleY)
    
    GUISetState(@SW_SHOW)
    
    ; Event loop
    While 1
        Local $nMsg = GUIGetMsg()
        Switch $nMsg
            Case $GUI_EVENT_CLOSE, $hCancelButton
                GUIDelete($hPreviewGUI)
                ExitLoop
            Case $hCropButton
                ; Crop the selected face
                If _CropDetectedFace($sCurrentImagePath, $sCurrentOutputPath, $aDetectedFaces[$iSelectedFace]) Then
                    Local $sMsg = "AI Face Detection Successful!" & @CRLF & _
                                 "Faces detected: " & UBound($aDetectedFaces) & @CRLF & _
                                 "Cropped selected face to 1:1 ratio" & @CRLF & _
                                 "Saved: " & StringRegExpReplace($sCurrentOutputPath, "^.*\\", "")

                    If UBound($aDetectedFaces) > 1 Then
                        $sMsg &= @CRLF & @CRLF & "Note: " & UBound($aDetectedFaces) - 1 & " additional face(s) were detected but not cropped."
                    EndIf
                    MsgBox(64, "Success", $sMsg)
                    _ShowStatus("Face crop successful: " & StringRegExpReplace($sCurrentOutputPath, "^.*\\", ""), 1)
                    _LogAction("Cropped", $sCurrentOutputPath)
                Else
                    MsgBox(16, "Error", "Failed to crop selected face")
                    _ShowStatus("Failed to crop face", 0)
                EndIf
                GUIDelete($hPreviewGUI)
                ExitLoop
            Case $hPreviewImage
                ; Handle click on image
                Local $aMousePos = GUIGetCursorInfo($hPreviewGUI)
                If IsArray($aMousePos) Then
                    Local $iMouseX = $aMousePos[0] - 10  ; Adjust for image position
                    Local $iMouseY = $aMousePos[1] - 10
                    
                    ; Check if click is within any face region
                    ConsoleWrite("Mouse Click: X=" & $iMouseX & ", Y=" & $iMouseY & @CRLF)
                    ConsoleWrite("Face regions count: " & UBound($aFaceRegions) & @CRLF)
                    
                    For $i = 0 To UBound($aFaceRegions) - 1
                        Local $sRegion = $aFaceRegions[$i]
                        ConsoleWrite("Checking Face " & ($i + 1) & ": " & $sRegion & @CRLF)
                        Local $aCoords = StringSplit($sRegion, "|", 2)
                        If IsArray($aCoords) And UBound($aCoords) = 4 Then
                            ConsoleWrite("Face " & ($i + 1) & " region: X1=" & $aCoords[0] & ", Y1=" & $aCoords[1] & ", X2=" & $aCoords[2] & ", Y2=" & $aCoords[3] & @CRLF)
                            If $iMouseX >= $aCoords[0] And $iMouseX <= $aCoords[2] And _
                               $iMouseY >= $aCoords[1] And $iMouseY <= $aCoords[3] Then
                                ConsoleWrite("Face " & ($i + 1) & " SELECTED!" & @CRLF)
                                $iSelectedFace = $i
                                GUICtrlSetData($hSelectionLabel, "Selected: Face " & ($i + 1))
                                ; Enable the crop button and change its text
                                GUICtrlSetState($hCropButton, $GUI_ENABLE)
                                GUICtrlSetData($hCropButton, "Crop Face")
                                ; Update preview to show selection
                                _CreateInteractivePreviewImage($sCurrentImagePath, $aDetectedFaces, $iPreviewWidth, $iPreviewHeight, $fScaleX, $fScaleY, $iSelectedFace)
                                ExitLoop
                            EndIf
                        EndIf
                    Next
                EndIf
        EndSwitch
    WEnd
    
    _GDIPlus_ImageDispose($hOriginalImage)
EndFunc

Func _CreateInteractivePreviewImage($sImagePath, $aFaces, $iPreviewWidth, $iPreviewHeight, $fScaleX, $fScaleY, $iSelectedIndex = -1)
    ; Load original image
    Local $hOriginalImage = _GDIPlus_ImageLoadFromFile($sImagePath)
    
    ; Create bitmap for preview
    Local $hBitmap = _GDIPlus_BitmapCreateFromScan0($iPreviewWidth, $iPreviewHeight)
    Local $hGraphics = _GDIPlus_ImageGetGraphicsContext($hBitmap)
    
    ; Draw original image scaled to preview size
    _GDIPlus_GraphicsDrawImageRect($hGraphics, $hOriginalImage, 0, 0, $iPreviewWidth, $iPreviewHeight)
    
    ; Clear face regions array
    Global $aFaceRegions[0]
    
    ; Draw rectangles around detected faces
    For $i = 0 To UBound($aFaces) - 1
        Local $aFace = $aFaces[$i]
        Local $iFaceX = $aFace[0] * $fScaleX
        Local $iFaceY = $aFace[1] * $fScaleY
        Local $iFaceWidth = $aFace[2] * $fScaleX
        Local $iFaceHeight = $aFace[3] * $fScaleY
        
        ; Store face region for click detection (expanded slightly for easier clicking)
        Local $iExpandedX = _Max(0, $iFaceX - 5)
        Local $iExpandedY = _Max(0, $iFaceY - 5)
        Local $iExpandedWidth = _Min($iPreviewWidth - $iExpandedX, $iFaceWidth + 10)
        Local $iExpandedHeight = _Min($iPreviewHeight - $iExpandedY, $iFaceHeight + 10)
        
        Local $sRegion = $iExpandedX & "|" & $iExpandedY & "|" & ($iExpandedX + $iExpandedWidth) & "|" & ($iExpandedY + $iExpandedHeight)
        ConsoleWrite("Adding Face " & ($i + 1) & " region: " & $sRegion & @CRLF)
        
        ; Manually add to array since _ArrayAdd might not work properly
        Local $iSize = UBound($aFaceRegions)
        ReDim $aFaceRegions[$iSize + 1]
        $aFaceRegions[$iSize] = $sRegion
        
        ; Choose pen color based on selection
        Local $iPenColor, $iBrushColor
        If $i = $iSelectedIndex Then
            $iPenColor = 0xFF00FF00  ; Green for selected face
            $iBrushColor = 0xFF00FF00
        Else
            $iPenColor = 0xFFFF0000  ; Red for unselected faces
            $iBrushColor = 0xFFFF0000
        EndIf
        
        ; Create pen for face rectangles
        Local $hPen = _GDIPlus_PenCreate($iPenColor, 3)
        
        ; Draw rectangle
        _GDIPlus_GraphicsDrawRect($hGraphics, $iFaceX, $iFaceY, $iFaceWidth, $iFaceHeight, $hPen)
        
        ; Add face number label
        Local $hBrush = _GDIPlus_BrushCreateSolid($iBrushColor)
        Local $hFont = _GDIPlus_FontCreate(_GDIPlus_FontFamilyCreate("Arial"), 14)
        Local $hFormat = _GDIPlus_StringFormatCreate()
        _GDIPlus_GraphicsDrawString($hGraphics, "Face " & ($i + 1), $iFaceX + 5, $iFaceY + 5, $hFont, $hFormat, $hBrush)
        
        _GDIPlus_BrushDispose($hBrush)
        _GDIPlus_FontDispose($hFont)
        _GDIPlus_StringFormatDispose($hFormat)
        _GDIPlus_PenDispose($hPen)
    Next
    
    ; Save preview to temporary file
    Local $sTempFile = @TempDir & "\face_preview_" & @YEAR & @MON & @MDAY & "_" & @HOUR & @MIN & @SEC & ".jpg"
    _GDIPlus_ImageSaveToFile($hBitmap, $sTempFile)
    
    ; Update the picture control
    GUICtrlSetImage($hPreviewImage, $sTempFile)
    
    ; Cleanup
    _GDIPlus_GraphicsDispose($hGraphics)
    _GDIPlus_BitmapDispose($hBitmap)
    _GDIPlus_ImageDispose($hOriginalImage)
    
    ; Delete temp file on exit
    OnAutoItExitRegister("_DeleteTempPreview")
EndFunc

Func _DetectFacesWithMicrosoft($sImagePath)
    ConsoleWrite("Calling Microsoft Face API..." & @CRLF)
    ConsoleWrite("Using Endpoint: " & $sFaceAPIEndpoint & @CRLF)

    Local $hFile = FileOpen($sImagePath, 16) ; binary
    If $hFile = -1 Then
        ConsoleWrite("Error: Could not open image file" & @CRLF)
        Return SetError(3, 0, 0)
    EndIf
    Local $bImageData = FileRead($hFile)
    FileClose($hFile)

    ; Build API URL - Remove recognition-related parameters to avoid approval requirements
    Local $sApiUrl = $sFaceAPIEndpoint & "face/v1.0/detect" & _
                    "?returnFaceId=false" & _  ; Disables face recognition
                    "&returnFaceLandmarks=false" & _
                    "&returnFaceAttributes=" & _
                    "&detectionModel=detection_03"  ; Keep newer detection model

    ConsoleWrite("API URL: " & $sApiUrl & @CRLF)

    Local $sResponse = _INetPost($sApiUrl, $bImageData, "application/octet-stream", $sFaceAPIKey)
    If @error Then
        ConsoleWrite("INetPost Error: " & @error & @CRLF)
        Return SetError(2, 0, 0)
    EndIf

    If StringInStr($sResponse, "error") Or $sResponse = "" Then
        ConsoleWrite("API Error Response: " & $sResponse & @CRLF)
        Return SetError(1, 0, 0)
    EndIf

    Local $aFaces = _ParseFaceResponse($sResponse)
    Return $aFaces
EndFunc   ;==>_DetectFacesWithMicrosoft


; ------------------------------------------------------------------
;  POST via WinHttp (with proper error checking)
; ------------------------------------------------------------------
Func _INetPost($sURL, $vData, $sContentType = "application/octet-stream", $sApiKey = "")
    ; Simple network connectivity test before attempting HTTP request
    Local $iPingResult = Ping("8.8.8.8", 1000) ; Ping Google DNS with 1 second timeout
    If $iPingResult = 0 Then
        ConsoleWrite("Network connectivity test failed - no internet connection" & @CRLF)
        Return SetError(2, 0, "") ; Network error
    EndIf
    
    Local $oHTTP = ObjCreate("WinHttp.WinHttpRequest.5.1")
    If @error Then
        ConsoleWrite("Error creating HTTP object: " & @error & @CRLF)
        Return SetError(1, 0, "")
    EndIf
    
    ; Use error trapping for the Open method to catch network failures
    Local $iOpenError = 0
    $oHTTP.Open("POST", $sURL, False)
    If @error Then
        $iOpenError = @error
        ConsoleWrite("Error opening HTTP connection: " & $iOpenError & @CRLF)
        ; Check if this is a network connectivity error
        If $iOpenError = -2147012867 Or $iOpenError = -2147012889 Then ; Common network errors
            Return SetError(2, $iOpenError, "") ; Network error
        Else
            Return SetError(2, $iOpenError, "") ; Other open error
        EndIf
    EndIf
    
    $oHTTP.SetRequestHeader("Content-Type", $sContentType)
    If @error Then
        ConsoleWrite("Error setting Content-Type header: " & @error & @CRLF)
        Return SetError(3, 0, "")
    EndIf
    
    $oHTTP.SetRequestHeader("Ocp-Apim-Subscription-Key", $sApiKey)
    If @error Then
        ConsoleWrite("Error setting API key header: " & @error & @CRLF)
        Return SetError(4, 0, "")
    EndIf
    
    ; Set longer timeout (30 seconds)
    $oHTTP.SetTimeouts(30000, 30000, 30000, 30000)
    
    ; Use error trapping for the Send method to catch network failures
    Local $iSendError = 0
    $oHTTP.Send($vData)
    If @error Then
        $iSendError = @error
        ConsoleWrite("Error sending HTTP request: " & $iSendError & @CRLF)
        ; Check if this is a network connectivity error
        If $iSendError = -2147012867 Or $iSendError = -2147012889 Then ; Common network errors
            Return SetError(2, $iSendError, "") ; Network error
        Else
            Return SetError(5, $iSendError, "") ; Other send error
        EndIf
    EndIf

    ConsoleWrite("HTTP Status: " & $oHTTP.Status & " " & $oHTTP.StatusText & @CRLF)

    If $oHTTP.Status = 200 Then
        Return $oHTTP.ResponseText
    Else
        ConsoleWrite("HTTP Error: " & $oHTTP.Status & " - " & $oHTTP.StatusText & @CRLF)
        ConsoleWrite("Response: " & $oHTTP.ResponseText & @CRLF)
        
        ; Enhanced error classification based on HTTP status codes
        ; API configuration errors (wrong key, endpoint, quota, etc.)
        If $oHTTP.Status = 401 Or $oHTTP.Status = 403 Or $oHTTP.Status = 404 Or $oHTTP.Status = 429 Then
            Return SetError(6, $oHTTP.Status, "") ; API configuration error
        Else
            Return SetError(7, $oHTTP.Status, "") ; Other HTTP error (network/server issues)
        EndIf
    EndIf
EndFunc   ;==>_INetPost

; ------------------------------------------------------------------
;  crude JSON extractor
; ------------------------------------------------------------------
Func _ParseFaceResponse($sJsonResponse)
    Local $aFaces[0]
    
    ; Validate JSON response
    If StringLeft($sJsonResponse, 1) <> "[" Then
        ConsoleWrite("Invalid JSON response: " & $sJsonResponse & @CRLF)
        Return $aFaces
    EndIf
    
    Local $aFaceBlocks = StringSplit($sJsonResponse, '{"faceRectangle"', 1)
    If $aFaceBlocks[0] < 2 Then 
        ConsoleWrite("No faces found in response" & @CRLF)
        Return $aFaces
    EndIf

    For $i = 2 To $aFaceBlocks[0]
        Local $sFaceBlock = $aFaceBlocks[$i]
        Local $iLeft   = _ExtractJsonValue($sFaceBlock, "left")
        Local $iTop    = _ExtractJsonValue($sFaceBlock, "top")
        Local $iWidth  = _ExtractJsonValue($sFaceBlock, "width")
        Local $iHeight = _ExtractJsonValue($sFaceBlock, "height")

        If $iLeft <> "" And $iTop <> "" And $iWidth <> "" And $iHeight <> "" Then
            ReDim $aFaces[UBound($aFaces) + 1]
            Local $aFace[4] = [$iLeft, $iTop, $iWidth, $iHeight]
            $aFaces[UBound($aFaces) - 1] = $aFace

            ConsoleWrite("Detected Face: X=" & $iLeft & ", Y=" & $iTop & _
                        ", W=" & $iWidth & ", H=" & $iHeight & @CRLF)
        EndIf
    Next
    Return $aFaces
EndFunc   ;==>_ParseFaceResponse

Func _ExtractJsonValue($sJson, $sKey)
    Local $sPattern = '"' & $sKey & '":\s*(\d+)'
    Local $aMatch = StringRegExp($sJson, $sPattern, 1)
    If @error Then Return ""
    Return $aMatch[0]
EndFunc   ;==>_ExtractJsonValue

; ------------------------------------------------------------------
;  Crop around the detected face (1:1)
; ------------------------------------------------------------------
Func _CropDetectedFace($sImagePath, $sOutputPath, $aFace)
    _GDIPlus_Startup()
    Local $hImage = _GDIPlus_ImageLoadFromFile($sImagePath)
    If @error Then
        _GDIPlus_Shutdown()
        Return False
    EndIf

    Local $iImgWidth  = _GDIPlus_ImageGetWidth($hImage)
    Local $iImgHeight = _GDIPlus_ImageGetHeight($hImage)

    Local $iFaceX      = $aFace[0]
    Local $iFaceY      = $aFace[1]
    Local $iFaceWidth  = $aFace[2]
    Local $iFaceHeight = $aFace[3]

    Local $iCenterX     = $iFaceX + ($iFaceWidth / 2)
    Local $iCenterY     = $iFaceY + ($iFaceHeight / 2)
    Local $iSquareSize  = _Max($iFaceWidth, $iFaceHeight) * 1.6 ; 60 % padding

    Local $iCropX = $iCenterX - ($iSquareSize / 2)
    Local $iCropY = $iCenterY - ($iSquareSize / 2)

    $iCropX        = _Max(0, $iCropX)
    $iCropY        = _Max(0, $iCropY)
    
    ; Fixed: Use Min3 function for three values
    $iSquareSize   = _Min3($iSquareSize, $iImgWidth - $iCropX, $iImgHeight - $iCropY)

    If $iSquareSize < 10 Then
        _GDIPlus_Shutdown()
        Return False
    EndIf

    Local $hCropped = _GDIPlus_BitmapCloneArea($hImage, $iCropX, $iCropY, $iSquareSize, $iSquareSize)
    Local $bResult  = _GDIPlus_ImageSaveToFile($hCropped, $sOutputPath)

    _GDIPlus_ImageDispose($hImage)
    _GDIPlus_ImageDispose($hCropped)
    _GDIPlus_Shutdown()

    If $bResult Then ConsoleWrite("Cropped to: " & $iSquareSize & "x" & $iSquareSize & @CRLF)
    Return $bResult
EndFunc   ;==>_CropDetectedFace

; ------------------------------------------------------------------
;  Simple center crop
; ------------------------------------------------------------------
Func _CenterCrop($sInputPath, $sOutputPath)
    _GDIPlus_Startup()
    Local $hImage = _GDIPlus_ImageLoadFromFile($sInputPath)
    If @error Then 
        _GDIPlus_Shutdown()
        Return False
    EndIf

    Local $iWidth  = _GDIPlus_ImageGetWidth($hImage)
    Local $iHeight = _GDIPlus_ImageGetHeight($hImage)
    
    ; Simple center crop to 1:1 ratio
    Local $iSize = _Min($iWidth, $iHeight)
    Local $iX = ($iWidth - $iSize) / 2
    Local $iY = ($iHeight - $iSize) / 2

    $iX = _Max(0, $iX)
    $iY = _Max(0, $iY)
    
    ; Fixed: Use Min3 function for three values
    $iSize = _Min3($iSize, $iWidth - $iX, $iHeight - $iY)

    Local $hCropped = _GDIPlus_BitmapCloneArea($hImage, $iX, $iY, $iSize, $iSize)
    Local $bResult  = _GDIPlus_ImageSaveToFile($hCropped, $sOutputPath)

    _GDIPlus_ImageDispose($hImage)
    _GDIPlus_ImageDispose($hCropped)
    _GDIPlus_Shutdown()
    
    Return $bResult
EndFunc   ;==>_CenterCrop

; ------------------------------------------------------------------
;  Min / Max helpers - FIXED VERSIONS
; ------------------------------------------------------------------
Func _Max($a, $b)
    Return $a > $b ? $a : $b
EndFunc   ;==>_Max

Func _Min($a, $b)
    Return $a < $b ? $a : $b
EndFunc   ;==>_Min

; NEW: Min function that accepts three parameters
Func _Min3($a, $b, $c)
    Local $iMin = $a
    If $b < $iMin Then $iMin = $b
    If $c < $iMin Then $iMin = $c
    Return $iMin
EndFunc   ;==>_Min3

Func _DeleteTempPreview()
    ; Clean up temporary preview files
    Local $aFiles = _FileListToArray(@TempDir, "face_preview_*.jpg", 1)
    If IsArray($aFiles) Then
        For $i = 1 To $aFiles[0]
            FileDelete(@TempDir & "\" & $aFiles[$i])
        Next
    EndIf
EndFunc

; ==============================================================
; Common Functions (from Copy_and_Resize.au3)
; ==============================================================
Func _UpdatePreview()
    Local $sMode = GUICtrlRead($modeCombo)
    Local $sSrc = StringStripWS(GUICtrlRead($srcInput), 3)
    
    ; Handle SFTP mode with multiple files - show first image as preview
    If $sMode = "Upload_to_SFTP" And StringInStr($sSrc, "|") Then
        Local $aFiles = StringSplit($sSrc, "|", 2) ; No count flag
        If UBound($aFiles) > 0 Then
            ; Use the first file for preview
            $sSrc = $aFiles[0]
        EndIf
    EndIf
    
    If $sSrc = "" Or Not FileExists($sSrc) Then
        _ClearPreview()
        Return
    EndIf
    
    ; Validate file type
    Local $sExt = StringLower(StringTrimLeft($sSrc, StringInStr($sSrc, ".", 0, -1)))
    If $sExt <> "jpg" And $sExt <> "jpeg" And $sExt <> "png" Then
        _ClearPreview()
        Return
    EndIf
    
    ; Clear previous preview
    _ClearPreview()
    
    ; Load image using GDI+
    Local $hImage = _GDIPlus_ImageLoadFromFile($sSrc)
    If @error Then
        _ClearPreview()
        Return
    EndIf
    
    ; Get image dimensions
    Local $iWidth = _GDIPlus_ImageGetWidth($hImage)
    Local $iHeight = _GDIPlus_ImageGetHeight($hImage)
    
    ; Calculate scaled dimensions to fit 300x300 preview
    Local $iNewWidth, $iNewHeight
    If $iWidth > $iHeight Then
        $iNewWidth = 300
        $iNewHeight = Int($iHeight * 300 / $iWidth)
    Else
        $iNewHeight = 300
        $iNewWidth = Int($iWidth * 300 / $iHeight)
    EndIf
    
    ; Create bitmap for preview
    Local $hBitmap = _GDIPlus_BitmapCreateFromScan0($iNewWidth, $iNewHeight)
    Local $hGraphics = _GDIPlus_ImageGetGraphicsContext($hBitmap)
    _GDIPlus_GraphicsSetInterpolationMode($hGraphics, $GDIP_INTERPOLATIONMODE_HIGHQUALITYBICUBIC)
    _GDIPlus_GraphicsDrawImageRect($hGraphics, $hImage, 0, 0, $iNewWidth, $iNewHeight)
    
    ; Save bitmap to file for Pic control
    Local $sTempBMP = @TempDir & "\preview_" & @MSEC & ".bmp"
    _GDIPlus_ImageSaveToFile($hBitmap, $sTempBMP)
    
    ; Cleanup GDI+ objects
    _GDIPlus_GraphicsDispose($hGraphics)
    _GDIPlus_ImageDispose($hBitmap)
    _GDIPlus_ImageDispose($hImage)
    
    ; Show preview
    If FileExists($sTempBMP) Then
        GUICtrlSetImage($previewPic, $sTempBMP)
        GUICtrlSetState($previewPic, $GUI_SHOW)
        $hPreviewBitmap = $hBitmap
    Else
        _ClearPreview()
    EndIf
EndFunc

Func _ClearPreview()
    GUICtrlSetState($previewPic, $GUI_HIDE)
    GUICtrlSetImage($previewPic, "")
    
    ; Clean up temporary preview files
    Local $aFiles = _FileListToArray(@TempDir, "preview_*.bmp", 1)
    If Not @error Then
        For $i = 1 To $aFiles[0]
            FileDelete(@TempDir & "\" & $aFiles[$i])
        Next
    EndIf
EndFunc

Func _ShowStatus($sText, $bOK)
    GUICtrlSetData($statusLabel, $sText)
    If $bOK Then
        GUICtrlSetColor($statusLabel, 0x008000)
    Else
        GUICtrlSetColor($statusLabel, 0xFF0000)
    EndIf
EndFunc

Func _LogAction($sAction, $sFilename)
    ; Log actions to Log_file.log with timestamp
    Local $sLogFile = $sScriptDir & "\Log_file.log"
    Local $sTimestamp = @YEAR & "-" & @MON & "-" & @MDAY & " " & @HOUR & ":" & @MIN & ":" & @SEC
    Local $sLogEntry = $sTimestamp & " - " & $sAction & ": " & $sFilename & @CRLF
    
    ; Open log file in append mode
    Local $hLogFile = FileOpen($sLogFile, 1) ; 1 = append mode
    If $hLogFile <> -1 Then
        FileWrite($hLogFile, $sLogEntry)
        FileClose($hLogFile)
        ConsoleWrite("Logged: " & $sLogEntry)
    Else
        ConsoleWrite("ERROR: Could not open log file: " & $sLogFile & @CRLF)
    EndIf
EndFunc

Func _LoadConfig()
    Local $sConfig = $sScriptDir & "\config.file"
    If Not FileExists($sConfig) Then Return
    Local $aLines = FileReadToArray($sConfig)
    For $i = 0 To UBound($aLines) - 1
        Local $sLine = StringStripWS($aLines[$i], 3)
        If StringInStr($sLine, "Folder-1") Then
            Local $sPath = StringRegExpReplace($sLine, '(?i)^.*=\s*"?([^"]+)"?\s*$', "$1")
            ; Expand environment variables in the path
            $sPath = _ExpandEnvironmentVariables($sPath)
            GUICtrlSetData($destInput, $sPath)
        ElseIf StringInStr($sLine, "Folder-2") Then
            ; Store Folder-2 path for face detection operations
            $sFolder2Path = StringRegExpReplace($sLine, '(?i)^.*=\s*"?([^"]+)"?\s*$', "$1")
            ; Expand environment variables in the path
            $sFolder2Path = _ExpandEnvironmentVariables($sFolder2Path)
        ElseIf StringInStr($sLine, "Folder-3") Then
            ; Store Folder-3 path for SFTP operations
            $sFolder3Path = StringRegExpReplace($sLine, '(?i)^.*=\s*"?([^"]+)"?\s*$', "$1")
            ; Expand environment variables in the path
            $sFolder3Path = _ExpandEnvironmentVariables($sFolder3Path)
        EndIf
    Next
EndFunc

Func _UpdatePreviewWithFaces($sImagePath, $aFaces)
    ; Clear previous preview
    _ClearPreview()
    
    ; Load image using GDI+
    Local $hImage = _GDIPlus_ImageLoadFromFile($sImagePath)
    If @error Then
        _ClearPreview()
        Return
    EndIf
    
    ; Get image dimensions
    Local $iWidth = _GDIPlus_ImageGetWidth($hImage)
    Local $iHeight = _GDIPlus_ImageGetHeight($hImage)
    
    ; Calculate scaled dimensions to fit 300x300 preview
    Local $iNewWidth, $iNewHeight
    If $iWidth > $iHeight Then
        $iNewWidth = 300
        $iNewHeight = Int($iHeight * 300 / $iWidth)
    Else
        $iNewHeight = 300
        $iNewWidth = Int($iWidth * 300 / $iHeight)
    EndIf
    
    ; Calculate scale factors
    Local $fScaleX = $iNewWidth / $iWidth
    Local $fScaleY = $iNewHeight / $iHeight
    
    ; Create bitmap for preview
    Local $hBitmap = _GDIPlus_BitmapCreateFromScan0($iNewWidth, $iNewHeight)
    Local $hGraphics = _GDIPlus_ImageGetGraphicsContext($hBitmap)
    _GDIPlus_GraphicsSetInterpolationMode($hGraphics, $GDIP_INTERPOLATIONMODE_HIGHQUALITYBICUBIC)
    _GDIPlus_GraphicsDrawImageRect($hGraphics, $hImage, 0, 0, $iNewWidth, $iNewHeight)
    
    ; Clear face regions array
    Global $aFaceRegions[0]
    
    ; Draw rectangles around detected faces
    For $i = 0 To UBound($aFaces) - 1
        Local $aFace = $aFaces[$i]
        Local $iFaceX = $aFace[0] * $fScaleX
        Local $iFaceY = $aFace[1] * $fScaleY
        Local $iFaceWidth = $aFace[2] * $fScaleX
        Local $iFaceHeight = $aFace[3] * $fScaleY
        
        ; Store face region for click detection (expanded slightly for easier clicking)
        Local $iExpandedX = _Max(0, $iFaceX - 5)
        Local $iExpandedY = _Max(0, $iFaceY - 5)
        Local $iExpandedWidth = _Min($iNewWidth - $iExpandedX, $iFaceWidth + 10)
        Local $iExpandedHeight = _Min($iNewHeight - $iExpandedY, $iFaceHeight + 10)
        
        Local $sRegion = $iExpandedX & "|" & $iExpandedY & "|" & ($iExpandedX + $iExpandedWidth) & "|" & ($iExpandedY + $iExpandedHeight)
        
        ; Manually add to array
        Local $iSize = UBound($aFaceRegions)
        ReDim $aFaceRegions[$iSize + 1]
        $aFaceRegions[$iSize] = $sRegion
        
        ; Choose pen color based on selection
        Local $iPenColor, $iBrushColor
        If $i = $iSelectedFace Then
            $iPenColor = 0xFF00FF00  ; Green for selected face
            $iBrushColor = 0xFF00FF00
        Else
            $iPenColor = 0xFFFF0000  ; Red for unselected faces
            $iBrushColor = 0xFFFF0000
        EndIf
        
        ; Create pen for face rectangles
        Local $hPen = _GDIPlus_PenCreate($iPenColor, 3)
        
        ; Draw rectangle
        _GDIPlus_GraphicsDrawRect($hGraphics, $iFaceX, $iFaceY, $iFaceWidth, $iFaceHeight, $hPen)
        
        ; Add face number label
        Local $hBrush = _GDIPlus_BrushCreateSolid($iBrushColor)
        Local $hFont = _GDIPlus_FontCreate(_GDIPlus_FontFamilyCreate("Arial"), 14)
        Local $hFormat = _GDIPlus_StringFormatCreate()
        _GDIPlus_GraphicsDrawString($hGraphics, "Face " & ($i + 1), $iFaceX + 5, $iFaceY + 5, $hFont, $hFormat, $hBrush)
        
        _GDIPlus_BrushDispose($hBrush)
        _GDIPlus_FontDispose($hFont)
        _GDIPlus_StringFormatDispose($hFormat)
        _GDIPlus_PenDispose($hPen)
    Next
    
    ; Save preview to temporary file
    Local $sTempBMP = @TempDir & "\face_preview_" & @MSEC & ".bmp"
    _GDIPlus_ImageSaveToFile($hBitmap, $sTempBMP)
    
    ; Cleanup GDI+ objects
    _GDIPlus_GraphicsDispose($hGraphics)
    _GDIPlus_ImageDispose($hBitmap)
    _GDIPlus_ImageDispose($hImage)
    
    ; Show preview
    If FileExists($sTempBMP) Then
        GUICtrlSetImage($previewPic, $sTempBMP)
        GUICtrlSetState($previewPic, $GUI_SHOW)
        $hPreviewBitmap = $hBitmap
    Else
        _ClearPreview()
    EndIf
EndFunc

Func _HandlePreviewClick()
    If UBound($aDetectedFaces) = 0 Then Return
    
    ; Handle click on preview image
    Local $aMousePos = GUIGetCursorInfo($hGui)
    If IsArray($aMousePos) Then
        Local $iMouseX = $aMousePos[0] - 480  ; Adjust for image position
        Local $iMouseY = $aMousePos[1] - 40
        
        ; Check if click is within any face region
        For $i = 0 To UBound($aFaceRegions) - 1
            Local $sRegion = $aFaceRegions[$i]
            Local $aCoords = StringSplit($sRegion, "|", 2)
            If IsArray($aCoords) And UBound($aCoords) = 4 Then
                If $iMouseX >= $aCoords[0] And $iMouseX <= $aCoords[2] And _
                   $iMouseY >= $aCoords[1] And $iMouseY <= $aCoords[3] Then
                    $iSelectedFace = $i
                    _ShowStatus("Selected: Face " & ($i + 1) & " - Click 'Crop Face' to crop", 1)
                    GUICtrlSetState($cropBtn, $GUI_ENABLE)
                    ; Update preview to show selection
                    _UpdatePreviewWithFaces($sCurrentImagePath, $aDetectedFaces)
                    ExitLoop
                EndIf
            EndIf
        Next
    EndIf
EndFunc

; ==============================================================
; Manual Face Selection Functions
; ==============================================================
Func _ToggleManualSelection()
    ; Enhanced validation for source image
    Local $sSrc = StringStripWS(GUICtrlRead($srcInput), 3)
    If $sSrc = "" Or Not FileExists($sSrc) Then
        _ShowStatus("Please select a source image first", 0)
        ConsoleWrite("ERROR: No source image selected for manual selection" & @CRLF)
        Return
    EndIf
    
    ; Enhanced path validation
    Local $sAbsolutePath = _GetAbsolutePath($sSrc)
    If Not FileExists($sAbsolutePath) Then
        _ShowStatus("Source image not found: " & $sAbsolutePath, 0)
        ConsoleWrite("ERROR: Source image not found: " & $sAbsolutePath & @CRLF)
        Return
    EndIf
    
    ; Always update current image path to the current source image
    ; This ensures we're always working with the currently selected image
    $sCurrentImagePath = $sAbsolutePath
    ConsoleWrite("ToggleManualSelection - Updated CurrentImagePath: " & $sCurrentImagePath & @CRLF)
    
    ; Enhanced output path generation for manual selection
    If $sCurrentOutputPath = "" Then
        Local $sDest = StringStripWS(GUICtrlRead($destInput), 3)
        Local $sName = StringStripWS(GUICtrlRead($nameInput), 3)
        
        If $sDest = "" Then
            _ShowStatus("Please select destination folder first", 0)
            ConsoleWrite("ERROR: No destination folder for manual selection" & @CRLF)
            Return
        EndIf
        
        If $sName = "" Or StringLen($sName) <> 5 Or Not StringIsDigit($sName) Then
            _ShowStatus("Please enter valid 5-digit client number first", 0)
            ConsoleWrite("ERROR: Invalid client number for manual selection" & @CRLF)
            Return
        EndIf
        
        ; Generate output path with client number only
        $sCurrentOutputPath = $sDest & "\" & $sName & ".jpg"
    EndIf
    
    ConsoleWrite("ToggleManualSelection - CurrentOutputPath: " & $sCurrentOutputPath & @CRLF)
    
    $bManualSelectionMode = Not $bManualSelectionMode
    If $bManualSelectionMode Then
        GUICtrlSetData($manualSelectBtn, "Cancel Manual")
        _ShowStatus("Manual Face Selection Mode: Click and drag on preview to select face area", 1)
        ; Reset manual selection with enhanced initialization
        $aManualSelection[0] = -1
        $aManualSelection[1] = -1
        $aManualSelection[2] = -1
        $aManualSelection[3] = -1
        ; Enhanced preview update with error handling - always use current source image
        If FileExists($sCurrentImagePath) Then
            _UpdatePreviewWithManualSelection($sCurrentImagePath, $aDetectedFaces)
        Else
            _ShowStatus("Source image not available for preview", 0)
            ConsoleWrite("ERROR: Source image not found for preview: " & $sCurrentImagePath & @CRLF)
            _UpdatePreview()
        EndIf
    Else
        GUICtrlSetData($manualSelectBtn, "Manual Face Select")
        _ShowStatus("Manual face selection cancelled", 1)
        ; Reset manual selection
        $aManualSelection[0] = -1
        $aManualSelection[1] = -1
        $aManualSelection[2] = -1
        $aManualSelection[3] = -1
        ; Enhanced preview restoration with error handling - always use current source image
        If FileExists($sCurrentImagePath) Then
            If IsArray($aDetectedFaces) And UBound($aDetectedFaces) > 0 Then
                _UpdatePreviewWithFaces($sCurrentImagePath, $aDetectedFaces)
            Else
                _UpdatePreview()
            EndIf
        Else
            _ShowStatus("Source image not available for preview", 0)
            ConsoleWrite("ERROR: Source image not found for preview restoration: " & $sCurrentImagePath & @CRLF)
            _UpdatePreview()
        EndIf
    EndIf
EndFunc

Func _HandleManualSelection()
    Local $aMousePos = GUIGetCursorInfo($hGui)
    If Not IsArray($aMousePos) Then Return
    
    Local $iMouseX = $aMousePos[0] - 480  ; Adjust for image position
    Local $iMouseY = $aMousePos[1] - 40
    
    ; Check if click is within preview area
    If $iMouseX < 0 Or $iMouseX > 300 Or $iMouseY < 0 Or $iMouseY > 300 Then Return
    
    ; Start selection on first click
    If $aManualSelection[0] = -1 Then
        $aManualSelection[0] = $iMouseX
        $aManualSelection[1] = $iMouseY
        $aManualSelection[2] = $iMouseX
        $aManualSelection[3] = $iMouseY
        _ShowStatus("Manual Selection: Drag to define selection area, release to complete", 1)
        
        ; Create a temporary preview with the initial selection
        _UpdatePreviewWithManualSelection($sCurrentImagePath, $aDetectedFaces)
        
        ; Start tracking mouse movement for drag operation
        Local $iLastUpdateTime = TimerInit()
        While $bManualSelectionMode
            Local $aCurrentPos = GUIGetCursorInfo($hGui)
            If IsArray($aCurrentPos) Then
                Local $iCurrentX = $aCurrentPos[0] - 480
                Local $iCurrentY = $aCurrentPos[1] - 40
                
                ; Update selection rectangle while dragging
                If $iCurrentX >= 0 And $iCurrentX <= 300 And $iCurrentY >= 0 And $iCurrentY <= 300 Then
                    ; Only update if coordinates changed significantly or enough time has passed
                    If Abs($iCurrentX - $aManualSelection[2]) > 2 Or Abs($iCurrentY - $aManualSelection[3]) > 2 Or TimerDiff($iLastUpdateTime) > 50 Then
                        $aManualSelection[2] = $iCurrentX
                        $aManualSelection[3] = $iCurrentY
                        
                        ; Update preview only when necessary to reduce flickering
                        _UpdatePreviewWithManualSelection($sCurrentImagePath, $aDetectedFaces)
                        $iLastUpdateTime = TimerInit()
                    EndIf
                EndIf
                
                ; Check if mouse button is released
                If Not $aCurrentPos[2] Then ; Mouse button released
                    ExitLoop
                EndIf
            EndIf
            Sleep(5) ; Small delay to prevent high CPU usage
        WEnd
        
        ; Ensure proper rectangle coordinates (top-left to bottom-right)
        Local $iX1 = _Min($aManualSelection[0], $aManualSelection[2])
        Local $iY1 = _Min($aManualSelection[1], $aManualSelection[3])
        Local $iX2 = _Max($aManualSelection[0], $aManualSelection[2])
        Local $iY2 = _Max($aManualSelection[1], $aManualSelection[3])
        
        $aManualSelection[0] = $iX1
        $aManualSelection[1] = $iY1
        $aManualSelection[2] = $iX2
        $aManualSelection[3] = $iY2
        
        ; Update preview with final selection
        _UpdatePreviewWithManualSelection($sCurrentImagePath, $aDetectedFaces)
        
        ; Enable crop button for manual selection
        GUICtrlSetState($cropBtn, $GUI_ENABLE)
        _ShowStatus("Manual selection complete - Click 'Crop Face' to crop selected area", 1)
        
        ; Exit manual selection mode
        $bManualSelectionMode = False
        GUICtrlSetData($manualSelectBtn, "Manual Face Select")
    EndIf
EndFunc

Func _UpdatePreviewWithManualSelection($sImagePath, $aFaces)
    ; Clear previous preview
    _ClearPreview()
    
    ; Check if we have a valid image path
    If $sImagePath = "" Or Not FileExists($sImagePath) Then
        _ShowStatus("No valid image selected for preview", 0)
        Return
    EndIf
    
    ; Load image using GDI+
    Local $hImage = _GDIPlus_ImageLoadFromFile($sImagePath)
    If @error Then
        _ClearPreview()
        _ShowStatus("Error loading image for preview: " & @error, 0)
        Return
    EndIf
    
    ; Get image dimensions
    Local $iWidth = _GDIPlus_ImageGetWidth($hImage)
    Local $iHeight = _GDIPlus_ImageGetHeight($hImage)
    
    ; Calculate scaled dimensions to fit 300x300 preview
    Local $iNewWidth, $iNewHeight
    If $iWidth > $iHeight Then
        $iNewWidth = 300
        $iNewHeight = Int($iHeight * 300 / $iWidth)
    Else
        $iNewHeight = 300
        $iNewWidth = Int($iWidth * 300 / $iHeight)
    EndIf
    
    ; Calculate scale factors
    Local $fScaleX = $iNewWidth / $iWidth
    Local $fScaleY = $iNewHeight / $iHeight
    
    ; Create bitmap for preview
    Local $hBitmap = _GDIPlus_BitmapCreateFromScan0($iNewWidth, $iNewHeight)
    Local $hGraphics = _GDIPlus_ImageGetGraphicsContext($hBitmap)
    _GDIPlus_GraphicsSetInterpolationMode($hGraphics, $GDIP_INTERPOLATIONMODE_HIGHQUALITYBICUBIC)
    _GDIPlus_GraphicsDrawImageRect($hGraphics, $hImage, 0, 0, $iNewWidth, $iNewHeight)
    
    ; Clear face regions array
    Global $aFaceRegions[0]
    
    ; Draw rectangles around detected faces (if any)
    For $i = 0 To UBound($aFaces) - 1
        Local $aFace = $aFaces[$i]
        Local $iFaceX = $aFace[0] * $fScaleX
        Local $iFaceY = $aFace[1] * $fScaleY
        Local $iFaceWidth = $aFace[2] * $fScaleX
        Local $iFaceHeight = $aFace[3] * $fScaleY
        
        ; Store face region for click detection
        Local $iExpandedX = _Max(0, $iFaceX - 5)
        Local $iExpandedY = _Max(0, $iFaceY - 5)
        Local $iExpandedWidth = _Min($iNewWidth - $iExpandedX, $iFaceWidth + 10)
        Local $iExpandedHeight = _Min($iNewHeight - $iExpandedY, $iFaceHeight + 10)
        
        Local $sRegion = $iExpandedX & "|" & $iExpandedY & "|" & ($iExpandedX + $iExpandedWidth) & "|" & ($iExpandedY + $iExpandedHeight)
        
        ; Manually add to array
        Local $iSize = UBound($aFaceRegions)
        ReDim $aFaceRegions[$iSize + 1]
        $aFaceRegions[$iSize] = $sRegion
        
        ; Draw face rectangles in red
        Local $hPen = _GDIPlus_PenCreate(0xFFFF0000, 2)
        _GDIPlus_GraphicsDrawRect($hGraphics, $iFaceX, $iFaceY, $iFaceWidth, $iFaceHeight, $hPen)
        
        ; Add face number label
        Local $hBrush = _GDIPlus_BrushCreateSolid(0xFFFF0000)
        Local $hFont = _GDIPlus_FontCreate(_GDIPlus_FontFamilyCreate("Arial"), 12)
        Local $hFormat = _GDIPlus_StringFormatCreate()
        _GDIPlus_GraphicsDrawString($hGraphics, "Face " & ($i + 1), $iFaceX + 5, $iFaceY + 5, $hFont, $hFormat, $hBrush)
        
        _GDIPlus_BrushDispose($hBrush)
        _GDIPlus_FontDispose($hFont)
        _GDIPlus_StringFormatDispose($hFormat)
        _GDIPlus_PenDispose($hPen)
    Next
    
    ; Draw manual selection rectangle if active
    If $aManualSelection[0] <> -1 And $aManualSelection[1] <> -1 And _
       $aManualSelection[2] <> -1 And $aManualSelection[3] <> -1 Then
        
        Local $iSelX1 = $aManualSelection[0]
        Local $iSelY1 = $aManualSelection[1]
        Local $iSelX2 = $aManualSelection[2]
        Local $iSelY2 = $aManualSelection[3]
        Local $iSelWidth = $iSelX2 - $iSelX1
        Local $iSelHeight = $iSelY2 - $iSelY1
        
        ; Draw manual selection rectangle in green
        Local $hPen = _GDIPlus_PenCreate(0xFF00FF00, 3)
        _GDIPlus_GraphicsDrawRect($hGraphics, $iSelX1, $iSelY1, $iSelWidth, $iSelHeight, $hPen)
        
        ; Add manual selection label
        Local $hBrush = _GDIPlus_BrushCreateSolid(0xFF00FF00)
        Local $hFont = _GDIPlus_FontCreate(_GDIPlus_FontFamilyCreate("Arial"), 14)
        Local $hFormat = _GDIPlus_StringFormatCreate()
        _GDIPlus_GraphicsDrawString($hGraphics, "Manual Selection", $iSelX1 + 5, $iSelY1 + 5, $hFont, $hFormat, $hBrush)
        
        _GDIPlus_BrushDispose($hBrush)
        _GDIPlus_FontDispose($hFont)
        _GDIPlus_StringFormatDispose($hFormat)
        _GDIPlus_PenDispose($hPen)
    EndIf
    
    ; Save preview to temporary file
    Local $sTempBMP = @TempDir & "\face_preview_" & @MSEC & ".bmp"
    _GDIPlus_ImageSaveToFile($hBitmap, $sTempBMP)
    
    ; Cleanup GDI+ objects
    _GDIPlus_GraphicsDispose($hGraphics)
    _GDIPlus_ImageDispose($hBitmap)
    _GDIPlus_ImageDispose($hImage)
    
    ; Show preview
    If FileExists($sTempBMP) Then
        GUICtrlSetImage($previewPic, $sTempBMP)
        GUICtrlSetState($previewPic, $GUI_SHOW)
        $hPreviewBitmap = $hBitmap
    Else
        _ClearPreview()
    EndIf
EndFunc

Func _CropSelectedFace()
    ConsoleWrite("=== Starting CropSelectedFace ===" & @CRLF)
    
    ; Check if manual selection is active
    Local $bManualMode = ($aManualSelection[0] <> -1 And $aManualSelection[1] <> -1 And _
                          $aManualSelection[2] <> -1 And $aManualSelection[3] <> -1)
    
    ConsoleWrite("Manual mode: " & $bManualMode & @CRLF)
    ConsoleWrite("Selected face: " & $iSelectedFace & @CRLF)
    ConsoleWrite("Manual selection: X1=" & $aManualSelection[0] & ", Y1=" & $aManualSelection[1] & ", X2=" & $aManualSelection[2] & ", Y2=" & $aManualSelection[3] & @CRLF)
    
    If $iSelectedFace = -1 And Not $bManualMode Then
        ConsoleWrite("No selection made" & @CRLF)
        MsgBox(48, "No Selection", "Please either:" & @CRLF & _
               "1. Click on a detected face in the preview, OR" & @CRLF & _
               "2. Use 'Manual Select' to draw a selection area")
        Return
    EndIf
    
    ; Validate that we have the required paths
    If $sCurrentImagePath = "" Or Not FileExists($sCurrentImagePath) Then
        ConsoleWrite("ERROR: No valid source image selected" & @CRLF)
        MsgBox(16, "Error", "Please select a valid source image first")
        Return
    EndIf

    ; Build output filename
    Local $sOutputPath
    If $bManualMode Then
        ; For manual mode, always generate a unique filename with timestamp and milliseconds
        Local $sDest = StringStripWS(GUICtrlRead($destInput), 3)
        Local $sName = StringStripWS(GUICtrlRead($nameInput), 3)
        If $sDest = "" Or $sName = "" Then
            ConsoleWrite("ERROR: Missing destination folder or client number" & @CRLF)
            MsgBox(16, "Error", "Please select destination folder and enter client number")
            Return
        EndIf
        ; Generate filename with client number only
        $sOutputPath = $sDest & "\" & $sName & ".jpg"
    Else
        ; For face detection mode, use client number only
        $sOutputPath = $sCurrentOutputPath
    EndIf

    ; Check if output file already exists and ask for overwrite confirmation
    If FileExists($sOutputPath) Then
        ; Ask user for overwrite confirmation (modal to main window)
        Local $iResponse = MsgBox(4 + 32 + 262144, "File Exists", "File " & StringRegExpReplace($sOutputPath, "^.*\\", "") & " already exists." & @CRLF & "Do you want to overwrite it?")
        
        If $iResponse <> 6 Then ; 6 = Yes
            _ShowStatus("Operation cancelled by user.", 0)
            Return
        EndIf
    EndIf
    
    ConsoleWrite("Output path: " & $sOutputPath & @CRLF)
    ConsoleWrite("Current image path: " & $sCurrentImagePath & @CRLF)
    
    Local $bSuccess = False
    Local $sMsg = ""
    
    If $bManualMode Then
        ConsoleWrite("Attempting manual crop..." & @CRLF)
        ; Crop manual selection
        $bSuccess = _CropManualSelection($sCurrentImagePath, $sOutputPath)
        If $bSuccess Then
            $sMsg = "Manual Crop Successful!" & @CRLF & _
                   "Cropped selection to 1:1 ratio" & @CRLF & _
                   "Saved: " & $sOutputPath
        Else
            ConsoleWrite("Manual crop failed" & @CRLF)
        EndIf
    Else
        ConsoleWrite("Attempting face detection crop..." & @CRLF)
        ; Crop detected face
        ConsoleWrite("Calling _CropDetectedFace with face index: " & $iSelectedFace & @CRLF)
        If IsArray($aDetectedFaces) And $iSelectedFace >= 0 And $iSelectedFace < UBound($aDetectedFaces) Then
            Local $aFace = $aDetectedFaces[$iSelectedFace]
            If IsArray($aFace) And UBound($aFace) = 4 Then
                ConsoleWrite("Face data: X=" & $aFace[0] & ", Y=" & $aFace[1] & ", W=" & $aFace[2] & ", H=" & $aFace[3] & @CRLF)
                $bSuccess = _CropDetectedFace($sCurrentImagePath, $sOutputPath, $aFace)
            Else
                ConsoleWrite("ERROR: Invalid face data structure" & @CRLF)
                $bSuccess = False
            EndIf
        Else
            ConsoleWrite("ERROR: Invalid face selection or face data" & @CRLF)
            $bSuccess = False
        EndIf
        If $bSuccess Then
            $sMsg = "AI Face Detection Successful!" & @CRLF & _
                   "Faces detected: " & UBound($aDetectedFaces) & @CRLF & _
                   "Cropped selected face to 1:1 ratio" & @CRLF & _
                   "Saved: " & StringRegExpReplace($sOutputPath, "^.*\\", "")

            If UBound($aDetectedFaces) > 1 Then
                $sMsg &= @CRLF & @CRLF & "Note: " & UBound($aDetectedFaces) - 1 & " additional face(s) were detected but not cropped."
            EndIf
        Else
            ConsoleWrite("Face detection crop failed" & @CRLF)
        EndIf
    EndIf
    
    If $bSuccess Then
        ConsoleWrite("Crop operation successful" & @CRLF)
        MsgBox(64, "Success", $sMsg)
        _ShowStatus("Crop successful: " & StringRegExpReplace($sOutputPath, "^.*\\", ""), 1)
        _LogAction("Cropped", $sOutputPath)
        
        ; Reset manual selection after successful crop
        $aManualSelection[0] = -1
        $aManualSelection[1] = -1
        $aManualSelection[2] = -1
        $aManualSelection[3] = -1
        
        ; Open the CROPPED folder to show the result
        Local $sCroppedFolder = StringRegExpReplace($sOutputPath, "(.*)\\.*$", "$1")
        If FileExists($sCroppedFolder) Then
            ShellExecute($sCroppedFolder)
            ConsoleWrite("Opened CROPPED folder: " & $sCroppedFolder & @CRLF)
        Else
            ConsoleWrite("WARNING: CROPPED folder not found: " & $sCroppedFolder & @CRLF)
        EndIf
    Else
        ConsoleWrite("Crop operation failed" & @CRLF)
        MsgBox(16, "Error", "Failed to crop selected area")
        _ShowStatus("Failed to crop", 0)
    EndIf
    ConsoleWrite("=== Finished CropSelectedFace ===" & @CRLF)
EndFunc

Func _CropManualSelection($sImagePath, $sOutputPath)
    ConsoleWrite("=== Starting manual crop ===" & @CRLF)
    ConsoleWrite("Image path: " & $sImagePath & @CRLF)
    ConsoleWrite("Output path: " & $sOutputPath & @CRLF)
    ConsoleWrite("Manual selection coordinates: X1=" & $aManualSelection[0] & ", Y1=" & $aManualSelection[1] & ", X2=" & $aManualSelection[2] & ", Y2=" & $aManualSelection[3] & @CRLF)
    
    _GDIPlus_Startup()
    Local $hImage = _GDIPlus_ImageLoadFromFile($sImagePath)
    If @error Then
        ConsoleWrite("ERROR: Failed to load image: " & @error & @CRLF)
        _GDIPlus_Shutdown()
        Return False
    EndIf

    ; Get original image dimensions
    Local $iImgWidth = _GDIPlus_ImageGetWidth($hImage)
    Local $iImgHeight = _GDIPlus_ImageGetHeight($hImage)
    ConsoleWrite("Original image dimensions: " & $iImgWidth & "x" & $iImgHeight & @CRLF)
    
    ; Calculate scale factors from preview to original image
    ; The preview is always scaled to fit within 300x300 while maintaining aspect ratio
    Local $iPreviewWidth, $iPreviewHeight
    If $iImgWidth > $iImgHeight Then
        $iPreviewWidth = 300
        $iPreviewHeight = Int($iImgHeight * 300 / $iImgWidth)
    Else
        $iPreviewHeight = 300
        $iPreviewWidth = Int($iImgWidth * 300 / $iImgHeight)
    EndIf
    ConsoleWrite("Preview dimensions: " & $iPreviewWidth & "x" & $iPreviewHeight & @CRLF)
    
    ; Calculate scale factors from preview to original image
    Local $fScaleX = $iImgWidth / $iPreviewWidth
    Local $fScaleY = $iImgHeight / $iPreviewHeight
    ConsoleWrite("Scale factors: X=" & $fScaleX & ", Y=" & $fScaleY & @CRLF)
    
    ; The preview image is drawn at (0,0) in the preview bitmap, not centered
    ; So we don't need to adjust for centering offsets
    ConsoleWrite("Preview dimensions: " & $iPreviewWidth & "x" & $iPreviewHeight & " (no centering offsets)" & @CRLF)

    ; Manual selection coordinates are already relative to the actual image area
    ; since the preview image is drawn at (0,0) in the 300x300 control
    Local $iAdjustedSelX1 = $aManualSelection[0]
    Local $iAdjustedSelY1 = $aManualSelection[1]
    Local $iAdjustedSelX2 = $aManualSelection[2]
    Local $iAdjustedSelY2 = $aManualSelection[3]
    
    ; Ensure adjusted coordinates are within image bounds
    $iAdjustedSelX1 = _Max(0, $iAdjustedSelX1)
    $iAdjustedSelY1 = _Max(0, $iAdjustedSelY1)
    $iAdjustedSelX2 = _Min($iPreviewWidth, $iAdjustedSelX2)
    $iAdjustedSelY2 = _Min($iPreviewHeight, $iAdjustedSelY2)
    
    ConsoleWrite("Adjusted manual selection (preview): X1=" & $iAdjustedSelX1 & ", Y1=" & $iAdjustedSelY1 & ", X2=" & $iAdjustedSelX2 & ", Y2=" & $iAdjustedSelY2 & @CRLF)
    
    ; Convert manual selection coordinates from preview to original image
    Local $iSelX1 = Int($iAdjustedSelX1 * $fScaleX)
    Local $iSelY1 = Int($iAdjustedSelY1 * $fScaleY)
    Local $iSelX2 = Int($iAdjustedSelX2 * $fScaleX)
    Local $iSelY2 = Int($iAdjustedSelY2 * $fScaleY)
    ConsoleWrite("Manual selection (original): X1=" & $iSelX1 & ", Y1=" & $iSelY1 & ", X2=" & $iSelX2 & ", Y2=" & $iSelY2 & @CRLF)
    
    ; Calculate selection dimensions
    Local $iSelWidth = $iSelX2 - $iSelX1
    Local $iSelHeight = $iSelY2 - $iSelY1
    ConsoleWrite("Selection dimensions: " & $iSelWidth & "x" & $iSelHeight & @CRLF)
    
    ; For manual selection, make it 1:1 aspect ratio (square) like AI face detection
    ; Take the larger dimension and center the crop area
    Local $iSquareSize = _Max($iSelWidth, $iSelHeight)
    
    ; Center the crop area around the selection
    Local $iCenterX = $iSelX1 + Int($iSelWidth / 2)
    Local $iCenterY = $iSelY1 + Int($iSelHeight / 2)
    
    Local $iCropX = $iCenterX - Int($iSquareSize / 2)
    Local $iCropY = $iCenterY - Int($iSquareSize / 2)
    
    ConsoleWrite("Square crop dimensions: " & $iSquareSize & "x" & $iSquareSize & @CRLF)
    ConsoleWrite("Crop coordinates: X=" & $iCropX & ", Y=" & $iCropY & @CRLF)
    
    ; Ensure crop area stays within image bounds
    $iCropX = _Max(0, $iCropX)
    $iCropY = _Max(0, $iCropY)
    $iSquareSize = _Min3($iSquareSize, $iImgWidth - $iCropX, $iImgHeight - $iCropY)
    ConsoleWrite("Adjusted crop coordinates: X=" & $iCropX & ", Y=" & $iCropY & ", Size=" & $iSquareSize & @CRLF)
    
    If $iSquareSize < 10 Then
        ConsoleWrite("ERROR: Crop size too small: " & $iSquareSize & @CRLF)
        _GDIPlus_Shutdown()
        Return False
    EndIf
    
    ; Check if crop area is valid
    If $iCropX < 0 Or $iCropY < 0 Or $iCropX + $iSquareSize > $iImgWidth Or $iCropY + $iSquareSize > $iImgHeight Then
        ConsoleWrite("ERROR: Crop area outside image bounds" & @CRLF)
        ConsoleWrite("Crop area: X=" & $iCropX & ", Y=" & $iCropY & ", Size=" & $iSquareSize & @CRLF)
        ConsoleWrite("Image bounds: Width=" & $iImgWidth & ", Height=" & $iImgHeight & @CRLF)
        _GDIPlus_Shutdown()
        Return False
    EndIf
    
    ; Perform the crop
    ConsoleWrite("Attempting to clone bitmap area..." & @CRLF)
    Local $hCropped = _GDIPlus_BitmapCloneArea($hImage, $iCropX, $iCropY, $iSquareSize, $iSquareSize)
    If @error Then
        ConsoleWrite("ERROR: Failed to clone bitmap area: " & @error & @CRLF)
        _GDIPlus_Shutdown()
        Return False
    EndIf
    
    ConsoleWrite("Attempting to save cropped image..." & @CRLF)
    Local $bResult = _GDIPlus_ImageSaveToFile($hCropped, $sOutputPath)
    If @error Then
        ConsoleWrite("ERROR: Failed to save cropped image: " & @error & @CRLF)
    EndIf
    
    _GDIPlus_ImageDispose($hImage)
    _GDIPlus_ImageDispose($hCropped)
    _GDIPlus_Shutdown()
    
    If $bResult Then
        ConsoleWrite("SUCCESS: Manual crop completed: " & $iSquareSize & "x" & $iSquareSize & " (1:1)" & @CRLF)
        ConsoleWrite("Crop area: X=" & $iCropX & ", Y=" & $iCropY & @CRLF)
    Else
        ConsoleWrite("ERROR: Manual crop failed - check file permissions or disk space" & @CRLF)
    EndIf
    
    ConsoleWrite("=== Finished manual crop ===" & @CRLF)
    Return $bResult
EndFunc

; ==============================================================
; UI Guidance Functions
; ==============================================================
Func _UpdateStepGuidance()
    Local $sSrc = StringStripWS(GUICtrlRead($srcInput), 3)
    Local $sDest = StringStripWS(GUICtrlRead($destInput), 3)
    Local $sName = StringStripWS(GUICtrlRead($nameInput), 3)
    Local $sMode = GUICtrlRead($modeCombo)
    
    ; Reset all controls to default appearance
    GUICtrlSetBkColor($srcBrowseBtn, 0xF0F0F0)
    GUICtrlSetBkColor($destBrowseBtn, 0xF0F0F0)
    GUICtrlSetBkColor($nameInput, 0xFFFFFF)
    GUICtrlSetBkColor($processBtn, 0xF0F0F0)
    
    ; Determine current step based on what's filled and mode
    If $sMode = "Upload_to_SFTP" Then
        ; SFTP mode requires client number and source files
        If $sSrc = "" Then
            ; Step 1: Select source files
            $iCurrentStep = 1
            GUICtrlSetData($hStepLabel, "Step 1: Click 'Browse' to select image files to upload")
            GUICtrlSetBkColor($srcBrowseBtn, 0x90EE90) ; Light green
        ElseIf $sName = "" Or StringLen($sName) <> 5 Or Not StringIsDigit($sName) Then
            ; Step 2: Enter valid client number
            $iCurrentStep = 2
            GUICtrlSetData($hStepLabel, "Step 2: Enter 5-digit client number")
            GUICtrlSetBkColor($nameInput, 0x90EE90) ; Light green
        Else
            ; Step 3: Ready to upload
            $iCurrentStep = 3
            GUICtrlSetData($hStepLabel, "Step 3: Click 'Upload to SFTP' to upload selected files")
            GUICtrlSetBkColor($processBtn, 0x90EE90) ; Light green
            GUICtrlSetBkColor($nameInput, 0xFFFFFF) ; White background
        EndIf
    Else
        ; Other modes (Resize_and_Copy, Detect_Face_and_Crop)
        If $sSrc = "" Then
            ; Step 1: Select source file
            $iCurrentStep = 1
            GUICtrlSetData($hStepLabel, "Step 1: Click 'Browse' to select source image")
            GUICtrlSetBkColor($srcBrowseBtn, 0x90EE90) ; Light green
        ElseIf $sDest = "" Then
            ; Step 2: Select destination folder
            $iCurrentStep = 2
            GUICtrlSetData($hStepLabel, "Step 2: Click 'Browse' to select destination folder")
            GUICtrlSetBkColor($destBrowseBtn, 0x90EE90) ; Light green
        ElseIf $sName = "" Or StringLen($sName) <> 5 Or Not StringIsDigit($sName) Then
            ; Step 3: Enter valid client number
            $iCurrentStep = 3
            GUICtrlSetData($hStepLabel, "Step 3: Enter 5-digit client number")
            GUICtrlSetBkColor($nameInput, 0x90EE90) ; Light green
        Else
            ; Step 4: Ready to process
            $iCurrentStep = 4
            GUICtrlSetData($hStepLabel, "Step 4: Click '" & GUICtrlRead($processBtn) & "' to process")
            GUICtrlSetBkColor($processBtn, 0x90EE90) ; Light green
            GUICtrlSetBkColor($nameInput, 0xFFFFFF) ; White background
        EndIf
    EndIf
EndFunc

Func _HighlightNextStep()
    ; This function is called after each action to update the guidance
    _UpdateStepGuidance()
EndFunc

; ==============================================================
; Cardpresso Functions
; ==============================================================
Func _RunCardpresso()
    Local $sName = StringStripWS(GUICtrlRead($nameInput), 3)
    
    GUICtrlSetData($statusLabel, "")
    
    If $sName = "" Then
        _ShowStatus("Please enter client number.", 0)
        Return
    EndIf
    If StringLen($sName) <> 5 Or Not StringIsDigit($sName) Then
        _ShowStatus("Client number must be exactly 5 digits.", 0)
        Return
    EndIf
    
    ; Show progress bar
    GUICtrlSetState($progressBar, $GUI_SHOW)
    GUICtrlSetData($progressBar, 0)
    _ShowStatus("Starting Cardpresso lookup for client: " & $sName, 1)
    
    ; Update progress
    GUICtrlSetData($progressBar, 30)
    _ShowStatus("Looking up client in Master.xlsx...", 1)
    
    ; Call the Cardpresso lookup function
    Local $bSuccess = _CardpressoLookupAndExport($sName)
    
    ; Complete progress
    GUICtrlSetData($progressBar, 100)
    Sleep(500)
    GUICtrlSetState($progressBar, $GUI_HIDE)
    
    If $bSuccess Then
        _ShowStatus("Cardpresso lookup successful - CSV exported to CardPresso.csv", 1)
        _LogAction("Cardpresso Export", "Client: " & $sName)
    Else
        _ShowStatus("Cardpresso lookup failed. Check Master.xlsx file and client number.", 0)
    EndIf
EndFunc

Func _CardpressoLookupAndExport($sClientID)
    ConsoleWrite("=== Starting Cardpresso lookup for client: " & $sClientID & " ===" & @CRLF)
    
    ; Load configuration if not already loaded
    If $sCardpressoMasterFile = "" Then
        _LoadCardpressoConfig()
    EndIf
    
    ; Use configuration paths
    Local $sMasterFile = $sCardpressoMasterFile
    Local $sTempCSV = $sCardpressoCSVFile
    Local $sPhotoBaseDir = $sCardpressoPhotoBaseDir
    
    ; Check if Master.xlsx exists
    If Not FileExists($sMasterFile) Then
        ConsoleWrite("ERROR: Master file not found: " & $sMasterFile & @CRLF)
        MsgBox(16, "File Missing", "Master.xlsx file not found: " & @CRLF & $sMasterFile)
        Return False
    EndIf
    
    ; Try to open Excel
    Local $oExcel = ObjCreate("Excel.Application")
    If @error Or Not IsObj($oExcel) Then
        ConsoleWrite("ERROR: Unable to create Excel object" & @CRLF)
        MsgBox(16, "Excel Error", "Unable to open Excel. Please ensure Microsoft Excel is installed.")
        Return False
    EndIf
    
    $oExcel.Visible = False
    $oExcel.DisplayAlerts = False
    
    Local $oWorkbook = $oExcel.Workbooks.Open($sMasterFile)
    If @error Or Not IsObj($oWorkbook) Then
        ConsoleWrite("ERROR: Unable to open workbook: " & $sMasterFile & @CRLF)
        $oExcel.Quit()
        Return False
    EndIf
    
    Local $oSheet = $oWorkbook.ActiveSheet
    Local $lastRow = $oSheet.UsedRange.Rows.Count
    Local $bFound = False
    Local $sFoundID = ""
    Local $sFoundName = ""
    Local $sFoundPhoto = ""
    
    ; Search for client ID in column A (ID column)
    For $i = 2 To $lastRow ; Assuming row 1 is headers
        Local $sCurrentID = String($oSheet.Cells($i, 1).Value) ; Column A = ID
        If $sCurrentID = $sClientID Then
            $sFoundID = $sCurrentID
            $sFoundName = String($oSheet.Cells($i, 2).Value) ; Column B = Name
            $sFoundPhoto = String($oSheet.Cells($i, 3).Value) ; Column C = Photo
            $bFound = True
            ConsoleWrite("Found client: ID=" & $sFoundID & ", Name=" & $sFoundName & ", Photo=" & $sFoundPhoto & @CRLF)
            ExitLoop
        EndIf
    Next
    
    ; Close Excel
    $oWorkbook.Close(False)
    $oExcel.Quit()
    
    If Not $bFound Then
        ConsoleWrite("ERROR: Client ID " & $sClientID & " not found in Master.xlsx" & @CRLF)
        MsgBox(16, "Not Found", "Client ID " & $sClientID & " not found in Master.xlsx.")
        _ClearPreview() ; Clear previously loaded image
        Return False
    EndIf
    
    ; Show photo preview
    Local $bPhotoFound = _ShowCardpressoPhoto($sFoundPhoto, $sClientID)
    If Not $bPhotoFound Then
        Return False ; Stop processing if photo not found
    EndIf
    
    ; Export to CSV
    Local $hFile = FileOpen($sTempCSV, 2) ; 2 = overwrite mode
    If $hFile = -1 Then
        ConsoleWrite("ERROR: Unable to open CSV file for writing: " & $sTempCSV & @CRLF)
        MsgBox(16, "Write Error", "Unable to create CSV file: " & @CRLF & $sTempCSV)
        Return False
    EndIf
    
    ; Write CSV header and data
    FileWriteLine($hFile, "ID_Number,Name,Photo")
    Local $sCSVLine = _CSV($sFoundID) & "," & _CSV($sFoundName) & "," & _CSV($sFoundPhoto)
    FileWriteLine($hFile, $sCSVLine)
    FileClose($hFile)
    
    ConsoleWrite("CSV exported successfully: " & $sTempCSV & @CRLF)
    ConsoleWrite("CSV content: " & $sCSVLine & @CRLF)
    
    Return True
EndFunc

Func _OpenCardpresso($sTemplateType = "Single-ID")
    ; Load configuration if not already loaded
    If $sCardpressoExePath = "" Then
        _LoadCardpressoConfig()
    EndIf
    
    Local $sCardPressoPath = $sCardpressoExePath
    
    If Not FileExists($sCardPressoPath) Then
        ConsoleWrite("WARNING: Cardpresso not found at: " & $sCardPressoPath & @CRLF)
        MsgBox(48, "Cardpresso Not Found", "Cardpresso executable not found at:" & @CRLF & $sCardPressoPath & @CRLF & @CRLF & "Please ensure Cardpresso is installed in the correct location.")
        Return False
    EndIf
    
    ; Build template file path
    Local $sTemplatePath = "%BasePath%\Templates\" & $sTemplateType & ".card"
    $sTemplatePath = _ExpandEnvironmentVariables($sTemplatePath)
    
    ; Check if template file exists
    If Not FileExists($sTemplatePath) Then
        ConsoleWrite("WARNING: Template file not found: " & $sTemplatePath & @CRLF)
        MsgBox(48, "Template Not Found", "Template file not found:" & @CRLF & $sTemplatePath & @CRLF & @CRLF & "Opening Cardpresso without template.")
        ; Open Cardpresso without template as fallback
        ConsoleWrite("Opening Cardpresso: " & $sCardPressoPath & @CRLF)
        Run('"' & $sCardPressoPath & '"')
    Else
        ; Open Cardpresso with template
        ConsoleWrite("Opening Cardpresso with template: " & $sCardPressoPath & " " & $sTemplatePath & @CRLF)
        Run('"' & $sCardPressoPath & '" "' & $sTemplatePath & '"')
    EndIf
    Return True
EndFunc

Func _ShowCardpressoPhoto($sPhotoPathRaw, $sClientID = "")
    ; Handle absolute/relative paths; normalize slashes
    Local $sPhotoPath = StringReplace($sPhotoPathRaw, "/", "\")
    
    If $sPhotoPath = "" Then
        _ShowStatus("No photo path provided.", 0)
        GUICtrlSetImage($previewPic, "")
        Return False
    EndIf
    
    ; Load configuration if not already loaded
    If $sCardpressoMasterFile = "" Then
        _LoadCardpressoConfig()
    EndIf
    
    ; Always resolve to absolute path for photo base directory
    ; Use simpler path resolution - get directory of Master file
    Local $sMasterDir = StringRegExpReplace($sCardpressoMasterFile, "\\[^\\]*$", "")
    If $sMasterDir = "" Then
        _ShowStatus("Error: Could not determine Master file directory.", 0)
        GUICtrlSetImage($previewPic, "")
        Return False
    EndIf
    
    ; Build absolute path to photo base directory + filename
    Local $sAbsolutePhotoPath = $sMasterDir & "\" & $sCardpressoPhotoBaseDir & "\" & $sPhotoPath
    
    If FileExists($sAbsolutePhotoPath) Then
        $sCardpressoPhotoPath = $sAbsolutePhotoPath
    Else
        ; File doesn't exist at expected location
        Local $sMessage = "Photo " & $sClientID & " not found." & @CRLF & @CRLF & "Please either upload the photo " & $sClientID & " or update the master file."
        MsgBox(16, "Photo Not Found", $sMessage) ; 16 = Critical icon (red X)
        GUICtrlSetImage($previewPic, "")
        Return False ; Return false to prevent CSV injection
    EndIf
    
    ; Show image in preview area
    Local $bImageResult = GUICtrlSetImage($previewPic, $sCardpressoPhotoPath)
    If $bImageResult Then
        GUICtrlSetState($previewPic, $GUI_SHOW) ; Make sure preview is visible
    EndIf
    Return True ; Return true when photo is successfully found and displayed
    If Not $bImageResult Then
        _ShowStatus("Failed to load preview:" & @CRLF & "File may be corrupted or unsupported format.", 0)
        GUICtrlSetImage($previewPic, "") ; Clear any previous image
    Else
        GUICtrlSetState($previewPic, $GUI_SHOW)
        _ShowStatus("Photo preview loaded successfully", 1)
    EndIf
EndFunc

Func _CSV($s)
    ; Quote if needed; escape internal quotes
    Local $t = String($s)
    If StringInStr($t, '"') Then $t = StringReplace($t, '"', '""')
    If StringRegExp($t, "[,\r\n]") Then $t = '"' & $t & '"'
    Return $t
EndFunc

; ==============================================================
; SFTP Upload Functions
; ==============================================================
Func _RunSFTPUpload()
    Local $sName = StringStripWS(GUICtrlRead($nameInput), 3)
    Local $sSrc = StringStripWS(GUICtrlRead($srcInput), 3)

    GUICtrlSetData($statusLabel, "")

    If $sName = "" Then
        _ShowStatus("Please enter client number.", 0)
        Return
    EndIf
    If StringLen($sName) <> 5 Or Not StringIsDigit($sName) Then
        _ShowStatus("Client number must be exactly 5 digits.", 0)
        Return
    EndIf
    If $sSrc = "" Then
        _ShowStatus("Please select source image files to upload.", 0)
        Return
    EndIf

    ; Load SFTP configuration
    _LoadSFTPConfig()
    
    ; Validate SFTP configuration
    If $sSFTPHost = "" Or $sSFTPUsername = "" Then
        _ShowStatus("SFTP configuration incomplete. Please check config.file", 0)
        Return
    EndIf

    ; Get SFTP password securely (try saved password first, then prompt if needed)
    Local $sSFTPPassword = _GetSFTPPassword()
    If $sSFTPPassword = "" Then
        _ShowStatus("SFTP upload cancelled - no password provided", 0)
        Return
    EndIf

    ; Get files from source input (could be single file or multiple files separated by |)
    Local $aFiles = StringSplit($sSrc, "|", 2) ; No count flag
    Local $iTotalFiles = UBound($aFiles)
    
    ; If only one file was selected, check if it's a valid file
    If $iTotalFiles = 1 Then
        If Not FileExists($aFiles[0]) Then
            _ShowStatus("Selected file not found: " & $aFiles[0], 0)
            Return
        EndIf
    EndIf

    ; Show progress bar
    GUICtrlSetState($progressBar, $GUI_SHOW)
    GUICtrlSetData($progressBar, 0)
    _ShowStatus("Starting batch SFTP upload... Found " & $iTotalFiles & " files", 1)

    Local $iSuccessCount = 0

    ; Upload each file
    For $i = 0 To $iTotalFiles - 1
        Local $sLocalFile = $aFiles[$i]
        
        ; Validate file exists
        If Not FileExists($sLocalFile) Then
            _ShowStatus("File not found: " & $sLocalFile, 0)
            ContinueLoop
        EndIf

        ; Update progress
        Local $iProgress = Int(($i / $iTotalFiles) * 100)
        GUICtrlSetData($progressBar, $iProgress)
        _ShowStatus("Uploading " & ($i + 1) & " of " & $iTotalFiles & ": " & StringRegExpReplace($sLocalFile, "^.*\\", ""), 1)

        ; Build remote filename - keep original filename only
        Local $sRemoteFilename = StringRegExpReplace($sLocalFile, "^.*\\", "")
        Local $sRemotePath = $sSFTPRemotePath
        If StringRight($sRemotePath, 1) <> "/" Then
            $sRemotePath &= "/"
        EndIf
        $sRemotePath &= $sRemoteFilename

        ; Check if file exists on SFTP server and ask for overwrite confirmation
        Local $bFileExists = _CheckSFTPFileExists($sRemotePath, $sSFTPPassword)
        If $bFileExists Then
            Local $iResponse = MsgBox(4 + 32 + 262144, "File Exists", "File " & $sRemoteFilename & " already exists on SFTP server." & @CRLF & "Do you want to overwrite it?")
            If $iResponse <> 6 Then ; 6 = Yes
                _ShowStatus("Skipped: " & $sRemoteFilename & " (file exists)", 1)
                ContinueLoop
            EndIf
        EndIf

        ; Upload file via SFTP
        If _UploadFileViaSFTP($sLocalFile, $sRemotePath, $sSFTPPassword) Then
            $iSuccessCount += 1
            _LogAction("SFTP Upload", $sRemotePath)
        EndIf
    Next

    ; Clear password from memory immediately after use
    $sSFTPPassword = ""

    ; Complete progress
    GUICtrlSetData($progressBar, 100)
    Sleep(500)
    GUICtrlSetState($progressBar, $GUI_HIDE)

    If $iSuccessCount > 0 Then
        _ShowStatus("SFTP upload completed: " & $iSuccessCount & " of " & $iTotalFiles & " files uploaded successfully", 1)
        
        ; Ask user if they want to clear the source files (only if all were successful)
        If $iSuccessCount = $iTotalFiles Then
            MsgBox(64, "Upload Complete", "Upload complete " & $iSuccessCount & " files uploaded successfully!")
        EndIf
    Else
        _ShowStatus("SFTP upload failed for all files. Check SFTP configuration and connection.", 0)
    EndIf
EndFunc

Func _UploadFileViaSFTP($sLocalFile, $sRemotePath, $sPassword)
    ConsoleWrite("=== Starting SFTP upload ===" & @CRLF)
    _LogSFTPError("UploadFileViaSFTP", "Starting SFTP upload", "Local: " & $sLocalFile & " | Remote: " & $sRemotePath)
    ConsoleWrite("Local file: " & $sLocalFile & @CRLF)
    ConsoleWrite("Remote path: " & $sRemotePath & @CRLF)
    ConsoleWrite("SFTP Host: " & $sSFTPHost & @CRLF)
    ConsoleWrite("SFTP Username: " & $sSFTPUsername & @CRLF)
    ConsoleWrite("SFTP Password length: " & StringLen($sPassword) & " characters" & @CRLF)

    ; Check if WinSCP is available - use portable version first
    Local $sWinSCPPath = "C:\Tools\WinSCP-6.5.5-Portable\WinSCP.com"
    If Not FileExists($sWinSCPPath) Then
        $sWinSCPPath = "C:\Tools\WinSCP-6.5.5-Portable\WinSCP.exe"
        If Not FileExists($sWinSCPPath) Then
            ; Fallback to default Program Files location
            $sWinSCPPath = @ProgramFilesDir & "\WinSCP\WinSCP.com"
            If Not FileExists($sWinSCPPath) Then
                $sWinSCPPath = @ProgramFilesDir & "\WinSCP\WinSCP.exe"
                If Not FileExists($sWinSCPPath) Then
                    MsgBox(16, "WinSCP Not Found", "WinSCP is not found in portable location (C:\Tools\WinSCP-6.5.5-Portable) or default Program Files location." & @CRLF & _
                           "Please ensure WinSCP is available in the portable location or install WinSCP from https://winscp.net/")
                    Return False
                EndIf
            EndIf
        EndIf
    EndIf

    ConsoleWrite("Using WinSCP path: " & $sWinSCPPath & @CRLF)

    ; Create WinSCP script file
    Local $sScriptFile = @TempDir & "\winscp_upload_" & @MSEC & ".txt"
    Local $hScriptFile = FileOpen($sScriptFile, 2) ; 2 = overwrite mode
    If $hScriptFile = -1 Then
        ConsoleWrite("ERROR: Could not create WinSCP script file" & @CRLF)
        _LogSFTPError("UploadFileViaSFTP", "Could not create WinSCP script file")
        Return False
    EndIf

    ; Write WinSCP script
    FileWriteLine($hScriptFile, "option batch abort")
    FileWriteLine($hScriptFile, "option confirm off")
    FileWriteLine($hScriptFile, "open sftp://" & $sSFTPUsername & ":" & $sPassword & "@" & $sSFTPHost & " -hostkey=*")
    FileWriteLine($hScriptFile, "put """ & $sLocalFile & """ """ & $sRemotePath & """")
    FileWriteLine($hScriptFile, "exit")
    FileClose($hScriptFile)

    ConsoleWrite("Created WinSCP script: " & $sScriptFile & @CRLF)

    ; Update progress
    GUICtrlSetData($progressBar, 30)
    _ShowStatus("Connecting to SFTP server...", 1)

    ; Run WinSCP with the script
    Local $sCmd
    Local $sLogFile = @TempDir & "\winscp_log_" & @MSEC & ".txt"
    If StringRight($sWinSCPPath, 4) = ".com" Then
        $sCmd = '"' & $sWinSCPPath & '" /script="' & $sScriptFile & '" /log="' & $sLogFile & '"'
    Else
        $sCmd = '"' & $sWinSCPPath & '" /console /script="' & $sScriptFile & '" /log="' & $sLogFile & '"'
    EndIf

    ConsoleWrite("Running WinSCP command: " & $sCmd & @CRLF)

    ; Update progress
    GUICtrlSetData($progressBar, 60)
    _ShowStatus("Uploading file to SFTP server...", 1)

    Local $iPID = Run($sCmd, "", @SW_HIDE)
    If @error Then
        ConsoleWrite("ERROR: Failed to start WinSCP: " & @error & @CRLF)
        _LogSFTPError("UploadFileViaSFTP", "Failed to start WinSCP", "Error: " & @error)
        FileDelete($sScriptFile)
        Return False
    EndIf

    ; Wait for process to complete with timeout
    Local $iTimeout = 30000 ; 30 seconds
    Local $hTimer = TimerInit()
    While ProcessExists($iPID)
        If TimerDiff($hTimer) > $iTimeout Then
            ConsoleWrite("ERROR: WinSCP process timeout after " & $iTimeout & "ms" & @CRLF)
            _LogSFTPError("UploadFileViaSFTP", "WinSCP process timeout", "Timeout: " & $iTimeout & "ms")
            ProcessClose($iPID)
            FileDelete($sScriptFile)
            Return False
        EndIf
        Sleep(100)
    WEnd
    Local $iExitCode = @extended

    ConsoleWrite("WinSCP process completed with exit code: " & $iExitCode & @CRLF)

    ; Update progress
    GUICtrlSetData($progressBar, 90)
    _ShowStatus("Finalizing upload...", 1)

    ; Check exit code and log file
    Local $bSuccess = False
    Local $sLogContent = ""
    
    If FileExists($sLogFile) Then
        $sLogContent = FileRead($sLogFile)
        ConsoleWrite("WinSCP Log Content:" & @CRLF & $sLogContent & @CRLF)
        
        ; Check for success indicators based on actual WinSCP log patterns
        If StringInStr($sLogContent, "Transfer done") Or StringInStr($sLogContent, "successful") Or StringInStr($sLogContent, "Uploading file") Or StringInStr($sLogContent, "100%") Then
            $bSuccess = True
        ElseIf StringInStr($sLogContent, "Authentication failed") Then
            ConsoleWrite("ERROR: Authentication failed - check username/password" & @CRLF)
            _LogSFTPError("UploadFileViaSFTP", "Authentication failed", $sLogContent)
        ElseIf StringInStr($sLogContent, "Connection failed") Then
            ConsoleWrite("ERROR: Connection failed - check host/port" & @CRLF)
            _LogSFTPError("UploadFileViaSFTP", "Connection failed", $sLogContent)
        ElseIf StringInStr($sLogContent, "No such file or directory") Then
            ConsoleWrite("ERROR: Remote directory doesn't exist" & @CRLF)
            _LogSFTPError("UploadFileViaSFTP", "Remote directory doesn't exist", $sLogContent)
        ElseIf StringInStr($sLogContent, "Permission denied") Then
            ConsoleWrite("ERROR: Permission denied - check user permissions" & @CRLF)
            _LogSFTPError("UploadFileViaSFTP", "Permission denied", $sLogContent)
        ElseIf $sLogContent <> "" Then
            _LogSFTPError("UploadFileViaSFTP", "Unknown error in WinSCP log", $sLogContent)
        EndIf
        
        FileDelete($sLogFile)
    Else
        ConsoleWrite("WARNING: No WinSCP log file found" & @CRLF)
    EndIf

    ; Clean up script file
    FileDelete($sScriptFile)

    If $bSuccess Then
        ConsoleWrite("SFTP upload successful" & @CRLF)
        _LogSFTPError("UploadFileViaSFTP", "SFTP upload successful")
    Else
        ; Check if upload actually succeeded despite error detection
        If $iExitCode = 0 And StringInStr($sLogContent, "Transfer done") Then
            ConsoleWrite("SFTP upload actually succeeded (Transfer done detected)" & @CRLF)
            _LogSFTPError("UploadFileViaSFTP", "SFTP upload successful (Transfer done detected)")
            $bSuccess = True
        Else
            ConsoleWrite("SFTP upload failed. Exit code: " & $iExitCode & @CRLF)
            If $sLogContent = "" Then
                ConsoleWrite("No detailed error information available" & @CRLF)
                _LogSFTPError("UploadFileViaSFTP", "No detailed error information", "Exit code: " & $iExitCode)
            EndIf
        EndIf
    EndIf

    ConsoleWrite("=== Finished SFTP upload ===" & @CRLF)
    Return $bSuccess
EndFunc

; ==============================================================
; SFTP File Existence Check Function
; ==============================================================
Func _CheckSFTPFileExists($sRemotePath, $sPassword)
    ConsoleWrite("=== Checking if file exists on SFTP server ===" & @CRLF)
    ConsoleWrite("Remote path: " & $sRemotePath & @CRLF)
    
    ; Check if WinSCP is available - use portable version first
    Local $sWinSCPPath = "C:\Tools\WinSCP-6.5.5-Portable\WinSCP.com"
    If Not FileExists($sWinSCPPath) Then
        $sWinSCPPath = "C:\Tools\WinSCP-6.5.5-Portable\WinSCP.exe"
        If Not FileExists($sWinSCPPath) Then
            ; Fallback to default Program Files location
            $sWinSCPPath = @ProgramFilesDir & "\WinSCP\WinSCP.com"
            If Not FileExists($sWinSCPPath) Then
                $sWinSCPPath = @ProgramFilesDir & "\WinSCP\WinSCP.exe"
                If Not FileExists($sWinSCPPath) Then
                    ConsoleWrite("WARNING: WinSCP not found, assuming file doesn't exist" & @CRLF)
                    Return False
                EndIf
            EndIf
        EndIf
    EndIf

    ; Create WinSCP script file to check file existence
    Local $sScriptFile = @TempDir & "\winscp_check_" & @MSEC & ".txt"
    Local $hScriptFile = FileOpen($sScriptFile, 2) ; 2 = overwrite mode
    If $hScriptFile = -1 Then
        ConsoleWrite("ERROR: Could not create WinSCP script file" & @CRLF)
        Return False
    EndIf

    ; Write WinSCP script to check file existence
    FileWriteLine($hScriptFile, "option batch abort")
    FileWriteLine($hScriptFile, "option confirm off")
    FileWriteLine($hScriptFile, "open sftp://" & $sSFTPUsername & ":" & $sPassword & "@" & $sSFTPHost & " -hostkey=*")
    FileWriteLine($hScriptFile, "stat """ & $sRemotePath & """")
    FileWriteLine($hScriptFile, "exit")
    FileClose($hScriptFile)

    ConsoleWrite("Created WinSCP check script: " & $sScriptFile & @CRLF)

    ; Run WinSCP with the script
    Local $sCmd
    Local $sLogFile = @TempDir & "\winscp_check_log_" & @MSEC & ".txt"
    If StringRight($sWinSCPPath, 4) = ".com" Then
        $sCmd = '"' & $sWinSCPPath & '" /script="' & $sScriptFile & '" /log="' & $sLogFile & '"'
    Else
        $sCmd = '"' & $sWinSCPPath & '" /console /script="' & $sScriptFile & '" /log="' & $sLogFile & '"'
    EndIf

    ConsoleWrite("Running WinSCP check command: " & $sCmd & @CRLF)

    Local $iPID = Run($sCmd, "", @SW_HIDE)
    If @error Then
        ConsoleWrite("ERROR: Failed to start WinSCP for file check: " & @error & @CRLF)
        FileDelete($sScriptFile)
        Return False
    EndIf

    ; Wait for process to complete with timeout
    Local $iTimeout = 15000 ; 15 seconds
    Local $hTimer = TimerInit()
    While ProcessExists($iPID)
        If TimerDiff($hTimer) > $iTimeout Then
            ConsoleWrite("ERROR: WinSCP check process timeout" & @CRLF)
            ProcessClose($iPID)
            FileDelete($sScriptFile)
            Return False
        EndIf
        Sleep(100)
    WEnd

    ; Check log file for file existence
    Local $bFileExists = False
    If FileExists($sLogFile) Then
        Local $sLogContent = FileRead($sLogFile)
        ConsoleWrite("WinSCP Check Log Content:" & @CRLF & $sLogContent & @CRLF)
        
        ; Check if file exists - if stat command succeeds, file exists
        ; If file doesn't exist, WinSCP will show an error message
        If Not StringInStr($sLogContent, "No such file or directory") And _
           Not StringInStr($sLogContent, "File or folder") And _
           Not StringInStr($sLogContent, "does not exist") Then
            $bFileExists = True
        EndIf
    EndIf

    ; Cleanup
    FileDelete($sScriptFile)
    If FileExists($sLogFile) Then FileDelete($sLogFile)

    ConsoleWrite("File exists on SFTP server: " & $bFileExists & @CRLF)
    Return $bFileExists
EndFunc

Func _GetSFTPPassword()
    ; First try to get saved password from encrypted file storage
    Local $sSavedPassword = _GetSavedSFTPPassword()
    
    If $sSavedPassword <> "" Then
        ConsoleWrite("SFTP password retrieved from encrypted storage" & @CRLF)
        Return $sSavedPassword
    Else
        ConsoleWrite("No saved SFTP password found - prompting for new password" & @CRLF)
        Return _PromptForSFTPPassword()
    EndIf
EndFunc

Func _PromptForSFTPPassword()
    ; Check if we have a saved password
    Local $bHasSavedPassword = (_GetSavedSFTPPassword() <> "")
    
    ; Create secure password input dialog with save option
    Local $hPasswordGUI = GUICreate("SFTP Authentication", 350, 220, -1, -1, $WS_CAPTION + $WS_SYSMENU)
    GUISetFont(9, 400, 0, "Segoe UI")
    
    ; Server info
    GUICtrlCreateLabel("Server: " & $sSFTPHost, 20, 20, 310, 20)
    GUICtrlCreateLabel("Username: " & $sSFTPUsername, 20, 40, 310, 20)
    
    ; Password status info
    If $bHasSavedPassword Then
        GUICtrlCreateLabel("Saved password found - please re-enter for security", 20, 60, 310, 20)
        GUICtrlSetColor(-1, 0x008000) ; Green text
    Else
        GUICtrlCreateLabel("No saved password found - enter new password", 20, 60, 310, 20)
        GUICtrlSetColor(-1, 0xFF0000) ; Red text
    EndIf
    
    ; Password input
    GUICtrlCreateLabel("Enter SFTP Password:", 20, 90, 310, 20)
    Local $hPasswordInput = GUICtrlCreateInput("", 20, 110, 310, 24, $ES_PASSWORD) ; Password style
    GUICtrlSetState($hPasswordInput, $GUI_FOCUS)
    
    ; Save password checkbox
    Local $hSavePassword = GUICtrlCreateCheckbox("Save password in encrypted file", 20, 140, 310, 20)
    If $bHasSavedPassword Then
        GUICtrlSetState($hSavePassword, $GUI_CHECKED) ; Default to checked if password exists
        GUICtrlSetData($hSavePassword, "Update saved password")
    Else
        GUICtrlSetState($hSavePassword, $GUI_CHECKED) ; Default to checked for new passwords
    EndIf
    
    ; Info about where password is saved
    GUICtrlCreateLabel("Password will be saved in encrypted file: sftp_password.dat", 20, 165, 310, 20)
    GUICtrlSetFont(-1, 8) ; Smaller font
    GUICtrlSetColor(-1, 0x666666) ; Gray text
    
    ; Buttons
    Local $hOKButton = GUICtrlCreateButton("OK", 20, 190, 150, 25)
    Local $hCancelButton = GUICtrlCreateButton("Cancel", 180, 190, 150, 25)
    
    GUISetState(@SW_SHOW)
    
    Local $sPassword = ""
    Local $bCancelled = False
    Local $bSavePassword = True
    
    ; Event loop
    While 1
        Switch GUIGetMsg()
            Case $GUI_EVENT_CLOSE, $hCancelButton
                $bCancelled = True
                ExitLoop
            Case $hOKButton
                $sPassword = GUICtrlRead($hPasswordInput)
                $bSavePassword = (GUICtrlRead($hSavePassword) = $GUI_CHECKED)
                ExitLoop
        EndSwitch
    WEnd
    
    GUIDelete($hPasswordGUI)
    
    If $bCancelled Then
        Return ""
    Else
        ; Save password if requested
        If $bSavePassword And $sPassword <> "" Then
            _SaveSFTPPassword($sPassword)
        EndIf
        Return $sPassword
    EndIf
EndFunc

Func _SaveSFTPPassword($sPassword)
    ; Save password securely using encrypted file storage
    Local $sPasswordFile = @ScriptDir & "\sftp_password.dat"
    
    ; Encrypt the password before saving
    Local $sEncryptedPassword = _EncryptPassword($sPassword)
    
    ; Save encrypted password to file
    Local $hFile = FileOpen($sPasswordFile, 2) ; 2 = overwrite mode
    If $hFile = -1 Then
        ConsoleWrite("ERROR: Could not create password file: " & $sPasswordFile & @CRLF)
        Return
    EndIf
    
    FileWrite($hFile, $sEncryptedPassword)
    FileClose($hFile)
    
    ; Set file attributes to hidden for additional security
    FileSetAttrib($sPasswordFile, "+H")
    
    ConsoleWrite("SFTP password successfully saved to encrypted storage" & @CRLF)
EndFunc

Func _GetSavedSFTPPassword()
    ; Try to retrieve saved password from encrypted file storage
    Local $sPasswordFile = @ScriptDir & "\sftp_password.dat"
    
    If Not FileExists($sPasswordFile) Then
        ConsoleWrite("No saved SFTP password file found" & @CRLF)
        Return ""
    EndIf
    
    ; Read encrypted password from file
    Local $hFile = FileOpen($sPasswordFile, 0) ; 0 = read mode
    If $hFile = -1 Then
        ConsoleWrite("ERROR: Could not open password file: " & $sPasswordFile & @CRLF)
        Return ""
    EndIf
    
    Local $sEncryptedPassword = FileRead($hFile)
    FileClose($hFile)
    
    ; Decrypt the password
    Local $sPassword = _DecryptPassword($sEncryptedPassword)
    
    If $sPassword <> "" Then
        ConsoleWrite("SFTP password successfully retrieved from encrypted storage" & @CRLF)
        Return $sPassword
    Else
        ConsoleWrite("ERROR: Failed to decrypt SFTP password" & @CRLF)
        Return ""
    EndIf
EndFunc

; ==============================================================
; Password Encryption/Decryption Functions
; ==============================================================
Func _EncryptPassword($sPassword)
    ; Simple XOR encryption with a fixed key for basic obfuscation
    ; Note: This is not military-grade encryption but provides basic protection
    Local $sKey = "PhotoManagerSFTP2024" ; Fixed encryption key
    Local $sEncrypted = ""
    
    For $i = 1 To StringLen($sPassword)
        Local $cChar = StringMid($sPassword, $i, 1)
        Local $cKey = StringMid($sKey, Mod($i - 1, StringLen($sKey)) + 1, 1)
        Local $iEncryptedChar = BitXOR(Asc($cChar), Asc($cKey))
        $sEncrypted &= Hex($iEncryptedChar, 2)
    Next
    
    Return $sEncrypted
EndFunc

Func _DecryptPassword($sEncrypted)
    ; Decrypt XOR-encrypted password
    Local $sKey = "PhotoManagerSFTP2024" ; Same encryption key
    Local $sPassword = ""
    
    ; Check if the encrypted string has even length (hex pairs)
    If Mod(StringLen($sEncrypted), 2) <> 0 Then
        ConsoleWrite("ERROR: Invalid encrypted password format" & @CRLF)
        Return ""
    EndIf
    
    For $i = 1 To StringLen($sEncrypted) Step 2
        Local $sHexPair = StringMid($sEncrypted, $i, 2)
        Local $iEncryptedChar = Dec($sHexPair)
        Local $cKey = StringMid($sKey, Mod(($i - 1) / 2, StringLen($sKey)) + 1, 1)
        Local $cChar = Chr(BitXOR($iEncryptedChar, Asc($cKey)))
        $sPassword &= $cChar
    Next
    
    Return $sPassword
EndFunc

; ==============================================================
; SFTP Logging Functions
; ==============================================================
Func _LogSFTPError($sFunction, $sMessage, $sDetails = "")
    Local $sLogFile = @ScriptDir & "\SFTP_Error_Log.txt"
    Local $sTimestamp = @YEAR & "-" & @MON & "-" & @MDAY & " " & @HOUR & ":" & @MIN & ":" & @SEC
    Local $sLogEntry = $sTimestamp & " - " & $sFunction & " - " & $sMessage & @CRLF
    
    If $sDetails <> "" Then
        $sLogEntry &= "Details: " & $sDetails & @CRLF
    EndIf
    
    $sLogEntry &= "Configuration:" & @CRLF
    $sLogEntry &= "  Host: " & $sSFTPHost & @CRLF
    $sLogEntry &= "  Username: " & $sSFTPUsername & @CRLF
    $sLogEntry &= "  Remote Path: " & $sSFTPRemotePath & @CRLF
    $sLogEntry &= "----------------------------------------" & @CRLF & @CRLF
    
    Local $hLogFile = FileOpen($sLogFile, 1) ; 1 = append mode
    If $hLogFile <> -1 Then
        FileWrite($hLogFile, $sLogEntry)
        FileClose($hLogFile)
        ConsoleWrite("SFTP Error logged to: " & $sLogFile & @CRLF)
    EndIf
EndFunc

Func _SaveWinSCPLog($sLogContent, $sOperation)
    Local $sLogFile = @ScriptDir & "\WinSCP_Log_" & $sOperation & "_" & @YEAR & @MON & @MDAY & "_" & @HOUR & @MIN & @SEC & ".txt"
    Local $hLogFile = FileOpen($sLogFile, 2) ; 2 = overwrite mode
    If $hLogFile <> -1 Then
        FileWrite($hLogFile, $sLogContent)
        FileClose($hLogFile)
        ConsoleWrite("WinSCP log saved to: " & $sLogFile & @CRLF)
        Return $sLogFile
    EndIf
    Return ""
EndFunc

; ==============================================================
; SFTP Test Function
; ==============================================================
Func _TestSFTPConnection()
    ConsoleWrite("=== Testing SFTP Connection ===" & @CRLF)
    _LogSFTPError("TestSFTPConnection", "Starting SFTP connection test")
    
    ; Load SFTP configuration
    _LoadSFTPConfig()
    
    ; Validate SFTP configuration
    If $sSFTPHost = "" Or $sSFTPUsername = "" Then
        _LogSFTPError("TestSFTPConnection", "SFTP configuration incomplete")
        MsgBox(16, "SFTP Configuration Error", "SFTP configuration incomplete. Please check config.file")
        Return
    EndIf
    
    ConsoleWrite("SFTP Configuration:" & @CRLF)
    ConsoleWrite("  Host: " & $sSFTPHost & @CRLF)
    ConsoleWrite("  Username: " & $sSFTPUsername & @CRLF)
    ConsoleWrite("  Remote Path: " & $sSFTPRemotePath & @CRLF)
    
    ; Get SFTP password
    Local $sSFTPPassword = _GetSFTPPassword()
    If $sSFTPPassword = "" Then
        _LogSFTPError("TestSFTPConnection", "No SFTP password provided")
        MsgBox(16, "SFTP Test", "No SFTP password provided. Test cancelled.")
        Return
    EndIf
    
    ; Check if WinSCP is available
    Local $sWinSCPPath = "C:\Tools\WinSCP-6.5.5-Portable\WinSCP.com"
    If Not FileExists($sWinSCPPath) Then
        $sWinSCPPath = "C:\Tools\WinSCP-6.5.5-Portable\WinSCP.exe"
        If Not FileExists($sWinSCPPath) Then
            _LogSFTPError("TestSFTPConnection", "WinSCP not found in portable location")
            MsgBox(16, "WinSCP Not Found", "WinSCP not found in portable location: C:\Tools\WinSCP-6.5.5-Portable")
            Return
        EndIf
    EndIf
    
    ConsoleWrite("Using WinSCP: " & $sWinSCPPath & @CRLF)
    
    ; Create test script
    Local $sScriptFile = @TempDir & "\winscp_test_" & @MSEC & ".txt"
    Local $hScriptFile = FileOpen($sScriptFile, 2)
    If $hScriptFile = -1 Then
        _LogSFTPError("TestSFTPConnection", "Could not create test script file")
        MsgBox(16, "Error", "Could not create test script file")
        Return
    EndIf
    
    ; Write test script
    FileWriteLine($hScriptFile, "option batch abort")
    FileWriteLine($hScriptFile, "option confirm off")
    FileWriteLine($hScriptFile, "open sftp://" & $sSFTPUsername & ":" & $sSFTPPassword & "@" & $sSFTPHost & " -hostkey=*")
    FileWriteLine($hScriptFile, "ls """ & $sSFTPRemotePath & """")
    FileWriteLine($hScriptFile, "exit")
    FileClose($hScriptFile)
    
    ; Run test
    Local $sCmd
    Local $sLogFile = @TempDir & "\winscp_test_log_" & @MSEC & ".txt"
    If StringRight($sWinSCPPath, 4) = ".com" Then
        $sCmd = '"' & $sWinSCPPath & '" /script="' & $sScriptFile & '" /log="' & $sLogFile & '"'
    Else
        $sCmd = '"' & $sWinSCPPath & '" /console /script="' & $sScriptFile & '" /log="' & $sLogFile & '"'
    EndIf
    
    ConsoleWrite("Running test command: " & $sCmd & @CRLF)
    
    Local $iPID = Run($sCmd, "", @SW_HIDE)
    If @error Then
        _LogSFTPError("TestSFTPConnection", "Failed to start WinSCP test", "Error: " & @error)
        MsgBox(16, "Test Failed", "Failed to start WinSCP test")
        FileDelete($sScriptFile)
        Return
    EndIf
    
    ; Wait for process
    ProcessWaitClose($iPID)
    Local $iExitCode = @extended
    
    ; Read log file
    Local $sLogContent = ""
    If FileExists($sLogFile) Then
        $sLogContent = FileRead($sLogFile)
        ConsoleWrite("Test Log:" & @CRLF & $sLogContent & @CRLF)
        ; Save WinSCP log permanently
        _SaveWinSCPLog($sLogContent, "Test")
        FileDelete($sLogFile)
    EndIf
    
    ; Clean up
    FileDelete($sScriptFile)
    
    ; Analyze results
    If $iExitCode = 0 And (StringInStr($sLogContent, "successful") Or StringInStr($sLogContent, "11111_IMG-001.png") Or StringInStr($sLogContent, "Listing directory")) Then
        _LogSFTPError("TestSFTPConnection", "SFTP connection test successful")
        MsgBox(64, "SFTP Test Success", "SFTP connection test successful!" & @CRLF & @CRLF & _
               "Connection to " & $sSFTPHost & " established successfully." & @CRLF & @CRLF & _
               "Uploaded file is visible in /uploads directory.")
    Else
        Local $sErrorMsg = "SFTP connection test failed." & @CRLF & @CRLF
        $sErrorMsg &= "Exit code: " & $iExitCode & @CRLF & @CRLF
        
        If StringInStr($sLogContent, "Authentication failed") Then
            $sErrorMsg &= "Error: Authentication failed" & @CRLF & "Check username and password"
            _LogSFTPError("TestSFTPConnection", "Authentication failed", $sLogContent)
        ElseIf StringInStr($sLogContent, "Connection failed") Then
            $sErrorMsg &= "Error: Connection failed" & @CRLF & "Check host address and network connectivity"
            _LogSFTPError("TestSFTPConnection", "Connection failed", $sLogContent)
        ElseIf StringInStr($sLogContent, "No such file or directory") Then
            $sErrorMsg &= "Error: Remote directory doesn't exist" & @CRLF & "Check remote path configuration"
            _LogSFTPError("TestSFTPConnection", "Remote directory doesn't exist", $sLogContent)
        ElseIf StringInStr($sLogContent, "Permission denied") Then
            $sErrorMsg &= "Error: Permission denied" & @CRLF & "Check user permissions on server"
            _LogSFTPError("TestSFTPConnection", "Permission denied", $sLogContent)
        ElseIf $sLogContent <> "" Then
            $sErrorMsg &= "Log output:" & @CRLF & $sLogContent
            _LogSFTPError("TestSFTPConnection", "Unknown error", $sLogContent)
        Else
            $sErrorMsg &= "No detailed error information available"
            _LogSFTPError("TestSFTPConnection", "No log output available", "Exit code: " & $iExitCode)
        EndIf
        
        MsgBox(16, "SFTP Test Failed", $sErrorMsg)
    EndIf
    
    ConsoleWrite("=== Finished SFTP Connection Test ===" & @CRLF)
EndFunc

Func _ClearSavedSFTPPassword()
    ; Clear saved SFTP password from encrypted file storage
    Local $sPasswordFile = @ScriptDir & "\sftp_password.dat"
    
    If FileExists($sPasswordFile) Then
        FileDelete($sPasswordFile)
        ConsoleWrite("SFTP password cleared from encrypted storage" & @CRLF)
    Else
        ConsoleWrite("No saved SFTP password file found to clear" & @CRLF)
    EndIf
EndFunc

; ==============================================================
; Environment Variable Expansion Function
; ==============================================================
Func _ExpandEnvironmentVariables($sPath)
    ; First expand Windows environment variables (like %USERPROFILE%)
    If StringInStr($sPath, "%") Then
        ; Use Windows API to expand environment variables
        Local $aResult = DllCall("kernel32.dll", "dword", "ExpandEnvironmentStringsW", "wstr", $sPath, "wstr", "", "dword", 4096)
        If Not @error And $aResult[0] > 0 Then
            $sPath = $aResult[2]
        EndIf
    EndIf
    
    ; Now handle custom %BasePath% variable from config file
    If StringInStr($sPath, "%BasePath%") Then
        ; Get the BasePath value from config file
        Local $sBasePath = _GetBasePathFromConfig()
        If $sBasePath <> "" Then
            $sPath = StringReplace($sPath, "%BasePath%", $sBasePath)
        EndIf
    EndIf
    
    Return $sPath
EndFunc

; ==============================================================
; Get BasePath from Config File
; ==============================================================
Func _GetBasePathFromConfig()
    Local $sConfig = $sScriptDir & "\config.file"
    If Not FileExists($sConfig) Then Return ""
    
    Local $aLines = FileReadToArray($sConfig)
    For $i = 0 To UBound($aLines) - 1
        Local $sLine = StringStripWS($aLines[$i], 3)
        If StringInStr($sLine, "BasePath") And Not StringInStr($sLine, "Folder") Then
            Local $sBasePath = StringRegExpReplace($sLine, '(?i)^.*=\s*"?([^"]+)"?\s*$', "$1")
            ; Expand any Windows environment variables in BasePath itself
            If StringInStr($sBasePath, "%") Then
                Local $aResult = DllCall("kernel32.dll", "dword", "ExpandEnvironmentStringsW", "wstr", $sBasePath, "wstr", "", "dword", 4096)
                If Not @error And $aResult[0] > 0 Then
                    $sBasePath = $aResult[2]
                EndIf
            EndIf
            Return $sBasePath
        EndIf
    Next
    Return ""
EndFunc

Func _ClearCroppedFolder()
    ; Clear all files from the Cropped folder
    Local $sCroppedFolder = $sScriptDir & "\Cropped"
    
    If Not FileExists($sCroppedFolder) Then
        Return False
    EndIf
    
    Local $aFiles = _FileListToArray($sCroppedFolder, "*", 1)
    If @error Or Not IsArray($aFiles) Or $aFiles[0] = 0 Then
        Return True ; Folder is already empty
    EndIf
    
    Local $iDeletedCount = 0
    For $i = 1 To $aFiles[0]
        Local $sFile = $sCroppedFolder & "\" & $aFiles[$i]
        If FileDelete($sFile) Then
            $iDeletedCount += 1
        EndIf
    Next
    
    ConsoleWrite("Cleared Cropped folder: " & $iDeletedCount & " files deleted" & @CRLF)
    _ShowStatus("Cropped folder cleared: " & $iDeletedCount & " files removed", 1)
    Return True
EndFunc
