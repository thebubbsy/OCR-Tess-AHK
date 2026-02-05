#SingleInstance Force

; ====================================================================
; OCR Clipboard to Text
; Hotkey: Ctrl+Shift+V
;
; Extracts text from clipboard images using Tesseract OCR with
; ImageMagick preprocessing for improved accuracy
; ====================================================================

; Configuration
TESSERACT_ENV_VAR := "TESSERACT_PATH"
IMAGEMAGICK_ENV_VAR := "MAGICKBUBBSY"
TEMP_DIR := A_Temp "\ClipboardOCR\"
HOTKEY := "^+V"

; Initialize: Detect Tesseract Path
EnvGet, TESSERACT_PATH, %TESSERACT_ENV_VAR%
if (TESSERACT_PATH = "" || !FileExist(TESSERACT_PATH)) {
    MsgBox, 0x40, Setup Required, Tesseract OCR not found. Please locate tesseract.exe
    FileSelectFile, SelectedPath, 3,, Please locate the Tesseract executable (tesseract.exe)
    if (SelectedPath = "" || !(InStr(SelectedPath, "tesseract.exe"))) {
        MsgBox, 0x10, Error, Invalid file selected. Script will now exit.
        ExitApp
    }
    RunWait, % "setx " TESSERACT_ENV_VAR " """ SelectedPath """", , Hide
    TESSERACT_PATH := SelectedPath
    MsgBox, 0x40, Success, Tesseract path saved. Script is ready to use.
}

; Initialize: Detect ImageMagick Path
EnvGet, IMAGEMAGICK_PATH, %IMAGEMAGICK_ENV_VAR%
if (IMAGEMAGICK_PATH = "" || !FileExist(IMAGEMAGICK_PATH)) {
    MsgBox, 0x40, Setup Required, ImageMagick not found. Please locate magick.exe
    FileSelectFile, SelectedPath, 3,, Please locate the ImageMagick executable (magick.exe)
    if (SelectedPath = "" || !(InStr(SelectedPath, "magick.exe"))) {
        MsgBox, 0x10, Error, Invalid file selected. Script will now exit.
        ExitApp
    }
    RunWait, % "setx " IMAGEMAGICK_ENV_VAR " """ SelectedPath """", , Hide
    IMAGEMAGICK_PATH := SelectedPath
    MsgBox, 0x40, Success, ImageMagick path saved. Script is ready to use.
}

; Register Hotkey: Ctrl + Shift + V
Hotkey, %HOTKEY%, PerformOCR
return

; ====================================================================
; PerformOCR - Main OCR workflow
; Triggered by: Ctrl+Shift+V
; Process:
;   1. Save clipboard image to temp file
;   2. Preprocess with ImageMagick
;   3. Run Tesseract OCR
;   4. Copy result to clipboard
;   5. Clean up temp files
; ====================================================================
PerformOCR:
{
    FileCreateDir, %TEMP_DIR%
    TempBmpFile := TEMP_DIR "clipboard_image.bmp"

    if !SaveClipboardImage(TempBmpFile) {
        ToolTip, No image found on clipboard
        SetTimer, RemoveToolTip, -2000
        return
    }

    ToolTip, Processing image...
    if !RunImageMagick(IMAGEMAGICK_PATH, TempBmpFile) {
        ToolTip, Failed to preprocess the image
        SetTimer, RemoveToolTip, -3000
        FileRemoveDir, %TEMP_DIR%, 1
        return
    }

    ProcessedFile := TEMP_DIR "clipboard_image_processed.bmp"
    TempResultFile := TEMP_DIR "ocr_result"
    
    ToolTip, Running OCR...
    if !RunTesseractOCR(TESSERACT_PATH, ProcessedFile, TempResultFile) {
        ToolTip, Tesseract failed to generate output
        SetTimer, RemoveToolTip, -3000
        FileRemoveDir, %TEMP_DIR%, 1
        return
    }

    FileRead, OCRResult, % TempResultFile ".txt"
    if (OCRResult = "") {
        ToolTip, No text detected in image
        SetTimer, RemoveToolTip, -3000
    } else {
        Clipboard := OCRResult
        ClipWait, 1
        ToolTip, Text copied to clipboard!
        SetTimer, RemoveToolTip, -2000
    }

    FileRemoveDir, %TEMP_DIR%, 1
}
return

RemoveToolTip:
    ToolTip
return

; ====================================================================
; SaveClipboardImage - Extract bitmap from clipboard and save as BMP
; Parameters:
;   FilePath - Destination path for the BMP file
; Returns: true if successful, false otherwise
; ====================================================================
SaveClipboardImage(FilePath) {
    static CF_DIB := 8

    if !DllCall("OpenClipboard", "Ptr", 0) {
        return false
    }

    if !DllCall("IsClipboardFormatAvailable", "UInt", CF_DIB) {
        DllCall("CloseClipboard")
        return false
    }

    hDIB := DllCall("GetClipboardData", "UInt", CF_DIB, "Ptr")
    if !hDIB {
        DllCall("CloseClipboard")
        return false
    }

    pData := DllCall("GlobalLock", "Ptr", hDIB, "Ptr")
    if !pData {
        DllCall("CloseClipboard")
        return false
    }

    Size := DllCall("GlobalSize", "Ptr", hDIB, "UInt")
    if (Size <= 0) {
        DllCall("GlobalUnlock", "Ptr", hDIB)
        DllCall("CloseClipboard")
        return false
    }

    FileDelete, %FilePath%
    File := FileOpen(FilePath, "w")
    if !File {
        DllCall("GlobalUnlock", "Ptr", hDIB)
        DllCall("CloseClipboard")
        return false
    }

    VarSetCapacity(Buffer, Size, 0)
    DllCall("RtlMoveMemory", "Ptr", &Buffer, "Ptr", pData, "UInt", Size)
    File.RawWrite(&Buffer, Size)
    File.Close()

    DllCall("GlobalUnlock", "Ptr", hDIB)
    DllCall("CloseClipboard")

    return FileExist(FilePath)
}

; ====================================================================
; RunImageMagick - Preprocess image for optimal OCR results
; Applies: density increase, grayscale, deskew, resize, contrast, 
;          sharpen, and despeckle filters
; Parameters:
;   IMAGEMAGICK_PATH - Path to magick.exe
;   InputFile - Source image file path
; Returns: true if processed file was created successfully
; ====================================================================
RunImageMagick(IMAGEMAGICK_PATH, InputFile) {
    ProcessedFile := RegExReplace(InputFile, "\.bmp$", "_processed.bmp", 1)
    RunWait, % IMAGEMAGICK_PATH . " """ . InputFile . """ -density 300 -strip -colorspace Gray -deskew 40% -resize 150% -contrast-stretch 5%x5% -sharpen 0x1.0 -despeckle """ . ProcessedFile """", , Hide
    return FileExist(ProcessedFile)
}

; ====================================================================
; RunTesseractOCR - Execute Tesseract OCR with optimized parameters
; Parameters:
;   TESSERACT_PATH - Path to tesseract.exe
;   InputFile - Preprocessed image file path
;   OutputFile - Base path for output (without extension)
; Returns: true if OCR output file was created successfully
; ====================================================================
RunTesseractOCR(TESSERACT_PATH, InputFile, OutputFile) {
    RunWait, "%TESSERACT_PATH%" "%InputFile%" "%OutputFile%" -l eng --oem 3 --psm 6 --dpi 600 -c preserve_interword_spaces=1 -c tessedit_char_unrej_keep_perfect_wd=1, , Hide
    return FileExist(OutputFile ".txt")
}
