#SingleInstance Force  ; Prevent multiple instances

; Configuration
TESSERACT_ENV_VAR := "TESSERACT_PATH"
IMAGEMAGICK_ENV_VAR := "MAGICKBUBBSY"
TEMP_DIR := A_Temp "\ClipboardOCR\"
HOTKEY := "^+V"

; Detect Tesseract Path
EnvGet, TESSERACT_PATH, %TESSERACT_ENV_VAR%
if (TESSERACT_PATH = "") {
    MsgBox, Tesseract not found in environment variables. Please locate it manually.
    FileSelectFile, SelectedPath, 3,, Please locate the Tesseract executable (tesseract.exe)
    if (SelectedPath = "" || !(InStr(SelectedPath, "tesseract.exe"))) {
        MsgBox, Invalid file selected. Exiting.
        ExitApp
    }
    RunWait, % "setx " TESSERACT_ENV_VAR " """ SelectedPath """", , Hide
    TESSERACT_PATH := SelectedPath
    MsgBox, Tesseract path saved as %TESSERACT_ENV_VAR%. Continuing with the script...
}

; Detect ImageMagick Path
EnvGet, IMAGEMAGICK_PATH, %IMAGEMAGICK_ENV_VAR%
if (IMAGEMAGICK_PATH = "" || !FileExist(IMAGEMAGICK_PATH)) {
    MsgBox, ImageMagick (magick.exe) not found. Please locate it manually.
    FileSelectFile, SelectedPath, 3,, Please locate the ImageMagick executable (magick.exe)
    if (SelectedPath = "" || !(InStr(SelectedPath, "magick.exe"))) {
        MsgBox, Invalid file selected. Exiting.
        ExitApp
    }
    RunWait, % "setx " IMAGEMAGICK_ENV_VAR " """ SelectedPath """", , Hide
    IMAGEMAGICK_PATH := SelectedPath
    MsgBox, ImageMagick path saved as %IMAGEMAGICK_ENV_VAR%. Continuing with the script...
}

; Hotkey: Ctrl + Shift + V
Hotkey, %HOTKEY%, PerformOCR
return

PerformOCR:
{
    FileCreateDir, %TEMP_DIR%
    TempBmpFile := TEMP_DIR "clipboard_image.bmp"

    if !SaveClipboardImage(TempBmpFile) {
        MsgBox, No image found on the clipboard. Please copy an image first.
        return
    }

    if !RunImageMagick(IMAGEMAGICK_PATH, TempBmpFile) {
        MsgBox, Failed to preprocess the image.
        return
    }

    ProcessedFile := TEMP_DIR "clipboard_image_processed.bmp"
    TempResultFile := TEMP_DIR "ocr_result"
    if !RunTesseractOCR(TESSERACT_PATH, ProcessedFile, TempResultFile) {
        MsgBox, Tesseract failed to generate output. Check the input image.
        return
    }

    FileRead, OCRResult, % TempResultFile ".txt"
    Clipboard := OCRResult
    ClipWait, 1
    MsgBox, OCR result successfully copied to clipboard.

    ; Cleanup
    FileRemoveDir, %TEMP_DIR%, 1
}
return

; Utility: Get a registry value
GetRegistryValue(SubKey, ValueName) {
    RegRead, Value, HKEY_LOCAL_MACHINE, %SubKey%, %ValueName%
    return Value
}

; Save clipboard image to BMP file
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

; Utility: Run ImageMagick to preprocess the image
RunImageMagick(IMAGEMAGICK_PATH, InputFile) {
    ProcessedFile := RegExReplace(InputFile, "\.bmp$", "_processed.bmp", 1)
    RunWait, % IMAGEMAGICK_PATH . " """ . InputFile . """ -density 300 -strip -colorspace Gray -deskew 40% -resize 150% -contrast-stretch 5%x5% -sharpen 0x1.0 -despeckle """ . ProcessedFile """", , Hide
    return FileExist(ProcessedFile)
}

; Run Tesseract OCR with improved parameters
RunTesseractOCR(TESSERACT_PATH, InputFile, OutputFile) {
    RunWait, "%TESSERACT_PATH%" "%InputFile%" "%OutputFile%" -l eng --oem 3 --psm 6 --dpi 600 -c preserve_interword_spaces=1 -c tessedit_char_unrej_keep_perfect_wd=1, , Hide
    return FileExist(OutputFile ".txt")
}
