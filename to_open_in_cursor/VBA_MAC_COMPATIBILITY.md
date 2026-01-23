# VBA Mac Compatibility Guide

This document summarizes critical differences and gotchas when writing VBA code that needs to work on both Windows and Mac (specifically MS Office 2026 for Mac).

---

## ⚠️ CRITICAL: Do NOT Use Conditional Compilation Around Declare Statements

**The #1 mistake that breaks Mac compatibility:**

```vba
' ❌ BAD - This BREAKS module compilation on Mac!
#If Mac Then
    ' Mac alternative
#Else
    Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As Long)
#End If

' ✅ GOOD - Leave Declare statements as-is!
Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As Long)
' The Declare only crashes when CALLED, not when it EXISTS
```

**Why:** Conditional compilation around module-level `Declare` statements or `Sub`/`Function` definitions breaks VBA's module compilation on Mac. The module won't load at all.

**Solution:** Leave all `Declare PtrSafe` statements unchanged. They're safe - they only crash when the function is actually CALLED, not when they exist in the module. Handle the Mac alternative at the CALL SITE instead.

---

## 1. Windows API Declarations - Call Site Handling

**Problem:** Any `Declare PtrSafe` statements for Windows DLLs (`kernel32`, `User32`, etc.) will cause an **immediate crash** on Mac - even if the code path isn't executed.

**Examples of problematic declarations:**
```vba
' These CRASH on Mac just by existing in the module:
Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Declare PtrSafe Function OpenClipboard Lib "User32" (ByVal hWnd As Long) As Long
Declare PtrSafe Function GetClipboardSequenceNumber Lib "User32" () As Long
```

**Solution:** Wrap ALL Windows API declarations in conditional compilation:

```vba
#If Mac Then
    ' Mac alternative implementation
    Public Sub Sleep(ByVal dwMilliseconds As Long)
        Dim endTime As Double
        endTime = Timer + (dwMilliseconds / 1000)
        Do While Timer < endTime
            DoEvents
        Loop
    End Sub
#Else
    ' Windows API declaration
    Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If
```

---

## 2. `IsEmpty()` Does NOT Work on Object Variables

**Problem:** `IsEmpty()` is designed for Variant variables, NOT object references. On Mac, using `IsEmpty()` on a Dictionary, Collection, or other object causes **Error 424: Object required**.

**BAD - Will crash on Mac:**
```vba
If IsEmpty(MyDictionary) Then  ' CRASH!
If Not IsEmpty(MyDictionary) And MyDictionary.Count > 0 Then  ' CRASH!
```

**GOOD - Mac compatible:**
```vba
' Method 1: Use Is Nothing
If MyDictionary Is Nothing Then

' Method 2: Defensive approach with error handling (most reliable)
Dim dictCount As Long
On Error Resume Next
dictCount = 0
dictCount = MyDictionary.Count
On Error GoTo 0

If dictCount > 0 Then
    ' Safe to use dictionary
End If
```

---

## 3. VBA `And`/`Or` Operators Do NOT Short-Circuit

**Problem:** Unlike most languages, VBA evaluates BOTH sides of `And`/`Or` expressions, even if the first condition determines the result.

**BAD - Will crash if MyDictionary is Nothing:**
```vba
' VBA evaluates BOTH sides - .Count fails if object is Nothing
If Not (MyDictionary Is Nothing) And MyDictionary.Count > 0 Then
```

**GOOD - Use nested If statements:**
```vba
If Not (MyDictionary Is Nothing) Then
    If MyDictionary.Count > 0 Then
        ' Safe to proceed
    End If
End If
```

**GOOD - Or use error handling:**
```vba
On Error Resume Next
dictCount = MyDictionary.Count
If Err.Number = 0 And dictCount > 0 Then
    ' Safe to proceed
End If
On Error GoTo 0
```

---

## 4. `WithEvents` May Have Limited Support on Mac

**Problem:** Event handling via `WithEvents` may not work reliably on Mac, causing crashes or silent failures.

**Solution:** Wrap event class initialization and event handlers in error handling:

```vba
' In initialization code:
On Error Resume Next
Set X = New CEventClass
Set X.PPTEvent = Application
On Error GoTo 0

' In event handler:
Private Sub PPTEvent_WindowSelectionChange(ByVal Sel As Selection)
    On Error GoTo EventErrorHandler
    ' ... event handling code ...
    Exit Sub
    
EventErrorHandler:
    ' Silently ignore errors on Mac
    On Error GoTo -1
End Sub
```

---

## 5. Timer/Performance Counter Differences

**Problem:** `QueryPerformanceCounter` and `QueryPerformanceFrequency` are Windows-only APIs.

**Solution:** Use VBA's built-in `Timer` function on Mac:

```vba
#If Mac Then
    Private m_StartTime As Double
    
    Public Sub StartCounter()
        m_StartTime = Timer
    End Sub
    
    Property Get TimeElapsed() As Double
        TimeElapsed = (Timer - m_StartTime) * 1000  ' milliseconds
    End Property
#Else
    ' Windows: Use QueryPerformanceCounter for higher precision
    Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" ...
#End If
```

---

## 6. Clipboard Handling Differences

**Problem:** Windows clipboard APIs (`OpenClipboard`, `CloseClipboard`, `GetClipboardSequenceNumber`, etc.) don't exist on Mac.

**Solution:** Use simpler approaches on Mac:

```vba
#If Mac Then
    ' Mac: Give clipboard time to process with DoEvents loop
    Dim macWaitEnd As Double
    macWaitEnd = Timer + 0.5  ' Wait up to 500ms
    Do While Timer < macWaitEnd
        DoEvents
    Loop
#Else
    ' Windows: Use clipboard APIs for precise control
    ' ... Windows clipboard code ...
#End If
```

---

## 7. Path Separator Differences

**Problem:** Windows uses `\` for paths, Mac uses `/` (and sometimes `:` in legacy HFS format).

**Solution:** Use `Application.PathSeparator` or check platform:

```vba
#If Mac Then
    Const PATH_SEP As String = "/"
#Else
    Const PATH_SEP As String = "\"
#End If

' Or dynamically:
Dim pathSep As String
pathSep = Application.PathSeparator
```

**Path conversion between formats:**
```vba
' Mac colon format to POSIX (for AppleScript)
posixPath = Replace(macPath, ":", "/")

' POSIX to Windows-style (for internal consistency)
winStylePath = Replace(posixPath, "/", Chr(92))

' Remove "Macintosh HD" prefix if present
cleanPath = Replace(cleanPath, "Macintosh HD", "")
```

**CRITICAL: Parsing filenames from paths on Mac:**
```vba
' BAD - Windows-only, causes freeze/error on Mac:
inputFileName = StrReverse(Left(reversedPath, InStr(1, reversedPath, "\") - 1))

' GOOD - Handle both / and \ separators:
#If Mac Then
    Dim backslashPos As Long, forwardslashPos As Long
    backslashPos = InStr(1, reversedPath, "\")
    forwardslashPos = InStr(1, reversedPath, "/")
    
    If backslashPos = 0 And forwardslashPos = 0 Then
        inputFileName = fullPath  'No separator - use whole path
    ElseIf backslashPos = 0 Then
        pathSeparatorPos = forwardslashPos
    ElseIf forwardslashPos = 0 Then
        pathSeparatorPos = backslashPos
    Else
        pathSeparatorPos = IIf(backslashPos < forwardslashPos, backslashPos, forwardslashPos)
    End If
    
    If pathSeparatorPos > 0 Then
        inputFileName = StrReverse(Left(reversedPath, pathSeparatorPos - 1))
    End If
#Else
    pathSeparatorPos = InStr(1, reversedPath, "\")
    inputFileName = StrReverse(Left(reversedPath, pathSeparatorPos - 1))
#End If
```

---

## 8. File Dialogs and Excel Access on Mac - AppleScript Workaround

**Problem:** `Application.FileDialog` and direct Excel automation don't work reliably on Mac. VBA's built-in file picker may crash or return incorrect paths.

**Solution:** Use an external AppleScript file called via `AppleScriptTask()`.

### Required Setup

1. **Create/obtain the AppleScript file:** `MyBoss-browse_files_on_mac.scpt`

2. **Place it in the correct location:**
   ```
   ~/Library/Application Scripts/com.microsoft.PowerPoint/
   ```
   (For Excel add-ins, use `com.microsoft.Excel` instead)

3. **The AppleScript should implement these handlers:**
   - `openDialog` - Show file picker dialog
   - `openExcelGetSheets` - Open Excel file and return sheet names
   - `checkExcelExists` - Verify Excel file exists
   - `checkSheetExists` - Verify sheet exists in workbook
   - `openExcelGetSheetData` - Read data from Excel sheet

### VBA Code Pattern

```vba
Function browseFile() As String
    #If Mac Then
        seedPath = Replace(ActivePresentation.Path, "/", ":")
    
        If CheckAppleScriptTaskPowerPointScriptFile("MyBoss-browse_files_on_mac.scpt") = True Then
            filePath = CStr(AppleScriptTask("MyBoss-browse_files_on_mac.scpt", "openDialog", seedPath))
            
            ' Convert Mac path to Windows-style for internal use
            browseFile = Replace(filePath, ":", Chr(92))
            browseFile = Replace(browseFile, "Macintosh HD", "")
        Else
            MsgBox "AppleScript file not found in correct location"
            browseFile = ""
        End If
    #Else
        ' Windows: Use standard FileDialog
        Set diaFolder = Application.FileDialog(msoFileDialogFilePicker)
        diaFolder.AllowMultiSelect = False
        diaFolder.Show
        
        If diaFolder.SelectedItems.Count > 0 Then
            browseFile = diaFolder.SelectedItems.Item(1)
        Else
            browseFile = ""
        End If
    #End If
End Function

' Helper to verify AppleScript file exists
Function CheckAppleScriptTaskPowerPointScriptFile(ScriptFileName As String) As Boolean
    Dim AppleScriptTaskFolder As String
    Dim TestStr As String

    AppleScriptTaskFolder = MacScript("return POSIX path of (path to desktop folder) as string")
    AppleScriptTaskFolder = Replace(AppleScriptTaskFolder, "/Desktop", "") & _
        "Library/Application Scripts/com.microsoft.PowerPoint/"

    On Error Resume Next
    TestStr = Dir(AppleScriptTaskFolder & ScriptFileName, vbDirectory)
    On Error GoTo 0
    
    CheckAppleScriptTaskPowerPointScriptFile = (TestStr <> vbNullString)
End Function
```

### Reading Excel Data on Mac

```vba
#If Mac Then
    ' Convert path for AppleScript (use forward slashes)
    Dim parameterString As String
    parameterString = Replace(inputFilePath, Chr(92), "/") & "|" & sheetName
    
    ' Get sheet names
    worksheetNames = Split(AppleScriptTask("MyBoss-browse_files_on_mac.scpt", _
                           "openExcelGetSheets", parameterString), ",")
    
    ' Get sheet data
    excelSheetData = CStr(AppleScriptTask("MyBoss-browse_files_on_mac.scpt", _
                          "openExcelGetSheetData", parameterString))
#Else
    ' Windows: Use direct Excel automation
    Set oXL = CreateObject("Excel.Application")
    Set oWB = oXL.Workbooks.Open(inputFilePath)
    ' ... direct Excel object manipulation ...
#End If
```

### Important Notes

- The `.scpt` file is a **compiled binary** AppleScript
- User must manually copy the `.scpt` file to the correct location
- The script folder requires user permission to access
- Error `-1728` from AppleScript typically means file/folder not found
- Reference: https://www.rondebruin.nl/mac/applescripttask.htm

---

## 9. Array Initialization Checks

**Problem:** Calling `UBound()` on an uninitialized array causes **Error 13: Type mismatch** on Mac.

**Solution:** Use error handling:

```vba
Sub getNumberOfItems(ByRef returnedVal)
    On Error Resume Next
    returnedVal = UBound(MyArray) + 1
    If Err.Number <> 0 Then
        returnedVal = 0
        Err.Clear
    End If
    On Error GoTo 0
End Sub
```

---

## 10. General Best Practices for Mac Compatibility

### Always Use Conditional Compilation for Platform-Specific Code
```vba
#If Mac Then
    ' Mac-specific code
#Else
    ' Windows-specific code
#End If
```

### Defensive Error Handling Pattern
```vba
Sub SafeFunction()
    On Error Resume Next
    
    ' Attempt operation
    result = SomeObject.SomeProperty
    
    If Err.Number <> 0 Then
        ' Handle failure gracefully
        result = DefaultValue
        Err.Clear
    End If
    
    On Error GoTo 0
End Sub
```

### Test Object Validity Before Access
```vba
' Don't assume objects are initialized
Dim objCount As Long
On Error Resume Next
objCount = MyObject.Count
If Err.Number <> 0 Or objCount = 0 Then
    ' Object not ready - return safe default
    Exit Sub
End If
On Error GoTo 0
```

---

## Common Error Codes Reference

| Error | Description | Common Mac Cause |
|-------|-------------|------------------|
| 424 | Object required | Using `IsEmpty()` on objects, accessing `Nothing` object |
| 13 | Type mismatch | `UBound()` on uninitialized array |
| 429 | ActiveX component can't create object | Windows-only COM objects |
| 453 | Specified DLL function not found | Windows API calls |

---

## Files Typically Needing Mac Fixes

1. **Modules with API declarations** - Wrap in `#If Mac Then`
2. **Timer/Performance classes** - Replace Windows APIs with `Timer`
3. **Clipboard handling code** - Simplify for Mac
4. **Event handler classes** - Add robust error handling
5. **UI callback functions** - Defensive checks for uninitialized objects/arrays
6. **File dialog/browsing code** - Use AppleScriptTask workaround
7. **Excel automation code** - Route through AppleScript on Mac

## Required External Files for Mac

| File | Location | Purpose |
|------|----------|---------|
| `MyBoss-browse_files_on_mac.scpt` | `~/Library/Application Scripts/com.microsoft.PowerPoint/` | File dialogs & Excel access |

**User Setup Required:** The AppleScript file must be manually copied to the location above. VBA cannot do this automatically due to sandboxing.

---

*Last updated: January 2026*
*Based on debugging VBA add-in for MS Office 2026 for Mac*
