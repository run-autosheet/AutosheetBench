on run argv
  set wbName to item 1 of argv
  set holdSeconds to (item 2 of argv) as real
  set stepRows to (item 3 of argv) as integer
  set maxSteps to (item 4 of argv) as integer
  set scrollDelay to (item 5 of argv) as real
  set sheetNames to items 6 thru -1 of argv
  set toured to {}
  tell application "Microsoft Excel"
    set wb to workbook wbName
    set startSheet to name of active sheet
    repeat with sheetName in sheetNames
      try
        set ws to worksheet (sheetName as text) of wb
        activate object ws
        set scroll row of active window to 1
        set scroll column of active window to 1
        delay holdSeconds
        set rowCount to count of rows of used range of ws
        set topRow to 1
        set steps to 0
        repeat while (topRow + stepRows) <= rowCount and steps < maxSteps
          set topRow to topRow + stepRows
          set scroll row of active window to topRow
          delay scrollDelay
          set steps to steps + 1
        end repeat
        set end of toured to ((sheetName as text) & ":" & rowCount & ":" & steps)
      on error errorMessage
        set end of toured to ((sheetName as text) & ":error:" & errorMessage)
      end try
    end repeat
    try
      activate object worksheet startSheet of wb
      set scroll row of active window to 1
      set scroll column of active window to 1
    end try
  end tell
  set AppleScript's text item delimiters to "|"
  return toured as text
end run
