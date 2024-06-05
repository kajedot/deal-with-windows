#Requires AutoHotkey 2

ProcessedWindows := Map()  ; Initialize an empty object to store processed window IDs

; List of the apps to monitor
apps := ["OUTLOOK.EXE", "ms-teams.exe"]

id   := [] 

; Monitor for new windows every 1000 ms
SetTimer MonitorWindows, 1000

MonitorWindows(){
   
   For (app in apps) {
      id.Push( WinGetList("ahk_exe " app)* )
   }
   
   For (this_id in id) {
      ;MsgBox this_id, "Now processing window"
  
      if WinExist("ahk_id " this_id) {
         ;MsgBox this_id, "WinExists"
   
         if WinGetMinMax("ahk_id " this_id) != -1 {
            ;MsgBox this_id, "Win is not minimized"
   
            if !ProcessedWindows.Has(this_id) {
               ;MsgBox this_id, "Win was not processed" 
               
               WinMinimize "ahk_id " this_id
               
               ProcessedWindows[this_id] := true  ; Mark the window as processed
               
               Sleep 500 ; Wait for 0.5 seconds
               
               try {
                   WinRestore "ahk_id " this_id
               }
               catch {
               } 
            }
         }
      }
   }
   return
}
