dim wshShell
dim pDir

set wshShell = createObject("WScript.Shell") 
pDir = session.property("CustomActionData")
set wshEnv = wshShell.environment("SYSTEM") 

newPath = ""
s = split(wshEnv("Path"), ";")
for each e in s
    if e <> pDir & "\bin\" and e <> pDir & "\lib\" then
        newPath = newPath & e & ";"
    end if
next

wshEnv("Path") = newPath
wshEnv.remove("DUBBEL_LIBPATH")