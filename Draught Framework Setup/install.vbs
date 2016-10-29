dim wshShell
dim pDir
dim isPresent1
dim isPresent2

set wshShell = createObject("WScript.Shell")
pDir = session.property("CustomActionData")
set wshEnv = wshShell.environment("SYSTEM")

wshEnv("DUBBEL_LIBPATH") = pDir & "\lib\"

isPresent1 = false
isPresent2 = false
s = split(wshEnv("Path"), ";")
for each e in s
    if e = pDir & "\bin\" then
        isPresent1 = true
        exit for
    elseif e = pDir & "\lib\" then
        isPresent2 = true
        exit for
    end if
next

if not isPresent1 then
    wshEnv("Path") = wshEnv("Path") & ";" & pDir & "\bin\"
end if
if not isPresent2 then
    wshEnv("Path") = wshEnv("Path") & ";" & pDir & "\lib\"
end if
