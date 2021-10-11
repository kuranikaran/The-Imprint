set shell =createobject("wscript.shell")
strtimes = inputbox("Enter a Automate number")
strtext=inputbox("Enter a Automate Keyword ")
if not isnumeric(strtimes)then
lol=msgbox("Please Entere again")
wscript.quit
end if
msgbox"Automated Keyword will be Printed ... "
wscript.sleep(3000)
for i=1 to strtimes
shell.sendkeys(strtext & "")
Shell.Sendkeys"{Enter}"
wscript.sleep(75)
next