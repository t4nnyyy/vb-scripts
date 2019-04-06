option explicit

dim david, zira

set david=createobject("sapi.spvoice")
set david.voice=david.getvoices.item(0)


set zira=createobject("sapi.spvoice")
set zira.voice=zira.getvoices.item(1)

zira.rate=3
zira.speak "David can i tell you a secret?"
david.speak "sure zira tell me, what is it?"

zira.rate=0
zira.volume=80
zira.speak "i'm Lesbian."

david.speak "hmm! i already know that. hahahaha"

zira.speak "hehehe"