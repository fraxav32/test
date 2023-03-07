set wscrpt = CreateObject("WScript.Shell")
Do

WScript.Sleep (5*1000)

wscrpt.SendKeys ("{SCROLLLOCK 2}")

Loop  