import sys
html=open("/workspace/page.html","rb").read().decode("utf-8","ignore")
i=html.lower().find("<body")
print((html[i:i+1200]).replace("><","\n<"))
