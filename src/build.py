# Сборка: python3 src/build.py  → index.html, director.html и exec.html
import os
d=os.path.dirname(os.path.abspath(__file__));r=os.path.dirname(d)
css=open(os.path.join(d,'base.css')).read();logo=open(os.path.join(d,'logo.txt')).read()
for tpl,out in [('app.tpl.html','index.html'),('director.tpl.html','director.html'),('exec.tpl.html','exec.html')]:
    t=open(os.path.join(d,tpl)).read().replace('{{CSS}}',css).replace('{{LOGO}}',logo)
    open(os.path.join(r,out),'w').write(t)
print('ok')
