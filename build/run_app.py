import sys, importlib
sys.path.insert(0,'build')
import appendix as A
import zones_b
A.cover(); A.contents()
for sec in A.SECTIONS:
    A.section(sec, list(sec["ids"]), A.ZONES)
    for i in sec["ids"]:
        A.card(A.ZONES[i], i, len(A.ZONES)); A.detail(A.ZONES[i], i, len(A.ZONES))
A.ending()
A.prs.save('Sokol_приложение_зоны.pptx')
print("slides:", len(A.prs.slides.__iter__.__self__._sldIdLst), "zones:", len(A.ZONES))
