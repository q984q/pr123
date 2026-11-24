import sys, os
sys.path.insert(0,'build')
import appendix as A
import zones_b
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

prs=A.prs; sld=A.sld; rect=A.rect; txt=A.txt; hair=A.hair
M=A.M; CW=A.CW
W=A.W; MUT=A.MUT; ACC=A.ACC; DIM=A.DIM; LINE=A.LINE; SUB=A.SUB; BG=A.BG; CARD=A.CARD
ZONES=A.ZONES

PHOTOS={
 "01":"assets3/z01_dc.jpg","02":"assets3/z02_robotics.jpg","03":"assets3/z03_aviation.jpg",
 "04":"assets3/z04_biomed.jpg","05":"assets3/z05_electronics.jpg","06":"assets3/z06_auto.jpg",
 "07":"assets3/z07_rail.jpg","08":"assets3/z08_station.jpg","09":"assets3/z09_truckhub.jpg",
}
PHOTOS2={"06":"assets3/z06_tire.jpg"}

def pic(s,path,x,y,w,h):
    if not path or not os.path.exists(path): return False
    s.shapes.add_picture(path,Inches(x),Inches(y),Inches(w),Inches(h)); return True

def short(t):
    t=t.strip()
    for sep in (". ",):
        i=t.find(sep)
        if 0<i<95: return t[:i+1]
    return t if len(t)<95 else t[:92].rsplit(' ',1)[0]+"…"

# ---------- 1. cover with photo ----------
def z_cover(z,idx,total):
    s=sld()
    ph=PHOTOS.get(z['n'])
    if ph:
        pic(s,ph,0,0,13.333,7.5)
        rect(s,0,0,7.15,7.5,BG)
    else:
        rect(s,0,0,13.333,7.5,RGBColor(0x12,0x16,0x1C))
        rect(s,7.15,0,0.012,7.5,RGBColor(0x22,0x29,0x33))
        txt(s,7.85,1.02,4.9,2.15,z['n'],118,True,RGBColor(0x1D,0x24,0x2E),sp=0.86)
        hair(s,3.30,7.85,4.35)
        txt(s,7.85,3.54,4.35,0.26,"ПЛОЩАДЬ",7.2,True,ACC,caps=True)
        txt(s,7.85,3.86,4.35,0.7,z['area'],34,True,RGBColor(0x66,0x71,0x7F),sp=0.98)
        txt(s,7.85,4.72,4.30,0.9,z['plots'].replace("\n"," "),8.4,False,RGBColor(0x4E,0x58,0x64),sp=1.3)
        hair(s,6.55,7.85,4.35)
    txt(s,M,0.72,6.2,0.26,f"ЗОНА {z['n']} ИЗ {total} · {z['sec']}",8,True,ACC,caps=True)
    txt(s,M,1.62,6.1,1.9,z['name'],33,True,W,sp=1.04)
    rect(s,M,3.72,0.62,0.02,ACC)
    txt(s,M,3.96,5.75,1.1,z['purpose'],10.5,False,SUB,sp=1.38)
    txt(s,M,6.28,2.6,0.24,"ПЛОЩАДЬ",7.4,True,ACC,caps=True)
    txt(s,M,6.52,2.4,0.4,z['area'],15,True,W)
    txt(s,M+2.55,6.52,3.4,0.55,z['plots'].replace("\n"," "),8,False,DIM,sp=1.25)

# ---------- 2. состав ----------
def z_comp(z,idx,total):
    s=sld()
    txt(s,M,0.74,9.6,0.3,f"{z['n']} · СОСТАВ",8,True,ACC,caps=True)
    txt(s,M,1.10,11.4,0.6,"Что входит в зону",27,True,W)
    hair(s,1.96)
    col=z['inside']
    for j,it in enumerate(col):
        y=2.20+j*0.74
        txt(s,M,y,0.55,0.3,f"0{j+1}",9,True,RGBColor(0x8A,0x6A,0x34))
        txt(s,M+0.62,y-0.02,5.9,0.62,it,10.6,False,SUB,sp=1.26)
    # right column: residents + why
    rect(s,8.10,2.10,4.51,1.95,CARD)
    txt(s,8.34,2.28,3.9,0.24,"КТО РЕЗИДЕНТ",7.2,True,ACC,caps=True)
    for j,it in enumerate(z['who']):
        txt(s,8.34,2.56+j*0.44,4.0,0.4,it,8.8,False,MUT,sp=1.2)
    txt(s,M,6.10,7.0,0.24,"ЗАЧЕМ ЭТО",7.2,True,ACC,caps=True)
    txt(s,M,6.34,7.2,0.75,z['why'],9,False,MUT,sp=1.30)

# ---------- 3. процесс ----------
def z_flow(z,idx,total):
    s=sld()
    ph=PHOTOS2.get(z['n']) or PHOTOS.get(z['n'])
    if ph:
        pic(s,ph,0,0,13.333,2.62)
        # затемняющая полоса у нижней кромки фото под подпись
        rect(s,0,2.34,13.333,0.28,RGBColor(0x0E,0x11,0x16))
    txt(s,M,2.92,9.0,0.3,f"{z['n']} · КАК ЭТО РАБОТАЕТ",8,True,ACC,caps=True)
    txt(s,M,3.28,11.2,0.62,z['ptitle'],26,True,W,sp=1.04)
    hair(s,4.22)
    for i,(h,d) in enumerate(z['process']):
        x=M+i*3.02
        txt(s,x,4.46,0.7,0.3,f"0{i+1}",10,True,ACC)
        txt(s,x,4.80,2.76,0.34,h,11,True,W)
        txt(s,x,5.20,2.76,1.05,d,8.8,False,MUT,sp=1.30)
        if i<3: rect(s,x+2.86,4.46,0.010,1.85,LINE)
    txt(s,M,6.86,CW,0.24,"Цифры, инженерия и ограничения зоны — на следующем слайде",7.4,False,DIM)

# ---------- 4. инженерия и границы ----------
def z_facts(z,idx,total):
    s=sld()
    txt(s,M,0.74,9.6,0.3,f"{z['n']} · ЦИФРЫ И ОГРАНИЧЕНИЯ",8,True,ACC,caps=True)
    txt(s,M,1.10,11.4,0.6,"Что известно точно, а что нет",27,True,W)
    hair(s,1.96)
    for i,(n,u) in enumerate(z['tech']):
        x=M+i*3.02
        txt(s,x,2.24,2.7,0.5,n,19,True,W,sp=1.0)
        txt(s,x,2.82,2.76,0.55,u,8,False,DIM,sp=1.22)
    hair(s,3.62)
    txt(s,M,3.82,2.4,0.24,"ИНЖЕНЕРИЯ",7.2,True,ACC,caps=True)
    yy=4.10
    for it in z['eng']:
        hh=0.30 if len(it)<=80 else 0.50
        txt(s,M,yy,5.85,hh,"—  "+it,8.6,False,MUT,sp=1.18); yy+=hh
    txt(s,6.95,3.82,2.6,0.24,"ЧЕГО НЕТ",7.2,True,ACC,caps=True)
    yy=4.10
    for it in z['limits']:
        hh=0.30 if len(it)<=80 else 0.50
        txt(s,6.95,yy,5.66,hh,"—  "+it,8.6,False,MUT,sp=1.18); yy+=hh
    rect(s,M,6.62,CW,0.52,RGBColor(0x1A,0x15,0x0E))
    rect(s,M,6.62,0.05,0.52,ACC)
    txt(s,M+0.24,6.72,6.9,0.34,"СТАТУС: "+z['status'],8.6,True,RGBColor(0xE8,0xC9,0x9A))
    txt(s,M+7.30,6.72,4.9,0.34,"СВЯЗИ: "+z['links'],8.0,False,DIM)

def sec_div(sec,ids,zs):
    s=sld()
    n=len(ids)
    txt(s,M,0.80,CW,0.3,"РАЗДЕЛ",8,True,ACC,caps=True)
    txt(s,M,1.24,9.4,0.95,sec["title"],40,True,W,sp=1.02)
    hair(s,2.72)
    txt(s,M,2.96,6.9,0.95,sec["lead"],10.5,False,MUT,sp=1.38)
    txt(s,M,4.56,2.0,0.24,"ЗОНЫ РАЗДЕЛА",7.2,True,ACC,caps=True)
    for j,i in enumerate(ids):
        z=zs[i]
        txt(s,M,4.84+j*0.36,7.8,0.32,f"{z['n']}   {z['name']}   —   {z['area']}",9,False,SUB,sp=1.1)
    txt(s,9.55,2.90,3.0,0.26,"ВСЕГО",7.2,True,ACC,caps=True)
    txt(s,9.55,3.20,3.0,0.7,sec["sum"],32,True,W,sp=0.98)
    txt(s,9.55,4.10,3.0,1.9,sec["note"],8.2,False,MUT,sp=1.30)

if __name__=="__main__":
    A.cover()
    A.contents()
    for sec in A.SECTIONS:
        ids=list(sec["ids"])
        sec_div(sec,ids,ZONES)
        for i in ids:
            z=ZONES[i]
            z_cover(z,i,len(ZONES)); z_comp(z,i,len(ZONES))
            z_flow(z,i,len(ZONES)); z_facts(z,i,len(ZONES))
    A.ending()
    prs.save('Sokol_приложение_зоны_v2.pptx')
    print("slides:",len(prs.slides.__iter__.__self__._sldIdLst),"zones:",len(ZONES))
