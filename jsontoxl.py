import openpyxl as op
import json
from tkinter import *
w=Tk()
def callback():
    global data
    data=txt.get()
    a=json.loads(data)
    wb=op.Workbook()
    s=wb.active
    s['A1']='GSTIN'
    s['B1']=a['gstin']
    s['A2']="Year"
    s['B2']=a['fp'][2:6]
    s['A3']='month'
    s['B3']=a['fp'][0:2]
    s['A4']='GSTIN'
    s['B4']='inv no'
    s['C4']='invdte'
    s['D4']='inv value'
    s['E4']='pos'
    s['F4']='revcharg'
    s['G4']='taxable value'
    s['H4']='rate'
    s['I4']='CGST'
    s['J4']='SGST'
    s['K4']='Cess'
    j=5
    s.title="b2b"
    for i in a['b2b']:
        for k in i['inv'][0]['itms']:
            s['A'+str(j)]=i['ctin']
            s['B'+str(j)]=i['inv'][0]['inum']
            s['C'+str(j)]=i['inv'][0]['idt']
            s['D'+str(j)]=i['inv'][0]['val']
            s['E'+str(j)]=i['inv'][0]['pos']
            s['F'+str(j)]=i['inv'][0]['rchrg']
            v='F'
            c=1
            s[chr(ord(v)+c)+str(j)]=k['itm_det']['txval']
            c=c+1
            s[chr(ord(v)+c)+str(j)]=k['itm_det']['rt']
            c=c+1
            s[chr(ord(v)+c)+str(j)]=k['itm_det']['camt']
            c=c+1
            s[chr(ord(v)+c)+str(j)]=k['itm_det']['samt']
            c=c+1
            s[chr(ord(v)+c)+str(j)]=k['itm_det']['csamt']
            j=j+1
    wb.save("sample.xlsx")
    print("done")
    w.destroy()
w.title("Application")
w.geometry('200x80')
global txt
txt=Entry(w,width=200)
txt.grid(row=0,column=0)
txt.focus_set()
b=Button(w,text="submit",command=callback)
b.place(relx=0.5,rely=0.5,anchor=CENTER)

w.mainloop()


