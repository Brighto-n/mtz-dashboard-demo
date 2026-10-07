"""
Зведення файлів доставки з папки (напр. синхронізованої папки Nextcloud) в одну таблицю.

Запуск:  python zvedennia.py <папка_з_файлами> <dovidnyk_pidrozdiliv.csv> <папка_результату>

Результат:
  dostavky_flat.csv  — один рядок = одна доставка (для Superset, широкий формат)
  dostavky_long.csv  — один рядок = доставка × категорія (зручно для графіків по категоріях)
  pomylky.csv        — журнал проблем (порожні поля, невідомі позиції, некоректні числа)
  stan_zapysiv.csv   — службовий: коли кожен рядок уперше з'явився у файлах (НЕ видаляти)
Запускати щодня: таблиця щоразу перебудовується з усіх файлів (без дублікатів),
а дата «вперше побачено» для старих рядків зберігається.
Потрібно: Python 3 + openpyxl (pip install openpyxl). Інтернет не потрібен.
"""
import sys, os, csv, glob
from datetime import datetime, date
from openpyxl import load_workbook

CATS=[("Набір продуктів, кг","food_kg"),("Пайок, кг","ration_kg"),("Вода, кг","water_kg"),("ЗУ, кг","zu_kg"),
      ("ЗББ, кг","zbb_kg"),("АтППО, кг","atppo_kg"),("речове майно, кг","clothing_kg"),("ПММ,кг","fuel_kg"),
      ("засоби звʼязку/АКБ","comms_kg"),("Інше МТЗ, кг","other_kg")]

def norm(s): return " ".join(str(s).split()).lower() if s is not None else ""

def main(src,dov,out):
    os.makedirs(out,exist_ok=True)
    # довідник: шлях за (в/ч, підрозділ, позиція)
    path_by={}; names={}
    with open(dov,encoding="utf-8-sig") as f:
        rows=list(csv.DictReader(f))
    for r in rows: names[r["code"]]=r["name"]
    for r in rows:
        if r["level"]=="позиція":
            parts=r["path"].split("/")   # група/в/ч/підрозділ/позиція
            path_by[(norm(parts[1]),norm(parts[2]),norm(parts[3]))]=parts
    flat=[]; errs=[]; today=date.today().isoformat()
    st_fn=os.path.join(out,"stan_zapysiv.csv"); seen={}
    if os.path.exists(st_fn):
        for r in csv.DictReader(open(st_fn,encoding="utf-8-sig")): seen[r["key"]]=r["first_seen"]
    for fn in sorted(glob.glob(os.path.join(src,"**","*.xlsx"),recursive=True)):
        if os.path.basename(fn).startswith("~$"): continue
        ws=load_workbook(fn,data_only=True).worksheets[0]
        head=[norm(c.value) for c in ws[1]]
        def col(name):
            n=norm(name)
            return head.index(n) if n in head else None
        need=["військова частина","підрозділ","позиція","дата, час доставки"]+[c[0] for c in CATS]
        miss=[n for n in need if col(n) is None]
        if miss: errs.append([os.path.basename(fn),"",f"немає колонок: {', '.join(miss)}"]); continue
        meth=[i for i,h in enumerate(head) if h=="спосіб доставки"]
        for row in ws.iter_rows(min_row=2,values_only=True):
            vch,sub,pos=row[col("військова частина")],row[col("підрозділ")],row[col("позиція")]
            if not any([vch,sub,pos]): continue
            rn=f"{os.path.basename(fn)}"
            key=(norm(vch),norm(sub),norm(pos))
            p=path_by.get(key)
            if not p: errs.append([rn,f"{vch}/{sub}/{pos}","позиції немає в довіднику"]); continue
            dt=row[col("дата, час доставки")]
            if not isinstance(dt,datetime): errs.append([rn,f"{vch}/{sub}/{pos}",f"некоректна дата: {dt}"]); continue
            kg={}; bad=False
            for h,k in CATS:
                v=row[col(h)]
                try: kg[k]=float(v) if v not in (None,"") else 0.0
                except (TypeError,ValueError): bad=True; errs.append([rn,f"{vch}/{sub}/{pos}",f"не число у «{h}»: {v}"])
            if bad: continue
            g=lambda n: row[col(n)] if col(n) is not None else None
            ed=g("дата запису")
            if not isinstance(ed,datetime): errs.append([rn,f"{vch}/{sub}/{pos}","немає дати запису (рядок враховано)"]); ed=None
            key="|".join([norm(vch),norm(sub),norm(pos),dt.strftime("%Y-%m-%d %H:%M")])
            seen.setdefault(key,today)
            flat.append(dict(delivery_dt=dt.strftime("%Y-%m-%d %H:%M"),delivery_date=dt.strftime("%Y-%m-%d"),
                grp=names.get(p[0],p[0]),unit=p[1],subunit=p[2],position=p[3],
                path="/".join(p),personnel=g("кількість в/с на позиції") or 0,**kg,total_kg=round(sum(kg.values()),1),
                responsible=g("відповідальний за пакування") or "",
                method=row[meth[0]] if meth else "",vehicle=row[meth[1]] if len(meth)>1 else "",
                confirm_link=g("посилання на пітвердження доставки") or "",note=g("примітка") or "",
                entry_dt=ed.strftime("%Y-%m-%d %H:%M") if ed else "",first_seen=seen[key],
                source_file=os.path.basename(fn)))
    for i,r in enumerate(flat,1): r["delivery_id"]=i
    cols=["delivery_id"]+[k for k in flat[0] if k!="delivery_id"]
    with open(os.path.join(out,"dostavky_flat.csv"),"w",newline="",encoding="utf-8-sig") as f:
        w=csv.DictWriter(f,fieldnames=cols); w.writeheader(); w.writerows(flat)
    with open(os.path.join(out,"dostavky_long.csv"),"w",newline="",encoding="utf-8-sig") as f:
        w=csv.writer(f); w.writerow(["delivery_id","delivery_date","grp","unit","subunit","position","path","method","category","kg"])
        for r in flat:
            for h,k in CATS:
                if r[k]>0: w.writerow([r["delivery_id"],r["delivery_date"],r["grp"],r["unit"],r["subunit"],r["position"],r["path"],r["method"],h.replace(", кг","").replace(",кг",""),r[k]])
    with open(st_fn,"w",newline="",encoding="utf-8-sig") as f:
        w=csv.writer(f); w.writerow(["key","first_seen"]); w.writerows(sorted(seen.items()))
    with open(os.path.join(out,"pomylky.csv"),"w",newline="",encoding="utf-8-sig") as f:
        w=csv.writer(f); w.writerow(["файл","позиція","проблема"]); w.writerows(errs)
    print(f"Доставок: {len(flat)}, проблем: {len(errs)}")

if __name__=="__main__":
    if len(sys.argv)!=4: print(__doc__); sys.exit(1)
    main(*sys.argv[1:])
