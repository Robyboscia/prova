import math
W, H, WC, HIP, CALF, WRIST, AGE = 74.0, 174.0, 97.0, 100.0, 34.0, 18.0, 50
HDL, TG = 41.0, 148.0
Hm = H/100
print("=== BMI ===")
print("  74/1.74^2           =", round(W/Hm**2, 4), "  <-- referto dichiara 24,09")
print("  peso che darebbe 24,09 a 174 cm =", round(24.09*Hm**2, 2), "kg")
print("  altezza che darebbe 24,09 a 74 kg =", round(100*math.sqrt(W/24.09), 1), "cm")
BMI_calc = W/Hm**2
BMI_ref = 24.09

print("\n=== morfotipo / pesi ideali ===")
print("  altezza/polso       =", round(H/WRIST,2), "(referto 9,7)")
print("  Lorenz H-100-(H-150)/4 =", round(H-100-(H-150)/4,2), "(referto 68,0)")
print("  Creff (H-100+eta/10)*0,9 =", round((H-100+AGE/10)*0.9,2), "(referto 71,1)")

print("\n=== forma ===")
print("  WHtR = WC/H         =", round(WC/H,4), "(referto 0,56)")
print("  WHR  = WC/HIP       =", round(WC/HIP,4), "(referto 0,97)")
ci = (WC/100)/(0.109*math.sqrt(W/Hm))
print("  Conicity            =", round(ci,4), "(referto 1,36)")

print("\n=== adiposita ===")
for name,bmi in (("BMI calc",BMI_calc),("BMI referto",BMI_ref)):
    d = 1.20*bmi + 0.23*AGE - 10.8*1 - 5.4
    print(f"  Deurenberg %BF ({name}) =", round(d,2), "(referto 23)")
print("  RFM = 64-20*(H/WC)  =", round(64-20*(H/WC),2))
bai = HIP/(Hm**1.5) - 18
print("  BAI = HIP/H^1.5 -18 =", round(bai,3), "(referto 25)")
bri = 364.2 - 365.5*math.sqrt(1-(((WC/100)/(2*math.pi))/(0.5*Hm))**2)
print("  BRI                 =", round(bri,4), "(referto 4,50)")
avi = (2*WC**2 + 0.7*(WC-HIP)**2)/1000
print("  AVI standard        =", round(avi,3), "(referto 11,89)  <-- NON coincide")
print("  AVI/1000 su metri?  =", round((2*(WC/100)**2+0.7*((WC-HIP)/100)**2),4))

print("\n=== ibridi (conversioni SI) ===")
TGm = TG/88.57; HDLm = HDL/38.67
print("  TG mmol/L =", round(TGm,4), " HDL mmol/L =", round(HDLm,4))
print("  LAP = (WC-65)*TG_mmol =", round((WC-65)*TGm,3), "(referto 53,5)")
print("  CMI = WHtR*(TG/HDL) mg/dL =", round((WC/H)*(TG/HDL),4))
print("  CMI = WHtR*(TG/HDL) mmol  =", round((WC/H)*(TGm/HDLm),4), "(referto 0,88)")
for name,bmi in (("BMI calc",BMI_calc),("BMI referto",BMI_ref)):
    vai = (WC/(39.68+1.88*bmi))*(TGm/1.03)*(1.31/HDLm)
    print(f"  VAI ({name})       =", round(vai,4), "(referto 2,29)")

print("\n=== ABSI ===")
for name,bmi in (("BMI calc",BMI_calc),("BMI referto",BMI_ref)):
    absi = (WC/100)/((bmi**(2/3))*(Hm**0.5))
    print(f"  ABSI ({name})      =", round(absi,6))
print("  z-score referto 1,52 (richiede tabella NHANES eta/sesso, non pubblicata)")

print("\n=== polpaccio ===")
print("  CC grezza =", CALF, " referto corretto 33,5  soglia 34,0")
for k in (0.1,0.12,0.2):
    print(f"   CC - {k}*(BMIcalc-20) =", round(CALF-k*(BMI_calc-20),3),
          f"| BMIref =", round(CALF-k*(BMI_ref-20),3))
print("  EWGSOP2/Gonzalez categoriale (BMI 18,5-24,9 -> correzione 0) =", CALF)

print("\n=== conteggio fuori soglia (cut-off dell'articolo) ===")
vals = [("WHtR",WC/H,0.52,"alto"),("WHR",WC/HIP,0.89,"alto"),("Conicity",ci,1.25,"alto"),
        ("%BF",23,25,"alto"),("BAI",bai,23,"alto"),("BRI",bri,4.71,"alto"),
        ("AVI",11.89,24.5,"alto"),("LAP",(WC-65)*TGm,26.7,"alto"),
        ("CMI",(WC/H)*(TGm/HDLm),0.39,"alto"),("VAI",2.29,1.92,"alto"),
        ("ABSI z",1.52,0.228,"alto"),("CC corretta",33.5,34.0,"basso")]
fuori=0
for n,v,c,d in vals:
    out = v>c if d=="alto" else v<c
    fuori += out
    print(f"   {n:12s} {round(v,3):>8} vs {c:>6}  {'FUORI' if out else 'dentro'}")
print("  TOTALE FUORI SOGLIA =", fuori, "su 12   <-- l'articolo dichiara 6")
