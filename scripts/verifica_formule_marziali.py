import math
FE, FEadj, TRF, TIBC, UIBC = 13.6, 17.1, 39.2, 78.4, 64.8
SAT, SATadj, FER, FERadj, DEP, RATIO = 17.3, 21.8, 160.9, 120.7, 724.0, 1.63

print("=== relazioni interne al pannello ===")
print("  TIBC = 2 x transferrina      =", round(2*TRF,2), "(referto 78,4)")
print("  UIBC = TIBC - sideremia      =", round(TIBC-FE,2), "(referto 64,8)")
print("  Sat  = sideremia/TIBC*100    =", round(FE/TIBC*100,3), "(referto 17,3)")
print("  Sat aggiustata = Fe_adj/TIBC =", round(FEadj/TIBC*100,3), "(referto 21,8)")
print("  Depositi = ferritina * 4,5   =", round(FER*4.5,2), "(referto 724)")
print("     -> 4,5 = 4500 atomi di Fe per molecola di ferritina, /1000 per pM->nM")
print("  Depositi da ferritina AGGIUSTATA =", round(FERadj*4.5,2), "(dentro 301,5-706,5)")

print("\n=== fattori di aggiustamento BRINDA dedotti ===")
print("  ferritina: 120,7/160,9      =", round(FERadj/FER,4))
print("  sideremia: 17,1/13,6        =", round(FEadj/FE,4))
print("  (i marcatori d'infiammazione non sono pubblicati in questo articolo)")

print("\n=== transferrina / log(ferritina) ===")
for mw_f in (440000, 450000, 474000):
    fer_ug = FER*1e-12*mw_f*1e6   # pM -> g/L -> ug/L
    for mw_t in (76500, 79570, 80000):
        trf_gl = TRF*1e-6*mw_t
        print(f"  MWfer={mw_f} MWtrf={mw_t}: ferritina={fer_ug:.1f} ug/L, "
              f"transferrina={trf_gl:.3f} g/L, rapporto={trf_gl/math.log10(fer_ug):.3f}")
print("  referto 1,63")

print("\n=== conversioni verso unita convenzionali ===")
print("  sideremia  13,6 uM * 5,585  =", round(FE*5.585,1), "ug/dL")
print("  TIBC       78,4 uM * 5,585  =", round(TIBC*5.585,1), "ug/dL")
print("  UIBC       64,8 uM * 5,585  =", round(UIBC*5.585,1), "ug/dL")
print("  transferrina 39,2 uM*7,957  =", round(TRF*7.957,1), "mg/dL")
print("  ferritina 160,9 pM * 0,45   =", round(FER*0.45,1), "ug/L (ng/mL)")
print("  ferritina agg 120,7 pM*0,45 =", round(FERadj*0.45,1), "ug/L")
print("  soglia ottimale 283,1 pM    =", round(283.1*0.45,1), "ug/L")

print("\n=== conteggio fuori finestra ottimale ===")
rows = [("Sideremia",FE,13,27),("Sideremia aggiustata",FEadj,15,23),("Transferrina",TRF,25.1,50.3),
        ("TIBC",TIBC,45,76),("UIBC",UIBC,22,61),("Saturazione transferrina",SAT,20,48),
        ("Saturazione aggiustata",SATadj,24,35),("Ferritina",FER,54,854),
        ("Ferritina aggiustata",FERadj,283.1,99999),("Depositi di ferro",DEP,301.5,706.5),
        ("Transferrina/log(ferritina)",RATIO,0,1.70)]
fuori=0
for n,v,lo,hi in rows:
    out = v<lo or v>hi
    fuori += out
    print(f"   {n:30s} {v:>8} in [{lo}, {hi}]  {'FUORI' if out else 'dentro'}")
print("  TOTALE FUORI =", fuori, "  <-- l'articolo dichiara 5")
