NLR,MLR,PLR,SII,SIRI,AISI = 3.4,0.38,158.0,917.0,2.23,606.56
MHR,HBRDW,MPVL,MPVP,RDWP,UHR = 16.1,0.96,6.84,4.3,0.05,16.3
HDL,URIC = 41.0,6.7

print("=== inversione degli indici per ricavare l'emocromo ===")
P = AISI/SIRI;            print(f"  Piastrine = AISI/SIRI            = {P:.2f}  x10^9/L")
M = AISI/SII;             print(f"  Monociti  = AISI/SII             = {M:.4f} x10^9/L")
L = M/MLR;                print(f"  Linfociti = monociti/MLR         = {L:.4f} x10^9/L")
N = SIRI*L/M;             print(f"  Neutrofili= SIRI*linf/monociti   = {N:.4f} x10^9/L")
MPV = MPVL*L;             print(f"  MPV       = (MPV/linf)*linfociti = {MPV:.3f} fL")

print("\n=== verifica incrociata dei valori ricostruiti ===")
def chk(nome, calc, pub):
    d = (calc-pub)/pub*100
    print(f"  {nome:26s} calcolato {calc:10.3f}  pubblicato {pub:>8}   scarto {d:+6.2f}%")
chk("NLR = N/L", N/L, NLR)
chk("MLR = M/L", M/L, MLR)
chk("PLR = P/L", P/L, PLR)
chk("SII = P*N/L", P*N/L, SII)
chk("SIRI = N*M/L", N*M/L, SIRI)
chk("AISI = N*M*P/L", N*M*P/L, AISI)
chk("Mono/HDL x1000 (per mille)", M/HDL*1000, MHR)
chk("MPV/Linfociti", MPV/L, MPVL)
chk("MPV/Piastrine x100", MPV/P*100, MPVP)
chk("Uricemia/HDL x100", URIC/HDL*100, UHR)

print("\n=== RDW ed emoglobina (sistema sottodeterminato) ===")
print("  RDW/Piastrine = 0,05 con 2 decimali -> RDW fra", round(0.045*P,2), "e", round(0.055*P,2))
for RDW in (13.6, 14.0, 14.5, 15.0):
    Hb = HBRDW*RDW
    print(f"   RDW {RDW:5.1f} -> Hb = {Hb:5.2f} g/dL | RDW/P = {RDW/P:.4f} (pubbl. 0,05) | Hb/RDW = {Hb/RDW:.3f}")
print("  scelta: RDW 14,5 e Hb 13,9 -> Hb/RDW = %.4f, RDW/P = %.4f" % (13.9/14.5, 14.5/P))

print("\n=== percentuali citate nell'articolo ===")
print("  neutrofili 69% / linfociti 20,5% -> NLR =", round(69/20.5,3), "(pubblicato 3,4)")
print("  monociti dedotti in percentuale     =", round(M/L*20.5,2), "%")

print("\n=== conteggio fuori riferimento (colonna 'Riferimento' dell'articolo) ===")
rows=[("PCR",0.3,0,0.5),("VES",14,1,25),("NLR",3.4,0.73,3.33),("MLR",0.38,0.12,0.38),
      ("PLR",158,63,209),("SII",917,131,901),("SIRI",2.23,0,0.68),("AISI",606.56,0,147.16),
      ("Monociti/HDL",16.1,0,6),("Emoglobina/RDW",0.96,1,99),("MPV/Linfociti",6.84,0,5.55),
      ("MPV/Piastrine",4.3,0,4.0),("RDW/Piastrine",0.05,0,0.07),("Omocisteina",14.6,3,15),
      ("Uricemia",6.7,3.5,7.2),("Uricemia/HDL",16.3,0,12.2)]
f=0; lim=0
for n,v,lo,hi in rows:
    out = v<lo or v>hi
    atlim = (v==hi or v==lo) and not out
    f+=out; lim+=atlim
    print(f"   {n:16s} {v:>8} in [{lo}, {hi}]  {'FUORI' if out else ('SUL LIMITE' if atlim else 'dentro')}")
print(f"  fuori = {f}, esattamente sul limite = {lim}, totale non-dentro = {f+lim}  <-- l'articolo dichiara 10")

print("\n=== coerenza con l'articolo 3 (correzione BRINDA della ferritina) ===")
print("  PCR = 0,3 mg/dL = 3 mg/L, sotto la soglia convenzionale di 5 mg/L")
print("  l'art.3 applicava alla ferritina una riduzione del 25% (fattore 0,7502)")
