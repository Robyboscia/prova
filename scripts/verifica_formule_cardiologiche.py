SBP, DBP, HR, SV, BSA, AGE = 128, 84, 82, 66, 1.9, 50
PP = SBP - DBP
print("PP =", PP, "(referto 44)")

print("\n--- MAP: quale formula da' 101? ---")
print("  (SBP+2*DBP)/3        =", round((SBP+2*DBP)/3, 2))
print("  DBP + PP/3           =", round(DBP+PP/3, 2))
print("  DBP + 0.4*PP         =", round(DBP+0.4*PP, 2))
print("  DBP + 0.412*PP       =", round(DBP+0.412*PP, 2))

print("\n--- FC max ---")
print("  220 - eta            =", 220-AGE)
print("  Tanaka 208-0.7*eta   =", 208-0.7*AGE)
print("  referto (logistico)  = 175")
print("  banda 90-110% di 175 =", 0.9*175, "-", 1.1*175, "(referto 158-192)")

print("\n--- derivati ---")
print("  Riserva = 175-82     =", 175-HR)
print("  MSI = HR/MAP         =", round(HR/101, 3), "(referto 0,81)")
print("  MSI errato = HR/SBP  =", round(HR/SBP, 3), "(articolo cita 0,64)")
print("  RPP = HR*SBP         =", HR*SBP)

print("\n--- pompa ---")
CO = SV*HR/1000
CI = CO/BSA
print("  CO = SV*HR           =", CO, "L/min")
print("  CI = CO/BSA          =", round(CI,3), "(referto 2,8)")
# potenza: CI[L/min/m2] * MAP[mmHg] * (1/60000 m3/s) * 133.322 Pa = mW/m2
K = (1/60000)*133.322*1000
print("  costante CI*MAP->mW  =", round(K,4))
print("  LCWI = CI*MAP*2.222  =", round(CI*101*K,1), "(referto 627)")
print("  Rigidita = PP/MAP    =", round(PP/101,3), "(referto 0,44)")
print("  Rigidita = PP/SBP    =", round(PP/SBP,3))
for cvp in (0,3,5):
    print(f"  SVRI 80*(MAP-{cvp})/CI  =", round(80*(101-cvp)/CI), "dyn*s*cm-5*m2 (referto 2,7k)")

print("\n--- zone come % di FCmax 175 ---")
for lim in (87,104,105,121,122,139,140,156,157,175):
    print(f"   {lim} bpm = {round(100*lim/175,1)}%")

print("\n--- battiti in eccesso ---")
print("  (82-76)*60*24*365    =", (82-76)*60*24*365, "(articolo 3.153.600)")
print("\n--- RR frequenza a riposo (meta-analisi 1,09 per +10bpm) ---")
for ref in (65, 76):
    print(f"  vs {ref} bpm: RR = 1.09^(({HR}-{ref})/10) =", round(1.09**((HR-ref)/10),3))
