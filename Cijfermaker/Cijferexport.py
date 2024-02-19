import pandas as pd
import os
# voor gebruik installeer eerst, numpy, pandas, openpyxl, door: pip install "extensie", te doen

excel_bestand = "cijfers.xlsx"
if os.path.exists(excel_bestand):
    os.remove(excel_bestand)
    
norminput = str(input("Vul norm in: norm%."))
maxaantalinput = int(input("Vul max aantal punten in: punten."))

# Functie om cijfer te berekenen 
def bereken_cijfer(aantal_behaalde_punten, totaal_aantal_punten, procent_norm):
    if "%" in procent_norm: #verwijderen van het procent teken
        norm = int(procent_norm.replace("%", ""))
    else:
        norm = int(procent_norm)
    anorm = (norm - 50) * 2 # de norm omzetten naar hoeveel procent je nodigt hebt voor een 1
    a = (10 - 1) / (100 - int(anorm)) # de a berekenen van y=ax+b
    b = -(a*100) + 10 # de b berekenen van y=ax+b
    x = aantal_behaalde_punten / totaal_aantal_punten * 100 # de x berekenen, zodat je het cijfer kunt uitlezen dat overeekomt met het aantal procent dat je goed hebt
    if b > 0: # je kan niet lager hebben dan een 1
        cijfer = round(a*x-b, 1) # afronden op 1 decimaal
    else:
        cijfer = round(a*x+b, 1)
    return cijfer # terug geven van cijfer

# Lijst met aantal behaalde punten van 100 naar 0 met stapgrootte van -0.5
aantal_behaalde_punten_lijst = list(range(maxaantalinput * 2, 0, -1)) # Verdubbel het maxaantalinput zodat de stapgrootte van 0.5 wordt bereikt
aantal_behaalde_punten_lijst = [punt / 2 for punt in aantal_behaalde_punten_lijst] # Converteer de punten terug naar floats

# Het totaal aantal punten blijft hetzelfde
totaal_aantal_punten_lijst = [maxaantalinput] * len(aantal_behaalde_punten_lijst)

# Lijsten om cijfers op te slaan
cijfers = []

# Bereken cijfers voor elk paar aantal behaalde punten en totaal aantal punten
for aantal_behaalde_punten, totaal_aantal_punten in zip(aantal_behaalde_punten_lijst, totaal_aantal_punten_lijst):
    cijfer = bereken_cijfer(aantal_behaalde_punten, totaal_aantal_punten, norminput)
    if cijfer < 1:
        cijfer = 1
    cijfers.append(cijfer)

# Creëer een dataframe met de gegevens
data = {'Aantal behaalde punten': aantal_behaalde_punten_lijst,
        'Totaal aantal punten': totaal_aantal_punten_lijst,
        'Cijfer': cijfers}
df = pd.DataFrame(data)

# Bepaal het pad naar het Excel-bestand in dezelfde map als het script
script_dir = os.path.dirname(__file__)
excel_bestand = os.path.join(script_dir, "cijfers.xlsx")

# Exporteer het dataframe naar Excel
df.to_excel(excel_bestand, index=False)

print("Cijfers zijn geëxporteerd naar", excel_bestand)
