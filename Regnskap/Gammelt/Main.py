from contextlib import nullcontext

import pandas as pd
from Skrive import printxl

# Les Excel-filen
df = pd.read_excel("Transaksjoner.xlsx", sheet_name="Transaksjoner")  # eller bruk sheet_name=0 for første ark



class Categories:
    def __init__(self,kidInn="¤#",kidUt="¤#",numrefInn="irrelevant",numrefUt="irrelevant"):
        df = pd.read_excel("Transaksjoner.xlsx",sheet_name="Transaksjoner")  # eller bruk sheet_name=0 for første ark
        self.kidInn=kidInn
        self.kidUt=kidUt
        self.numrefInn = numrefInn
        self.numrefUt = numrefUt
        self.df = df

        #Fikse kid nummere
        allKids = [kidInn,kidUt]
        self.allKids = "|".join(map(str, allKids))

        """
        All
        """
        self.all = df[df["Melding/KID/Fakt.nr"].str.contains(self.allKids,case=False, na=False)
                      | df["Numref"].astype(str).str.contains(self.numrefInn,case=False, na=False)
                      | df["Numref"].astype(str).str.contains(self.numrefUt, case=False, na=False)
        ]
        self.all["utført dato"] = pd.to_datetime(self.all["Utført dato"], format="%d.%m.%Y", errors="coerce")
        self.all = self.all.sort_values(by="Utført dato", ascending=False)

        """
        Inn
        """
        self.inn = df[df["Melding/KID/Fakt.nr"].str.contains(self.kidInn,case=False, na=False)
                  | df["Numref"].astype(str).str.contains(self.numrefInn, case=False, na=False)
        ]
        self.inn["utført dato"] = pd.to_datetime(self.inn["Utført dato"], format="%d.%m.%Y", errors="coerce")
        self.inn = self.inn.sort_values(by="Utført dato", ascending=False)

        """
        Ut
        """
        self.ut = df[df["Melding/KID/Fakt.nr"].str.contains(self.kidUt,case=False, na=False)
                     |df["Numref"].astype(str).str.contains(self.numrefUt,case=False, na=False)
        ]
        self.ut["utført dato"] = pd.to_datetime(self.ut["Utført dato"], format="%d.%m.%Y", errors="coerce")
        self.ut = self.ut.sort_values(by="Utført dato",ascending=False)
        
class Remaining:
    def __init__(self,inn,ut):
        self.inn = inn
        self.ut = ut


Nattcup = Categories("26180|standinntekt","nattcup")
Kiosk = Categories("126635|970538","kiosk",numrefUt="175694291|62748")
Milkshake = Categories("625974","milks",numrefUt="96985710000")
Barrista = Categories("834150","baris|barris",numrefUt="96985720000|96985730000")
Isruss = Categories("625973","isrus")
Basar = Categories("625975","basar")
Drottningborgrussen = Categories("115851","drottningborgrussen")
Terminal_Nets = Categories("769821|5351127","Terminal1")
Kjøregodtgjørelse = Categories("ikke relevant","kjør")
Pant = Categories("pant")
Misjonsløp = Categories("¤¤kidnummer uvist","løputgift|iphone|soundbox|airpods|multimedia")
Skole = Categories(numrefInn="33753760001",numrefUt="41418400000|79708430000")
Bokbind = Categories(numrefInn="702660")
Måneskinstur = Categories("625976","månes")
Krympefest = Categories("33392","krympef")
Premier_leaugue = Categories("723903", "Premier lea|borgenpl")
Redaksjonen = Categories("MÅ FIKSES", "redaksjonen")



# Slå sammen alle filtrerte DataFrames til én
#df_filtered = pd.concat([Nattcup.all,Kiosk.all,Milkshake.all,Barrista.all,Isruss.all,Basar.all,Drottningborgrussen.all,Terminal_Nets.all,Kjøregodtgjørelse.all,Pant.all,Misjonsløp.all,Skole.all,Bokbind.all,Måneskin  stur.all])
#df_filtered = pd.concat([Drottningborgrussen.all,Måneskinstur.all,Bokbind.all,Nattcup.all,Barrista.all])
df_filtered = pd.concat([Kiosk.all,Isruss.all,Milkshake.all,Pant.all,Basar.all,Terminal_Nets.all,Drottningborgrussen.all,Måneskinstur.all,Bokbind.all,Nattcup.all,Barrista.all,Skole.all,Misjonsløp.all,Kjøregodtgjørelse.all,Krympefest.all,Premier_leaugue.all,Redaksjonen.all])

# Bruk merge med indikator, behold kun rader som ikke er i df_filtered
df_remaining = df.merge(df_filtered, how='outer', indicator=True)
df_remaining = df_remaining[df_remaining['_merge'] == 'left_only'].drop(columns=['_merge'])
dfRemainingInn = df_remaining[df_remaining["Beløp inn"].notna()]
dfRemainingUt = df_remaining[df_remaining["Beløp inn"].notna() == False]
dfRemaining = Remaining(dfRemainingInn,dfRemainingUt)




def printUtInn(Kategori,sheet,kordUt="C12",kordInn="J12"):
    if kordUt != "N":
        printxl(Kategori.ut[["Utført dato", "Beløp ut", "Mottakernavn", "Numref", "Melding/KID/Fakt.nr"]], str(sheet), kordUt)
    if kordInn != "N":
        printxl(Kategori.inn[["Utført dato", "Beløp inn", "Mottakernavn", "Numref", "Melding/KID/Fakt.nr"]],str(sheet), kordInn)


"""
printUtInn(Isruss,"safe Python")

printUtInn(Måneskinstur,"Måneskinstur")
printUtInn(Pant,"Panteruss","J12","C12")
printUtInn(Milkshake,"Milkshake")
printUtInn(Barrista,"Barrista")
printUtInn(Basar,"Basar")
printUtInn(Nattcup,"Nattcup")
printUtInn(Kjøregodtgjørelse,"Kjøregodtgjørelse","R3","N")
#printUtInn(Kiosk,"Vippsutskrifter Python","N","A80")
#printUtInn(Kiosk,"Inn Kiosk Python", "A85","N")
#printUtInn(Terminal_Nets,"Terminalutskrifter Python","N","A80")
#printUtInn(Drottningborgrussen,"Master","N","Q5")
printUtInn(Krympefest,"Krympefest")
printUtInn(Premier_leaugue,"Premiere Leaugue")
printUtInn(Redaksjonen,"Redaksjonen")

printUtInn(dfRemaining,"Master","AD5","X5")


printxl(Misjonsløp.ut[["Beløp ut", "Mottakernavn", "Til konto","Numref","Utført dato", "Melding/KID/Fakt.nr"]],"Misjonsløp", "B10")
printxl(Misjonsløp.inn[["Beløp inn", "Mottakernavn", "Til konto","Numref","Utført dato", "Melding/KID/Fakt.nr"]],"Misjonsløp", "I10")
printxl(Skole.inn[["Beløp inn"]],"Master", "E4",headers=False)
printxl(Skole.ut[["Beløp ut"]],"Master", "D2",headers=False)
"""