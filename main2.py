# full_script.py
import pandas as pd
import xlwings as xw

# -----------------------------
# Read transactions from Excel
# -----------------------------
df = pd.read_excel("Transaksjoner.xlsx", sheet_name="Transaksjoner")


# -----------------------------
# Helper class to categorize data
# -----------------------------
class Categories:
    def __init__(self, kidInn="¤#", kidUt="¤#", numrefInn="irrelevant", numrefUt="irrelevant",Undertype="Irrelevant"):
        self.kidInn = kidInn
        self.kidUt = kidUt
        self.numrefInn = numrefInn
        self.numrefUt = numrefUt
        self.df = df
        self.Undertype = Undertype

        allKids = [kidInn, kidUt]
        self.allKids = "|".join(map(str, allKids))

        # -----------------------------
        # All
        # -----------------------------
        self.all = df[
            df["Melding/KID/Fakt.nr"].str.contains(self.allKids, case=False, na=False)
            | df["Numref"].astype(str).str.contains(self.numrefInn, case=False, na=False)
            | df["Numref"].astype(str).str.contains(self.numrefUt, case=False, na=False)
            | df["Undertype"].astype(str).str.contains(self.Undertype, case=False, na=False)
        ].copy()
        self.all["utført dato"] = pd.to_datetime(self.all["Utført dato"], format="%d.%m.%Y", errors="coerce")
        self.all = self.all.sort_values(by="Utført dato", ascending=False)

        # -----------------------------
        # Inn
        # -----------------------------
        self.inn = df[
            df["Melding/KID/Fakt.nr"].str.contains(self.kidInn, case=False, na=False)
            | df["Numref"].astype(str).str.contains(self.numrefInn, case=False, na=False)
            | df["Undertype"].astype(str).str.contains(self.Undertype, case=False,na=False)
        ].copy()
        self.inn["utført dato"] = pd.to_datetime(self.inn["Utført dato"], format="%d.%m.%Y", errors="coerce")
        self.inn = self.inn.sort_values(by="Utført dato", ascending=False)

        # -----------------------------
        # Ut
        # -----------------------------
        self.ut = df[
            df["Melding/KID/Fakt.nr"].str.contains(self.kidUt, case=False, na=False)
            | df["Numref"].astype(str).str.contains(self.numrefUt, case=False, na=False)
        ].copy()
        self.ut["utført dato"] = pd.to_datetime(self.ut["Utført dato"], format="%d.%m.%Y", errors="coerce")
        self.ut = self.ut.sort_values(by="Utført dato", ascending=False)


# -----------------------------
# Class for remaining transactions
# -----------------------------
class Remaining:
    def __init__(self, inn, ut):
        self.inn = inn
        self.ut = ut


# -----------------------------
# Initialize categories
# -----------------------------
Nattcup = Categories("26180|standinntekt", "nattcup")
Kiosk = Categories("126635|970538", "kiosk", numrefUt="175694291|62748|172399792")
Milkshake = Categories("625974", "milks", numrefUt="96985710000")
Barrista = Categories("834150", "baris|barris", numrefUt="96985720000|96985730000")
Isruss = Categories("625973", "isrus")
Basar = Categories("625975", "basar")
Drottningborgrussen = Categories("115851", "drottningborgrussen")
Terminal_Nets = Categories("769821|5351127", "Terminal1")
Kjøregodtgjørelse = Categories("ikke relevant", "kjør")
Pant = Categories("pant")
Misjonsløp = Categories("¤¤kidnummer uvist", "løputgift|iphone|soundbox|airpods|multimedia",Undertype="OCR")
Skole = Categories(numrefInn="33753760001", numrefUt="41418400000|79708430000")
Bokbind = Categories(numrefInn="702660")
Måneskinstur = Categories("625976", "månes")
Krympefest = Categories("33392", "krympef")
Premier_leaugue = Categories("723903", "Premier lea|borgenpl")
Redaksjonen = Categories("MÅ FIKSES", "redaksjonen")


# -----------------------------
# Compute remaining transactions
# -----------------------------
df_filtered = pd.concat([
    Kiosk.all, Isruss.all, Milkshake.all, Pant.all, Basar.all, Terminal_Nets.all,
    Drottningborgrussen.all, Måneskinstur.all, Bokbind.all, Nattcup.all, Barrista.all,
    Skole.all, Misjonsløp.all, Kjøregodtgjørelse.all, Krympefest.all, Premier_leaugue.all, Redaksjonen.all
]).drop_duplicates()


df_remaining = df.merge(df_filtered, how='outer', indicator=True)
df_remaining = df_remaining[df_remaining['_merge'] == 'left_only'].drop(columns=['_merge'])
dfRemainingInn = df_remaining[df_remaining["Beløp inn"].notna()]
dfRemainingUt = df_remaining[df_remaining["Beløp inn"].isna()]
dfRemaining = Remaining(dfRemainingInn, dfRemainingUt)


# -----------------------------
# Function to write to Excel safely using xlwings
# -----------------------------
def printxl(transactions, sheetname, start_cell, filename="Regnskap/Russens Regnskap Helautomatisert.xlsx", headers=True):
    # Open Excel app
    app = xw.App(visible=False)  # Change to True if you want to see Excel
    app.display_alerts = False   # Avoid popups

    # Try to open workbook, otherwise create new
    try:
        wb = xw.Book(filename)
    except FileNotFoundError:
        wb = xw.Book()
        wb.save(filename)

    # Get or create sheet
    if sheetname in [s.name for s in wb.sheets]:
        ws = wb.sheets[sheetname]
    else:
        ws = wb.sheets.add(sheetname)

    start_row = ws.range(start_cell).row
    start_col = ws.range(start_cell).column

    # Write headers
    if headers:
        for j, col in enumerate(transactions.columns):
            ws.cells(start_row, start_col + j).value = col

    # Write data
    for i, row in enumerate(transactions.values):
        for j, val in enumerate(row):
            ws.cells(start_row + 1 + i, start_col + j).value = val

    wb.save()
    wb.close()
    app.quit()  # Close Excel app
    print(f"✅ Data written to '{filename}' → sheet '{sheetname}' starting at {start_cell}")



# -----------------------------
# Function to print Inn/Ut for a category
# -----------------------------
def printUtInn(Kategori, sheet, kordUt="C12", kordInn="J12"):
    if kordUt != "N":
        printxl(Kategori.ut[["Utført dato", "Beløp ut", "Mottakernavn", "Numref", "Melding/KID/Fakt.nr"]], sheet, kordUt)
    if kordInn != "N":
        printxl(Kategori.inn[["Utført dato", "Beløp inn", "Mottakernavn", "Numref", "Melding/KID/Fakt.nr"]], sheet, kordInn)



# -----------------------------
# Write all categories to Excel
# -----------------------------

printUtInn(Kiosk,"Vippsutskrifter Python","N","A80")
printUtInn(Kiosk,"Inn Kiosk Python", "A85","N")
printUtInn(Terminal_Nets,"Terminalutskrifter Python","N","A80")

printUtInn(Isruss, "Isruss")
printUtInn(Måneskinstur, "Måneskinstur")
printUtInn(Pant, "Panteruss", "J12", "C12")
printUtInn(Milkshake, "Milkshake")
printUtInn(Barrista, "Barrista")
printUtInn(Basar, "Basar")
printUtInn(Nattcup, "Nattcup")
printUtInn(Kjøregodtgjørelse, "Kjøregodtgjørelse", "R3", "N")
printUtInn(Krympefest, "Krympefest")
printUtInn(Premier_leaugue, "Premiere Leaugue")
printUtInn(Redaksjonen, "Redaksjonen")

printUtInn(dfRemaining, "Master", "AD5", "X5")
printUtInn(Drottningborgrussen,"Master","N","Q5")


# Individual category prints
printxl(Misjonsløp.ut[["Beløp ut", "Mottakernavn", "Til konto", "Numref", "Utført dato", "Melding/KID/Fakt.nr"]], "Misjonsløp", "B10")
printxl(Misjonsløp.inn[["Beløp inn", "Mottakernavn", "Til konto", "Numref", "Utført dato", "Melding/KID/Fakt.nr"]], "Misjonsløp", "I10")
printxl(Skole.inn[["Beløp inn"]], "Master", "E4", headers=False)
printxl(Skole.ut[["Beløp ut"]], "Master", "D2", headers=False)
