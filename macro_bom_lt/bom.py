class Litze:
    def __init__(self):
        self.material_multipliers_strand: dict[str : dict[str:float]] = {
            "04": {"steel": 0.0044, "PE": 4, "GFK": 0.025, "schrumpfschlauch": 0.4, "AZBAN": 0.12},
            "05": {"steel": 0.0, "PE": 0.0, "GFK": 0.0, "schrumpfschlauch": 0.0, "AZBAN": 0.0},
            "07": {"steel": 0.0, "PE": 0.0, "GFK": 0.0, "schrumpfschlauch": 0.0, "AZBAN": 0.0},
            "08": {"steel": 0.0, "PE": 0.0, "GFK": 0.0, "schrumpfschlauch": 0.0, "AZBAN": 0.0},
            "10": {"steel": 0.0, "PE": 0.0, "GFK": 0.0, "schrumpfschlauch": 0.0, "AZBAN": 0.0},
        }

    def calculate_consumption(self, gesamt, lfr, stk, steel_mp, pe_mp, gfk_mp, schlauch_mp, azban_mp):
        # * run all methods
        L00VLI01 = self.L00VLI01(gesamt, steel_mp)
        PE = self.PE_00_00_3009(lfr, gesamt, pe_mp)
        GFK = self.GFK_klebeband_00_00_6070(lfr, gesamt, gfk_mp)
        Schrumpfschlauch = self.Schrumpfschlauch_010_7027(stk, gesamt, schlauch_mp)
        AZBAN5 = self.AZBAN5(stk, gesamt, azban_mp)

        return (L00VLI01, PE, GFK, Schrumpfschlauch, AZBAN5)

    def L00VLI01(self, gesamt, steel_mp):
        stahl_weight = steel_mp * 1000  # * convert to kg; 1m anchor per strand = 1,1kg
        return stahl_weight * gesamt

    def PE_00_00_3009(self, lfr, gesamt, pe_mp):
        return pe_mp * (lfr / gesamt)

    def GFK_klebeband_00_00_6070(self, lfr, gesamt, gfk_mp):
        return (gfk_mp * (lfr / gesamt)) * gesamt

    def Schrumpfschlauch_010_7027(self, stk, gesamt, schlauch_mp):
        return schlauch_mp * (stk / gesamt)

    def AZBAN5(self, stk, gesamt, azban_mp):
        return azban_mp * (stk / gesamt)
