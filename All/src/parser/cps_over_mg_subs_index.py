class CpsOverMgSubsIndexHandler:
    def __init__(self, new_row):
        self.new_row = new_row

    def add_additional_fields(self):
        self.new_row['YEAR'] = self.calculate_year()
        # Index
        self.new_row['CT_INDEX'] = self.calculate_minimum_garanti()
        # Minimum garanti
        self.new_row['CT_MG'] = self.calculate_minimum_garanti()
        # variable fees
        self.new_row['CT_VARFEE'] = self.calculate_variable_fee()
        self.new_row['CT_VARFEE_NEW'] = self.calculate_variable_fee()
        # audience
        self.new_row['CNT_NAME_PRD_MINS'] = self.calculate_cnt_name_prd_mins()
        self.new_row['CNT_NAME_PRD_MINS_TOT'] = self.calculate_cnt_name_prd_mins_tot()
        self.new_row['AUDSHARE_ALLOC_KEY'] = self.generate_audshare_alloc_key()
        self.new_row['TOTAL_COST_NETWORK'] = self.calculate_total_cost_network()
        self.new_row['FIX_YEARLY_COST_ALLOC'] = self.calculate_fix_yearly_cost_alloc()
        #Subs here ??
        self.new_row['CT_MIN_SUBS'] = self.calculate_minimum_subs()
        # CAPEX / COGS
        self.new_row['CAPEX'] = self.calculate_capex()
        self.new_row['COGS'] = self.calculate_cogs()


    def calculate_year(self):
        pass

    def calculate_minimum_garanti(self):
        pass

    def calculate_variable_fee(self):
        pass

    def calculate_cnt_name_prd_mins(self):
        pass

    def calculate_cnt_name_prd_mins_tot(self):
        pass

    def generate_audshare_alloc_key(self):
        pass

    def calculate_fix_yearly_cost_alloc(self):
        pass

    def calculate_total_cost_network(self):
        pass

    def calculate_capex(self):
        pass

    def calculate_cogs(self):
        pass

    def calculate_minimum_subs(self):
        pass



