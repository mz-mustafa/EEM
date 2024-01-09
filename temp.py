from source_meta import GasGenMeta
import openpyxl
class FuelTariff:

    def __init__(self, input_file_path='input_data.xlsx', sheet_name='tariff'):

        wb = openpyxl.load_workbook(input_file_path, data_only=True)
        sheet = wb[sheet_name]

        self.tariffs = {
            'NG': {
                'tariff': sheet['B5'].value,
                'inflation': sheet['B6'].value,
                'co2_emission': sheet['B7'].value
            },
            'RLNG': {
                'tariff': sheet['B10'].value,
                'inflation': sheet['B11'].value,
                'co2_emission': sheet['B12'].value
            },
            'LPG': {
                'tariff': sheet['B15'].value,
                'inflation': sheet['B16'].value,
                'co2_emission': sheet['B17'].value
            },
            'Biogas': {
                'tariff': sheet['B20'].value,
                'inflation': sheet['B21'].value,
                'co2_emission': sheet['B22'].value
            },
            'HFO': {
                'tariff': sheet['B25'].value,
                'inflation': sheet['B26'].value,
                'co2_emission': sheet['B27'].value
            },
            'Diesel': {
                'tariff': sheet['B30'].value,
                'inflation': sheet['B31'].value,
                'co2_emission': sheet['B32'].value
            }
        }

    def get_tariff_and_inflation(self, fuel_name):
        return self.tariffs.get(fuel_name, {'tariff': None, 'inflation': None, 'co2_emission': None})




class Source:
    def __init__(self, n, source_type, src_priority):
        self.n = n
        self.inputs = self.create_input_structure()
        self.outputs = self.create_output_structure()
        self.source_type = source_type
        self.priority = src_priority

    def create_input_structure(self):
        return {
            year: {
                'count_prim_units': 0,
                'rating_prim_units': 0
            }
            for year in range(0, self.n + 1)
        }

    def create_output_structure(self):
        output = {}
        for year in range(0, self.n + 1):
            output[year] = {
                'capital_cost': 0,
                'depreciation_cost': 0,
            }
            for month in range(1, 13):
                output[year][month] = {
                    'energy_output_prim_units': 0,
                    'fixed_opex': 0,
                    'num_pot_failures': 0,
                    'num_failures': 0,
                    'failure_duration': 0,
                    'co2_emissions': 0
                }
                for hour in range(1, 25):
                    output[year][month][hour] = {
                        'power_output_prim_units': 0,
                        'loading_prim_units': 0
                    }
        return output

class GasGenSource(Source):
    def __init__(self, n, src_p):
        super().__init__(n, 'Gas Generator', src_p)
        self.meta = GasGenMeta()
        self.extend_input_structure()
        self.extend_output_structure()

    def extend_input_structure(self):
        # Extend input structure with GasGenSource specific keys
        for year in range(self.n + 1):  # Use range based on n to iterate over years
            self.inputs[year]['rating_backup_units'] = 0
            self.inputs[year]['count_backup_units'] = 0
            self.inputs[year]['perc_rated_output'] = 0

        # These two keys are not associated with a specific year, so they remain the same
        self.inputs['chp_operation'] = False
        self.inputs['fuel_type'] = 'NG'

    def extend_output_structure(self):
        # Extend output structure with GasGenSource specific keys
        for year in range(self.n + 1):
            for month in range(1, 13):  # Use range for months
                self.outputs[year][month]['energy_output_backup_units'] = 0
                self.outputs[year][month]['energy_free_cooling'] = 0
                self.outputs[year][month]['var_opex'] = 0
                self.outputs[year][month]['fuel_charges'] = 0
                for hour in range(1, 25):  # Use range for hours
                    self.outputs[year][month][hour]['power_output_backup_units'] = 0
                    self.outputs[year][month][hour]['loading_backup_units'] = 0

class Scenario:
    def __init__(self, name, client_name, input_file_path='input_data.xlsx', n=5):
        self.name = name
        self.client_name = client_name
        self.timestamp = datetime.datetime.now()
        self.scenario_spec = {}
        self.sources_dict = {}
        self.sources_list = []
        self.ip_site_data = {}
        self.ip_load_data = {}
        self.ip_enr_data = {}
        #OUTPUT DATAFRAMES
        self.power_df = []
        self.energy_df = []
        self.capex_df = []
        self.opex_df = []
        self.emissions_df = []

        def add_source(self, source):

            self.sources_dict[source.source_type] = source
            self.sources_list.append(source)
            self.sources_list.sort(key= lambda src: src.priority)

def opex_calculation(self):
        fuel_tariff = FuelTariff()  # Create an instance of the FuelTariff class
        for y in range(1, self.n + 1):

            for m in range(1, 13):

                month_data = {'year': y, 'month': m}
                interrupt_loss = 0
                outage_loss = 0

                for src in self.sources_list:

                    src_name = src.source_type
                    # Calculate total capacity up till current year
                    total_capacity = sum(
                        yr_data['count_prim_units'] * yr_data['rating_prim_units']
                        for yr, yr_data in src.inputs.items() if isinstance(yr, int) and yr <= y
                    )
                    if src_name != 'Grid' and src_name != 'PPA':

                        total_capex = sum([src.outputs[y]['capital_cost'] for y in range(y + 1)])
                        # Compute the annual depreciation
                        src.outputs[y]['depreciation_cost'] = total_capex / src.meta.useful_life

                    inflation_rate = pow(1 + src.meta.opex_inflation_rate, y)
                    src_mnth_op = src.outputs[y][m]

                    # Calculate and save fixed OPEX
                    if src_name != 'Grid' and src_name != 'PPA':
                        src_mnth_op['fixed_opex'] = total_capacity * \
                                                    src.meta.fixed_opex_baseline * inflation_rate

                        month_data[f'{src_name} Depreciation Cost, M PKR'] = \
                            src.outputs[y]['depreciation_cost'] / (12 * 1000000)
                        month_data[f'{src_name} Fixed Opex, M PKR'] = src_mnth_op['fixed_opex'] / 1000000
                    else:
                        src_mnth_op['fixed_charges'] = total_capacity * \
                                                       src.meta.tariff_baseline_fixed * inflation_rate
                        month_data[f'{src_name} Fixed Opex, M PKR'] = src_mnth_op['fixed_charges'] / 1000000

                    # Calculate and save variable OPEX
                    if src_name not in ['BESS', 'Solar', 'Wind', 'PPA','Grid']:
                        src_mnth_op['var_opex'] = src_mnth_op['energy_output_prim_units'] * \
                                                  src.meta.var_opex_baseline * inflation_rate
                        month_data[f'{src_name} Var OPEX, M PKR'] = src_mnth_op['var_opex'] / 1000000

                    # Calculate and save energy charges for Grid
                    if src_name == 'Grid':
                        src_mnth_op['peak_enr_charges'] = src_mnth_op['energy_output_peak'] * \
                                                          src.meta.tariff_baseline_var_peak * inflation_rate
                        src_mnth_op['offpeak_enr_charges'] = src_mnth_op['energy_output_offpeak'] * \
                                                             src.meta.tariff_baseline_var_offpeak * inflation_rate
                        month_data['Grid Peak Rate Energy Cost, M PKR'] = src_mnth_op['peak_enr_charges'] / 1000000
                        month_data['Grid Offpeak Rate Energy Cost, M PKR'] = src_mnth_op['offpeak_enr_charges'] / 1000000
                    
                    if src_name == 'PPA':
                        src_mnth_op['enr_charges'] = src_mnth_op['energy_output_prim_units'] * \
                                                          src.meta.tariff_baseline_var * inflation_rate
                        month_data['PPA Energy Cost, M PKR'] = src_mnth_op['enr_charges'] / 1000000

                    # Calculate and save fuel costs
                    if src_name in ['Gas Generator', 'HFO Generator', 'HFO+Gas Generator', 'Diesel Generator']:

                        fuel_type = src.inputs['fuel_type']
                        fuel_data = fuel_tariff.get_tariff_and_inflation(fuel_type)

                        src_mnth_op['fuel_charges'] = src_mnth_op['energy_output_prim_units'] * fuel_data[
                            'tariff'] * pow(1 + fuel_data['inflation'], y)
                        month_data[f'Fuel Charges for {fuel_type}, M PKR'] = src_mnth_op['fuel_charges'] / 1000000

                        if src_name == 'HFO+Gas Generator':
                            sec_fuel_type = src.inputs['sec_fuel_type']
                            fuel_data = fuel_tariff.get_tariff_and_inflation(sec_fuel_type)
                            src_mnth_op['fuel_charges_sec'] = src_mnth_op['energy_output_prim_units_sec'] * \
                                                              fuel_data['tariff'] * pow(1 + fuel_data['inflation'], y)
                            month_data[f'Fuel Charges for {sec_fuel_type}, M PKR'] = \
                                src_mnth_op['fuel_charges_sec'] / 1000000

                    # Calculate the cost of interruptions
                    interrupt_loss += src.outputs[y][m]['num_failures'] * \
                                      self.ip_site_data['fail_loss_immediate'] / 1000000

                    outage_loss += src.outputs[y][m]['failure_duration'] * \
                                   self.ip_site_data['fail_loss_over_time'] / 1000000
                month_data['Loss due to Interruptions, M PKR'] = interrupt_loss
                month_data['Loss due to Outage, M PKR'] = outage_loss
                self.opex_df.append(month_data)
        self.opex_df = pd.DataFrame(self.opex_df)