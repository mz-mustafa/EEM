import tkinter as tk
from tkinter import scrolledtext

import openpyxl


# Create a function to display the summaries in a GUI
def display_data(solar_instance, wind_instance, grid_instance):
    # Create a new Tkinter window
    window = tk.Tk()
    window.title("Data Summary")

    # Create a scrolled text widget to display the summaries
    text_widget = scrolledtext.ScrolledText(window, width=150, height=35)
    text_widget.pack()

    # Get the summaries from the instances
    solar_summary = solar_instance.summary()
    wind_summary = wind_instance.summary()
    grid_summary = grid_instance.summary()

    # Display the summaries in the text widget
    text_widget.insert(tk.END, "Solar Data Summary:\n\n")
    text_widget.insert(tk.END, "Solar Data:\n")
    text_widget.insert(tk.END, solar_summary + "\n\n")
    text_widget.insert(tk.END, "Wind Data Summary:\n\n")
    text_widget.insert(tk.END, "Wind Data:\n")
    text_widget.insert(tk.END, wind_summary + "\n\n")
    text_widget.insert(tk.END, "Grid Data Summary:\n\n")
    text_widget.insert(tk.END, "Grid Data:\n")
    text_widget.insert(tk.END, grid_summary + "\n\n")

    # Start the Tkinter main loop
    window.mainloop()


class Utility:
    def __init__(self, sanctioned_load):
        self.sanctioned_load = sanctioned_load
        self.rated_output_of_plant = self.sanctioned_load
        self.available = 1 if self.rated_output_of_plant > 0 else 0

    def display(self):
        print("Sanctioned Load:", self.sanctioned_load, "MW")
        print("Rated Output of Plant:", self.rated_output_of_plant, "MW")
        print("Available:", "Yes" if self.available == 1 else "No")


# Sample values
utility1 = Utility(500)
utility2 = Utility(800)

# Display sample utility information
print("Utility 1 Information:")
utility1.display()

print("\nUtility 2 Information:")
utility2.display()


class DGSET:
    # Class attributes
    max_power_output_percentage = 98.5
    min_power_output_percentage = 30.0

    def __init__(self, num_generators, installed_capacity_per_unit, units_for_operation, units_for_backup):
        self.num_generators = num_generators
        self.installed_capacity_per_unit = installed_capacity_per_unit
        self.units_for_operation = units_for_operation
        self.units_for_backup = units_for_backup
        self.total_units_installed = self.units_for_operation + self.units_for_backup
        self.rated_output_of_plant = self.installed_capacity_per_unit * self.units_for_operation
        self.available = 1 if self.total_units_installed > 0 else 0

    def display(self):
        print("Number of Generators:", self.num_generators)
        print("Installed Capacity per Unit:", self.installed_capacity_per_unit, "MW")
        print("Units for Operation:", self.units_for_operation)
        print("Units for Backup:", self.units_for_backup)
        print("Total Units Installed:", self.total_units_installed)
        print("Rated Output of Plant:", self.rated_output_of_plant, "MW")
        print("Max Power Output per Unit (%):", self.max_power_output_percentage, "%")
        print("Min Power Output per Unit (%):", self.min_power_output_percentage, "%")
        print("Available:", "Yes" if self.available == 1 else "No")


# Sample values
dgset1 = DGSET(4, 200, 3, 1)
dgset2 = DGSET(2, 150, 1, 1)

# Display sample DGSET information
print("DGSET 1 Information:")
dgset1.display()

print("\nDGSET 2 Information:")
dgset2.display()


class SolarOngrid:
    # Class attributes
    max_power_output_percentage = 98.5
    min_power_output_percentage = 30.0

    def __init__(self, installed_capacity_per_unit):
        self.installed_capacity_per_unit = installed_capacity_per_unit
        self.rated_output_of_plant = self.installed_capacity_per_unit
        self.available = 1 if self.rated_output_of_plant > 0 else 0

    def display(self):
        print("Installed Capacity per Unit:", self.installed_capacity_per_unit, "MW")
        print("Rated Output of Plant:", self.rated_output_of_plant, "MW")
        print("Max Power Output per Unit (%):", self.max_power_output_percentage, "%")
        print("Min Power Output per Unit (%):", self.min_power_output_percentage, "%")
        print("Available:", "Yes" if self.available == 1 else "No")


# Sample values
solar1 = SolarOngrid(50)
solar2 = SolarOngrid(75)

# Display sample SolarOngrid information
print("SolarOngrid 1 Information:")
solar1.display()

print("\nSolarOngrid 2 Information:")
solar2.display()


class SolarHybrid:
    # Class attributes
    max_power_output_percentage = 100.0
    min_power_output_percentage = 0.0

    def __init__(self, installed_capacity_per_unit, rated_battery_mwh, bess_charging_time_hrs):
        self.installed_capacity_per_unit = installed_capacity_per_unit
        self.rated_output_of_plant = self.installed_capacity_per_unit
        self.rated_battery_mwh = rated_battery_mwh
        self.bess_charging_time_hrs = bess_charging_time_hrs
        self.available = 1 if self.rated_output_of_plant > 0 else 0

    def display(self):
        print("Installed Capacity per Unit:", self.installed_capacity_per_unit, "MW")
        print("Rated Output of Plant:", self.rated_output_of_plant, "MW")
        print("Rated Battery MWh:", self.rated_battery_mwh, "MWh")
        print("BESS Charging Time (Hrs):", self.bess_charging_time_hrs, "hrs")
        print("Max Power Output per Unit (%):", self.max_power_output_percentage, "%")
        print("Min Power Output per Unit (%):", self.min_power_output_percentage, "%")
        print("Available:", "Yes" if self.available == 1 else "No")


# Sample values
solar_hybrid1 = SolarHybrid(50, 20, 2)
solar_hybrid2 = SolarHybrid(75, 30, 3)

# Display sample SolarHybrid information
print("SolarHybrid 1 Information:")
solar_hybrid1.display()

print("\nSolarHybrid 2 Information:")
solar_hybrid2.display()


class Wind:
    # Class attributes
    max_power_output_percentage = 100.0
    min_power_output_percentage = 0.0

    def __init__(self, installed_capacity_per_unit, units_for_operation, units_for_backup):
        self.installed_capacity_per_unit = installed_capacity_per_unit
        self.units_for_operation = units_for_operation
        self.total_units_installed = units_for_operation + units_for_backup
        self.rated_output_of_plant = self.installed_capacity_per_unit * self.total_units_installed
        self.available = 1 if self.rated_output_of_plant > 0 else 0

    def display(self):
        print("Installed Capacity per Unit:", self.installed_capacity_per_unit, "MW")
        print("Units for Operation:", self.units_for_operation)
        print("Total Units Installed:", self.total_units_installed)
        print("Rated Output of Plant:", self.rated_output_of_plant, "MW")
        print("Max Power Output per Unit (%):", self.max_power_output_percentage, "%")
        print("Min Power Output per Unit (%):", self.min_power_output_percentage, "%")
        print("Available:", "Yes" if self.available == 1 else "No")


# Sample values
wind1 = Wind(2, 10, 2)
wind2 = Wind(3, 15, 3)

# Display sample Wind information
print("Wind 1 Information:")
wind1.display()

print("\nWind 2 Information:")
wind2.display()

##############################################



class GridMetaArchived:
    def __init__(self, file_path, existing_pp_baseline, total_energy_req, chp_gas_vol, chp_mwh, ke_peak_mwh,
                 ke_offpeak_mwh,
                 operating_conditions, time, weight_ke, weight_pp):
        # Initialize Grid data from the Excel file
        self.data = {}

        workbook = openpyxl.load_workbook(file_path, data_only=True)
        sheet = workbook['Sheet3']

        """
        self.headers = []  # Initialize a list to store the headers
        # Read headers starting from cell E12 to AB12
        for col in range(5, 29, 2):  # Assuming headers are in columns E to AB
            header = sheet.cell(row=12, column=col).value
            self.headers.extend([f'{header} KE', f'{header} PP'])
        """

        # Read other data
        self.data['Unofficial charges'] = sheet['A2'].value
        self.data['NOC Fee'] = sheet['B2'].value
        self.data['ROW cost'] = sheet['C2'].value
        self.data['Supervision charges'] = sheet['D2'].value
        self.data['Cost of right of way'] = sheet['E2'].value
        self.data['Security Deposit for 11kV'] = sheet['F2'].value
        self.data['Tariff baseline fixed'] = sheet['G2'].value
        self.data['Tariff baseline variable- offpeak'] = sheet['H2'].value
        self.data['Tariff baseline variable- peak'] = sheet['I2'].value
        self.data['Failure Rate'] = sheet['J2'].value
        self.data['Average outage time per failure'] = sheet['M2'].value
        self.data['Average time of failure outage per month'] = sheet['N2'].value
        self.data['Average time of availability per month'] = sheet['O2'].value
        self.data['Sanctioned Load (MW)'] = sheet['P2'].value
        self.data['Sanctioned Load half'] = sheet['Q2'].value

        """
        self.data['Failure Loss (Immediate)'] = sheet['K2'].value
        self.data['Failure Loss (Length of time)'] = sheet['L2'].value
        """

        # Define units for each attribute
        self.units = {
            'Unofficial charges': 'Rs./MW',
            'NOC Fee': 'Rs.',
            'ROW cost': 'Rs.',
            'Supervision charges': 'Rs.',
            'Cost of right of way': 'Rs.',
            'Security Deposit for 11kV': 'Rs.',
            'Tariff baseline fixed': 'Rs./MW',
            'Tariff baseline variable- offpeak': 'Rs./MW/month',
            'Tariff baseline variable- peak': 'Rs. Per kWh',
            'Failure Rate': 'Rs. Per kWh',
            'Failure Loss (Immediate)': 'No. of Failures/Year',
            'Failure Loss (Length of time)': 'Rupees/Failure',
            'Average outage time per failure': 'Rupees/one hour of failure',
            'Average time of failure outage per month': 'hours/failure',
            'Average time of availability per month': 'hours/month',
            'Sanctioned Load (MW)': 'MW',
            'Sanctioned Load half': 'MW'
        }

        # Initialize GDE data
        # Dont think most of this is needed.
        """
        self.existing_pp_baseline = existing_pp_baseline
        self.total_energy_req = total_energy_req
        self.chp_gas_vol = chp_gas_vol
        self.chp_mwh = chp_mwh
        self.ke_peak_mwh = ke_peak_mwh
        self.ke_offpeak_mwh = ke_offpeak_mwh
        self.operating_conditions = operating_conditions
        self.time = time
        self.weight_ke = weight_ke
        self.weight_pp = weight_pp

        # Define month names
        self.months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sept', 'Oct', 'Nov', 'Dec']
        """

    # Quite sure this function is not very useful
    def calculate_month_ke(self, month_index):
        month_ke = []
        for i, condition in enumerate(self.operating_conditions):
            if condition == 'off-peak':
                ke_value = (self.ke_offpeak_mwh[month_index] * self.weight_ke[i]) / 30
            else:
                ke_value = (self.ke_peak_mwh[month_index] * self.weight_ke[i]) / 30
            month_ke.append(ke_value)
        return month_ke

    def calculate_month_pp(self, month_index):
        month_pp = []
        for i, condition in enumerate(self.operating_conditions):
            if condition in ('off-peak', 'peak'):
                pp_value = (self.chp_mwh[month_index] * self.weight_pp[i]) / 30
                month_pp.append(pp_value)
        return month_pp

    def summary(self):
        summary = "Grid_Instances Summary:\n"
        # Include the Grid data in the summary
        for key, value in self.data.items():
            summary += f"{key}: {value} {self.units[key]}\n"

        # Include the headers in the summary
        summary += "Headers:\n" + ", ".join(self.headers) + "\n"

        # Include the GDE data in the summary
        for attribute, value in self.__dict__.items():
            if attribute != 'data' and attribute != 'headers' and attribute != 'units' and attribute != 'months':
                summary += f"{attribute}: {value}\n"

        # Generate summaries for all months
        for month in self.months:
            month_ke_values = self.calculate_month_ke(0)
            summary += f"{month} KE:\n" + ", ".join([f"{value:.2f}" for value in month_ke_values]) + "\n"

            month_pp_values = self.calculate_month_pp(0)
            summary += f"{month} PP:\n" + ", ".join([f"{value:.2f}" for value in month_pp_values]) + "\n"

        return summary

