import sys
import os
import pandas as pd
from PyQt5.QtWidgets import QApplication, QMainWindow, QTableView, QVBoxLayout, QWidget, QComboBox, QSizePolicy, QPushButton, QLineEdit, QLabel, QDateEdit, QDoubleSpinBox, QMessageBox, QDialog, QHBoxLayout, QInputDialog, QDialogButtonBox
from PyQt5.QtCore import QAbstractTableModel, Qt, QDate
from datetime import datetime, timedelta
import yfinance as yf
import requests
import caldav
from dateutil.relativedelta import relativedelta
import zipfile

# Excel filename and the path to find it
filename = 'Investment_Tracker.xlsx'
path = '/Users/.../Investment_Tracker.xlsx'

# iCloud Calendar Information
username = ''
password = ''
calendar_name = ''

## TO IMPLEMENT #########################################
# NS&I Holdings and Winnigs Sheets
# Winnings - same
# Save them to excel
# Check excel for entries from that month, if so dont run code
# Check if up to 6 months previous have entries, if not check them
# Holdings - Starting Bond Number, Ending Bond Number
########################################################


## Update NS&I Premium Bonds ###########################
def generate_bond_numbers(path, sheet_name):
    # Read bond ranges from an Excel sheet
    df = pd.read_excel(path, sheet_name=sheet_name, engine='openpyxl')

    # Initialize an empty list to hold all generated bond numbers
    held_bond_numbers = []

    # Iterate through each row in the DataFrame to get start and end bond numbers
    for index, row in df.iterrows():
        start_bond, end_bond = row['Starting Bond Number'], row['Ending Bond Number']

        # Find the shared prefix for the start and end bond numbers
        prefix = find_shared_prefix(start_bond, end_bond)

        # Extract the numeric part by removing the prefix
        start_seq = int(start_bond[len(prefix):])
        end_seq = int(end_bond[len(prefix):])

        # Generate all bond numbers in the range and add them to the list
        held_bond_numbers.extend([f"{prefix}{str(seq).zfill(len(start_bond) - len(prefix))}" for seq in range(start_seq, end_seq + 1)])

    return held_bond_numbers

def generate_next_id(df):
    if df.empty or 'Unique Identifier' not in df.columns or not df['Unique Identifier'].str.startswith('P').any():
        return 'P1'  # Start from 'P1' if DataFrame is empty or no ID starts with 'P'
    else:
        max_id = df['Unique Identifier'].str.extract(r'P(\d+)').astype(int).max().iloc[0]
        return f'P{max_id + 1}'

def format_and_save_winnings_df(winnings_df, filename):
    # Drop the 'Year-Month' column if it exists
    if 'Year-Month' in winnings_df.columns:
        winnings_df.drop(columns=['Year-Month'], inplace=True)

    # Ensure 'Draw Date' is in datetime format
    winnings_df['Draw Date'] = pd.to_datetime(winnings_df['Draw Date'], dayfirst=True)

    # Convert 'Draw Date' to 'dd/mm/YYYY' format
    winnings_df['Draw Date'] = winnings_df['Draw Date'].dt.strftime('%d/%m/%Y')

    # Save the DataFrame to an Excel file
    with pd.ExcelWriter(filename, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        winnings_df.to_excel(writer, sheet_name="NS&I Winnings", index=False)

    print("Winnings data formatted and saved successfully.")

def parse_content(content):
        bond_prizes = {}
        lines = content.splitlines()
        current_prize = None

        for line in lines:
            if '£' in line:
                # This assumes that the line with £ contains the prize amount
                current_prize = line.split('£')[1].split()[0].replace(',', '').strip()
            else:
                # Extract bond numbers and associate them with the current prize
                bond_numbers = line.split()
                for number in bond_numbers:
                    if number.startswith('510SQ1'):
                        bond_prizes[number] = current_prize
        return bond_prizes

def find_shared_prefix(a, b):
    min_length = min(len(a), len(b))
    for i in range(min_length):
        if a[i] != b[i]:
            return a[:i]
    return a[:min_length]  # In case one string is a complete prefix of the other

def find_missing_months(months_to_check, winnings_df):
    data_missing = False
    missing_months = []

    # Convert winnings DataFrame 'Year-Month' column to a list of Periods for easier checking
    recorded_periods = winnings_df['Year-Month'].unique().tolist()

    # Get the period for the current month
    current_month_period = pd.Timestamp(datetime.now()).to_period('M')

    # Track the status of the previous month to identify isolated missing months
    prev_month_missing = False

    for i, month in enumerate(months_to_check):
        month_period = pd.Timestamp(month).to_period('M')

        # Check if the current month is missing
        if month_period not in recorded_periods:
            # If it's the most recent month, mark it as missing
            if month_period == current_month_period:
                data_missing = True
                missing_months.append(month)
                prev_month_missing = True
            else:
                # Check if the next month is also missing or if the previous month was missing
                if (i < len(months_to_check) - 1 and pd.Timestamp(months_to_check[i + 1]).to_period('M') not in recorded_periods) or prev_month_missing:
                    data_missing = True
                    missing_months.append(month)
                    prev_month_missing = True
                else:
                    # If the next month is not missing and the previous month was not missing, it's an isolated missing month
                    prev_month_missing = False
        else:
            prev_month_missing = False

    return data_missing, missing_months

def generate_months_to_check(num_months_back):
    current_date = datetime.now()
    months_to_check = [(current_date - relativedelta(months=i)).replace(day=1) for i in range(num_months_back, -1, -1)]
    return months_to_check

filename = 'Investment_Tracker.xlsx'
path = '/Users/DanPage/Library/Mobile Documents/com~apple~CloudDocs/Documents/Projects/Investments/Investment_Tracker.xlsx'
holdings = 'NS&I Holdings'

# Read the 'Easy Access' sheet to find the highest interest rate
easy_access_df = pd.read_excel(path, sheet_name='Easy Access')
highest_interest_rate = easy_access_df['Interest Rate'].max()

# Read the 'NS&I Winnings' sheet from the Excel file, without parsing dates initially
winnings_df = pd.read_excel(path, sheet_name='NS&I Winnings', engine='openpyxl')
winnings_df['Draw Date'] = pd.to_datetime(winnings_df['Draw Date'], dayfirst=True)
winnings_df['Year-Month'] = winnings_df['Draw Date'].dt.to_period('M')

months_to_check = generate_months_to_check(6)
data_missing, missing_months = find_missing_months(months_to_check, winnings_df)

if data_missing:
    for missing_month_datetime in missing_months:
        # Format the missing month and year from the datetime object
        missing_month_str = missing_month_datetime.strftime("%m")
        missing_year_str = missing_month_datetime.strftime("%Y")
        
        # Convert the formatted month and year to integers if needed
        missing_month_int = int(missing_month_str)
        missing_year_int = int(missing_year_str)

        print(f"NS&I data missing for {missing_month_datetime.strftime('%B %Y')}.")

        # Calculate Held Bond Numbers
        held_bond_numbers = generate_bond_numbers(path, holdings)
    
        # Construct the URL for the missing month
        url = f'https://www.nsandi.com/files/asset/zip/premium-bonds-winning-bond-numbers-{missing_month_str}-{missing_year_str}.zip'

        response = requests.get(url)
        if response.status_code == 200:
            with open('premium_bonds.zip', 'wb') as file:
                file.write(response.content)

            extracted_files = []
            with zipfile.ZipFile('premium_bonds.zip', 'r') as zip_ref:
                zip_ref.extractall()
                extracted_files.extend(zip_ref.namelist())

                for file_name in zip_ref.namelist():
                    with zip_ref.open(file_name) as file:
                        content = file.read().decode('ISO-8859-1')
                        bond_prizes = parse_content(content)

                        # Create a datetime object for the first day of the missing month
                        draw_date = datetime(missing_year_int, missing_month_int, 1)

                        if isinstance(winnings_df, pd.DataFrame):
                            # Initialize a list to hold new rows DataFrames
                            new_rows = []
                            for bond in held_bond_numbers:
                                if bond in bond_prizes:
                                    unique_id = generate_next_id(winnings_df)
                                    winnings_amount = int(bond_prizes[bond])
                                    # Create a new DataFrame for the current winning detail
                                    new_row_df = pd.DataFrame({
                                        'Bond Number': [bond], 
                                        'Draw Date': [draw_date],
                                        'Winnings': [winnings_amount],
                                        'Unique Identifier': [unique_id],
                                        'Max Interest': [highest_interest_rate],
                                    })

                                    # Add the new DataFrame to the list
                                    new_rows.append(new_row_df)

                                    # Update winnings_df with the new row to ensure the next ID is unique
                                    winnings_df = pd.concat([winnings_df, new_row_df], ignore_index=True)         
                        else:
                            print("Error: winnings_df is not a DataFrame.") 

            # Cleanup: Delete extracted files
            for file_name in extracted_files:
                os.remove(file_name)
            # Cleanup: Delete the ZIP file
            os.remove('premium_bonds.zip')
        else:
            print("Failed to download the file.")

format_and_save_winnings_df(winnings_df, filename)

## Check for and Handle Matured Investments #################################################################
def check_for_matured_investments(excel_file_path):
    try:
        fixed_term_df = pd.read_excel(excel_file_path, sheet_name='Fixed Term')
        current_date = pd.to_datetime("today")

        for index, row in fixed_term_df.iterrows():
            sell_date = pd.to_datetime(row['Sell Date'], dayfirst=True)
            if sell_date <= current_date:
                return True  # There is at least one investment to mature
    except Exception as e:
        print(f"Error checking for matured investments: {e}")

    return False  # No investments to mature

def update_matured_investments_gui(excel_file_path, app):
    # Load the "Fixed Term" sheet into a DataFrame
    fixed_term_df = pd.read_excel(excel_file_path, sheet_name='Fixed Term')

    # Get the current date
    current_date = pd.to_datetime("today")

    # Attempt to load the "Matured Fixed Term" sheet; if it doesn't exist, create an empty DataFrame
    try:
        matured_df = pd.read_excel(excel_file_path, sheet_name='Matured Fixed Term')
    except Exception:
        matured_df = pd.DataFrame(columns=fixed_term_df.columns)

    # Iterate through the DataFrame
    for index, row in fixed_term_df.iterrows():
        # Convert 'Sell Date' to datetime and compare with current date
        sell_date = pd.to_datetime(row['Sell Date'], dayfirst=True)
        if sell_date <= current_date:
            # Use QInputDialog to get the actual sell amount from the user
            actual_amount, ok = QInputDialog.getDouble(None, "Enter Actual Sell Price", 
                                                       f"Enter the actual sell amount for investment {row['Unique Identifier']}:", 
                                                       decimals=2)
            if ok:
                # Update the row with the actual sell amount
                row['Actual Sell Price'] = actual_amount
                row['Sell Date'] = sell_date.strftime("%d/%m/%Y")

                # Append the updated row to the matured investments DataFrame
                matured_df = pd.concat([matured_df, pd.DataFrame([row])], ignore_index=True)
                
                # Drop the matured investment from the fixed term DataFrame
                fixed_term_df.drop(index, inplace=True)

    # Save the updated DataFrames back to the Excel file
    with pd.ExcelWriter(excel_file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        fixed_term_df.to_excel(writer, sheet_name='Fixed Term', index=False)
        matured_df.to_excel(writer, sheet_name='Matured Fixed Term', index=False)

## Add Calendar ##################################################################################################################
def find_event_by_unique_id(target_calendar, unique_id):
    # Search for events in the calendar
    results = target_calendar.events()

    for event in results:
        e = event.vobject_instance.vevent
        # Check if the unique identifier is in the event summary
        if f"[{unique_id}]" in str(e.summary.value):
            return event  # Return the event object if found

    return None  # Return None if no event is found with the unique identifier

def update_or_add_maturity_event_optimized(username, password, calendar_name, df_fixed):
    url = 'https://caldav.icloud.com/'

    # Connect to the CalDAV server
    client = caldav.DAVClient(url, username=username, password=password)
    principal = client.principal()
    calendars = principal.calendars()

    # Find the specific calendar by name
    target_calendar = None
    for calendar in calendars:
        if calendar.name == calendar_name:
            target_calendar = calendar
            break

    if not target_calendar:
        print(f"Calendar named '{calendar_name}' not found.")
        return

    for index, row in df_fixed.iterrows():
        unique_id = row['Unique Identifier']
        event_name = f"Maturity of {row['Bank Name']} Investment {unique_id}"
        event_date = pd.to_datetime(row['Sell Date']).date()
        
        # Check for existing events with the same unique identifier
        existing_event = find_event_by_unique_id(target_calendar, unique_id)
        if existing_event:
            e = existing_event.vobject_instance.vevent
            existing_event_date = e.dtstart.value
            # Check if the existing event matches the desired summary and date
            if str(e.summary.value).startswith(event_name) and existing_event_date == event_date:
                print(f"Event for investment {unique_id} on {event_date} already matches. No update needed.")
                continue  # Skip to the next iteration if the event is already correct

            # If the date is different, delete the old event
            existing_event.delete()
            print(f"Deleted outdated event for investment {unique_id}")

        # Add the new or updated event
        add_all_day_event(target_calendar, event_name, event_date, unique_id)
        print(f"Added/Updated event for maturity of investment {unique_id} on {event_date}")

def add_all_day_event(target_calendar, event_name, event_date, unique_id):
    next_day = event_date + timedelta(days=1)

    target_calendar.add_event(
        f"""BEGIN:VCALENDAR
VERSION:2.0
BEGIN:VEVENT
SUMMARY:{event_name} [{unique_id}]
DTSTART;VALUE=DATE:{event_date.strftime('%Y%m%d')}
DTEND;VALUE=DATE:{next_day.strftime('%Y%m%d')}
END:VEVENT
END:VCALENDAR"""
    )

def has_file_been_opened_today(filename):
    # Get the last modification time of the file
    mod_time = os.path.getmtime(filename)
    
    # Convert the modification time to a datetime object
    mod_date = datetime.fromtimestamp(mod_time)
    
    # Get the current date
    current_date = datetime.now()
    
    # Check if the modification date is the same as the current date
    return mod_date.date() == current_date.date()

if not has_file_been_opened_today(filename):
        df = pd.read_excel(filename, sheet_name='Fixed Term')
        update_or_add_maturity_event_optimized(username, password, calendar_name, df)
        print("Code executed and events updated.")
else:
    print("The code has already been run today. Skipping execution.")

##################################################################################################################

## Fund and Crypto Pricing ###################
def get_crypto_price(coin):
    url = f"https://api.coingecko.com/api/v3/simple/price?ids={coin}&vs_currencies=gbp"
    try:
        response = requests.get(url)
        data = response.json()
        return data[coin]['gbp']
    except Exception as e:
        print(f"Error fetching price for {coin}: {e}")
        return None

def isValidCrypto(coin_name):
    # Get the current price from the API
    current_price = get_crypto_price(coin_name)
    if current_price is not None:
        return True
    else:
        return False

def get_price(symbol):
    ticker = yf.Ticker(symbol)
    hist = ticker.history(period="1mo")
    return hist['Close'].iloc[-1]  # Return the last closing price

## Update crypto and fund pricing
def update_isa_prices():
    isaDf = pd.read_excel(filename, sheet_name='ISAs')
    for index, row in isaDf.iterrows():
        isaTicker = row['ISA Ticker']
        currentPricePerUnit = get_price(isaTicker)/100  # Fetch current price using the ticker
        numberOfUnits = row['Number of Units']
        initialInvestment = numberOfUnits * row['Average Price Per Unit']

        # Calculate current value and gain/loss
        currentValue = numberOfUnits * currentPricePerUnit
        gainLoss = currentValue - initialInvestment

        # Update the DataFrame with new values
        isaDf.at[index, 'Current Price Per Unit'] = currentPricePerUnit
        isaDf.at[index, 'Current Value'] = currentValue
        isaDf.at[index, 'Gain/Loss'] = gainLoss

    # Save the updated DataFrame back to the Excel file
    with pd.ExcelWriter(filename, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        isaDf.to_excel(writer, sheet_name='ISAs', index=False)

def update_fund_prices():
    fundDf = pd.read_excel(filename, sheet_name='Funds')
    for index, row in fundDf.iterrows():
        fundTicker = row['Fund Ticker']
        currentPricePerUnit = get_price(fundTicker)/100  # Fetch current price using the ticker
        numberOfUnits = row['Number of Units']
        initialInvestment = numberOfUnits * row['Average Price Per Unit']

        # Calculate current value and gain/loss
        currentValue = numberOfUnits * currentPricePerUnit
        gainLoss = currentValue - initialInvestment

        # Update the DataFrame with new values
        fundDf.at[index, 'Current Price Per Unit'] = currentPricePerUnit
        fundDf.at[index, 'Current Value'] = currentValue
        fundDf.at[index, 'Gain/Loss'] = gainLoss

    # Save the updated DataFrame back to the Excel file
    with pd.ExcelWriter(filename, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        fundDf.to_excel(writer, sheet_name='Funds', index=False)

def update_crypto_prices():
    cryptoDf = pd.read_excel(filename, sheet_name='Crypto')
    for index, row in cryptoDf.iterrows():
        coinName = row['Coin Name'].lower()
        currentPricePerUnit = get_crypto_price(coinName)  # Directly call the function without 'self'
      # Fetch current price using the CoinGecko API
        if currentPricePerUnit is None:
            print(f"Failed to fetch price for {coinName}. Skipping...")
            continue

        originalQuantity = row['Original Quantity']
        moneySpent = row['Money Spent'] if 'Money Spent' in row else 0  # Assuming there's a 'Money Spent' column

        # Calculate current value and gain/loss
        currentValue = originalQuantity * currentPricePerUnit
        gainLoss = currentValue - moneySpent

        # Update the DataFrame with new values
        cryptoDf.at[index, 'Current Price'] = currentPricePerUnit
        cryptoDf.at[index, 'Current Amount'] = originalQuantity  # Assuming you want to keep the original quantity
        cryptoDf.at[index, 'Current Value'] = currentValue
        cryptoDf.at[index, 'Gain/Loss'] = gainLoss

    # Save the updated DataFrame back to the Excel file
    with pd.ExcelWriter(filename, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        cryptoDf.to_excel(writer, sheet_name='Crypto', index=False)

update_isa_prices()
update_fund_prices()
update_crypto_prices()

## Update Easy Access Accounts
def calculate_interest(principal, rate, compound_frequency, time_in_years):
    n = 12 if compound_frequency == "Monthly" else 1
    interest = principal * (1 + rate / n) ** (n * time_in_years) - principal
    interest = round(interest,2)
    return interest

def parse_date(date_str):
    for fmt in ("%Y-%m-%d", "%d/%m/%Y"):  # Try both formats
        try:
            return datetime.strptime(date_str, fmt)
        except ValueError:
            pass  # if the format doesn't match, pass and try the next format
    raise ValueError(f"time data {date_str} does not match any supported format")

def update_easy_access_interest():
    # Load the Easy Access Accounts data
    df = pd.read_excel(filename, sheet_name='Easy Access')
    #print(df.columns)

    # Iterate over each row to calculate and update interest and sell price
    for index, row in df.iterrows():
        # Check if 'Start Date' is already a datetime object
        if isinstance(row['Start Date'], datetime):
            start_date = row['Start Date']
        else:
            start_date = parse_date(row['Start Date'])
        
        # Convert datetime object to QDate
        start_qdate = QDate(start_date.year, start_date.month, start_date.day)
        end_qdate = QDate.currentDate()  # Use current date as the end date
        days_between = start_qdate.daysTo(end_qdate)
        years_between = days_between / 365.25

        # Calculate interest
        interest_earned = calculate_interest(row['Initial Investment'], row['Interest Rate'] / 100, row['Compound Frequency'], years_between)
        est_sell_price = row['Initial Investment'] + interest_earned

        # Update the DataFrame
        df.at[index, 'Interest Earned'] = interest_earned
        df.at[index, 'Est. Sell Price'] = est_sell_price

    # Save the updated DataFrame back to the Excel file
    with pd.ExcelWriter(filename, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        df.to_excel(writer, sheet_name='Easy Access', index=False)

update_easy_access_interest()

# Handles the table view
class PandasModel(QAbstractTableModel):
    def __init__(self, data):
        super(PandasModel, self).__init__()
        self._data = data

    def rowCount(self, parent=None):
        return self._data.shape[0]

    def columnCount(self, parent=None):
        return self._data.shape[1]

    def data(self, index, role=Qt.DisplayRole):
        if index.isValid() and role == Qt.DisplayRole:
            value = self._data.iloc[index.row(), index.column()]
            if pd.isna(value):  # Check if the value is NaN
                return ""  # Return an empty string for NaN values
            if isinstance(value, (int, float)):
                # Format the number as a string with 2 decimal places
                return "{:.2f}".format(value)
            return str(value)
        return None

    def headerData(self, section, orientation, role):
        if role == Qt.DisplayRole:
            if orientation == Qt.Horizontal:
                return str(self._data.columns[section])
            else:
                return str(self._data.index[section])
        return None

# Handles selling investments
class SellInvestmentDialog(QDialog):
    def __init__(self, parent=None):
        super(SellInvestmentDialog, self).__init__(parent)
        self.setWindowTitle("Sell Investment")
        layout = QVBoxLayout(self)

        # Setup ID, Sell Amount, and Sell Date inputs
        self.idLabel = QLabel("ID:")
        self.idInput = QLineEdit()
        self.sellAmountLabel = QLabel("Sell Amount:")
        self.sellAmountInput = QLineEdit()
        self.sellDateLabel = QLabel("Sell Date:")
        self.sellDateInput = QDateEdit()
        self.sellDateInput.setCalendarPopup(True)
        self.sellDateInput.setDate(QDate.currentDate())

        # Add widgets to layout
        layout.addWidget(self.idLabel)
        layout.addWidget(self.idInput)
        layout.addWidget(self.sellAmountLabel)
        layout.addWidget(self.sellAmountInput)
        layout.addWidget(self.sellDateLabel)
        layout.addWidget(self.sellDateInput)

        # Setup Ok and Cancel buttons
        self.buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        self.buttons.accepted.connect(self.accept)
        self.buttons.rejected.connect(self.reject)
        layout.addWidget(self.buttons)

    def accept(self):
        # Collect the input values
        investment_id = self.idInput.text().strip().upper()
        sell_amount = float(self.sellAmountInput.text())
        sell_date = self.sellDateInput.date().toString("dd/MM/yyyy")

        # Signal the main window to process the sell operation
        self.parent().processSellOperation(investment_id, sell_amount, sell_date)
        super().accept()

# Handles Easy Access Rollover 
class SellPricesDialog(QDialog):
    def __init__(self, investments, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Enter Sell Prices")
        layout = QVBoxLayout(self)
        self.sell_price_inputs = {}
        for investment in investments:
            row_layout = QHBoxLayout()
            label = QLabel(f"Investment {investment['Unique Identifier']}:")
            row_layout.addWidget(label)
            sell_price_input = QLineEdit()
            row_layout.addWidget(sell_price_input)
            self.sell_price_inputs[investment['Unique Identifier']] = sell_price_input
            layout.addLayout(row_layout)
        self.buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        self.buttons.accepted.connect(self.accept)
        self.buttons.rejected.connect(self.reject)
        layout.addWidget(self.buttons)

    def getSellPrices(self):
        sell_prices = {}
        for id, input in self.sell_price_inputs.items():
            sell_prices[id] = float(input.text())
        return sell_prices

# Handles Deletion of Investments
class DeleteInvestmentDialog(QDialog):
    def __init__(self, parent=None):
        super(DeleteInvestmentDialog, self).__init__(parent)
        self.setWindowTitle("Delete Investment")
        layout = QVBoxLayout(self)

        self.idLabel = QLabel("ID:")
        self.idInput = QLineEdit()
        layout.addWidget(self.idLabel)
        layout.addWidget(self.idInput)

        # Setup Ok and Cancel buttons
        self.buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        self.buttons.accepted.connect(self.accept)
        self.buttons.rejected.connect(self.reject)
        layout.addWidget(self.buttons)

    def accept(self):
        # Collect the input values
        investment_id = self.idInput.text().strip().upper()

        # Signal the main window to process the sell operation
        self.parent().processDelete(investment_id)
        super().accept()

# Handles most things
class InvestmentApp(QMainWindow):

    ## Tax Year ##########
    def parse_tax_year(self, tax_year_str):
        start_year, end_year_suffix = tax_year_str.split('/')
        start_date = datetime.strptime(f"04/06/{start_year}", "%m/%d/%Y")
        end_year = int(start_year) + 1
        end_date = datetime.strptime(f"04/05/{end_year}", "%m/%d/%Y")
        next_tax_year_start_date = datetime.strptime(f"04/06/{end_year}", "%m/%d/%Y")
        return start_date, end_date, next_tax_year_start_date
    
    def sell_and_recreate_easy_access(self, tax_year_str):
        start_year, end_year_suffix = tax_year_str.split('/')
        start_year = int(start_year)
        end_year = int(start_year) + 1
        next_tax_year_start_date = datetime.strptime(f"04/05/{end_year}", "%m/%d/%Y")

        df_easy_access = pd.read_excel(self.excelFile, sheet_name='Easy Access')
        df_sold_easy_access = pd.read_excel(self.excelFile, sheet_name='Sold Easy Access')

        combined_ids = pd.concat([df_easy_access['Unique Identifier'], df_sold_easy_access['Unique Identifier']])
        highest_id = combined_ids.apply(lambda id: int(id[1:])).max()

        for index, row in df_easy_access.iterrows():
            actual_sell_price, ok = QInputDialog.getDouble(self, "Sell Price", f"Enter the Actual Sell Price for investment {row['Unique Identifier']}: ")
            
            if ok:
                sold_row = row.copy().to_frame().T
                sold_row['Actual Sell Price'] = actual_sell_price
                sold_row['Sell Date'] = next_tax_year_start_date.strftime("%d/%m/%Y")
                sold_row['Actual Interest'] = actual_sell_price - row['Initial Investment']
                df_sold_easy_access = pd.concat([df_sold_easy_access, sold_row], ignore_index=True)

                highest_id += 1
                new_investment = row.copy().to_frame().T
                new_investment['Unique Identifier'] = f'E{highest_id}'
                new_investment['Initial Investment'] = actual_sell_price
                new_investment['Start Date'] = next_tax_year_start_date.strftime("%d/%m/%Y")

                # Drop columns if they exist
                for column in ['Sell Date', 'Actual Sell Price', 'Actual Interest']:
                    if column in new_investment.columns:
                        new_investment.drop(column, axis=1, inplace=True)

                df_easy_access = pd.concat([df_easy_access, new_investment], ignore_index=True)

        df_easy_access = df_easy_access[~df_easy_access['Unique Identifier'].isin(df_sold_easy_access['Unique Identifier'])]

        with pd.ExcelWriter(self.excelFile, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            df_easy_access.to_excel(writer, sheet_name='Easy Access', index=False)
            df_sold_easy_access.to_excel(writer, sheet_name='Sold Easy Access', index=False)

        QMessageBox.information(self, "Completed", "Easy Access investments have been rolled over for the new tax year.")
        self.excelFile = pd.ExcelFile(filename)
        self.loadSelectedSheet()

    def handleEndOfTaxYear(self):
        tax_year, ok = QInputDialog.getText(self, "Tax Year", "Enter the tax year (XXXX/XX):")
        if ok and tax_year:
            self.sell_and_recreate_easy_access(tax_year)
            try:
                # Parse the tax year to get the start and end dates
                start_date, end_date, _ = self.parse_tax_year(tax_year)
                
                # Initialize interest amounts for each investment type
                interest_isa = interest_funds = interest_crypto = interest_easy_access = interest_fixed_term = 0
                
                # Matured Fixed Term
                matured_fixed_term_df = pd.read_excel(self.excelFile, sheet_name='Matured Fixed Term')
                if 'Sell Date' in matured_fixed_term_df.columns:
                    matured_fixed_term_df['Sell Date'] = pd.to_datetime(matured_fixed_term_df['Sell Date'], dayfirst=True)
                    filtered_df = matured_fixed_term_df[(matured_fixed_term_df['Sell Date'] >= start_date) & (matured_fixed_term_df['Sell Date'] <= end_date)]
                    interest_fixed_term += filtered_df['Actual Interest'].sum()

                # Process other investment types
                for sheet, interest_var in zip(['ISAs', 'Funds', 'Crypto'], [interest_isa, interest_funds, interest_crypto]):
                    df = pd.read_excel(self.excelFile, sheet_name=sheet)
                    if 'Sell Date' in df.columns:
                        df['Sell Date'] = pd.to_datetime(df['Sell Date'], dayfirst=True)
                        filtered_df = df[(df['Sell Date'] >= start_date) & (df['Sell Date'] <= end_date)]
                        locals()[interest_var] += filtered_df['Actual Interest'].sum()

                # Process Easy Access (both active and sold)
                easy_access_df = pd.read_excel(self.excelFile, sheet_name='Easy Access')
                sold_easy_access_df = pd.read_excel(self.excelFile, sheet_name='Sold Easy Access')
                for df in [easy_access_df, sold_easy_access_df]:
                    if 'Sell Date' in df.columns:
                        #print(df)
                        df['Sell Date'] = pd.to_datetime(df['Sell Date'], dayfirst=True)
                        filtered_df = df[(df['Sell Date'] >= start_date) & (df['Sell Date'] <= end_date)]
                        interest_easy_access += filtered_df['Actual Interest'].sum()

                # Calculate the total interest
                total_interest = interest_isa + interest_funds + interest_crypto + interest_easy_access + interest_fixed_term

                # Update the Interest Generated sheet
                interest_generated_df = pd.DataFrame({
                    'Tax Year': [tax_year],
                    'ISA': [interest_isa],
                    'Funds': [interest_funds],
                    'Crypto': [interest_crypto],
                    'Easy Access': [interest_easy_access],
                    'Fixed Term': [interest_fixed_term],
                    'Total': [total_interest]
                })

                with pd.ExcelWriter(self.excelFile, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
                    interest_generated_df.to_excel(writer, sheet_name='Interest Generated', index=False, header=False, startrow=writer.sheets['Interest Generated'].max_row)

                QMessageBox.information(self, "Tax Year Summary", f"Interest Generated sheet updated for tax year {tax_year}.")
                # Reload the Excel file to ensure it contains the latest data
                self.excelFile = pd.ExcelFile(filename)
                self.loadSelectedSheet()  # This now updates the model and emits layoutChanged

            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to process tax year data: {e}")

    ## Sell Investments
    def processSellOperation(self, investment_id, sell_amount, sell_date):
        investments_sheets = {
            'F': 'Funds', 'E': 'Easy Access', 'C': 'Crypto', 'I': 'ISAs'
        }
        sold_sheets = {
            'F': 'Sold Funds', 'E': 'Sold Easy Access', 'C': 'Sold Crypto', 'I': 'Sold ISAs'
        }

        # Determine the type of investment based on the ID prefix
        investment_type = investment_id[0]
        
        if investment_type not in investments_sheets:
            QMessageBox.warning(self, "Error", "Invalid investment type.")
            return
        
        try:
            investments_df = pd.read_excel(self.excelFile, sheet_name=investments_sheets[investment_type])
            sold_df = pd.read_excel(self.excelFile, sheet_name=sold_sheets[investment_type])
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Error loading sheets: {e}")
            return

        # Find the investment row
        if investment_id not in investments_df['Unique Identifier'].values:
            QMessageBox.warning(self, "Error", "Investment ID not found.")
            return

        row_index = investments_df.index[investments_df['Unique Identifier'] == investment_id].tolist()[0]
        investment_row = investments_df.loc[row_index].copy()

        initial_investment = investment_row['Initial Investment']


        # Update and move the investment
        investment_row['Actual Sell Price'] = sell_amount
        investment_row['Sell Date'] = sell_date
        investment_row['Actual Interest'] = sell_amount - initial_investment
        sold_df = pd.concat([sold_df, pd.DataFrame([investment_row])], ignore_index=True)
        investments_df.drop(row_index, inplace=True)

        # Save the updated DataFrames back to the Excel file
        with pd.ExcelWriter(self.excelFile, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            investments_df.to_excel(writer, sheet_name=investments_sheets[investment_type], index=False)
            sold_df.to_excel(writer, sheet_name=sold_sheets[investment_type], index=False)

        try:
            self.excelFile = pd.ExcelFile(filename)  # Adjust the path if necessary
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Failed to reload Excel file: {e}")
            return

        QMessageBox.information(self, "Success", "Investment sold successfully.")
        self.loadSelectedSheet()

    def openSellInvestmentDialog(self):
        dialog = SellInvestmentDialog(self)
        dialog.excel_file_path = self.excelFile  # Assuming self.excelFile is the path to your Excel file
        if dialog.exec_():
            # The dialog will handle the sell operation internally
            pass

    def processDelete(self, investment_id):
        investments_sheets = {
            'F': 'Funds', 'E': 'Easy Access', 'C': 'Crypto', 'I': 'ISAs', 'S': 'Savings'
            # Add other mappings as necessary
        }
        sold_sheets = {
            'F': 'Sold Funds', 'E': 'Sold Easy Access', 'C': 'Sold Crypto', 'I': 'Sold ISAs'
            # Add other mappings for sold investments if necessary
        }

        # Determine the type of investment based on the ID prefix
        investment_type = investment_id[0]

        if investment_type not in investments_sheets:
            QMessageBox.warning(self, "Error", "Invalid investment type.")
            return

        sheet_found = False
        for sheet_dict in [investments_sheets, sold_sheets]:
            if investment_type in sheet_dict:
                try:
                    # Load the DataFrame for the specific investment sheet
                    investments_df = pd.read_excel(self.excelFile, sheet_name=sheet_dict[investment_type])
                    if investment_id in investments_df['Unique Identifier'].values:
                        # Find the row index for the investment ID and delete it
                        row_index = investments_df.index[investments_df['Unique Identifier'] == investment_id].tolist()[0]
                        investments_df = investments_df.drop(row_index)

                        # Save the updated DataFrame back to the Excel file
                        with pd.ExcelWriter(self.excelFile, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                            investments_df.to_excel(writer, sheet_name=sheet_dict[investment_type], index=False)

                        QMessageBox.information(self, "Success", "Investment deleted successfully.")
                        self.excelFile = pd.ExcelFile(self.excelFile)  # Reload the Excel file
                        self.loadSelectedSheet()  # Refresh the table view to reflect changes
                        sheet_found = True
                        break  # Stop looking through other sheets once the investment is found and deleted
                except Exception as e:
                    QMessageBox.warning(self, "Error", f"Error loading sheet: {e}")
                    return

        if not sheet_found:
            QMessageBox.warning(self, "Error", "Investment ID not found in any sheet.")

    def openDeleteInvestmentDialog(self):
        dialog = DeleteInvestmentDialog(self)
        dialog.excel_file_path = self.excelFile  # Assuming self.excelFile is the path to your Excel file
        if dialog.exec_():
            # The dialog will handle the sell operation internally
            pass
    
    def create_input_field(self, label_text, input_widget, main_layout, maximum=None, prefix=None, decimals=None, suffix=None, placeholder=None, combo_options=None):
        # Horizontal layout for each label-input pair
        input_layout = QHBoxLayout()

        # Create and configure label
        label = QLabel(label_text)
        input_layout.addWidget(label, 0)  # The second argument '0' means no stretch factor for the label

        # Create and configure input based on widget type
        if input_widget == QLineEdit:
            input_field = QLineEdit()
            if placeholder:
                input_field.setPlaceholderText(placeholder)
        elif input_widget == QDoubleSpinBox:
            input_field = QDoubleSpinBox()
            if maximum:
                input_field.setMaximum(maximum)
            if prefix:
                input_field.setPrefix(prefix)
            if decimals is not None:
                input_field.setDecimals(decimals)
            if suffix:
                input_field.setSuffix(suffix)
            if placeholder is not None:
                input_field.setValue(placeholder)
        elif input_widget == QComboBox:
            input_field = QComboBox()
            if combo_options:
                input_field.addItems(combo_options)
        elif input_widget == QDateEdit:
            input_field = QDateEdit()
            input_field.setCalendarPopup(True)
            if placeholder:  # Assuming this is for setting a default date
                input_field.setDate(placeholder)

        input_layout.addWidget(input_field, 1)  # The second argument '1' means the input field will stretch

        # Add the horizontal layout to the main layout
        main_layout.addLayout(input_layout)

        # Set the input field to expand and fill the space
        input_field.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)

        return label, input_field

    ## Total Investment
    def calculate_total_investments(self):
        total = 0
        
        # For Fixed Term
        fixed_term_df = pd.read_excel(self.excelFile, sheet_name='Fixed Term')
        total += fixed_term_df['Initial Investment'].sum()

        # For ISAs
        isa_df = pd.read_excel(self.excelFile, sheet_name='ISAs')
        total += isa_df['Current Value'].sum()

        # For Crypto
        crypto_df = pd.read_excel(self.excelFile, sheet_name='Crypto')
        total += crypto_df['Current Value'].sum()

        # For Funds
        funds_df = pd.read_excel(self.excelFile, sheet_name='Funds')
        total += funds_df['Current Value'].sum()

        # For Easy Access
        easy_access_df = pd.read_excel(self.excelFile, sheet_name='Easy Access')
        total += easy_access_df['Est. Sell Price'].sum()

        return total
    
    def update_total_investments(self):
            total = self.calculate_total_investments()
            self.totalInvestmentsLabel.setText(f"Total Investments: £{total:,.2f}")

    ## Premium Bonds Functions
    def calculateBondsInterest(self, excel_file_path):
        # Load NS&I sheet
        nsai_df = pd.read_excel(excel_file_path, sheet_name='NS&I Winnings')
        
        # Filter rows that have Unique Identifier starting with 'P'
        p_rows = nsai_df[nsai_df['Unique Identifier'].str.startswith('P', na=False)]
        
        # Check if there are any 'P' investments
        if p_rows.empty:
            return "No 'P' investments found"

        # Get the start date from the 'Draw Date' column of the earliest 'P' investment
        start_date = pd.to_datetime(p_rows['Draw Date'].min())
        #print(f"Earliest draw date: {start_date}")

        # Get the current date
        current_date = datetime.now()
        
        # Calculate the time difference in years
        time_diff_years = (current_date - start_date).days / 365.25
        #print(f"Time difference in years: {time_diff_years:.2f}")

        # Sum the total prize money won from all 'P' investments
        total_prize_money = p_rows['Winnings'].sum() if 'Winnings' in p_rows.columns else 0
        #print(f"Total prize money from 'P' investments: {total_prize_money}")

        # Calculate the interest as a percentage of the £50,000 investment
        interest = (total_prize_money / 50000) * 100
        #print(f"Interest percentage: {interest:.2f}%")

        # Return the interest calculated over the time period
        return interest

    def updateBondsInterest(self):
                total = self.calculateBondsInterest(path)
                self.bondsInterestRate.setText(f"Interest Rate: {total:,.2f}%")

    def calculateMaximumInterest(self, excel_file_path):
        # Load NS&I sheet
        nsai_df = pd.read_excel(excel_file_path, sheet_name='NS&I Winnings')
        
        # Ensure there's a 'Maximum Interest' column
        if 'Max Interest' not in nsai_df.columns:
            return "The 'Maximum Interest' column is missing."

        # Calculate the average maximum interest
        average_max_interest = nsai_df['Max Interest'].mean()

        #print(f"Average maximum interest from traditional banking: {average_max_interest:.2f}%")

        # Return the average maximum interest
        return average_max_interest

    def updateMaximumInterest(self):
        total = self.calculateMaximumInterest(path)
        #print(total)
        self.maxInterestRate.setText(f"Maximum Interest Rate: {total:,.2f}%")
    
    def easyAccessMaxInterest(self, excel_file_path):
        # Load the 'Easy Access' sheet
        easy_access_df = pd.read_excel(excel_file_path, sheet_name='Easy Access')
        
        # Assuming the interest rates are stored in a column named 'Interest Rate'
        # and are already in percentage format (e.g., 1.5 for 1.5%)
        highest_interest_rate = easy_access_df['Interest Rate'].max()
        #print('hi')
        #print(highest_interest_rate)
        
        return highest_interest_rate 

    ## init
    def __init__(self):
        super().__init__()
        self.excelFile = pd.ExcelFile(filename)

        if check_for_matured_investments(filename):
            # Call your GUI function for maturing investments
            update_matured_investments_gui(filename, app)
            self.excelFile = pd.ExcelFile(filename)
        
        super().__init__()
        self.setWindowTitle('Investment Tracker Table View')
        self.setGeometry(100, 100, 800, 600)

        # Main layout setup
        layout = QVBoxLayout()
        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

        # Dropdown to select the sheet
        self.sheetSelector = QComboBox()
        layout.addWidget(self.sheetSelector)

        # Total Investments Label
        self.totalInvestmentsLabel = QLabel("Total Investments: £0")
        layout.insertWidget(0, self.totalInvestmentsLabel)  # Insert at the top of the layout


        # Setup the table view
        self.tableView = QTableView()
        self.tableView.setMinimumSize(600, 200)  # Set minimum size (width, height)
        layout.addWidget(self.tableView)
        
        self.update_total_investments()

        # Load Excel file and populate sheet selector
        self.excelFile = pd.ExcelFile(filename)
        self.sheetNames = self.excelFile.sheet_names
        self.sheetSelector.addItems(self.sheetNames)

        # Container for input fields
        self.inputContainer = QWidget()
        inputLayout = QVBoxLayout(self.inputContainer)

        # Assuming 'inputLayout' is already defined in your class __init__ method.

        # Bank Name
        self.bankNameLabel, self.bankNameEdit = self.create_input_field("Bank Name:", QLineEdit, inputLayout)

        # Initial Investment
        self.initialInvestmentLabel, self.initialInvestmentEdit = self.create_input_field("Initial Investment:", QDoubleSpinBox, inputLayout, maximum=100000000, prefix="£ ", decimals=2)

        # Interest Rate
        self.interestRateLabel, self.interestRateEdit = self.create_input_field("Interest Rate (%):", QDoubleSpinBox, inputLayout, maximum=100, suffix=" %", decimals=2)

        # Compound Frequency
        self.compoundFrequencyLabel, self.compoundFrequencyEdit = self.create_input_field("Compound Frequency:", QComboBox, inputLayout, combo_options=["Monthly", "Annual"])

        # Start Date
        self.startDateLabel, self.startDateEdit = self.create_input_field("Start Date:", QDateEdit, inputLayout, placeholder=QDate.currentDate())

        # Sell Date
        self.sellDateLabel, self.sellDateEdit = self.create_input_field("Sell Date:", QDateEdit, inputLayout, placeholder=QDate.currentDate().addYears(1))

        # ISA Name
        self.isaNameLabel, self.isaNameEdit = self.create_input_field("ISA Name:", QLineEdit, inputLayout)

        # ISA Ticker
        self.isaTickerLabel, self.isaTickerEdit = self.create_input_field("ISA Ticker:", QLineEdit, inputLayout)

        # Number of Units (for ISAs, Funds, etc.)
        self.numberOfUnitsLabel, self.numberOfUnitsEdit = self.create_input_field("Number of Units:", QDoubleSpinBox, inputLayout, maximum=1000000)

        # Average Price Per Unit (for ISAs, Funds, etc.)
        self.averagePriceLabel, self.averagePriceEdit = self.create_input_field("Average Price Per Unit:", QDoubleSpinBox, inputLayout, maximum=1000000, prefix="£ ", decimals=2)

        # Tax Year (for ISAs)
        self.taxYearLabel, self.taxYearEdit = self.create_input_field("Tax Year:", QLineEdit, inputLayout, placeholder="XXXX/XX")

        # Fund Name
        self.fundNameLabel, self.fundNameEdit = self.create_input_field("Fund Name:", QLineEdit, inputLayout)

        # Fund Ticker
        self.fundTickerLabel, self.fundTickerEdit = self.create_input_field("Fund Ticker:", QLineEdit, inputLayout)

        # Coin Name (for Crypto)
        self.coinNameLabel, self.coinNameEdit = self.create_input_field("Coin Name:", QLineEdit, inputLayout)

        # Original Quantity (for Crypto)
        self.originalQuantityLabel, self.originalQuantityEdit = self.create_input_field("Original Quantity:", QDoubleSpinBox, inputLayout, maximum=1000000)

        # Average Buy Price (for Crypto)
        self.averageBuyPriceLabel, self.averageBuyPriceEdit = self.create_input_field("Average Buy Price:", QDoubleSpinBox, inputLayout, maximum=1000000, prefix="£ ", decimals=2)

        # Interest Rate (for Crypto, if applicable)
        self.cryptoInterestRateLabel, self.cryptoInterestRateEdit = self.create_input_field("Interest Rate (%):", QDoubleSpinBox, inputLayout, maximum=100, suffix=" %", decimals=2)

        # 'Number of Units' input field for Funds
        self.fundUnitsLabel, self.fundUnitsEdit = self.create_input_field("Number of Units:", QDoubleSpinBox, inputLayout, maximum=1000000)

        # 'Average Price Per Unit' input field for Funds
        self.fundPriceLabel, self.fundPriceEdit = self.create_input_field("Average Price Per Unit:", QDoubleSpinBox, inputLayout, maximum=1000000, prefix="£ ", decimals=2)

        # Draw Date for NS&I
        self.drawDateLabel, self.drawDateEdit = self.create_input_field("Draw Date:", QDateEdit, inputLayout, placeholder=QDate(QDate.currentDate().year(), QDate.currentDate().month(), 1))

        # Winnings Enter
        self.winningsLabel, self.winningsEdit = self.create_input_field("Winnings", QDoubleSpinBox, inputLayout, maximum=1000000, prefix="£ ", decimals=2)

        # Bond Number
        self.bondNumLabel, self.bondNumEdit = self.create_input_field("Bond Number:", QLineEdit, inputLayout)

        # Max Interest Rate for NS&I
        self.bondIntLabel, self.bondIntEdit = self.create_input_field("Max Interest Rate (%):", QDoubleSpinBox, inputLayout, maximum=100, suffix=" %", decimals=2, placeholder=self.easyAccessMaxInterest(path))


        # Notes (common for all types)
        self.notesLabel, self.notesEdit = self.create_input_field("Notes:", QLineEdit, inputLayout)

        # Add Investment Button
        self.addButton = QPushButton("Add Investment")
        self.addButton.clicked.connect(self.addInvestment)
        inputLayout.addWidget(self.addButton)

        # Sell Investment Button
        self.sellButton = QPushButton("Sell Investment")
        self.sellButton.clicked.connect(self.openSellInvestmentDialog)
        inputLayout.addWidget(self.sellButton)  # Add the button to the same layout as the Add button

        # End of Tax Year Button
        self.endOfTaxYearButton = QPushButton("End of Tax Year")
        self.endOfTaxYearButton.clicked.connect(self.handleEndOfTaxYear)
        inputLayout.addWidget(self.endOfTaxYearButton)  # Add the button to the layout

        # Add Winnings Button
        self.winningsButton = QPushButton("Add Winnings")
        self.winningsButton.clicked.connect(self.addWinnings)
        inputLayout.addWidget(self.winningsButton)

        # Add Delete Investment Button
        self.deleteButton = QPushButton("Delete Investment")
        self.deleteButton.clicked.connect(self.openDeleteInvestmentDialog)
        inputLayout.addWidget(self.deleteButton)

        # Premium Bonds Interest Rate
        self.bondsInterestRate = QLabel("Interest Rate: 0.0%")
        layout.insertWidget(0,self.bondsInterestRate)
        self.bondsInterestRate.setVisible(False)

        # Maximum Interest Rate
        self.maxInterestRate = QLabel("Maximum Interest Rate: 0.0%")
        layout.insertWidget(0,self.maxInterestRate)
        self.maxInterestRate.setVisible(False)
        self.updateBondsInterest()
        self.updateMaximumInterest()

        # Initially hide the input container
        self.inputContainer.setVisible(False)
        layout.addWidget(self.inputContainer)

        # Connect sheet selection change to update UI
        self.sheetSelector.currentIndexChanged.connect(self.loadSelectedSheet)

         # Group all fields and labels into a dictionary for easier management
        self.allFields = {
            'bankName': (self.bankNameLabel, self.bankNameEdit),
            'initialInvestment': (self.initialInvestmentLabel, self.initialInvestmentEdit),
            'interestRate': (self.interestRateLabel, self.interestRateEdit),
            'compoundFrequency': (self.compoundFrequencyLabel, self.compoundFrequencyEdit),
            'startDate': (self.startDateLabel, self.startDateEdit),
            'sellDate': (self.sellDateLabel, self.sellDateEdit),
            'notes': (self.notesLabel, self.notesEdit),
            'isaName': (self.isaNameLabel, self.isaNameEdit),
            'isaTicker': (self.isaTickerLabel, self.isaTickerEdit),
            'numberOfUnits': (self.numberOfUnitsLabel, self.numberOfUnitsEdit),
            'averagePrice': (self.averagePriceLabel, self.averagePriceEdit),
            'taxYear': (self.taxYearLabel, self.taxYearEdit),
            'fundName': (self.fundNameLabel, self.fundNameEdit),
            'fundTicker': (self.fundTickerLabel, self.fundTickerEdit),
            'fundUnits': (self.fundUnitsLabel, self.fundUnitsEdit),
            'fundPrice': (self.fundPriceLabel, self.fundPriceEdit),
            'coinName': (self.coinNameLabel, self.coinNameEdit),
            'originalQuantity': (self.originalQuantityLabel, self.originalQuantityEdit),
            'averageBuyPrice': (self.averageBuyPriceLabel, self.averageBuyPriceEdit),
            'cryptoInterestRate': (self.cryptoInterestRateLabel, self.cryptoInterestRateEdit),
            'winnings': (self.winningsLabel, self.winningsEdit),
            'drawDate': (self.drawDateLabel, self.drawDateEdit),
            'maxInterest': (self.bondIntLabel, self.bondIntEdit),
            'bondNumber': (self.bondNumLabel, self.bondNumEdit),

        }
        # Initially load the first sheet
        self.loadSelectedSheet()
    
    ## Set GUI Visibility
    def clearInputFields(self):
        # Clear text fields
        self.bankNameEdit.clear()
        self.isaNameEdit.clear()
        self.isaTickerEdit.clear()
        self.fundNameEdit.clear()
        self.fundTickerEdit.clear()
        self.coinNameEdit.clear()
        self.taxYearEdit.clear()
        self.notesEdit.clear()
        self.bondNumEdit.clear()

        # Reset numeric fields to default values
        self.initialInvestmentEdit.setValue(0)
        self.interestRateEdit.setValue(0)
        self.numberOfUnitsEdit.setValue(0)
        self.averagePriceEdit.setValue(0)
        self.fundUnitsEdit.setValue(0)
        self.fundPriceEdit.setValue(0)
        self.originalQuantityEdit.setValue(0)
        self.averageBuyPriceEdit.setValue(0)
        self.cryptoInterestRateEdit.setValue(0)
        self.winningsEdit.setValue(0)
        self.bondIntEdit.setValue(self.easyAccessMaxInterest(path))

        # Reset date fields to defualt dates
        self.startDateEdit.setDate(QDate.currentDate())
        self.sellDateEdit.setDate(QDate.currentDate().addYears(1))
        self.drawDateEdit.setDate(QDate(QDate.currentDate().year(), QDate.currentDate().month(), 1))

    def loadSelectedSheet(self):
        selectedSheet = self.sheetSelector.currentText()
        df = pd.read_excel(self.excelFile, sheet_name=selectedSheet)

        if hasattr(self, 'model'):
            self.model.beginResetModel()  # Notify the model that a reset is about to occur
            self.model._data = df  # Update the model's data
            self.model.endResetModel()  # Notify the model that the reset is complete
        else:
            # If the model doesn't exist, create it
            self.model = PandasModel(df)
            self.tableView.setModel(self.model)

        # Adjust visibility of input fields based on the selected sheet
        self.adjustInputFieldsVisibility(selectedSheet)

        # Show or hide the "Sell Investment" button based on the selected sheet
        self.adjustSellButtonVisibility(selectedSheet)  

        self.adjustEndOfTaxYearButtonVisibility(selectedSheet)

        self.adjustAddButtonVisibility(selectedSheet)

        self.adjustWinningsButtonVisibility(selectedSheet)

        self.adjustDeleteButtonVisibility(selectedSheet)

        # Resize the window to fit the table view
        self.resizeWindowToFitTableView()

    def adjustInputFieldsVisibility(self, selectedSheet):
        # Define all sheets in excel
        visibilitySettings = {
            'Easy Access': ['bankName', 'initialInvestment', 'interestRate', 'compoundFrequency', 'startDate', 'notes'],
            'Fixed Term': ['bankName', 'initialInvestment', 'interestRate', 'compoundFrequency', 'startDate', 'sellDate', 'notes'],
            'ISAs': ['isaName', 'isaTicker', 'numberOfUnits', 'averagePrice', 'taxYear', 'notes'],
            'Funds': ['fundName', 'fundTicker', 'fundUnits', 'fundPrice', 'notes'],
            'Crypto': ['coinName', 'originalQuantity', 'averageBuyPrice', 'cryptoInterestRate', 'notes'],
            'Interest Generated': [],
            'NS&I Winnings': ['winnings', 'drawDate', 'bondNumber', 'maxInterest'],
            'Sold Easy Access': [],
            'Matured Fixed Term': [],
            'Sold ISAs': [],
            'Sold Funds': [],
            'Sold Crypto': [],
        }

        # Hide all fields initially
        for field in self.allFields.values():
            field[0].setVisible(False)
            field[1].setVisible(False)

        # Show fields based on the selected sheet
        if selectedSheet in visibilitySettings:
            self.inputContainer.setVisible(True)
            for fieldName in visibilitySettings[selectedSheet]:
                self.allFields[fieldName][0].setVisible(True)
                self.allFields[fieldName][1].setVisible(True)
        else:
            self.inputContainer.setVisible(False)
    
    def adjustSellButtonVisibility(self, selectedSheet):
        # Define sheets where the Sell button should be visible
        sheetsWithSellOption = ['Easy Access', 'ISAs', 'Funds', 'Crypto']

        # Show or hide the Sell button based on the selected sheet
        if selectedSheet in sheetsWithSellOption:
            self.sellButton.setVisible(True)
        else:
            self.sellButton.setVisible(False)

    def adjustEndOfTaxYearButtonVisibility(self, selectedSheet):
        #print(f"Adjusting End of Tax Year Button Visibility for sheet: {selectedSheet}")
        if selectedSheet == 'Interest Generated':
            #print("Showing End of Tax Year Button")
            self.endOfTaxYearButton.setVisible(True)
        else:
            #print("Hiding End of Tax Year Button")
            self.endOfTaxYearButton.setVisible(False)
    
    def adjustWinningsButtonVisibility(self, selectedSheet):
        #print(f"Adjusting End of Tax Year Button Visibility for sheet: {selectedSheet}")
        if selectedSheet == 'NS&I Winnings':
            self.winningsButton.setVisible(True)
            self.bondsInterestRate.setVisible(True)
            self.maxInterestRate.setVisible(True)
        else:
            self.winningsButton.setVisible(False)
            self.bondsInterestRate.setVisible(False)
            self.maxInterestRate.setVisible(False)

    def adjustAddButtonVisibility(self, selectedSheet):
        # Define sheets where the Add Investment button should be visible
        sheetsWithAddOption = ['Easy Access', 'ISAs', 'Funds', 'Crypto', 'Fixed Term']

        # Show or hide the Add Investment button based on the selected sheet
        if selectedSheet in sheetsWithAddOption:
            self.addButton.setVisible(True)
        else:
            self.addButton.setVisible(False)
    
    def adjustDeleteButtonVisibility(self, selectedSheet):
        # Define sheets where the Add Investment button should be visible
        sheetsWithDelOption = ['Easy Access', 'ISAs', 'Funds', 'Crypto', 'Fixed Term', 'Sold Easy Access', 'NS&I Winnings', 'Sold Funds', 'Matured Fixed Term', 'Sold Crypto', 'Sold ISAs']

        # Show or hide the Add Investment button based on the selected sheet
        if selectedSheet in sheetsWithDelOption:
            self.deleteButton.setVisible(True)
        else:
            self.deleteButton.setVisible(False)

    def resizeWindowToFitTableView(self):
        self.tableView.resizeColumnsToContents()
        totalWidth = self.tableView.verticalHeader().width()
        totalWidth += self.tableView.horizontalScrollBar().height()
        for column in range(self.tableView.model().columnCount()):
            totalWidth += self.tableView.columnWidth(column)
        totalWidth += self.tableView.frameWidth() * 2
        self.resize(totalWidth, self.height())
    ## Add Investments & Checks
    def isValidTicker(self, ticker):
            try:
                current_price_per_unit = get_price(ticker)  # Use your existing function to fetch the price
                if pd.isna(current_price_per_unit):
                    return False
                return True
            except Exception:
                return False
    
    def addInvestment(self):
        selectedSheet = self.sheetSelector.currentText()
        bankName = self.bankNameEdit.text()
        initialInvestment = self.initialInvestmentEdit.value()
        interestRate = self.interestRateEdit.value() / 100  # Convert percentage to decimal
        compoundFrequency = self.compoundFrequencyEdit.currentText()
        startDate = self.startDateEdit.date().toString("dd/MM/yyyy")
        notes = self.notesEdit.text()

        # Initialize newData dictionary without 'Sell Date'
        newData = {
            'Bank Name': [bankName],
            'Initial Investment': [initialInvestment],
            'Interest Rate': [interestRate * 100],
            'Compound Frequency': [compoundFrequency],
            'Start Date': [startDate],
            'Notes': [notes]
        }

        #print(f"Selected Sheet: {selectedSheet}")

        # Only set up and add 'Sell Date' for Fixed Term Savings
        if selectedSheet == 'Fixed Term':
            sellDate = self.sellDateEdit.date().toString("dd/MM/yyyy")
            newData['Sell Date'] = [sellDate]  # Add 'Sell Date' only here

            # Calculate the time period 't' for interest calculation
            startQDate = QDate.fromString(startDate, "dd/MM/yyyy")
            endQDate = QDate.fromString(sellDate, "dd/MM/yyyy")
            daysBetween = startQDate.daysTo(endQDate)
            t = daysBetween / 365.25
            newData['Sell Date'] = [sellDate]

            # Generate new unique identifier
            fixedTermDf = pd.read_excel(self.excelFile, sheet_name='Fixed Term')
            maturedDf = pd.read_excel(self.excelFile, sheet_name='Matured Fixed Term')

            # Function to extract numeric part from the identifier
            def extract_id(identifier):
                if pd.isnull(identifier):
                    return 0
                numbers = ''.join(filter(str.isdigit, identifier))
                return int(numbers) if numbers else 0

            # Apply extract_id function to get numeric IDs and find the max
            maxIdFixed = fixedTermDf['Unique Identifier'].apply(extract_id).max() if 'Unique Identifier' in fixedTermDf.columns else 0
            maxIdMatured = maturedDf['Unique Identifier'].apply(extract_id).max() if 'Unique Identifier' in maturedDf.columns else 0
            maxId = max(maxIdFixed, maxIdMatured) + 1  # Increment the highest ID found
            newId = maxId + 1 if maxId > 0 else 1

            
            n = 12 if compoundFrequency == "Monthly" else 1
            A = initialInvestment * (1 + interestRate / n) ** (n * t)
            interestEarned = A - initialInvestment

            newData.update({'Unique Identifier': [f'S{newId}'], 'Interest Earned': [interestEarned], 'Est. Sell Price': [A]})

        if selectedSheet == 'Easy Access':
            # Generate new unique identifier for Easy Access Accounts
            easyAccessDf = pd.read_excel(self.excelFile, sheet_name='Easy Access')
            soldEasyAccessDf = pd.read_excel(self.excelFile, sheet_name='Sold Easy Access')

            # Extract numeric part from the identifier and find the max
            maxIdEasyAccess = easyAccessDf['Unique Identifier'].apply(lambda x: int(x[1:]) if pd.notnull(x) else 0).max()
            maxIdSoldEasyAccess = soldEasyAccessDf['Unique Identifier'].apply(lambda x: int(x[1:]) if pd.notnull(x) else 0).max()
            maxId = max(maxIdEasyAccess, maxIdSoldEasyAccess) + 1  # Increment the highest ID found
            newId = maxId + 1 if maxId > 0 else 1

            newData['Unique Identifier'] = [f'E{newId}']
        
        if selectedSheet == 'ISAs':
            isaTicker = self.isaTickerEdit.text().upper()  # Ensure ticker is uppercase

            # Validate ISA Ticker
            if not self.isValidTicker(isaTicker):
                QMessageBox.warning(self, "Invalid Ticker", "The ISA ticker is invalid. Please enter a valid ticker.")
                return  # Exit the method without adding the investment
            # Generate new unique identifier for ISAs
            isaDf = pd.read_excel(self.excelFile, sheet_name='ISAs')
            soldIsaDf = pd.read_excel(self.excelFile, sheet_name='Sold ISAs')

            # Extract numeric part from the identifier and find the max
            maxIdIsa = isaDf['Unique Identifier'].apply(lambda x: int(x[1:]) if pd.notnull(x) and x.startswith('I') else 0).max()
            maxIdSoldIsa = soldIsaDf['Unique Identifier'].apply(lambda x: int(x[1:]) if pd.notnull(x) and x.startswith('I') else 0).max()
            maxId = max(maxIdIsa, maxIdSoldIsa) + 1  # Increment the highest ID found
            newId = maxId + 1 if maxId > 0 else 1

            isaName = self.isaNameEdit.text()
            isaTicker = self.isaTickerEdit.text()
            numberOfUnits = self.numberOfUnitsEdit.value()
            averagePrice = self.averagePriceEdit.value()
            taxYear = self.taxYearEdit.text()

            # Calculate the initial investment as the product of the number of units and the average price per unit
            initialInvestment = numberOfUnits * averagePrice

            # Construct the data dictionary for ISAs including the Unique Identifier
            newData = {
                'Unique Identifier': [f'I{newId}'],
                'ISA Name': [isaName],
                'ISA Ticker': [isaTicker],
                'Number of Units': [numberOfUnits],
                'Average Price Per Unit': [averagePrice],
                'Initial Investment': [initialInvestment],  # Add the calculated initial investment
                'Tax Year': [taxYear],
                'Notes': [notes]
            }

        if selectedSheet == 'Funds':
            fundTicker = self.fundTickerEdit.text().upper()

            # Validate ticker
            if not self.isValidTicker(fundTicker):
                QMessageBox.warning(self, "Invalid Ticker", "The fund ticker is invalid. Please enter a valid ticker.")
                return  # Exit the method without adding the investment

            # Generate new unique identifier
            fundDf = pd.read_excel(self.excelFile, sheet_name='Funds')
            soldFundDf = pd.read_excel(self.excelFile, sheet_name='Sold Funds')
            
            # Extract numeric part from the identifier and find the max
            maxIdFund = fundDf['Unique Identifier'].apply(lambda x: int(x[1:]) if pd.notnull(x) and x.startswith('F') else 0).max()
            maxIdSoldFund = soldFundDf['Unique Identifier'].apply(lambda x: int(x[1:]) if pd.notnull(x) and x.startswith('F') else 0).max()
            maxId = max(maxIdFund, maxIdSoldFund) + 1  # Increment the highest ID found
            newId = maxId + 1 if maxId > 0 else 1

            fundName = self.fundNameEdit.text()
            fundTicker = self.fundTickerEdit.text()
            numberOfUnits = self.fundUnitsEdit.value()
            averagePrice = self.fundPriceEdit.value()

            initialInvestment = numberOfUnits * averagePrice
            # Construct the data dictionary for ISAs including the Unique Identifier
            newData = {
                'Unique Identifier': [f'F{newId}'],
                'Fund Name': [fundName],
                'Fund Ticker': [fundTicker],
                'Number of Units': [numberOfUnits],
                'Average Price Per Unit': [averagePrice],
                'Initial Investment': [initialInvestment],  # Add the calculated initial investment
                'Notes': [notes]
            }

        if selectedSheet == 'Crypto':
            coinName = self.coinNameEdit.text().lower()

            # Validate coin name
            if not isValidCrypto(coinName):
                QMessageBox.warning(self, "Invalid Coin", "The coin name is invalid or not supported. Please enter a valid coin name.")
                return  # Exit the method without adding the investment
            
                    # Generate new unique identifier
            cryptoDf = pd.read_excel(self.excelFile, sheet_name='Crypto')
            soldCryptoDf = pd.read_excel(self.excelFile, sheet_name='Sold Crypto')
            
            # Extract numeric part from the identifier and find the max
            maxIdCrypto = cryptoDf['Unique Identifier'].apply(lambda x: int(x[1:]) if pd.notnull(x) and x.startswith('C') else 0).max()
            maxIdSoldCrypto = soldCryptoDf['Unique Identifier'].apply(lambda x: int(x[1:]) if pd.notnull(x) and x.startswith('C') else 0).max()
            maxId = max(maxIdCrypto, maxIdSoldCrypto) + 1  # Increment the highest ID found
            newId = maxId + 1 if maxId > 0 else 1

            coinName = self.coinNameEdit.text()
            originalQuantity = self.originalQuantityEdit.value()
            averageBuyPrice = self.averageBuyPriceEdit.value()
            cryptoInterestRate = self.cryptoInterestRateEdit.value()  # Assuming interest rate is relevant for your crypto investment

            # Calculate the initial investment as the product of the original quantity and the average buy price
            initialInvestment = originalQuantity * averageBuyPrice

            # Construct the data dictionary for Crypto including the Unique Identifier
            newData = {
                'Unique Identifier': [f'C{newId}'],
                'Coin Name': [coinName],
                'Original Quantity': [originalQuantity],
                'Initial Investment': [initialInvestment],
                'Average Buy Price': [averageBuyPrice],
                'Interest Rate': [cryptoInterestRate],  # Assuming you want to store this
                'Notes': [notes]
            }

        self.clearInputFields()

        
        newDfRow = pd.DataFrame(newData)
        df = pd.read_excel(self.excelFile, sheet_name=selectedSheet)
        df = pd.concat([df, newDfRow], ignore_index=True)

        with pd.ExcelWriter(filename, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            df.to_excel(writer, sheet_name=selectedSheet, index=False)

        # Explicitly reload the Excel file to reflect changes
        self.excelFile = pd.ExcelFile(filename)
        
        # Reload the currently selected sheet to update the table view
        self.loadSelectedSheet()

    def addWinnings(self):
            selectedSheet = self.sheetSelector.currentText()
            
            if selectedSheet == 'NS&I Winnings':

                # Generate new unique identifier
                nsaiDf = pd.read_excel(self.excelFile, sheet_name='NS&I Winnings')
                
                # Extract numeric part from the identifier and find the max
                # Assuming nsaiDf is your DataFrame and 'Unique Identifier' is the column of interest
                if nsaiDf.empty or not nsaiDf['Unique Identifier'].str.startswith('P').any():
                    newId = '1'  # Start from 'P1' if DataFrame is empty or no ID starts with 'P'
                else:
                    # Extract numeric part from the identifier, find the max, and generate the next ID
                    maxIdNSI = nsaiDf['Unique Identifier'].apply(lambda x: int(x[1:]) if pd.notnull(x) and x.startswith('P') else 0).max()
                    newId = str(maxIdNSI + 1)


                #newId = maxIdNSI + 1

                winnings = self.winningsEdit.value()
                drawDate = self.drawDateEdit.date().toString("dd/MM/yyyy")
                bondNumber = self.bondNumEdit.text()
                maxInterest = self.bondIntEdit.value()

                newData = {
                    'Unique Identifier': [f'P{newId}'],
                    'Winnings': [winnings],
                    'Draw Date': [drawDate],
                    'Bond Number': [bondNumber],
                    'Max Interest': [maxInterest],
                }

            self.clearInputFields()

            
            newDfRow = pd.DataFrame(newData)
            df = pd.read_excel(self.excelFile, sheet_name=selectedSheet)
            df = pd.concat([df, newDfRow], ignore_index=True)

            with pd.ExcelWriter(filename, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                df.to_excel(writer, sheet_name=selectedSheet, index=False)

            # Explicitly reload the Excel file to reflect changes
            self.excelFile = pd.ExcelFile(filename)
            
            # Reload the currently selected sheet to update the table view
            self.loadSelectedSheet()


if __name__ == '__main__':
    app = QApplication(sys.argv)
    investmentApp = InvestmentApp()
    investmentApp.show()
    
    sys.exit(app.exec_())
    


