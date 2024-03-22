# Investment-Tracker
A Python-based investment tracker that uses an Excel file for data storage, but a UI for adding and viewing investments.

It allows for the handling of:
- Fixed Term Savings Accounts
- Easy Access Savings Accounts
- Crypto Currency Investments
- Fund Investments
- ISA Investments
- NS&I Bonds


To Install:
1. Download all files
2. Save the files to the required location
3. Install the required packages by running: pip install -r requirements.txt
4. Update the NS&I Holdings sheet with the bond numbers that you hold
5. Open investmentTracker.py and update the path and filename for the Excel file. If you wish to receive iCloud calendar notifications then your iCloud username (email), password (an application-specific password can be generated) and the name of the calendar to update need specifying. 
6. The code is now ready to be ran and a UI interface will open.

Using the UI:
Each investment type can have new investments added via the UI by selecting the investments corresponding sheet type. When an investment is sold or matured it moves to the sold sheets. 

Features:
- Automatic Maturity of Fixed Term Investments - Upon launching the script the accounts are checked and if they have matured you are prompted to input the sell price of the investments.
- End of Tax Year - Upon clicking the button you will be prompted to enter the tax year (ie 2023/24), following this the interest for each investment type generated in that tax year will be saved and summed. For active Easy Access accounts tax is payable on them therefore upon clicking the button all active investments are sold and then remade with the start date being the first date of the new tax year.
- Automatic Premium Bond Checking - Each month the held premium bonds are compared with the winnings bonds and all winnings are then recorded.
- Current Investment Pricing - For funds, ISAs and crypto the current value is updated each time the script is ran.
