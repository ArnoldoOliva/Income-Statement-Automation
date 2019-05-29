import gspread
import PyPDF2
from oauth2client.service_account import ServiceAccountCredentials

#This is the authorization command/items needed (use creds to create a client to interact with the Google Drive API)
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name('client_secret.json', scope)
client = gspread.authorize(creds)

#Find a workbook to work with and get all data
gs = client.open("FY'19 Income Statement").sheet1

#Global Variables
close_month = int(raw_input("What month are we closing? (Please enter in format ##): "))
close_month_col = close_month + 1

cs_password = raw_input("Password for Bank Statement (Charles Schwab): ")
wf_password = raw_input("Password for Mortgage Statement (Wells Fargo): ")
ao_password = raw_input("Password for Management Statement (Acorn & Oak): ")

#Open and obtain all necessary information from Bank Statement (Interest, Ending Cash Balance)
pdf = PyPDF2.PdfFileReader('Month End Documents\BankStatement0430191219 - 04.pdf')
page = pdf.getPage(1)
text = page.extractText()

begin_string_interest = text.find("Interest Paid") + len("Interest Paid")
end_string_interest = text.find("Withdrawals")
interest_revenue = "$" + text[begin_string_interest:end_string_interest]

begin_string_cash = text.find("Ending Balance") + len("Ending Balance")
end_string_cash = text.find("Nonsufficient Funds Fees")
end_cash_balance = text[begin_string_cash:end_string_cash]

#Open and obtain all necessary information from Mortgage Statement (Mortgage Interest)
pdf = PyPDF2.PdfFileReader('Month End Documents\Mortgage Statement - 04.pdf')
page = pdf.getPage(0)
text = page.extractText()

start_point_interest = text.find("Explanation of amount due")
begin_string_mort_interest = text.find("Interest", start_point_interest) + len("Interest")
end_string_mort_interest = text.find("Escrow", begin_string_mort_interest)
mort_interest_expense = text[begin_string_mort_interest:end_string_mort_interest]

#Rental Revenue
rent_revenue = 0
rent_row = 0

#Interest Revenue
interest_row = gs.find('Interest').row
gs.update_cell(interest_row, close_month_col, interest_revenue)

#Dividends Revenue
dividends_revenue = 0
dividends_revenue_row = 0

#Other Revenue
other_revenue = 0
other_revenue_row = 0

#Mortgage Interest Expense
mortgage_interest_row = gs.find('Mortgage Interest').row
gs.update_cell(mortgage_interest_row, close_month_col, mort_interest_expense)

#Update Cash Flow
end_cash_row = gs.find('Per Bank').row
gs.update_cell(end_cash_row, close_month_col, end_cash_balance)




