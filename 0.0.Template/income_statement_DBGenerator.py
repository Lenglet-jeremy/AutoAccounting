from openpyxl import *

ENTREPRISE_NAME = "Entreprise fictive";
EMPLOYEES = 5000;
TEMPORARY_WORKERS = 500;

income_statement_DB_file = Workbook();
DB_ws = income_statement_DB_file.active;
DB_ws.title = "DB";

DB_ws["A1"] = "Hello world !";

income_statement_DB_file.save("C:/Users/JL/Documents/2.0-AutoAccounting/0.0.Template/income_statement_DB_file.xlsx");