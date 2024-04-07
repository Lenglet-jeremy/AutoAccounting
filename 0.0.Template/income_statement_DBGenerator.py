from openpyxl import *

ENTREPRISE_NAME = "Entreprise fictive";
EMPLOYEES = 5000;
TEMPORARY_WORKERS = 500;

income_statement_DB_file = Workbook();
DB_ws = income_statement_DB_file.active;
DB_ws.title = "DB";

DB_ws["A1"] = "Nature comptable";
DB_ws["B1"] = "Designation comptable";

DB_ws["C1"] = "Centre de coût";
DB_ws["D1"] = "Designation centre de coût";

DB_ws["E1"] = "Centre de profit";
DB_ws["F1"] = "Designation centre de profit";

DB_ws["G1"] = "Montant";
DB_ws["H1"] = "Type Piece";

DB_ws["I1"] = "Nom";
DB_ws["J1"] = "Prenom";
DB_ws["K1"] = "Matricule";

DB_ws["L1"] = "Periode d'effet";
DB_ws["M1"] = "Debut periode";
DB_ws["N1"] = "Fin periode";

DB_ws["O1"] = "N° piece reference";

DB_ws["P1"] = "Utilisateur ecriture";

DB_ws["Q1"] = "Date piece";
DB_ws["R1"] = "Date comptable";
DB_ws["S1"] = "Date de saisie";

DB_ws["T1"] = "Compte contre partie";
DB_ws["U1"] = "Designation compte contre partie";

DB_ws["V1"] = "N° Ecriture";

DB_ws["W1"] = "Commentaire ecriture";

DB_ws["X1"] = "N° contre passation";
DB_ws["Y1"] = "Commentaire contre passation";

DB_ws["Z1"] = "Devise";
DB_ws["AA1"] = "Convertion en euros";
DB_ws["AB1"] = "Date convertion";
DB_ws["AC1"] = "Taux convertion";
DB_ws["AD1"] = "Source convertion";

DB_ws["AE1"] = "Societe";
DB_ws["AF1"] = "Designation societe";

DB_ws["AG1"] = "Unité de quantité";

DB_ws["AH1"] = "Quantité";
DB_ws["AI1"] = "Taux unité de quantité";

DB_ws["AJ1"] = "Code mouvement";
DB_ws["AK1"] = "Designation mouvement";



income_statement_DB_file.save("C:/Users/JL/Documents/2.0-AutoAccounting/0.0.Template/income_statement_DB_file.xlsx");