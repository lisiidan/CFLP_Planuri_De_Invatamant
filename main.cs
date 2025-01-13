using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;

class Program
{
    public class Discipline
    {
        public string CodDisciplina { get; set; }
        public string NumeDisciplina { get; set; }
        public string NrCredite { get; set; }
        public string FormaEvaluare { get; set; }
        public string NrOreCurs { get; set; }
        public string NrOreSeminar { get; set; }
        public string NrOreLaborator { get; set; }
        public int RowNumber { get; set; } 
        public int ColumnNumber { get; set; }
    }

    static void Main(string[] args)
    {
        Console.OutputEncoding = Encoding.UTF8; // For Romanian characters

        string inputFilePath = @"..\..\inputFiles\2023-2027_AC_PI_TI-RO.xlsx";
        string outputExcelFilePath = @"..\..\outputFiles\output_disciplines.xlsx";
        string outputCsvFilePath = @"..\..\outputFiles\output_disciplines.csv";

        List<Discipline> disciplines = new List<Discipline>();
        bool exit = false;

        while (!exit)
        {
            Console.WriteLine("Meniu:");
            Console.WriteLine("1. Extrage discipline din Excel");
            Console.WriteLine("2. Generează fișier Excel");
            Console.WriteLine("3. Generează fișier CSV");
            Console.WriteLine("4. Afișează disciplinele în consolă");
            Console.WriteLine("5. Ieși");
            Console.Write("Alege o opțiune (1-5): ");

            string choice = Console.ReadLine();

            switch (choice)
            {
                case "1":
                    disciplines = ExtractDisciplinesFromExcel(inputFilePath);
                    Console.WriteLine("Disciplina a fost extrasă cu succes.");
                    break;
                case "2":
                    if (disciplines.Count > 0)
                    {
                        GenerateExcelFile(disciplines, outputExcelFilePath);
                        Console.WriteLine("Fișierul Excel a fost generat cu succes.");
                    }
                    else
                    {
                        Console.WriteLine("Nu există discipline extrase. Te rog să extragi disciplinele mai întâi.");
                    }
                    break;
                case "3":
                    if (disciplines.Count > 0)
                    {
                        GenerateCsvFile(disciplines, outputCsvFilePath);
                        Console.WriteLine("Fișierul CSV a fost generat cu succes.");
                    }
                    else
                    {
                        Console.WriteLine("Nu există discipline extrase. Te rog să extragi disciplinele mai întâi.");
                    }
                    break;
                case "4":
                    if (disciplines.Count > 0)
                    {
                        PrintDisciplines(disciplines);
                    }
                    else
                    {
                        Console.WriteLine("Nu există discipline extrase. Te rog să extragi disciplinele mai întâi.");
                    }
                    break;
                case "5":
                    exit = true;
                    Console.WriteLine("Ieșire din program.");
                    break;
                default:
                    Console.WriteLine("Opțiune invalidă. Te rog să alegi o opțiune între 1 și 5.");
                    break;
            }

            Console.WriteLine(); // Print a blank line for better readability
        }
    }

    static List<Discipline> ExtractDisciplinesFromExcel(string filePath)
    {
        var disciplines = new List<Discipline>();
        var regex = new Regex(@"^L\d{3}\.\d{2}\.\d{2}\.[A-Za-z0-9-]+$"); // Regex pentru a verifica formatul codului

        using (var workbook = new XLWorkbook(filePath))
        {
            var worksheet = workbook.Worksheet(2); // Luăm a doua foaie de lucru

            foreach (var row in worksheet.RowsUsed())
            {
                foreach (var cell in row.CellsUsed())
                {
                    string potentialCod = cell.GetString();
                    if (regex.IsMatch(potentialCod))
                    {
                        var discipline = new Discipline
                        {
                            NumeDisciplina = worksheet.Cell(row.RowNumber() - 2, cell.Address.ColumnNumber).GetString(),
                            CodDisciplina = potentialCod,
                            NrCredite = worksheet.Cell(row.RowNumber(), cell.Address.ColumnNumber + 3).GetValue<string>(),
                            FormaEvaluare = worksheet.Cell(row.RowNumber(), cell.Address.ColumnNumber + 4).GetValue<string>(),
                            NrOreCurs = worksheet.Cell(row.RowNumber(), cell.Address.ColumnNumber + 5).GetValue<string>(),
                            NrOreSeminar = worksheet.Cell(row.RowNumber(), cell.Address.ColumnNumber + 6).GetValue<string>(),
                            NrOreLaborator = worksheet.Cell(row.RowNumber(), cell.Address.ColumnNumber + 7).GetValue<string>(),
                            RowNumber = row.RowNumber(),
                            ColumnNumber = cell.Address.ColumnNumber
                        };
                        if(discipline.NumeDisciplina.Equals("Semestrul 1"))
                            return disciplines; // Am facut asta fiindca aparent undeva dupa randul 400 avem o gramada de coduri de disciplini, fara nume, si nu inteleg ce cauta acolo, cu acest if, cand ajungem acolo, IMEDIAT iesim din cautare.
                        disciplines.Add(discipline);

                    }
                }
            }
        }

        return disciplines;
    }
    static void GenerateExcelFile(List<Discipline> disciplines, string filePath)
    {
        using (var workbook = new XLWorkbook())
        {
            var worksheet = workbook.Worksheets.Add("Disciplines");

            worksheet.Cell(2, 2).Value = "Legenda";
            worksheet.Cell(4, 2).Value = "Nume Disciplina";
            worksheet.Cell(5, 2).Value = "Cod Disciplina";
            worksheet.Cell(5, 3).Value = "Nr Credite";
            worksheet.Cell(5, 4).Value = "Forma Evaluare";
            worksheet.Cell(5, 5).Value = "Nr Ore Curs";
            worksheet.Cell(5, 6).Value = "Nr Ore Seminar";
            worksheet.Cell(5, 7).Value = "Nr Ore Laborator";

            for (int i = 0; i < disciplines.Count; i++)
            {
                foreach (var discipline in disciplines)
                {
                    worksheet.Cell(discipline.RowNumber - 1, discipline.ColumnNumber).Value = discipline.NumeDisciplina;
                    worksheet.Cell(discipline.RowNumber, discipline.ColumnNumber).Value = discipline.CodDisciplina;
                    worksheet.Cell(discipline.RowNumber, discipline.ColumnNumber + 1).Value = discipline.NrCredite;
                    worksheet.Cell(discipline.RowNumber, discipline.ColumnNumber + 2).Value = discipline.FormaEvaluare;
                    worksheet.Cell(discipline.RowNumber, discipline.ColumnNumber + 3).Value = discipline.NrOreCurs;
                    worksheet.Cell(discipline.RowNumber, discipline.ColumnNumber + 4).Value = discipline.NrOreSeminar;
                    worksheet.Cell(discipline.RowNumber, discipline.ColumnNumber + 5).Value = discipline.NrOreLaborator;
                }
            }

            workbook.SaveAs(filePath);
        }
    }

    static void GenerateCsvFile(List<Discipline> disciplines, string filePath)
    {
        using (var writer = new StreamWriter(filePath, false, Encoding.UTF8))
        {
            writer.WriteLine("Nume Disciplina,Cod Disciplina,Nr Credite,Forma Evaluare,Nr Ore Curs,Nr Ore Seminar,Nr Ore Laborator");

            foreach (var discipline in disciplines)
            {
                writer.WriteLine($"{discipline.NumeDisciplina},{discipline.CodDisciplina},{discipline.NrCredite},{discipline.FormaEvaluare},{discipline.NrOreCurs},{discipline.NrOreSeminar},{discipline.NrOreLaborator}");
            }
        }
    }
    static void PrintDisciplines(List<Discipline> disciplines)
    {
        Console.WriteLine("Nume Disciplina\tCod Disciplina\tNr Credite\tForma Evaluare\tOre Curs\tOre Seminar\tOre Laborator");

        foreach (var discipline in disciplines)
        {
            Console.WriteLine($"{discipline.NumeDisciplina}\t{discipline.CodDisciplina}\t{discipline.NrCredite}\t{discipline.FormaEvaluare}\t{discipline.NrOreCurs}\t{discipline.NrOreSeminar}\t{discipline.NrOreLaborator}");
        }
    }
}