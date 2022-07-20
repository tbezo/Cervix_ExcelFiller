using System;
using System.Linq;
using System.Text.RegularExpressions;
using System.IO;
using System.Windows;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using VMS.TPS.Common.Model.API;
using VMS.TPS.Common.Model.Types;

// TODO: Replace the following version attributes by creating AssemblyInfo.cs. You can do this in the properties of the Visual Studio project.
[assembly: AssemblyVersion("1.0.0.3")]
[assembly: AssemblyFileVersion("1.0.0.3")]
[assembly: AssemblyInformationalVersion("1.0")]

// TODO: Uncomment the following line if the script requires write access.
// [assembly: ESAPIScript(IsWriteable = true)]

namespace VMS.TPS
{
    public class Script
    {

        const string SCRIPT_NAME = "Cervix_ExcelFiller";
        public Script()
        {
        }

        [MethodImpl(MethodImplOptions.NoInlining)]
        public void Execute(ScriptContext context /*, System.Windows.Window window, ScriptEnvironment environment*/)
        {
            List<FractionData> Fractions = new List<FractionData>();
            Patient pat = context.Patient;
            string patientInfo = pat.FirstName + " " + pat.LastName + ", " + pat.Id;

            // Excel template to copy and fill. (original file from https://www.estro.org/ESTRO/media/ESTRO/About/hdr-gyn_biol-physik-formular_2017.xlsx)
            string templateFile = @"\\path\to\hdr-gyn_biol-physik-formular_2017.xlsx";

            // Construct output file name
            string currentYear = DateTime.Now.Year.ToString();
            string patientName = context.Patient.LastName;
            string outFileExcel = @"\\path\to\patients\folder\" + currentYear + @"\hdr-gyn_biol-physik-formular - " + patientName + ".xlsx";

            // Extract data from brachy plans which contain two zeros in their file names (e.g. g1BA 001).
            try
            {
                foreach (BrachyPlanSetup bps in context.Course.BrachyPlanSetups.Where(x => x.Id.Contains("00")))
                {
                    FractionData tmpFraction = ExtractFractionData(bps);
                    if (tmpFraction != null)
                    {
                        Fractions.Add(tmpFraction);
                    }
                }
            }
            catch (Exception)
            {
                MessageBox.Show("Error while extracting brachytherapy plan data!");
                return;
            }

            if ( Fractions.Count > 0)
            {
                // Copy template to patient folder if not already done.
                if (!File.Exists(outFileExcel))
                {
                    try
                    {
                        File.Copy(templateFile, outFileExcel);
                    }
                    catch (Exception)
                    {
                        MessageBox.Show("Could not copy template file.");
                        return;
                    }
                }

                // write data from all fractions to given Excel file.
                FillExcel(outFileExcel, Fractions, patientInfo);
                // print Id's of used plans
                string report = "Fraction data written from:\n";
                foreach (FractionData frac in Fractions)
                {
                    report += frac.PlanId + "\n";
                }
                MessageBox.Show(report);
            }
            else { MessageBox.Show("Could not find any matching plans!"); }

        }


        /// <summary>
        /// Given text and a SharedStringTablePart, creates a SharedStringItem with the specified text
        /// and inserts it into the SharedStringTablePart. If the item already exists, returns its index.
        /// </summary>
        /// <param name="text">Text to add to SharedStringTable</param>
        /// <param name="shareStringPart">The SharedStringTablePart</param>
        /// <returns></returns>
        private static int InsertSharedStringItem(string text, SharedStringTablePart shareStringPart)
        {
            // If the part does not contain a SharedStringTable, create one.
            if (shareStringPart.SharedStringTable == null)
            {
                shareStringPart.SharedStringTable = new SharedStringTable();
            }

            int i = 0;

            // Iterate through all the items in the SharedStringTable. If the text already exists, return its index.
            foreach (SharedStringItem item in shareStringPart.SharedStringTable.Elements<SharedStringItem>())
            {
                if (item.InnerText == text)
                {
                    return i;
                }
                i++;
            }

            // The text does not exist in the part. Create the SharedStringItem and return its index.
            shareStringPart.SharedStringTable.AppendChild(new SharedStringItem(new DocumentFormat.OpenXml.Spreadsheet.Text(text)));
            shareStringPart.SharedStringTable.Save();

            return i;
        }


        /// <summary>
        /// Given a worksheetpart and text, writes the text to cell with given columnname and rowindex of the first worksheet.
        /// </summary>
        /// <param name="text">Text to insert into worksheet</param>
        /// <param name="worksheetPart">Worksheet to insert text into</param>
        public static void InsertText(string text, string columnName, uint rowIndex, WorksheetPart worksheetPart)
        {
            // Get the SharedStringTablePart. If it does not exist, create a new one.
            WorkbookPart WbPart = (WorkbookPart)worksheetPart.GetParentParts().First();
            SharedStringTablePart shareStringPart;
            if (WbPart.GetPartsOfType<SharedStringTablePart>().Count() > 0)
            {
                shareStringPart = WbPart.GetPartsOfType<SharedStringTablePart>().First();
            }
            else
            {
                shareStringPart = WbPart.AddNewPart<SharedStringTablePart>();
            }

            // Insert the text into the SharedStringTablePart.
            int index = InsertSharedStringItem(text, shareStringPart);

            // Insert cell A1 into the new worksheet.
            Cell cell = InsertCellInWorksheet(columnName, rowIndex, worksheetPart);

            // Set the value of cell A1.
            cell.CellValue = new CellValue(index.ToString());
            cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);

            // Save the new worksheet.
            worksheetPart.Worksheet.Save();
        }


        /// <summary>
        /// Given a column name, a row index, and a WorksheetPart, inserts a cell into the worksheet.
        /// If the cell already exists, returns it.
        /// </summary>
        /// <param name="columnName">Column to insert into.</param>
        /// <param name="rowIndex">Row to insert into.</param>
        /// <param name="worksheetPart">WorksheetPart to insert data into first found worksheet.</param>
        /// <returns></returns>
        private static Cell InsertCellInWorksheet(string columnName, uint rowIndex, WorksheetPart worksheetPart)
        {
            Worksheet worksheet = worksheetPart.Worksheet;
            SheetData sheetData = worksheet.GetFirstChild<SheetData>();
            string cellReference = columnName + rowIndex;

            // If the worksheet does not contain a row with the specified row index, insert one.
            Row row;
            if (sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).Count() != 0)
            {
                row = sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).First();
            }
            else
            {
                row = new Row() { RowIndex = rowIndex };
                sheetData.Append(row);
            }

            // If there is not a cell with the specified column name, insert one.  
            if (row.Elements<Cell>().Where(c => c.CellReference.Value == columnName + rowIndex).Count() > 0)
            {
                return row.Elements<Cell>().Where(c => c.CellReference.Value == cellReference).First();
            }
            else
            {
                // Cells must be in sequential order according to CellReference. Determine where to insert the new cell.
                Cell refCell = null;
                foreach (Cell cell in row.Elements<Cell>())
                {
                    if (cell.CellReference.Value.Length == cellReference.Length)
                    {
                        if (string.Compare(cell.CellReference.Value, cellReference, true) > 0)
                        {
                            refCell = cell;
                            break;
                        }
                    }
                }

                Cell newCell = new Cell() { CellReference = cellReference };
                row.InsertBefore(newCell, refCell);

                worksheet.Save();
                return newCell;
            }
        }


        /// <summary>
        /// Extract fraction data from given BrachyPlanSetup
        /// </summary>
        /// <param name="ps"></param>
        /// <returns></returns>
        private static FractionData ExtractFractionData(BrachyPlanSetup ps)
        {
            FractionData fraction = new FractionData(ps.Id);
            string strIndex = ps.Id.Last().ToString();
            int index = Convert.ToInt32(strIndex);

            if (index < 1 || index > 6)
            {
                MessageBox.Show("Ignoring plan id " + ps.Id.ToString() + " ,\nindex out of excel data cell range", SCRIPT_NAME);
                return null;
            }
            //else
            //{
            //    fraction = new FractionData(ps.Id);
            //}

            StructureSet ss = ps.StructureSet;

            Structure GTV = ss.Structures.FirstOrDefault(x => Regex.IsMatch(x.Id, @"GTV[_-]res", RegexOptions.IgnoreCase));
            if (GTV != null)
            {
                fraction.Structures["GTV"]["Vol"]["value"] = GTV.Volume;
                fraction.Structures["GTV"]["D98%"]["value"] = (double)ps.GetDoseAtVolume(GTV, 98.0, VolumePresentation.Relative, DoseValuePresentation.Absolute).Dose;
            }
            Structure CTV = ss.Structures.FirstOrDefault(x => Regex.IsMatch(x.Id, @"CTV[_-]hr", RegexOptions.IgnoreCase));
            if (CTV != null)
            {
                fraction.Structures["CTV"]["Vol"]["value"] = CTV.Volume;
                fraction.Structures["CTV"]["D98%"]["value"] = ps.GetDoseAtVolume(CTV, 98.0, VolumePresentation.Relative, DoseValuePresentation.Absolute).Dose;
                fraction.Structures["CTV"]["D90%"]["value"] = ps.GetDoseAtVolume(CTV, 90.0, VolumePresentation.Relative, DoseValuePresentation.Absolute).Dose;
                fraction.Structures["CTV"]["D50%"]["value"] = ps.GetDoseAtVolume(CTV, 50.0, VolumePresentation.Relative, DoseValuePresentation.Absolute).Dose;
            }
            Structure Rektum = ss.Structures.FirstOrDefault(x => Regex.IsMatch(x.Id, @"^OAR.*rektum", RegexOptions.IgnoreCase));
            if (Rektum != null)
            {
                fraction.Structures["Rectum"]["Vol"]["value"] = Rektum.Volume;
                fraction.Structures["Rectum"]["D0.1cc"]["value"] = ps.GetDoseAtVolume(Rektum, 0.1, VolumePresentation.AbsoluteCm3, DoseValuePresentation.Absolute).Dose;
                fraction.Structures["Rectum"]["D2.0cc"]["value"] = ps.GetDoseAtVolume(Rektum, 2.0, VolumePresentation.AbsoluteCm3, DoseValuePresentation.Absolute).Dose;
            }
            Structure Blase = ss.Structures.FirstOrDefault(x => Regex.IsMatch(x.Id, @"^OAR.*bla", RegexOptions.IgnoreCase));
            if (Blase != null)
            {
                fraction.Structures["Bladder"]["Vol"]["value"] = Blase.Volume;
                fraction.Structures["Bladder"]["D0.1cc"]["value"] = ps.GetDoseAtVolume(Blase, 0.1, VolumePresentation.AbsoluteCm3, DoseValuePresentation.Absolute).Dose;
                fraction.Structures["Bladder"]["D2.0cc"]["value"] = ps.GetDoseAtVolume(Blase, 2.0, VolumePresentation.AbsoluteCm3, DoseValuePresentation.Absolute).Dose;
            }
            Structure Sigma = ss.Structures.FirstOrDefault(x => Regex.IsMatch(x.Id, @"^OAR.*sigma", RegexOptions.IgnoreCase));
            if (Sigma != null)
            {
                fraction.Structures["Sigmoid"]["Vol"]["value"] = Sigma.Volume;
                fraction.Structures["Sigmoid"]["D0.1cc"]["value"] = ps.GetDoseAtVolume(Sigma, 0.1, VolumePresentation.AbsoluteCm3, DoseValuePresentation.Absolute).Dose;
                fraction.Structures["Sigmoid"]["D2.0cc"]["value"] = ps.GetDoseAtVolume(Sigma, 2.0, VolumePresentation.AbsoluteCm3, DoseValuePresentation.Absolute).Dose;
            }
            
            Structure Darm = ss.Structures.FirstOrDefault(x => Regex.IsMatch(x.Id, @"^OAR.*darm", RegexOptions.IgnoreCase));
            if (Darm != null)
            {
                fraction.Structures["Bowel"]["Vol"]["value"] = Darm.Volume;
                fraction.Structures["Bowel"]["D0.1cc"]["value"] = ps.GetDoseAtVolume(Darm, 0.1, VolumePresentation.AbsoluteCm3, DoseValuePresentation.Absolute).Dose;
                fraction.Structures["Bowel"]["D2.0cc"]["value"] = ps.GetDoseAtVolume(Darm, 2.0, VolumePresentation.AbsoluteCm3, DoseValuePresentation.Absolute).Dose;
            }
            
            IEnumerable<ReferencePoint> RefPts = ps.ReferencePoints;
            ps.DoseValuePresentation = DoseValuePresentation.Absolute;
            //if (RefPts.Any(x => x.HasLocation(ps))) { }
            foreach (ReferencePoint rp in RefPts)
            {
                if (rp.HasLocation(ps))
                {
                    if (rp.Id.Contains("li"))
                    {
                        fraction.Structures["DoseA"]["A_li"]["value"] = ps.Dose.GetDoseToPoint(rp.GetReferencePointLocation(ps.StructureSet.Image)).Dose;
                    }
                    if (rp.Id.Contains("re"))
                    {
                        fraction.Structures["DoseA"]["A_re"]["value"] = ps.Dose.GetDoseToPoint(rp.GetReferencePointLocation(ps.StructureSet.Image)).Dose;
                    }
                }
            }
            // mean is calculated in excel, but should it?
            //fraction.Structures["DoseA"]["A_mean"]["value"] = (fraction.Structures["DoseA"]["A_li"]["value"] + fraction.Structures["DoseA"]["A_re"]["value"]) / 2;

            IEnumerable<Catheter> cats = ps.Catheters;
            if (cats.Count() > 1)
            {
                double ttime = ps.Catheters.Sum(x => x.GetTotalDwellTime());
                fraction.Structures["TRAK"]["Sum"]["value"] = (ttime * 4.07 / 3600);

                if (ps.Catheters.Any<Catheter>(x => x.Id.Equals("Tandem")))
                {
                    fraction.Structures["TRAK"]["Tandem"]["value"] = ps.Catheters.First<Catheter>(x => x.Id.Equals("Tandem")).GetTotalDwellTime() / ttime;
                }
                // Ring not present if using Aarhus Template
                if (ps.Catheters.Any<Catheter>(x => x.Id.Equals("Ring")))
                {
                    fraction.Structures["TRAK"]["Ring"]["value"] = ps.Catheters.First<Catheter>(x => x.Id.Equals("Ring")).GetTotalDwellTime() / ttime;
                }

                double needleSum = 0.0;
                foreach (Catheter cat in cats.Where<Catheter>(x => !(new int[] { 1, 3 }.Contains(x.ChannelNumber))))
                {
                    needleSum += cat.GetTotalDwellTime();
                }
                fraction.Structures["TRAK"]["Needles"]["value"] = needleSum / ttime;

            }
            
            return fraction;
        }


        /// <summary>
        /// Fills first sheet of excel file with data from Fraktion list.
        /// </summary>
        /// <param name="outFile">Filename to open.</param>
        /// <param name="fractions">List of FraktionData Elements.</param>
        private static void FillExcel(string outFile, List<FractionData> fractions, string patientInfo)
        {
            // Open the document for editing.
            try
            {
                using (SpreadsheetDocument spreadSheet = SpreadsheetDocument.Open(@outFile, true))
                {
                    // Get a reference to the workbook part
                    WorkbookPart wbPart = spreadSheet.WorkbookPart;

                    // Search reference for first worksheet in sheets definitions in the workbook
                    Sheet theSheet = wbPart.Workbook.Descendants<Sheet>().FirstOrDefault();
                    if (theSheet == null)
                    {
                        throw new ArgumentException("sheetname");
                    }

                    WorksheetPart worksheetPart = (WorksheetPart)(wbPart.GetPartById(theSheet.Id));

                    // insert Patient Info in Worksheet
                    InsertText(patientInfo, "C", (uint)4, worksheetPart);

                    // Get a reference to the calculation chain
                    CalculationChainPart calculationChainPart = wbPart.CalculationChainPart;
                    CalculationChain calculationChain = calculationChainPart.CalculationChain;

                    // List of all elements of the calculation chain
                    var calculationCells = calculationChain.Elements<CalculationCell>().ToList();


                    // Loop over fractions
                    foreach (FractionData fraction in fractions)
                    {
                        if (fraction == null) { throw new Exception("No data in fraction"); }
                        string columnName = fraction.ColumnName;

                        bool nan_error = false;

                        // Insert cell data into the worksheet.
                        foreach (Dictionary<string, Dictionary<string, double>> StructureData in fraction.Structures.Values)
                        {
                            foreach (Dictionary<string, double> DataCells in StructureData.Values)
                            {
                                // Access cell
                                Cell cell = InsertCellInWorksheet(columnName, (uint)DataCells["row"], worksheetPart);

                                // test if cell contains formula, if so, remove the formula
                                if (cell.CellFormula != null)
                                {
                                    // delete formula from calc chain
                                    CalculationCell calculationCell = calculationCells.Where(c => c.CellReference == cell.CellReference).FirstOrDefault();
                                    calculationCell.Remove();
                                    // delete formula from workbook
                                    cell.CellFormula.Remove();
                                }

                                // write new cell values if not NaN
                                if (!Double.IsNaN(DataCells["value"]))
                                {
                                    cell.CellValue = new CellValue(DataCells["value"].ToString("##.000", CultureInfo.InvariantCulture));
                                    cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                                }
                                else
                                {
                                    nan_error = true;
                                }
                            }
                        }
                                                if (nan_error)
                        {
                            MessageBox.Show("Not all cell values were valid! Please check for consistency");
                        }
                    }

                    // apply changes
                    spreadSheet.WorkbookPart.Workbook.CalculationProperties.ForceFullCalculation = true;
                }
            }
            catch (Exception)
            {
                MessageBox.Show("Error while writing data, maybe the file is open somewhere else!");
            }
        }

    }


    /// <summary>
    /// Class that holds the relevant fraction data in a dictionary of dictionaries and fraction number.
    /// The fraction column is generated from the fraction number and only has a get method.
    /// Row numbers differ from original Excel Sheet! (https://www.estro.org/ESTRO/media/ESTRO/About/hdr-gyn_biol-physik-formular_2017.xlsx)
    /// </summary>
    public class FractionData
    {
        private string planId;
        private int fractionNumber;
        private string columnName;

        public int FractionNumber
        {
            get { return fractionNumber; }
            set
            {
                fractionNumber = value;
                FractionToColumn();
            }
        }

        public string ColumnName
        {
            get { return columnName; }
        }

        public string PlanId
        {
            get { return planId; }
            set 
            {
                planId = value;
                string strIndex = value.Last().ToString();
                FractionNumber = Convert.ToInt32(strIndex);
            }
        }

        /// <summary>
        /// Converts fraction number to column string.
        /// </summary>
        public string FractionToColumn()
        {
            columnName = ((char)(fractionNumber + 66)).ToString(); // converts number to letter (1=C, 2=D etc.)
            return columnName;
        }

        public FractionData(string id)
        {
            PlanId = id;
        }

        public Dictionary<string, Dictionary<string, Dictionary<string, double>>> Structures = new Dictionary<string, Dictionary<string, Dictionary<string, double>>>()
        {

            ["TRAK"] = new Dictionary<string, Dictionary<string, double>>()
            {
                ["Sum"] = new Dictionary<string, double>() { ["row"] = 24, ["value"] = 0.0 },
                ["Tandem"] = new Dictionary<string, double>() { ["row"] = 25, ["value"] = 0.0 },
                ["Ring"] = new Dictionary<string, double>() { ["row"] = 26, ["value"] = 0.0 },
                ["Needles"] = new Dictionary<string, double>() { ["row"] = 27, ["value"] = 0.0 },
            },
            ["DoseA"] = new Dictionary<string, Dictionary<string, double>>()
            {
                ["A_li"] = new Dictionary<string, double>() { ["row"] = 32, ["value"] = 0.0 },
                ["A_re"] = new Dictionary<string, double>() { ["row"] = 34, ["value"] = 0.0 },
                // Mean is calculated in excel.
                //["A_mean"] = new Dictionary<string, double>() { ["row"] = 36, ["value"] = 0.0 },
            },
            ["GTV"] = new Dictionary<string, Dictionary<string, double>>()
            {
                ["Vol"] = new Dictionary<string, double>() { ["row"] = 39, ["value"] = 0.0 },
                ["D98%"] = new Dictionary<string, double>() { ["row"] = 40, ["value"] = 0.0 },
            },
            ["CTV"] = new Dictionary<string, Dictionary<string, double>>()
            {
                ["Vol"] = new Dictionary<string, double>() { ["row"] = 43, ["value"] = 0.0 },
                ["D98%"] = new Dictionary<string, double>() { ["row"] = 44, ["value"] = 0.0 },
                ["D90%"] = new Dictionary<string, double>() { ["row"] = 46, ["value"] = 0.0 },
                ["D50%"] = new Dictionary<string, double>() { ["row"] = 48, ["value"] = 0.0 },
            },
            ["Bladder"] = new Dictionary<string, Dictionary<string, double>>()
            {
                ["Vol"] = new Dictionary<string, double>() { ["row"] = 51, ["value"] = 0.0 },
                ["D0.1cc"] = new Dictionary<string, double>() { ["row"] = 52, ["value"] = 0.0 },
                ["D2.0cc"] = new Dictionary<string, double>() { ["row"] = 54, ["value"] = 0.0 },
            },
            ["Rectum"] = new Dictionary<string, Dictionary<string, double>>()
            {
                ["Vol"] = new Dictionary<string, double>() { ["row"] = 57, ["value"] = 0.0 },
                ["D0.1cc"] = new Dictionary<string, double>() { ["row"] = 58, ["value"] = 0.0 },
                ["D2.0cc"] = new Dictionary<string, double>() { ["row"] = 60, ["value"] = 0.0 },
            },
            ["Sigmoid"] = new Dictionary<string, Dictionary<string, double>>()
            {
                ["Vol"] = new Dictionary<string, double>() { ["row"] = 63, ["value"] = 0.0 },
                ["D0.1cc"] = new Dictionary<string, double>() { ["row"] = 64, ["value"] = 0.0 },
                ["D2.0cc"] = new Dictionary<string, double>() { ["row"] = 66, ["value"] = 0.0 },
            },
            ["Bowel"] = new Dictionary<string, Dictionary<string, double>>()
            {
                ["Vol"] = new Dictionary<string, double>() { ["row"] = 63, ["value"] = 0.0 },
                ["D0.1cc"] = new Dictionary<string, double>() { ["row"] = 64, ["value"] = 0.0 },
                ["D2.0cc"] = new Dictionary<string, double>() { ["row"] = 66, ["value"] = 0.0 },
            },
        };
    }


}
