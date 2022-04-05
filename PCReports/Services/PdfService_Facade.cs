using PdfSharp.Pdf;
using PdfSharp.Drawing;
using PdfSharp.Pdf.AcroForms;
using PdfSharp.Pdf.IO;
using PdfSharp.Pdf.Security;
using PCReports.Models.Facade;

namespace PCReports.Services.Facade
{
    public class PdfService
    // 
    // This service creates and saves a pdf file of the input and output data.
    // 
    //     Steps:
    //
    //      1. copy the background form to a temperary location
    //      2. Get user to specify a filename
    //      3. Fill the form entries in the pdf File with the results
    //      4. Display the graphics data.
    //
    //
    {
        // --------- Main entery point ------------------------------------------------
        public PdfService()
        {
            _resourceFolderPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"Resources\");
        }

        private void AddReportGuidToFields(PdfDocument report, string reportGuid)
        {
            var fields = report.AcroForm.Fields;

            for (int i = 0; i < report.AcroForm.Fields.Count(); i++)
            {
                PdfAcroField field = report.AcroForm.Fields[i];
                string fieldName = $"{reportGuid}_{field.Name}";
                field.Elements.SetString("/T", fieldName);
            }
        }


        public string GenerateReport(string pdfFilePath, PDFResult pdfResult)
        {
            string reportGuid = Path.GetFileNameWithoutExtension(pdfFilePath);
            var source = string.Empty;

            // copy template to destination
            string source1 = _resourceFolderPath + @"templates\SPS Facade Structural report tempelate_Part1 - EN.pdf";
            string source2_transom = _resourceFolderPath + @"templates\SPS Facade Structural report tempelate_Part2_Transom - EN.pdf";
            string source2_mullion = _resourceFolderPath + @"templates\SPS Facade Structural report tempelate_Part2_Mullion - EN.pdf";
            string source3 = _resourceFolderPath + @"templates\SPS Facade Structural report tempelate_Part3 - EN.pdf";

            if (Thread.CurrentThread.CurrentCulture.Name == "de-DE")
            {
                source1 = _resourceFolderPath + @"templates\SPS Facade Structural report tempelate_Part1 - DE.pdf";
                source2_transom = _resourceFolderPath + @"templates\SPS Facade Structural report tempelate_Part2_Transom - DE.pdf";
                source2_mullion = _resourceFolderPath + @"templates\SPS Facade Structural report tempelate_Part2_Mullion - DE.pdf";
                source3 = _resourceFolderPath + @"templates\SPS Facade Structural report tempelate_Part3 - DE.pdf";
            }
            if (Thread.CurrentThread.CurrentCulture.Name == "fr-FR")
            {
                source1 = _resourceFolderPath + @"templates\SPS Facade Structural report tempelate_Part1 - FR.pdf";
                source2_transom = _resourceFolderPath + @"templates\SPS Facade Structural report tempelate_Part2_Transom - FR.pdf";
                source2_mullion = _resourceFolderPath + @"templates\SPS Facade Structural report tempelate_Part2_Mullion - FR.pdf";
                source3 = _resourceFolderPath + @"templates\SPS Facade Structural report tempelate_Part3 - FR.pdf";
            }

            string destination = CreateNewPdf(source1, pdfFilePath);
            // Create tempPDFfile for each member
            if (!string.IsNullOrEmpty(destination))
            {
                // open the Pdf file, fill general data
                PdfDocument myReport = PdfReader.Open(pdfFilePath, PdfDocumentOpenMode.Modify);
                PdfInputGeneralData(pdfResult, myReport);
                PdfDrawStructure(pdfResult, myReport);

                //PdfDocument source2Document = PdfReader.Open(source2, PdfDocumentOpenMode.Import);

                int memberCount = pdfResult.PDFMemberResults.Count();
                for (int i = 0; i < memberCount; i++)
                {
                    PDFMemberResult pdfMemberResult = pdfResult.PDFMemberResults[i];

                    // create temp file
                    string tempFileLoc = Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, $"Content\\temp\\tempPDF{Guid.NewGuid()}.pdf"));
                    string source2 = pdfMemberResult.MemberType == 5 ? source2_transom : source2_mullion;
                    string tempPDF = CreateNewPdf(source2, tempFileLoc);
                    PdfDocument tempDoc = PdfReader.Open(tempPDF, PdfDocumentOpenMode.Modify);
                    PdfInputMemberData(pdfResult, i + 1, tempDoc);

                    // draw load cases
                    PdfDrawLoadCaseImage(pdfMemberResult, tempDoc);

                    // draw profile
                    PdfDrawProfileImage(pdfMemberResult, tempDoc);

                    // draw result curves
                    PdfDrawResultCurves(pdfMemberResult, tempDoc);

                    // draw section and subsection title
                    //string subSectionTitle = $"6.{i+1} Result for Structural Member {pdfMemberResult.MemberID}";
                    PdfDrawSectionTitle(tempDoc.Pages[0], i);

                    // add report stamp to all fields
                    AddReportGuidToFields(tempDoc, reportGuid);

                    tempDoc.Save(tempFileLoc);
                    tempDoc.Close();
                    tempDoc = PdfReader.Open(tempFileLoc, PdfDocumentOpenMode.Import);
                    PdfPage tempPage = tempDoc.Pages[0];
                    myReport.AddPage(tempPage);
                    tempPage = tempDoc.Pages[1];
                    myReport.AddPage(tempPage);
                    //tempPage = tempDoc.Pages[2];
                    //myReport.AddPage(tempPage);
                    tempDoc.Close();
                    File.Delete(tempFileLoc);
                }
                //source2Document.Close();

                // source3Document
                PdfDocument source3Document = PdfReader.Open(source3, PdfDocumentOpenMode.Import);
                string tempFileLoc3 = Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, $"Content\\temp\\tempPDF{Guid.NewGuid()}.pdf"));
                string tempPDF3 = CreateNewPdf(source3, tempFileLoc3);
                PdfDocument tempDoc3 = PdfReader.Open(tempPDF3, PdfDocumentOpenMode.Modify);

                // fill appendix data
                PdfInputAppendixData(pdfResult, tempDoc3);

                // add report stamp to all fields
                AddReportGuidToFields(tempDoc3, reportGuid);

                tempDoc3.Save(tempFileLoc3);
                tempDoc3.Close();
                tempDoc3 = PdfReader.Open(tempFileLoc3, PdfDocumentOpenMode.Import);
                PdfPage tempPage3 = tempDoc3.Pages[0];
                myReport.AddPage(tempPage3);

                source3Document.Close();
                File.Delete(tempFileLoc3);

                // draw page number
                PdfDrawPageNumber(myReport);

                // add report stamp to all fields
                AddReportGuidToFields(myReport, reportGuid);

                // flatten and close the pdf file
                PdfSecuritySettings securitySettings = myReport.SecuritySettings;
                securitySettings.PermitFormsFill = false;

                // Restrict some rights.
                securitySettings.PermitAccessibilityExtractContent = false;
                securitySettings.PermitAnnotations = false;
                securitySettings.PermitAssembleDocument = false;
                securitySettings.PermitExtractContent = false;
                securitySettings.PermitFormsFill = false;
                securitySettings.PermitFullQualityPrint = true;
                securitySettings.PermitModifyDocument = false;
                securitySettings.PermitPrint = true;

                myReport.Flatten();
                myReport.Save(pdfFilePath);
                myReport.Close();
            }

            return destination;
        }

        public void PdfInputAppendixData(PDFResult pdfResult, PdfDocument myTemplate)
        {
            DateTime dateTime = DateTime.Today;
            var dateFormat = "dd. MMM. yyyy"; //german format
            //Thread.CurrentThread.CurrentCulture = currentUser.DefaultLanguage;
            if (Thread.CurrentThread.CurrentCulture.Name.Equals("en-US"))
            {
                dateFormat = "MMM. dd. yyyy"; //english format
            }

            InsertValue(myTemplate, "ProjectName", pdfResult.ProjectName);
            InsertValue(myTemplate, "Location", pdfResult.Location);
            InsertValue(myTemplate, "Date", DateTime.Now.ToString(dateFormat));
            InsertValue(myTemplate, "User", pdfResult.UserName);

            //List<string> section_list = new List<string>();

            //int trackNo = 0;
            List<FacadeSection> sectionToList = pdfResult.FacadeSections.GroupBy(x => x.ArticleName).Select(grp => grp.First()).ToList();
            for (int i = 0; i < sectionToList.Count; i++)
            {
                int secType = sectionToList[i].SectionType;

                InsertValue(myTemplate, $"Section.{i}", sectionToList[i].ArticleName);

                //General information
                if (Math.Abs(sectionToList[i].BTDepth - 0) < 0.001)
                {
                    InsertValue(myTemplate, $"Depth.{i}", "--");
                }
                else
                {
                    InsertValue(myTemplate, $"Depth.{i}", $"{sectionToList[i].BTDepth / 10,8:0.000}");  //convert mm to cm
                }

                if (Math.Abs(sectionToList[i].Width - 0) < 0.001)
                {
                    InsertValue(myTemplate, $"Width.{i}", "--");
                }
                else
                {
                    InsertValue(myTemplate, $"Width.{i}", $"{sectionToList[i].Width / 10,8:0.000}");  //convert mm to cm
                }

                if (Math.Abs(sectionToList[i].A - 0) < 0.001)
                {
                    InsertValue(myTemplate, $"Area.{i}", "--");
                }
                else
                {
                    InsertValue(myTemplate, $"Area.{i}", $"{sectionToList[i].A / 100,8:0.000}");   //convert mm2 to cm2
                }


                // Gyration
                if (Math.Abs(sectionToList[i].Ry - 0) < 0.001)
                {
                    InsertValue(myTemplate, $"rx.{i}", "--");
                }
                else
                {
                    InsertValue(myTemplate, $"rx.{i}", $"{sectionToList[i].Ry / 10,8:0.000}");   //convert mm to cm
                }

                if (Math.Abs(sectionToList[i].Rz - 0) < 0.001)
                {
                    InsertValue(myTemplate, $"ry.{i}", "--");
                }
                else
                {
                    InsertValue(myTemplate, $"ry.{i}", $"{sectionToList[i].Rz / 10,8:0.000}");   //convert mm to cm
                }

                // moment of inertia
                if (Math.Abs(sectionToList[i].Iyy - 0) < 0.001)
                {
                    InsertValue(myTemplate, $"Ix.{i}", "--");
                }
                else
                {
                    InsertValue(myTemplate, $"Ix.{i}", $"{sectionToList[i].Iyy / 10000,8:0.000}");   //convert mm4 to cm4
                }

                if (Math.Abs(sectionToList[i].Izz - 0) < 0.001)
                {
                    InsertValue(myTemplate, $"Iy.{i}", "--");
                }
                else
                {
                    InsertValue(myTemplate, $"Iy.{i}", $"{sectionToList[i].Izz / 10000,8:0.000}");   //convert mm4 to cm4
                }

                // section modulus
                if (Math.Abs(sectionToList[i].Wyp - 0) < 0.001)
                {
                    InsertValue(myTemplate, $"Spx.{i}", "--");
                }
                else
                {
                    InsertValue(myTemplate, $"Spx.{i}", $"{sectionToList[i].Wyp / 1000,8:0.000}");  //convert mm3 to cm3
                }

                if (Math.Abs(sectionToList[i].Wyn - 0) < 0.001)
                {
                    InsertValue(myTemplate, $"Snx.{i}", "--");
                }
                else
                {
                    InsertValue(myTemplate, $"Snx.{i}", $"{sectionToList[i].Wyn / 1000,8:0.000}"); //convert mm3 to cm3
                }

                if (Math.Abs(sectionToList[i].Wzp - 0) < 0.001)
                {
                    InsertValue(myTemplate, $"Spy.{i}", "--");
                }
                else
                {
                    InsertValue(myTemplate, $"Spy.{i}", $"{sectionToList[i].Wzp / 1000,8:0.000}"); //convert mm3 to cm3
                }

                if (Math.Abs(sectionToList[i].Wzn - 0) < 0.001)
                {
                    InsertValue(myTemplate, $"Sny.{i}", "--");
                }
                else
                {
                    InsertValue(myTemplate, $"Sny.{i}", $"{sectionToList[i].Wzn / 1000,8:0.000}"); //convert mm3 to cm3
                }

                // Torsional constants
                if (Math.Abs(sectionToList[i].J - 0) < 0.001)
                {
                    InsertValue(myTemplate, $"J.{i}", "--");
                }
                else
                {
                    InsertValue(myTemplate, $"J.{i}", $"{sectionToList[i].J / 10000,8:0.000}");    //convert mm4 to cm4
                }

                if (Math.Abs(sectionToList[i].Cw - 0) < 0.001)
                {
                    InsertValue(myTemplate, $"Cw.{i}", "--");
                }
                else
                {
                    InsertValue(myTemplate, $"Cw.{i}", $"{sectionToList[i].Cw / 1000000,8:0.000}");   //convert mm6 to cm6
                }

                if (Math.Abs(sectionToList[i].Beta_torsion - 0) < 0.001)
                {
                    InsertValue(myTemplate, $"Betax.{i}", "--");
                }
                else
                {
                    InsertValue(myTemplate, $"Betax.{i}", $"{sectionToList[i].Beta_torsion / 10,4:0.00}"); //convert mm to cm
                }


                // Shear center
                InsertValue(myTemplate, $"Xs.{i}", $"{sectionToList[i].Ys / 10,4:0.00}");  //convert mm to cm
                InsertValue(myTemplate, $"Ys.{i}", $"{sectionToList[i].Zs / 10,4:0.00}");  //convert mm to cm


                // Plastic property
                if (Math.Abs(sectionToList[i].Zy - 0) < 0.001)
                {
                    InsertValue(myTemplate, $"Zx.{i}", "--");
                }
                else
                {
                    InsertValue(myTemplate, $"Zx.{i}", $"{sectionToList[i].Zy / 1000,8:0.000}");  //convert mm3 to cm3
                }

                if (Math.Abs(sectionToList[i].Zz - 0) < 0.001)
                {
                    InsertValue(myTemplate, $"Zy.{i}", "--");
                }
                else
                {
                    InsertValue(myTemplate, $"Zy.{i}", $"{sectionToList[i].Zz / 1000,8:0.000}");  //convert mm3 to cm3
                }

            }

        }


        public void PdfDrawLoadCaseImage(PDFMemberResult pdfMemberResult, PdfDocument myTemplate)
        {
            bool isTransom = pdfMemberResult.MemberType == 5;

            // set y location of each drawing
            bool existHorizontalLiveLoad = true;
            if ((Math.Abs(pdfMemberResult.ReactionForces[2] - 0) < 0.0001) && (Math.Abs(pdfMemberResult.ReactionForces[3] - 0) < 0.0001))
            {
                existHorizontalLiveLoad = false;
            }

            bool existDeadLoad = false;
            if (pdfMemberResult.MemberType == 5)
            {
                existDeadLoad = true;
            }

            double y1 = 330;
            double y2 = 450;
            double y3 = 550;
            double y4 = 610;
            double y5 = 690;

            if (existHorizontalLiveLoad && existDeadLoad == false)
            {
                y1 = 330;
                y2 = 480;
                y3 = 600;
                y4 = 690;
            }

            if (existHorizontalLiveLoad == false && existDeadLoad)
            {
                y1 = 330;
                y2 = 480;
                y4 = 580;
                y5 = 680;
            }

            if (existHorizontalLiveLoad == false && existDeadLoad == false)
            {
                y1 = 330;
                y2 = 510;
                y4 = 650;
            }

            PdfDrawBeam(pdfMemberResult, myTemplate, 1, y1);
            PdfDrawLoadCaseImageTitle(myTemplate, isTransom, 1, y1);
            PdfDrawPinSupport(pdfMemberResult, myTemplate, 1, y1);
            PdfDrawLoads(pdfMemberResult, myTemplate, 1, y1);
            PdfDrawCoordinate(pdfMemberResult, myTemplate, 1, y1);

            PdfDrawBeam(pdfMemberResult, myTemplate, 2, y2);
            PdfDrawLoadCaseImageTitle(myTemplate, isTransom, 2, y2);
            PdfDrawPinSupport(pdfMemberResult, myTemplate, 2, y2);
            PdfDrawLoads(pdfMemberResult, myTemplate, 2, y2);

            if (existHorizontalLiveLoad)
            {
                PdfDrawBeam(pdfMemberResult, myTemplate, 3, y3);
                PdfDrawLoadCaseImageTitle(myTemplate, isTransom, 3, y3); //use chartNo=3 to draw support for reaction force
                PdfDrawPinSupport(pdfMemberResult, myTemplate, 3, y3);
                PdfDrawLoads(pdfMemberResult, myTemplate, 3, y3);
            }

            PdfDrawLoadCaseImageTitle(myTemplate, isTransom, 4, y4); //use chartNo=4 to draw support for reaction force
            PdfDrawPinSupport(pdfMemberResult, myTemplate, 4, y4);
            PdfDrawReactions(pdfMemberResult, myTemplate.Pages[0], 4, y4);

            if (existDeadLoad)
            {
                PdfDrawBeam(pdfMemberResult, myTemplate, 5, y5);
                PdfDrawLoadCaseImageTitle(myTemplate, isTransom, 5, y5); //use chartNo=4 to draw support for reaction force
                PdfDrawPinSupport(pdfMemberResult, myTemplate, 5, y5);
                PdfDrawLoads(pdfMemberResult, myTemplate, 5, y5);
                PdfDrawCoordinate(pdfMemberResult, myTemplate, 5, y5);
                PdfDrawReactions(pdfMemberResult, myTemplate.Pages[0], 5, y5);
            }
        }



        // ---------- I/O routine ------------------------------------------------------
        public void PdfInputGeneralData(PDFResult pdfResult, PdfDocument myTemplate)
        {
            DateTime dateTime = DateTime.Today;
            var dateFormat = "dd. MMM. yyyy"; //german format
            //Thread.CurrentThread.CurrentCulture = currentUser.DefaultLanguage;
            if (Thread.CurrentThread.CurrentCulture.Name.Equals("en-US"))
            {
                dateFormat = "MMM. dd. yyyy"; //english format
            }
            // insert values
            InsertValue(myTemplate, "CoverPageProjectName", pdfResult.ProjectName);
            InsertValue(myTemplate, "CoverPageLocation", pdfResult.Location);
            InsertValue(myTemplate, "CoverPageConfiguration", pdfResult.ConfigurationName);
            InsertValue(myTemplate, "CoverPageDate", DateTime.Now.ToString(dateFormat));
            InsertValue(myTemplate, "UserNotes", pdfResult.UserNotes);

            InsertValue(myTemplate, "ProjectName", pdfResult.ProjectName);
            InsertValue(myTemplate, "Location", pdfResult.Location);
            InsertValue(myTemplate, "Date", DateTime.Now.ToString(dateFormat));
            InsertValue(myTemplate, "User", pdfResult.UserName);

            InsertValue(myTemplate, "ProfileSystem", pdfResult.ProfileSystem);

            // Transom profile
            if (!(string.IsNullOrEmpty(pdfResult.TransomProfile)))
            {
                InsertValue(myTemplate, "TransomProfile", $"{pdfResult.TransomProfile,3:0.##}");
                InsertValue(myTemplate, "TransomProfileWeight", $"{pdfResult.TransomProfileWeight * 9.81 * 1000,3:0.0#} N/m"); // input kg/mm
            }
            else
            {
                InsertValue(myTemplate, "TransomProfile", "--");
                InsertValue(myTemplate, "TransomProfileWeight", "--");
            }

            // Major Mullion profile
            if (!(string.IsNullOrEmpty(pdfResult.MajorMullionProfile)))
            {
                InsertValue(myTemplate, "MajorMullionProfile", $"{pdfResult.MajorMullionProfile,3:0.##}");
                InsertValue(myTemplate, "MajorMullionProfileWeight", $"{pdfResult.MajorMullionProfileWeight * 9.81 * 1000,3:0.0#} N/m");
            }
            else
            {
                InsertValue(myTemplate, "MajorMullionProfile", "--");
                InsertValue(myTemplate, "MajorMullionProfileWeight", "--");
            }

            // Minor Mullion Profile
            if (!(string.IsNullOrEmpty(pdfResult.MinorMullionProfile)))
            {
                InsertValue(myTemplate, "MinorMullionProfile", $"{pdfResult.MinorMullionProfile,3:0.##}");
                InsertValue(myTemplate, "MinorMullionProfileWeight", $"{pdfResult.MinorMullionProfileWeight * 9.81 * 1000,3:0.0#} N/m");
            }
            else
            {
                InsertValue(myTemplate, "MinorMullionProfile", "--");
                InsertValue(myTemplate, "MinorMullionProfileWeight", "--");
            }

            // block distance
            InsertValue(myTemplate, "BlockDistance", $"{pdfResult.BlockDistance,3:0.##} mm");


            int showGlassCount = Math.Min(pdfResult.GlassTypes.Count(), 5);
            for (int i = 0; i < showGlassCount; i++)
            {
                InsertValue(myTemplate, $"Glass.{i}", $"{pdfResult.GlassTypes[i].GlassIDs}        {pdfResult.GlassTypes[i].Weight}    {pdfResult.GlassTypes[i].Description}");
            }
            if (showGlassCount < pdfResult.GlassTypes.Count())
            {
                InsertValue(myTemplate, $"Glass.Overflow", "...");
            }

            // Section 2
            InsertValue(myTemplate, "WindLoad", $"{pdfResult.WindLoad,5:0.00#}");
            InsertValue(myTemplate, "CpeString", pdfResult.CpeString);
            InsertValue(myTemplate, "pCpiString", pdfResult.pCpiString);
            InsertValue(myTemplate, "nCpiString", pdfResult.nCpiString);
            InsertValue(myTemplate, "HorizontalLiveLoad", $"{pdfResult.HorizontalLiveLoad,4:0.00}");
            InsertValue(myTemplate, "HorizontalLiveLoadHeight", $"{pdfResult.HorizontalLiveLoadHeight,4:0.##}");

            InsertValue(myTemplate, "WindLoadFactor", $"{pdfResult.WindLoadFactor,4:0.00}");
            InsertValue(myTemplate, "HorizontalLiveLoadFactor", $"{pdfResult.HorizontalLiveLoadFactor,4:0.00}");
            InsertValue(myTemplate, "DeadLoadFactor", $"{1.35,4:0.00}");
            // Section 3
            string codeType = "";

            if (pdfResult.AllowableDeflectionLine1 == "DIN EN 13830:2003.")
            {
                codeType = "DIN (EN) 13830, Curtain wall standard: 2003-09";
            }
            else if (pdfResult.AllowableDeflectionLine1 == "DIN EN 13830:2015/2020.")
            {
                codeType = "DIN (EN) 13830, Curtain wall standard: 2015-07";
            }
            else if (pdfResult.AllowableDeflectionLine1 == "DIN EN 14351-1-2016 Class B." || pdfResult.AllowableDeflectionLine1 == "DIN EN 14351-1-2016 Class C.")
            {
                codeType = "DIN (EN) 14351-1, Windows and doors_Product standard, performance characteristics, 2016";
            }
            else if (pdfResult.AllowableDeflectionLine1 == "US ASCE.")
            {
                codeType = "ASCE/SEI 7-16, Minimum Design Loads andAssociated Criteria forBuildings and Other Structures";
            }
            else if (pdfResult.AllowableDeflectionLine1 == "user defined.")
            {
                codeType = "Permissible deflection is user defined";
            }
            else
            {
                codeType = "Design code not found";
            }
            InsertValue(myTemplate, "UDcode", $"{codeType}");

            // Section 4
            InsertValue(myTemplate, "AllowableDeflectionLine1", $"{pdfResult.AllowableDeflectionLine1}");
            InsertValue(myTemplate, "AllowableDeflectionLine2", $"{pdfResult.AllowableDeflectionLine2}");
            double inplaneDispIndex = pdfResult.PDFMemberResults.FirstOrDefault(x => x.MemberType == 5).VerticalDispIndex;
            InsertValue(myTemplate, "AllowableInplaneDeflectionLine", $"{pdfResult.AllowableInplaneDeflectionLine}");
            // Section 5
            InsertValue(myTemplate, "Alloys", pdfResult.Alloys);
            InsertValue(myTemplate, "Beta", $"{pdfResult.Beta,4:0.00} MPa");
            if (Math.Abs(pdfResult.AluminumReinforcementBeta - 0) < 0.00001)
            {
                InsertValue(myTemplate, "AluminumReinforcementBeta", "  --");
            }
            else
            {
                InsertValue(myTemplate, "AluminumReinforcementBeta", $"{pdfResult.AluminumReinforcementBeta,4:0.00} MPa");
            }
            if (Math.Abs(pdfResult.SteelReinforcementBeta - 0) < 0.00001)
            {
                InsertValue(myTemplate, "SteelReinforcementBeta", "  --");
            }
            else
            {
                InsertValue(myTemplate, "SteelReinforcementBeta", $"{pdfResult.SteelReinforcementBeta,4:0.00} MPa");
            }
        }


        public void PdfInputMemberData(PDFResult pdfResult, int subSectionNo, PdfDocument myTemplate)
        {
            var dateFormat = "dd MMM yyyy"; //german format
            //Thread.CurrentThread.CurrentCulture = currentUser.DefaultLanguage;
            if (Thread.CurrentThread.CurrentCulture.Name.Equals("en-US"))
            {
                dateFormat = "MMM dd yyyy"; //english format
            }
            PDFMemberResult pdfMemberResult = pdfResult.PDFMemberResults[subSectionNo - 1];
            InsertValue(myTemplate, "MemberID", $"{pdfMemberResult.MemberID}");
            InsertValue(myTemplate, "ArticleName", $"{pdfMemberResult.ArticleName}");

            InsertValue(myTemplate, "Depth", $"{pdfMemberResult.Depth,4:0}");
            InsertValue(myTemplate, "Length", $"{pdfMemberResult.Length,4:0}");

            InsertValue(myTemplate, "CrossArea", $"{pdfMemberResult.A / 100.0,4:0.00}"); // convert from mm2 to cm2
            InsertValue(myTemplate, "Ix", $"{pdfMemberResult.Iyy / 10000.0,4:0.00}");  // the Ix in the report is the strong axis and corresponds to the Iyy in the back calculation
            InsertValue(myTemplate, "Iy", $"{pdfMemberResult.Izz / 10000.0,4:0.00}");  // the Iy in the report is the weak axis and corresponds to the Izz in the back calculation

            if (!string.IsNullOrEmpty(pdfMemberResult.ReinfArticleName))
            {
                InsertValue(myTemplate, "Reinf.No", $"{pdfMemberResult.ReinfArticleName,4:0.00}");
            }
            else
            {
                InsertValue(myTemplate, "Reinf.No", "  --");
            }

            if (Math.Abs(pdfMemberResult.ReinfWidth - 0) > 0.0001)
            {
                InsertValue(myTemplate, "Reinf.Width", $"{pdfMemberResult.ReinfWidth,4:0}");
            }
            else
            {
                InsertValue(myTemplate, "Reinf.Width", "  --");
            }

            if (Math.Abs(pdfMemberResult.ReinfDepth - 0) > 0.0001)
            {
                InsertValue(myTemplate, "Reinf.Depth", $"{pdfMemberResult.ReinfDepth,4:0}");
            }
            else
            {
                InsertValue(myTemplate, "Reinf.Depth", "  --");
            }

            if (Math.Abs(pdfMemberResult.ReinfA - 0) > 0.0001)
            {
                InsertValue(myTemplate, "Reinf.Area", $"{pdfMemberResult.ReinfA / 100.0,4:0.00}");
            }
            else
            {
                InsertValue(myTemplate, "Reinf.Area", "  --");
            }

            if (Math.Abs(pdfMemberResult.ReinfIyy - 0) > 0.0001)
            {
                InsertValue(myTemplate, "Reinf.Ix", $"{pdfMemberResult.ReinfIyy / 10000.0,4:0.00}");  // the Ix in the report is the strong axis and corresponds to the Iyy in the back calculation
            }
            else
            {
                InsertValue(myTemplate, "Reinf.Ix", "  --");
            }

            if (Math.Abs(pdfMemberResult.ReinfIzz - 0) > 0.0001)
            {
                InsertValue(myTemplate, "Reinf.Iy", $"{pdfMemberResult.ReinfIzz / 10000.0,4:0.00}");  // the Iy in the report is the weak axis and corresponds to the Izz in the back calculation
            }
            else
            {
                InsertValue(myTemplate, "Reinf.Iy", "  --");
            }

            InsertValue(myTemplate, "TributaryArea", $"{pdfMemberResult.TributaryArea / 1000000,3:0.00}");  // convert from mm2 to m2
            InsertValue(myTemplate, "Cp", $"{pdfMemberResult.Cp,5:0.00#}");
            InsertValue(myTemplate, "MemberWindLoad", $"{pdfMemberResult.AppliedWindLoad,5:0.00#}");

            InsertValue(myTemplate, "ProjectName", pdfResult.ProjectName);
            InsertValue(myTemplate, "Location", pdfResult.Location);
            InsertValue(myTemplate, "Date", DateTime.Now.ToString(dateFormat));
            InsertValue(myTemplate, "User", pdfResult.UserName);

            // Utilization convert unit
            double[,] UtilCheckTable = new double[2, 8];

            for (int i = 0; i < 2; i++)
            {
                for (int j = 0; j <= 7; j++)
                {
                    if (j == 0 || j == 2 || j == 4 || j == 6)
                    {
                        UtilCheckTable[i, j] = pdfMemberResult.UtilizationCheckTable[i * 8 + j] / 1000.0;  // convert mm3 to cm3
                    }
                    else if (j == 1 || j == 3 || j == 5)
                    {
                        UtilCheckTable[i, j] = pdfMemberResult.UtilizationCheckTable[i * 8 + j] / 10000.0;  // convert mm4 to cm4
                    }
                    else
                    {
                        UtilCheckTable[i, j] = pdfMemberResult.UtilizationCheckTable[i * 8 + j];  // N/mm2 
                    }
                }
            }

            // Utilization check
            for (int j = 0; j <= 7; j++)
            {
                for (int i = 0; i < 2; i++)
                {
                    if (Math.Abs(UtilCheckTable[i, j]) > 0.0001)
                    {
                        InsertValue(myTemplate, $"UtilzationCheck.{i}.{j}", $"{UtilCheckTable[i, j],4:0.00}");
                    }
                    else
                    {
                        InsertValue(myTemplate, $"UtilzationCheck.{i}.{j}", $"--");
                    }
                }
                for (int i = 2; i < 3; i++)
                {
                    if (Math.Abs(pdfMemberResult.UtilizationCheckTable[i * 8 + j] - 0) > 0.0001)
                    {
                        if (pdfMemberResult.UtilizationCheckTable[i * 8 + j] > 1)
                        {
                            InsertValue(myTemplate, $"UtilzationCheck.{i}.{j}", $"{pdfMemberResult.UtilizationCheckTable[i * 8 + j],0:P1}", true);
                        }
                        else
                        {
                            InsertValue(myTemplate, $"UtilzationCheck.{i}.{j}", $"{pdfMemberResult.UtilizationCheckTable[i * 8 + j],0:P1}");
                        }
                    }
                    else
                    {
                        InsertValue(myTemplate, $"UtilzationCheck.{i}.{j}", $"--");
                    }
                }


                // status check
                if (Math.Abs(pdfMemberResult.UtilizationCheckTable[16+j]) < 0.0001 || Double.IsNaN(pdfMemberResult.UtilizationCheckTable[16+j]))
                {
                    InsertValue(myTemplate, $"StatCheck.{j}", "--", false);
                }
                else
                {
                    if (pdfMemberResult.UtilizationCheckTable[16+j] > 1)
                    {
                        if (Thread.CurrentThread.CurrentCulture.Name == "de-DE")
                        {
                            InsertValue(myTemplate, $"StatCheck.{j}", "Scheitern", true);
                        }
                        else if (Thread.CurrentThread.CurrentCulture.Name == "fr-FR")
                        {
                            InsertValue(myTemplate, $"StatCheck.{j}", "non OK", true);
                        }
                        else
                        {
                            InsertValue(myTemplate, $"StatCheck.{j}", "Fail", true);
                        }

                    }
                    else
                    {
                        InsertValue(myTemplate, $"StatCheck.{j}", "OK", false);
                    }
                }

            }

            // deflection check table
            // title
            if (pdfMemberResult.MemberType == 5)
            {
                if (Thread.CurrentThread.CurrentCulture.Name == "de-DE")
                {
                    InsertValue(myTemplate, "Field", "Verschiebung", false);
                }
                else if (Thread.CurrentThread.CurrentCulture.Name == "fr-FR")
                {
                    InsertValue(myTemplate, "Field", "Déplacement", false);
                }
                else
                {
                    InsertValue(myTemplate, "Field", "Displacement", false);
                }

                InsertValue(myTemplate, "No.0", "\u03B4y", false);
                InsertValue(myTemplate, "No.1", "\u03B4x", false);
            }
            else
            {
                if (Thread.CurrentThread.CurrentCulture.Name == "de-DE")
                {
                    InsertValue(myTemplate, "Field", "Feld", false);
                }
                else if (Thread.CurrentThread.CurrentCulture.Name == "fr-FR")
                {
                    InsertValue(myTemplate, "Field", "Travée", false);
                }
                else
                {
                    InsertValue(myTemplate, "Field", "Field", false);
                }

                for (int j = 0; j < 4; j++)
                {
                    if (Math.Abs(pdfMemberResult.DeflectionCheckTable[j] - 0) < 0.0001)
                    {
                        InsertValue(myTemplate, $"No.{j}", "--", false);
                    }
                    else
                    {
                        InsertValue(myTemplate, $"No.{j}", $"{j + 1,2:0.}", false);
                    }
                }
            }
            // values
            for (int j = 0; j < 4; j++)
            {
                if (Math.Abs(pdfMemberResult.DeflectionCheckTable[j] - 0) < 0.0001)
                {
                    InsertValue(myTemplate, $"DeflectionCheck.{0}.{j}", "--", false);
                    InsertValue(myTemplate, $"DeflectionCheck.{1}.{j}", "--", false);
                    InsertValue(myTemplate, $"DeflectionCheck.{2}.{j}", "--", false);
                }
                else
                {
                    InsertValue(myTemplate, $"DeflectionCheck.{0}.{j}", $"{pdfMemberResult.DeflectionCheckTable[j],4:0.}");
                    InsertValue(myTemplate, $"DeflectionCheck.{1}.{j}", $"{pdfMemberResult.DeflectionCheckTable[5+j],4:0.00}");
                    InsertValue(myTemplate, $"DeflectionCheck.{2}.{j}", $"{pdfMemberResult.DeflectionCheckTable[10+j],4:0.00}");
                }
                // status
                if (Math.Abs(pdfMemberResult.DeflectionCheckTable[5+j] - 0) < 0.0001 && Math.Abs(pdfMemberResult.DeflectionCheckTable[10+j] - 0) < 0.0001)
                {
                    InsertValue(myTemplate, $"DeflectionCheck.{3}.{j}", "--", false);
                }
                else
                {
                    if (pdfMemberResult.DeflectionCheckTable[10+j] < pdfMemberResult.DeflectionCheckTable[5+j])
                    {
                        if (Thread.CurrentThread.CurrentCulture.Name == "de-DE")
                        {
                            InsertValue(myTemplate, $"DeflectionCheck.{3}.{j}", "Scheitern", true);
                        }
                        else if (Thread.CurrentThread.CurrentCulture.Name == "fr-FR")
                        {
                            InsertValue(myTemplate, $"DeflectionCheck.{3}.{j}", "non OK", true);
                        }
                        else
                        {
                            InsertValue(myTemplate, $"DeflectionCheck.{3}.{j}", "Fail", true);
                        }

                    }
                    else
                    {
                        InsertValue(myTemplate, $"DeflectionCheck.{3}.{j}", "OK", false);
                    }
                }

            }

            // Load case
            int loadCaseNum = pdfMemberResult.HLLWorstLoadCase.Length;
            int validLoadCaseNum;

            string LC1string, LC2string, LC3string, LCFieldString2;
            string LC1LiveLoad, LC2LiveLoad, LC1WindLoad, LC2WindLoad, LC3DeadLoad;
            string[] LCstringField;
            // get the number of non-zero load case
            validLoadCaseNum = 0;
            for (int i = 0; i < loadCaseNum; i++)
            {
                int loadCase = pdfMemberResult.HLLWorstLoadCase[i];
                if (loadCase != 0)
                {
                    validLoadCaseNum++;
                }
            }
            // Convert the valid load case number to field string
            LCstringField = new string[validLoadCaseNum];
            int jj = 0;
            for (int i = 0; i < loadCaseNum; i++)
            {
                int loadCase = pdfMemberResult.HLLWorstLoadCase[i];
                if (loadCase != 0)
                {
                    if (Thread.CurrentThread.CurrentCulture.Name == "de-DE")
                    {
                        LCstringField[jj] = "Feld " + (i + 1).ToString();
                        continue;
                    }
                    else if (Thread.CurrentThread.CurrentCulture.Name == "fr-FR")
                    {
                        LCstringField[jj] = "Travée " + (i + 1).ToString();
                        continue;
                    }
                    LCstringField[jj] = "Field " + (i + 1).ToString();
                    jj++;
                }
            }
            // Combine all the field string into one string
            LCFieldString2 = "";
            for (int i = 0; i < LCstringField.Length; i++)
            {
                if (i == 0)
                {
                    LCFieldString2 = LCstringField[i];
                }
                else
                {
                    LCFieldString2 += "+" + LCstringField[i];
                }
            }

            if (string.IsNullOrEmpty(LCFieldString2))
            {
                LC1LiveLoad = "";
                LC2LiveLoad = "";
            }
            else
            {
                LC1LiveLoad = " + " + pdfResult.HorizontalLiveLoadFactor + "*0.7*Live load (" + LCFieldString2 + ")";
                LC2LiveLoad = " + " + pdfResult.HorizontalLiveLoadFactor + "*Live load (" + LCFieldString2 + ")";

                if (Thread.CurrentThread.CurrentCulture.Name == "de-DE")
                {
                    LC1LiveLoad = " + " + pdfResult.HorizontalLiveLoadFactor + "*0,7*Nutzlast (" + LCFieldString2 + ")";
                    LC2LiveLoad = " + " + pdfResult.HorizontalLiveLoadFactor + "*Nutzlast (" + LCFieldString2 + ")";
                }
                else if (Thread.CurrentThread.CurrentCulture.Name == "fr-FR")
                {
                    LC1LiveLoad = " + " + pdfResult.HorizontalLiveLoadFactor + "*0,7*Charge d'exploitation (" + LCFieldString2 + ")";
                    LC2LiveLoad = " + " + pdfResult.HorizontalLiveLoadFactor + "*Charge d'exploitation (" + LCFieldString2 + ")";
                }
            }

            LC1WindLoad = pdfResult.WindLoadFactor + "*Wind load";
            LC2WindLoad = pdfResult.WindLoadFactor + "*0.6" + "*Wind load";
            LC3DeadLoad = pdfResult.DeadLoadFactor + "*Dead load";

            if (Thread.CurrentThread.CurrentCulture.Name == "de-DE")
            {
                LC1WindLoad = pdfResult.WindLoadFactor + "*Windlast";
                LC2WindLoad = pdfResult.WindLoadFactor + "*0,6" + "*Windlast";
                LC3DeadLoad = pdfResult.DeadLoadFactor + "Eigengewicht";
            }
            else if (Thread.CurrentThread.CurrentCulture.Name == "fr-FR")
            {
                LC1WindLoad = pdfResult.WindLoadFactor + "*Charge de vent";
                LC2WindLoad = pdfResult.WindLoadFactor + "*0,6" + "*Charge de vent";
                LC3DeadLoad = pdfResult.DeadLoadFactor + "Dead load";
            }

            // construct all LC strings
            LC1string = "Max moment with LC1: " + $"{pdfMemberResult.MaxMyLC1 / 10000,3:0.##}" + "kN*cm " + "(" + LC1WindLoad + LC1LiveLoad + ")";
            LC2string = "Max moment with LC2: " + $"{pdfMemberResult.MaxMyLC2 / 10000,3:0.##}" + "kN*cm " + "(" + LC2WindLoad + LC2LiveLoad + ")";
            LC3string = "Max moment with LC3: " + $"{pdfMemberResult.MaxMzLC3 / 10000,3:0.##}" + "kN*cm " + "(" + LC3DeadLoad + ")";

            if (Thread.CurrentThread.CurrentCulture.Name == "de-DE")
            {
                LC1string = "Maximaler moment mit LC1: " + $"{pdfMemberResult.MaxMyLC1 / 10000,3:0.##}" + "kN*cm " + "(" + LC1WindLoad + LC1LiveLoad + ")";
                LC2string = "Maximaler moment mit LC2: " + $"{pdfMemberResult.MaxMyLC2 / 10000,3:0.##}" + "kN*cm " + "(" + LC2WindLoad + LC2LiveLoad + ")";
                LC3string = "Maximaler moment mit LC3: " + $"{pdfMemberResult.MaxMzLC3 / 10000,3:0.##}" + "kN*cm " + "(" + LC3DeadLoad + ")";
            }
            else if (Thread.CurrentThread.CurrentCulture.Name == "fr-FR")
            {
                LC1string = "Moment maximum avec LC1: " + $"{pdfMemberResult.MaxMyLC1 / 10000,3:0.##}" + "kN*cm " + "(" + LC1WindLoad + LC1LiveLoad + ")";
                LC2string = "Moment maximum avec LC2: " + $"{pdfMemberResult.MaxMyLC2 / 10000,3:0.##}" + "kN*cm " + "(" + LC2WindLoad + LC2LiveLoad + ")";
                LC3string = "Moment maximum avec LC3: " + $"{pdfMemberResult.MaxMzLC3 / 10000,3:0.##}" + "kN*cm " + "(" + LC3DeadLoad + ")";
            }

            InsertValue(myTemplate, "LC1", LC1string);
            InsertValue(myTemplate, "LC2", LC2string);
            if (pdfMemberResult.MemberType == 5) InsertValue(myTemplate, "LC3", LC3string);

            PdfChangeFieldName(myTemplate, subSectionNo);
        }

        public void PdfDrawStructure(PDFResult pdfResult, PdfDocument myTemplate, int pageNumber = 1, double xLocation = 77, double yLocation = 290, double ImageHeight = 300, double ImageWidth = 430)
        {
            XGraphics g = XGraphics.FromPdfPage(myTemplate.Pages[pageNumber]);
            XPen grayPen = new XPen(XColor.FromArgb(38, 38, 38), .25);
            XPen grayDashPen = new XPen(XColor.FromArgb(38, 38, 38), .6);
            grayDashPen.DashStyle = XDashStyle.DashDot;
            grayDashPen.DashPattern = new double[4] { 8, 3, 2, 3 };
            XPen blackPen = new XPen(XColor.FromArgb(0, 0, 0), 0.75);
            XBrush greenBrush = new XSolidBrush(XColor.FromArgb(120, 185, 40));
            XBrush grayBrush = new XSolidBrush(XColor.FromArgb(102, 102, 102));
            XBrush blackBrush = new XSolidBrush(XColor.FromArgb(0, 0, 0));
            XFont drawFont = new XFont("Univers for Schueco 330 Light", 8, XFontStyle.Regular);
            string drawString = string.Empty;

            // set the scope of the drawing
            double modelAspectRatio = pdfResult.ModelHeight / pdfResult.ModelWidth;
            if (modelAspectRatio > 0.7)
            {
                ImageWidth = ImageHeight / modelAspectRatio;
            }
            else  // height/width < 0.7, width governs
            {
                ImageHeight = modelAspectRatio * ImageWidth;
            }
            double xs = 0, ys = 0;

            //xLocation += (ImageWidth / 2 - 50);
            yLocation += ImageHeight;


            XRect rectMemberID, rectMemberID_large, rectGlassID, rectGlass, rectVent;
            foreach (MemberGeometricInfo memberGeometricInfo in pdfResult.MemberGeometricInfos)
            {
                // draw member lines
                double x0 = pdfResult.ModelOriginX;
                double y0 = pdfResult.ModelOriginY;
                double x1 = memberGeometricInfo.PointCoordinates[0];
                double y1 = memberGeometricInfo.PointCoordinates[1];
                double x2 = memberGeometricInfo.PointCoordinates[2];
                double y2 = memberGeometricInfo.PointCoordinates[3];

                double x3 = x1;
                double y3 = y1;

                double x4 = x2;
                double y4 = y2;

                // draw main structure
                if (memberGeometricInfo.MemberType == 4)  // membertype 4: major mullion, 5: transom, 6: minor mullion
                {
                    double offA = memberGeometricInfo.offsetA;
                    double offB = memberGeometricInfo.offsetB;

                    x1 -= memberGeometricInfo.width / 2;
                    x3 += memberGeometricInfo.width / 2;
                    x2 -= memberGeometricInfo.width / 2;
                    x4 += memberGeometricInfo.width / 2;

                    y1 += offA;
                    y3 += offA;
                    y2 += offB;
                    y4 += offB;
                }
                else if (memberGeometricInfo.MemberType == 5)
                {
                    x1 += memberGeometricInfo.offsetA;
                    x2 += memberGeometricInfo.offsetB;
                }
                else if (memberGeometricInfo.MemberType == 6)
                {
                    y1 += memberGeometricInfo.offsetA;
                    y2 += memberGeometricInfo.offsetB;
                }

                x1 = xLocation + (x1 - x0) / pdfResult.ModelWidth * ImageWidth;
                y1 = yLocation - (y1 - y0) / pdfResult.ModelHeight * ImageHeight;
                x2 = xLocation + (x2 - x0) / pdfResult.ModelWidth * ImageWidth;
                y2 = yLocation - (y2 - y0) / pdfResult.ModelHeight * ImageHeight;
                x3 = xLocation + (x3 - x0) / pdfResult.ModelWidth * ImageWidth;
                y3 = yLocation - (y3 - y0) / pdfResult.ModelHeight * ImageHeight;
                x4 = xLocation + (x4 - x0) / pdfResult.ModelWidth * ImageWidth;
                y4 = yLocation - (y4 - y0) / pdfResult.ModelHeight * ImageHeight;

                double w = memberGeometricInfo.width / pdfResult.ModelWidth * ImageWidth;  // scaled width of the beam

                double triangleSize = Math.Min(ImageHeight, ImageWidth) / 40;   // for anchor size

                if (memberGeometricInfo.MemberType == 4)  // major mullion
                {
                    g.DrawLine(grayPen, x1, y1, x2, y2);
                    g.DrawLine(grayPen, x2, y2, x4, y4);
                    g.DrawLine(grayPen, x1, y1, x3, y3);
                    g.DrawLine(grayPen, x3, y3, x4, y4);

                    // draw mullion ends anchors
                    // bottom
                    double eaX = (x1 + x3) / 2 + memberGeometricInfo.width / 2 / pdfResult.ModelHeight * ImageHeight;
                    double eaY = Math.Min(memberGeometricInfo.PointCoordinates[1], memberGeometricInfo.PointCoordinates[3]);
                    eaY = yLocation - (eaY - y0) / pdfResult.ModelHeight * ImageHeight;
                    g.DrawLine(grayPen, eaX, eaY, eaX + triangleSize, eaY + 0.6 * triangleSize);
                    g.DrawLine(grayPen, eaX + triangleSize, eaY + 0.6 * triangleSize, eaX + triangleSize, eaY - 0.6 * triangleSize);
                    g.DrawLine(grayPen, eaX + triangleSize, eaY - 0.6 * triangleSize, eaX, eaY);
                    g.DrawLine(grayPen, eaX + triangleSize, eaY - 0.45 * triangleSize, eaX + 1.4 * triangleSize, eaY - 0.65 * triangleSize);
                    g.DrawLine(grayPen, eaX + triangleSize, eaY - 0.15 * triangleSize, eaX + 1.4 * triangleSize, eaY - 0.35 * triangleSize);
                    g.DrawLine(grayPen, eaX + triangleSize, eaY + 0.15 * triangleSize, eaX + 1.4 * triangleSize, eaY - 0.05 * triangleSize);
                    g.DrawLine(grayPen, eaX + triangleSize, eaY + 0.45 * triangleSize, eaX + 1.4 * triangleSize, eaY + 0.25 * triangleSize);
                    // top
                    eaY = Math.Max(memberGeometricInfo.PointCoordinates[1], memberGeometricInfo.PointCoordinates[3]);
                    eaY = yLocation - (eaY - y0) / pdfResult.ModelHeight * ImageHeight;
                    g.DrawLine(grayPen, eaX, eaY, eaX + triangleSize, eaY + 0.6 * triangleSize);
                    g.DrawLine(grayPen, eaX + triangleSize, eaY + 0.6 * triangleSize, eaX + triangleSize, eaY - 0.6 * triangleSize);
                    g.DrawLine(grayPen, eaX + triangleSize, eaY - 0.6 * triangleSize, eaX, eaY);
                    g.DrawLine(grayPen, eaX + 1.4 * triangleSize, eaY - 0.6 * triangleSize, eaX + 1.4 * triangleSize, eaY + 0.6 * triangleSize);
                }
                else if (memberGeometricInfo.MemberType == 5) // transom
                {
                    //g.DrawLine(grayDashPen, x1, y1, x2, y2); // dash dot line
                    g.DrawLine(grayPen, x1, y1 - w / 2, x2, y2 - w / 2);
                    g.DrawLine(grayPen, x1, y1 + w / 2, x2, y2 + w / 2);
                }
                else if (memberGeometricInfo.MemberType == 6) // minor mullion
                {
                    //g.DrawLine(grayDashPen, x1, y1, x2, y2); // dash dot line
                    g.DrawLine(grayPen, x1 - w / 2, y1, x2 - w / 2, y2);
                    g.DrawLine(grayPen, x1 + w / 2, y1, x2 + w / 2, y2);
                }

                // draw reinforcement                
                PDFMemberResult memberResult = pdfResult.PDFMemberResults.FirstOrDefault(x => x.MemberID == memberGeometricInfo.MemberID);

                if (memberResult.isReinforced == true)
                {
                    double offsetA_temp = (memberGeometricInfo.offsetA * 0.2) / pdfResult.ModelHeight * ImageHeight;
                    double offsetB_temp = (memberGeometricInfo.offsetB * 0.2) / pdfResult.ModelHeight * ImageHeight;

                    g.DrawLine(grayPen, (x1 + x3) / 2, y1 + offsetA_temp, (x1 + x3) / 2, y2 + offsetB_temp);
                }

                // draw slab anchor     
                for (int i = 0; i < memberResult.SlabAnchorX.Count; i++)
                {
                    double y = memberResult.SlabAnchorX[i];
                    double saX = (x1 + x3) / 2 + memberGeometricInfo.width / 2 / pdfResult.ModelHeight * ImageHeight;
                    double yA = memberGeometricInfo.PointCoordinates[1];
                    double yB = memberGeometricInfo.PointCoordinates[3];
                    double saY = yA < yB ? yA + y : yA - y;
                    saY = yLocation - (saY - y0) / pdfResult.ModelHeight * ImageHeight;


                    g.DrawLine(grayPen, saX, saY, saX + triangleSize, saY + 0.6 * triangleSize);
                    g.DrawLine(grayPen, saX + triangleSize, saY + 0.6 * triangleSize, saX + triangleSize, saY - 0.6 * triangleSize);
                    g.DrawLine(grayPen, saX + triangleSize, saY - 0.6 * triangleSize, saX, saY);

                    string saType = memberResult.SlabAnchorType[i];
                    if (saType == "Sliding")
                    {
                        g.DrawLine(grayPen, saX + 1.4 * triangleSize, saY - 0.6 * triangleSize, saX + 1.4 * triangleSize, saY + 0.6 * triangleSize);
                    }
                    else if (saType == "Fixed")
                    {
                        g.DrawLine(grayPen, saX + triangleSize, saY - 0.45 * triangleSize, saX + 1.4 * triangleSize, saY - 0.65 * triangleSize);
                        g.DrawLine(grayPen, saX + triangleSize, saY - 0.15 * triangleSize, saX + 1.4 * triangleSize, saY - 0.35 * triangleSize);
                        g.DrawLine(grayPen, saX + triangleSize, saY + 0.15 * triangleSize, saX + 1.4 * triangleSize, saY - 0.05 * triangleSize);
                        g.DrawLine(grayPen, saX + triangleSize, saY + 0.45 * triangleSize, saX + 1.4 * triangleSize, saY + 0.25 * triangleSize);
                    }
                }

                // draw splice joint
                for (int i = 0; i < memberResult.SpliceJointX.Count; i++)
                {
                    double y = memberResult.SpliceJointX[i];
                    double sjX1 = (x1 + x3) / 2 - memberGeometricInfo.width / 2 / pdfResult.ModelHeight * ImageHeight;
                    double sjX2 = (x1 + x3) / 2 + memberGeometricInfo.width / 2 / pdfResult.ModelHeight * ImageHeight;
                    double yA = memberGeometricInfo.PointCoordinates[1];
                    double yB = memberGeometricInfo.PointCoordinates[3];
                    double sjY = yA < yB ? yA + y : yA - y;
                    double spliceJointSectionLenght = 20;
                    double sjY1 = sjY + spliceJointSectionLenght / 2;
                    sjY1 = yLocation - (sjY1 - y0) / pdfResult.ModelHeight * ImageHeight;
                    double sjY2 = sjY - spliceJointSectionLenght / 2;
                    sjY2 = yLocation - (sjY2 - y0) / pdfResult.ModelHeight * ImageHeight;

                    g.DrawLine(grayPen, sjX1, sjY1, sjX2, sjY1);
                    g.DrawLine(grayPen, sjX1, sjY2, sjX2, sjY2);
                }


                // draw member label
                double LabelOffset = memberGeometricInfo.width * Math.Max(ImageWidth / pdfResult.ModelWidth, ImageHeight / pdfResult.ModelHeight) + 2;
                xs = (x1 + x2) / 2 - LabelOffset - 5;
                ys = (y1 + y2) / 2 - LabelOffset;
                drawString = $"{memberGeometricInfo.MemberID}";
                rectMemberID = new XRect(xs - 2, ys - 7.5, 9, 9);
                rectMemberID_large = new XRect(xs - 2, ys - 7.5, 15, 9);

                if (memberGeometricInfo.MemberType == 4 || memberGeometricInfo.MemberType == 5)
                {
                    XFont memberFont = new XFont("Univers for Schueco 330 Light", 8, XFontStyle.Bold);
                    g.DrawString(drawString, memberFont, blackBrush, xs, ys);
                    if (memberGeometricInfo.MemberID < 10)
                    {
                        g.DrawRectangle(blackPen, rectMemberID);
                    }
                    else
                    {
                        g.DrawRectangle(blackPen, rectMemberID_large);
                    }
                }
                else
                {
                    XFont memberFont = new XFont("Univers for Schueco 330 Light", 8, XFontStyle.Regular);
                    g.DrawString(drawString, drawFont, grayBrush, xs, ys);
                    if (memberGeometricInfo.MemberID < 10)
                    {
                        g.DrawRectangle(grayPen, rectMemberID);
                    }
                    else
                    {
                        g.DrawRectangle(grayPen, rectMemberID_large);
                    }
                }
            }

            foreach (GlassGeometricInfo glassGeometricInfo in pdfResult.GlassGeometricInfos)
            {
                double x0 = pdfResult.ModelOriginX;
                double y0 = pdfResult.ModelOriginY;
                xs = glassGeometricInfo.PointCoordinates[0];
                ys = glassGeometricInfo.PointCoordinates[1];
                xs = xLocation + (xs - x0) / pdfResult.ModelWidth * ImageWidth;
                ys = yLocation - (ys - y0) / pdfResult.ModelHeight * ImageHeight;
                drawString = $"{glassGeometricInfo.GlassID}";

                if (glassGeometricInfo.GlassID < 10)
                {
                    g.DrawString(drawString, drawFont, grayBrush, xs, ys);
                }
                if (glassGeometricInfo.GlassID >= 10)
                {
                    g.DrawString(drawString, drawFont, grayBrush, xs - 1, ys);
                }

                if (glassGeometricInfo.GlassID < 10)
                {
                    rectGlassID = new XRect(xs - 2, ys - 7.5, 9, 9);
                    g.DrawArc(grayPen, rectGlassID, 0F, 360F);
                }
                if (glassGeometricInfo.GlassID >= 10)
                {
                    rectGlassID = new XRect(xs - 2.5, ys - 8.5, 12, 12);
                    g.DrawArc(grayPen, rectGlassID, 0F, 360F);
                }

                // draw glass and vent
                x0 = pdfResult.ModelOriginX;
                y0 = pdfResult.ModelOriginY;
                double x1 = glassGeometricInfo.CornerCoordinates[0];
                double y1 = glassGeometricInfo.CornerCoordinates[1];
                double x2 = glassGeometricInfo.CornerCoordinates[2];
                double y2 = glassGeometricInfo.CornerCoordinates[3];
                double dx = (x2 - x1) / pdfResult.ModelWidth * ImageWidth;
                double dy = (y2 - y1) / pdfResult.ModelHeight * ImageHeight;
                x1 = xLocation + (x1 - x0) / pdfResult.ModelWidth * ImageWidth;
                y1 = yLocation - (y1 - y0) / pdfResult.ModelHeight * ImageHeight;
                rectGlass = new XRect(x1, y1 - dy, dx, dy);
                g.DrawRectangle(grayPen, rectGlass);

                if (!(glassGeometricInfo.VentCoordinates is null))
                {
                    x1 = glassGeometricInfo.VentCoordinates[0];
                    y1 = glassGeometricInfo.VentCoordinates[1];
                    x2 = glassGeometricInfo.VentCoordinates[2];
                    y2 = glassGeometricInfo.VentCoordinates[3];
                    dx = (x2 - x1) / pdfResult.ModelWidth * ImageWidth;
                    dy = (y2 - y1) / pdfResult.ModelHeight * ImageHeight;
                    x1 = xLocation + (x1 - x0) / pdfResult.ModelWidth * ImageWidth;
                    y1 = yLocation - (y1 - y0) / pdfResult.ModelHeight * ImageHeight;
                    rectVent = new XRect(x1, y1 - dy, dx, dy);
                    g.DrawRectangle(grayPen, rectVent);
                    // draw symbol
                    XPen vSymbolPen;
                    if (glassGeometricInfo.VentOpeningDirection.Contains("Inward"))
                    {
                        vSymbolPen = grayDashPen;
                    }
                    else
                    {
                        vSymbolPen = grayPen;
                    }
                    if (glassGeometricInfo.VentOperableType.Contains("Tilt"))
                    {
                        if ((glassGeometricInfo.VentOperableType.Contains("Right") && glassGeometricInfo.VentOpeningDirection.Contains("Inward")) ||
                            (glassGeometricInfo.VentOperableType.Contains("Left") && glassGeometricInfo.VentOpeningDirection.Contains("Outward")))
                        {
                            g.DrawLine(vSymbolPen, x1, y1, x1 + dx / 2, y1 - dy);
                            g.DrawLine(vSymbolPen, x1 + dx / 2, y1 - dy, x1 + dx, y1);
                            g.DrawLine(vSymbolPen, x1, y1, x1 + dx, y1 - dy / 2);
                            g.DrawLine(vSymbolPen, x1 + dx, y1 - dy / 2, x1, y1 - dy);
                        }
                        else
                        {
                            g.DrawLine(vSymbolPen, x1, y1, x1 + dx / 2, y1 - dy);
                            g.DrawLine(vSymbolPen, x1 + dx / 2, y1 - dy, x1 + dx, y1);
                            g.DrawLine(vSymbolPen, x1, y1 - dy / 2, x1 + dx, y1);
                            g.DrawLine(vSymbolPen, x1, y1 - dy / 2, x1 + dx, y1 - dy);
                        }

                    }
                    else if (glassGeometricInfo.VentOperableType.Contains("Side"))
                    {
                        if ((glassGeometricInfo.VentOperableType.Contains("Right") && glassGeometricInfo.VentOpeningDirection.Contains("Inward")) ||
                            (glassGeometricInfo.VentOperableType.Contains("Left") && glassGeometricInfo.VentOpeningDirection.Contains("Outward")))
                        {
                            g.DrawLine(vSymbolPen, x1, y1, x1 + dx, y1 - dy / 2);
                            g.DrawLine(vSymbolPen, x1 + dx, y1 - dy / 2, x1, y1 - dy);
                        }
                        else
                        {
                            g.DrawLine(vSymbolPen, x1, y1 - dy / 2, x1 + dx, y1);
                            g.DrawLine(vSymbolPen, x1, y1 - dy / 2, x1 + dx, y1 - dy);
                        }
                    }
                    else if (glassGeometricInfo.VentOperableType.Contains("Bottom"))
                    {
                        g.DrawLine(vSymbolPen, x1, y1, x1 + dx / 2, y1 - dy);
                        g.DrawLine(vSymbolPen, x1 + dx / 2, y1 - dy, x1 + dx, y1);
                    }
                    else if (glassGeometricInfo.VentOperableType.Contains("Parallel"))
                    {
                        g.DrawLine(vSymbolPen, x1, y1, x1 + dx / 2, y1 - dy / 4);
                        g.DrawLine(vSymbolPen, x1 + dx / 2, y1 - dy / 4, x1 + dx, y1);
                        g.DrawLine(vSymbolPen, x1 + dx / 2, y1 - dy / 4, x1 + dx / 2, y1 - dy * 3 / 4);
                        g.DrawLine(vSymbolPen, x1, y1 - dy, x1 + dx / 2, y1 - dy * 3 / 4);
                        g.DrawLine(vSymbolPen, x1 + dx / 2, y1 - dy * 3 / 4, x1 + dx, y1 - dy);
                    }
                }
            }

            // get dimension info
            double[] xdimensions = pdfResult.MemberGeometricInfos.Select(item => item.PointCoordinates[0]).Distinct().ToArray();
            List<double> yds1 = pdfResult.MemberGeometricInfos.Select(item => item.PointCoordinates[1]).ToList();
            List<double> yds3 = pdfResult.MemberGeometricInfos.Select(item => item.PointCoordinates[3]).ToList();
            yds1.AddRange(yds3);
            double[] ydimensions = yds1.Distinct().ToArray();
            Array.Sort(xdimensions);
            Array.Sort(ydimensions);
            double modelLength = xdimensions.Last() - xdimensions.First();
            double modelHeight = ydimensions.Last() - ydimensions.First();
            xdimensions = xdimensions.Select(item => (item - xdimensions.First()) / (xdimensions.Last() - xdimensions.First())).ToArray();
            ydimensions = ydimensions.Select(item => (item - ydimensions.First()) / (ydimensions.Last() - ydimensions.First())).ToArray();

            // get ydimension_SlabAnchor
            List<double> ydsTemp = pdfResult.PDFMemberResults[0].SlabAnchorX.Select(x => x).ToList();
            for (int i = 1; i < pdfResult.PDFMemberResults.Count; i++)
            {
                ydsTemp.AddRange(pdfResult.PDFMemberResults[i].SlabAnchorX);
            }
            double[] ydimensions_SlabAnchor = ydsTemp.Distinct().ToArray();
            Array.Sort(ydimensions_SlabAnchor);
            ydimensions_SlabAnchor = ydimensions_SlabAnchor.Select(item => (item - ydimensions_SlabAnchor.First()) / (ydimensions_SlabAnchor.Last() - ydimensions_SlabAnchor.First())).ToArray();

            // get ydimension_SpliceJoint
            ydsTemp = pdfResult.PDFMemberResults[0].SpliceJointX.Select(x => x).ToList();
            for (int i = 1; i < pdfResult.PDFMemberResults.Count; i++)
            {
                ydsTemp.AddRange(pdfResult.PDFMemberResults[i].SpliceJointX);
            }
            ydsTemp.Add(0);
            ydsTemp.Add(modelHeight);
            double[] ydimensions_SpliceJoint = ydsTemp.Distinct().ToArray();
            Array.Sort(ydimensions_SpliceJoint);

            ydimensions_SpliceJoint = ydimensions_SpliceJoint.Select(item => (item - ydimensions_SpliceJoint.First()) / (ydimensions_SpliceJoint.Last() - ydimensions_SpliceJoint.First())).ToArray();

            // draw legend
            xs = xLocation + 5;
            ys = yLocation + 30 + 5 * xdimensions.Count();
            drawString = $"n   Glass ID";
            if (Thread.CurrentThread.CurrentCulture.Name == "de-DE")
            {
                drawString = $"n   Glas-Position";
            }
            else if (Thread.CurrentThread.CurrentCulture.Name == "fr-FR")
            {
                drawString = $"n   Position du verre";
            }

            g.DrawString(drawString, drawFont, grayBrush, xs, ys);
            rectGlassID = new XRect(xs - 2, ys - 7.5, 9, 9);
            g.DrawArc(grayPen, rectGlassID, 0, 360);

            ys = ys + 15;
            drawString = $"n   Structural Member ID";
            if (Thread.CurrentThread.CurrentCulture.Name == "de-DE")
            {
                drawString = $"n   Statik-Position";
            }
            else if (Thread.CurrentThread.CurrentCulture.Name == "fr-FR")
            {
                drawString = $"n   ID des profilés";
            }

            drawFont = new XFont("Univers for Schueco 330 Light", 8, XFontStyle.Bold);
            g.DrawString(drawString, drawFont, blackBrush, xs, ys);
            rectMemberID = new XRect(xs - 2, ys - 7.5, 9, 9);
            g.DrawRectangle(blackPen, rectMemberID);

            g.Save();
            g.Dispose();

            // draw dimensions
            double xdimLocation = xLocation;
            double ydimLocation = yLocation + 20;
            for (int i = 1; i < xdimensions.Count(); i++)
            {
                PdfDrawDimension(myTemplate.Pages[pageNumber], xdimensions[i - 1], xdimensions[i], 0, (xdimensions[i] - xdimensions[i - 1]) * modelLength, xdimLocation, ydimLocation, ImageWidth);
                //PdfDrawDimension(PdfPage myPage, double xStart, double xEnd, int nc, double value, double xLocation, double yLocation, double Length, bool isVertical = false)
            }
            PdfDrawDimension(myTemplate.Pages[pageNumber], 0, xdimensions[xdimensions.Count() - 1], 1, xdimensions[xdimensions.Count() - 1] * modelLength, xdimLocation, ydimLocation, ImageWidth);
            double dimOffset = ImageWidth < 200 ? 5 : 10;
            xdimLocation = xLocation + ImageWidth + dimOffset;
            ydimLocation = yLocation;
            for (int i = 1; i < ydimensions.Count(); i++)
            {
                PdfDrawDimension(myTemplate.Pages[pageNumber], ydimensions[i - 1], ydimensions[i], 0, (ydimensions[i] - ydimensions[i - 1]) * modelHeight, xdimLocation, ydimLocation, ImageHeight, true);
            }
            // draw slab anchor dimensions
            ydimensions = ydimensions_SlabAnchor;
            for (int i = 1; i < ydimensions.Count(); i++)
            {
                PdfDrawDimension(myTemplate.Pages[pageNumber], ydimensions[i - 1], ydimensions[i], 1, (ydimensions[i] - ydimensions[i - 1]) * modelHeight, xdimLocation, ydimLocation, ImageHeight, true);
            }
            // draw splice joint dimensions
            ydimensions = ydimensions_SpliceJoint;
            for (int i = 1; i < ydimensions.Count(); i++)
            {
                PdfDrawDimension(myTemplate.Pages[pageNumber], ydimensions[i - 1], ydimensions[i], 2, (ydimensions[i] - ydimensions[i - 1]) * modelHeight, xdimLocation, ydimLocation, ImageHeight, true);
            }
            PdfDrawDimension(myTemplate.Pages[pageNumber], 0, ydimensions[ydimensions.Count() - 1], 3, ydimensions[ydimensions.Count() - 1] * modelHeight, xdimLocation, ydimLocation, ImageHeight, true);
        }


        public void PdfDrawBeam(PDFMemberResult pdfMemberResult, PdfDocument myTemplate, int chartNo, double y)
        {
            XGraphics g = XGraphics.FromPdfPage(myTemplate.Pages[0]);

            XPen grayPen = new XPen(XColor.FromArgb(38, 38, 38), .25);
            XPen whitePen = new XPen(XColor.FromArgb(255, 255, 255), .10);
            XPen blackPen = new XPen(XColor.FromArgb(0, 0, 0), .10);

            XBrush blueBrush = new XSolidBrush(XColor.FromArgb(0, 162, 209));
            XBrush blackBrush = new XSolidBrush(XColor.FromArgb(0, 0, 0));
            XBrush grayBrush = new XSolidBrush(XColor.FromArgb(102, 102, 102));
            XBrush whiteBrush = new XSolidBrush(XColor.FromArgb(255, 255, 255));

            double x, width, height;

            // Draw top beam
            // Draw top beam
            x = 100;

            width = 400;
            height = 10;
            if (chartNo == 5)
            {
                height = 5;
                g.DrawRectangle(grayPen, grayBrush, x, y, width, height);
            }
            else
            {
                g.DrawRectangle(grayPen, blueBrush, x, y, width, height);
            }

            // draw reinforcement
            if (pdfMemberResult.isReinforced == true)
            {
                g.DrawRectangle(blackPen, blackBrush, x, y + 4, width, height - 8);
            }

            // draw splice joint
            List<string> spliceJointType = pdfMemberResult.SpliceJointType;
            List<double> spliceJointX = pdfMemberResult.SpliceJointX;
            string drawString = "";

            XFont drawFont = new XFont("Univers for Schueco 330 Light", 6, XFontStyle.Regular);


            int rowcount = spliceJointX.Count;
            for (int i = 0; i < rowcount; i++)
            {
                string SJT = spliceJointType[i];
                double SJX = spliceJointX[i] / pdfMemberResult.Length * 400;

                double xLocation = x + SJX - 5;

                XRect Splice = new XRect(xLocation, y + 0.5, 10 - 1, 10 - 1);
                if (SJT == "?")  // between rigid and hinged
                {
                    g.DrawPie(blackPen, blackBrush, Splice, -90, 180);
                    g.DrawPie(whitePen, whiteBrush, Splice, -90, -180);
                    g.DrawArc(blackPen, Splice, 0, 360);
                }

                if (SJT == "Hinged")
                {
                    g.DrawPie(whitePen, whiteBrush, Splice, 0, 360);
                    g.DrawArc(blackPen, Splice, 0, 360);
                    drawString = "Splice joint: Hinge";
                    //textSize = g.MeasureString(drawString, drawFont);
                    g.DrawString(drawString, drawFont, whiteBrush, xLocation + 10, y + 7);
                }
                if (SJT == "Rigid")
                {
                    g.DrawPie(blackPen, blackBrush, Splice, 0, 360);
                    g.DrawArc(blackPen, Splice, 0, 360);
                    drawString = "Splice joint: Rigid";
                    //textSize = g.MeasureString(drawString, drawFont);
                    g.DrawString(drawString, drawFont, whiteBrush, xLocation + 10, y + 7);
                }

            }

            g.Save();
            g.Dispose();
        }

        public void PdfDrawPinSupport(PDFMemberResult pdfMemberResult, PdfDocument myTemplate, int chartNo, double yref)
        {
            XGraphics g = XGraphics.FromPdfPage(myTemplate.Pages[0]);

            XPen grayPen = new XPen(XColor.FromArgb(38, 38, 38), .25);
            XPen grayPen2 = new XPen(XColor.FromArgb(102, 102, 102), .25);
            XPen blackPen = new XPen(XColor.FromArgb(0, 0, 0), .10);
            XBrush blueBrush = new XSolidBrush(XColor.FromArgb(0, 162, 209));
            XBrush grayBrush = new XSolidBrush(XColor.FromArgb(102, 102, 102));
            string anchorType = "Fixed";


            // draw support
            double x = 120;
            double y = 0;
            switch (chartNo)
            {
                case 1:
                    y = yref + 10;
                    break;
                case 2:
                    y = yref + 10;
                    break;
                case 3:
                    y = yref + 10;
                    break;
                case 4:
                    y = yref + 10;
                    break;
                case 5:
                    y = yref + 5;
                    break;
            }

            // draw end support
            x = 100;
            if (pdfMemberResult.MemberType == 4) //mullion
            {
                for (int k = 0; k < pdfMemberResult.SlabAnchorX.Count; k++)
                {
                    if (Math.Abs(pdfMemberResult.SlabAnchorX[k] - 0) < 0.0001)
                    {
                        anchorType = pdfMemberResult.SlabAnchorType[k];
                        continue;
                    }
                }
            }
            else
            {
                anchorType = "Fixed";
            }

            DrawSupport(g, x, y, anchorType);

            x = 500;
            if (pdfMemberResult.MemberType == 4) //mullion
            {
                for (int k = 0; k < pdfMemberResult.SlabAnchorX.Count; k++)
                {
                    if (Math.Abs(pdfMemberResult.SlabAnchorX[k] - pdfMemberResult.Length) < 0.0001)
                    {
                        anchorType = pdfMemberResult.SlabAnchorType[k];
                        continue;
                    }
                }
            }
            else
            {
                anchorType = "Fixed";
            }

            DrawSupport(g, x, y, anchorType);

            // draw the support in the span (concrete slab support)
            List<LoadData> loadCaseData = pdfMemberResult.WindLoadDataList;
            int rowcount = loadCaseData.Count;
            for (int i = 0; i < rowcount; i++)
            {
                LoadData LD = loadCaseData[i];

                if (LD.isConcentratedLoad == true && LD.isBoundaryReaction == true)  // plot the reaction boundary force
                {
                    if (LD.x1 == 0 || LD.x1 == pdfMemberResult.Length)
                    {
                        continue;
                    }

                    double loc = LD.x1;

                    int tempNum = 0;
                    for (int k = 0; k < pdfMemberResult.SlabAnchorX.Count; k++)
                    {
                        if (pdfMemberResult.SlabAnchorX[k] == loc)
                        {
                            tempNum = k;
                            break;
                        }
                    }
                    anchorType = pdfMemberResult.SlabAnchorType[tempNum];

                    x = 100 + LD.x1 / pdfMemberResult.Length * 400;
                    DrawSupport(g, x, y, anchorType);

                }
            }

            g.Save();
            g.Dispose();
        }

        public void DrawSupport(XGraphics g, double x, double y, string anchorType)
        {
            XPen grayPen = new XPen(XColor.FromArgb(38, 38, 38), .25);
            XPen grayPen2 = new XPen(XColor.FromArgb(102, 102, 102), .25);
            XPen blackPen = new XPen(XColor.FromArgb(0, 0, 0), .10);
            XBrush blueBrush = new XSolidBrush(XColor.FromArgb(0, 162, 209));
            XBrush grayBrush = new XSolidBrush(XColor.FromArgb(102, 102, 102));


            if (anchorType == "Fixed")
            {
                XPoint point1 = new XPoint(x, y);
                XPoint point2 = new XPoint(x - 5, y + 10);
                XPoint point3 = new XPoint(x + 5, y + 10);
                XPoint[] curvePointsCen = { point1, point2, point3 };

                g.DrawPolygon(grayPen, grayBrush, curvePointsCen, XFillMode.Alternate);
                g.DrawRectangle(grayPen, blueBrush, x - 7, y + 10, 14, 2);
            }
            if (anchorType == "Sliding")
            {

                XRect Splice = new XRect(x - 5, y, 10, 10);

                g.DrawPie(grayPen2, grayBrush, Splice, 0, 360);
                g.DrawArc(grayPen, Splice, 0, 360);

                g.DrawRectangle(grayPen, blueBrush, x - 7, y + 10, 14, 2);
            }
        }


        public void PdfDrawLoadCaseImageTitle(PdfDocument myTemplate, bool isTransom, int chartNo, double yref)
        {
            double x = 220, y = 0;
            switch (chartNo)
            {
                case 1:
                    x = 220;
                    y = yref - 50;
                    break;
                case 2:
                    x = 220;
                    y = yref - 50;
                    break;
                case 3:
                    x = 160;
                    y = yref - 30;
                    break;
                case 4:
                    x = 255;
                    y = yref;
                    break;
                case 5:
                    x = 260;
                    y = yref - 30;
                    break;
            }

            XGraphics g = XGraphics.FromPdfPage(myTemplate.Pages[0]);
            XPen pen = new XPen(XColor.FromArgb(38, 38, 38), .25);
            XBrush whitebrush = new XSolidBrush(XColor.FromArgb(255, 255, 255));
            XBrush graybrush = new XSolidBrush(XColor.FromArgb(102, 102, 102));
            //XBrush brush = chartNo == 4 ? graybrush : whitebrush;
            XBrush brush = graybrush;

            XFont drawFont = new XFont("Univers for Schueco 330 Light", 10, XFontStyle.Bold);
            string drawString = "";
            if (chartNo == 1)
            {
                drawString = isTransom ? "Wind Load on Transom Top Side" : "Wind Load on Mullion Left Side";
                if (Thread.CurrentThread.CurrentCulture.Name == "de-DE")
                {
                    drawString = isTransom ? "Last auf der Oberseite des Riegels" : "Last auf der linken Seite des Pfostens";
                }
                else if (Thread.CurrentThread.CurrentCulture.Name == "fr-FR")
                {
                    drawString = isTransom ? "Charge de vent sur le dessus du tableau arrière" : "Charge de vent sur la trame gauche du poteau";
                }
            }
            else if (chartNo == 2)
            {
                drawString = isTransom ? "Wind Load on Transom Bottom Side" : "Wind Load on Mullion Right Side";
                if (Thread.CurrentThread.CurrentCulture.Name == "de-DE")
                {
                    drawString = isTransom ? "Last auf der Unterseite des Riegels" : "Last auf der rechten Seite des Pfostens";
                }
                else if (Thread.CurrentThread.CurrentCulture.Name == "fr-FR")
                {
                    drawString = isTransom ? "Charge de vent sur le côté inférieur du tableau arrière" : "Charge de vent sur la trame droite du poteau";
                }
            }
            else if (chartNo == 3)
            {
                drawString = "Horizontal Live Load (most unfavorable combination for ULS)";
                if (Thread.CurrentThread.CurrentCulture.Name == "de-DE")
                {
                    drawString = "Horizontale Nutzlast (maßgebende Lastfallkombination für GZT)";
                }
                else if (Thread.CurrentThread.CurrentCulture.Name == "fr-FR")
                {
                    drawString = "Hauteur d'application de la charge d'exploitation (most unfavorable combination)";
                }
            }
            else if (chartNo == 4)
            {
                drawString = "Reaction Force (horizontal)";
                if (Thread.CurrentThread.CurrentCulture.Name == "de-DE")
                {
                    drawString = "Auflagerkräfte (horizontale)";
                }
                else if (Thread.CurrentThread.CurrentCulture.Name == "fr-FR")
                {
                    drawString = "Réactions aux appuis (horizontale)";
                }
            }
            else if (chartNo == 5)
            {
                drawString = "Dead Load (vertical)";
                if (Thread.CurrentThread.CurrentCulture.Name == "de-DE")
                {
                    drawString = "Eigengewicht (vertikal)";
                }
                else if (Thread.CurrentThread.CurrentCulture.Name == "fr-FR")
                {
                    drawString = "Poids propre (vertical)";
                }
            }

            XSize size = g.MeasureString(drawString, drawFont);

            g.DrawString(drawString, drawFont, brush, x, y);

            if (chartNo == 5)
            {
                string drawNoteString = "*Self-weight is not shown here";
                if (Thread.CurrentThread.CurrentCulture.Name == "de-DE")
                {
                    drawNoteString = "*Self-weight is not shown here";
                }
                else if (Thread.CurrentThread.CurrentCulture.Name == "fr-FR")
                {
                    drawNoteString = "*Self-weight is not shown here";
                }
                XFont drawNoteFont = new XFont("Univers for Schueco 330 Light", 8, XFontStyle.Bold);
                g.DrawString(drawNoteString, drawNoteFont, brush, 365, y);

            }


            g.Save();
            g.Dispose();
        }

        public void PdfDrawLoads(PDFMemberResult pdfMemberResult, PdfDocument myTemplate, int chartNo, double y5)
        {
            double len = pdfMemberResult.Length;

            List<LoadData> loadCaseData;

            //int rowcount; // 16 numbers of loaddata
            double[] dimData = new double[200];

            if (chartNo == 1)
            {
                loadCaseData = pdfMemberResult.WindLoadDataList;
                List<LoadData> loadCaseDataLoadside1 = loadCaseData.Where(x => x.LoadSide == 1).OrderBy(x => x.x1).ToList();

                PdfDrawLoadDimension(pdfMemberResult, myTemplate, loadCaseDataLoadside1, chartNo, y5);
            }
            else if (chartNo == 2)
            {
                loadCaseData = pdfMemberResult.WindLoadDataList;
                List<LoadData> loadCaseDataLoadside2 = loadCaseData.Where(x => x.LoadSide == 2).OrderBy(x => x.x1).ToList();

                PdfDrawLoadDimension(pdfMemberResult, myTemplate, loadCaseDataLoadside2, chartNo, y5);
            }
            else if (chartNo == 3)// merge concentrated load for horizontal live load, since it has only one chart, left and right load may overlap each other
            {
                loadCaseData = pdfMemberResult.HLLDataList;
                MergeLoadCaseData(loadCaseData);
                PdfDrawLoadDimension(pdfMemberResult, myTemplate, loadCaseData, chartNo, y5);
            }
            else if (chartNo == 5) // merge concentrated load for vertical load, since it has only one chart
            {
                loadCaseData = pdfMemberResult.WeightDataList.Where(x => x.isBoundaryReaction == false).ToList();
                MergeLoadCaseData(loadCaseData);
                PdfDrawLoadDimension(pdfMemberResult, myTemplate, loadCaseData, chartNo, y5);
            }
        }

        internal void MergeLoadCaseData(List<LoadData> loadCaseData)  // merge concentrated loads at same location
        {
            int rowcount = loadCaseData.Count;
            for (int i = 0; i < rowcount - 1; i++)
            {
                rowcount = loadCaseData.Count;
                for (int j = i + 1; j < rowcount; j++)
                {
                    if (Math.Abs(loadCaseData[i].x1 - loadCaseData[j].x1) < 0.00001 &&
                        Math.Abs(loadCaseData[i].x2 - loadCaseData[j].x2) < 0.00001)
                    {
                        loadCaseData[i].q1 = loadCaseData[i].q1 + loadCaseData[j].q1;
                        loadCaseData[i].q2 = loadCaseData[i].q2 + loadCaseData[j].q2;
                        loadCaseData[j].q1 = 0;
                        loadCaseData[j].q2 = 0;
                        break;
                    }
                }
            }
        }

        public void PdfDrawLoadDimension(PDFMemberResult pdfMemberResult, PdfDocument myTemplate, List<LoadData> loadCaseData, int chartNo, double yref)
        {
            double xStart, xEnd, valueStart, valueEnd, valueMax, valueMax_con;
            bool isHLiveLoad;

            int dimCounter = 0;
            double[] dimData = new double[200];

            double slope_cur = 0, slope_next = -999999;
            bool sameSlope = false;

            double len = pdfMemberResult.Length;

            /////////////////////////////////////////
            /// Get the maximum value for distributed & concentrated load
            /////////////////////////////////////////
            int rowcount = loadCaseData.Count;
            // distributed load
            valueMax = 0;
            for (int i = 0; i < rowcount; i++)
            {
                LoadData LD = loadCaseData[i];
                if (LD.isConcentratedLoad == false)
                {
                    if (Math.Abs(LD.q1) > valueMax)
                    {
                        valueMax = Math.Abs(LD.q1);
                    }

                    if (Math.Abs(LD.q2) > valueMax)
                    {
                        valueMax = Math.Abs(LD.q2);
                    }
                }
            }
            // allow some room for text
            valueMax = valueMax * 1.25;

            // concentrated load
            valueMax_con = 0;
            for (int i = 0; i < rowcount; i++)
            {
                LoadData LD = loadCaseData[i];
                if (LD.isConcentratedLoad == true)
                {
                    if (Math.Abs(LD.q1) > valueMax_con)
                    {
                        valueMax_con = Math.Abs(LD.q1);
                    }

                    if (Math.Abs(LD.q2) > valueMax_con)
                    {
                        valueMax_con = Math.Abs(LD.q2);
                    }
                }
            }
            // allow some room for text
            valueMax_con = valueMax_con * 1.25;

            /////////////////////////////////////////
            /// Plot the load
            ///////////////////////////////////////// 
            /// distributed load
            List<LoadData> tempDistData = loadCaseData.Where(x => x.isConcentratedLoad == false).OrderBy(x => x.x1).ToList();
            int rowCountDist = tempDistData.Count;
            for (int i = 0; i < rowCountDist; i++)
            {
                LoadData LD = tempDistData[i];

                xStart = LD.x1 / len;
                xEnd = LD.x2 / len;
                valueStart = LD.q1;
                valueEnd = LD.q2;

                slope_cur = (valueEnd - valueStart) / (xEnd - xStart);

                if (i < rowCountDist - 1)
                {
                    slope_next = (tempDistData[i + 1].q2 - tempDistData[i + 1].q1) / (tempDistData[i + 1].x2 / len - tempDistData[i + 1].x1 / len);
                }
                else if (i == rowCountDist - 1)
                {
                    slope_next = -999999;
                }

                sameSlope = Math.Abs(slope_cur - slope_next) < 0.00001;

                if (slope_cur == 0 && slope_next == 0)
                {
                    if (LD.q1 != tempDistData[i + 1].q1)
                    {
                        sameSlope = false;      // step function, with different q1 value, we need to plot the value
                    }
                }


                isHLiveLoad = chartNo == 3;

                if (Math.Abs(valueStart) > 0.0001 || Math.Abs(valueEnd) > 0.0001)
                {
                    PdfDrawForce(myTemplate, xStart, xEnd, valueStart, valueEnd, valueMax, chartNo, yref, sameSlope, isHLiveLoad);
                    dimData[dimCounter] = xStart;
                    dimData[dimCounter + 1] = xEnd;
                    dimCounter = dimCounter + 2;
                }
            }
            /// concentrated load
            List<LoadData> tempConData = loadCaseData.Where(x => x.isConcentratedLoad == true).OrderBy(x => x.x1).ToList();
            int rowCountCon = tempConData.Count;
            for (int i = 0; i < rowCountCon; i++)
            {
                LoadData LD = tempConData[i];
                bool isConcentrated = true;

                if (LD.isBoundaryReaction == true) { continue; }

                xStart = LD.x1 / len;
                xEnd = LD.x2 / len;
                valueStart = LD.q1;
                valueEnd = LD.q2;

                isHLiveLoad = chartNo == 3;

                if (Math.Abs(valueStart) > 0.0001 || Math.Abs(valueEnd) > 0.0001)
                {
                    PdfDrawForce(myTemplate, xStart, xEnd, valueStart, valueEnd, valueMax_con, chartNo, yref, sameSlope, isHLiveLoad, isConcentrated);
                    dimData[dimCounter] = xStart;
                    dimData[dimCounter + 1] = xEnd;
                    dimCounter = dimCounter + 2;
                }
            }

            // add one entery to dimData for the overall dimension and start
            dimData[dimCounter] = 0;
            dimData[dimCounter + 1] = 1;

            // sort the dimension lines
            dimData = dimData.OrderBy(i => i).ToArray();
            double xLocation = 100;
            double yLocation = yref + 40;
            double Length = 400;
            int dimNumber = 0;
            // dimension for loads
            for (int i = 1; i < dimData.Length; i++)
            {
                if (dimData[i] != 0 && dimData[i] != dimData[i - 1] && (dimData[i] - dimData[i - 1]) != 1)
                {
                    PdfDrawDimension(myTemplate.Pages[0], dimData[i - 1], dimData[i], 0, (dimData[i] - dimData[i - 1]) * len, xLocation, yLocation, Length);
                    dimNumber = dimNumber + 1;
                }
            }
            // dimension for whole beam
            if (dimNumber == 0)
            {
                PdfDrawDimension(myTemplate.Pages[0], 0, dimData[dimData.Length - 1], 0, dimData[dimData.Length - 1] * len, xLocation, yLocation, Length);
            }
            else if (chartNo == 1 || chartNo == 2)
            {
                PdfDrawDimension(myTemplate.Pages[0], 0, dimData[dimData.Length - 1], 1, dimData[dimData.Length - 1] * len, xLocation, yLocation, Length);
            }
        }

        public void PdfDrawDimensionSupport(PDFMemberResult pdfMemberResult, PdfDocument myTemplate, bool existHorizontalLiveLoad)
        {
            List<int> SJXs = pdfMemberResult.SpliceJointX.Select(x => Convert.ToInt32(x)).ToList();
            List<int> SAXs = pdfMemberResult.SlabAnchorX.Select(x => Convert.ToInt32(x)).ToList();
            List<int> featureXs = SJXs.Union(SAXs).ToList();
            featureXs.Sort();

            int count = featureXs.Count;

            int xLocation = 100, yLocation = existHorizontalLiveLoad ? 715 : 640;
            int Length = 400;
            double len = pdfMemberResult.Length;
            int x1, x2;

            for (int i = 0; i < count - 1; i++)  // remove first and the last, because they are beam's two ends
            {

                x1 = featureXs[i];
                x2 = featureXs[i + 1]; // next LD

                PdfDrawDimension(myTemplate.Pages[0], x1 / len, x2 / len, 0,
                    x2 - x1, xLocation, yLocation, Length);
            }
        }

        public void PdfDrawField(PDFMemberResult pdfMemberResult, PdfPage myPage, bool existHorizontalLiveLoad)
        {
            XGraphics g = XGraphics.FromPdfPage(myPage);
            List<LoadData> loadCaseData = pdfMemberResult.WindLoadDataList;
            List<LoadData> boundaryData = loadCaseData.Where(x => x.isBoundaryReaction == true).OrderBy(x => x.x1).ToList();
            int rowCountBou = boundaryData.Count;

            XPen bluePen = new XPen(XColor.FromArgb(0, 162, 209), 0.25);
            XBrush blueBrush = new XSolidBrush(XColor.FromArgb(0, 162, 209));
            XFont drawFont = new XFont("Univers for Schueco 330 Light", 6, XFontStyle.Regular);

            String drawString;
            XSize textSize;
            double y = existHorizontalLiveLoad ? 685 : 615; ;

            double xtart = 100;
            double len = pdfMemberResult.Length;
            double beamLength = 400;

            for (int i = 0; i < rowCountBou - 1; i++)
            {
                double left = xtart + boundaryData[i].x1 / len * beamLength;
                double right = xtart + boundaryData[i + 1].x1 / len * beamLength;

                if (Thread.CurrentThread.CurrentCulture.Name == "de-DE")
                {
                    drawString = "Feld " + (i + 1).ToString();
                }
                else if (Thread.CurrentThread.CurrentCulture.Name == "fr-FR")
                {
                    drawString = "Travée " + (i + 1).ToString();
                }
                else
                {
                    drawString = "Field " + (i + 1).ToString();
                }

                textSize = g.MeasureString(drawString, drawFont);
                g.DrawString(drawString, drawFont, blueBrush, (left + right) / 2 - textSize.Width, y);
            }

            g.Save();
            g.Dispose();

        }


        public void PdfDrawReactions(PDFMemberResult pdfMemberResult, PdfPage myPage, int chartNo, double yref)
        {

            double[] reactionForce = pdfMemberResult.ReactionForces;

            // draw reactions
            double xLocation, xtart = 100;
            double yLocation = yref + 22.5;
            if (chartNo == 5) yLocation = yref + 17.5;
            double beamLength = 400;
            double x, y;

            if (chartNo == 4)
            {
                // Wind reaction force
                List<LoadData> loadCaseData = pdfMemberResult.WindLoadDataList;
                int rowcount = loadCaseData.Count;
                for (int i = 0; i < rowcount; i++)
                {
                    LoadData LD = loadCaseData[i];

                    if (LD.isConcentratedLoad == true && LD.isBoundaryReaction == true)  // plot the reaction boundary force
                    {
                        //if (LD.x1 == 0 || LD.x1 == pdfMemberResult.Length) { continue; }

                        // draw the boundary load 
                        xLocation = xtart + LD.x1 / pdfMemberResult.Length * beamLength;
                        x = xLocation;
                        y = yLocation;

                        if (LD.x1 == pdfMemberResult.Length)
                        {
                            PdfDrawVerticalReaction(myPage, x, y, LD.q1, 0, "right");
                        }
                        else
                        {
                            PdfDrawVerticalReaction(myPage, x, y, LD.q1, 0, "left");
                        }

                    }
                }

                // HLL reaction force
                loadCaseData = pdfMemberResult.HLLDataList;
                rowcount = loadCaseData.Count;
                for (int i = 0; i < rowcount; i++)
                {
                    LoadData LD = loadCaseData[i];

                    if (LD.isConcentratedLoad == true && LD.isBoundaryReaction == true)  // plot the reaction boundary force
                    {
                        //if (LD.x1 == 0 || LD.x1 == pdfMemberResult.Length) { continue; }
                        if (Math.Abs(LD.q1) < 1e-5 && Math.Abs(LD.q2) < 1e-5) continue;

                        // draw the boundary load 
                        xLocation = xtart + LD.x1 / pdfMemberResult.Length * beamLength;
                        x = xLocation;
                        y = yLocation;

                        if (LD.x1 == pdfMemberResult.Length)
                        {
                            PdfDrawVerticalReaction(myPage, x, y, 0, LD.q1, "right");
                        }
                        else
                        {
                            PdfDrawVerticalReaction(myPage, x, y, 0, LD.q1, "left");
                        }
                    }
                }

                // Weight reaction force
                if (pdfMemberResult.MemberType == 4)
                {
                    loadCaseData = pdfMemberResult.WeightDataList;
                    rowcount = loadCaseData.Count;
                    for (int i = 0; i < rowcount; i++)
                    {
                        LoadData LD = loadCaseData[i];

                        if (LD.isConcentratedLoad == true && LD.isBoundaryReaction == true)  // plot the reaction boundary force
                        {
                            if (LD.x1 != 0 && LD.x1 != pdfMemberResult.Length) continue;
                            if (Math.Abs(LD.q1) < 1e-3 && Math.Abs(LD.q2) < 1e-3) continue;

                            // draw the boundary load 
                            xLocation = xtart + LD.x1 / pdfMemberResult.Length * beamLength;
                            x = xLocation;
                            y = yLocation - 12.5;

                            if (LD.x1 == pdfMemberResult.Length)
                            {
                                PdfDrawHorizontalReaction(myPage, x, y, LD.q1, "right");
                            }
                            else
                            {
                                PdfDrawHorizontalReaction(myPage, x, y, LD.q1, "left");
                            }
                        }
                    }
                }
            }
            else if (chartNo == 5)
            {
                // Weight reaction force
                List<LoadData> loadCaseData = pdfMemberResult.WeightDataList;
                int rowcount = loadCaseData.Count;
                for (int i = 0; i < rowcount; i++)
                {
                    LoadData LD = loadCaseData[i];

                    if (LD.isConcentratedLoad == true && LD.isBoundaryReaction == true)  // plot the reaction boundary force
                    {
                        //if (LD.x1 == 0 || LD.x1 == pdfMemberResult.Length) { continue; }

                        // draw the boundary load 
                        xLocation = xtart + LD.x1 / pdfMemberResult.Length * beamLength;
                        x = xLocation;
                        y = yLocation;

                        if (LD.x1 == pdfMemberResult.Length)
                        {
                            PdfDrawVerticalReaction(myPage, x, y, 0, 0, "right", -LD.q1);
                        }
                        else
                        {
                            PdfDrawVerticalReaction(myPage, x, y, 0, 0, "left", -LD.q1);
                        }

                    }
                }
            }
        }

        public void PdfDrawVerticalReaction(PdfPage myPage, double x, double y, double wlRF, double hllRF, string side, double dlRF = 0)
        {
            XGraphics g = XGraphics.FromPdfPage(myPage);
            XSize textSize;

            XFont drawFont = new XFont("Univers for Schueco 330 Light", 6, XFontStyle.Regular);
            XPen grayPen = new XPen(XColor.FromArgb(38, 38, 38), .25);
            XPen grayPenThick = new XPen(XColor.FromArgb(102, 102, 102), 1.5);
            XBrush blueBrush = new XSolidBrush(XColor.FromArgb(0, 162, 209));
            XBrush grayBrush = new XSolidBrush(XColor.FromArgb(102, 102, 102));

            XPoint point1 = new XPoint(x, y);
            XPoint point2 = new XPoint(x - 1.75, y + 5);
            XPoint point3 = new XPoint(x - 0.75, y + 5);
            XPoint point4 = new XPoint(x - 0.75, y + 15);
            XPoint point5 = new XPoint(x + 0.75, y + 15);
            XPoint point6 = new XPoint(x + 0.75, y + 5);
            XPoint point7 = new XPoint(x + 1.75, y + 5);

            XPoint[] curvePoints = { point1, point2, point3, point4, point5, point6, point7 };

            g.DrawPolygon(grayPen, grayBrush, curvePoints, XFillMode.Alternate);


            String drawString = $"{-wlRF / 1000,3:0.00#}kN";
            textSize = g.MeasureString(drawString, drawFont);
            if (wlRF != 0)
            {
                if (side == "left")
                {
                    g.DrawString(drawString, drawFont, grayBrush, x - textSize.Width - 3, y + 8.5);
                }
                else
                {
                    g.DrawString(drawString, drawFont, grayBrush, x + 3, y + 8.5);
                }
            }


            drawString = $"{-hllRF / 1000,3:0.00#}kN";
            textSize = g.MeasureString(drawString, drawFont);
            if (hllRF != 0)
            {
                if (side == "left")
                {
                    g.DrawString(drawString, drawFont, blueBrush, x - textSize.Width - 3, y + 18.5);
                }
                else
                {
                    g.DrawString(drawString, drawFont, blueBrush, x + 3, y + 18.5);
                }
            }

            drawString = $"{-dlRF / 1000,3:0.00#}kN";
            textSize = g.MeasureString(drawString, drawFont);
            if (dlRF != 0)
            {
                if (side == "left")
                {
                    g.DrawString(drawString, drawFont, grayBrush, x - textSize.Width - 3, y + 18.5);
                }
                else
                {
                    g.DrawString(drawString, drawFont, grayBrush, x + 3, y + 18.5);
                }
            }
            g.Save();
            g.Dispose();
        }

        public void PdfDrawHorizontalReaction(PdfPage myPage, double x, double y, double dlRF, string side)
        {
            XGraphics g = XGraphics.FromPdfPage(myPage);
            XSize textSize;

            XFont drawFont = new XFont("Univers for Schueco 330 Light", 6, XFontStyle.Regular);
            XPen grayPen = new XPen(XColor.FromArgb(38, 38, 38), .25);
            XPen grayPenThick = new XPen(XColor.FromArgb(102, 102, 102), 1.5);
            XBrush blueBrush = new XSolidBrush(XColor.FromArgb(0, 162, 209));
            XBrush grayBrush = new XSolidBrush(XColor.FromArgb(102, 102, 102));

            XPoint point1 = new XPoint(x, y);
            XPoint point2 = new XPoint(x - 5, y - 1.75);
            XPoint point3 = new XPoint(x - 5, y - 0.75);
            XPoint point4 = new XPoint(x - 15, y - 0.75);
            XPoint point5 = new XPoint(x - 15, y + 0.75);
            XPoint point6 = new XPoint(x - 5, y + 0.75);
            XPoint point7 = new XPoint(x - 5, y + 1.75);

            XPoint[] curvePoints = { point1, point2, point3, point4, point5, point6, point7 };

            g.DrawPolygon(grayPen, grayBrush, curvePoints, XFillMode.Alternate);


            String drawString = $"{dlRF / 1000,3:0.00#}kN (Vertical)";
            textSize = g.MeasureString(drawString, drawFont);
            if (dlRF != 0)
            {
                if (side == "left")
                {
                    g.DrawString(drawString, drawFont, grayBrush, x - textSize.Width / 2, y - 3);
                }
                else
                {
                    g.DrawString(drawString, drawFont, grayBrush, x - textSize.Width / 2, y - 3);
                }
            }

            g.Save();
            g.Dispose();
        }

        public void PdfDrawForce(PdfDocument myTemplate, double xs, double xe, double qs, double qe, double qmax, int chartNo, double yref, bool sameSlope, bool isHLiveLoad = false, bool isConcentrated = false)
        {
            XGraphics g = XGraphics.FromPdfPage(myTemplate.Pages[0]);

            XPen grayPen = new XPen(XColor.FromArgb(38, 38, 38), 0.25);
            XPen bluePen = new XPen(XColor.FromArgb(0, 162, 209), 0.25);
            XBrush grayBrush = new XSolidBrush(XColor.FromArgb(102, 102, 102));
            XBrush blueBrush = new XSolidBrush(XColor.FromArgb(0, 162, 209));

            double xLocation = 100;
            double yLocation = 0;
            double maxTailLength = 40;
            switch (chartNo)
            {
                case 1:
                    yLocation = yref;
                    maxTailLength = isConcentrated ? 40 : 30;
                    break;
                case 2:
                    yLocation = yref;
                    maxTailLength = isConcentrated ? 40 : 30;
                    break;
                case 3:
                    yLocation = yref;
                    maxTailLength = 20;
                    break;
                case 5:
                    yLocation = yref;
                    maxTailLength = 20;
                    break;
            }

            double beamLength = 400;
            double arrowSpaceing = isHLiveLoad ? 40 : 20;
            int numberArrows = Convert.ToInt32(Math.Abs(xe - xs) * beamLength / arrowSpaceing);
            double x, y;
            double slope, force, deltaX;

            // set inital values
            XPen arrowPen = isHLiveLoad ? bluePen : grayPen;
            XBrush arrowBrush = isHLiveLoad ? blueBrush : grayBrush;
            if (xs == xe)
            {
                slope = 0;
                deltaX = 0;
                numberArrows = 0;
            }
            else
            {
                slope = (qe - qs) / (xe - xs);
                arrowSpaceing = (xe - xs) * beamLength / numberArrows;
                deltaX = (xe - xs) / numberArrows;
            }

            // draw the arrows 
            if (numberArrows != 0)
            {
                for (int i = 0; i <= numberArrows; i++)
                {
                    x = xLocation + (xs * beamLength) + i * arrowSpaceing;
                    force = qs + slope * (i * deltaX);
                    y = yLocation - maxTailLength * force / qmax;


                    // draw the arrow tail
                    g.DrawLine(arrowPen, x, yLocation, x, y);

                    // draw the arrow head
                    if (Math.Abs(y - yLocation) > 5)
                    {
                        XPoint point1 = new XPoint(x, yLocation);
                        XPoint point2 = new XPoint(x - 1.5, yLocation - 5);
                        XPoint point3 = new XPoint(x + 1.5, yLocation - 5);
                        XPoint[] curvePoints = { point1, point2, point3 };
                        g.DrawPolygon(arrowPen, arrowBrush, curvePoints, XFillMode.Alternate);
                    }
                }
            }

            // draw the load line
            double x1, x2, y1, y2;
            x1 = xLocation + (xs * beamLength);
            x2 = xLocation + (xe * beamLength);
            y1 = yLocation - maxTailLength * Math.Abs(qs) / qmax;
            y2 = yLocation - maxTailLength * Math.Abs(qe) / qmax;
            if (xs != xe)
            {
                g.DrawLine(arrowPen, x1, y1, x2, y2);
            }

            // draw concentrated load
            if (isConcentrated == true)
            {
                numberArrows = 1;
                x = xLocation + (xs * beamLength);
                force = Math.Abs(qs);
                y = yLocation - maxTailLength * force / qmax;

                if (Math.Abs(yLocation - y) < 10)
                {
                    y = yLocation - 10; // set the minumum value for the concentrated load
                }

                // draw the arrow tail
                g.DrawLine(arrowPen, x, yLocation, x, y);
                // draw the arrow head
                if (Math.Abs(y - yLocation) > 3)
                {
                    XPoint point1 = new XPoint(x, yLocation);
                    XPoint point2 = new XPoint(x - 1.5, yLocation - 5);
                    XPoint point3 = new XPoint(x + 1.5, yLocation - 5);
                    XPoint[] curvePoints = { point1, point2, point3 };
                    g.DrawPolygon(arrowPen, arrowBrush, curvePoints, XFillMode.Alternate);
                }
            }



            // draw the load value
            XFont drawFont = new XFont("Univers for Schueco 330 Light", 6, XFontStyle.Regular);
            if (xs == xe)  // CONCENTRATED LOAD
            {
                x = xLocation + (xs * beamLength);
                y = yLocation - maxTailLength * Math.Abs(qe) / qmax - 2.5;
                String drawString = (qe / 1000).ToString("F3") + "kN";
                g.DrawString(drawString, drawFont, arrowBrush, x, y);
            }
            else
            {

                if (!sameSlope)
                {
                    String drawString = qe.ToString("F3") + "kN/m";
                    if (qe != 0)
                    {
                        double tempSlope = (y2 - y1) / (x2 - x1);
                        if (x1 == 100 && x2 == 100 + beamLength && Math.Abs(tempSlope - 0) < 0.0001)  // the case for only one uniform load over the whole span, plot at the center
                        {
                            g.DrawString(drawString, drawFont, arrowBrush, xLocation + (0.5 * beamLength), y2 - 2.5);
                        }
                        else
                        {
                            g.DrawString(drawString, drawFont, arrowBrush, x2, y2 - 2.5);
                        }

                    }
                }
            }
            g.Save();
            g.Dispose();
        }


        public void PdfDrawDimension(PdfPage myPage, double xStart, double xEnd, int nc, double value, double xLocation, double yLocation, double Length, bool isVertical = false)
        {
            // Draw beam Dimention Line
            XGraphics g = XGraphics.FromPdfPage(myPage);

            if (isVertical)
            {
                XPoint rotationCenter = new XPoint(xLocation, yLocation);
                g.RotateAtTransform(-90, rotationCenter);
                yLocation = yLocation + 20;
            }

            XPen pen = new XPen(XColor.FromArgb(38, 38, 38), .25);
            XBrush brush = new XSolidBrush(XColor.FromArgb(102, 102, 102));

            double width, x, y, beamLength;

            beamLength = Length;

            yLocation -= 5;


            x = xLocation + xStart * beamLength;
            double dimOffset = Length <= 200 ? 10 : 15;
            y = yLocation + nc * dimOffset;
            width = (xEnd - xStart) * beamLength;

            g.DrawLine(pen, x, y, x + width, y);
            g.DrawLine(pen, x, y + 2, x, y - 11);
            g.DrawLine(pen, x + width, y + 2, x + width, y - (nc + 1) * 3);

            XPoint point1 = new XPoint(x, y);
            XPoint point2 = new XPoint(x + 5, y - 1);
            XPoint point3 = new XPoint(x + 5, y + 1);
            XPoint[] curvePoints = { point1, point2, point3 };

            if (width >= 15)
            {
                g.DrawPolygon(pen, brush, curvePoints, XFillMode.Alternate);
            }


            curvePoints[0] = new XPoint(x + width, y);
            curvePoints[1] = new XPoint(x + width - 5, y - 1);
            curvePoints[2] = new XPoint(x + width - 5, y + 1);

            if (width >= 15)
            {
                g.DrawPolygon(pen, brush, curvePoints, XFillMode.Alternate);
            }


            // draw the dimension text
            double fontSize = 6;
            if (xEnd - xStart < 0.05) fontSize = 4;
            XFont drawFont = new XFont("Univers for Schueco 330 Light", fontSize, XFontStyle.Regular);
            String drawString = $"{value / 1000,0:0.00}m";

            XSize size = g.MeasureString(drawString, drawFont);

            double xt = x + width / 2 - size.Width / 2;

            g.DrawString(drawString, drawFont, brush, xt, y - 1.75);

            g.Save();
            g.Dispose();

        }

        public void PdfDrawCoordinate(PDFMemberResult pdfMemberResult, PdfDocument myTemplate, int ChartNo, double yref)
        {
            string imageFilePath = $"{_resourceFolderPath}Image\\YZCoordinate.PNG";
            if (ChartNo == 5)
            {
                imageFilePath = $"{_resourceFolderPath}Image\\XZCoordinate.PNG";
            }

            XGraphics g = XGraphics.FromPdfPage(myTemplate.Pages[0]);
            XImage image = XImage.FromFile(imageFilePath);
            double xLocation = 64;
            double yLocation = yref - 30;

            g.DrawImage(image, xLocation, yLocation, 28, 32);

            g.Save();
            g.Dispose();

        }


        public void PdfDrawProfileImage(PDFMemberResult pdfMemberResult, PdfDocument myTemplate)
        {
            string imageFilePath;

            if (string.IsNullOrEmpty(pdfMemberResult.ReinfArticleName))
            {
                imageFilePath = $"{_resourceFolderPath}article-jpeg\\{pdfMemberResult.ArticleName}.jpg";
            }
            else
            {
                imageFilePath = $"{_resourceFolderPath}article-jpeg\\{pdfMemberResult.ArticleName}_{pdfMemberResult.ReinfArticleName}.jpg";
                if (!File.Exists(imageFilePath))
                {
                    imageFilePath = $"{_resourceFolderPath}article-jpeg\\{pdfMemberResult.ArticleName}.jpg";
                }
            }

            if (!File.Exists(imageFilePath))
            {
                imageFilePath = $"{_resourceFolderPath}article-jpeg\\CustomArticle.jpg";
            }


            XGraphics g = XGraphics.FromPdfPage(myTemplate.Pages[0]);
            XImage image = XImage.FromFile(imageFilePath);
            double xLocation = 430;
            double yLocation = 150;
            double scaledWidth, scaledHeight;
            double imageAspectRatio = image.PointWidth / image.PointHeight;
            if (imageAspectRatio > 1)
            {
                scaledWidth = 100;
                scaledHeight = scaledWidth / imageAspectRatio;
            }
            else
            {
                scaledHeight = 100;
                scaledWidth = imageAspectRatio * scaledHeight;
            }

            g.DrawImage(image, xLocation + 50 - scaledWidth / 2, yLocation, scaledWidth, scaledHeight);

            //// draw title and parameters
            //XBrush grayBrush = new XSolidBrush(XColor.FromArgb(102, 102, 102));
            //XFont drawFont = new XFont("Univers for Schueco 330 Light", 9, XFontStyle.Regular);
            ////double xs = xLocation + 137;
            ////double ys = yLocation;
            ////int iStart = imageFilePath.LastIndexOf(@"\");
            ////string drawString = imageFilePath.Substring(iStart+1);
            ////g.DrawString(drawString, drawFont, grayBrush, xs, ys);

            //double xs = xLocation + 160;
            //double ys = yLocation;
            //string drawString = $"{pdfMemberResult.Iy,-4:0.00}";
            //g.DrawString(drawString, drawFont, grayBrush, xs, ys);
            //ys = ys + 15;
            //drawString = $"{pdfMemberResult.Il,-4:0.00}";
            //g.DrawString(drawString, drawFont, grayBrush, xs, ys);
            //ys = ys + 15;
            //drawString = $"{pdfMemberResult.Is,-4:0.00}";
            //g.DrawString(drawString, drawFont, grayBrush, xs, ys);
            //ys = ys + 15;
            //drawString = $"{pdfMemberResult.Iv,-4:0.00}";
            //g.DrawString(drawString, drawFont, grayBrush, xs, ys);
            //ys = ys + 15;
            //drawString = $"{pdfMemberResult.v,4:0.0%}";
            //g.DrawString(drawString, drawFont, grayBrush, xs, ys);
            //ys = ys + 15;
            //drawString = $"{pdfMemberResult.lamda,4:0.00}";
            //g.DrawString(drawString, drawFont, grayBrush, xs, ys);

            g.Save();
            g.Dispose();

        }

        public void PdfDrawResultCurves(PDFMemberResult pdfMemberResult, PdfDocument myTemplate)
        {
            XGraphics g = XGraphics.FromPdfPage(myTemplate.Pages[1]);
            bool isTransom = pdfMemberResult.MemberType == 5;

            double x0w = 96;
            double x0d = 349;
            double y0w = 130.7;
            double y0l = 201.8;

            double Yoffset = 163;
            double graphXsize = isTransom ? 190.5 : 438;
            double graphYsize = 31;

            double x0, y0;

            int curveCount = pdfMemberResult.CurveX.Length - 1;
            double[] curveX = pdfMemberResult.CurveX;
            double xLength = curveX[curveCount];

            //for Tv  -- wind
            x0 = x0w;
            y0 = y0w;
            double[] curveY = pdfMemberResult.ShearCurveWind.Select(x => x / 1000).ToArray();
            PdfDrawCurve(g, x0, y0, graphXsize, graphYsize, curveX, curveY);

            //for Tv  -- HLL
            x0 = x0w;
            y0 = y0l;
            curveY = pdfMemberResult.ShearCurveHLL.Select(x => x / 1000).ToArray();
            PdfDrawCurve(g, x0, y0, graphXsize, graphYsize, curveX, curveY);

            //for Tv  -- Dead Load
            if (isTransom)
            {
                x0 = x0d;
                y0 = y0w;
                curveY = pdfMemberResult.ShearCurveWeight.Select(x => x / 1000).ToArray();
                PdfDrawCurve(g, x0, y0, graphXsize, graphYsize, curveX, curveY);
            }

            //for moment   -- wind
            x0 = x0w;
            y0 = y0w + Yoffset;
            curveY = pdfMemberResult.MomentCurveWind.Select(x => x / 10000).ToArray();
            PdfDrawCurve(g, x0, y0, graphXsize, graphYsize, curveX, curveY);

            //for moment   -- HLL
            x0 = x0w;
            y0 = y0l + Yoffset;
            curveY = pdfMemberResult.MomentCurveHLL.Select(x => x / 10000).ToArray();
            PdfDrawCurve(g, x0, y0, graphXsize, graphYsize, curveX, curveY);

            //for moment   -- Dead load
            if (isTransom)
            {
                x0 = x0d;
                y0 = y0w + Yoffset;
                curveY = pdfMemberResult.MomentCurveWeight.Select(x => x / 10000).ToArray();
                PdfDrawCurve(g, x0, y0, graphXsize, graphYsize, curveX, curveY);
            }

            //for out-of - plane displacement
            x0 = x0w;
            y0 = isTransom ? y0w + 2 * Yoffset + 62 : y0w + 2 * Yoffset + 37.5;
            curveY = pdfMemberResult.OutOfPlaneDeflectionCurveWind.Select(x => x).ToArray();
            PdfDrawCurve(g, x0, y0, graphXsize, graphYsize, curveX, curveY);

            //for In-plane displacement
            if (isTransom)
            {
                x0 = x0d;
                y0 = y0w + 2 * Yoffset + 62;
                curveY = pdfMemberResult.InPlaneDeflectionCurveWeight.Select(x => x).ToArray();
                PdfDrawCurve(g, x0, y0, graphXsize, graphYsize, curveX, curveY);
            }
            g.Save();
            g.Dispose();
        }

        private void PdfDrawCurve(XGraphics g, double x0, double y0, double graphXsize, double graphYsize, double[] curveX, double[] curveY)
        {
            XPen Pen_lightgray = new XPen(XColor.FromArgb(211, 211, 211), 0.5);
            Pen_lightgray.DashStyle = XDashStyle.Dot;
            XPen Pen_darkgray = new XPen(XColor.FromArgb(128, 128, 128), 1);
            XPen Pen_blue = new XPen(XColor.FromArgb(0, 162, 209), 1);


            XBrush blackFontBrush = new XSolidBrush(XColor.FromArgb(0, 0, 0));
            XFont drawFont = new XFont("Univers for Schueco 330 Light", 8);

            // draw grid x lines
            int XsectNo = 10;
            for (int i = 0; i <= XsectNo; i++)
            {
                XPen pen = Pen_lightgray;
                if (i == 0 || i == XsectNo) pen = Pen_darkgray;
                g.DrawLine(pen, x0 + i * graphXsize / XsectNo, y0 - graphYsize, x0 + i * graphXsize / XsectNo, y0 + graphYsize);
            }

            // draw grid Y lines
            // draw y=0 grid line first
            g.DrawLine(Pen_darkgray, x0, y0, x0 + graphXsize, y0);
            int YsectNo = 4;
            for (int i = 1; i <= YsectNo; i++)
            {
                XPen pen = Pen_lightgray;
                if (i == YsectNo) pen = Pen_darkgray;
                g.DrawLine(pen, x0, y0 + i * graphYsize / YsectNo, x0 + graphXsize, y0 + i * graphYsize / YsectNo);
                g.DrawLine(pen, x0, y0 - i * graphYsize / YsectNo, x0 + graphXsize, y0 - i * graphYsize / YsectNo);
            }

            // draw Y value text
            double yMax = curveY.Select(x => Math.Abs(x)).Max();
            double yLimit = 0;
            if (yMax < 30)
            {
                yLimit = Math.Max(Math.Ceiling(yMax / 4) * 4, 1);
            }
            else
            {
                yLimit = Math.Max(Math.Ceiling(yMax / 20) * 20, 1);
            }

            g.DrawString($"{yLimit,3:0.#}", drawFont, blackFontBrush, x0 - 18, y0 - graphYsize + 3);
            g.DrawString($"{-yLimit,3:0.#}", drawFont, blackFontBrush, x0 - 18, y0 + graphYsize + 3);

            // draw Curve
            int curveCount = curveX.Length - 1;
            double xLength = curveX[curveCount];
            double x1, x2, y1, y2;
            for (int j = 1; j <= curveCount; j++)
            {
                x1 = x0 + curveX[j - 1] / xLength * graphXsize;
                x2 = x0 + curveX[j] / xLength * graphXsize;
                y1 = y0 - curveY[j - 1] / yLimit * graphYsize;
                y2 = y0 - curveY[j] / yLimit * graphYsize;
                g.DrawLine(Pen_blue, x1, y1, x2, y2);
            }
        }

        public void PdfDrawSectionTitle(PdfPage myPage, int i)
        {
            XGraphics g = XGraphics.FromPdfPage(myPage);

            // draw subsection title
            double xs = 68;
            double ys = 143;
            string drawString = $"{i + 1}";
            XBrush blackBrush = new XSolidBrush(XColor.FromArgb(0, 0, 0));
            XFont drawFont = new XFont("Univers for Schueco 330 Light", 10, XFontStyle.Bold);
            g.DrawString(drawString, drawFont, blackBrush, xs, ys);

            // draw section title
            //if (i == 0)
            //{
            //    xs = 58;
            //    ys = 120;
            //    drawString = "6. Result";
            //    drawFont = new XFont("Univers for Schueco 330 Light", 14, XFontStyle.Bold);
            //    g.DrawString(drawString, drawFont, blackBrush, xs, ys);
            //}

            g.Save();
            g.Dispose();
        }


        public void PdfDrawPageNumber(PdfDocument tempDoc)
        {
            for (int i = 1; i < tempDoc.Pages.Count; i++)
            {
                PdfPage tempPage = tempDoc.Pages[i];
                XGraphics g = XGraphics.FromPdfPage(tempPage);
                XBrush blackBrush = new XSolidBrush(XColor.FromArgb(0, 0, 0));
                XFont drawFont = new XFont("Univers for Schueco 330 Light", 10, XFontStyle.Bold);
                double xs = 290, ys = 810;
                string drawString = $"{i}/{tempDoc.Pages.Count - 1}";
                g.DrawString(drawString, drawFont, blackBrush, xs, ys);
                g.Save();
                g.Dispose();
            }
        }




        public void PdfChangeFieldName(PdfDocument tempDoc, int subSectionNo)
        {
            for (int i = 0; i < tempDoc.AcroForm.Fields.Count(); i++)
            {
                PdfAcroField field = tempDoc.AcroForm.Fields[i];
                if (field.Name == "ProjectName" || field.Name == "Location" || field.Name == "User" || field.Name == "Date") continue;
                string fieldName = $"subSection{subSectionNo}_" + field.Name;
                field.Elements.SetString("/T", fieldName);
            }
        }


        public string CreateNewPdf(string source, string destination)
        {
            var fopen = File.Open(destination, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None);
            fopen.Close();

            destination = Path.GetFullPath(destination);
            File.Copy(source, destination, true);
            return destination;
        }

        public void InsertValue(PdfDocument myTemplate, string fieldName, string fieldValue, bool isRed = false)
        {
            try
            {

                PdfAcroForm form = myTemplate.AcroForm;

                if (form.Elements.ContainsKey("/NeedAppearances"))
                {
                    form.Elements["/NeedAppearances"] = new PdfSharp.Pdf.PdfBoolean(true);
                }
                else
                {
                    form.Elements.Add("/NeedAppearances", new PdfSharp.Pdf.PdfBoolean(true));
                }

                // Get all form fields of the whole document

                PdfAcroField.PdfAcroFieldCollection fields = form.Fields;

                // this sets the value for the field selected
                if (fieldValue == null) fieldValue = "";

                PdfAcroField field = fields[fieldName];
                PdfTextField txtField;
                if ((txtField = field as PdfTextField) != null)
                {
                    txtField.ReadOnly = false;

                    if (fieldValue.Contains(" > 1.0     NG") || isRed)
                    {
                        txtField.Elements.SetString(PdfTextField.Keys.DA, "/UniversForSchueco-630Bold 6 Tf 1 0 0 rg");
                    }
                    txtField.Value = new PdfString(fieldValue);
                    txtField.ReadOnly = true;
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        string _resourceFolderPath;
    }
}
