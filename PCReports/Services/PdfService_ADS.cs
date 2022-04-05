using PdfSharp.Pdf;
using PdfSharp.Drawing;
using PdfSharp.Pdf.AcroForms;
using PdfSharp.Pdf.IO;
using PdfSharp.Pdf.Security;
using PCReports.Models.ADS;

namespace PCReports.Services.ADS
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

    {
        // --------- Main entery point ------------------------------------------------
        public PdfService()
        {
            _resourceFolderPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"Resources\");
        }

        public string GenerateReport(string pdfFilePath, PDFResult pdfResult)
        {
            string reportGuid = Path.GetFileNameWithoutExtension(pdfFilePath);
            string output = pdfFilePath;
            var source = string.Empty;

            // copy template to destination
            string source1 = _resourceFolderPath + @"templates\SPS ADS Structural report tempelate_Part1 - EN.pdf";
            string source2 = _resourceFolderPath + @"templates\SPS ADS Structural report tempelate_Part2 - EN.pdf";
            string source3 = _resourceFolderPath + @"templates\SPS ADS Structural report tempelate_Part3 - EN.pdf";
            if (Thread.CurrentThread.CurrentCulture.Name == "de-DE")
            {
                source1 = _resourceFolderPath + @"templates\SPS ADS Structural report tempelate_Part1 - DE.pdf";
                source2 = _resourceFolderPath + @"templates\SPS ADS Structural report tempelate_Part2 - DE.pdf";
                source3 = _resourceFolderPath + @"templates\SPS ADS Structural report tempelate_Part3 - DE.pdf";
            }
            if (Thread.CurrentThread.CurrentCulture.Name == "fr-FR")
            {
                source1 = _resourceFolderPath + @"templates\SPS ADS Structural report tempelate_Part1 - FR.pdf";
                source2 = _resourceFolderPath + @"templates\SPS ADS Structural report tempelate_Part2 - FR.pdf";
                source3 = _resourceFolderPath + @"templates\SPS ADS Structural report tempelate_Part3 - FR.pdf";
            }

            string destination = CreateNewPdf(source1, output);

            // Create tempPDFfile for each member
            if (!string.IsNullOrEmpty(destination))
            {
                // open the Pdf file, fill general data
                PdfDocument myReport = PdfReader.Open(output, PdfDocumentOpenMode.Modify);
                PdfInputGeneralData(pdfResult, myReport);
                PdfDrawStructure(pdfResult, myReport);

                PdfDocument source2Document = PdfReader.Open(source2, PdfDocumentOpenMode.Import);
                PdfDocument source3Document = PdfReader.Open(source3, PdfDocumentOpenMode.Import);

                //Directory.CreateDirectory(Path.Combine(_resourceFolderPath, "temp\\"));
                int memberCount = pdfResult.PDFMemberResults.Count();
                for (int i = 0; i < memberCount; i++)
                {
                    string tempFileLoc;
                    string tempPDF;
                    PdfDocument tempDoc;

                    // draw load cases
                    PDFMemberResult pdfMemberResult = pdfResult.PDFMemberResults[i];
                    if (pdfMemberResult.MemberType == 3)
                    {
                        // create temp file
                        //tempFileLoc = Path.GetFullPath(Path.Combine(_resourceFolderPath, $"temp\\tempPDF{Guid.NewGuid().ToString()}.pdf"));
                        tempFileLoc = Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, $"Content\\temp\\tempPDF{Guid.NewGuid().ToString()}.pdf"));
                        tempPDF = CreateNewPdf(source3, tempFileLoc);
                        tempDoc = PdfReader.Open(tempPDF, PdfDocumentOpenMode.Modify);
                    }
                    else
                    {
                        // create temp file
                        //tempFileLoc = Path.GetFullPath(Path.Combine(_resourceFolderPath, $"temp\\tempPDF{Guid.NewGuid().ToString()}.pdf"));
                        tempFileLoc = Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, $"Content\\temp\\tempPDF{Guid.NewGuid().ToString()}.pdf"));
                        tempPDF = CreateNewPdf(source2, tempFileLoc);
                        tempDoc = PdfReader.Open(tempPDF, PdfDocumentOpenMode.Modify);
                    }

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

                    AddReportGuidToFields(tempDoc, reportGuid);

                    tempDoc.Save(tempFileLoc);
                    tempDoc.Close();
                    tempDoc = PdfReader.Open(tempFileLoc, PdfDocumentOpenMode.Import);
                    PdfPage tempPage = tempDoc.Pages[0];
                    myReport.AddPage(tempPage);
                    tempPage = tempDoc.Pages[1];
                    myReport.AddPage(tempPage);
                    tempPage = tempDoc.Pages[2];
                    myReport.AddPage(tempPage);
                    tempDoc.Close();
                    File.Delete(tempFileLoc);
                }
                source2Document.Close();
                source3Document.Close();

                // draw page number
                PdfDrawPageNumber(myReport);

                // add report stamp to all fields
                AddReportGuidToFields(myReport, reportGuid);

                // flatten and close the pdf file
                PdfSecuritySettings securitySettings = myReport.SecuritySettings;
                securitySettings.PermitFormsFill = false;

                // Setting one of the passwords automatically sets the security level to
                // PdfDocumentSecurityLevel.Encrypted128Bit.
                //securitySettings.UserPassword = "vcl";
                //securitySettings.OwnerPassword = "tjdeganyar";

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
                myReport.Save(output);
                myReport.Close();
            }

            return destination;
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
            InsertValue(myTemplate, "ProjectName", pdfResult.ProjectName);
            InsertValue(myTemplate, "Location", pdfResult.Location);
            InsertValue(myTemplate, "Date", DateTime.Now.ToString(dateFormat));
            InsertValue(myTemplate, "User", pdfResult.UserName);

            InsertValue(myTemplate, "UserNotes", pdfResult.UserNotes);

            InsertValue(myTemplate, "ProfileSystem", pdfResult.ProfileSystem);
            InsertValue(myTemplate, "FrameProfile", pdfResult.FrameProfile);
            InsertValue(myTemplate, "FrameProfileWeight", $"{pdfResult.FrameProfileWeight * 100,3:0.#} N/m");  // convert N/cm to N/m
            InsertValue(myTemplate, "TransomProfile", pdfResult.TransomProfile);
            InsertValue(myTemplate, "TransomProfileWeight", $"{pdfResult.TransomProfileWeight * 100,3:0.#} N/m");  // convert N/cm to N/m
            InsertValue(myTemplate, "MullionProfile", pdfResult.MullionProfile);
            InsertValue(myTemplate, "MullionProfileWeight", $"{pdfResult.MullionProfileWeight * 100,3:0.#} N/m");  // convert N/cm to N/m
            InsertValue(myTemplate, "BlockDistance", $"{pdfResult.BlockDistance} mm");

            int showGlassCount = Math.Min(pdfResult.GlassTypes.Count(), 5);
            for (int i = 0; i < showGlassCount; i++)
            {
                InsertValue(myTemplate, $"Glass.{i}", $"{pdfResult.GlassTypes[i].GlassIDs}            {pdfResult.GlassTypes[i].Weight}      {pdfResult.GlassTypes[i].Description}");
            }
            if (showGlassCount < pdfResult.GlassTypes.Count())
            {
                InsertValue(myTemplate, $"Glass.Overflow", "...");
            }

            InsertValue(myTemplate, "WindLoad", $"{pdfResult.WindLoad,4:0.00}");
            InsertValue(myTemplate, "CpeString", pdfResult.CpeString);
            InsertValue(myTemplate, "pCpiString", pdfResult.pCpiString);
            InsertValue(myTemplate, "nCpiString", pdfResult.nCpiString);
            InsertValue(myTemplate, "HorizontalLiveLoad", $"{pdfResult.HorizontalLiveLoad,4:0.00}");
            InsertValue(myTemplate, "HorizontalLiveLoadHeight", $"{pdfResult.HorizontalLiveLoadHeight,4:0.##}");
            InsertValue(myTemplate, "SummerTempDiff", $"{pdfResult.SummerTempDiff,2}");
            InsertValue(myTemplate, "WinterTempDiff", $"{pdfResult.WinterTempDiff,2}");
            InsertValue(myTemplate, "WindLoadFactor", $"{pdfResult.WindLoadFactor,4:0.00}");
            InsertValue(myTemplate, "HorizontalLiveLoadFactor", $"{pdfResult.HorizontalLiveLoadFactor,4:0.00}");
            InsertValue(myTemplate, "DeadLoadFactor", $"{pdfResult.DeadLoadFactor,4:0.00}");

            InsertValue(myTemplate, "TemperatureLoadFactor", $"{pdfResult.TemperatureLoadFactor,4:0.00}");

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


            InsertValue(myTemplate, "AllowableDeflectionLine1", $"{pdfResult.AllowableDeflectionLine1}");
            InsertValue(myTemplate, "AllowableDeflectionLine2", $"{pdfResult.AllowableDeflectionLine2}");
            if (pdfResult.PDFMemberResults.Any(x => x.MemberType == 3))
            {
                InsertValue(myTemplate, "AllowableInplaneDeflectionLine", $"{pdfResult.AllowableInplaneDeflectionLine}");
            }
            else
            {
                InsertValue(myTemplate, "AllowableInplaneDeflectionLine", "--");
            }

            InsertValue(myTemplate, "Alloys", pdfResult.Alloys);
            InsertValue(myTemplate, "Beta", $"{pdfResult.Beta,3}");

            string insulatingBarType = "";
            // Need to check the translations
            switch (pdfResult.InsulatingBarType)
            {
                case "0":
                case "Polythermid Coated Before":
                    if (Thread.CurrentThread.CurrentCulture.Name == "fr-FR") { insulatingBarType = "(PT) Revêtement avant"; }
                    else if (Thread.CurrentThread.CurrentCulture.Name == "de-DE") { insulatingBarType = "Beschichtung vor Verbund (PT)"; }
                    else { insulatingBarType = "Coated Before rolling process (PT)"; }
                    break;
                case "1":
                case "Polythermid Anodized Before":
                    if (Thread.CurrentThread.CurrentCulture.Name == "fr-FR") { insulatingBarType = "(PT) Anodisation avant"; }
                    else if (Thread.CurrentThread.CurrentCulture.Name == "de-DE") { insulatingBarType = "Eloxal vor Verbund (PT)"; }
                    else { insulatingBarType = "Anodized before rolling process (PT)"; }
                    break;
                case "2":
                case "Polyamide Coated Before":
                    if (Thread.CurrentThread.CurrentCulture.Name == "fr-FR") { insulatingBarType = "(PA) Laquée avant"; }
                    else if (Thread.CurrentThread.CurrentCulture.Name == "de-DE") { insulatingBarType = "Beschichtung vor Verbund (PA)"; }
                    else { insulatingBarType = "Coated before rolling process (PA)"; }
                    break;
                case "3":
                case "Polyamide Coated After":
                    if (Thread.CurrentThread.CurrentCulture.Name == "fr-FR") { insulatingBarType = "(PA) Revêtement après"; }
                    else if (Thread.CurrentThread.CurrentCulture.Name == "de-DE") { insulatingBarType = "Beschichtung nach Verbund (PA)"; }
                    else { insulatingBarType = "Coated after rolling process (PA)"; }
                    break;
                case "4":
                case "Polyamide Anodized Before":
                    if (Thread.CurrentThread.CurrentCulture.Name == "fr-FR") { insulatingBarType = "(PA) Anodisation avant"; }
                    else if (Thread.CurrentThread.CurrentCulture.Name == "de-DE") { insulatingBarType = "Eloxal vor Verbund (PA)"; }
                    else { insulatingBarType = "Anodized before rolling process (PA)"; }
                    break;
                case "5":
                case "Polyamide Anodized After":
                    if (Thread.CurrentThread.CurrentCulture.Name == "fr-FR") { insulatingBarType = "(PA) Anodisé après"; }
                    else if (Thread.CurrentThread.CurrentCulture.Name == "de-DE") { insulatingBarType = "Eloxal nach Verbund (PA)"; }
                    else { insulatingBarType = "Anodized after rolling process (PA)"; }
                    break;
                case "User Defined":
                    if (Thread.CurrentThread.CurrentCulture.Name == "fr-FR") { insulatingBarType = "Défini par l'utilisateur"; }
                    else if (Thread.CurrentThread.CurrentCulture.Name == "de-DE") { insulatingBarType = "User Defined"; }
                    else { insulatingBarType = "User Defined"; }
                    break;
                default:
                    if (Thread.CurrentThread.CurrentCulture.Name == "fr-FR") { insulatingBarType = "(PT) Revêtement avant"; }
                    else if (Thread.CurrentThread.CurrentCulture.Name == "de-DE") { insulatingBarType = "Beschichtung vor Verbund (PT)"; }
                    else { insulatingBarType = "Coated Before rolling process(PT)"; }
                    break;
            }


            InsertValue(myTemplate, "InsulatingBarType", insulatingBarType);

            InsertValue(myTemplate, "InsulatingBarDataNote", pdfResult.InsulatingBarDataNote);
            InsertValue(myTemplate, "RSn20", $"{pdfResult.RSn20,3}");
            InsertValue(myTemplate, "RSp80", $"{pdfResult.RSp80,3}");
            InsertValue(myTemplate, "RTn20", $"{pdfResult.RTn20,3}");
            InsertValue(myTemplate, "RTp80", $"{pdfResult.RTp80,3}");
            InsertValue(myTemplate, "Cn20", $"{pdfResult.Cn20,3}");
            InsertValue(myTemplate, "Cp20", $"{pdfResult.Cp20,3}");
            InsertValue(myTemplate, "Cp80", $"{pdfResult.Cp80,3}");
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
            InsertValue(myTemplate, "ArticleNo", $"{pdfMemberResult.ArticleName}");

            InsertValue(myTemplate, "Depth", $"{pdfMemberResult.Depth,4:0.00}");
            InsertValue(myTemplate, "Length", $"{pdfMemberResult.Length,4:0.00}");
            InsertValue(myTemplate, "Weight", $"{pdfMemberResult.Weight * 100,4:0.00}");
            InsertValue(myTemplate, "Iy", $"{pdfMemberResult.Iy,4:0.00}");
            InsertValue(myTemplate, "Il", $"{pdfMemberResult.Il,4:0.00}");
            InsertValue(myTemplate, "Is", $"{pdfMemberResult.Is,4:0.00}");
            InsertValue(myTemplate, "Iv", $"{pdfMemberResult.Iv,4:0.00}");
            InsertValue(myTemplate, "v", $"{pdfMemberResult.v,4:0.0%}");
            InsertValue(myTemplate, "TributaryArea", $"{pdfMemberResult.TributaryArea,3:0.0}");  //m2
            InsertValue(myTemplate, "MemberWindLoad", $"{pdfMemberResult.AppliedWindLoad,4:0.00}");
            InsertValue(myTemplate, "Cp", $"{pdfMemberResult.Cp,4:0.00}");
            InsertValue(myTemplate, "lamdan20", $"{pdfMemberResult.lamdan20,4:0.00}");
            InsertValue(myTemplate, "lamdap20", $"{pdfMemberResult.lamdap20,4:0.00}");
            InsertValue(myTemplate, "lamdap80", $"{pdfMemberResult.lamdap80,4:0.00}");
            InsertValue(myTemplate, "ProjectName", pdfResult.ProjectName);
            InsertValue(myTemplate, "Location", pdfResult.Location);
            InsertValue(myTemplate, "Date", DateTime.Now.ToString(dateFormat));
            InsertValue(myTemplate, "User", pdfResult.UserName);
            double[] GeometricInfo = pdfResult.MemberGeometricInfos.Single(item => item.MemberID == pdfMemberResult.MemberID).PointCoordinates;

            for (int i = 0; i < 4; i++)
            {
                for (int j = 0; j < 4; j++)
                    // convert moment unit from kNm to kNcm
                    InsertValue(myTemplate, $"MomentMatrix.{i}.{j}", $"{pdfMemberResult.MomentMatrix[i, j] * 100,6:0.0#}");
            }

            for (int i = 0; i < 10; i++)
            {
                for (int j = 0; j < 5; j++)
                    InsertValue(myTemplate, $"StressMatrix.{i}.{j}", $"{pdfMemberResult.StressMatrix[i, j],6:0.0#}");
            }


            //////////////////////////
            ///




            if (pdfMemberResult.MemberType == 3)
            {
                InsertValue(myTemplate, "MomaxDead", $"{pdfMemberResult.MomentMatrix[4, 0] * 100,4:0.00}");
                InsertValue(myTemplate, "MumaxDead", $"{pdfMemberResult.MomentMatrix[4, 1] * 100,4:0.00}");

                InsertValue(myTemplate, "SoDead", $"{pdfMemberResult.StressMatrix[10, 0],4:0.00}");
                InsertValue(myTemplate, "SuDead", $"{pdfMemberResult.StressMatrix[10, 1],4:0.00}");

                InsertValue(myTemplate, "SoLC3", $"{pdfMemberResult.StressMatrix[11, 0],4:0.00}");
                InsertValue(myTemplate, "SuLC3", $"{pdfMemberResult.StressMatrix[11, 1],4:0.00}");

            }
            //////////////////////////
            ///





            if (pdfMemberResult.MaxInplaneDeflection == 0)
            {
                InsertValue(myTemplate, "VerticalDisplacement", $"--");
                InsertValue(myTemplate, "InplaneDispIndexLoc2", $"--");
                InsertValue(myTemplate, "AllowableVerticalDisplacement", $"--");
                InsertValue(myTemplate, "VerticalDisplacementRatio", $"--");
            }
            else
            {
                InsertValue(myTemplate, "VerticalDisplacement", $"{pdfMemberResult.MaxInplaneDeflection * 10,3:0.00} mm");
                InsertValue(myTemplate, "InplaneDispIndexLoc2", $"{Math.Round(pdfMemberResult.InplaneDispIndex)}");
                InsertValue(myTemplate, "AllowableVerticalDisplacement", $"{pdfMemberResult.AllowableInplaneDeflecton * 10,3:0.00} mm");
                string strInplaneDeflecionRatio = $"{pdfMemberResult.InplaneDeflectionRatio,3:P1}";

                if (Thread.CurrentThread.CurrentCulture.Name.Equals("de-DE"))
                {
                    if (pdfMemberResult.MemberType == 3)
                    {
                        PdfCheckCriteria(myTemplate, strInplaneDeflecionRatio, 425, 676);
                    }
                    else
                    {
                        PdfCheckCriteria(myTemplate, strInplaneDeflecionRatio, 419, 658);
                    }
                }
                else
                {
                    if (pdfMemberResult.MemberType == 3)
                    {
                        PdfCheckCriteria(myTemplate, strInplaneDeflecionRatio, 391, 678);
                    }
                    else
                    {
                        PdfCheckCriteria(myTemplate, strInplaneDeflecionRatio, 391, 668);
                    }
                }
            }

            InsertValue(myTemplate, "Displacement", $"{pdfMemberResult.MaxOutofplaneDeflection * 10,3:0.00} mm");
            InsertValue(myTemplate, "AllowableDisplacement", $"{pdfMemberResult.AllowableOutofplaneDeflecton * 10,3:0.00} mm");
            string strOutofplaneDeflecionRatio = $"{pdfMemberResult.OutofplaneDeflectionRatio,3:P1}";

            if (Thread.CurrentThread.CurrentCulture.Name.Equals("de-DE"))
            {
                if (pdfMemberResult.MemberType == 3)
                {
                    PdfCheckCriteria(myTemplate, strOutofplaneDeflecionRatio, 140, 679);
                }
                else
                {
                    PdfCheckCriteria(myTemplate, strOutofplaneDeflecionRatio, 140, 675);
                }
            }
            else
            {
                if (pdfMemberResult.MemberType == 3)
                {
                    PdfCheckCriteria(myTemplate, strOutofplaneDeflecionRatio, 140, 680);
                }
                else
                {
                    PdfCheckCriteria(myTemplate, strOutofplaneDeflecionRatio, 132, 675);
                }
            }



            if (pdfMemberResult.MemberType == 3)
            {
                string tempstring = $"{pdfMemberResult.strStressRatio,3:P1}";
                PdfCheckCriteria(myTemplate, tempstring, 165, 487);

                tempstring = $"{pdfMemberResult.strSummerShearRatio,3:P1}";
                PdfCheckCriteria(myTemplate, tempstring, 210, 472 + 37);
                tempstring = $"{pdfMemberResult.strWinterShearRatio,3:P1}";
                PdfCheckCriteria(myTemplate, tempstring, 210, 495 + 37);
                tempstring = $"{pdfMemberResult.strSummerTransverseRatio,3:P1}";
                PdfCheckCriteria(myTemplate, tempstring, 210, 517 + 37);
                tempstring = $"{pdfMemberResult.strWinterTransverseRatio,3:P1}";
                PdfCheckCriteria(myTemplate, tempstring, 210, 540 + 37);

                tempstring = $"{pdfMemberResult.strSummerCompositeRatio,3:P1}";
                PdfCheckCriteria(myTemplate, tempstring, 240, 700);
                tempstring = $"{pdfMemberResult.strWinterCompositeRatio,3:P1}";
                PdfCheckCriteria(myTemplate, tempstring, 240, 722.8);
            }
            else
            {
                string tempstring = $"{pdfMemberResult.strStressRatio,3:P1}";
                PdfCheckCriteria(myTemplate, tempstring, 162, 464);
                tempstring = $"{pdfMemberResult.strSummerShearRatio,3:P1}";
                PdfCheckCriteria(myTemplate, tempstring, 210, 489);
                tempstring = $"{pdfMemberResult.strWinterShearRatio,3:P1}";
                PdfCheckCriteria(myTemplate, tempstring, 210, 511);
                tempstring = $"{pdfMemberResult.strSummerTransverseRatio,3:P1}";
                PdfCheckCriteria(myTemplate, tempstring, 210, 533);
                tempstring = $"{pdfMemberResult.strWinterTransverseRatio,3:P1}";
                PdfCheckCriteria(myTemplate, tempstring, 210, 556);
                tempstring = $"{pdfMemberResult.strSummerCompositeRatio,3:P1}";
                PdfCheckCriteria(myTemplate, tempstring, 236, 699);
                tempstring = $"{pdfMemberResult.strWinterCompositeRatio,3:P1}";
                PdfCheckCriteria(myTemplate, tempstring, 236, 725);
            }


            PdfChangeFieldName(myTemplate, subSectionNo);
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

        public void PdfCheckCriteria(PdfDocument myTemplate, string ratioString, double xs, double ys)
        {
            XGraphics g = XGraphics.FromPdfPage(myTemplate.Pages[2]);
            if (ratioString.Last().Equals('<'))
            {
                XBrush blackBrush = new XSolidBrush(XColor.FromArgb(0, 0, 0));
                XFont drawFont = new XFont("Univers for Schueco 330 Light", 9, XFontStyle.Regular);
                g.DrawString(ratioString + "1  OK", drawFont, blackBrush, xs, ys);
            }
            else
            {
                XBrush redBrush = new XSolidBrush(XColor.FromArgb(255, 0, 0));
                XFont drawFont = new XFont("Univers for Schueco 330 Light", 9, XFontStyle.Regular);
                g.DrawString(ratioString + "1  NG", drawFont, redBrush, xs, ys);
            }
            g.Save();
            g.Dispose();
        }

        // --------- Graphic routines --------------------------------------------------
        public void PdfDrawStructure(PDFResult pdfResult, PdfDocument myTemplate, int pageNumber = 1, double xLocation = 320, double yLocation = 300)
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

            double ImageHeight, ImageWidth;

            double modelAspectRatio = pdfResult.ModelWidth / pdfResult.ModelHeight;
            if (modelAspectRatio > 1)
            {
                ImageWidth = 175;
                yLocation = 300 - (175 - ImageWidth / modelAspectRatio);
                ImageHeight = ImageWidth / modelAspectRatio;
            }
            else
            {
                ImageHeight = 175;
                ImageWidth = modelAspectRatio * ImageHeight;
            }
            double xs = 0, ys = 0;

            XRect rectMemberID, rectGlassID, rectGlass, rectVent;
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
                if (memberGeometricInfo.MemberType == 1)
                {
                    double offA = memberGeometricInfo.offsetA;
                    double offB = memberGeometricInfo.offsetB;
                    if (memberGeometricInfo.outerFrameSide == 0)  // vertical outer frame
                    {
                        x4 = x1 + memberGeometricInfo.width;
                        y4 = y1 + offA;
                        x3 = x4;
                        y3 = y2 + offB;
                    }
                    else if (memberGeometricInfo.outerFrameSide == 1)
                    {
                        x4 = x1 + offA;
                        y4 = y1 - memberGeometricInfo.width;
                        x3 = x2 + offB;
                        y3 = y4;
                    }
                    else if (memberGeometricInfo.outerFrameSide == 2)
                    {
                        x4 = x1 - memberGeometricInfo.width;
                        y4 = y1 + offA;
                        x3 = x4;
                        y3 = y2 + offB;
                    }
                    else if (memberGeometricInfo.outerFrameSide == 3)
                    {
                        x4 = x1 + offA;
                        y4 = y1 + memberGeometricInfo.width;
                        x3 = x2 + offB;
                        y3 = y4;
                    }
                }
                else if (memberGeometricInfo.MemberType == 2)
                {
                    y1 = y1 + memberGeometricInfo.offsetA;
                    y2 = y2 + memberGeometricInfo.offsetB;
                }
                else if (memberGeometricInfo.MemberType == 3)
                {
                    x1 = x1 + memberGeometricInfo.offsetA;
                    x2 = x2 + memberGeometricInfo.offsetB;
                }
                else if (memberGeometricInfo.MemberType == 31 || memberGeometricInfo.MemberType == 33)
                {
                    double offA = memberGeometricInfo.offsetA;
                    double offB = memberGeometricInfo.offsetB;
                    x4 = x1 + offA;
                    y4 = y1 + memberGeometricInfo.width;
                    x3 = x2 + offB;
                    y3 = y4;
                }
                x1 = xLocation + (x1 - x0) / pdfResult.ModelWidth * ImageWidth;
                y1 = yLocation - (y1 - y0) / pdfResult.ModelHeight * ImageHeight;
                x2 = xLocation + (x2 - x0) / pdfResult.ModelWidth * ImageWidth;
                y2 = yLocation - (y2 - y0) / pdfResult.ModelHeight * ImageHeight;
                x3 = xLocation + (x3 - x0) / pdfResult.ModelWidth * ImageWidth;
                y3 = yLocation - (y3 - y0) / pdfResult.ModelHeight * ImageHeight;
                x4 = xLocation + (x4 - x0) / pdfResult.ModelWidth * ImageWidth;
                y4 = yLocation - (y4 - y0) / pdfResult.ModelHeight * ImageHeight;
                double w = memberGeometricInfo.width / pdfResult.ModelWidth * ImageWidth;
                if (memberGeometricInfo.MemberType == 1)
                {
                    g.DrawLine(grayPen, x1, y1, x2, y2);
                    g.DrawLine(grayPen, x2, y2, x3, y3);
                    g.DrawLine(grayPen, x3, y3, x4, y4);
                    g.DrawLine(grayPen, x4, y4, x1, y1);
                }
                else if (memberGeometricInfo.MemberType == 2) //mullion
                {
                    //g.DrawLine(grayDashPen, x1, y1, x2, y2); // dash dot line
                    g.DrawLine(grayPen, x1 - w / 2, y1, x2 - w / 2, y2);
                    g.DrawLine(grayPen, x1 + w / 2, y1, x2 + w / 2, y2);
                }
                else if (memberGeometricInfo.MemberType == 3) //transom
                {
                    //g.DrawLine(grayDashPen, x1, y1, x2, y2); // dash dot line
                    g.DrawLine(grayPen, x1, y1 - w / 2, x2, y2 - w / 2);
                    g.DrawLine(grayPen, x1, y1 + w / 2, x2, y2 + w / 2);
                }
                else if (memberGeometricInfo.MemberType == 31 || memberGeometricInfo.MemberType == 33)
                {
                    g.DrawLine(grayPen, x1, y1, x2, y2);
                    g.DrawLine(grayPen, x3, y3, x4, y4);
                }

                // draw member label
                double LabelOffset = pdfResult.OuterFrameWidth * Math.Max(ImageWidth / pdfResult.ModelWidth, ImageHeight / pdfResult.ModelHeight) + 2;
                xs = (x1 + x2) / 2 + LabelOffset + 2;
                ys = (y1 + y2) / 2 - LabelOffset;
                drawString = $"{memberGeometricInfo.MemberID}";
                rectMemberID = new XRect(xs - 2, ys - 7.5, 9, 9);
                if (memberGeometricInfo.MemberType == 2 || memberGeometricInfo.MemberType == 3)
                {
                    XFont memberFont = new XFont("Univers for Schueco 330 Light", 8, XFontStyle.Bold);
                    g.DrawString(drawString, memberFont, blackBrush, xs, ys);
                    g.DrawRectangle(blackPen, rectMemberID);
                }
                else
                {
                    XFont memberFont = new XFont("Univers for Schueco 330 Light", 8, XFontStyle.Regular);
                    g.DrawString(drawString, drawFont, grayBrush, xs, ys);
                    g.DrawRectangle(grayPen, rectMemberID);
                }
            }


            // to be updated for door
            foreach (GlassGeometricInfo glassGeometricInfo in pdfResult.GlassGeometricInfos)
            {
                double x0 = pdfResult.ModelOriginX;
                double y0 = pdfResult.ModelOriginY;
                xs = glassGeometricInfo.PointCoordinates[0];
                ys = glassGeometricInfo.PointCoordinates[1];
                xs = xLocation + (xs - x0) / pdfResult.ModelWidth * ImageWidth;
                ys = yLocation - (ys - y0) / pdfResult.ModelHeight * ImageHeight;
                drawString = $"{glassGeometricInfo.GlassID}";
                g.DrawString(drawString, drawFont, grayBrush, xs, ys);

                rectGlassID = new XRect(xs - 2, ys - 7.5, 9, 9);
                g.DrawArc(grayPen, rectGlassID, 0, 360);

                // draw glass
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
                x2 = xLocation + (x2 - x0) / pdfResult.ModelWidth * ImageWidth;
                y2 = yLocation - (y2 - y0) / pdfResult.ModelHeight * ImageHeight;
                rectGlass = new XRect(x1, y1 - dy, dx, dy);
                g.DrawRectangle(grayPen, rectGlass);

                // draw vent
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
                }

                // draw door
                if (!(glassGeometricInfo.DoorArticleWidths is null))
                {
                    // "Single-Door-Left"; "Single-Door-Right"; "Double-Door-Active-Left"; "Double-Door-Active-Right";
                    double leafWidthX = glassGeometricInfo.DoorArticleWidths[0] / pdfResult.ModelWidth * ImageWidth;
                    double leafWidthY = glassGeometricInfo.DoorArticleWidths[0] / pdfResult.ModelHeight * ImageHeight;
                    double passiveJambWidthX = glassGeometricInfo.DoorArticleWidths[1] / pdfResult.ModelWidth * ImageWidth;
                    double sillWidthY = glassGeometricInfo.DoorArticleWidths[2] / pdfResult.ModelWidth * ImageWidth;
                    double midX = (x1 + x2) / 2;

                    // door leafs
                    g.DrawLine(grayPen, x1 + leafWidthX, y1, x1 + leafWidthX, y2 + leafWidthY);
                    g.DrawLine(grayPen, x1 + leafWidthX, y2 + leafWidthY, x2 - leafWidthX, y2 + leafWidthY);
                    g.DrawLine(grayPen, x2 - leafWidthX, y2 + leafWidthY, x2 - leafWidthX, y1);

                    if (glassGeometricInfo.VentOperableType == "Single-Door-Left" || glassGeometricInfo.VentOperableType == "Single-Door-Right")
                    {
                        // door sill
                        g.DrawLine(grayPen, x1 + leafWidthX, y1 - sillWidthY, x2 - leafWidthX, y1 - sillWidthY);
                    }
                    else if (glassGeometricInfo.VentOperableType == "Double-Door-Active-Left")
                    {
                        // door jamb
                        g.DrawLine(grayPen, midX, y1, midX, y2 + leafWidthY);
                        g.DrawLine(grayPen, midX - leafWidthX, y1, midX - leafWidthX, y2 + leafWidthY);
                        g.DrawLine(grayPen, midX + passiveJambWidthX, y1, midX + passiveJambWidthX, y2 + leafWidthY);

                        // door sill
                        g.DrawLine(grayPen, x1 + leafWidthX, y1 - sillWidthY, midX - leafWidthX, y1 - sillWidthY);
                        g.DrawLine(grayPen, midX + passiveJambWidthX, y1 - sillWidthY, x2 - leafWidthX, y1 - sillWidthY);
                    }
                    else if (glassGeometricInfo.VentOperableType == "Double-Door-Active-Right")
                    {
                        // door jamb
                        g.DrawLine(grayPen, midX, y1, midX, y2 + leafWidthY);
                        g.DrawLine(grayPen, midX - passiveJambWidthX, y1, midX - passiveJambWidthX, y2 + leafWidthY);
                        g.DrawLine(grayPen, midX + leafWidthX, y1, midX + leafWidthX, y2 + leafWidthY);

                        // door sill
                        g.DrawLine(grayPen, x1 + leafWidthX, y1 - sillWidthY, midX - passiveJambWidthX, y1 - sillWidthY);
                        g.DrawLine(grayPen, midX + leafWidthX, y1 - sillWidthY, x2 - leafWidthX, y1 - sillWidthY);
                    }
                }

            }

            // get dimension info
            double[] xdimensions = pdfResult.MemberGeometricInfos.Select(item => item.PointCoordinates[0]).Distinct().ToArray();
            double[] ydimensions = pdfResult.MemberGeometricInfos.Select(item => item.PointCoordinates[1]).Distinct().ToArray();
            Array.Sort(xdimensions);
            Array.Sort(ydimensions);
            double modelLength = xdimensions.Last() - xdimensions.First();
            double modelHeight = ydimensions.Last() - ydimensions.First();
            xdimensions = xdimensions.Select(item => (item - xdimensions.First()) / xdimensions.Last()).ToArray();
            ydimensions = ydimensions.Select(item => (item - ydimensions.First()) / ydimensions.Last()).ToArray();

            // draw legend
            xs = xLocation + 5;
            ys = yLocation + 15 + 11 * xdimensions.Count();
            drawString = $"n   Glass ID";
            if (Thread.CurrentThread.CurrentCulture.Name == "de-DE")
            {
                drawString = $"n   Glas-Position";
            }
            if (Thread.CurrentThread.CurrentCulture.Name == "fr-FR")
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
            if (Thread.CurrentThread.CurrentCulture.Name == "fr-FR")
            {
                drawString = $"n   ID de Membre Statique";
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
                PdfDrawDimension(myTemplate.Pages[pageNumber], 0, xdimensions[i], i - 1, xdimensions[i] * modelLength, xdimLocation, ydimLocation, ImageWidth);
            }

            xdimLocation = xLocation + ImageWidth;
            ydimLocation = yLocation;
            for (int i = 1; i < ydimensions.Count(); i++)
            {
                PdfDrawDimension(myTemplate.Pages[pageNumber], 0, ydimensions[i], i - 1, ydimensions[i] * modelHeight, xdimLocation, ydimLocation, ImageHeight, true);
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

            g.Save();
            g.Dispose();
        }

        public void PdfDrawTriangle(XGraphics g, double xs, double ys, XPen pen)
        {
            XPoint point1 = new XPoint(xs - 3.5, ys + 1.5);
            XPoint point2 = new XPoint(xs + 7.5, ys + 1.5);
            XPoint point3 = new XPoint(xs + 2, ys - 8.5);
            XPoint[] curvePoints = { point1, point2, point3 };
            g.DrawPolygon(pen, curvePoints);
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
            string imageFilePath = $"{_resourceFolderPath}article-jpeg\\{pdfMemberResult.ArticleName}.jpg";
            if (!File.Exists(imageFilePath))
            {
                imageFilePath = $"{_resourceFolderPath}article-jpeg\\CustomArticle.jpg";
            }

            XGraphics g = XGraphics.FromPdfPage(myTemplate.Pages[0]);
            XImage image = XImage.FromFile(imageFilePath);
            double xLocation = 400;
            double yLocation = 145;
            double scaledWidth, scaledHeight;
            double imageAspectRatio = image.PointWidth / image.PointHeight;
            if (imageAspectRatio > 1.2)
            {
                scaledWidth = 120;
                scaledHeight = scaledWidth / imageAspectRatio;
            }
            else
            {
                scaledHeight = 100;
                scaledWidth = imageAspectRatio * scaledHeight;
            }

            g.DrawImage(image, xLocation, yLocation, scaledWidth, scaledHeight);

            g.Save();
            g.Dispose();

        }

        //public void PdfDrawLoadCaseImage(PDFMemberResult pdfMemberResult, PdfDocument myTemplate)
        //{
        //    bool isTransom = pdfMemberResult.MemberType == 3;
        //    PdfDrawBeam(myTemplate, 1);
        //    PdfDrawLoadCaseImageTitle(myTemplate, isTransom, 1);
        //    PdfDrawPinSupport(myTemplate, 1);
        //    PdfDrawLoads(pdfMemberResult, myTemplate, 1);
        //    PdfDrawBeam(myTemplate, 2);
        //    PdfDrawLoadCaseImageTitle(myTemplate, isTransom, 2);
        //    PdfDrawPinSupport(myTemplate, 2);
        //    PdfDrawLoads(pdfMemberResult, myTemplate, 2);

        //    bool existHorizontalLiveLoad = true;
        //    if( pdfMemberResult.ReactionForce[1, 0] == 0 && pdfMemberResult.ReactionForce[1, 1] == 0)
        //    {
        //        existHorizontalLiveLoad = false;
        //    }
        //    if (existHorizontalLiveLoad)
        //    {
        //        PdfDrawBeam(myTemplate, 3);
        //        PdfDrawLoadCaseImageTitle(myTemplate, isTransom, 3); //use chartNo=3 to draw support for reaction force
        //        PdfDrawPinSupport(myTemplate, 3);
        //        PdfDrawLoads(pdfMemberResult, myTemplate, 3);
        //    }
        //    PdfDrawLoadCaseImageTitle(myTemplate, isTransom, 4, existHorizontalLiveLoad); //use chartNo=4 to draw support for reaction force
        //    PdfDrawPinSupport(myTemplate, 4, existHorizontalLiveLoad);
        //    PdfDrawReactions(pdfMemberResult.ReactionForce, myTemplate.Pages[0], existHorizontalLiveLoad);
        //}

        public void PdfDrawLoadCaseImage(PDFMemberResult pdfMemberResult, PdfDocument myTemplate)
        {
            bool isTransom = pdfMemberResult.MemberType == 5;

            // set y location of each drawing
            bool existHorizontalLiveLoad = true;
            if ((Math.Abs(pdfMemberResult.ReactionForce[1, 0]) < 0.0001) && (Math.Abs(pdfMemberResult.ReactionForce[1, 1]) < 0.0001))
            {
                existHorizontalLiveLoad = false;
            }

            bool existDeadLoad = false;
            if (pdfMemberResult.MemberType == 3)
            {
                existDeadLoad = true;
            }
            // set existDeadLoad to false for now, will uncomment after chartNo.5 is done
            //if (pdfMemberResult.MemberType == 3)
            //{
            //    existDeadLoad = true;
            //}

            double y1 = 310;
            double y2 = 420;
            double y3 = 510;
            double y4 = 570;
            double y5 = 680;

            if (existHorizontalLiveLoad && !existDeadLoad)
            {
                y1 = 320;
                y2 = 460;
                y3 = 570;
                y4 = 670;
            }

            if (!existHorizontalLiveLoad && existDeadLoad)
            {
                y1 = 320;
                y2 = 460;
                y4 = 560;
                y5 = 670;
            }

            if (!existHorizontalLiveLoad && !existDeadLoad)
            {
                y1 = 320;
                y2 = 500;
                y4 = 650;
            }

            PdfDrawBeam(myTemplate, 1, y1);
            PdfDrawLoadCaseImageTitle(myTemplate, isTransom, 1, y1);
            PdfDrawPinSupport(myTemplate, 1, y1);
            PdfDrawLoads(pdfMemberResult, myTemplate, 1, y1);
            PdfDrawCoordinate(pdfMemberResult, myTemplate, 1, y1);

            PdfDrawBeam(myTemplate, 2, y2);
            PdfDrawLoadCaseImageTitle(myTemplate, isTransom, 2, y2);
            PdfDrawPinSupport(myTemplate, 2, y2);
            PdfDrawLoads(pdfMemberResult, myTemplate, 2, y2);

            if (existHorizontalLiveLoad)
            {
                PdfDrawBeam(myTemplate, 3, y3);
                PdfDrawLoadCaseImageTitle(myTemplate, isTransom, 3, y3); //use chartNo=3 to draw horizontal load
                PdfDrawPinSupport(myTemplate, 3, y3);
                PdfDrawLoads(pdfMemberResult, myTemplate, 3, y3);
            }

            PdfDrawLoadCaseImageTitle(myTemplate, isTransom, 4, y4); //use chartNo=4 to draw support for reaction force
            PdfDrawPinSupport(myTemplate, 4, y4);
            PdfDrawReactions(pdfMemberResult, myTemplate.Pages[0], 4, y4);

            if (existDeadLoad)
            {
                PdfDrawBeam(myTemplate, 5, y5);
                PdfDrawLoadCaseImageTitle(myTemplate, isTransom, 5, y5); //use chartNo=5 to draw dead load
                PdfDrawPinSupport(myTemplate, 5, y5);
                PdfDrawLoads(pdfMemberResult, myTemplate, 5, y5);
                PdfDrawCoordinate(pdfMemberResult, myTemplate, 5, y5);
                PdfDrawReactions(pdfMemberResult, myTemplate.Pages[0], 5, y5);
            }
        }

        public void PdfDrawLoadCaseImageTitle(PdfDocument myTemplate, bool isTransom, int chartNo, double yref)
        {
            double x = 245;
            double y = chartNo == 5 ? yref - 40 : yref + 13;

            XGraphics g = XGraphics.FromPdfPage(myTemplate.Pages[0]);
            XPen pen = new XPen(XColor.FromArgb(38, 38, 38), .25);
            XBrush whitebrush = new XSolidBrush(XColor.FromArgb(255, 255, 255));
            XBrush graybrush = new XSolidBrush(XColor.FromArgb(102, 102, 102));
            XBrush brush = chartNo == 4 ? graybrush : whitebrush;

            XFont drawFont = new XFont("Univers for Schueco 330 Light", 8, XFontStyle.Regular);
            string drawString = "";
            if (chartNo == 1)
            {
                drawString = isTransom ? "Wind Load on Transom Top Side" : "Wind Load on Mullion Left Side";
                if (Thread.CurrentThread.CurrentCulture.Name == "de-DE")
                {
                    drawString = isTransom ? "Riegel: Belastung aus oberem Feld" : "Pfosten: Belastung aus linkem Feld";
                }

                if (Thread.CurrentThread.CurrentCulture.Name == "fr-FR")
                {
                    drawString = isTransom ? "Charge de vent sur la demi trame au dessus de la traverse" : "Charge sur le côté gauche du poteau";
                }
            }
            else if (chartNo == 2)
            {
                drawString = isTransom ? "Wind Load on Transom Bottom Side" : "Wind Load on Mullion Right Side";
                if (Thread.CurrentThread.CurrentCulture.Name == "de-DE")
                {
                    drawString = isTransom ? "Riegel: Belastung aus unterem Feld" : "Pfosten: Belastung aus rechtem Feld";
                }

                if (Thread.CurrentThread.CurrentCulture.Name == "fr-FR")
                {
                    drawString = isTransom ? "Charge de vent sur la trame en dessous de la traverse" : "Charge sur le côté droit du poteau";
                }
            }
            else if (chartNo == 3)
            {
                drawString = "Horizontal Live Load";
                if (Thread.CurrentThread.CurrentCulture.Name == "de-DE")
                {
                    drawString = "Horizontale Nutzlast";
                }

                if (Thread.CurrentThread.CurrentCulture.Name == "fr-FR")
                {
                    drawString = "Charge utile horizontale";
                }
            }
            else if (chartNo == 4)
            {
                drawString = "Reaction Force (horizontal)";
                if (Thread.CurrentThread.CurrentCulture.Name == "de-DE")
                {
                    drawString = "Auflagerkräfte";
                }

                if (Thread.CurrentThread.CurrentCulture.Name == "fr-FR")
                {
                    drawString = "Réactions aux appuis";
                }
            }
            else if (chartNo == 5)
            {
                drawString = "Dead Load (vertical)";
                if (Thread.CurrentThread.CurrentCulture.Name == "de-DE")
                {
                    drawString = "Eigengewicht";
                }
                else if (Thread.CurrentThread.CurrentCulture.Name == "fr-FR")
                {
                    drawString = "Poids propre";
                }
            }
            XSize size = g.MeasureString(drawString, drawFont);

            g.DrawString(drawString, drawFont, brush, x, y);

            g.Save();
            g.Dispose();
        }


        public void PdfDrawBeam(PdfDocument myTemplate, int chartNo, double y)
        {
            XGraphics g = XGraphics.FromPdfPage(myTemplate.Pages[0]);

            XPen grayPen = new XPen(XColor.FromArgb(38, 38, 38), .25);
            XBrush blueBrush = new XSolidBrush(XColor.FromArgb(0, 162, 209));
            XBrush grayBrush = new XSolidBrush(XColor.FromArgb(102, 102, 102));

            double x, width, height;

            x = 100;
            width = 400;

            if (chartNo == 5)
            {
                height = 5;
                g.DrawRectangle(grayPen, grayBrush, x, y, width, height);
            }
            else
            {
                // Draw top beam
                height = 5;
                g.DrawRectangle(grayPen, blueBrush, x, y, width, height);
                g.DrawRectangle(grayPen, x, y, width, height - 4);

                // Draw the Thermal Break
                y = y + 5;
                height = 10;
                g.DrawRectangle(grayPen, grayBrush, x, y, width, height);

                // Draw Bottom beam
                y = y + 10;
                height = 2.5;
                g.DrawRectangle(grayPen, blueBrush, x, y, width, height);
                g.DrawRectangle(grayPen, blueBrush, x, y, width, height - 1);
            }

            g.Save();
            g.Dispose();
        }

        public void PdfDrawPinSupport(PdfDocument myTemplate, int chartNo, double yref)
        {
            XGraphics g = XGraphics.FromPdfPage(myTemplate.Pages[0]);

            XPen grayPen = new XPen(XColor.FromArgb(38, 38, 38), .25);
            XBrush blueBrush = new XSolidBrush(XColor.FromArgb(0, 162, 209));
            XBrush grayBrush = new XSolidBrush(XColor.FromArgb(102, 102, 102));

            // draw support
            double x = 120;
            double y = chartNo == 5 ? yref + 5 : yref + 18;

            x = 100;
            XPoint point1 = new XPoint(x, y);
            XPoint point2 = new XPoint(x - 5, y + 10);
            XPoint point3 = new XPoint(x + 5, y + 10);
            XPoint[] curvePoints = { point1, point2, point3 };
            g.DrawPolygon(grayPen, grayBrush, curvePoints, XFillMode.Alternate);
            g.DrawRectangle(grayPen, blueBrush, x - 7, y + 10, 14, 2);
            x = 500;
            point1 = new XPoint(x, y);
            point2 = new XPoint(x - 5, y + 10);
            point3 = new XPoint(x + 5, y + 10);
            XPoint[] curvePoints2 = { point1, point2, point3 };
            g.DrawPolygon(grayPen, grayBrush, curvePoints2, XFillMode.Alternate);
            g.DrawRectangle(grayPen, blueBrush, x - 7, y + 10, 14, 2);

            g.Save();
            g.Dispose();
        }

        public void PdfDrawLoads(PDFMemberResult pdfMemberResult, PdfDocument myTemplate, int chartNo, double yref)
        {
            double xStart, xEnd, valueStart, valueEnd, valueMax;
            bool isHLiveLoad;
            double len = pdfMemberResult.Length;
            double[,] loadCaseData = pdfMemberResult.WindLoadData;

            int rowcount = loadCaseData.GetLength(0);
            double[] dimData = new double[42];

            if (chartNo == 1 || chartNo == 2)        // merge concentrated load for horizontal live load, since it has only one chart, left and right load may overlap each other
            {
                loadCaseData = pdfMemberResult.WindLoadData;
                MergeLoadCaseData(loadCaseData);
            }
            if (chartNo == 3)        // merge concentrated load for horizontal live load, since it has only one chart, left and right load may overlap each other
            {
                loadCaseData = pdfMemberResult.HorizontalLiveLoadData;
                MergeLoadCaseData(loadCaseData);
            }

            if (chartNo == 5)        // merge concentrated load for horizontal live load, since it has only one chart, left and right load may overlap each other
            {
                loadCaseData = pdfMemberResult.VerticalLoadData;
                MergeLoadCaseData(loadCaseData);
            }

            valueMax = 0;
            for (int i = 0; i < rowcount; i++)
            {
                if (Math.Abs(loadCaseData[i, 2]) > valueMax)
                {
                    valueMax = Math.Abs(loadCaseData[i, 2]);
                }

                if (Math.Abs(loadCaseData[i, 3]) > valueMax)
                {
                    valueMax = Math.Abs(loadCaseData[i, 3]);
                }
            }

            // allow some room for text
            valueMax = valueMax * 1.25;

            // Draw forces
            int dimCounter = 0;
            for (int i = 0; i < rowcount; i++)
            {
                if ((chartNo == 1 || chartNo == 2) && Convert.ToInt32(loadCaseData[i, 7]) != chartNo) continue;   // check loadSide for wind load
                xStart = loadCaseData[i, 0] / len;
                xEnd = loadCaseData[i, 1] / len;
                valueStart = loadCaseData[i, 2];
                valueEnd = loadCaseData[i, 3];

                isHLiveLoad = chartNo == 3;

                if (valueStart != 0 || valueEnd != 0)
                {
                    PdfDrawForce(myTemplate, xStart, xEnd, valueStart, valueEnd, valueMax, chartNo, yref, isHLiveLoad);
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
            for (int i = 1; i < dimData.Length; i++)
            {
                if (dimData[i] != 0 && dimData[i] != dimData[i - 1])
                {
                    PdfDrawDimension(myTemplate.Pages[0], dimData[i - 1], dimData[i], 0, (dimData[i] - dimData[i - 1]) * len * 10, xLocation, yLocation, Length);
                    dimNumber = dimNumber + 1;
                }
            }
        }

        internal void MergeLoadCaseData(double[,] loadCaseData)  // merge concentrated loads at same location
        {
            int rowcount = loadCaseData.GetLength(0);
            for (int i = 0; i < rowcount - 1; i++)
            {
                if (Convert.ToInt32(loadCaseData[i, 8]) != 2) continue;
                for (int j = i + 1; j < rowcount; j++)
                {
                    if (loadCaseData[i, 0] == loadCaseData[j, 0] && loadCaseData[i, 1] == loadCaseData[j, 1])
                    {
                        loadCaseData[i, 2] = loadCaseData[i, 2] + loadCaseData[j, 2];
                        loadCaseData[i, 3] = loadCaseData[i, 3] + loadCaseData[j, 3];
                        loadCaseData[j, 2] = 0;
                        loadCaseData[j, 3] = 0;
                        break;

                    }
                }
            }
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

            x = xLocation + xStart * beamLength;
            y = yLocation + nc * 11;
            width = (xEnd - xStart) * beamLength;

            g.DrawLine(pen, x, y, x + width, y);
            g.DrawLine(pen, x, y + 2, x, y - 11);
            g.DrawLine(pen, x + width, y + 2, x + width, y - (nc + 1) * 11);

            XPoint point1 = new XPoint(x, y);
            XPoint point2 = new XPoint(x + 5, y - 1.5);
            XPoint point3 = new XPoint(x + 5, y + 1.5);
            XPoint[] curvePoints = { point1, point2, point3 };
            g.DrawPolygon(pen, brush, curvePoints, XFillMode.Alternate);

            curvePoints[0] = new XPoint(x + width, y);
            curvePoints[1] = new XPoint(x + width - 5, y - 1.5);
            curvePoints[2] = new XPoint(x + width - 5, y + 1.5);

            g.DrawPolygon(pen, brush, curvePoints, XFillMode.Alternate);

            // draw the dimension text
            XFont drawFont = new XFont("Univers for Schueco 330 Light", 6, XFontStyle.Regular);
            String drawString = $"{value,0:N0}mm";

            XSize size = g.MeasureString(drawString, drawFont);

            double xt = x + width / 2 - size.Width / 2;

            g.DrawString(drawString, drawFont, brush, xt, y - .75);

            g.Save();
            g.Dispose();

        }

        public void PdfDrawForce(PdfDocument myTemplate, double xs, double xe, double qs, double qe, double qmax, int chartNo, double yref, bool isHLiveLoad = false, bool isConcentrated = false)
        {
            XGraphics g = XGraphics.FromPdfPage(myTemplate.Pages[0]);

            XPen grayPen = new XPen(XColor.FromArgb(38, 38, 38), 0.25);
            XPen bluePen = new XPen(XColor.FromArgb(0, 162, 209), 0.25);
            XBrush grayBrush = new XSolidBrush(XColor.FromArgb(102, 102, 102));
            XBrush blueBrush = new XSolidBrush(XColor.FromArgb(0, 162, 209));

            double xLocation = 100;
            double yLocation = yref;
            double beamLength = 400;
            double arrowSpaceing = isHLiveLoad ? 40 : 20;
            int numberArrows = Convert.ToInt32(Math.Abs(xe - xs) * beamLength / arrowSpaceing);
            double maxTailLength = 20;
            switch (chartNo)
            {
                case 1:
                    maxTailLength = isConcentrated ? 40 : 30;
                    break;
                case 2:
                    maxTailLength = isConcentrated ? 40 : 30;
                    break;
                case 3:
                    maxTailLength = 20;
                    break;
                case 5:
                    maxTailLength = 40;
                    break;
            }

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
            for (int i = 0; i <= numberArrows; i++)
            {
                x = xLocation + (xs * beamLength) + i * arrowSpaceing;
                force = qs + slope * (i * deltaX);
                y = yLocation - maxTailLength * Math.Abs(force) / qmax;


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


            // draw the load line
            double x1, x2, y1, y2;
            x1 = xLocation + (xs * beamLength);
            x2 = xLocation + (xe * beamLength);
            y1 = yLocation - maxTailLength * qs / qmax;
            y2 = yLocation - maxTailLength * qe / qmax;
            if (xs != xe)
            {
                g.DrawLine(arrowPen, x1, y1, x2, y2);
            }

            // draw the load value
            XFont drawFont = new XFont("Univers for Schueco 330 Light", 8, XFontStyle.Regular);
            if (xs == xe)

            {
                x = xLocation + (xs * beamLength);
                y = yLocation - maxTailLength * Math.Abs(qs) / qmax - 2.5;
                String drawString = qs.ToString("F2") + "kN";
                g.DrawString(drawString, drawFont, arrowBrush, x, y);
            }
            else
            {
                String drawString = qs.ToString("F2") + "kN/m";
                if (qs != 0)
                {
                    g.DrawString(drawString, drawFont, arrowBrush, x1, y1 - 2.5);
                }
                drawString = qe.ToString("F2") + "kN/m";

                if (qe != 0)
                {
                    g.DrawString(drawString, drawFont, arrowBrush, x2, y2 - 2.5);
                }
            }
            g.Save();
            g.Dispose();
        }

        public void PdfDrawReactions(PDFMemberResult pdfMemberResult, PdfPage myPage, int chartNo, double yref)
        {
            double[,] reactionForce = pdfMemberResult.ReactionForce;
            // draw reactions
            double xLocation = 100;
            double yLocation = chartNo == 5 ? yref + 18 : yref + 31;
            double beamLength = 400;
            double x, y;

            // draw the left vertical reaction
            if (Math.Abs(reactionForce[0, 0]) > 0.0001)
            {
                x = xLocation;
                y = yLocation;
                PdfDrawVerticalReaction(myPage, x, y, reactionForce[0, 0], reactionForce[1, 0], "left");
            }

            // draw the right vertical reaction
            if (Math.Abs(reactionForce[0, 1]) > 0.0001)
            {
                x = xLocation + beamLength;
                y = yLocation;
                PdfDrawVerticalReaction(myPage, x, y, reactionForce[0, 1], reactionForce[1, 1], "right");
            }
        }

        public void PdfDrawVerticalReaction(PdfPage myPage, double x, double y, double wlRF, double hllRF, string side)
        {
            XGraphics g = XGraphics.FromPdfPage(myPage);
            XSize textSize;

            XFont drawFont = new XFont("Univers for Schueco 330 Light", 8, XFontStyle.Regular);
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


            String drawString = $"{wlRF,3:0.0#}kN";
            textSize = g.MeasureString(drawString, drawFont);
            if (side == "left")
            {
                g.DrawString(drawString, drawFont, grayBrush, x - textSize.Width - 3, y + 8.5);
            }
            else
            {
                g.DrawString(drawString, drawFont, grayBrush, x + 3, y + 8.5);
            }
            drawString = $"{hllRF,3:0.0#}kN";
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
            g.Save();
            g.Dispose();
        }

        public void PdfDrawMomentReaction(PdfDocument myTemplate, double x, double y, double value, string side)
        {
            XGraphics g = XGraphics.FromPdfPage(myTemplate.Pages[0]);
            XSize textSize;

            XFont drawFont = new XFont("Univers for Schueco 330 Light", 8, XFontStyle.Regular);
            XPen grayPen = new XPen(XColor.FromArgb(38, 38, 38), .25);
            XPen grayPenThick = new XPen(XColor.FromArgb(102, 102, 102), 1.5);
            XBrush greenBrush = new XSolidBrush(XColor.FromArgb(120, 185, 40));
            XBrush grayBrush = new XSolidBrush(XColor.FromArgb(102, 102, 102));

            double a1, a2;
            double r = 20;
            if (side == "left")
            {
                a1 = 230 * Math.PI / 180;
                a2 = 220 * Math.PI / 180;

                g.DrawArc(grayPenThick, x, y, r, r, 140, 80);
            }
            else
            {
                a1 = -50 * Math.PI / 180;
                a2 = -40 * Math.PI / 180;
                g.DrawArc(grayPenThick, x, y, r, r, 40, -80);
            }

            r = r / 2;
            x = x + r;
            y = y + r;

            XPoint point1 = new XPoint(x + r * Math.Cos(a1), y + r * Math.Sin(a1));
            XPoint point2 = new XPoint((x + (r + .9) * Math.Cos(a2)), (y + (r + .9) * Math.Sin(a2)));
            XPoint point3 = new XPoint((x + (r - 1.1) * Math.Cos(a2)), (y + (r - 1.1) * Math.Sin(a2)));
            XPoint[] curvePoints = { point1, point2, point3 };

            g.DrawPolygon(grayPenThick, grayBrush, curvePoints, XFillMode.Alternate);

            String drawString = Math.Abs(value).ToString("F2");
            textSize = g.MeasureString(drawString, drawFont);

            if (side == "left")
            {
                g.DrawString(drawString, drawFont, grayBrush, x - textSize.Width - 3, y + 15);
            }
            else
            {
                g.DrawString(drawString, drawFont, grayBrush, x + 3, y + 15);
            }

            g.Save();
            g.Dispose();

        }

        public void PdfDrawResultCurves(PDFMemberResult pdfMemberResult, PdfDocument myTemplate)
        {
            // wip
            XGraphics g = XGraphics.FromPdfPage(myTemplate.Pages[1]);
            XPen Pen1 = new XPen(XColor.FromArgb(0, 162, 209), .75);
            XPen Pen2 = new XPen(XColor.FromArgb(241, 135, 0), .75);
            XPen Pen3 = new XPen(XColor.FromArgb(132, 206, 228), .75);
            XPen Pen4 = new XPen(XColor.FromArgb(0, 86, 129), .75);

            XBrush blackBrush = new XSolidBrush(XColor.FromArgb(0, 0, 0));
            XFont drawFont = new XFont("Univers for Schueco 330 Light", 8);

            double x0s = 69;
            double x0w = 323;
            double y0r1 = 163.2;
            double Yoffset = 160;
            double graphXsize = 216;
            double graphYsize = 54;
            string[] seasons = new string[2];
            seasons[0] = "summer";
            seasons[1] = "winter";
            double x0, y0, x1, y1, x2, y2;
            double yOutofPlaneDeflectionLimit = CalculateLimit(pdfMemberResult.MaxOutofplaneDeflection * 10);
            double yInPlaneDeflectionLimit = CalculateLimit(pdfMemberResult.MaxInplaneDeflection * 10);
            double[] MmaxArray = new double[] {pdfMemberResult.MomentMatrix[0, 0], pdfMemberResult.MomentMatrix[1, 0], pdfMemberResult.MomentMatrix[0, 1], pdfMemberResult.MomentMatrix[1, 1],
                pdfMemberResult.MomentMatrix[0, 2], pdfMemberResult.MomentMatrix[1, 2] };
            double yMLimit = CalculateLimit(MmaxArray.Max() * 100);
            double[] StressMaxArray = new double[] {pdfMemberResult.StressMatrix[0, 0], pdfMemberResult.StressMatrix[3, 0], pdfMemberResult.StressMatrix[0, 1], pdfMemberResult.StressMatrix[3, 1],
                pdfMemberResult.StressMatrix[0, 2], pdfMemberResult.StressMatrix[3, 2],pdfMemberResult.StressMatrix[0, 3], pdfMemberResult.StressMatrix[3, 3]};
            double yStressLimit = CalculateLimit(StressMaxArray.Max());
            double[] ShearMaxArray = new double[] { pdfMemberResult.StressMatrix[0, 4], pdfMemberResult.StressMatrix[3, 4] };
            double yShearLimit = CalculateLimit(ShearMaxArray.Max());
            XPen pen = Pen1;

            foreach (string season in seasons)
            {
                x0 = season == "winter" ? x0w : x0s;
                double[,] resultCurve = season == "winter" ? pdfMemberResult.winterResultCurves : pdfMemberResult.summerResultCurves;

                // for moment
                y0 = y0r1;
                PdfGridLines(g, x0, y0, graphXsize, graphYsize, yMLimit);
                for (int i = 1; i <= 3; i++)
                {
                    switch (i)
                    {
                        case 1:
                            pen = Pen1;
                            break;
                        case 2:
                            pen = Pen2;
                            break;
                        case 3:
                            pen = Pen3;
                            break;
                    }
                    for (int j = 1; j <= 100; j++)
                    {
                        x1 = x0 + (j - 1) / 100.0 * graphXsize;
                        x2 = x0 + j / 100.0 * graphXsize;
                        y1 = y0 - resultCurve[i, j - 1] * 100 / yMLimit * graphYsize;
                        y2 = y0 - resultCurve[i, j] * 100 / yMLimit * graphYsize;
                        g.DrawLine(pen, x1, y1, x2, y2);
                    }
                }


                //for sigma
                y0 = y0r1 + Yoffset;
                PdfGridLines(g, x0, y0, graphXsize, graphYsize, yStressLimit);
                for (int i = 4; i <= 7; i++)
                {
                    switch (i - 3)
                    {
                        case 1:
                            pen = Pen1;
                            break;
                        case 2:
                            pen = Pen2;
                            break;
                        case 3:
                            pen = Pen3;
                            break;
                        case 4:
                            pen = Pen4;
                            break;
                    }
                    for (int j = 1; j <= 100; j++)
                    {
                        x1 = x0 + (j - 1) / 100.0 * graphXsize;
                        x2 = x0 + j / 100.0 * graphXsize;
                        y1 = y0 - resultCurve[i, j - 1] / yStressLimit * graphYsize;
                        y2 = y0 - resultCurve[i, j] / yStressLimit * graphYsize;
                        g.DrawLine(pen, x1, y1, x2, y2);
                    }
                }


                //for Tv
                y0 = y0r1 + 2 * Yoffset;
                PdfGridLines(g, x0, y0, graphXsize, graphYsize, yShearLimit);
                for (int j = 1; j <= 100; j++)
                {
                    x1 = x0 + (j - 1) / 100.0 * graphXsize;
                    x2 = x0 + j / 100.0 * graphXsize;
                    y1 = y0 - resultCurve[8, j - 1] / yShearLimit * graphYsize;
                    y2 = y0 - resultCurve[8, j] / yShearLimit * graphYsize;
                    g.DrawLine(Pen1, x1, y1, x2, y2);
                }
            }

            //for out-of-plane displacement
            x0 = x0s;
            y0 = y0r1 + 3 * Yoffset;
            PdfGridLines(g, x0, y0, graphXsize, graphYsize, yOutofPlaneDeflectionLimit);
            for (int j = 1; j <= 100; j++)
            {
                x1 = x0 + (j - 1) / 100.0 * graphXsize;
                x2 = x0 + j / 100.0 * graphXsize;
                y1 = y0 - pdfMemberResult.ambientResultCurves[0, j - 1] * 10 / yOutofPlaneDeflectionLimit * graphYsize;  // the disp curve in summer and winter are the same. Both are based on ambient condition;
                y2 = y0 - pdfMemberResult.ambientResultCurves[0, j] * 10 / yOutofPlaneDeflectionLimit * graphYsize;     // convert from cm to mm
                g.DrawLine(Pen1, x1, y1, x2, y2);
            }

            //for In-plane displacement
            if (!(pdfMemberResult.verticalResultCurveX is null))
            {
                x0 = x0w;
                y0 = y0r1 + 3 * Yoffset;
                PdfGridLines(g, x0, y0, graphXsize, graphYsize, yInPlaneDeflectionLimit);
                int iCount = pdfMemberResult.verticalResultCurveX.Count();
                double verticalCurveXmax = pdfMemberResult.verticalResultCurveX[iCount - 1];
                for (int j = 1; j <= iCount - 1; j++)
                {
                    x1 = x0 + pdfMemberResult.verticalResultCurveX[j - 1] / verticalCurveXmax * graphXsize;
                    x2 = x0 + pdfMemberResult.verticalResultCurveX[j] / verticalCurveXmax * graphXsize;
                    y1 = y0 - pdfMemberResult.verticalResultCurveY[j - 1] * 10 / yInPlaneDeflectionLimit * graphYsize;
                    y2 = y0 - pdfMemberResult.verticalResultCurveY[j] * 10 / yInPlaneDeflectionLimit * graphYsize;
                    g.DrawLine(Pen1, x1, y1, x2, y2);
                }
            }
        }

        private double CalculateLimit(double yMax)
        {
            double yLimit = 0;
            if (yMax < 30)
            {
                yLimit = Math.Max(Math.Ceiling(yMax / 4) * 4, 1);
            }
            else
            {
                yLimit = Math.Max(Math.Ceiling(yMax / 20) * 20, 1);
            }

            return yLimit;
        }


        private void PdfGridLines(XGraphics g, double x0, double y0, double graphXsize, double graphYsize, double yLimit)
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
            g.DrawString($"{yLimit,3:0.#}", drawFont, blackFontBrush, x0 - 18, y0 - graphYsize + 3);
            g.DrawString($"{-yLimit,3:0.#}", drawFont, blackFontBrush, x0 - 18, y0 + graphYsize + 3);
        }

        // --------- Utility Pdf Functions  (created by Arash)  ------------------------------------

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
                        txtField.Elements.SetString(PdfTextField.Keys.DA, "/UniversForSchueco-630Bold 7 Tf 1 0 0 rg");
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
