namespace PCReports.Models.ADS
{
    public class PDFResult
    {
        public PDFResult()
        {
            GlassTypes = new List<GlassType>();
            PDFMemberResults = new List<PDFMemberResult>();
        }
        // Section 0. Structure Geometry Info
        public List<MemberGeometricInfo> MemberGeometricInfos { get; set; }
        public List<GlassGeometricInfo> GlassGeometricInfos { get; set; }
        public double ModelWidth;
        public double ModelHeight;
        public double ModelOriginX;
        public double ModelOriginY;
        public double OuterFrameWidth;

        // Section 0. Project Info
        public string ProjectGuid { get; set; }
        public string ProblemGuid { get; set; }
        public string Language { get; set; }
        public string ProjectName { get; set; }
        public string Location { get; set; }
        public string ConfigurationName { get; set; }
        public string UserName { get; set; }
        public string UserNotes { get; set; }

        // Section 1. Window Information
        public string ProfileSystem { get; set; }
        public string FrameProfile { get; set; }
        public string MullionProfile { get; set; }
        public string TransomProfile { get; set; }
        public double FrameProfileWeight { get; set; }
        public double MullionProfileWeight { get; set; }
        public double TransomProfileWeight { get; set; }
        public List<GlassType> GlassTypes { get; set; }
        public int BlockDistance { get; set; }

        // Section 2. Applied Load
        public double WindLoad { get; set; }
        public string CpeString { get; set; }
        public string pCpiString { get; set; }
        public string nCpiString { get; set; }
        public double HorizontalLiveLoad { get; set; }
        public double HorizontalLiveLoadHeight { get; set; }
        public double SummerTempDiff { get; set; }
        public double WinterTempDiff { get; set; }
        public double WindLoadFactor { get; set; }
        public double HorizontalLiveLoadFactor { get; set; }
        public double TemperatureLoadFactor { get; set; }
        public double DeadLoadFactor { get; set; }

        // Section 4.
        public string AllowableDeflectionLine1 { get; set; }
        public string AllowableDeflectionLine2 { get; set; }
        public string AllowableInplaneDeflectionLine { get; set; }

        // Section 5.
        public string Alloys { get; set; }
        public double Beta { get; set; }
        public string InsulatingBarType { get; set; }
        public string InsulatingBarDataNote { get; set; }
        public double RSn20 { get; set; }
        public double RSp80 { get; set; }
        public double RTn20 { get; set; }
        public double RTp80 { get; set; }
        public double Cn20 { get; set; }
        public double Cp20 { get; set; }
        public double Cp80 { get; set; }

        // Section 6. Results
        public List<PDFMemberResult> PDFMemberResults { get; set; }
    }

    public class PDFMemberResult
    {
        // 6.1 member
        public int MemberID { get; set; }
        public int MemberType { get; set; }
        public string ArticleName { get; set; }
        public double Depth { get; set; }
        public double Width { get; set; }
        public double Length { get; set; }
        public double Weight { get; set; }
        public double TributaryArea { get; set; }
        public double Cp { get; set; }
        public double TributaryAreaFactor { get; set; }
        public double AppliedWindLoad { get; set; }

        // profile parameters
        public double Iy { get; set; }
        public double Il { get; set; }
        public double Is { get; set; }
        public double Iv { get; set; }
        public double v { get; set; }
        public double lamdan20 { get; set; }
        public double lamdap20 { get; set; }
        public double lamdap80 { get; set; }

        // beam data
        public double[,] WindLoadData { get; set; }           // xs, xe, qs, qe, Phi_summer, Phi_winter, Phi_disp, loadSide, loadType
        public double[,] HorizontalLiveLoadData { get; set; }  // load type - 1- windload, 2- horizontal live load, 3- reaction force, 4 - vertical force
        public double[,] VerticalLoadData { get; set; }       // vertical load data follow wind load data format
        public double[,] ReactionForce { get; set; }
        public double[] VerticalReactionForce { get; set; }
        public double Ak { get; set; }
        public double Ad { get; set; }
        public double Bk { get; set; }
        public double Bd { get; set; }

        // peak moment
        public double[,] MomentMatrix { get; set; }

        // Maximum deflection
        public double MaxInplaneDeflection { get; set; }
        public double AllowableInplaneDeflecton { get; set; }
        public double InplaneDispIndex { get; set; }
        public string InplaneDeflectionRatio { get; set; }
        public double MaxOutofplaneDeflection { get; set; }
        public double AllowableOutofplaneDeflecton { get; set; }
        public string OutofplaneDeflectionRatio { get; set; }

        // Peak stress
        public double[,] StressMatrix { get; set; }

        public double stressMax { get; set; }
        public double winterShearMax { get; set; }
        public double summerShearMax { get; set; }
        public double stressRatio { get; set; }
        public double summerShearRatio { get; set; }
        public double winterShearRatio { get; set; }
        public double summerCompositeRatio { get; set; }
        public double winterCompositeRatio { get; set; }
        public double summerTransverseRatio { get; set; }
        public double winterTransverseRatio { get; set; }
        public string strStressRatio { get; set; }
        public string strSummerShearRatio { get; set; }
        public string strWinterShearRatio { get; set; }
        public string strSummerCompositeRatio { get; set; }
        public string strWinterCompositeRatio { get; set; }
        public string strSummerTransverseRatio { get; set; }
        public string strWinterTransverseRatio { get; set; }

        // curves
        public double[,] summerResultCurves { get; set; } // 0 - out of plane displacements, 1 - Mo, 2 - Mv, 3- Mu, 4 - Sigmaoo, 5 - Sigmaou, 6 - Sigmauo, 7 - Sigmauu, 8 - Tv
        public double[,] winterResultCurves { get; set; } // 0 - out of plane displacements, 1 - Mo, 2 - Mv, 3- Mu, 4 - Sigmaoo, 5 - Sigmaou, 6 - Sigmauo, 7 - Sigmauu, 8 - Tv
        public double[,] ambientResultCurves { get; set; } // 0 - out of plane displacements, 1 - Mo, 2 - Mv, 3- Mu, 4 - Sigmaoo, 5 - Sigmaou, 6 - Sigmauo, 7 - Sigmauu, 8 - Tv
        public double[] verticalResultCurveX { get; set; }
        public double[] verticalResultCurveY { get; set; }

    }

    public class GlassType
    {
        public string GlassIDs { get; set; }
        public string Weight { get; set; }
        public string Description { get; set; }
    }

    public class MemberGeometricInfo
    {
        public int MemberID { get; set; }
        public int MemberType { get; set; }
        public double[] PointCoordinates { get; set; }
        public double offsetA { get; set; }
        public double offsetB { get; set; }
        public double width { get; set; }
        public int outerFrameSide { get; set; }
    }

    public class GlassGeometricInfo
    {
        public int GlassID { get; set; }
        public double[] PointCoordinates { get; set; }
        public double[] CornerCoordinates { get; set; } // InsertUnit Outer BD
        public double[] VentCoordinates { get; set; } // InsertUnit Inner BD
        public double[] DoorArticleWidths { get; set; } //[DoorleafOutsideW, DoorPassiveJambOutsideW, DoorSillOUtsideW]
        public double[] InsertOuterFrameCoordinates { get; set; }
        public string VentOpeningDirection { get; set; }
        public string VentOperableType { get; set; }
    }
}
