namespace PCReports.Models.Window
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

    public class GlassType
    {
        public string GlassIDs { get; set; }
        public string Weight { get; set; }
        public string Description { get; set; }
    }

    public class FacadeSection : ICloneable
    {
        public object Clone()
        {
            return this.MemberwiseClone();
        }

        public int SectionID;           // = 1,2,3,4 SectionType = SectionID + 3
        public int SectionType;         // same as MemberType. 4: major mullion, 5: transom, 6: minor mullion, 7: reinforcement
                                        // 21:UDC Top Frame; 22: UDC Vertical;  23: UDC Bottom Frame; 24: UDC Vertical Glazing Bar; 25: UDC Horizontal Glazing Bar;
        public string ArticleName;      // For Schuco standard article, use article ID.
        public bool isCustomProfile;    // is custom profile
        public double OutsideW;         // for display in pdf only
        public double BTDepth;
        public double Width;
        public double Zo;          // depth from central axis to the top
        public double Zu;          // depth from central axis to the bottom
        public double Zl;          // depth from central axis to the mullion left
        public double Zr;          // depth from central axis to the mullion right
        public double A;           //Area
        public string Material;        // Aluminum or Steel
        public double beta;           // MPa; depending on Alloy
        public double Weight;      //section weight kg/m
        public double Iyy;         //Moment of inertia for bending about the y-axis, out-of-plane
        public double Izz;           //Moment of inertia for bending about the z-axis, in-plane
        public double Wyy;          // min(Iyy / Zo, Izz/ Zu)
        public double Wzz;          // min(Izz / Zyl, Izz/ Zyr)
        public double Asy;         //Shear area in y-direction
        public double Asz;         //Shear area in z-direction
        public double J;           //Torsional constant 
        public double E;            //Young's modulus
        public double G;            //Torsional shear modulus
        public double EA;                   // Derived EA
        public double GAsy;                 // Derived GA
        public double GAsz;                 // Derived GA
        public double EIy;                  // Derived EIy
        public double EIz;                  // Derived EIz
        public double GJ;                   // Derived GJ
        // sectional properties for display only
        public double Ys = 0;           // Shear center y coordinate
        public double Zs = 0;           // Shear center z coordinate
        public double Ry = 0;           // Radius of Gyration about the y-axis
        public double Rz = 0;           // Radius of Gyration about the z-axis
        public double Wyp = 0;          // Section Modulus about the positive semi-axis of the y-axis 
        public double Wyn = 0;          // Section Modulus about the negative semi-axis of the y-axis 
        public double Wzp = 0;          // Section Modulus about the positive semi-axis of the z-axis 
        public double Wzn = 0;          // Section Modulus about the negative semi-axis of the z-axis 
        public double Cw = 0;
        public double Beta_torsion = 0;
        public double Zy = 0;           // Plastic Property 
        public double Zz = 0;           // Plastic Property 

        // Reinforcement properties, for physics core internal only
        public double ReinforcementEIy;
        public double ReinforcementEIz;
        public double ReinforcementWeight;
    }

    public class Section
    {
        public int SectionID;           // = 1,2,3 , SectionID and SectionType is same
        public int SectionType;        // same as MemberType. 1: Outer Frame, 2: Mullion, 3: transom
        public string ArticleName;     // article name
        public bool isCustomProfile;
        public double InsideW;
        public double OutsideW;
        public double LeftRebate;
        public double RightRebate;
        public double DistBetweenIsoBars;
        public double d;                    // mm.  article parameters for strucutural analysis. from d to alpha. 
        public double Weight;               // N/m
        public double Ao;                   // mm2
        public double Au;                   // mm2
        public double Io;                   // mm4
        public double Iu;                   // mm4
        public double Ioyy;                 // mm4
        public double Iuyy;                 // mm4

        public double Zoo;                  // mm
        public double Zuo;                  // mm
        public double Zou;                  // mm
        public double Zuu;                  // mm
        public double Zol;                  // mm
        public double Zul;                  // mm
        public double Zor;                  // mm
        public double Zur;                  // mm

        public double RSn20;                // N/m  
        public double RSp80;
        public double RTn20;
        public double RTp80;
        public double Cn20;                 // N/mm2  
        public double Cp20;
        public double Cp80;
        public double beta;                   // MPa; depending on Alloy
        public double gammaM;                 // Optional, Partial factor for material property
        public double A2;                     // Optional, for future use
        public double E;                      // Optional, for future use
        public double alpha;                  // Optional, for future use

        public double Woyp = 0;          // Section Modulus about the positive semi-axis of the y-axis 
        public double Woyn = 0;          // Section Modulus about the negative semi-axis of the y-axis 
        public double Wozp = 0;          // Section Modulus about the positive semi-axis of the z-axis 
        public double Wozn = 0;          // Section Modulus about the negative semi-axis of the z-axis 
        public double Wuyp = 0;          // Section Modulus about the positive semi-axis of the y-axis 
        public double Wuyn = 0;          // Section Modulus about the negative semi-axis of the y-axis 
        public double Wuzp = 0;          // Section Modulus about the positive semi-axis of the z-axis 
        public double Wuzn = 0;          // Section Modulus about the negative semi-axis of the z-axis 
    }
}
