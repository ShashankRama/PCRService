namespace PCReports.Models.Facade
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
        public double ModelWidth { get; set; }
        public double ModelHeight { get; set; }
        public double ModelOriginX { get; set; }
        public double ModelOriginY { get; set; }

        // Section 0. Project Info
        public string ProjectGuid { get; set; }
        public string ProblemGuid { get; set; }
        public string Language { get; set; }
        public string ProjectName { get; set; }
        public string Location { get; set; }
        public string ConfigurationName { get; set; }
        public string UserName { get; set; }
        public string UserNotes { get; set; }

        // Section 1. Facade Information
        public string ProfileSystem { get; set; }
        public string MajorMullionProfile { get; set; }
        public string TransomProfile { get; set; }
        public string MinorMullionProfile { get; set; }
        public double MajorMullionProfileWeight { get; set; }
        public double TransomProfileWeight { get; set; }
        public double MinorMullionProfileWeight { get; set; }
        public List<GlassType> GlassTypes { get; set; }
        public double BlockDistance { get; set; }

        // Section 2. Applied Load
        public double WindLoad { get; set; }
        public string CpeString { get; set; }
        public string pCpiString { get; set; }
        public string nCpiString { get; set; }
        public double HorizontalLiveLoad { get; set; }
        public double HorizontalLiveLoadHeight { get; set; }
        public double WindLoadFactor { get; set; }
        public double HorizontalLiveLoadFactor { get; set; }
        public double DeadLoadFactor { get; set; }

        // Section 4.
        public string AllowableDeflectionLine1 { get; set; }
        public string AllowableDeflectionLine2 { get; set; }
        public string AllowableInplaneDeflectionLine { get; set; }

        // Section 5.
        public string Alloys { get; set; }
        public double Beta { get; set; }
        public double AluminumReinforcementBeta { get; set; }
        public double SteelReinforcementBeta { get; set; }

        // Section Appendix
        public List<FacadeSection> FacadeSections { get; set; }

        // Section 6. Results
        public List<PDFMemberResult> PDFMemberResults { get; set; }
    }

    public class PDFMemberResult
    {
        // 6.1 member
        public int MemberID { get; set; }
        public int MemberType { get; set; }
        public string ArticleName { get; set; }
        public double Length { get; set; }

        public double TributaryArea { get; set; }
        public double Cp { get; set; }
        public double TributaryAreaFactor { get; set; }
        public double AppliedWindLoad { get; set; }

        public double VerticalDispIndex { get; set; }

        // profile parameters
        public double Depth { get; set; }
        public double Width { get; set; }
        public double A { get; set; }
        public double Iyy { get; set; }
        public double Izz { get; set; }
        public double Wyy { get; set; }
        public double Wzz { get; set; }

        // reinforcement
        public string? ReinfArticleName { get; set; }
        public double ReinfDepth { get; set; }
        public double ReinfWidth { get; set; }
        public double ReinfA { get; set; }
        public double ReinfIyy { get; set; }
        public double ReinfIzz { get; set; }
        public double ReinfWyy { get; set; }
        public double ReinfWzz { get; set; }
        public double ReinfBeta { get; set; }
        public double ReinfEIy { get; set; }
        public double ReinfEIz { get; set; }

        // splice joint
        public List<double> SpliceJointX { get; set; }
        public List<string> SpliceJointType { get; set; }

        // slab anchor
        public List<double> SlabAnchorX { get; set; }
        public List<string> SlabAnchorType { get; set; }

        // reinforcement
        public bool isReinforced { get; set; }

        // beam data
        public List<LoadData> WindLoadDataList { get; set; }
        public List<LoadData> HLLDataList { get; set; }
        public List<LoadData> WeightDataList { get; set; }
        public double[] ReactionForces { get; set; }   // [windLeftReactionForce, windRightReactionForce, hllLeftReactionForce, hllRightReactionForce]
        public int[] HLLWorstLoadCase { get; set; }
        public double MaxMyLC1 { get; set; }
        public double MaxMyLC2 { get; set; }
        public double MaxMzLC3 { get; set; }

        // curves
        public double[] CurveX { get; set; }
        public double[] ShearCurveWind { get; set; }
        public double[] ShearCurveHLL { get; set; }
        public double[] ShearCurveWeight { get; set; }
        public double[] MomentCurveWind { get; set; }
        public double[] MomentCurveHLL { get; set; }
        public double[] MomentCurveWeight { get; set; }
        public double[] OutOfPlaneDeflectionCurveWind { get; set; }
        public double[] InPlaneDeflectionCurveWeight { get; set; }

        // table
        public double[] UtilizationCheckTable { get; set; }

        public double[] DeflectionCheckTable { get; set; }
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
        public double[] CornerCoordinates { get; set; }
        public double[]? VentCoordinates { get; set; }
        public double[]? InsertOuterFrameCoordinates { get; set; }
        public string VentOpeningDirection { get; set; }
        public string VentOperableType { get; set; }
    }

    public class GlassType
    {
        public string GlassIDs { get; set; }
        public string Weight { get; set; }
        public string Description { get; set; }
    }
    public class LoadData
    {
        public int LoadType { get; set; }
        public int LoadSide { get; set; }
        public double x1 { get; set; }
        public double x2 { get; set; }
        public double q1 { get; set; }
        public double q2 { get; set; }
        public bool isConcentratedLoad { get; set; }
        public bool isBoundaryReaction { get; set; }
    }

    public class FacadeSection : ICloneable
    {
        public object Clone()
        {
            return this.MemberwiseClone();
        }

        public int SectionID { get; set; }           // = 1,2,3,4 SectionType = SectionID + 3
        public int SectionType { get; set; }         // same as MemberType. 4: major mullion, 5: transom, 6: minor mullion, 7: reinforcement
                                        // 21:UDC Top Frame; 22: UDC Vertical;  23: UDC Bottom Frame; 24: UDC Vertical Glazing Bar; 25: UDC Horizontal Glazing Bar;
        public string ArticleName { get; set; }      // For Schuco standard article, use article ID.
        public bool isCustomProfile { get; set; }    // is custom profile
        public double OutsideW { get; set; }         // for display in pdf only
        public double BTDepth { get; set; }
        public double Width { get; set; }
        public double Zo { get; set; }          // depth from central axis to the top
        public double Zu { get; set; }          // depth from central axis to the bottom
        public double Zl { get; set; }          // depth from central axis to the mullion left
        public double Zr { get; set; }          // depth from central axis to the mullion right
        public double A { get; set; }           //Area
        public string Material { get; set; }        // Aluminum or Steel
        public double beta { get; set; }          // MPa; depending on Alloy
        public double Weight { get; set; }      //section weight kg/m
        public double Iyy { get; set; }         //Moment of inertia for bending about the y-axis, out-of-plane
        public double Izz { get; set; }           //Moment of inertia for bending about the z-axis, in-plane
        public double Wyy { get; set; }          // min(Iyy / Zo, Izz/ Zu)
        public double Wzz { get; set; }          // min(Izz / Zyl, Izz/ Zyr)
        public double Asy { get; set; }         //Shear area in y-direction
        public double Asz { get; set; }         //Shear area in z-direction
        public double J { get; set; }           //Torsional constant 
        public double E { get; set; }            //Young's modulus
        public double G { get; set; }            //Torsional shear modulus
        public double EA { get; set; }                   // Derived EA
        public double GAsy { get; set; }                 // Derived GA
        public double GAsz { get; set; }                 // Derived GA
        public double EIy { get; set; }                  // Derived EIy
        public double EIz { get; set; }                  // Derived EIz
        public double GJ { get; set; }                   // Derived GJ
        // sectional properties for display only
        public double Ys { get; set; }          // Shear center y coordinate
        public double Zs { get; set; }           // Shear center z coordinate
        public double Ry { get; set; }           // Radius of Gyration about the y-axis
        public double Rz { get; set; }          // Radius of Gyration about the z-axis
        public double Wyp { get; set; }         // Section Modulus about the positive semi-axis of the y-axis 
        public double Wyn { get; set; }         // Section Modulus about the negative semi-axis of the y-axis 
        public double Wzp { get; set; }         // Section Modulus about the positive semi-axis of the z-axis 
        public double Wzn { get; set; }         // Section Modulus about the negative semi-axis of the z-axis 
        public double Cw { get; set; }
        public double Beta_torsion { get; set; }
        public double Zy { get; set; }          // Plastic Property 
        public double Zz { get; set; }        // Plastic Property 

        // Reinforcement properties, for physics core internal only
        public double ReinforcementEIy;
        public double ReinforcementEIz;
        public double ReinforcementWeight;
    }
}
