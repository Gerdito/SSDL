using System;
using System.Collections.Generic;
using System.Collections;
using System.Text;
using System.IO;
using System.Data;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.Reflection;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace SSDLAdmin
{
    public class TLDDataSheetClass : SSDLBaseAttachmentClass
    {
        private int _TLDDataID = -1;
        private int _SetID = -1;

        ////private int _UnitID = -1;
        //private string _UnitCode = string.Empty;
        ////private int _BeamID = -1;
        //private string _BeamCode = string.Empty;
        ////private int _BatchID = -1;
        private string _BatchNo = string.Empty; // Keep to load from PDF forms
        //private string _SetNo = string.Empty;
        //private string _AuditType = string.Empty;
        //private string _ParticipationType = string.Empty;

        //private string _BeamType = string.Empty; // Accelerator | Co60

        //private int _DataSheetType = -1; // 1-FirstIrradiation | 2-FollowUp
        //private int _DataSheetID = -1; // Original TLDDataID from the FirstIrradiation


        //private string _DataSheetStatus = string.Empty;
        //private string _ReasonForDeviation = string.Empty;
        //private string _DeviationDescription = string.Empty;
        //private string _DataSheetComments = string.Empty;

        //private int _ContactID = -1;
        private string _ContactTitle = string.Empty;
        private string _ContactFamilyName = string.Empty;
        private string _ContactFirstName = string.Empty;
        private string _ContactPosition = string.Empty;
        private string _ContactDepartment = string.Empty;
        private string _ContactEmail = string.Empty;
        private string _ContactTelephone1 = string.Empty;
        private string _ContactTelephone2 = string.Empty;

        private string _OperatorName = string.Empty;
        private string _Street = string.Empty;
        private string _City = string.Empty;
        private string _POBox = string.Empty;
        private string _Zip = string.Empty;
        private string _State = string.Empty;
        private string _Country = string.Empty;

        private string _InstitutionalEmail = string.Empty;
        private string _InstitutionalTelephone1 = string.Empty;
        private string _InstitutionalFax = string.Empty;

        private string _CompletedByPersonFamilyName = string.Empty;
        private string _CompletedByPersonFirstName = string.Empty;
        private DateTime _CompletedDate = DateTime.MinValue;
        
        private string _IrradiatedByPersonTitle = "Off";
        private string _IrradiatedByPersonFamilyName = string.Empty;
        private string _IrradiatedByPersonFirstName = string.Empty;
        private string _IrradiatedByPersonPosition = "Off";

        private string _IrradiatedByPersonTitle2 = "Off";
        private string _IrradiatedByPersonFamilyName2 = string.Empty;
        private string _IrradiatedByPersonFirstName2 = string.Empty;
        private string _IrradiatedByPersonPosition2 = "Off";

        private string _IrradiatedByPersonTitle3 = "Off";
        private string _IrradiatedByPersonFamilyName3 = string.Empty;
        private string _IrradiatedByPersonFirstName3 = string.Empty;
        private string _IrradiatedByPersonPosition3 = "Off";

        private string _IrradiatedByPersonTitle4 = "Off";
        private string _IrradiatedByPersonFamilyName4 = string.Empty;
        private string _IrradiatedByPersonFirstName4 = string.Empty;
        private string _IrradiatedByPersonPosition4 = "Off";

        //private string _IrradiatedByPersonDepartment = string.Empty;
        //private string _IrradiatedByPersonEmail = string.Empty;
        //private string _IrradiatedByPersonTelephone1 = string.Empty;
        //private string _IrradiatedByPersonTelephone2 = string.Empty;

        private string _PreviousParticipation = "Off";
        private string _ParticipationOrganiser = "Off";

        private string _ParticipationOrganiserOther = string.Empty;

        private string _ParticipationYear = "Off";
        private string _Equipment = "Off";
        private string _Irradiator = "Off"; // RP

        private string _EquipmentCo60 = "Off";
        private string _EquipmentLinac = "Off";

        private string _EquipmentOther = string.Empty;
        private string _EquipmentProductionYear = "Off";
        private string _EquipmentInstallationYear = "Off";
        private string _EquipmentLastSourceReplacementYear = "Off";
        
        private int _EquipmentEnergy = 0;


        private string _EquipmentSerialNumber = string.Empty;
        private Double _EquipmentSourceStrength = 0;
        private string _EquipmentSourceStrengthUnits = "Off";
        private DateTime _EquipmentSourceStrengthOnDate = DateTime.MinValue;

        private string _BeamQuality = "Off";
        private Double _BeamQualityD20D10 = 0;
        private Double _BeamQualityTPR20 = 0;
        private Double _BeamQualityTPR20Distance = 0;
        private Double _BeamQualityOther = 0;
        private string _BeamQualityOtherConditions = string.Empty;

        private Double _BeamQualityR50 = 0;
        private Double _BeamQIrrFieldSize1 = 0;
        private Double _BeamQIrrFieldSize2 = 0;
        private Double _BeamQIrrDistance = 0;

        private Double _ElectronZref = 0;
        private Double _Electrondmax = 0;

        private string _DepthCurves = string.Empty;

        private string _IrradiationDepthType = string.Empty;
        private string _ChamberCalibration = string.Empty;
        private string _CalibrationEnergy = string.Empty;
        private string _CrossCalibrationExplanations = string.Empty;
        private Double _ConversionFactor = 0;


        private DateTime _IrradiationDate = DateTime.MinValue;
        private Double _IrradiationDepth = 0;
        private Double _IrradiationFieldSize1 = 0;
        private Double _IrradiationFieldSize2 = 0;
        private Double _IrradiationDistance = 0;
        private string _IrradiationDistanceType = "Off";
        private string _BeamGeometry = "Off";
        private Double _IrradiationSetting1 = 0;
        private string _IrradiationUnits1 = "Off";
        private Double _UserDose1 = 0;
        private Double _AirKerma1 = 0;
        private Double _IrradiationSetting2 = 0;
        private string _IrradiationUnits2 = "Off";
        private Double _UserDose2 = 0;
        private Double _AirKerma2 = 0;
        private Double _IrradiationSetting3 = 0;
        private string _IrradiationUnits3 = "Off";
        private Double _UserDose3 = 0;
        private Double _AirKerma3 = 0;
        private string _Factors = string.Empty;
        private Double _BeamOutput = 0;
        private string _BeamUnits = "Off";
        private DateTime _BeamOutputDate = DateTime.MinValue;
        private string _Conditions = string.Empty;
        private string _MeasuredByPersonFamilyName = string.Empty;
        private string _MeasuredByPersonFirstName = string.Empty;

        private string _MeasuredByPosition = string.Empty;
        private DateTime _MeasuredDate = DateTime.MinValue;
        private string _IonisationChamber = "Off";
        private string _IonisationChamberOther = string.Empty;

        private string _Electrometer = "Off";
        private string _ElectrometerOther = string.Empty;
        private string _CalibrationType = "Off";
        private Double _CalibrationValue = 0;
        private string _CalibrationUnit = "Off";
        private string _CalibrationLaboratory = string.Empty;
        private DateTime _CalibrationDate = DateTime.MinValue;
        private Double _Temperature = 0;
        private Double _Pressure = 0;
        private string _PressureUnit = "Off";
        private string _PhantomType = "Off";
        private string _PhantomMaterial = "Off";
        private Double _ChamberIrradiationFieldSize1 = 0;
        private Double _ChamberIrradiationFieldSize2 = 0;
        private Double _ChamberIrradiationDistance = 0;
        private string _ChamberIrradiationDistanceType = "Off";
        private string _ChamberIrradiationMeasuringPoint = "Off";
        private Double _ChamberIrradiationDepth = 0;
        private string _CapMaterial = "Off";
        private Double _CapThickness = 0;
        private Double _ReadingUncorrected = 0;
        private Double _ReadingMeasurementSetting = 0;
        private string _ReadingMeasurementSettingUnits = "Off";
        private Double _ReadingTemperature = 0;
        private Double _ReadingPressure = 0;
        private string _ReadingPressureUnit = "Off";
        private string _DosimetryProtocol = "Off";
        private Double _CorrectionSetting = 0;
        private string _CorrectionUnits = "Off";
        private string _DetailedExplanations = string.Empty;

        //private bool _Archived = false;

        private DateTime _CreatedOn = DateTime.MinValue;
        private string _CreatedBy = string.Empty;

        private DateTime _LastUpdate = DateTime.MinValue;
        private string _UpdateComment = string.Empty;

        //public TLDEvaluationClass Evaluation;

        public TLDDataSheetClass(SSDLBaseClass ParentObject)
            : base(ParentObject) // Call base-class constructor first
        {
            if (this.ParentObject is TLDSetClass)
            {
                this._SetID = (this.ParentObject as TLDSetClass).SetID;

                if ((this.ParentObject as TLDSetClass).ParentObject is OperatorClass)
                {
                    this.OperatorID = ((this.ParentObject as TLDSetClass).ParentObject as OperatorClass).OperatorID;
                    this.CCode = ((this.ParentObject as TLDSetClass).ParentObject as OperatorClass).CCode;
                    this.LabCode = ((this.ParentObject as TLDSetClass).ParentObject as OperatorClass).LabCode;
                }
            }

            this._CreatedOn = DateTime.Now;
            this._CreatedBy = "Created by user: [" + GlobalClass.User.UserName + "] IP Address: [" + GlobalClass.GetIPAddress() + "] Date: [" + this._CreatedOn.ToString() + "]";
        }


        public sealed override OperatorClass Operator
        {
            get { return null; }
        }

        public TLDSetClass TLDSet
        {
            get
            {
                if (this.ParentObject is TLDSetClass)
                    return (this.ParentObject as TLDSetClass);
                else
                    return null;
            }
        }

        #region Public TLDDataSheetClass

        public int TLDDataID
        {
            get { return this._TLDDataID; }
            set
            {
                if (this._TLDDataID != value)
                {
                    this._TLDDataID = value;
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public int SetID
        {
            get { return this._SetID; }
            set
            {
                if (this._SetID != value)
                {
                    this._SetID = value;
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public override int DocumentID
        {
            get { return _TLDDataID; }
        }


        public string BatchNo
        {
            get { return this._BatchNo; }
            set
            {
                if (this._BatchNo.Trim() != value.Trim())
                {
                    this._BatchNo = value.Replace(" ", string.Empty).Trim();
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public string ContactTitle
        {
            get { return this._ContactTitle; }
            set
            {
                if (this._ContactTitle.Trim() != value.Trim())
                {
                    this._ContactTitle = value.Trim();
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public string ContactFirstName
        {
            get { return this._ContactFirstName; }
            set
            {
                if (this._ContactFirstName.Trim() != value.Trim())
                {
                    this._ContactFirstName = value.Trim();
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public string ContactFamilyName
        {
            get { return this._ContactFamilyName; }
            set
            {
                if (this._ContactFamilyName.Trim() != value.Trim())
                {
                    this._ContactFamilyName = value.Trim();
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public string ContactPosition
        {
            get { return this._ContactPosition; }
            set
            {
                if (this._ContactPosition.Trim() != value.Trim())
                {
                    this._ContactPosition = value.Trim();
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        //Addedd by Gerd
        public string IrradiationDepthType
        {
            get { return this._IrradiationDepthType; }
            set
            {
                if (this._IrradiationDepthType.Trim() != value.Trim())
                {
                    this._IrradiationDepthType = value.Trim();
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public string ChamberCalibration
        {
            get { return this._ChamberCalibration; }
            set
            {
                if (this._ChamberCalibration.Trim() != value.Trim())
                {
                    this._ChamberCalibration = value.Trim();
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public string CalibrationEnergy
        {
            get { return this._CalibrationEnergy; }
            set
            {
                if (this._CalibrationEnergy.Trim() != value.Trim())
                {
                    this._CalibrationEnergy = value.Trim();
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public string CrossCalibrationExplanations
        {
            get { return this._CrossCalibrationExplanations; }
            set
            {
                if (this._CrossCalibrationExplanations.Trim() != value.Trim())
                {
                    this._CrossCalibrationExplanations = value.Trim();
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public double ConversionFactor
        {
            get { return this._ConversionFactor; }
            set
            {
                if (this._ConversionFactor != value)
                {
                    this._ConversionFactor = value;
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }
        //end added by Gerd


        public string ContactDepartment
        {
            get { return this._ContactDepartment; }
            set
            {
                if (this._ContactDepartment.Trim() != value.Trim())
                {
                    this._ContactDepartment = value.Trim();
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public string ContactEmail
        {
            get { return this._ContactEmail; }
            set
            {
                if (this._ContactEmail.Trim() != value.Trim())
                {
                    this._ContactEmail = value.Trim();

                    this._ContactEmail = this._ContactEmail.Replace(" ", "").Replace(",", ";");

                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public string ContactTelephone1
        {
            get { return this._ContactTelephone1; }
            set
            {
                if (this._ContactTelephone1.Trim() != value.Trim())
                {
                    this._ContactTelephone1 = value.Trim();
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public string ContactTelephone2
        {
            get { return this._ContactTelephone2; }
            set
            {
                if (this._ContactTelephone2.Trim() != value.Trim())
                {
                    this._ContactTelephone2 = value.Trim();
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public string OperatorName
        {
            get { return this._OperatorName; }
            set
            {
                if (this._OperatorName.Trim() != value.Trim())
                {
                    this._OperatorName = value.Trim();
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public string Street
        {
            get { return this._Street; }
            set
            {
                if (this._Street.Trim() != value.Trim())
                {
                    this._Street = value.Trim();
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public string City
        {
            get { return this._City; }
            set
            {
                if (this._City.Trim() != value.Trim())
                {
                    this._City = value.Trim();
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public string POBox
        {
            get { return this._POBox; }
            set
            {
                if (this._POBox.Trim() != value.Trim())
                {
                    this._POBox = value.Trim();
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public string Zip
        {
            get { return this._Zip; }
            set
            {
                if (this._Zip.Trim() != value.Trim())
                {
                    this._Zip = value.Trim();
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public string State
        {
            get { return this._State; }
            set
            {
                if (this._State.Trim() != value.Trim())
                {
                    this._State = value.Trim();
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public string Country
        {
            get { return this._Country; }
            set
            {
                if (this._Country.Trim() != value.Trim())
                {
                    this._Country = value.Trim();
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }
        
        public string EmailInstitutional
        {
            get { return this._InstitutionalEmail; }
            set
            {
                if (this._InstitutionalEmail.Trim() != value.Trim().Replace(" ", "").Replace(",", ";"))
                {
                    this._InstitutionalEmail = value.Trim();
                    this._InstitutionalEmail = this._InstitutionalEmail.Replace(" ", "").Replace(",", ";");
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public string InstitutionalEmail
        {
            get { return this._InstitutionalEmail; }
            set
            {
                if (this._InstitutionalEmail.Trim() != value.Trim().Replace(" ", "").Replace(",", ";"))
                {
                    this._InstitutionalEmail = value.Trim();
                    this._InstitutionalEmail = this._InstitutionalEmail.Replace(" ", "").Replace(",", ";");
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public string TelephoneInstitutional
        {
            get { return this._InstitutionalTelephone1; }
            set
            {
                if (this._InstitutionalTelephone1.Trim() != value.Trim())
                {
                    this._InstitutionalTelephone1 = value.Trim();
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public string InstitutionalTelephone1
        {
            get { return this._InstitutionalTelephone1; }
            set
            {
                if (this._InstitutionalTelephone1.Trim() != value.Trim())
                {
                    this._InstitutionalTelephone1 = value.Trim();
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public string InstitutionalFax
        {
            get { return this._InstitutionalFax; }
            set
            {
                if (this._InstitutionalFax.Trim() != value.Trim())
                {
                    this._InstitutionalFax = value.Trim();
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public string CompletedByPersonFamilyName
        {
            get { return this._CompletedByPersonFamilyName; }
            set
            {
                if (this._CompletedByPersonFamilyName.Trim() != value.Trim())
                {
                    this._CompletedByPersonFamilyName = value.Trim();
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public string CompletedByPersonFirstName
        {
            get { return this._CompletedByPersonFirstName; }
            set
            {
                if (this._CompletedByPersonFirstName.Trim() != value.Trim())
                {
                    this._CompletedByPersonFirstName = value.Trim();
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public DateTime CompletedDate
        {
            get { return this._CompletedDate; }
            set
            {
                if (this._CompletedDate != value)
                {
                    this._CompletedDate = value;
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }
        
        public string IrradiatedByPersonTitle
        {
            get
            {
                string sReturn = this._IrradiatedByPersonTitle;

                if (sReturn == string.Empty)
                    sReturn = "Off";

                return sReturn;
            }
            set
            {
                if (value.Trim() == string.Empty)
                    value = "Off";

                if (this._IrradiatedByPersonTitle.Trim() != value.Trim())
                {
                    this._IrradiatedByPersonTitle = value.Trim();
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }
        
        public string IrradiatedByPersonFamilyName
        {
            get { return this._IrradiatedByPersonFamilyName; }
            set
            {
                if (this._IrradiatedByPersonFamilyName.Trim() != value.Trim())
                {
                    this._IrradiatedByPersonFamilyName = value.Trim();
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public string IrradiatedByPersonFirstName
        {
            get { return this._IrradiatedByPersonFirstName; }
            set
            {
                if (this._IrradiatedByPersonFirstName.Trim() != value.Trim())
                {
                    this._IrradiatedByPersonFirstName = value.Trim();
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public string IrradiatedByPersonPosition
        {
            get
            {
                string sReturn = this._IrradiatedByPersonPosition;

                if (this.TLDSet.ParticipationType == GlobalClass.sParticipationTypeHospitals)
                {
                    if (sReturn == string.Empty)
                        sReturn = "Off";
                }
                else
                {
                    if (sReturn == "Off")
                        sReturn = string.Empty;
                }
                return sReturn;
            }
            set
            {
                if (this.TLDSet.ParticipationType == GlobalClass.sParticipationTypeHospitals)
                    if (value.Trim() == string.Empty)
                        value = "Off";

                if (this._IrradiatedByPersonPosition.Trim() != value.Trim())
                {
                    this._IrradiatedByPersonPosition = value.Trim();
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public string IrradiatedByPersonTitle2
        {
            get
            {
                string sReturn = this._IrradiatedByPersonTitle2;

                if (sReturn == string.Empty)
                    sReturn = "Off";

                return sReturn;
            }
            set
            {
                if (value.Trim() == string.Empty)
                    value = "Off";

                if (this._IrradiatedByPersonTitle2.Trim() != value.Trim())
                {
                    this._IrradiatedByPersonTitle2 = value.Trim();
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public string IrradiatedByPersonFamilyName2
        {
            get { return this._IrradiatedByPersonFamilyName2; }
            set
            {
                if (this._IrradiatedByPersonFamilyName2.Trim() != value.Trim())
                {
                    this._IrradiatedByPersonFamilyName2 = value.Trim();
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public string IrradiatedByPersonFirstName2
        {
            get { return this._IrradiatedByPersonFirstName2; }
            set
            {
                if (this._IrradiatedByPersonFirstName2.Trim() != value.Trim())
                {
                    this._IrradiatedByPersonFirstName2 = value.Trim();
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public string IrradiatedByPersonPosition2
        {
            get
            {
                string sReturn = this._IrradiatedByPersonPosition2;

                if (this.TLDSet.ParticipationType == GlobalClass.sParticipationTypeHospitals)
                {
                    if (sReturn == string.Empty)
                        sReturn = "Off";
                }
                else
                {
                    if (sReturn == "Off")
                        sReturn = string.Empty;
                }

                return sReturn;
            }
            set
            {
                if (this.TLDSet.ParticipationType == GlobalClass.sParticipationTypeHospitals)
                    if (value.Trim() == string.Empty)
                        value = "Off";

                if (this._IrradiatedByPersonPosition2.Trim() != value.Trim())
                {
                    this._IrradiatedByPersonPosition2 = value.Trim();
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public string IrradiatedByPersonTitle3
        {
            get
            {
                string sReturn = this._IrradiatedByPersonTitle3;

                if (sReturn == string.Empty)
                    sReturn = "Off";

                return sReturn;
            }
            set
            {
                if (value.Trim() == string.Empty)
                    value = "Off";

                if (this._IrradiatedByPersonTitle3.Trim() != value.Trim())
                {
                    this._IrradiatedByPersonTitle3 = value.Trim();
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public string IrradiatedByPersonFamilyName3
        {
            get { return this._IrradiatedByPersonFamilyName3; }
            set
            {
                if (this._IrradiatedByPersonFamilyName3.Trim() != value.Trim())
                {
                    this._IrradiatedByPersonFamilyName3 = value.Trim();
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public string IrradiatedByPersonFirstName3
        {
            get { return this._IrradiatedByPersonFirstName3; }
            set
            {
                if (this._IrradiatedByPersonFirstName3.Trim() != value.Trim())
                {
                    this._IrradiatedByPersonFirstName3 = value.Trim();
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public string IrradiatedByPersonPosition3
        {
            get
            {
                string sReturn = this._IrradiatedByPersonPosition3;

                if (this.TLDSet.ParticipationType == GlobalClass.sParticipationTypeHospitals)
                {
                    if (sReturn == string.Empty)
                        sReturn = "Off";
                }
                else
                {
                    if (sReturn == "Off")
                        sReturn = string.Empty;
                }

                return sReturn;
            }
            set
            {
                if (this.TLDSet.ParticipationType == GlobalClass.sParticipationTypeHospitals)
                    if (value.Trim() == string.Empty)
                        value = "Off";

                if (this._IrradiatedByPersonPosition3.Trim() != value.Trim())
                {
                    this._IrradiatedByPersonPosition3 = value.Trim();
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }


        public string IrradiatedByPersonTitle4
        {
            get 
            {
                string sReturn = this._IrradiatedByPersonTitle4;

                if (sReturn == string.Empty)
                    sReturn = "Off";

                return sReturn;
            }
            set
            {
                if (value.Trim() == string.Empty)
                    value = "Off";

                if (this._IrradiatedByPersonTitle4.Trim() != value.Trim())
                {
                    this._IrradiatedByPersonTitle4 = value.Trim();
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public string IrradiatedByPersonFamilyName4
        {
            get { return this._IrradiatedByPersonFamilyName4; }
            set
            {
                if (this._IrradiatedByPersonFamilyName4.Trim() != value.Trim())
                {
                    this._IrradiatedByPersonFamilyName4 = value.Trim();
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public string IrradiatedByPersonFirstName4
        {
            get { return this._IrradiatedByPersonFirstName4; }
            set
            {
                if (this._IrradiatedByPersonFirstName4.Trim() != value.Trim())
                {
                    this._IrradiatedByPersonFirstName4 = value.Trim();
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public string IrradiatedByPersonPosition4
        {
            get 
            {
                string sReturn = this._IrradiatedByPersonPosition4;

                if (this.TLDSet.ParticipationType == GlobalClass.sParticipationTypeHospitals)
                {
                    if (sReturn == string.Empty)
                        sReturn = "Off";
                }
                else
                {
                    if (sReturn == "Off")
                        sReturn = string.Empty;
                }

                return sReturn;
            }
            set
            {
                if (this.TLDSet.ParticipationType == GlobalClass.sParticipationTypeHospitals)
                    if (value.Trim() == string.Empty)
                        value = "Off";

                if (this._IrradiatedByPersonPosition4.Trim() != value.Trim())
                {
                    this._IrradiatedByPersonPosition4 = value.Trim();
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public string PreviousParticipation
        {
            get 
            {
                string sReturn = this._PreviousParticipation; 

                if (sReturn == string.Empty)
                    sReturn = "Off";

                return sReturn;
            }
            set
            {
                if (value.Trim() == string.Empty)
                    value = "Off";

                if (this._PreviousParticipation.Trim() != value.Trim())
                {
                    this._PreviousParticipation = value.Trim();
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public string ParticipationOrganiser
        {
            get
            {
                string sReturn = this._ParticipationOrganiser;

                if (this.TLDSet != null)
                {
                    if (this.TLDSet.ParticipationType == GlobalClass.sParticipationTypeHospitals)
                    {
                        if (sReturn == string.Empty)
                            sReturn = "Off";
                    }
                    else
                    {
                        if (sReturn == "Off")
                            sReturn = string.Empty;
                    }
                }

                return sReturn;
            }
            set
            {
                if (this.TLDSet.ParticipationType == GlobalClass.sParticipationTypeHospitals)
                {
                    if (value.Trim() == string.Empty)
                        value = "Off";
                }
                else
                {
                    if (value.Trim() == "Off")
                        value = string.Empty;
                }

                if (this._ParticipationOrganiser.Trim() != value.Trim())
                {
                    this._ParticipationOrganiser = value.Trim();
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public string ParticipationOrganiserOther
        {
            get { return this._ParticipationOrganiserOther; }
            set
            {
                if (this._ParticipationOrganiserOther.Trim() != value.Trim())
                {
                    this._ParticipationOrganiserOther = value.Trim();
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public string ParticipationYear
        {
            get
            {
                string sReturn = this._ParticipationYear;

                if (sReturn == string.Empty)
                    sReturn = "Off";

                return sReturn;
            }
            set
            {
                if (value.Trim() == string.Empty)
                    value = "Off";

                if (this._ParticipationYear.Trim() != value.Trim())
                {
                    this._ParticipationYear = value.Trim();
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public string Equipment
        {
            get
            {
                string sReturn = this._Equipment;

                if (sReturn == string.Empty)
                    sReturn = "Off";

                return sReturn;
            }
            set
            {
                if (value.Trim() == string.Empty)
                    value = "Off";

                if (this._Equipment.Trim() != value.Trim())
                {
                    this._Equipment = value.Trim();
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public string Irradiator
        {
            get
            {
                string sReturn = this._Irradiator;

                if (sReturn == string.Empty)
                    sReturn = "Off";

                return sReturn;
            }
            set
            {
                if (value.Trim() == string.Empty)
                    value = "Off";

                if (this._Irradiator.Trim() != value.Trim())
                {
                    this._Irradiator = value.Trim();
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }


        public string EquipmentCo60
        {
            get
            {
                string sReturn = this._Equipment;

                if (sReturn == string.Empty)
                    sReturn = "Off";

                return sReturn;
            }
            set
            {
                if (value.Trim() == string.Empty)
                    value = "Off";

                if (this._EquipmentCo60.Trim() != value.Trim())
                {
                    this._EquipmentCo60 = value.Trim();
                    this._EquipmentLinac = string.Empty;
                    this._Equipment = value.Trim();
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public string EquipmentLinac
        {
            get
            {
                string sReturn = this._Equipment;

                if (sReturn == string.Empty)
                    sReturn = "Off";

                return sReturn;
            }
            set
            {
                if (value.Trim() == string.Empty)
                    value = "Off";

                if (this._EquipmentLinac.Trim() != value.Trim())
                {
                    this._EquipmentLinac = value.Trim();
                    this._EquipmentCo60 = string.Empty;
                    this._Equipment = value.Trim();
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public string EquipmentOther
        {
            get { return this._EquipmentOther; }
            set
            {
                if (this._EquipmentOther.Trim() != value.Trim())
                {
                    this._EquipmentOther = value.Trim();
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public string EquipmentProductionYear
        {
            get
            {
                string sReturn = this._EquipmentProductionYear;

                if (sReturn == string.Empty)
                    sReturn = "Off";

                return sReturn;
            }
            set
            {
                if (value.Trim() == string.Empty)
                    value = "Off";

                if (this._EquipmentProductionYear.Trim() != value.Trim())
                {
                    this._EquipmentProductionYear = value.Trim();
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public string EquipmentInstallationYear
        {
            get
            {
                string sReturn = this._EquipmentInstallationYear;

                if (sReturn == string.Empty)
                    sReturn = "Off";

                return sReturn;
            }
            set
            {
                if (value.Trim() == string.Empty)
                    value = "Off";

                if (this._EquipmentInstallationYear.Trim() != value.Trim())
                {
                    this._EquipmentInstallationYear = value.Trim();
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public string EquipmentLastSourceReplacementYear
        {
            get
            {
                string sReturn = this._EquipmentLastSourceReplacementYear;

                if (sReturn == string.Empty)
                    sReturn = "Off";

                return sReturn;
            }
            set
            {
                if (value.Trim() == string.Empty)
                    value = "Off";

                if (this._EquipmentLastSourceReplacementYear.Trim() != value.Trim())
                {
                    this._EquipmentLastSourceReplacementYear = value.Trim();
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }
        
        public int EquipmentEnergy
        {
            get { return this._EquipmentEnergy; }
            set
            {
                if (this._EquipmentEnergy != value)
                {
                    this._EquipmentEnergy = value;
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public string EquipmentSerialNumber
        {
            get { return this._EquipmentSerialNumber; }
            set
            {
                if (this._EquipmentSerialNumber.Trim() != value.Trim())
                {
                    this._EquipmentSerialNumber = value.Trim();
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public Double EquipmentSourceStrength
        {
            get { return this._EquipmentSourceStrength; }
            set
            {
                if (this._EquipmentSourceStrength != value)
                {
                    this._EquipmentSourceStrength = value;
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public string EquipmentSourceStrengthUnits
        {
            get
            {
                string sReturn = this._EquipmentSourceStrengthUnits;

                if (sReturn == string.Empty)
                    sReturn = "Off";

                return sReturn;
            }
            set
            {
                if (value.Trim() == string.Empty)
                    value = "Off";

                if (this._EquipmentSourceStrengthUnits.Trim() != value.Trim())
                {
                    this._EquipmentSourceStrengthUnits = value.Trim();
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public DateTime EquipmentSourceStrengthOnDate
        {
            get { return this._EquipmentSourceStrengthOnDate; }
            set
            {
                if (this._EquipmentSourceStrengthOnDate != value)
                {
                    this._EquipmentSourceStrengthOnDate = value;
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public string BeamQuality
        {
            get
            {
                string sReturn = this._BeamQuality;

                if (sReturn == string.Empty)
                    sReturn = "Off";

                return sReturn;
            }
            set
            {
                if (value.Trim() == string.Empty)
                    value = "Off";

                if (this._BeamQuality.Trim() != value.Trim())
                {
                    this._BeamQuality = value.Trim();
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public Double EquipmentQualityValue
        {
            get
            {
                Double dReturn = 0.00;

                if (this._BeamQuality == "D20/D10")
                    dReturn = this._BeamQualityD20D10;
                else if (this._BeamQuality == "TRP20/10")
                    dReturn = (this._BeamQualityTPR20 + 0.0595) / 1.2661;
                else if (this._BeamQuality == "R50")
                    dReturn = (this._BeamQualityR50);

                else if (this._BeamQuality == "Other")
                    dReturn = this._BeamQualityOther;

                return dReturn;
            }
        }

        public Double BeamQualityD20D10
        {
            get { return this._BeamQualityD20D10; }
            set
            {
                if (this._BeamQualityD20D10 != value)
                {
                    this._BeamQualityD20D10 = value;
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public Double BeamQualityTPR20
        {
            get { return this._BeamQualityTPR20; }
            set
            {
                if (this._BeamQualityTPR20 != value)
                {
                    this._BeamQualityTPR20 = value;
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public Double BeamQualityTPR20Distance
        {
            get { return this._BeamQualityTPR20Distance; }
            set
            {
                if (this._BeamQualityTPR20Distance != value)
                {
                    this._BeamQualityTPR20Distance = value;
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public Double BeamQualityOther
        {
            get { return this._BeamQualityOther; }
            set
            {
                if (this._BeamQualityOther != value)
                {
                    this._BeamQualityOther = value;
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public string BeamQualityOtherConditions
        {
            get { return this._BeamQualityOtherConditions; }
            set
            {
                if (this._BeamQualityOtherConditions.Trim() != value.Trim())
                {
                    this._BeamQualityOtherConditions = value.Trim();
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public Double BeamQualityR50
        {
            get { return this._BeamQualityR50; }
            set
            {
                if (this._BeamQualityR50 != value)
                {
                    this._BeamQualityR50 = value;
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public Double BeamQIrrFieldSize1
        {
            get { return this._BeamQIrrFieldSize1; }
            set
            {
                if (this._BeamQIrrFieldSize1 != value)
                {
                    this._BeamQIrrFieldSize1 = value;
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public Double BeamQIrrFieldSize2
        {
            get { return this._BeamQIrrFieldSize2; }
            set
            {
                if (this._BeamQIrrFieldSize2 != value)
                {
                    this._BeamQIrrFieldSize2 = value;
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public Double BeamQIrrDistance
        {
            get { return this._BeamQIrrDistance; }
            set
            {
                if (this._BeamQIrrDistance != value)
                {
                    this._BeamQIrrDistance = value;
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }


        public Double ElectronZref
        {
            get { return this._ElectronZref; }
            set
            {
                if (this._ElectronZref != value)
                {
                    this._ElectronZref = value;
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public Double Electrondmax
        {
            get { return this._Electrondmax; }
            set
            {
                if (this._Electrondmax != value)
                {
                    this._Electrondmax = value;
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public string DepthCurves
        {
            get { return this._DepthCurves; }
            set
            {
                if (this._DepthCurves.Trim() != value.Trim())
                {
                    this._DepthCurves = value.Trim();
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }
       
        public DateTime IrradiationDate
        {
            get { return this._IrradiationDate; }
            set
            {
                if (this._IrradiationDate != value)
                {
                    this._IrradiationDate = value;
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public Double IrradiationDepth
        {
            get { return this._IrradiationDepth; }
            set
            {
                if (this._IrradiationDepth != value)
                {
                    this._IrradiationDepth = value;
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public Double IrradiationFieldSize1
        {
            get { return this._IrradiationFieldSize1; }
            set
            {
                if (this._IrradiationFieldSize1 != value)
                {
                    this._IrradiationFieldSize1 = value;
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public Double IrradiationFieldSize2
        {
            get { return this._IrradiationFieldSize2; }
            set
            {
                if (this._IrradiationFieldSize2 != value)
                {
                    this._IrradiationFieldSize2 = value;
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public Double IrradiationDistance
        {
            get { return this._IrradiationDistance; }
            set
            {
                if (this._IrradiationDistance != value)
                {
                    this._IrradiationDistance = value;
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public string IrradiationDistanceType
        {
            get
            {
                string sReturn = this._IrradiationDistanceType;

                if (sReturn == string.Empty)
                    sReturn = "Off";

                return sReturn;
            }
            set
            {
                if (value.Trim() == string.Empty)
                    value = "Off";

                if (this._IrradiationDistanceType.Trim() != value.Trim())
                {
                    this._IrradiationDistanceType = value.Trim();
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public string BeamGeometry
        {
            get
            {
                string sReturn = this._BeamGeometry;

                if (sReturn == string.Empty)
                    sReturn = "Off";

                return sReturn;
            }
            set
            {
                if (value.Trim() == string.Empty)
                    value = "Off";

                if (this._BeamGeometry.Trim() != value.Trim())
                {
                    this._BeamGeometry = value.Trim();
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public Double IrradiationSetting1
        {
            get { return this._IrradiationSetting1; }
            set
            {
                if (this._IrradiationSetting1 != value)
                {
                    this._IrradiationSetting1 = value;
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public string IrradiationUnits1
        {
            get
            {
                string sReturn = this._IrradiationUnits1;

                if (sReturn == string.Empty)
                    sReturn = "Off";

                return sReturn;
            }
            set
            {
                if (value.Trim() == string.Empty)
                    value = "Off";

                if (this._IrradiationUnits1.Trim() != value.Trim())
                {
                    this._IrradiationUnits1 = value.Trim();
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public Double UserDose1
        {
            get { return this._UserDose1; }
            set
            {
                if (String.Format("{0:0.00000}", this._UserDose1) != String.Format("{0:0.00000}", value))
                {
                    this._UserDose1 = value;
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public Double AirKerma1
        {
            get { return this._AirKerma1; }
            set
            {
                if (String.Format("{0:0.00000}", this._AirKerma1) != String.Format("{0:0.00000}", value))
                {
                    this._AirKerma1 = value;
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public Double IrradiationSetting2
        {
            get { return this._IrradiationSetting2; }
            set
            {
                if (this._IrradiationSetting2 != value)
                {
                    this._IrradiationSetting2 = value;
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public string IrradiationUnits2
        {
            get
            {
                string sReturn = this._IrradiationUnits2;

                if (sReturn == string.Empty)
                    sReturn = "Off";

                return sReturn;
            }
            set
            {
                if (value.Trim() == string.Empty)
                    value = "Off";

                if (this._IrradiationUnits2.Trim() != value.Trim())
                {
                    this._IrradiationUnits2 = value.Trim();
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public Double UserDose2
        {
            get { return this._UserDose2; }
            set
            {
                if (String.Format("{0:0.00000}", this._UserDose2) != String.Format("{0:0.00000}", value))
                {
                    this._UserDose2 = value;
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public Double AirKerma2
        {
            get { return this._AirKerma2; }
            set
            {
                if (String.Format("{0:0.00000}", this._AirKerma2) != String.Format("{0:0.00000}", value))
                {
                    this._AirKerma2 = value;
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public Double IrradiationSetting3
        {
            get { return this._IrradiationSetting3; }
            set
            {
                if (this._IrradiationSetting3 != value)
                {
                    this._IrradiationSetting3 = value;
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public string IrradiationUnits3
        {
            get
            {
                string sReturn = this._IrradiationUnits3;

                if (sReturn == string.Empty)
                    sReturn = "Off";

                return sReturn;
            }
            set
            {
                if (value.Trim() == string.Empty)
                    value = "Off";

                if (this._IrradiationUnits3.Trim() != value.Trim())
                {
                    this._IrradiationUnits3 = value.Trim();
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public Double UserDose3
        {
            get { return this._UserDose3; }
            set
            {
                if (String.Format("{0:0.00000}", this._UserDose3) != String.Format("{0:0.00000}", value))
                {
                    this._UserDose3 = value;
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public Double AirKerma3
        {
            get { return this._AirKerma3; }
            set
            {
                if (String.Format("{0:0.00000}", this._AirKerma3) != String.Format("{0:0.00000}", value))
                {
                    this._AirKerma3 = value;
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public string Factors
        {
            get 
            {
                string sNewLine = "\n";
                string sCarriageReturn = "\r";

                string sCurrentValue = this._Factors;
                sCurrentValue = sCurrentValue.Replace(sNewLine, string.Empty);
                sCurrentValue = sCurrentValue.Replace(sCarriageReturn, sCarriageReturn + sNewLine);

                return sCurrentValue; 
            }
            set
            {
                string sNewLine = "\n";
                string sCarriageReturn = "\r";
                string sCurrentValue = value.Trim();
                sCurrentValue = sCurrentValue.Replace(sNewLine, string.Empty);
                sCurrentValue = sCurrentValue.Replace(sCarriageReturn, sCarriageReturn + sNewLine);

                if (this._Factors.Trim() != sCurrentValue.Trim())
                {
                    this._Factors = sCurrentValue;
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public Double BeamOutput
        {
            get { return this._BeamOutput; }
            set
            {
                if (String.Format("{0:0.00000}", this._BeamOutput) != String.Format("{0:0.00000}", value))
                {
                    this._BeamOutput = value;
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public string BeamUnits
        {
            get
            {
                string sReturn = this._BeamUnits;

                if (sReturn == string.Empty)
                    sReturn = "Off";

                return sReturn;
            }
            set
            {
                if (value.Trim() == string.Empty)
                    value = "Off";

                if (this._BeamUnits.Trim() != value.Trim())
                {
                    this._BeamUnits = value.Trim();
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public DateTime BeamOutputDate
        {
            get { return this._BeamOutputDate; }
            set
            {
                if (this._BeamOutputDate != value)
                {
                    this._BeamOutputDate = value;
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public string Conditions
        {
            get 
            { 
                string sNewLine = "\n";
                string sCarriageReturn = "\r";

                string sCurrentValue = this._Conditions;
                sCurrentValue = sCurrentValue.Replace(sNewLine, string.Empty);
                sCurrentValue = sCurrentValue.Replace(sCarriageReturn, sCarriageReturn + sNewLine);

                return sCurrentValue; 
            }
            set
            {
                string sNewLine = "\n";
                string sCarriageReturn = "\r";
                string sCurrentValue = value.Trim();
                sCurrentValue = sCurrentValue.Replace(sNewLine, string.Empty);
                sCurrentValue = sCurrentValue.Replace(sCarriageReturn, sCarriageReturn + sNewLine);

                if (this._Conditions.Trim() != sCurrentValue.Trim())
                {
                    this._Conditions = sCurrentValue.Trim();
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public string MeasuredByPersonFamilyName
        {
            get { return this._MeasuredByPersonFamilyName; }
            set
            {
                if (this._MeasuredByPersonFamilyName.Trim() != value.Trim())
                {
                    this._MeasuredByPersonFamilyName = value.Trim();
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public string MeasuredByPersonFirstName
        {
            get { return this._MeasuredByPersonFirstName; }
            set
            {
                if (this._MeasuredByPersonFirstName.Trim() != value.Trim())
                {
                    this._MeasuredByPersonFirstName = value.Trim();
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public string MeasuredByPosition
        {
            get { return this._MeasuredByPosition; }
            set
            {
                if (this._MeasuredByPosition.Trim() != value.Trim())
                {
                    this._MeasuredByPosition = value.Trim();
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public DateTime MeasuredDate
        {
            get { return this._MeasuredDate; }
            set
            {
                if (this._MeasuredDate != value)
                {
                    this._MeasuredDate = value;
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public string IonisationChamber
        {
            get
            {
                string sReturn = this._IonisationChamber;

                if (sReturn == string.Empty)
                    sReturn = "Off";

                return sReturn;
            }
            set
            {
                if (value.Trim() == string.Empty)
                    value = "Off";

                if (this._IonisationChamber.Trim() != value.Trim())
                {
                    this._IonisationChamber = value.Trim();
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public string IonisationChamberOther
        {
            get { return this._IonisationChamberOther; }
            set
            {
                if (this._IonisationChamberOther.Trim() != value.Trim())
                {
                    this._IonisationChamberOther = value.Trim();
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public string Electrometer
        {
            get
            {
                string sReturn = this._Electrometer;

                if (sReturn == string.Empty)
                    sReturn = "Off";

                return sReturn;
            }
            set
            {
                if (value.Trim() == string.Empty)
                    value = "Off";

                if (this._Electrometer.Trim() != value.Trim())
                {
                    this._Electrometer = value.Trim();
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public string ElectrometerOther
        {
            get { return this._ElectrometerOther; }
            set
            {
                if (this._ElectrometerOther.Trim() != value.Trim())
                {
                    this._ElectrometerOther = value.Trim();
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public string CalibrationType
        {
            get
            {
                string sReturn = this._CalibrationType;

                if (sReturn == string.Empty)
                    sReturn = "Off";

                return sReturn;
            }
            set
            {
                if (value.Trim() == string.Empty)
                    value = "Off";

                if (this._CalibrationType.Trim() != value.Trim())
                {
                    this._CalibrationType = value.Trim();
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public Double CalibrationValue
        {
            get { return this._CalibrationValue; }
            set
            {
                if (this._CalibrationValue != value)
                {
                    this._CalibrationValue = value;
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public string CalibrationUnit
        {
            get
            {
                string sReturn = this._CalibrationUnit;

                if (sReturn == string.Empty)
                    sReturn = "Off";

                return sReturn;
            }
            set
            {
                if (value.Trim() == string.Empty)
                    value = "Off";

                if (this._CalibrationUnit.Trim() != value.Trim())
                {
                    this._CalibrationUnit = value.Trim();
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public string CalibrationLaboratory
        {
            get { return this._CalibrationLaboratory; }
            set
            {
                if (this._CalibrationLaboratory.Trim() != value.Trim())
                {
                    this._CalibrationLaboratory = value.Trim();
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public DateTime CalibrationDate
        {
            get { return this._CalibrationDate; }
            set
            {
                if (this._CalibrationDate != value)
                {
                    this._CalibrationDate = value;
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public Double Temperature
        {
            get { return this._Temperature; }
            set
            {
                if (this._Temperature != value)
                {
                    this._Temperature = value;
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public Double Pressure
        {
            get { return this._Pressure; }
            set
            {
                if (this._Pressure != value)
                {
                    this._Pressure = value;
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public string PressureUnit
        {
            get
            {
                string sReturn = this._PressureUnit;

                if (sReturn == string.Empty)
                    sReturn = "Off";

                return sReturn;
            }
            set
            {
                if (value.Trim() == string.Empty)
                    value = "Off";

                if (this._PressureUnit.Trim() != value.Trim())
                {
                    this._PressureUnit = value.Trim();
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public string PhantomType
        {
            get
            {
                string sReturn = this._PhantomType;

                if (sReturn == string.Empty)
                    sReturn = "Off";

                return sReturn;
            }
            set
            {
                if (value.Trim() == string.Empty)
                    value = "Off";

                if (this._PhantomType.Trim() != value.Trim())
                {
                    this._PhantomType = value.Trim();
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public string PhantomMaterial
        {
            get
            {
                string sReturn = this._PhantomMaterial;

                if (sReturn == string.Empty)
                    sReturn = "Off";

                return sReturn;
            }
            set
            {
                if (value.Trim() == string.Empty)
                    value = "Off";

                if (this._PhantomMaterial.Trim() != value.Trim())
                {
                    this._PhantomMaterial = value.Trim();
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public Double ChamberIrradiationFieldSize1
        {
            get { return this._ChamberIrradiationFieldSize1; }
            set
            {
                if (this._ChamberIrradiationFieldSize1 != value)
                {
                    this._ChamberIrradiationFieldSize1 = value;
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public Double ChamberIrradiationFieldSize2
        {
            get { return this._ChamberIrradiationFieldSize2; }
            set
            {
                if (this._ChamberIrradiationFieldSize2 != value)
                {
                    this._ChamberIrradiationFieldSize2 = value;
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public Double ChamberIrradiationDistance
        {
            get { return this._ChamberIrradiationDistance; }
            set
            {
                if (this._ChamberIrradiationDistance != value)
                {
                    this._ChamberIrradiationDistance = value;
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public string ChamberIrradiationDistanceType
        {
            get
            {
                string sReturn = this._ChamberIrradiationDistanceType;

                if (sReturn == string.Empty)
                    sReturn = "Off";

                return sReturn;
            }
            set
            {
                if (value.Trim() == string.Empty)
                    value = "Off";

                if (this._ChamberIrradiationDistanceType.Trim() != value.Trim())
                {
                    this._ChamberIrradiationDistanceType = value.Trim();
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public string ChamberIrradiationMeasuringPoint
        {
            get
            {
                string sReturn = this._ChamberIrradiationMeasuringPoint;

                if (sReturn == string.Empty)
                    sReturn = "Off";

                return sReturn;
            }
            set
            {
                if (value.Trim() == string.Empty)
                    value = "Off";

                if (this._ChamberIrradiationMeasuringPoint.Trim() != value.Trim())
                {
                    this._ChamberIrradiationMeasuringPoint = value.Trim();
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public Double ChamberIrradiationDepth
        {
            get { return this._ChamberIrradiationDepth; }
            set
            {
                if (this._ChamberIrradiationDepth != value)
                {
                    this._ChamberIrradiationDepth = value;
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public string CapMaterial
        {
            get
            {
                string sReturn = this._CapMaterial;

                if (sReturn == string.Empty)
                    sReturn = "Off";

                return sReturn;
            }
            set
            {
                if (value.Trim() == string.Empty)
                    value = "Off";

                if (this._CapMaterial.Trim() != value.Trim())
                {
                    this._CapMaterial = value.Trim();
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public Double CapThickness
        {
            get { return this._CapThickness; }
            set
            {
                if (this._CapThickness != value)
                {
                    this._CapThickness = value;
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public Double ReadingUncorrected
        {
            get { return this._ReadingUncorrected; }
            set
            {
                if (String.Format("{0:0.00000}", this._ReadingUncorrected) != String.Format("{0:0.00000}", value))
                {
                    this._ReadingUncorrected = value;
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public Double ReadingMeasurementSetting
        {
            get { return this._ReadingMeasurementSetting; }
            set
            {
                if (String.Format("{0:0.00000}", this._ReadingMeasurementSetting) != String.Format("{0:0.00000}", value))
                {
                    this._ReadingMeasurementSetting = value;
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public string ReadingMeasurementSettingUnits
        {
            get
            {
                string sReturn = this._ReadingMeasurementSettingUnits;

                if (sReturn == string.Empty)
                    sReturn = "Off";

                return sReturn;
            }
            set
            {
                if (value.Trim() == string.Empty)
                    value = "Off";

                if (this._ReadingMeasurementSettingUnits.Trim() != value.Trim())
                {
                    this._ReadingMeasurementSettingUnits = value.Trim();
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public Double ReadingTemperature
        {
            get { return this._ReadingTemperature; }
            set
            {
                if (this._ReadingTemperature != value)
                {
                    this._ReadingTemperature = value;
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public Double ReadingPressure
        {
            get { return this._ReadingPressure; }
            set
            {
                if (this._ReadingPressure != value)
                {
                    this._ReadingPressure = value;
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public string ReadingPressureUnit
        {
            get
            {
                string sReturn = this._ReadingPressureUnit;

                if (sReturn == string.Empty)
                    sReturn = "Off";

                return sReturn;
            }
            set
            {
                if (value.Trim() == string.Empty)
                    value = "Off";

                if (this._ReadingPressureUnit.Trim() != value.Trim())
                {
                    this._ReadingPressureUnit = value.Trim();
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public string DosimetryProtocol
        {
            get
            {
                string sReturn = this._DosimetryProtocol;

                if (sReturn == string.Empty)
                    sReturn = "Off";

                return sReturn;
            }
            set
            {
                if (value.Trim() == string.Empty)
                    value = "Off";

                if (this._DosimetryProtocol.Trim() != value.Trim())
                {
                    this._DosimetryProtocol = value.Trim();
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public Double CorrectionSetting
        {
            get { return this._CorrectionSetting; }
            set
            {
                if (String.Format("{0:0.00000}", this._CorrectionSetting) != String.Format("{0:0.00000}", value))
                {
                    this._CorrectionSetting = value;
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public string CorrectionUnits
        {
            get
            {
                string sReturn = this._CorrectionUnits;

                if (sReturn == string.Empty)
                    sReturn = "Off";

                return sReturn;
            }
            set
            {
                if (value.Trim() == string.Empty)
                    value = "Off";

                if (this._CorrectionUnits.Trim() != value.Trim())
                {
                    this._CorrectionUnits = value.Trim();
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public string DetailedExplanations
        {
            get
            {
                string sNewLine = "\n";
                string sCarriageReturn = "\r";

                string sCurrentValue = this._DetailedExplanations;
                sCurrentValue = sCurrentValue.Replace(sNewLine, string.Empty);
                sCurrentValue = sCurrentValue.Replace(sCarriageReturn, sCarriageReturn + sNewLine);

                return sCurrentValue;
            }
            set
            {
                string sNewLine = "\n";
                string sCarriageReturn = "\r";
                string sCurrentValue = value.Trim();
                sCurrentValue = sCurrentValue.Replace(sNewLine, string.Empty);
                sCurrentValue = sCurrentValue.Replace(sCarriageReturn, sCarriageReturn + sNewLine);

                if (this._DetailedExplanations.Trim() != sCurrentValue.Trim())
                {
                    this._DetailedExplanations = sCurrentValue.Trim();
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public DateTime LastUpdate
        {
            get { return this._LastUpdate; }
            set
            {
                if (this._LastUpdate != value)
                {
                    this._LastUpdate = value;
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public string UpdateComment
        {
            get { return this._UpdateComment; }
            set
            {
                if (this._UpdateComment.Trim() != value.Trim())
                {
                    this._UpdateComment = value.Trim();
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public DateTime CreatedOn
        {
            get { return this._CreatedOn; }
            set
            {
                if (this._CreatedOn != value)
                {
                    this._CreatedOn = value;
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public string CreatedBy
        {
            get { return this._CreatedBy; }
            set
            {
                if (this._CreatedBy.Trim() != value.Trim())
                {
                    this._CreatedBy = value.Trim();
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public int GetHolderCorrectionID()
        {
            int iReturm = -1;

            if (this.TLDSet != null)
            {
                if (this.TLDSet.AuditType == GlobalClass.sAuditTypeRT)
                {
                    if (this._BeamGeometry == "Horizontal")
                        iReturm = 3;
                    else if (this._BeamGeometry == "Vertical")
                    {
                        if (this.TLDSet.SetBeamType == GlobalClass.sBeamTypeElectron)
                        {
                            iReturm = 4;
                        }
                        else
                        {
                            if (this._IrradiationDepth == 5)
                                iReturm = 1;
                            else if (this._IrradiationDepth == 10)
                                iReturm = 2;
                        }
                    }
                }
            }
            return iReturm;
        }

        public int GetQualityCorrectionID()
        {
            int _QualityCorrectionID = -1;
            if (this.TLDSet != null)
            {
                if (this.TLDSet.AuditType == GlobalClass.sAuditTypeRP)
                {
                    if (this.TLDSet.SetBeamType == GlobalClass.sBeamTypeCo60)
                        _QualityCorrectionID = 1;
                    else if (this.TLDSet.SetBeamType == GlobalClass.sBeamTypeCs137)
                        _QualityCorrectionID = 2;
                }
            }
            return _QualityCorrectionID;
        }


        public string GetCertificateIrradiatedByPersonFamilyName()
        {
            string sReturs = string.Empty;

            if (this.TLDSet.TLDDataSheet.IrradiatedByPersonFamilyName != string.Empty)
                sReturs = sReturs + this.TLDSet.TLDDataSheet.IrradiatedByPersonFamilyName + ", ";
            if (this.TLDSet.TLDDataSheet.IrradiatedByPersonFamilyName2 != string.Empty)
                sReturs = sReturs + this.TLDSet.TLDDataSheet.IrradiatedByPersonFamilyName2 + ", ";
            if (this.TLDSet.TLDDataSheet.IrradiatedByPersonFamilyName3 != string.Empty)
                sReturs = sReturs + this.TLDSet.TLDDataSheet.IrradiatedByPersonFamilyName3 + ", ";
            if (this.TLDSet.TLDDataSheet.IrradiatedByPersonFamilyName4 != string.Empty)
                sReturs = sReturs + this.TLDSet.TLDDataSheet.IrradiatedByPersonFamilyName4 + ", ";


            sReturs = sReturs.Trim().TrimEnd(',');
            return sReturs;
        }

        public string GetCertificateIrradiatedByPersonFirstName()
        {
            string sReturs = string.Empty;

            if (this.TLDSet.TLDDataSheet.IrradiatedByPersonFirstName != string.Empty)
                sReturs = sReturs + this.TLDSet.TLDDataSheet.IrradiatedByPersonFirstName + ", ";
            if (this.TLDSet.TLDDataSheet.IrradiatedByPersonFirstName2 != string.Empty)
                sReturs = sReturs + this.TLDSet.TLDDataSheet.IrradiatedByPersonFirstName2 + ", ";
            if (this.TLDSet.TLDDataSheet.IrradiatedByPersonFirstName3 != string.Empty)
                sReturs = sReturs + this.TLDSet.TLDDataSheet.IrradiatedByPersonFirstName3 + ", ";
            if (this.TLDSet.TLDDataSheet.IrradiatedByPersonFirstName4 != string.Empty)
                sReturs = sReturs + this.TLDSet.TLDDataSheet.IrradiatedByPersonFirstName4 + ", ";

            sReturs = sReturs.Trim().TrimEnd(',');
            return sReturs;
        }

        #endregion // TLDDataSheetClass

        public void PopulateRecord(DataRow SetDr, DataTable TLDAttachmentsTable)
        {
            if (SetDr != null)
            {
                Type TLDDataSheetType = this.GetType();
                PropertyInfo[] TLDDataSheetProperties = TLDDataSheetType.GetProperties();

                foreach (PropertyInfo Property in TLDDataSheetProperties)
                {
                    if (Property != null)
                    {
                        string sFieldName = Property.Name.Trim();
                        string sFieldType = Property.PropertyType.Name;

                        if (GlobalClass.ColumnExists(SetDr.Table, sFieldName) == true)
                        {
                            if (SetDr[sFieldName] != DBNull.Value)
                            {
                                try
                                {
                                    if (Property.CanWrite == true)
                                        Property.SetValue(this, SetDr[sFieldName], null);
                                }
                                catch
                                {
                                    MessageBox.Show("Incorrect value in the field TLDSet.TLDDataSheet." + sFieldName, "Value Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                }
                            }
                        }
                    }
                }

                if (SetDr["TLDDataSheetCreatedOn"] != DBNull.Value) { this.CreatedOn = (DateTime)SetDr["TLDDataSheetCreatedOn"]; };
                if (SetDr["TLDDataSheetCreatedBy"] != DBNull.Value) { this.CreatedBy = (string)SetDr["TLDDataSheetCreatedBy"]; };
                if (SetDr["TLDDataSheetLastUpdate"] != DBNull.Value) { this.LastUpdate = (DateTime)SetDr["TLDDataSheetLastUpdate"]; };
                if (SetDr["TLDDataSheetUpdateComment"] != DBNull.Value) { this.UpdateComment = (string)SetDr["TLDDataSheetUpdateComment"]; };

                
                this.AttachmentList.Clear();
                if (TLDAttachmentsTable != null)
                {
                    foreach (DataRow AttDr in TLDAttachmentsTable.Rows)
                    {
                        if ((int)AttDr["AttachmentType"] == GlobalClass.iAttachmentDataSheet)
                        {
                            if ((int)AttDr["OperatorID"] == this.OperatorID)
                            {
                                if ((int)AttDr["DocumentID"] == this.DocumentID)
                                {
                                    AttachmentClass Attachment = new AttachmentClass((int)AttDr["AttachmentType"]);
                                    Attachment.PopulateRecord(AttDr);
                                    this.AttachmentList.Add(Attachment);
                                }
                            }
                        }
                    }
                }
                else
                    this.PopulateAttachmentList();

                this.StateStatus = GlobalClass.sStateStatusClean;
            }
        }

        public string ExportTLDDataSheet(string sExportType)
        {
            string sReturn = string.Empty;
            string sPdfTemplate = string.Empty;

            //OperatorClass Operator = GlobalClass.Manager.GetOperator(this.OperatorID);

            if (this.TLDSet != null)
            {
                if (this.TLDSet.Operator != null)
                {
                    if (this.TLDSet.AuditType == GlobalClass.sAuditTypeRT)
                    {
                        if (this.TLDSet.ParticipationType == GlobalClass.sParticipationTypeHospitals)
                        {
                            if ((this.TLDSet.SetBeamType == GlobalClass.sBeamTypeCo60) || (this.TLDSet.SetBeamType == GlobalClass.sBeamTypePhoton))
                            {
                                sPdfTemplate = GlobalClass.TLDDataSheetTemplateRTHospital;

                                if (this.TLDSet.TLDPackage != null)
                                {
                                    if (this.TLDSet.TLDPackage.CommunicationLanguage == GlobalClass.sCommunicationLanguageSpanish)
                                        sPdfTemplate = GlobalClass.sApplicationStartupPath + "\\" + GlobalClass.TLDDataSheetTemplateRTHospitalSpanish;
                                    else if (this.TLDSet.TLDPackage.CommunicationLanguage == GlobalClass.sCommunicationLanguageRussian)
                                        sPdfTemplate = GlobalClass.sApplicationStartupPath + "\\" + GlobalClass.TLDDataSheetTemplateRTHospitalRussian;
                                }
                            }
                            else if (this.TLDSet.SetBeamType == GlobalClass.sBeamTypeElectron)
                            {
                                sPdfTemplate = GlobalClass.TLDDataSheetTemplateRTHospitalElectrones;
                            }
                        }
                        else
                        {
                            if ((this.TLDSet.SetBeamType == GlobalClass.sBeamTypeCo60) || (this.TLDSet.SetBeamType == GlobalClass.sBeamTypePhoton))
                                sPdfTemplate = GlobalClass.TLDDataSheetTemplateRTSSDL;
                            else if (this.TLDSet.SetBeamType == GlobalClass.sBeamTypeElectron)
                                sPdfTemplate = GlobalClass.TLDDataSheetTemplateRTSSDLElectrones;
                        }
                    }
                    else if (this.TLDSet.AuditType == GlobalClass.sAuditTypeRP)
                    {
                        if (this.TLDSet.ParticipationType == GlobalClass.sParticipationTypeHospitals)
                            sPdfTemplate = GlobalClass.TLDDataSheetTemplateRPHospital;
                        else
                            sPdfTemplate = GlobalClass.TLDDataSheetTemplateRPSSDL;
                    }

                    if (File.Exists(sPdfTemplate))
                    {
                        string sFilePath = GlobalClass.sApplicationTempPath;
                        // string sFileName = "TLDDataSheet" + this.TLDSet.AuditType + "_" + this.TLDSet.SetNo + ".pdf";
                        string sFileName = "OSLDDataSheet" + this.TLDSet.AuditType + "_" + this.TLDSet.SetNo + ".pdf"; //change done for Paulina May 11th, 2015

                        if (!Directory.Exists(sFilePath))
                            Directory.CreateDirectory(sFilePath);

                        PdfReader pdfReader = new PdfReader(sPdfTemplate);
                        PdfStamper pdfStamper = new PdfStamper(pdfReader, new FileStream(sFilePath + sFileName, FileMode.Create), '\0', true);
                        AcroFields pdfFormFields = pdfStamper.AcroFields;

                        pdfFormFields.SetField("FormStatus", "PrePopulated");

                        pdfFormFields.SetField("OperatorID", this.TLDSet.Operator.OperatorID.ToString());
                        pdfFormFields.SetField("CCode", this.TLDSet.Operator.CCode);
                        pdfFormFields.SetField("LabCode", this.TLDSet.Operator.LabCode);
                        
                        pdfFormFields.SetField("AuditType", this.TLDSet.AuditType);

                        //pdfFormFields.SetField("BeamType", this.TLDSet.SetBeamType);
                        //pdfFormFields.SetField("BeamTypeInfo", this.TLDSet.SetBeamType);

                        pdfFormFields.SetField("SetBeamType", this.TLDSet.SetBeamType);
                        pdfFormFields.SetField("SetBeamTypeInfo", this.TLDSet.SetBeamType);

                        if (this.TLDSet.SetBeamType != "Off")
                        {
                            //pdfFormFields.SetFieldProperty("BeamType", "setfflags", PdfFormField.FF_READ_ONLY, null);
                            //pdfFormFields.RegenerateField("BeamType");
                            //pdfFormFields.SetFieldProperty("BeamTypeInfo", "setfflags", PdfFormField.FF_READ_ONLY, null);
                            //pdfFormFields.RegenerateField("BeamTypeInfo");

                            pdfFormFields.SetFieldProperty("SetBeamType", "setfflags", PdfFormField.FF_READ_ONLY, null);
                            pdfFormFields.RegenerateField("SetBeamType");
                            pdfFormFields.SetFieldProperty("SetBeamTypeInfo", "setfflags", PdfFormField.FF_READ_ONLY, null);
                            pdfFormFields.RegenerateField("SetBeamTypeInfo");

                            //iTextSharp.text.Font Black = FontFactory.GetFont(FontFactory.HELVETICA, 10, iTextSharp.text.Font.BOLD, Color.BLACK);
                            //iTextSharp.text.Font LightGray = FontFactory.GetFont(FontFactory.HELVETICA, 10, iTextSharp.text.Font.BOLD, Color.RED);
                            //pdfFormFields.SetFieldProperty("lbLinacLine1", "textfont", Black.BaseFont, null);

                            if (this.TLDSet.SetBeamType == GlobalClass.sBeamTypeCo60)
                            {
                                pdfFormFields.SetFieldProperty("Equipment", "setfflags", PdfFormField.FF_READ_ONLY, null);
                                pdfFormFields.SetFieldProperty("Equipment", "flags", PdfAnnotation.FLAGS_PRINT | PdfAnnotation.FLAGS_HIDDEN, null);
                                pdfFormFields.RegenerateField("Equipment");

                                pdfFormFields.SetFieldProperty("EquipmentCo60", "flags", PdfAnnotation.FLAGS_PRINT, null);
                                pdfFormFields.RegenerateField("EquipmentCo60");
                                pdfFormFields.SetFieldProperty("EquipmentLinac", "flags", PdfAnnotation.FLAGS_PRINT | PdfAnnotation.FLAGS_HIDDEN, null);
                                pdfFormFields.RegenerateField("EquipmentLinac");

                                pdfFormFields.SetField("Equipment", this.Equipment);
                                pdfFormFields.SetField("EquipmentCo60", this.EquipmentCo60);
                                pdfFormFields.SetField("EquipmentSerialNumber", this.EquipmentSerialNumber);
                                pdfFormFields.SetField("EquipmentProductionYear", this.EquipmentProductionYear);
                                pdfFormFields.SetField("EquipmentInstallationYear", this.EquipmentInstallationYear);
                                pdfFormFields.SetField("EquipmentLastSourceReplacementYear", this.EquipmentLastSourceReplacementYear);

                                this.IrradiationUnits1 = "min"; pdfFormFields.SetField("IrradiationUnits1", this.IrradiationUnits1);
                                this.IrradiationUnits2 = "min"; pdfFormFields.SetField("IrradiationUnits2", this.IrradiationUnits2);
                                if (this.TLDSet.ParticipationType == GlobalClass.sParticipationTypeSSDL)
                                    this.IrradiationUnits3 = "min"; pdfFormFields.SetField("IrradiationUnits3", this.IrradiationUnits3);

                                this.ReadingMeasurementSettingUnits = "min"; pdfFormFields.SetField("ReadingMeasurementSettingUnits", this.ReadingMeasurementSettingUnits);
                                this.CorrectionUnits = "s"; pdfFormFields.SetField("CorrectionUnits", this.CorrectionUnits);

                                pdfFormFields.SetFieldProperty("EquipmentEnergy", "setfflags", PdfFormField.FF_READ_ONLY, null);
                                pdfFormFields.SetFieldProperty("EquipmentEnergy", "bgcolor", iTextSharp.text.Color.LIGHT_GRAY, null);
                                pdfFormFields.RegenerateField("EquipmentEnergy");

                                pdfFormFields.SetFieldProperty("BeamQualityD20D10", "setfflags", PdfFormField.FF_READ_ONLY, null);
                                pdfFormFields.SetFieldProperty("BeamQualityD20D10", "bgcolor", iTextSharp.text.Color.LIGHT_GRAY, null);
                                pdfFormFields.RegenerateField("BeamQualityD20D10");

                                pdfFormFields.SetFieldProperty("BeamQualityTPR20", "setfflags", PdfFormField.FF_READ_ONLY, null);
                                pdfFormFields.SetFieldProperty("BeamQualityTPR20", "bgcolor", iTextSharp.text.Color.LIGHT_GRAY, null);
                                pdfFormFields.RegenerateField("BeamQualityTPR20");

                                pdfFormFields.SetFieldProperty("BeamQualityTPR20Distance", "setfflags", PdfFormField.FF_READ_ONLY, null);
                                pdfFormFields.SetFieldProperty("BeamQualityTPR20Distance", "bgcolor", iTextSharp.text.Color.LIGHT_GRAY, null);
                                pdfFormFields.RegenerateField("BeamQualityTPR20Distance");

                                pdfFormFields.SetFieldProperty("BeamQualityOther", "setfflags", PdfFormField.FF_READ_ONLY, null);
                                pdfFormFields.SetFieldProperty("BeamQualityOther", "bgcolor", iTextSharp.text.Color.LIGHT_GRAY, null);
                                pdfFormFields.RegenerateField("BeamQualityOther");

                                pdfFormFields.SetFieldProperty("BeamQualityOtherConditions", "setfflags", PdfFormField.FF_READ_ONLY, null);
                                pdfFormFields.SetFieldProperty("BeamQualityOtherConditions", "bgcolor", iTextSharp.text.Color.LIGHT_GRAY, null);
                                pdfFormFields.RegenerateField("BeamQualityOtherConditions");

                                pdfFormFields.SetFieldProperty("BeamQuality", "setfflags", PdfFormField.FF_READ_ONLY, null);
                                pdfFormFields.SetFieldProperty("BeamQuality", "bgcolor", iTextSharp.text.Color.LIGHT_GRAY, null);
                                pdfFormFields.RegenerateField("BeamQuality");



                                pdfFormFields.SetFieldProperty("lbCo60Line1", "textcolor", Color.BLACK, null);
                                pdfFormFields.SetFieldProperty("lbLinacLine1", "textcolor", Color.LIGHT_GRAY, null);
                                pdfFormFields.SetFieldProperty("lbLinacLine1Star", "textcolor", Color.LIGHT_GRAY, null);
                                pdfFormFields.SetFieldProperty("lbLinacLine2", "textcolor", Color.LIGHT_GRAY, null);
                                pdfFormFields.SetFieldProperty("lbLinacLine3_1", "textcolor", Color.LIGHT_GRAY, null);
                                pdfFormFields.SetFieldProperty("lbLinacLine3_2", "textcolor", Color.LIGHT_GRAY, null);
                                pdfFormFields.SetFieldProperty("lbLinacLine4_1", "textcolor", Color.LIGHT_GRAY, null);
                                pdfFormFields.SetFieldProperty("lbLinacLine4_2", "textcolor", Color.LIGHT_GRAY, null);
                                pdfFormFields.SetFieldProperty("lbLinacLine5_1", "textcolor", Color.LIGHT_GRAY, null);
                                pdfFormFields.SetFieldProperty("lbLinacLine5_2", "textcolor", Color.LIGHT_GRAY, null);

                                pdfFormFields.SetFieldProperty("lbCo60Line2", "textcolor", Color.BLACK, null);
                                pdfFormFields.SetFieldProperty("lbCo60Line3_1", "textcolor", Color.BLACK, null);
                                pdfFormFields.SetFieldProperty("lbCo60Line3_2", "textcolor", Color.BLACK, null);
                            }
                            else if (this.TLDSet.SetBeamType == GlobalClass.sBeamTypePhoton)
                            {
                                pdfFormFields.SetFieldProperty("Equipment", "setfflags", PdfFormField.FF_READ_ONLY, null);
                                pdfFormFields.SetFieldProperty("Equipment", "flags", PdfAnnotation.FLAGS_PRINT | PdfAnnotation.FLAGS_HIDDEN, null);
                                pdfFormFields.RegenerateField("Equipment");

                                pdfFormFields.SetFieldProperty("EquipmentCo60", "flags", PdfAnnotation.FLAGS_PRINT | PdfAnnotation.FLAGS_HIDDEN, null);
                                pdfFormFields.RegenerateField("EquipmentCo60");
                                pdfFormFields.SetFieldProperty("EquipmentLinac", "flags", PdfAnnotation.FLAGS_PRINT, null);
                                pdfFormFields.RegenerateField("EquipmentLinac");

                                pdfFormFields.SetField("Equipment", this.Equipment);
                                pdfFormFields.SetField("EquipmentLinac", this.EquipmentLinac);
                                pdfFormFields.SetField("EquipmentSerialNumber", this.EquipmentSerialNumber);
                                pdfFormFields.SetField("EquipmentProductionYear", this.EquipmentProductionYear);
                                pdfFormFields.SetField("EquipmentInstallationYear", this.EquipmentInstallationYear);

                                pdfFormFields.SetField("EquipmentEnergy", this.EquipmentEnergy.ToString());
                                //pdfFormFields.SetField("BeamQuality", this.BeamQuality);
                                //pdfFormFields.SetField("BeamQualityD20D10", this.BeamQualityD20D10.ToString());
                                //pdfFormFields.SetField("BeamQualityTPR20", this.BeamQualityTPR20.ToString());
                                //pdfFormFields.SetField("BeamQualityTPR20Distance", this.BeamQualityTPR20Distance.ToString());
                                //pdfFormFields.SetField("BeamQualityOther", this.BeamQualityOther.ToString());
                                //pdfFormFields.SetField("BeamQualityOtherConditions", this.BeamQualityOtherConditions);


                                this.IrradiationUnits1 = "MU"; pdfFormFields.SetField("IrradiationUnits1", this.IrradiationUnits1);
                                this.IrradiationUnits2 = "MU"; pdfFormFields.SetField("IrradiationUnits2", this.IrradiationUnits2);
                                if (this.TLDSet.ParticipationType == GlobalClass.sParticipationTypeSSDL)
                                    this.IrradiationUnits3 = "MU"; pdfFormFields.SetField("IrradiationUnits3", this.IrradiationUnits3);
                                this.ReadingMeasurementSettingUnits = "MU"; pdfFormFields.SetField("ReadingMeasurementSettingUnits", this.ReadingMeasurementSettingUnits);
                                this.CorrectionUnits = "MU"; pdfFormFields.SetField("CorrectionUnits", this.CorrectionUnits);

                                //this.CapMaterial = "Off";
                                //this.CapThickness = -1;

                                pdfFormFields.SetFieldProperty("EquipmentLastSourceReplacementYear", "setfflags", PdfFormField.FF_READ_ONLY, null);
                                pdfFormFields.SetFieldProperty("EquipmentLastSourceReplacementYear", "bgcolor", iTextSharp.text.Color.LIGHT_GRAY, null);
                                pdfFormFields.RegenerateField("EquipmentLastSourceReplacementYear");

                                pdfFormFields.SetFieldProperty("CapThickness", "setfflags", PdfFormField.FF_READ_ONLY, null);
                                pdfFormFields.SetFieldProperty("CapThickness", "bgcolor", iTextSharp.text.Color.LIGHT_GRAY, null);
                                pdfFormFields.RegenerateField("CapThickness");

                                pdfFormFields.SetFieldProperty("CapMaterial", "setfflags", PdfFormField.FF_READ_ONLY, null);
                                pdfFormFields.SetFieldProperty("CapMaterial", "bgcolor", iTextSharp.text.Color.LIGHT_GRAY, null);
                                pdfFormFields.RegenerateField("CapMaterial");

                                pdfFormFields.SetFieldProperty("lbCo60Line1", "textcolor", Color.LIGHT_GRAY, null);
                                pdfFormFields.SetFieldProperty("lbLinacLine1", "textcolor", Color.BLACK, null);
                                pdfFormFields.SetFieldProperty("lbLinacLine1Star", "textcolor", Color.RED, null);
                                pdfFormFields.SetFieldProperty("lbLinacLine2", "textcolor", Color.BLACK, null);
                                pdfFormFields.SetFieldProperty("lbLinacLine3_1", "textcolor", Color.BLACK, null);
                                pdfFormFields.SetFieldProperty("lbLinacLine3_2", "textcolor", Color.BLACK, null);
                                pdfFormFields.SetFieldProperty("lbLinacLine4_1", "textcolor", Color.BLACK, null);
                                pdfFormFields.SetFieldProperty("lbLinacLine4_2", "textcolor", Color.BLACK, null);
                                pdfFormFields.SetFieldProperty("lbLinacLine5_1", "textcolor", Color.BLACK, null);
                                pdfFormFields.SetFieldProperty("lbLinacLine5_2", "textcolor", Color.BLACK, null);

                                pdfFormFields.SetFieldProperty("lbCo60Line2", "textcolor", Color.LIGHT_GRAY, null);
                                pdfFormFields.SetFieldProperty("lbCo60Line3_1", "textcolor", Color.LIGHT_GRAY, null);
                                pdfFormFields.SetFieldProperty("lbCo60Line3_2", "textcolor", Color.LIGHT_GRAY, null);
                            }
                            else if (this.TLDSet.SetBeamType == GlobalClass.sBeamTypeElectron)
                            {
                                pdfFormFields.SetFieldProperty("Equipment", "setfflags", PdfFormField.FF_READ_ONLY, null);
                                pdfFormFields.SetFieldProperty("Equipment", "flags", PdfAnnotation.FLAGS_PRINT | PdfAnnotation.FLAGS_HIDDEN, null);
                                pdfFormFields.RegenerateField("Equipment");

                                pdfFormFields.SetFieldProperty("EquipmentLinac", "flags", PdfAnnotation.FLAGS_PRINT, null);
                                pdfFormFields.RegenerateField("EquipmentLinac");

                                pdfFormFields.SetField("IrradiationDistanceType", this.IrradiationDistanceType);                                

                                pdfFormFields.SetField("Equipment", this.Equipment);
                                pdfFormFields.SetField("EquipmentLinac", this.EquipmentLinac);
                                pdfFormFields.SetField("EquipmentSerialNumber", this.EquipmentSerialNumber);
                                pdfFormFields.SetField("EquipmentProductionYear", this.EquipmentProductionYear);
                                pdfFormFields.SetField("EquipmentInstallationYear", this.EquipmentInstallationYear);

                                pdfFormFields.SetField("EquipmentEnergy", this.EquipmentEnergy.ToString());
                                pdfFormFields.SetField("BeamQuality", this.BeamQuality);

                                this.IrradiationUnits1 = "MU"; pdfFormFields.SetField("IrradiationUnits1", this.IrradiationUnits1);
                                this.IrradiationUnits2 = "MU"; pdfFormFields.SetField("IrradiationUnits2", this.IrradiationUnits2);
                                this.ReadingMeasurementSettingUnits = "MU"; pdfFormFields.SetField("ReadingMeasurementSettingUnits", this.ReadingMeasurementSettingUnits);
                                this.CorrectionUnits = "MU"; pdfFormFields.SetField("CorrectionUnits", this.CorrectionUnits);
                            }

                            pdfFormFields.RegenerateField("lbCo60Line1");
                            pdfFormFields.RegenerateField("lbLinacLine1");
                            pdfFormFields.RegenerateField("lbLinacLine1Star");
                            pdfFormFields.RegenerateField("lbLinacLine2");
                            pdfFormFields.RegenerateField("lbLinacLine3_1");
                            pdfFormFields.RegenerateField("lbLinacLine3_2");
                            pdfFormFields.RegenerateField("lbLinacLine4_1");
                            pdfFormFields.RegenerateField("lbLinacLine4_2");
                            pdfFormFields.RegenerateField("lbLinacLine5_1");
                            pdfFormFields.RegenerateField("lbLinacLine5_2");
                            pdfFormFields.RegenerateField("lbCo60Line2");
                            pdfFormFields.RegenerateField("lbCo60Line3_1");
                            pdfFormFields.RegenerateField("lbCo60Line3_2");
                        }

                        if (this.TLDSet.Batch != null)
                        {
                            if (this.BatchNo != this.TLDSet.Batch.BatchNo)
                                this.BatchNo = this.TLDSet.Batch.BatchNo;

                            pdfFormFields.SetField("BatchNo", this.TLDSet.Batch.BatchNo);
                            pdfFormFields.SetFieldProperty("BatchNo", "setfflags", PdfFormField.FF_READ_ONLY, null);
                            pdfFormFields.RegenerateField("BatchNo");
                        }

                        pdfFormFields.SetField("SetID", this.TLDSet.SetID.ToString());
                        pdfFormFields.SetField("SetNo", this.TLDSet.SetNo);
                        pdfFormFields.SetFieldProperty("SetNo", "setfflags", PdfFormField.FF_READ_ONLY, null);
                        pdfFormFields.RegenerateField("SetNo");

                        pdfFormFields.SetField("ContactTitle", this.ContactTitle);
                        pdfFormFields.SetField("ContactFamilyName", this.ContactFamilyName);
                        pdfFormFields.SetField("ContactFirstName", this.ContactFirstName);

                        pdfFormFields.SetField("ContactPosition", this.ContactPosition);
                        pdfFormFields.SetField("ContactDepartment", this.ContactDepartment);
                        pdfFormFields.SetField("ContactEmail", this.ContactEmail);
                        pdfFormFields.SetField("ContactTelephone1", this.ContactTelephone1);
                        pdfFormFields.SetField("ContactTelephone2", this.ContactTelephone2);

                        pdfFormFields.SetField("OperatorName", this.OperatorName);
                        pdfFormFields.SetField("Street", this.Street);
                        pdfFormFields.SetField("City", this.City);
                        pdfFormFields.SetField("POBox", this.POBox);
                        pdfFormFields.SetField("Zip", this.Zip);
                        pdfFormFields.SetField("State", this.State);
                        pdfFormFields.SetField("Country", this.Country);
                        pdfFormFields.SetField("InstitutionalEmail", this.InstitutionalEmail);
                        pdfFormFields.SetField("InstitutionalTelephone1", this.InstitutionalTelephone1);
                        pdfFormFields.SetField("InstitutionalFax", this.InstitutionalFax);

                        pdfFormFields.SetField("PreviousParticipation", this.PreviousParticipation);
                        pdfFormFields.SetField("ParticipationOrganiser", this.ParticipationOrganiser);
                        pdfFormFields.SetField("ParticipationOrganiserOther", this.ParticipationOrganiserOther);
                        pdfFormFields.SetField("ParticipationYear", this.ParticipationYear);

                        
                        pdfFormFields.SetField("PreviousParticipation", this.PreviousParticipation);
                        pdfFormFields.SetField("PreviousParticipation", this.PreviousParticipation);
                        pdfFormFields.SetField("PreviousParticipation", this.PreviousParticipation);

                        if (sExportType == "FullExport")
                        {
                            foreach (DictionaryEntry de in pdfReader.AcroFields.Fields)
                            {
                                string sCurrentField = de.Key.ToString().Trim();

                                Type myType = this.GetType();
                                PropertyInfo Property = myType.GetProperty(sCurrentField);

                                if (Property != null)
                                {
                                    if (Property.CanRead)
                                    {
                                        string sCurrentValue = string.Empty;
                                        if (Property.PropertyType.Name == "DateTime")
                                        {
                                            DateTime dCurrentValue = (DateTime)Property.GetValue(this, null);
                                            if (dCurrentValue != DateTime.MinValue)
                                                sCurrentValue = GlobalClass.FormatTLDStringDateTimeValue(dCurrentValue);
                                        }
                                        else
                                            sCurrentValue = Property.GetValue(this, null).ToString();

                                        if (sCurrentValue == "-1")
                                            sCurrentValue = string.Empty;

                                        pdfFormFields.SetField(sCurrentField, sCurrentValue);
                                    }
                                }
                            }
                        }

                        sReturn = sFilePath + sFileName;

                        // flatten the form to remove editting options, set it to false
                        // to leave the form open to subsequent manual edits
                        pdfStamper.FormFlattening = false;

                        //pdfStamper.FormFlattening = true;

                        pdfReader.Close();
                        // close the pdf
                        pdfStamper.Close();
                    }
                    else
                        MessageBox.Show("Template " + sPdfTemplate + " does not exist.", "File does not exist.", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            return sReturn;
        }

        public string CheckBeforeSave()
        {
            string sReturn = string.Empty;

            if (GlobalClass.User != null)
            {
                if ((this._TLDDataID == -1) && (GlobalClass.User.isUserPermissionValid("TLDDataSheetCreate") == false))
                    sReturn = sReturn + "User [" + GlobalClass.User.UserName + "] does not have permissions to create TLD Data Sheet." + System.Environment.NewLine;
                if ((this._TLDDataID != -1) && (GlobalClass.User.isUserPermissionValid("TLDDataSheetEdit") == false))
                    sReturn = sReturn + "User [" + GlobalClass.User.UserName + "] does not have permissions to edit TLD Data Sheet." + System.Environment.NewLine;
            }
            else
                sReturn = sReturn + "Unknown user does not have permissions to perform this operation." + System.Environment.NewLine;

            if (this.TLDSet != null)
            {
                if (this.TLDSet.Beam != null)
                {
                    if (this.TLDSet.SetType == 1) // 1-FirstIrradiation | 2-FollowUp 
                        if (this.IrradiationDate == DateTime.MinValue)
                            sReturn = sReturn + "Please specify Irradiation Date" + System.Environment.NewLine;
                }
            }
            else
                sReturn = sReturn + "TLD Set unknown" + System.Environment.NewLine;

            // DictionaryType --------------------------------------------------------------------------
           DictionaryTypeClass DictionaryType = GlobalClass.Dictionary.GetDictionaryType(GlobalClass.Dictionary.GetDictionary("ParticipationYear"), this._ParticipationYear);
           if (DictionaryType == null)
               sReturn = sReturn + "Unknown Participation Year = [" + this._ParticipationYear + "]" + System.Environment.NewLine;

           DictionaryType = GlobalClass.Dictionary.GetDictionaryType(GlobalClass.Dictionary.GetDictionary("EquipmentProductionYear"), this._EquipmentProductionYear);
           if (DictionaryType == null)
               sReturn = sReturn + "Unknown Equipment Production Year = [" + this._EquipmentProductionYear + "]" + System.Environment.NewLine;

           DictionaryType = GlobalClass.Dictionary.GetDictionaryType(GlobalClass.Dictionary.GetDictionary("EquipmentProductionYear"), this._EquipmentInstallationYear);
           if (DictionaryType == null)
               sReturn = sReturn + "Unknown Equipment Installation Year = [" + this._EquipmentInstallationYear + "]" + System.Environment.NewLine;

           DictionaryType = GlobalClass.Dictionary.GetDictionaryType(GlobalClass.Dictionary.GetDictionary("EquipmentLastSourceReplacementYear"), this._EquipmentLastSourceReplacementYear);
           if (DictionaryType == null)
               sReturn = sReturn + "Unknown Equipment Last Source Replacement Year = [" + this._EquipmentLastSourceReplacementYear + "]" + System.Environment.NewLine;


           DictionaryType = GlobalClass.Dictionary.GetDictionaryType(GlobalClass.Dictionary.GetDictionary("EquipmentSourceStrengthUnits"), this._EquipmentSourceStrengthUnits);
           if (DictionaryType == null)
               sReturn = sReturn + "Unknown Equipment Source Strength Units = [" + this._EquipmentSourceStrengthUnits + "]" + System.Environment.NewLine;

           DictionaryType = GlobalClass.Dictionary.GetDictionaryType(GlobalClass.Dictionary.GetDictionary("BeamQuality"), this._BeamQuality);
           if (DictionaryType == null)
               sReturn = sReturn + "Unknown Beam Quality = [" + this._BeamQuality + "]" + System.Environment.NewLine;

           if (this.EquipmentEnergy != 0)
               if (this.EquipmentEnergy < 0 && this.EquipmentEnergy > 25)
                   sReturn = sReturn + "Equipment Energy must be in the range [0..25]" + System.Environment.NewLine;




           if (this._IrradiationDepth != 0)
               if (this._IrradiationDepth < 1 && this._IrradiationDepth > 20)
                   sReturn = sReturn + "Capsule depth in phantom must be in the range [1..20]" + System.Environment.NewLine;

           if (this._IrradiationFieldSize1 != 0)
               if (this._IrradiationFieldSize1 < 5 && this._IrradiationFieldSize1 > 40)
                   sReturn = sReturn + "Field size 1 for capsules irradiation must be in the range [5..40]" + System.Environment.NewLine;

           if (this._IrradiationFieldSize2 != 0)
               if (this._IrradiationFieldSize2 < 5 && this._IrradiationFieldSize2 > 40)
                   sReturn = sReturn + "Field size 2 for capsules irradiation must be in the range [5..40]" + System.Environment.NewLine;

           if (this._IrradiationDistance != 0)
               if (this._IrradiationDistance < 50 && this._IrradiationDistance > 150)
                   sReturn = sReturn + "Capsule irradiation distance must be in the range [50..150]" + System.Environment.NewLine;



           if (this._ChamberIrradiationDepth != 0)
               if (this._ChamberIrradiationDepth < 1 && this._ChamberIrradiationDepth > 20)
                   sReturn = sReturn + "Chamber depth in phantom must be in the range [1..20]" + System.Environment.NewLine;

           if (this._ChamberIrradiationFieldSize1 != 0)
               if (this._ChamberIrradiationFieldSize1 < 5 && this._ChamberIrradiationFieldSize1 > 40)
                   sReturn = sReturn + "Field size 1 for chamber measurements must be in the range [5..40]" + System.Environment.NewLine;

           if (this._ChamberIrradiationFieldSize2 != 0)
               if (this._ChamberIrradiationFieldSize2 < 5 && this._ChamberIrradiationFieldSize2 > 40)
                   sReturn = sReturn + "Field size 2 for chamber measurements must be in the range [5..40]" + System.Environment.NewLine;

           if (this._ChamberIrradiationDistance != 0)
               if (this._ChamberIrradiationDistance < 50 && this._ChamberIrradiationDistance > 150)
                   sReturn = sReturn + "Chamber irradiation distance must be in the range [50..150]" + System.Environment.NewLine;


           //DictionaryType = GlobalClass.Dictionary.GetDictionaryType(GlobalClass.Dictionary.GetDictionary("IrradiationDistanceType"), this._IrradiationDistanceType);
           //if (DictionaryType == null)
           //    sReturn = sReturn + "Unknown Irradiation Distance Type = [" + this._IrradiationDistanceType + "]" + System.Environment.NewLine;
            
           DictionaryType = GlobalClass.Dictionary.GetDictionaryType(GlobalClass.Dictionary.GetDictionary("IrradiationUnits1"), this._IrradiationUnits1);
           if (DictionaryType == null)
               sReturn = sReturn + "Unknown Irradiation Units 1 = [" + this._IrradiationUnits1 + "]" + System.Environment.NewLine;

           DictionaryType = GlobalClass.Dictionary.GetDictionaryType(GlobalClass.Dictionary.GetDictionary("IrradiationUnits2"), this._IrradiationUnits2);
           if (DictionaryType == null)
               sReturn = sReturn + "Unknown Irradiation Units 2 = [" + this._IrradiationUnits2 + "]" + System.Environment.NewLine;

           DictionaryType = GlobalClass.Dictionary.GetDictionaryType(GlobalClass.Dictionary.GetDictionary("IrradiationUnits3"), this._IrradiationUnits3);
           if (DictionaryType == null)
               sReturn = sReturn + "Unknown Irradiation Units 3 = [" + this._IrradiationUnits3 + "]" + System.Environment.NewLine;

           DictionaryType = GlobalClass.Dictionary.GetDictionaryType(GlobalClass.Dictionary.GetDictionary("BeamUnits"), this._BeamUnits);
           if (DictionaryType == null)
               sReturn = sReturn + "Unknown Beam Units = [" + this._BeamUnits + "]" + System.Environment.NewLine;

           DictionaryType = GlobalClass.Dictionary.GetDictionaryType(GlobalClass.Dictionary.GetDictionary("CalibrationType"), this._CalibrationType);
           if (DictionaryType == null)
               sReturn = sReturn + "Unknown Calibration Type = [" + this._CalibrationType + "]" + System.Environment.NewLine;

           DictionaryType = GlobalClass.Dictionary.GetDictionaryType(GlobalClass.Dictionary.GetDictionary("CalibrationUnit"), this._CalibrationUnit);
           if (DictionaryType == null)
               sReturn = sReturn + "Unknown Calibration Units = [" + this._CalibrationUnit + "]" + System.Environment.NewLine;

           DictionaryType = GlobalClass.Dictionary.GetDictionaryType(GlobalClass.Dictionary.GetDictionary("PressureUnit"), this._PressureUnit);
           if (DictionaryType == null)
               sReturn = sReturn + "Unknown Pressure Units = [" + this._PressureUnit + "]" + System.Environment.NewLine;

           DictionaryType = GlobalClass.Dictionary.GetDictionaryType(GlobalClass.Dictionary.GetDictionary("PressureUnit"), this._PressureUnit);
           if (DictionaryType == null)
               sReturn = sReturn + "Unknown Pressure Units = [" + this._PressureUnit + "]" + System.Environment.NewLine;

           DictionaryType = GlobalClass.Dictionary.GetDictionaryType(GlobalClass.Dictionary.GetDictionary("PhantomMaterial"), this._PhantomMaterial);
           if (DictionaryType == null)
               sReturn = sReturn + "Unknown Phantom Material = [" + this._PhantomMaterial + "]" + System.Environment.NewLine;

           DictionaryType = GlobalClass.Dictionary.GetDictionaryType(GlobalClass.Dictionary.GetDictionary("ChamberIrradiationMeasuringPoint"), this._ChamberIrradiationMeasuringPoint);
           if (DictionaryType == null)
               sReturn = sReturn + "Unknown Chamber Irradiation Measuring Point = [" + this._ChamberIrradiationMeasuringPoint + "]" + System.Environment.NewLine;

           DictionaryType = GlobalClass.Dictionary.GetDictionaryType(GlobalClass.Dictionary.GetDictionary("CapMaterial"), this._CapMaterial);
           if (DictionaryType == null)
               sReturn = sReturn + "Unknown Cap Material = [" + this._CapMaterial + "]" + System.Environment.NewLine;

           DictionaryType = GlobalClass.Dictionary.GetDictionaryType(GlobalClass.Dictionary.GetDictionary("ReadingMeasurementSettingUnits"), this._ReadingMeasurementSettingUnits);
           if (DictionaryType == null)
               sReturn = sReturn + "Unknown Reading Measurement Setting Units = [" + this._ReadingMeasurementSettingUnits + "]" + System.Environment.NewLine;

           DictionaryType = GlobalClass.Dictionary.GetDictionaryType(GlobalClass.Dictionary.GetDictionary("ReadingPressureUnit"), this._ReadingPressureUnit);
           if (DictionaryType == null)
               sReturn = sReturn + "Unknown Reading Pressure Units = [" + this._ReadingPressureUnit + "]" + System.Environment.NewLine;


           DictionaryType = GlobalClass.Dictionary.GetDictionaryType(GlobalClass.Dictionary.GetDictionary("CorrectionUnits"), this._CorrectionUnits);
           if (DictionaryType == null)
               sReturn = sReturn + "Unknown Correction Units = [" + this._CorrectionUnits + "]" + System.Environment.NewLine;

           //Equipment --------------------------------------------------------------------------
            /*
           EquipmentTypeClass Equipment = GlobalClass.Dictionary.GetEquipmentItem(GlobalClass.Dictionary.GetEquipment("IonisationChamber"), this._IonisationChamber);
           if (Equipment == null)
               sReturn = sReturn + "Unknown Ionisation Chamber = [" + this._IonisationChamber + "]" + System.Environment.NewLine;

           Equipment = GlobalClass.Dictionary.GetEquipmentItem(GlobalClass.Dictionary.GetEquipment("Electrometer"), this._Electrometer);
           if (Equipment == null)
               sReturn = sReturn + "Unknown Electrometer = [" + this._Electrometer + "]" + System.Environment.NewLine;
            */
           EquipmentTypeClass Equipment = GlobalClass.Dictionary.GetEquipmentItem(GlobalClass.Dictionary.GetEquipment("DosimetryProtocol"), this._DosimetryProtocol);
           if (Equipment == null)
               sReturn = sReturn + "Unknown Dosimetry Protocol = [" + this._DosimetryProtocol + "]" + System.Environment.NewLine;


            //--------------------------------------------------------------------------

            if ((this._IrradiationSetting1 < 0) && (this._IrradiationSetting1 > 720))
                sReturn = sReturn + "Irradiation time for capsule I must be in the range 0-720 min" + System.Environment.NewLine;

            if ((this._IrradiationSetting2 < 0) && (this._IrradiationSetting2 > 720))
                sReturn = sReturn + "Irradiation time for capsule II must be in the range 0-720 min" + System.Environment.NewLine;

            if ((this._IrradiationSetting3 < 0) && (this._IrradiationSetting3 > 720))
                sReturn = sReturn + "Irradiation time for capsule III must be in the range 0-720 min" + System.Environment.NewLine;


            if (this.TLDSet.AuditType == GlobalClass.sAuditTypeRT)
            {
                if (this._UserDose1 > 0)
                {
                    if ((this._UserDose1 < 1) && (this._UserDose1 > 3))
                        sReturn = sReturn + "User Dose for capsule I must be in the range 1-3 Gy" + System.Environment.NewLine;
                }
                else if (this._UserDose1 < 0)
                    sReturn = sReturn + "User Dose for capsule I can to be negative" + System.Environment.NewLine;

                if (this._UserDose2 > 0)
                {
                    if ((this._UserDose2 < 1) && (this._UserDose2 > 3))
                        sReturn = sReturn + "User Dose for capsule II must be in the range 1-3 Gy" + System.Environment.NewLine;
                }
                else if (this._UserDose2 < 0)
                    sReturn = sReturn + "User Dose for capsule II can to be negative" + System.Environment.NewLine;

                if (this._UserDose3 > 0)
                {
                    if ((this._UserDose3 < 1) && (this._UserDose3 > 3))
                        sReturn = sReturn + "User Dose for capsule III must be in the range 1-3 Gy" + System.Environment.NewLine;
                }
                else if (this._UserDose3 < 0)
                    sReturn = sReturn + "User Dose for capsule III can to be negative" + System.Environment.NewLine;
            }
            else if (this.TLDSet.AuditType == GlobalClass.sAuditTypeRP)
            {
                if (this._AirKerma1 > 0)
                {
                    if ((this._AirKerma1 < 1) && (this._AirKerma1 > 10))
                        sReturn = sReturn + "Air kerma for capsule I must be in the range 1-10 mGy" + System.Environment.NewLine;
                }
                else if (this._AirKerma1 < 0)
                    sReturn = sReturn + "Air kerma for capsule I can to be negative" + System.Environment.NewLine;

                if (this._AirKerma1 > 0)
                {
                    if ((this._AirKerma2 < 1) && (this._AirKerma2 > 10))
                        sReturn = sReturn + "Air kerma for capsule II must be in the range 1-10 mGy" + System.Environment.NewLine;
                }
                else if (this._AirKerma2 < 0)
                    sReturn = sReturn + "Air kerma for capsule II can to be negative" + System.Environment.NewLine;

                if (this._AirKerma1 > 0)
                {
                    if ((this._AirKerma3 < 1) && (this._AirKerma3 > 10))
                        sReturn = sReturn + "Air kerma for capsule III must be in the range 1-10 mGy" + System.Environment.NewLine;
                }
                else if (this._AirKerma3 < 0)
                    sReturn = sReturn + "Air kerma for capsule III can to be negative" + System.Environment.NewLine;
            }


            if (this._PressureUnit == "kPa")
            {
                if ((this._Pressure < 60) && (this._Pressure > 110))
                    sReturn = sReturn + "Pressure must be in the range 60-110 kPa" + System.Environment.NewLine;
            }
            else if (this._PressureUnit == "mm Hg")
            {
                if ((this._Pressure < 450) && (this._Pressure > 825))
                    sReturn = sReturn + "Pressure must be in the range 450-825 mm Hg" + System.Environment.NewLine;
            }

            if (this._ReadingPressureUnit == "kPa")
            {
                if ((this._ReadingPressure < 60) && (this._ReadingPressure > 110))
                    sReturn = sReturn + "Pressure for chamber measurements must be in the range 60-110 kPa" + System.Environment.NewLine;
            }
            else if (this._ReadingPressureUnit == "mm Hg")
            {
                if ((this._ReadingPressure < 450) && (this._ReadingPressure > 825))
                    sReturn = sReturn + "Pressure for chamber measurements must be in the range 450-825 mm Hg" + System.Environment.NewLine;
            }

            if ((this._ReadingMeasurementSetting < 0) && (this._ReadingMeasurementSetting > 720))
                sReturn = sReturn + "Time for chamber measurements must be in the range 0-720 min" + System.Environment.NewLine;

            if ((this._Temperature < 0) && (this._Temperature > 50))
                sReturn = sReturn + "Temperature for calibration coefficient must be in the range 0-50 C" + System.Environment.NewLine;

            if ((this._ReadingTemperature < 0) && (this._ReadingTemperature > 50))
                sReturn = sReturn + "Temperature for chamber measurements must be in the range 0-50 C" + System.Environment.NewLine;

            return sReturn;
        }

        public int SaveTLDDataSheet()
        {
            int iReturn = 0;

            // Do not overwright Last Update values
            //this._LastUpdate = DateTime.Now;
            //this._UpdateComment = "Updated by user: [" + GlobalClass.User.UserName + "] IP Address: [" + GlobalClass.GetIPAddress() + "] Date: [" + _LastUpdate.ToString() + "]";

            string sSql = "SELECT TLDDataID FROM dbo.TLDDataSheet WHERE TLDDataID = '" + this._TLDDataID.ToString() + "'";
            DataTable dataTable = GlobalClass.GetDataTable("CheckRecord", sSql);

            if (dataTable != null)
            {
                if (dataTable.Rows.Count == 0)
                {
                    List<ParameterClass> ParameterList = new List<ParameterClass>();
                    string sSqlStatement = GlobalClass.GenerateInsertStatement("dbo.TLDDataSheet", this, ParameterList);

                    if (sSqlStatement != string.Empty)
                    {
                        try
                        {
                            if (ParameterList.Count > 0)
                                iReturn = GlobalClass.ExecuteSQL(sSqlStatement, ParameterList);
                            else
                                iReturn = GlobalClass.ExecuteSQL(sSqlStatement);

                            GlobalClass.LogUserAction(2, this.TLDSet.OperatorID, "User [#UserName#] Execute INSERT on dbo.TLDDataSheet table", sSql);

                            if (this._TLDDataID == -1)
                            {
                                dataTable = GlobalClass.GetDataTable("MaxValue", "SELECT MAX(TLDDataID) as TLDDataID FROM dbo.TLDDataSheet");

                                if (dataTable != null)
                                    if (dataTable.Rows.Count == 1)
                                        if (dataTable.Rows[0]["TLDDataID"] != DBNull.Value)
                                            this._TLDDataID = (int)dataTable.Rows[0]["TLDDataID"];

                                dataTable = null;
                            }

                            if (this._TLDDataID != -1)
                            {
                                foreach (AttachmentClass Attachment in this.AttachmentList)
                                {
                                    if (Attachment.DocumentID != this._TLDDataID)
                                        Attachment.DocumentID = this._TLDDataID;
                                }
                                this.SaveAttachments();
                            }
                        }
                        catch
                        {
                            iReturn = 0;

                            GlobalClass.LogUserAction(-2, this.TLDSet.OperatorID, "ERROR! User [#UserName#] Execute INSERT on dbo.TLDDataSheet table", sSql);
                            MessageBox.Show("Incorrect SQL statement.", "SQL Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }

                }
                else if (dataTable.Rows.Count == 1)
                {
                    List<ParameterClass> ParameterList = new List<ParameterClass>();
                    string sSqlStatement = GlobalClass.GenerateUpdateStatement("dbo.TLDDataSheet", this, ParameterList);

                    if (sSqlStatement != string.Empty)
                    {
                        try
                        {
                            if (ParameterList.Count > 0)
                                iReturn = GlobalClass.ExecuteSQL(sSqlStatement, ParameterList);
                            else
                                iReturn = GlobalClass.ExecuteSQL(sSqlStatement);

                            foreach (AttachmentClass Attachment in this.AttachmentList)
                            {
                                if (Attachment.DocumentID != this._TLDDataID)
                                    Attachment.DocumentID = this._TLDDataID;
                            }
                            this.SaveAttachments();

                            GlobalClass.LogUserAction(2, this.TLDSet.OperatorID, "User [#UserName#] Execute UPDATE on dbo.TLDDataSheet table", sSql);
                        }
                        catch
                        {
                            iReturn = 0;

                            GlobalClass.LogUserAction(-2, this.TLDSet.OperatorID, "ERROR! User [#UserName#] Execute UPDATE on dbo.TLDDataSheet table", sSql);
                            MessageBox.Show("Incorrect SQL statement.", "SQL Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }
            }

            // Update TLDSet ExportStatus
            if (this.TLDSet != null)
                this.TLDSet.UpdateTLDSetExportStatus();

            this.StateStatus = GlobalClass.sStateStatusClean;
            return iReturn;
        }

        public int DeleteTLDDataSheet()
        {
            int iReturn = 0;
            string sSql = string.Empty;
            DataTable dataTable = null;
            int iRowCount = 0;
            string sErrorMessage = string.Empty;

            if (this._TLDDataID > 0)
            {
                // Check TLDEvaluations Table
                sSql = "SELECT EvaluationID FROM dbo.TLDEvaluations WHERE SetID = " + this._SetID.ToString();
                dataTable = GlobalClass.GetDataTable("CheckRecord", sSql);

                if (dataTable != null)
                {
                    if (dataTable.Rows.Count > 0)
                    {
                        iRowCount = iRowCount + dataTable.Rows.Count;
                        sErrorMessage = sErrorMessage + dataTable.Rows.Count.ToString() + " Evaluation has been created." + System.Environment.NewLine;
                    }
                }

                // Check TLDCertificates Table
                sSql = "SELECT CertificateID FROM dbo.TLDCertificates WHERE SetID = " + this._SetID.ToString();
                dataTable = GlobalClass.GetDataTable("CheckRecord", sSql);

                if (dataTable != null)
                {
                    if (dataTable.Rows.Count > 0)
                    {
                        iRowCount = iRowCount + dataTable.Rows.Count;
                        sErrorMessage = sErrorMessage + dataTable.Rows.Count.ToString() + " Certificate has been created." + System.Environment.NewLine;
                    }
                }

                if (iRowCount == 0)
                {
                    sSql = string.Empty;
                    sSql = sSql + "DELETE FROM dbo.TLDAttachments WHERE AttachmentType = " + GlobalClass.iAttachmentDataSheet.ToString() + " and OperatorID = " + this.TLDSet.OperatorID.ToString() + " and DocumentID = " + this._TLDDataID.ToString() + ";" + System.Environment.NewLine;
                    sSql = sSql + "DELETE FROM dbo.TLDDataSheet WHERE TLDDataID = " + this._TLDDataID.ToString();


                    try
                    {
                        iReturn = GlobalClass.ExecuteSQL(sSql);
                        GlobalClass.LogUserAction(2, this.TLDSet.OperatorID, "User [#UserName#] Execute DELETE on TLDDataSheet " + this._TLDDataID.ToString(), sSql);
                    }
                    catch
                    {
                        iReturn = 0;
                        GlobalClass.LogUserAction(-2, this.TLDSet.OperatorID, "ERROR! User [#UserName#] Execute DELETE on TLDDataSheet " + this._TLDDataID.ToString(), sSql);

                        MessageBox.Show("Incorrect SQL statement.", "SQL Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }

                    // Update TLDSet ExportStatus
                    if (this.TLDSet != null)
                        this.TLDSet.UpdateTLDSetExportStatus();

                }
                else
                {
                    MessageBox.Show("Can not delete selected TLD DataSheet - Set No " + this.TLDSet.SetNo + System.Environment.NewLine + sErrorMessage, "Check integraty Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

            }

            return iReturn;
        }
    }
}
