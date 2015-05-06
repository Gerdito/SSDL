using System;
using System.Text;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Collections;
using Outlook = Microsoft.Office.Interop.Outlook;
using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.collection;

using System.Reflection;  // reflection namespace
using StatDescriptive;
using System.Net;
using System.Drawing;
using System.Drawing.Printing;
using System.Diagnostics;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;

//using Outlook = Microsoft.Office.Interop.Outlook;

namespace SSDLAdmin
{
    public class GlobalClass
    {
        //public static string sDB = "DEV";
        public static string sDB = "";

        //public static string ConnectionStringRead = "integrated security=SSPI;Data Source=NucMedicineDB" + sDB + ";persist security info=False;Initial Catalog=SSDL";
        //public static string ConnectionStringWrite = "integrated security=SSPI;Data Source=NucMedicineDB" + sDB + ";persist security info=False;Initial Catalog=SSDL";

        //public static string ConnectionStringRead = "Data Source=q120dev;Initial Catalog=SSDL;User Id=SSDLRead;Password=skdrhK57824$30957; Connection Timeout=60";
        //public static string ConnectionStringWrite = "Data Source=q120dev;Initial Catalog=SSDL;User Id=SSDLWrite;Password=Jalstkhwaffgds6785$; Connection Timeout=60";

        public static string ConnectionStringRead = "Data Source=SSDLDB" + sDB + ";Initial Catalog=SSDL;User Id=SSDLRead;Password=skdrhK57824$30957; Connection Timeout=60";
        public static string ConnectionStringWrite = "Data Source=SSDLDB" + sDB + ";Initial Catalog=SSDL;User Id=SSDLWrite;Password=Jalstkhwaffgds6785$; Connection Timeout=60";

        public static string ConnectionStringYP = "Data Source=SSDLDB" + sDB + ";Initial Catalog=SSDL; Integrated Security = SSPI; Connection Timeout=60";
        public static string ConnectionStringDIRACRead = "Data Source=DIRACDB" + sDB + ";Initial Catalog=DIRAC;User Id=NAHUNETRead;Password=JcdpN*p8!C-.zW0m; Connection Timeout=60";
        public static string ConnectionStringDIRACWrite = "Data Source=DIRACDB" + sDB + ";Initial Catalog=DIRAC;User Id=NAHUNETWrite;Password=k8ZO1e#qyDvX,7$D; Connection Timeout=60";


        public static string sApplicationVersion = "Version 2.03.23 from 2014-07-14";

        public static string sApplicationNameEmpty = string.Empty;
        public static string sApplicationNameTLD = "TLD";
        public static string sApplicationNameCalibration = "Calibration";
        public static string sApplicationNameAnnualReport = "AnnualReport";
        public static string sApplicationNameDosimetryAuditNetwork = "DosimetryAuditNetwork";
        public static string sApplicationNameDIRAC = "DIRAC";

        public static string sApplicationName = sApplicationNameTLD; //"TLD", "Calibration", "AnnualReport"

        public static string sApplicationStartupPath = Application.StartupPath;
        public static string sApplicationTemplatesFolder = "Templates\\";
        public static string sApplicationLabelsFolder = "Labels\\";
        public static string sApplicationTempPath = Path.GetTempPath() + "IAEA\\NAHU\\SSDL\\";
        public static string sAuditNetworkIAEA = "IAEA";
        public static string sAuditNetworkRPC = "RPC";

        public static string sStateStatusClean = "Clean";
        public static string sStateStatusDirty = "Dirty";

        public static string sExternalVariables = "ExternalVariables.ams";      

        // SSDL Templates
        public static string ContactInformationForm = sApplicationTemplatesFolder + "SSDLContactInformationForm.pdf";
        public static string AnnualReportTemplate = sApplicationTemplatesFolder + "SSDLAnnualReportTemplate.pdf";
        public static string SSDLSurveyTemplate = sApplicationTemplatesFolder + "SSDLSurvey.pdf";

        public static string EmailTemplateAnnualReport = sApplicationTemplatesFolder + "SSDLAnnualReport.txt";
        public static string EmailTemplateSSDLSurvey = sApplicationTemplatesFolder + "SSDLSurvey.txt";
        public static string EmailTemplateContactInformation = sApplicationTemplatesFolder + "SSDLContactInformationForm.txt";
        public static string EmailTemplateUserNotification = sApplicationTemplatesFolder + "UserNotification.txt";

        public static string EmailTemplateRequestSendPackege = sApplicationTemplatesFolder + "RequestSendPackage.txt";
        public static string EmailTemplateRequestGeneratePackage = sApplicationTemplatesFolder + "RequestGeneratePackage.txt";

        // TLD Templates
        public static string TLDOperatorReport = sApplicationTemplatesFolder + "TLDOperatorReport.dotx";
        public static string TLDCountryReport = sApplicationTemplatesFolder + "TLDCountryReport.dotx";
        
        public static string TLDApplicatioForm = sApplicationTemplatesFolder + "TLDApplicationForm.pdf";
        public static string TLDApplicatioFormRussian = sApplicationTemplatesFolder + "TLDApplicationFormRussian.pdf";
        public static string TLDApplicatioFormSpanish = sApplicationTemplatesFolder + "TLDApplicationFormSpanish.pdf";

        public static string TLDDataSheetTemplateRTSSDL = sApplicationTemplatesFolder + "TLDDataSheetTemplateRT-SSDL.pdf";
        // public static string TLDDataSheetTemplateRPSSDL = sApplicationTemplatesFolder + "TLDDataSheetTemplateRP-SSDL.pdf"; change done for Paulina April 6th, 2015
        public static string TLDDataSheetTemplateRPSSDL = sApplicationTemplatesFolder + "OSLDDataSheetTemplateRP-SSDL.pdf";

        public static string TLDDataSheetTemplateRTSSDLElectrones = sApplicationTemplatesFolder + "TLDDataSheetTemplateRT-SSDL-Electrones.pdf";

        public static string TLDDataSheetTemplateRTHospital = sApplicationTemplatesFolder + "TLDDataSheetTemplateRT-Hospital.pdf";
        public static string TLDDataSheetTemplateRTHospitalRussian = sApplicationTemplatesFolder + "TLDDataSheetTemplateRT-Hospital-Russian.pdf";
        public static string TLDDataSheetTemplateRTHospitalSpanish = sApplicationTemplatesFolder + "TLDDataSheetTemplateRT-Hospital-Spanish.pdf";

        public static string TLDDataSheetTemplateRTHospitalElectrones = sApplicationTemplatesFolder + "TLDDataSheetTemplateRT-Hospital-Electrones.pdf";


        public static string TLDInstructionSheetTemplateRTHospital = sApplicationTemplatesFolder + "TLDInstructionSheetTemplateRT-Hospital.pdf";
        public static string TLDInstructionSheetTemplateRTHospitalRussian = sApplicationTemplatesFolder + "TLDInstructionSheetTemplateRT-Hospital-Russian.pdf";
        public static string TLDInstructionSheetTemplateRTHospitalSpanish = sApplicationTemplatesFolder + "TLDInstructionSheetTemplateRT-Hospital-Spanish.pdf";

        public static string TLDInstructionSheetTemplateRTHospitalElectron = sApplicationTemplatesFolder + "TLDInstructionSheetTemplateRT-Hospital-Electron.pdf";

        public static string TLDInstructionSheetTemplateRTSSDL = sApplicationTemplatesFolder + "TLDInstructionSheetTemplateRT-SSDL.pdf";
        public static string TLDInstructionSheetTemplateRPSSDL = sApplicationTemplatesFolder + "TLDInstructionSheetTemplateRP-SSDL.pdf";

        public static string TLDDataSheetTemplateRPHospital = sApplicationTemplatesFolder + "TLDDataSheetTemplateRP-Hospital.pdf";

        public static string TLDEvaluationTemplateRT = sApplicationTemplatesFolder + "TLDEvaluationTemplateRT.xlsm";
        public static string TLDEvaluationTemplateRTFollowUp = sApplicationTemplatesFolder + "TLDEvaluationTemplateRT-FollowUp.xlsm";
        public static string TLDEvaluationTemplateRP = sApplicationTemplatesFolder + "TLDEvaluationTemplateRP.xlsm";
        public static string TLDEvaluationTemplateRPFollowUp = sApplicationTemplatesFolder + "TLDEvaluationTemplateRP-FollowUp.xlsm";

        public static string TLDCertificateTemplateRTHospital = sApplicationTemplatesFolder + "TLDCertificateTemplateRT-Hospital.dotx";
        public static string TLDCertificateTemplateRTHospital2 = sApplicationTemplatesFolder + "TLDCertificateTemplateRT-Hospital2.dotx";
        public static string TLDCertificateTemplateRTHospitalElectron = sApplicationTemplatesFolder + "TLDCertificateTemplateRT-Hospital-Electron.dotx";        

        public static string TLDCertificateTemplateRTHospitalSpanish = sApplicationTemplatesFolder + "TLDCertificateTemplateRT-HospitalSpanish.dotx";
        public static string TLDCertificateTemplateRTHospitalSpanish2 = sApplicationTemplatesFolder + "TLDCertificateTemplateRT-HospitalSpanish2.dotx";
        public static string TLDCertificateTemplateRTHospitalRussian = sApplicationTemplatesFolder + "TLDCertificateTemplateRT-HospitalRussian.dotx";
        public static string TLDCertificateTemplateRTHospitalRussian2 = sApplicationTemplatesFolder + "TLDCertificateTemplateRT-HospitalRussian2.dotx";

        public static string TLDCertificateTemplateRTSSDL = sApplicationTemplatesFolder + "TLDCertificateTemplateRT-SSDL.dotx";
        public static string TLDCertificateTemplateRTSSDL2 = sApplicationTemplatesFolder + "TLDCertificateTemplateRT-SSDL2.dotx";
        public static string TLDCertificateTemplateRTPrimary = sApplicationTemplatesFolder + "TLDCertificateTemplateRT-Primary.dotx";
        public static string TLDCertificateTemplateRTReference = sApplicationTemplatesFolder + "TLDCertificateTemplateRT-Reference.dotx";
        public static string TLDCertificateTemplateRPSSDL = sApplicationTemplatesFolder + "TLDCertificateTemplateRP-SSDL.dotx";
        public static string TLDCertificateTemplateRPSSDL2 = sApplicationTemplatesFolder + "TLDCertificateTemplateRP-SSDL2.dotx";
        public static string TLDCertificateTemplateRPPrimary = sApplicationTemplatesFolder + "TLDCertificateTemplateRP-Primary.dotx";

        public static string SSDLCalibrationCoveringLetterTemplateInternal = sApplicationTemplatesFolder + "SSDLCalibrationCoveringLetterTemplate-Internal.dotx";
        public static string SSDLCalibrationCoveringLetterTemplateExternal = sApplicationTemplatesFolder + "SSDLCalibrationCoveringLetterTemplate-External.dotx";

        public static string TLDCoveringLetterTemplateRTSSDL = sApplicationTemplatesFolder + "TLDCoveringLetterTemplateRT-SSDL.dotx";
        public static string TLDCoveringLetterTemplateRTSSDLPrimary = sApplicationTemplatesFolder + "TLDCoveringLetterTemplateRT-SSDL-Primary.dotx";
        public static string TLDCoveringLetterTemplateRPSSDL = sApplicationTemplatesFolder + "TLDCoveringLetterTemplateRP-SSDL.dotx";
        public static string TLDCoveringLetterTemplateRPSSDLPrimary = sApplicationTemplatesFolder + "TLDCoveringLetterTemplateRP-SSDL-Primary.dotx";
        public static string TLDCoveringLetterTemplateRTHospital = sApplicationTemplatesFolder + "TLDCoveringLetterTemplateRT-Hospital.dotx";
        public static string TLDCoveringLetterTemplateRTHospitalPrimary = sApplicationTemplatesFolder + "TLDCoveringLetterTemplateRT-Hospital-Primary.dotx";
        public static string TLDCoveringLetterTemplateRTHospitalFollowUp = sApplicationTemplatesFolder + "TLDCoveringLetterTemplateRT-Hospital-FolowUp.dotx";

        public static string TLDCoveringLetterDispatchTLDTemplateRTHospital = sApplicationTemplatesFolder + "TLDCoveringLetterDispatchTLDTemplateRT-Hospital.dotx";
        public static string TLDCoveringLetterDispatchTLDTemplateRTHospitalFolowUp = sApplicationTemplatesFolder + "TLDCoveringLetterDispatchTLDTemplateRT-Hospital-FolowUp.dotx";
        public static string TLDCoveringLetterDispatchTLDTemplateRTCountryCoordinator = sApplicationTemplatesFolder + "TLDCoveringLetterDispatchTLDTemplateRT-CountryCoordinator.dotx";
        public static string TLDCoveringLetterDispatchTLDTemplateRTCountryCoordinatorCC = sApplicationTemplatesFolder + "TLDCoveringLetterDispatchTLDTemplateRT-CountryCoordinatorCC.dotx";
        public static string TLDCoveringLetterDispatchTLDTemplateRTPAHO = sApplicationTemplatesFolder + "TLDCoveringLetterDispatchTLDTemplateRT-PAHO.dotx";
        public static string TLDCoveringLetterDispatchTLDTemplateRTPrimary = sApplicationTemplatesFolder + "TLDCoveringLetterDispatchTLDTemplateRT-Primary.dotx";

        public static string TLDCoveringLetterDispatchTLDTemplateRTSSDL = sApplicationTemplatesFolder + "TLDCoveringLetterDispatchTLDTemplateRT-SSDL.dotx";
        public static string TLDCoveringLetterDispatchTLDTemplateRTSSDLFolowUp = sApplicationTemplatesFolder + "TLDCoveringLetterDispatchTLDTemplateRT-SSDL-FolowUp.dotx";

        public static string TLDCoveringLetterProformaInvoiceHospital = sApplicationTemplatesFolder + "TLDCoveringLetterProformaInvoice-Hospital.dotx";
        public static string TLDCoveringLetterProformaInvoiceCountryCoordinator = sApplicationTemplatesFolder + "TLDCoveringLetterProformaInvoice-CountryCoordinator.dotx";
        public static string TLDCoveringLetterProformaInvoicePAHO = sApplicationTemplatesFolder + "TLDCoveringLetterProformaInvoice-PAHO.dotx";


        public static string TLDPrinciplesOfOperation = sApplicationTemplatesFolder + "Principles of Operation (English).pdf";
        public static string TLDPrinciplesOfOperationSpanish = sApplicationTemplatesFolder + "Principles of Operation (Spanish, Principios de Funcionamiento).pdf";
        public static string TLDPrinciplesOfOperationRussian = sApplicationTemplatesFolder + "Principles of Operation (Russian, Принципы предоставления).pdf";
        

        // -- e-Mails
        // Send Application Forms        
        public static string TLDEmailTemplateApplicationFormRTHospital = sApplicationTemplatesFolder + "TLDEmailTemplateApplicationFormRT-Hospital.txt";
        public static string TLDEmailTemplateApplicationFormRTHospitalCountryCoordinator = sApplicationTemplatesFolder + "TLDEmailTemplateApplicationFormRT-CountryCoordinator.txt";
        public static string TLDEmailTemplateApplicationFormRTSSDL = sApplicationTemplatesFolder + "TLDEmailTemplateApplicationFormRT-SSDL.txt";

        public static string TLDEmailTemplateApplicationFormRPSSDL = sApplicationTemplatesFolder + "TLDEmailTemplateApplicationFormRP-SSDL.txt";
        public static string TLDEmailTemplateApplicationFormRPHospital = sApplicationTemplatesFolder + "TLDEmailTemplateApplicationFormRP-Hospital.txt";

        // Send Reminder for Application Forms
        public static string TLDEmailTemplateApplicationFormRTHospitalReminder = sApplicationTemplatesFolder + "TLDEmailTemplateApplicationFormRT-Hospital-Reminder.txt";
        public static string TLDEmailTemplateApplicationFormRTCountryCoordinatorReminder = sApplicationTemplatesFolder + "TLDEmailTemplateApplicationFormRT-CountryCoordinator-Reminder.txt";
        public static string TLDEmailTemplateApplicationFormRTSSDLReminder = sApplicationTemplatesFolder + "TLDEmailTemplateApplicationFormRT-SSDL-Reminder.txt";

        // Send Data Sheets
        public static string TLDEmailTemplateSendDataSheetRTHospital = sApplicationTemplatesFolder + "TLDEmailTemplateSendDataSheetRT-Hospital.txt";
        public static string TLDEmailTemplateSendDataSheetRTHospitalCountryCoordinator = sApplicationTemplatesFolder + "TLDEmailTemplateSendDataSheetRT-HospitalCountryCoordinator.txt";
        public static string TLDEmailTemplateSendDataSheetRTHospitalCountryCoordinatorCC = sApplicationTemplatesFolder + "TLDEmailTemplateSendDataSheetRT-HospitalCountryCoordinatorCC.txt";
        public static string TLDEmailTemplateSendDataSheetRTHospitalPAHOEnglish = sApplicationTemplatesFolder + "TLDEmailTemplateSendDataSheetRT-HospitalPAHOEnglish.txt";
        public static string TLDEmailTemplateSendDataSheetRTHospitalPAHOSpanish = sApplicationTemplatesFolder + "TLDEmailTemplateSendDataSheetRT-HospitalPAHOSpanish.txt";
        public static string TLDEmailTemplateSendDataSheetRTSSDL = sApplicationTemplatesFolder + "TLDEmailTemplateSendDataSheetRT-SSDL.txt";
        public static string TLDEmailTemplateSendDataSheetRTReference = sApplicationTemplatesFolder + "TLDEmailTemplateSendDataSheetRT-Reference.txt";
        public static string TLDEmailTemplateSendDataSheetRTPrimary = sApplicationTemplatesFolder + "TLDEmailTemplateSendDataSheetRT-Primary.txt";

        // Send Data Sheets Follow-Up
        public static string TLDEmailTemplateSendDataSheetRTHospitalFollowUp = sApplicationTemplatesFolder + "TLDEmailTemplateSendDataSheetRT-Hospital-FollowUp.txt";
        public static string TLDEmailTemplateSendDataSheetRTHospitalCountryCoordinatorFollowUp = sApplicationTemplatesFolder + "TLDEmailTemplateSendDataSheetRT-HospitalCountryCoordinator-FollowUp.txt";
        public static string TLDEmailTemplateSendDataSheetRTHospitalCountryCoordinatorCCFollowUp = sApplicationTemplatesFolder + "TLDEmailTemplateSendDataSheetRT-HospitalCountryCoordinatorCC-FollowUp.txt";
        public static string TLDEmailTemplateSendDataSheetRTHospitalPAHOEnglishFollowUp = sApplicationTemplatesFolder + "TLDEmailTemplateSendDataSheetRT-HospitalPAHOEnglish-FollowUp.txt";
        public static string TLDEmailTemplateSendDataSheetRTHospitalPAHOSpanishFollowUp = sApplicationTemplatesFolder + "TLDEmailTemplateSendDataSheetRT-HospitalPAHOSpanish-FollowUp.txt";
        public static string TLDEmailTemplateSendDataSheetRTSSDLFollowUp = sApplicationTemplatesFolder + "TLDEmailTemplateSendDataSheetRT-SSDL-FollowUp.txt";

        public static string TLDEmailTemplateSendDataSheetRPSSDL = sApplicationTemplatesFolder + "TLDEmailTemplateSendDataSheetRP-SSDL.txt";
        public static string TLDEmailTemplateSendDataSheetRPReference = sApplicationTemplatesFolder + "TLDEmailTemplateSendDataSheetRP-Reference.txt";
        public static string TLDEmailTemplateSendDataSheetRPPrimary = sApplicationTemplatesFolder + "TLDEmailTemplateSendDataSheetRP-Primary.txt";

        // Send Reminder for Data Sheets
        public static string TLDEmailTemplateSendDataSheetRPReminder = sApplicationTemplatesFolder + "TLDEmailTemplateSendDataSheetRP-Reminder.txt";
        public static string TLDEmailTemplateSendDataSheetRTReminder = sApplicationTemplatesFolder + "TLDEmailTemplateSendDataSheetRT-Reminder.txt";

        // Send Dispatch / Certificates
        public static string TLDEmailTemplateDispatchCertificateRTHospital = sApplicationTemplatesFolder + "TLDEmailTemplateDispatchCertificateRT-Hospital.txt";
        public static string TLDEmailTemplateDispatchCertificateRTHospitalFollowUp = sApplicationTemplatesFolder + "TLDEmailTemplateDispatchCertificateRT-Hospital-FollowUp.txt";
        public static string TLDEmailTemplateDispatchCertificateRTSSDL = sApplicationTemplatesFolder + "TLDEmailTemplateDispatchCertificateRT-SSDL.txt";
        public static string TLDEmailTemplateDispatchCertificateRTSSDLFollowUp = sApplicationTemplatesFolder + "TLDEmailTemplateDispatchCertificateRT-SSDL-FollowUp.txt";
        public static string TLDEmailTemplateDispatchCertificateRTPrimary = sApplicationTemplatesFolder + "TLDEmailTemplateDispatchCertificateRT-Primary.txt";

        // Send Dispatch / Certificates Follow-Up
        public static string TLDEmailTemplateDispatchCertificateRPSSDL = sApplicationTemplatesFolder + "TLDEmailTemplateDispatchCertificateRP-SSDL.txt";
        public static string TLDEmailTemplateDispatchCertificateRPSSDLFollowUp = sApplicationTemplatesFolder + "TLDEmailTemplateDispatchCertificateRP-SSDL-FollowUp.txt";
        public static string TLDEmailTemplateDispatchCertificateRPPrimary = sApplicationTemplatesFolder + "TLDEmailTemplateDispatchCertificateRP-Primary.txt";

        // HTML Graph Templates
        public static string TLDGraphTreeRadial = sApplicationTemplatesFolder + "TLDGraphTreeRadial.html";        
        public static string TLDGraphSunbursts = sApplicationTemplatesFolder + "TLDGraphSunbursts.html";        
        public static string TLDGraphReingoldTilfordTree = sApplicationTemplatesFolder + "TLDGraphReingoldTilfordTree.html";
        public static string TLDGraphCollapsibleTreeLayout = sApplicationTemplatesFolder + "TLDGraphCollapsibleTreeLayout.html";
        public static string TLDGraphCirclePacking = sApplicationTemplatesFolder + "TLDGraphCirclePacking.html";
        public static string TLDGraphCodeFlower = sApplicationTemplatesFolder + "TLDGraphCodeFlower.html";
        

        public static string VisualDataYearBatchCountry = "VisualDataYearBatchCountry";
        public static string VisualDataBatchCountrySet = "VisualDataBatchCountrySet";
        public static string VisualDataCountryYearBatchSet = "VisualDataCountryYearBatchSet";

        public static string ImporterVersion_20081107 = "20081107";
        public static string ImporterVersion_Current = ImporterVersion_20081107;
        
        public static int SSDL_Admin_UserID = 100;
        public static char[] delimiterChars = { '|' };

        public static string sTaskDispatchRequest = "TaskDispatchRequest";
        //public static string sTaskDispatchChamber = "TaskDispatchChamber";
        //public static string sTaskDispatchElectrometer = "TaskDispatchElectrometer";        
        public static string sTaskDispatchEquipment = "TaskDispatchEquipment";
        public static string sTaskDispatchIAEAEquipment = "TaskDispatchIAEAEquipment";
        
        public static string sTaskValidateRequest = "TaskValidateRequest";
        public static string sTaskCreateCalibration = "TaskCreateCalibration";
        public static string sTaskVerifyRequest = "TaskVerifyRequest";
        public static string sTaskGeneratePackage = "TaskGeneratePackage";
        

        public static string sTaskReviewSummary = "TaskReviewSummary";
        public static string sTaskClearSummary = "TaskClearSummary"; //Pending Summaries (Waiting for SH Signature)
        public static string sTaskSignRequest = "TaskSignRequest"; // Sign Summary and Certificates by SH

        public static string sTaskChamberSummary = "TaskChamberSummary";
        public static string sTaskRequestSummary = "TaskRequestSummary";

        public static string sTaskSignCertificate = "TaskSignCertificate";

        public static string sTaskCalibrationListToday = "TaskCalibrationListToday";
        public static string sTaskCalibrationListDate = "TaskCalibrationListDate";
        public static string sTaskCalibrationListMonth = "TaskCalibrationListMonth";
        public static string sTaskCalibrationListYear = "TaskCalibrationListYear";


        public static string sTaskRequestList = "TaskRequestList";
        public static string sTaskEquipmentList = "TaskEquipmentList";

        public static string sTaskChambersSearch = "TaskChambersSearch";

        public static string sTaskElectrometersInDOLMembers = "TaskElectrometersInDOLMembers";
        public static string sTaskElectrometersInDOLIAEA = "TaskElectrometersInDOLIAEA";
        public static string sTaskElectrometersSearch = "TaskElectrometersSearch";

        public static string sTaskSignTLDCertificate = "TaskSignTLDCertificate";
        public static string sTaskIsueTLDCertificate = "TaskIsueTLDCertificate";
        public static string sTaskDispatchTLDCertificate = "TaskDispatchTLDCertificate";

        public static string sTaskOnlyCustomStatistics = "TaskOnlyCustomStatistics";

        public static string sTaskBatchSetUp = "TaskBatchSetUp";
        public static string sTaskBatchPlanning = "TaskBatchPlanning";
        public static string sTaskBatchDefinning = "TaskBatchDefinning";
        public static string sTaskBatchDispatching = "TaskBatchDispatching";
        public static string sTaskBatchInfoGraphics = "TaskBatchInfoGraphics";
        
        //public static string sTaskBatchArchiving = "TaskBatchArchiving";

        public static string sTaskInvitedDataSheets = "InvitedDataSheets";   // CreatedOn IS NOT NULL 
        public static string sTaskPendingDataSheets = "PendingDataSheets";   // CreatedOn IS NOT NULL and IrradiationDate IS NULL     and EvaluationID IS NULL     and CertificateID IS NULL 
        public static string sTaskRecivedDataSheets = "RecivedDataSheets";   // CreatedOn IS NOT NULL and IrradiationDate IS NOT NULL and EvaluationID IS NULL     and CertificateID IS NULL 
        public static string sTaskPendingEvaluations = "PendingEvaluations"; // CreatedOn IS NOT NULL                                 and EvaluationID IS NOT NULL and CertificateID IS NULL 
        public static string sTaskPendingTLDCertificates = "PendingTLDCertificates"; // CreatedOn IS NOT NULL and EvaluationID IS NOT NULL and CertificateID IS NOT NULL and SignByOfficerOn IS NULL and SignBySectionHeadOn IS NULL and DispatchedOn IS NULL
        public static string sTaskPendingTLDCertificatesSHSignature = "PendingTLDCertificatesSHSignature"; // CreatedOn IS NOT NULL and EvaluationID IS NOT NULL and CertificateID IS NOT NULL and SignByOfficerOn IS NOT NULL and SignBySectionHeadOn IS NULL and DispatchedOn IS NULL
        public static string sTaskDispatchTLDCertificates = "DispatchTLDCertificates"; // CreatedOn IS NOT NULL and EvaluationID IS NOT NULL and CertificateID IS NOT NULL and SignByOfficerOn IS NOT NULL and SignBySectionHeadOn IS NOT NULL and DispatchedOn IS NULL
        public static string sTaskArchiveTLDCertificates = "ArchiveTLDCertificates";
        public static string sTaskCreateTLDFollowUp = "CreateTLDFollowUp";


        public static string sDataColumnListDefault = "Default";
        public static string sDataColumnListBatchPlanning = "BatchPlanning";
        public static string sDataColumnListPendingDataSheets = "PendingDataSheets";
        public static string sDataColumnListPendingEvaluations = "PendingEvaluations";
        public static string sDataColumnListSignCertificate = "SignCertificate";
        public static string sDataColumnListBatchDipatching = "BatchDipatching";
        public static string sDataColumnListBatchArchiving = "BatchArchiving";
        public static string sDataColumnListSimpleDataList = "SimpleDataList";
        public static string sDataColumnListPackageDefault = "PackageDefault";
        public static string sDataColumnListPackageBatchPlanning = "PackageBatchPlanning";
        public static string sDataColumnListPackageBatchDefinning = "PackageBatchDefinning";        
        public static string sDataColumnListPackageBatchDipatching = "PackageBatchDipatching";
        public static string sDataColumnListPendingRequestsList = "PendingRequestsList";

        public static int iSignatureCalibrationCreated = 7;
        public static int iSignatureSummaryCreated = 8;
        public static int iSignatureSummaryChecked = 9;
        public static int iSignatureSummaryCleared = 10;
        public static int iSignatureCertificateCreated = 11; // Sign By Section Head
        public static int iSignatureCertificateIssued = 12; // Sign By Section Head
        //public static int iSignatureCertificateSend = 13;

        public static int iSignatureRequestCreated = 14;
        public static int iSignatureRequestValidated = 15;
        public static int iSignatureRequestCleared = 18; // Sign By Section Head
        public static int iSignaturePackageVerified = 16;
        public static int iSignatureRequestDispatched = 17;
        public static int iSignatureRequestCanceled = 19;

        public static int iSignatureTLDCertificateSignByOfficer = 41;
        public static int iSignatureTLDCertificateSignBySectionHead = 42; // Sign By Section Head
        public static int iSignatureTLDCertificateDispatched = 43;
        public static int iSignatureTLDCertificateArchived = 44;


        public static int iAttachmentRequestForm = 1;
        public static int iAttachmentRequestConfirmation = 2;
        public static int iAttachmentEquipmentReceivedConfirmation = 3;
        public static int iAttachmentPackageGenerated = 4;
        public static int iAttachmentRequestDispatchedConfirmation = 5;
        public static int iAttachmentEquipmentDispatchedConfirmation = 6;
        public static int iAttachmentSummaRyreadyForVerification = 7;
        public static int iAttachmentCalibrationCertificate = 8;
        public static int iAttachmentRequestCoverLetterInternal = 9;
        public static int iAttachmentRequestCoverLetterExternal = 19;
        public static int iAttachmentRequestAppendix = 10;

        public static int iAttachmentAnnualReport = 11;

        public static int iAttachmentApplicationForm = 20;
        public static int iAttachmentDataSheet = 21;
        public static int iAttachmentEvaluationReadings = 22;
        //public static int iAttachmentEvaluationSheetOriginal = 23;
        public static int iAttachmentEvaluationSheetVeryfied = 24;
        public static int iAttachmentTLDCertificate = 25;
        //public static int iAttachmentTLDCertificateCoverLetter = 26;
        public static int iAttachmentOther = 1000;
        

        public static string sTemplateObjectOperator = "Operator";
        public static string sTemplateObjectAnnualReport = "AnnualReport";
        public static string sTemplateObjectRequest = "Request";
        public static string sTemplateObjectSummary = "Summary";
        public static string sTemplateObjectCalibration = "Calibration";
        public static string sTemplateObjectCertificate = "Certificate";

        public static string sFormatPDF = "FormatPDF";
        public static string sFormatWord97 = "FormatWord97";
        public static string sFormatWord2010 = "FormatWord2010";
        
        public static string sSummaryDetailsCalibrationChValue = "CalibrationChValue";
        public static string sSummaryDetailsCalibrationSysValue = "CalibrationSysValue";
        public static string sSummaryDetailsKermaDoseValue = "KermaDoseValue";
        public static string sSummaryDetailsSourceChValue = "SourceChValue";
        public static string sSummaryDetailsSourceChUncertCoef = "SourceChUncertCoef";
        public static string sSummaryDetailsSourceSysValue = "SourceSysValue";
        public static string sSummaryDetailsSourceSysUncertCoef = "SourceSysUncertCoef";
        public static string sSummaryDetailsBeamUncertCoef = "BeamUncertCoef"; //8
        public static string sSummaryDetailsBeamUncertCoefa = "BeamUncertCoefa"; //8

        public static string sStatusCreated = "Created";  //Request
        public static string sStatusValidated = "Validated"; //Request
        public static string sStatusReceivedConfirmation = "Request Received Confirmation"; //Request
        public static string sStatusVerified = "Verified"; //Request
        public static string sStatusChecked = "Checked"; //Summary        
        public static string sStatusCleared = "Cleared"; //Summary
        public static string sStatusIssued = "Issued"; //Certificate
        public static string sStatusDispatched = "Dispatched"; //Request
        public static string sStatusDispatchedConfirmation = "Dispatched Confirmation"; //Request
        
        
        public static string sStatusCanceled = "Canceled"; //Request

        public static string sStatusEquipmentReceived = "Equipment Received"; //Request Equipment
        public static string sStatusEquipmentReceivedConfirmation = "Equipment Received Confirmation"; //Request Equipment
        public static string sStatusEquipmentDispatched = "Equipment Dispatched"; //Request Equipment
        public static string sStatusEquipmentDispatchedConfirmation = "Equipment Dispatched Confirmation"; //Request Equipment
        
        public static string sAuditTypeRT = "RT";
        public static string sAuditTypeRP = "RP";
        public static string sAuditTypeCalibration = "Calibration";
        public static string sAuditTypeDosimeteryAuditNetwork = "DosimeteryAuditNetwork";
        public static string sAuditTypeDIRAC = "DIRAC";

        public static string sParticipationTypeSSDL = "SSDL";
        public static string sParticipationTypeHospitals = "Hospitals";
        public static string sParticipationTypePrimary = "Primary";
        public static string sParticipationTypeReference = "Reference";
        //public static string sParticipationTypeDIRAC = "DIRAC";

        public static string sCommunicationLanguageEnglish = "English";
        public static string sCommunicationLanguageSpanish = "Spanish";
        public static string sCommunicationLanguageRussian = "Russian";

        public static string sCertificatetDosesDifferentDoses = "DifferentDoses";
        public static string sCertificatetDosesSameDoses = "SameDoses";

        public static string sExportStatusNew = "New";
        public static string sExportStatusUpdated = "Updated";
        public static string sExportStatusExported = "Exported";

        public static string sUnitTypeGroupRadionuclideTherapy = "RadionuclideTherapy";
        public static string sUnitTypeGroupLinearAccelerator = "LinearAccelerator";
        public static string sUnitTypeGroupCircularAccelerator = "CircularAccelerator";
        public static string sUnitTypeGroupXRayGenerator = "XRayGenerator";
        public static string sUnitTypeGroupParticleTherapy = "ParticleTherapy";
        public static string sUnitTypeGroupBrachyTherapy = "BrachyTherapy";
        public static string sUnitTypeGroupRPIrradiator = "RPIrradiator";

        public static string sUnitTypeGroupRT = sUnitTypeGroupRadionuclideTherapy + "," + sUnitTypeGroupLinearAccelerator + "," + sUnitTypeGroupCircularAccelerator + "," + sUnitTypeGroupXRayGenerator + "," + sUnitTypeGroupParticleTherapy;
        public static string sUnitTypeGroupDIRAC = sUnitTypeGroupRadionuclideTherapy + "," + sUnitTypeGroupLinearAccelerator + "," + sUnitTypeGroupCircularAccelerator + "," + sUnitTypeGroupXRayGenerator + "," + sUnitTypeGroupParticleTherapy + "," + sUnitTypeGroupBrachyTherapy;
        public static string sUnitTypeGroupRP = sUnitTypeGroupRPIrradiator;
        public static string sUnitTypeGroupBrachy = sUnitTypeGroupBrachyTherapy;

        public static string sUnitTypeCo60 = "Co60"; // Co-60 (Radionuclide Teletherapy)
        public static string sUnitTypeCo60SRT = "Co60SRT"; // Co-60 (Stereotactic Teletherapy)
        public static string sUnitTypeRadionuclideCs137 = "RadionuclideCs137"; // Cs-137 (Radionuclide Teletherapy)        

        public static string sUnitTypeAccelerator = "Accelerator"; //Linac (Clinical Accelerator)
        public static string sUnitTypeTomotherapy = "Tomotherapy"; //Tomotherapy (Linear Accelerator)
        public static string sUnitTypeLinacRobotic = "LinacRobotic"; //Robotic arm (Linear Accelerator)
        public static string sUnitTypeLinacIORT = "LinacIORT"; //IORT (Linear Accelerator)

        public static string sUnitTypeBetatron = "Betatron"; //Betatron (Circular Accelerator)
        public static string sUnitTypeMicrotron = "Microtron"; //Microtron (Circular Accelerator)

        public static string sUnitTypeXRayGenerator = "XRayGenerator"; // Only for DIRAC X-Ray Generator
        public static string sUnitTypeElectronicBrachytherapy = "ElectronicBrachytherapy"; // Only for DIRAC Electronic Brachytherapy

        public static string sUnitTypeSynchrotron = "Synchrotron"; //Proton: Synchrotron // Only for DIRAC
        public static string sUnitTypeCyclotron = "Cyclotron"; //Proton: Cyclotron // Only for DIRAC
        public static string sUnitTypeSynchroCyclotron = "SynchroCyclotron"; //Proton: SynchroCyclotron // Only for DIRAC

        public static string sUnitTypeCs137Irradiator = "Cs137Irradiator"; // Cs-137 Irradiator
        public static string sUnitTypeCo60Irradiator = "Co60Irradiator"; // Co-60 Irradiator
        public static string sUnitTypeOther = "Other"; // Only for DIRAC






        public static string sBeamTypeCo60 = "Co60";
        public static string sBeamTypeCs137 = "Cs137";
        public static string sBeamTypePhoton = "Photon";
        public static string sBeamTypeElectron = "Electron";
        public static string sBeamTypeXRayGenerator = "XRayGenerator";
        public static string sBeamTypeProton = "Proton";
        public static string sBeamTypeCarbonIon = "CarbonIon";
        public static string sBeamTypeNeutron = "Neutron";
        public static string sBeamTypeBrachy = "Brachy";
        

        public static string sEditModeNone = "None"; // Edit, None
        public static string sEditModeEdit = "Edit"; // Edit, None

        public static string sViewStyleInstitutions = "Institutions"; // Institutions, Units, Beams
        public static string sViewStyleUnits = "Units"; // Institutions, Units, Beams
        public static string sViewStyleBeams = "Beams"; // Institutions, Units, Beams

        public static string ReportYear = "2013"; // Read value from tblSettings
        public static int MaxVariableCount = 250;
        public static int TLDActiveYear = 2000;

        //Font iReportFontCurrent;
        public static iTextSharp.text.Font iReportFont = FontFactory.GetFont(FontFactory.COURIER, 10, iTextSharp.text.Font.COURIER);
        public static iTextSharp.text.Font iReportFont10 = FontFactory.GetFont(FontFactory.HELVETICA, 10, iTextSharp.text.Font.HELVETICA, new iTextSharp.text.Color(System.Drawing.Color.Black));
        public static iTextSharp.text.Font iReportFont10Blue = FontFactory.GetFont(FontFactory.HELVETICA, 10, iTextSharp.text.Font.HELVETICA, new iTextSharp.text.Color(System.Drawing.Color.Blue));
        public static iTextSharp.text.Font iReportFont12 = FontFactory.GetFont(FontFactory.HELVETICA, 12, iTextSharp.text.Font.HELVETICA, new iTextSharp.text.Color(System.Drawing.Color.Black));
        public static iTextSharp.text.Font iReportFont14 = FontFactory.GetFont(FontFactory.HELVETICA, 14, iTextSharp.text.Font.HELVETICA, new iTextSharp.text.Color(System.Drawing.Color.Black));
        public static iTextSharp.text.Font iReportFont14Blue = FontFactory.GetFont(FontFactory.HELVETICA, 14, iTextSharp.text.Font.HELVETICA, new iTextSharp.text.Color(System.Drawing.Color.Blue));
        public static iTextSharp.text.Font iReportFont8 = FontFactory.GetFont(FontFactory.HELVETICA, 8, iTextSharp.text.Font.HELVETICA, new iTextSharp.text.Color(System.Drawing.Color.Black));
        public static iTextSharp.text.Font iReportFont8Red = FontFactory.GetFont(FontFactory.HELVETICA, 8, iTextSharp.text.Font.HELVETICA, new iTextSharp.text.Color(System.Drawing.Color.Red));
        public static iTextSharp.text.Font iReportFont8Blue = FontFactory.GetFont(FontFactory.HELVETICA, 8, iTextSharp.text.Font.HELVETICA, new iTextSharp.text.Color(System.Drawing.Color.Blue));
        public static iTextSharp.text.Font iReportFont6 = FontFactory.GetFont(FontFactory.HELVETICA, 6, iTextSharp.text.Font.HELVETICA, new iTextSharp.text.Color(System.Drawing.Color.Black));
        public static iTextSharp.text.Font iReportFont6Red = FontFactory.GetFont(FontFactory.HELVETICA, 6, iTextSharp.text.Font.HELVETICA, new iTextSharp.text.Color(System.Drawing.Color.Red));
        public static iTextSharp.text.Font iReportFont6Blue = FontFactory.GetFont(FontFactory.HELVETICA, 6, iTextSharp.text.Font.HELVETICA, new iTextSharp.text.Color(System.Drawing.Color.Blue));
        

        public static ManagerClass Manager = new ManagerClass();
        public static GlobalDictionary Dictionary = new GlobalDictionary();
        public static UserClass User = new UserClass(null);
        public static DateTime GlobalStartTime = DateTime.Now;

        //event handler for StateChanged event
        protected static void OnStateChanged(object sender, StateChangeEventArgs e)
        {
            Console.WriteLine(
                "sql server state is -> {0}",
                e.CurrentState.ToString("g")
            );
        }

        //event handler for InfoMessage event
        protected static void OnInfoMessage(object sender, SqlInfoMessageEventArgs e)
        {
            foreach (SqlError err in e.Errors)
            {
                Console.WriteLine(err.Message);
            }
        }

        public static string[] Wrap(string text, int maxLength)
        {
            text = text.Replace("\n", " ");
            text = text.Replace("\r", " ");
            text = text.Replace(".", ". ");
            text = text.Replace(">", "> ");
            text = text.Replace("\t", " ");
            text = text.Replace(",", ", ");
            text = text.Replace(";", "; ");
            text = text.Replace("<br>", " ");
            text = text.Replace(" ", " ");

            string[] Words = text.Split(' ');
            int currentLineLength = 0;
            ArrayList Lines = new ArrayList(text.Length / maxLength);
            string currentLine = "";
            bool InTag = false;

            foreach (string currentWord in Words)
            {
                //ignore html
                if (currentWord.Length > 0)
                {

                    if (currentWord.Substring(0, 1) == "<")
                        InTag = true;

                    if (InTag)
                    {
                        //handle filenames inside html tags
                        if (currentLine.EndsWith("."))
                        {
                            currentLine += currentWord;
                        }
                        else
                            currentLine += " " + currentWord;

                        if (currentWord.IndexOf(">") > -1)
                            InTag = false;
                    }
                    else
                    {
                        if (currentLineLength + currentWord.Length + 1 < maxLength)
                        {
                            currentLine += " " + currentWord;
                            currentLineLength += (currentWord.Length + 1);
                        }
                        else
                        {
                            Lines.Add(currentLine);
                            currentLine = currentWord;
                            currentLineLength = currentWord.Length;
                        }
                    }
                }
            }
            if (currentLine != "")
                Lines.Add(currentLine);

            string[] textLinesStr = new string[Lines.Count];
            Lines.CopyTo(textLinesStr, 0);
            return textLinesStr;
        }

        public static string FormatStringValue(string inString, int iLength)
        {
            string sReturn = inString.Replace("'", "''");

            if (sReturn == "Off")
                sReturn = string.Empty;

            if (sReturn.Length > iLength)
                sReturn = sReturn.Substring(0, iLength);
            
            return sReturn;
        }

        public static string FormatIntegerValue(int inValue)
        {
            string sReturn = "NULL";

            if (inValue != -1)
                sReturn = inValue.ToString();

            return sReturn;
        }

        public static Double FormatStringDoubleValue(string inValue)
        {
            try
            {
                return Double.Parse(inValue);
            }
            catch
            {
                return 0.00;
            }
        }

        public static string FormatDateTimeValue(DateTime inDateTime)
        {
            string sReturn = string.Empty;

            if (inDateTime != DateTime.MinValue)
                sReturn = inDateTime.ToString().Substring(0, 10).Replace("0001-01-01", "");

            return sReturn;
        }

        public static DateTime FormatIAEAStringDateTimeValue(string sDateTime)
        {
            // {yyyy-mm-dd}
            try
            {
                int iYear = Convert.ToInt32(sDateTime.Trim().Substring(0, 4));
                int iMonth = Convert.ToInt32(sDateTime.Trim().Substring(5, 2));
                int iDay = Convert.ToInt32(sDateTime.Trim().Substring(8, 2));

                if (iYear >= 1800)
                    return new DateTime(iYear, iMonth, iDay);
                else
                    return DateTime.MinValue;
                //return DateTime.Parse(inDateTime);
            }
            catch
            {
                return DateTime.MinValue;
            }
        }

        public static DateTime FormatTLDStringDateTimeValue(string sDateTime)
        {
            // {dd/mm/yyyy}
            try
            {
                int iYear = Convert.ToInt32(sDateTime.Trim().Substring(6, 4));
                int iMonth = Convert.ToInt32(sDateTime.Trim().Substring(3, 2));
                int iDay = Convert.ToInt32(sDateTime.Trim().Substring(0, 2));

                if (iYear >= 1800)
                    return new DateTime(iYear, iMonth, iDay);
                else
                    return DateTime.MinValue;
                //return DateTime.Parse(inDateTime);
            }
            catch
            {
                return DateTime.MinValue;
            }
        }

        public static string FormatTLDStringDateTimeValue(DateTime dDateTime)
        {
            // {dd/mm/yyyy}
            try
            {
                int iYear = dDateTime.Year;
                int iMonth = dDateTime.Month;
                int iDay = dDateTime.Day;

                return iDay.ToString().PadLeft(2, '0') + "/" + iMonth.ToString().PadLeft(2, '0') + "/" + iYear.ToString();
                //return DateTime.Parse(inDateTime);
            }
            catch
            {
                return string.Empty;
            }
        }

        public static DateTime ConvertToDateTime(double excelDate)
        {
            if (excelDate < 1)
            {
                throw new ArgumentException("Excel dates cannot be smaller than 0.");
            }
            DateTime dateOfReference = new DateTime(1900, 1, 1);
            if (excelDate > 60d)
            {
                excelDate = excelDate - 2;
            }
            else
            {
                excelDate = excelDate - 1;
            }
            return dateOfReference.AddDays(excelDate);
        } 

        public static string FormatDoubleValue(Double inValue)
        {
            if (inValue == 0.00)
                return string.Empty;
            else
                return inValue.ToString();
        }

        public static string FormatBooleanValue(int inValue)
        {
            string sReturn = "False";
            if (inValue == 1)
                sReturn = "True";
            return sReturn;
        }

        public static int FormatBooleanValue(bool inValue)
        {
            int iReturn = 0;
            if (inValue)
                iReturn = 1;
            return iReturn;
        }

        public static int GetDecimals(Double inValue)
        {
            int iReturn = 0;

            if (inValue > 0)
            {
                int iValueLength = Convert.ToInt32(inValue).ToString().Length;

                if ((iValueLength >= 0) && (iValueLength <= 4))
                    iReturn = 4 - iValueLength;
            }

            return iReturn;
        }

        public static string GetDecimalFormat(Double inValue)
        {
            string sReturn = string.Empty;

            if (inValue > 0)
            {
                int iValueLength = Convert.ToInt32(inValue).ToString().Length;
                iValueLength = GetDecimals(inValue);

                if ((iValueLength >= 0) && (iValueLength <= 4))
                {
                    string sDecimals = string.Empty.PadLeft(iValueLength, '0');
                    sReturn = "{0:0." + sDecimals + "}";
                }
                else
                    sReturn = "{0:0}";
                /*
                if (inValue < 1)
                    sReturn = "{0:0.000}";
                else if (inValue >= 1 && inValue < 20)
                    sReturn = "{0:0.00}";
                else if (inValue >= 20 && inValue < 200)
                    sReturn = "{0:0.0}";
                else if (inValue >= 200)
                    sReturn = "{0:0}";
                 * */
            }

            return sReturn;
        }

        public static string GetKermaDecimalFormat(Double inValue)
        {
            string sReturn = string.Empty;

            if (inValue > 0)
            {
                if (inValue < 10)
                    sReturn = "{0:0.0}";
                else
                    sReturn = "{0:0}";
            }

            return sReturn;
        }

        public static string FormatCertificateResults(Double inValue1, Double inValue2)
        {
            string sReturn = string.Empty;

            if (inValue1 + inValue2 > 0)
            {
                /*
                int iDecimalLength = 1; // Minimum one decimal .0
               
                int iDecimalLength1 = 0;
                int iDecimalLength2 = 0;

                int iDecimalPosition1 = inValue1.ToString().IndexOf('.');
                if (iDecimalPosition1 > 0)
                    iDecimalLength1 = inValue1.ToString().Substring(iDecimalPosition1 + 1).Length;

                int iDecimalPosition2 = inValue2.ToString().IndexOf('.');
                if (iDecimalPosition2 > 0)
                    iDecimalLength2 = inValue2.ToString().Substring(iDecimalPosition2 + 1).Length;


                //iDecimalLength1 = Math.Round(Value1 - Math.Floor(Value1),5).ToString().Length - 2;
                //iDecimalLength2 = Math.Round(Value2 - Math.Floor(Value2),5).ToString().Length - 2;

                if (iDecimalLength1 > iDecimalLength)
                    iDecimalLength = iDecimalLength1;
                if (iDecimalLength2 > iDecimalLength)
                    iDecimalLength = iDecimalLength2;
                                
                 sFormat = "{0:0." + new string('0', iDecimalLength) + "}";
                 sReturn = String.Format("{0:0.00}", Value1) + " ± " + String.Format("{0:0.00}", Value2); 
                 
                */

                string sFormat = GetDecimalFormat(inValue1);
                
                if (sFormat != string.Empty)
                    sReturn = String.Format(sFormat, inValue1) + " ± " + String.Format(sFormat, inValue2);
            }
            else
                sReturn = "-";

            return sReturn;
        }

        public static string FormatKermaResults(Double inValue1)
        {
            string sReturn = string.Empty;

            if (inValue1 > 0)
            {
                string sFormat = GetKermaDecimalFormat(inValue1);

                if (sFormat != string.Empty)
                    sReturn = String.Format(sFormat, inValue1);
            }
            else
                sReturn = "-";

            return sReturn;
        }

        public static string FormatDoubleValue(Double inValue, int iMinDecimalLength)
        {
            string sReturn = string.Empty;

            int iDecimalLength = iMinDecimalLength; 
            int iDecimalLength1 = 0;

            int iDecimalPosition1 = inValue.ToString().IndexOf('.');
            if (iDecimalPosition1 > 0)
                iDecimalLength1 = inValue.ToString().Substring(iDecimalPosition1 + 1).Length;

            if (iDecimalLength1 > iDecimalLength)
                iDecimalLength = iDecimalLength1;

            string sFormat = "{0:0." + new string('0', iDecimalLength) + "}";

            //sReturn = String.Format("{0:0.00}", Value1) + " ± " + String.Format("{0:0.00}", Value2);
            sReturn = String.Format(sFormat, inValue);

            return sReturn;
        }

        public static bool isLabCodeExists(string sLabCode)
        {
            bool bReturn = false;
            foreach (CountryClass Country in GlobalClass.Manager.CountryList)
            {
                foreach (OperatorClass OperatorCCode in Country.OperatorList)
                {
                    if (OperatorCCode.LabCode == sLabCode)
                    {
                        bReturn = true;
                        break;
                    }
                }
            }

            return bReturn;
        }

        public static int ExecuteSQL(string sSql)
        {
            int iReturn = 0;
            try
            {
                SqlConnection c = new SqlConnection();

                //capture the infomessage event to capture
                //print commands and warnings from SQL server
                c.InfoMessage += new SqlInfoMessageEventHandler(OnInfoMessage);

                //capture the statechange event to pickup notifications
                //when the state of the connection changes
                c.StateChange += new StateChangeEventHandler(OnStateChanged);

                c.ConnectionString = ConnectionStringWrite;

                c.Open();

                //ensure db is open
                if (c.State == ConnectionState.Open)
                {
                    //begin transaction
                    SqlTransaction sqlTransaction = c.BeginTransaction();

                    //create new sql command (stored proc, sql text, etc.)
                    SqlCommand sqlCommand = new SqlCommand();

                    //set the current connection for this command to this
                    //running instance of sql server
                    sqlCommand.Connection = c;

                    //set the transaction to the one just created above
                    sqlCommand.Transaction = sqlTransaction;

                    //set the sql to execute
                    sqlCommand.CommandText = sSql;

                    //set the timeout in seconds
                    sqlCommand.CommandTimeout = 30;

                    //set the type of command. setting the type to text
                    //allows raw sql to be pumped in - translation - 
                    //any sql can be input this way, including stored procs.
                    sqlCommand.CommandType = CommandType.Text;

                    //execute the sql and return the records affected by
                    //the sql statement
                    int recsaffected = sqlCommand.ExecuteNonQuery();

                    try
                    { // Commit the transaction
                        sqlTransaction.Commit();
                        iReturn = 1;
                    }
                    catch (Exception e)
                    { // Rollback the transaction
                        sqlTransaction.Rollback();
                        MessageBox.Show(e.Message.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        iReturn = 0;
                    }

                    //cleanup
                    sqlCommand = null;
                    sqlTransaction = null;
                }
                //close db - this is required. You must close whatever you open
                c.Close();
                c = null;

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                iReturn = 0;
            }

            return iReturn;
        }

        public static int ExecuteSQL(string sSql, List<ParameterClass> ParameterList)
        {
            int iReturn = 0;
            try
            {
                SqlConnection c = new SqlConnection();

                //capture the infomessage event to capture
                //print commands and warnings from SQL server
                c.InfoMessage += new SqlInfoMessageEventHandler(OnInfoMessage);

                //capture the statechange event to pickup notifications
                //when the state of the connection changes
                c.StateChange += new StateChangeEventHandler(OnStateChanged);

                c.ConnectionString = ConnectionStringWrite;

                c.Open();

                //ensure db is open
                if (c.State == ConnectionState.Open)
                {
                    //begin transaction
                    SqlTransaction sqlTransaction = c.BeginTransaction();

                    //create new sql command (stored proc, sql text, etc.)
                    SqlCommand sqlCommand = new SqlCommand();

                    //set the current connection for this command to this
                    //running instance of sql server
                    sqlCommand.Connection = c;

                    //set the transaction to the one just created above
                    sqlCommand.Transaction = sqlTransaction;

                    //set the sql to execute
                    sqlCommand.CommandText = sSql;

                    //set the timeout in seconds
                    sqlCommand.CommandTimeout = 30;

                    //set the type of command. setting the type to text
                    //allows raw sql to be pumped in - translation - 
                    //any sql can be input this way, including stored procs.
                    sqlCommand.CommandType = CommandType.Text;

                    foreach (ParameterClass Parameter in ParameterList)
                        if (Parameter.Parameter != string.Empty)
                            sqlCommand.Parameters.Add(Parameter.Parameter, SqlDbType.Image, Parameter.Value.Length).Value = Parameter.Value;

                    //execute the sql and return the records affected by
                    //the sql statement
                    int recsaffected = sqlCommand.ExecuteNonQuery();

                    try
                    { // Commit the transaction
                        sqlTransaction.Commit();
                        iReturn = 1;
                    }
                    catch (Exception e)
                    { // Rollback the transaction
                        sqlTransaction.Rollback();
                        MessageBox.Show(e.Message.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        iReturn = 0;
                    }

                    //cleanup
                    sqlCommand = null;
                    sqlTransaction = null;
                }
                //close db - this is required. You must close whatever you open
                c.Close();
                c = null;

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                iReturn = 0;
            }

            return iReturn;
        }

        public static DataTable GetDataTable(string sTables, string sSql)
        {
            DataTable ReturnDataTable;
            ReturnDataTable = null;

            //create a new sql connection
            SqlConnection c = new SqlConnection();

            //capture the infomessage event to capture
            //print commands and warnings from SQL server
            c.InfoMessage += new SqlInfoMessageEventHandler(OnInfoMessage);

            //capture the statechange event to pickup notifications
            //when the state of the connection changes
            c.StateChange += new StateChangeEventHandler(OnStateChanged);

            //set the connstring
            c.ConnectionString = GlobalClass.ConnectionStringRead;

            //connect to this sql server
            c.Open();

            //ensure db is open
            if (c.State == ConnectionState.Open)
            {
                //same sql command call as before but with no transaction
                //and no timeout
                SqlCommand sqlCommand = new SqlCommand();

                //set to the current SqlConnection
                sqlCommand.Connection = c;

                //set the sql to run (select statement/stored proc 
                //that returns rows)
                sqlCommand.CommandText = sSql;

                //set the type to Text, the work horse command type.
                sqlCommand.CommandType = CommandType.Text;

                //create a data adapter. this object is in charge of
                //getting back any data from sql server in the form of
                //rows returned from an sql statement.
                SqlDataAdapter sqlDataAdapter = new SqlDataAdapter();

                //set the select command or stored proc call to
                //the current SqlCommand object
                sqlDataAdapter.SelectCommand = sqlCommand;

                //create a dataset to hold the results set. a dataset
                //is an in-memory cache of information, just like a hash table
                DataSet dataSet = new DataSet();

                //fill up the dataset with the results from the sql
                sqlDataAdapter.Fill(dataSet, sTables);

                //get back the first datatable. remember the NextRecordset()
                //function from ADODB.Recordset back in regular ADO? well the
                //Tables collection works the same way. What it returns is the
                //results of each sql run. in this example, only one sql statement
                //was run to get an rs so the first index is all we need back.
                ReturnDataTable = dataSet.Tables[sTables];

                //cleanup
                sqlDataAdapter = null;
                sqlCommand = null;

                //close db - this is required. You must close whatever you open
                c.Close();
            }

            c = null;
            return ReturnDataTable;
        }
        
        public static DataTable GetYPDataTable(string sTables, string sSql)
        {
            DataTable ReturnDataTable;
            ReturnDataTable = null;

            //create a new sql connection
            SqlConnection c = new SqlConnection();

            //capture the infomessage event to capture
            //print commands and warnings from SQL server
            c.InfoMessage += new SqlInfoMessageEventHandler(OnInfoMessage);

            //capture the statechange event to pickup notifications
            //when the state of the connection changes
            c.StateChange += new StateChangeEventHandler(OnStateChanged);

            //set the connstring
            c.ConnectionString = GlobalClass.ConnectionStringYP;

            //connect to this sql server
            c.Open();

            //ensure db is open
            if (c.State == ConnectionState.Open)
            {
                //same sql command call as before but with no transaction
                //and no timeout
                SqlCommand sqlCommand = new SqlCommand();

                //set to the current SqlConnection
                sqlCommand.Connection = c;

                //set the sql to run (select statement/stored proc 
                //that returns rows)
                sqlCommand.CommandText = sSql;

                //set the type to Text, the work horse command type.
                sqlCommand.CommandType = CommandType.Text;

                //create a data adapter. this object is in charge of
                //getting back any data from sql server in the form of
                //rows returned from an sql statement.
                SqlDataAdapter sqlDataAdapter = new SqlDataAdapter();

                //set the select command or stored proc call to
                //the current SqlCommand object
                sqlDataAdapter.SelectCommand = sqlCommand;

                //create a dataset to hold the results set. a dataset
                //is an in-memory cache of information, just like a hash table
                DataSet dataSet = new DataSet();

                //fill up the dataset with the results from the sql
                sqlDataAdapter.Fill(dataSet, sTables);

                //get back the first datatable. remember the NextRecordset()
                //function from ADODB.Recordset back in regular ADO? well the
                //Tables collection works the same way. What it returns is the
                //results of each sql run. in this example, only one sql statement
                //was run to get an rs so the first index is all we need back.
                ReturnDataTable = dataSet.Tables[sTables];

                //cleanup
                sqlDataAdapter = null;
                sqlCommand = null;

                //close db - this is required. You must close whatever you open
                c.Close();
            }

            c = null;
            return ReturnDataTable;
        }
        
        public static DataTable GetExcelDataTable(string sTables, string sFileName)
        {
            DataTable ReturnDataTable = null;
            
            OleDbConnection odcConnection = new OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0;data source=" + sFileName + ";Extended Properties=Excel 8.0;");
            StringBuilder stbQuery = new StringBuilder();
            stbQuery.Append("SELECT * FROM [" + sTables + "$]");
            OleDbDataAdapter sqlDataAdapter = new OleDbDataAdapter(stbQuery.ToString(), odcConnection);
            DataSet dataSet = new DataSet();
            sqlDataAdapter.Fill(dataSet, sTables);
            ReturnDataTable = dataSet.Tables[sTables];

            odcConnection.Close();

            stbQuery = null;
            sqlDataAdapter = null;

            return ReturnDataTable;
        }

        public static int ExecuteDIRACSQL(string sSQL)
        {
            int iReturn = 0;

            SqlConnection c = new SqlConnection();

            //capture the infomessage event to capture
            //print commands and warnings from SQL server
            c.InfoMessage += new SqlInfoMessageEventHandler(OnInfoMessage);

            //capture the statechange event to pickup notifications
            //when the state of the connection changes
            c.StateChange += new StateChangeEventHandler(OnStateChanged);

            c.ConnectionString = ConnectionStringDIRACWrite;

            c.Open();

            //ensure db is open
            if (c.State == ConnectionState.Open)
            {
                //begin transaction
                SqlTransaction trans = c.BeginTransaction();

                //create new sql command (stored proc, sql text, etc.)
                SqlCommand sqlCommand = new SqlCommand();

                //set the current connection for this command to this
                //running instance of sql server
                sqlCommand.Connection = c;

                //set the transaction to the one just created above
                sqlCommand.Transaction = trans;

                //set the sql to execute
                sqlCommand.CommandText = sSQL;

                //set the timeout in seconds
                sqlCommand.CommandTimeout = 30;

                //set the type of command. setting the type to text
                //allows raw sql to be pumped in - translation - 
                //any sql can be input this way, including stored procs.
                sqlCommand.CommandType = CommandType.Text;

                //execute the sql and return the records affected by
                //the sql statement
                int recsaffected = sqlCommand.ExecuteNonQuery();

                try
                { // Commit the transaction
                    trans.Commit();
                    iReturn = 1;
                }
                catch
                { // Rollback the transaction
                    trans.Rollback();
                    iReturn = 0;
                }

                //cleanup
                sqlCommand = null;
                trans = null;
            }
            //close db - this is required. You must close whatever you open
            c.Close();
            c = null;
            return iReturn;
        }

        public static DataTable GetDIRACDataTable(string sTables, string sSQL)
        {
            DataTable ReturnDataTable;
            ReturnDataTable = null;

            //create a new sql connection
            SqlConnection c = new SqlConnection();

            //capture the infomessage event to capture
            //print commands and warnings from SQL server
            c.InfoMessage += new SqlInfoMessageEventHandler(OnInfoMessage);

            //capture the statechange event to pickup notifications
            //when the state of the connection changes
            c.StateChange += new StateChangeEventHandler(OnStateChanged);

            //set the connstring
            c.ConnectionString = GlobalClass.ConnectionStringDIRACRead;

            //connect to this sql server
            c.Open();

            //ensure db is open
            if (c.State == ConnectionState.Open)
            {
                //same sql command call as before but with no transaction
                //and no timeout
                SqlCommand sqlCommand = new SqlCommand();

                //set to the current SqlConnection
                sqlCommand.Connection = c;

                //set the sql to run (select statement/stored proc 
                //that returns rows)
                sqlCommand.CommandText = sSQL;

                //set the type to Text, the work horse command type.
                sqlCommand.CommandType = CommandType.Text;

                //create a data adapter. this object is in charge of
                //getting back any data from sql server in the form of
                //rows returned from an sql statement.
                SqlDataAdapter sqlDataAdapter = new SqlDataAdapter();

                //set the select command or stored proc call to
                //the current SqlCommand object
                sqlDataAdapter.SelectCommand = sqlCommand;

                //create a dataset to hold the results set. a dataset
                //is an in-memory cache of information, just like a hash table
                DataSet dataSet = new DataSet();

                //fill up the dataset with the results from the sql
                sqlDataAdapter.Fill(dataSet, sTables);

                //get back the first datatable. remember the NextRecordset()
                //function from ADODB.Recordset back in regular ADO? well the
                //Tables collection works the same way. What it returns is the
                //results of each sql run. in this example, only one sql statement
                //was run to get an rs so the first index is all we need back.
                ReturnDataTable = dataSet.Tables[sTables];

                //cleanup
                sqlDataAdapter = null;
                sqlCommand = null;

                //close db - this is required. You must close whatever you open
                c.Close();
            }

            c = null;
            return ReturnDataTable;
        }

        public static int LogUserAction(int iUserActionID, int iOperatorID, string sUserAction, string sUserActionDetails)
        {
            int iReturn = 0;
            int iUserID = -1;
            string sUserName = "System User " + System.Windows.Forms.SystemInformation.UserName;

            if (GlobalClass.User != null)
            {
                iUserID = GlobalClass.User.UserID;
                sUserName = FormatStringValue(GlobalClass.User.UserName.Trim(), 50);
            }
            sUserAction = sUserAction.Replace("#UserName#", sUserName);

            string sIPAddress = "IP ADDRESS: [" + GetIPAddress() + "]";
            string sSessionID = ImporterVersion_Current;

            //----- Execute some SQL -----------
            //iReturn = ExecuteSQL("INSERT INTO dbo.UserLogFile (UserActionID, OperatorID, UserID, UserName, UserIP, SessionID, UserAction, UserActionDetails) VALUES (" + sUserActionID.ToString() + "," + iOperatorID.ToString() + "," + iUserID.ToString() + ",'" + sUserName + "','" + sIPAddress + "','" + sSessionID + "','" + sUserAction + "','" + sUserActionDetails + "')");
            string sFields = "UserActionID, OperatorID, UserID, UserName, UserIP, SessionID, UserAction, UserActionDetails";
            string sValues = string.Empty;
            sValues = sValues + iUserActionID.ToString() + ",";
            sValues = sValues + iOperatorID.ToString() + ",";
            sValues = sValues + iUserID.ToString() + ",";
            sValues = sValues + "'" + GlobalClass.FormatStringValue(sUserName.Trim(), 150) + "', ";
            sValues = sValues + "'" + GlobalClass.FormatStringValue(sIPAddress.Trim(), 150) + "', ";
            sValues = sValues + "'" + GlobalClass.FormatStringValue(sSessionID.Trim(), 150) + "', ";
            sValues = sValues + "'" + GlobalClass.FormatStringValue(sUserAction.Trim(), 150) + "', ";
            sValues = sValues + "'" + GlobalClass.FormatStringValue(sUserActionDetails.Trim(), 2000) + "' ";


            iReturn = ExecuteSQL("INSERT INTO dbo.UserLogFile (" + sFields + ") VALUES (" + sValues + ")");


            //----- Execute some SQL -----------
            //iReturn = ExecuteSQL("INSERT INTO dbo.UserLogFile (UserActionID, OperatorID, UserID, UserName, UserIP, SessionID, UserAction, UserActionDetails) VALUES (" + sUserActionID.ToString() + "," + iOperatorID.ToString() + "," + iUserID.ToString() + ",'" + sUserName + "','" + sIPAddress + "','" + sSessionID + "','" + sUserAction + "','" + sUserActionDetails + "')");
            //iReturn = ExecuteSQL("INSERT INTO dbo.UserLogFile (UserActionID, OperatorID, UserID, UserName, UserIP, SessionID, UserAction, UserActionDetails) VALUES (" + sUserActionID.ToString() + "," + iOperatorID.ToString() + "," + iUserID.ToString() + ",'" + sUserName + "','" + sIPAddress + "','" + sSessionID + "','" + sUserAction + "','" + sUserActionDetails + "')");

            return iReturn;
        }

        public static int LogDIRACUserAction(int iUserActionID, int iOperatorID, string sUserAction, string sUserActionDetails)
        {
            int iReturn = 0;
            int iUserID = -1;

            string sUserName = "System User " + System.Windows.Forms.SystemInformation.UserName;

            if (GlobalClass.User != null)
            {
                iUserID = GlobalClass.User.UserID;
                sUserName = FormatStringValue(GlobalClass.User.UserName.Trim(), 50);
            }

            //string sIPAddress = Path.GetFileName(Application.ExecutablePath);
            string sIPAddress = "IP ADDRESS: [" + GetIPAddress() + "]";
            string sSessionID = ImporterVersion_Current;
            sUserAction = sUserAction.Replace("#UserName#", sUserName);

            string sFields = "UserActionID, OperatorID, UserID, UserName, UserIP, SessionID, UserAction, UserActionDetails";
            string sValues = string.Empty;
            sValues = sValues + iUserActionID.ToString() + ",";
            sValues = sValues + iOperatorID.ToString() + ",";
            sValues = sValues + iUserID.ToString() + ",";
            sValues = sValues + "'" + GlobalClass.FormatStringValue(sUserName, 50) + "',";
            sValues = sValues + "'" + GlobalClass.FormatStringValue(sIPAddress, 50) + "',";
            sValues = sValues + "'" + GlobalClass.FormatStringValue(sSessionID, 50) + "',";
            sValues = sValues + "'" + GlobalClass.FormatStringValue(sUserAction, 150) + "',";
            sValues = sValues + "'" + GlobalClass.FormatStringValue(sUserActionDetails, 2000) + "'";

            //----- Execute some SQL -----------
            iReturn = ExecuteDIRACSQL("INSERT INTO dbo.UserLogFile (" + sFields + ") VALUES (" + sValues + ")");

            return iReturn;
        }

        public static bool IsEmail(string Email)
        {
            string strRegex = @"^([a-zA-Z0-9_\-\.]+)@((\[[0-9]{1,3}" +
                @"\.[0-9]{1,3}\.[0-9]{1,3}\.)|(([a-zA-Z0-9\-]+\" +
                @".)+))([a-zA-Z]{2,4}|[0-9]{1,3})(\]?)$";
            Regex re = new Regex(strRegex);
            if (re.IsMatch(Email))
                return (true);
            else
                return (false);
        }

        public static string GetIPAddress()
        {
            string sReturn = string.Empty;
            try
            {
                // Getting Ip address of local machine...
                // First get the host name of local machine.
                String strHostName = Dns.GetHostName();

                // Then using host name, get the IP address list..
                IPHostEntry ipEntry = Dns.GetHostEntry(strHostName);
                IPAddress[] addr = ipEntry.AddressList;

                System.OperatingSystem osInfo = System.Environment.OSVersion;
                if (osInfo.Platform == PlatformID.Win32NT)
                {
                    if (osInfo.Version.Major == 5) //Windows 2000 or Windows XP
                    {
                        for (int i = 0; i < addr.Length; i++)
                            sReturn = sReturn + addr[i].ToString() + ",";
                        sReturn = sReturn.Remove(sReturn.Length - 1, 1);
                    }
                    else if (osInfo.Version.Major == 6) //Windows 7
                    {
                        sReturn = addr[2].ToString();
                    }
                }
            }
            catch
            { }
            return sReturn;
        }

        public static void SetSystemUser()
        {
            try
            {
                string sUserName = Environment.UserName;

                GlobalClass.User = null;
                GlobalClass.User = GlobalClass.Manager.GetUserByUserName(sUserName);
                //GlobalClass.User = GlobalClass.Manager.GetUserByUserName("ypynda");
                if (GlobalClass.User != null)
                {
                    GlobalClass.User.LoadUserPermissions();
                    //GlobalClass.Manager.PopulateOperatorList();
                    //GlobalClass.Manager.ReportManager.PopulateReportList();
                    //MessageBox.Show(GlobalClass.User.UserName + " - welcome to the IAEA/WHO SSDL Network", "Welcome to the SSDL Network", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public static void DisablePanelWithTag(Control BaseControls, string sTag)
        {
            foreach (Control c in BaseControls.Controls)
            {
                if ((c is Panel))
                {
                    if ((c as Panel).Tag != null)
                    {
                        string sTagValue = (c as Panel).Tag.ToString();
                        if (sTagValue != string.Empty)
                        {
                            if (sTagValue == sTag)
                            {
                                GlobalClass.DisableControls(c);
                            }
                            else
                            {
                                if (sTag == "Sys")
                                    EnableControls(c, System.Drawing.Color.FromArgb(255, 255, 192));
                                else
                                    EnableControls(c, System.Drawing.Color.FromArgb(192, 255, 255));
                            }
                        }
                    }
                    DisablePanelWithTag(c, sTag);
                }
            }
        }

        public static void EnableControls(Control BaseControls, System.Drawing.Color iColor)
        {
            foreach (Control c in BaseControls.Controls)
            {
                if ((c is TextBox) || (c is CheckBox) || (c is ComboBox) || (c is RadioButton) ||
                    (c is Button) || (c is DateTimePicker) || (c is LinkLabel) || 
                    (c is ListView) || (c is TreeView) || (c is PictureBox))
                {
                    if (c is TextBox)
                    {
                        (c as TextBox).ReadOnly = false;
                        //if (iColor != null)
                        (c as TextBox).BackColor = iColor;
                        //(c as TextBox).BackColor = System.Drawing.SystemColors.Window;
                    }
                    else if (c is ComboBox)
                    {
                        (c as ComboBox).DropDownStyle = ComboBoxStyle.DropDownList;
                        (c as ComboBox).Enabled = true;
                    }
                    else if (c is CheckBox)
                        (c as CheckBox).Enabled = true;
                    else if (c is RadioButton)
                        (c as RadioButton).Enabled = true;
                    else if (c is Button)
                        (c as Button).Enabled = true;
                    else if (c is DateTimePicker)
                        (c as DateTimePicker).Enabled = true;
                    else if (c is LinkLabel)
                        (c as LinkLabel).Enabled = true;
                    else if (c is PictureBox)
                        (c as PictureBox).Enabled = true;
                    else if (c is ListView)
                    {
                        //(c as ListView).Enabled = false;
                        (c as ListView).ForeColor = System.Drawing.SystemColors.Window;
                        //(c as ListView).ContextMenuStrip = null;
                        //(c as ListView).MouseDoubleClick -= null;
                    }
                    else if (c is TreeView)
                    {
                        //(c as ListView).Enabled = false;
                        (c as TreeView).ForeColor = System.Drawing.SystemColors.Window;
                        //(c as TreeView).ContextMenuStrip = null;
                        //(c as TreeView).MouseDoubleClick -= null;
                    }
                }
                else
                {
                    EnableControls(c, iColor);
                }
            }
        }

        public static void DisableControls(Control BaseControls)
        {
            foreach (Control c in BaseControls.Controls)
            {
                if ((c is TextBox) || (c is CheckBox) || (c is ComboBox) || (c is RadioButton) ||
                    (c is Button) || (c is DateTimePicker) || (c is LinkLabel) ||
                    (c is ListView) || (c is TreeView) || (c is PictureBox))
                {
                    if (c is TextBox)
                    {
                        (c as TextBox).ReadOnly = true;
                        (c as TextBox).BackColor = System.Drawing.Color.FromName("Control");
                    }
                    else if (c is ComboBox)
                    {
                        (c as ComboBox).DropDownStyle = ComboBoxStyle.DropDownList;
                        (c as ComboBox).Enabled = false;
                    }
                    else if (c is CheckBox)
                        (c as CheckBox).Enabled = false;
                    else if (c is RadioButton)
                        (c as RadioButton).Enabled = false;
                    else if (c is Button)
                        (c as Button).Enabled = false;
                    else if (c is DateTimePicker)
                        (c as DateTimePicker).Enabled = false;
                    else if (c is LinkLabel)
                        (c as LinkLabel).Enabled = false;
                    else if (c is PictureBox)
                        (c as PictureBox).Enabled = false;
                    else if (c is ListView)
                    {
                        //(c as ListView).Enabled = false;
                        (c as ListView).ForeColor = System.Drawing.Color.Gray;
                        (c as ListView).ContextMenuStrip = null;
                        (c as ListView).MouseDoubleClick -= null;
                    }
                    else if (c is TreeView)
                    {
                        //(c as ListView).Enabled = false;
                        (c as TreeView).ForeColor = System.Drawing.Color.Gray;
                        //(c as TreeView).ContextMenuStrip = null;
                        //(c as TreeView).MouseDoubleClick -= null;
                    }
                }
                else 
                {
                    DisableControls(c);
                }
            }
        }
        
        public static Control FindControl(Control BaseControls, string sControlName)
        {
            Control cReturn = null;
            foreach (Control c in BaseControls.Controls)
            {
                if ((c is TextBox) || (c is CheckBox) || (c is ComboBox) || (c is RadioButton) ||
                    (c is Button) || (c is DateTimePicker) || (c is LinkLabel) ||
                    (c is ListView) || (c is TreeView) || (c is PictureBox))
                {
                    //if (c is DateTimePicker)
                    {
                        if (c.Name == sControlName)
                        {
                            cReturn = c;
                            break;
                        }
                    }
                }
                else
                {
                    cReturn = FindControl(c, sControlName);
                    if (cReturn != null)
                        break;
                }
            }
            return cReturn;
        }

        public static Control FindControlByTag(Control BaseControls, string sControlTag)
        {
            Control cReturn = null;
            foreach (Control c in BaseControls.Controls)
            {
                if ((c is TextBox) || (c is CheckBox) || (c is ComboBox) || (c is RadioButton) ||
                    (c is Button) || (c is DateTimePicker) || (c is LinkLabel) ||
                    (c is ListView) || (c is TreeView) || (c is PictureBox))
                {
                    if (c.Tag != null)
                    {
                        if (c.Tag.ToString() == sControlTag)
                        {
                            cReturn = c;
                            break;
                        }
                    }
                }
                else
                {
                    cReturn = FindControlByTag(c, sControlTag);
                    if (cReturn != null)
                        break;
                }
            }
            return cReturn;
        }

        public static System.Drawing.Image GetImageFromBytes(byte[] buffer)
        {
            System.Drawing.Image iReturn = null;
            if (buffer != null)
            {
                if (buffer.Length > 0)
                {
                    using (MemoryStream ms = new MemoryStream(buffer))
                    {
                        iReturn = System.Drawing.Image.FromStream(ms);
                    }
                }
            }
            return iReturn;
        }

        public static byte[] GetFile(string sFullFileName)
        {
            byte[] BinaryFile = null;

            if (File.Exists(sFullFileName))
            {
                FileStream stream = new FileStream(sFullFileName, FileMode.Open, FileAccess.Read);
                BinaryReader reader = new BinaryReader(stream);

                BinaryFile = reader.ReadBytes((int)stream.Length);

                reader.Close();
                stream.Close();
            }

            return BinaryFile;
        }

        public static int UploadFile(string sSql, string sFullFileName)
        {
            byte[] PDFFile = GetFile(sFullFileName);

            SqlConnection c = new SqlConnection();

            //capture the infomessage event to capture
            //print commands and warnings from SQL server
            c.InfoMessage += new SqlInfoMessageEventHandler(OnInfoMessage);

            //capture the statechange event to pickup notifications
            //when the state of the connection changes
            c.StateChange += new StateChangeEventHandler(OnStateChanged);

            c.ConnectionString = ConnectionStringWrite;

            c.Open();

            //create new sql command (stored proc, sql text, etc.)
            SqlCommand sqlCommand = new SqlCommand();

            //set the sql to execute
            sqlCommand.CommandText = sSql;
            //set the current connection for this command to this
            //running instance of sql server
            sqlCommand.Connection = c;

            sqlCommand.Parameters.Add("@PDFFile", SqlDbType.Image, PDFFile.Length).Value = PDFFile;

            sqlCommand.ExecuteNonQuery();

            return 1;
        }

        public static Bitmap ConvertToGrayscale(Bitmap source)
        {
            Bitmap bm = new Bitmap(source.Width, source.Height);

            for (int y = 0; y < bm.Height; y++)
            {
                for (int x = 0; x < bm.Width; x++)
                {
                    System.Drawing.Color c = source.GetPixel(x, y);
                    //if (((c.R + c.G + c.B) > 0) && ((c.R + c.G + c.B) < 755))
                    if ((c.R + c.G + c.B) > 0)
                    {
                        int luma = (int)(c.R * 0.3 + c.G * 0.59 + c.B * 0.11);
                        bm.SetPixel(x, y, System.Drawing.Color.FromArgb(luma, luma, luma));
                    }
                    //else
                    //    bm.SetPixel(x, y, Color.FromArgb(255, 255, 255));
                    
                }
            }
            return bm;
        }

        public static string GetTemplateObject(string sTemplate)
        {
            string sTemplateObject = string.Empty;
            string sLine = string.Empty;

            //Set the current directory.
            Directory.SetCurrentDirectory(GlobalClass.sApplicationStartupPath);

            if (File.Exists(sTemplate))
            {
                try
                {
                    string sExtension = Path.GetExtension(sTemplate).ToLower();
                    if (sExtension == ".txt")
                    {
                        // Create an instance of StreamReader to read from a file.
                        // The using statement also closes the StreamReader.
                        using (StreamReader sr = new StreamReader(sTemplate, System.Text.Encoding.Default, true))
                        {
                            // Read and display lines from the file until the end of 
                            // the file is reached.
                            while ((sLine = sr.ReadLine()) != null)
                            {
                                if (sLine.Contains("<<TemplateObject>>"))
                                    sTemplateObject = sLine.Replace("<<TemplateObject>>", "").Trim();
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Template Object", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    // The file could not be read
                }
            }
            else
                MessageBox.Show("Template " + sTemplate + " does not exist.", "Template Object", MessageBoxButtons.OK, MessageBoxIcon.Error);

            return sTemplateObject;
        }

        public static string GenerateEmail(string sTemplate, string sAttachment, string sSaveAs)
        {
            string sReturn = string.Empty;

            string sEmailFrom = string.Empty;
            string sEmailTo = string.Empty;
            string sEmailCC = string.Empty;
            string sSubject = string.Empty;
            string sBody = string.Empty;
            string sLine = string.Empty;
            string sAttachments = string.Empty;
            string sSignature = string.Empty;
            //string sSaveAs = string.Empty;

            //Set the current directory.
            Directory.SetCurrentDirectory(GlobalClass.sApplicationStartupPath);

            string[] sAttachmentList = sAttachment.Trim().Split(';');
            foreach (string sAttachmentFile in sAttachmentList)
            {
                if (File.Exists(sAttachmentFile.Trim()))
                    sAttachments = sAttachments + sAttachmentFile.Trim() + ";";
            }

            if (File.Exists(sTemplate))
            {
                try
                {
                    // Create an instance of StreamReader to read from a file.
                    // The using statement also closes the StreamReader.
                    using (StreamReader sr = new StreamReader(sTemplate, System.Text.Encoding.Default, true))
                    {
                        // Read and display lines from the file until the end of the file is reached.
                        while ((sLine = sr.ReadLine()) != null)
                        {
                            if (sLine.Contains("<<eMail.Subject>>"))
                                sSubject = sLine.Replace("<<eMail.Subject>>", "").Trim();
                            else if (sLine.Contains("<<eMail.TemplateObject>>"))
                            { }
                            else if (sLine.Contains("<<eMail.From>>"))
                                sEmailFrom = sLine.Replace("<<eMail.From>>", string.Empty).Trim();
                            else if (sLine.Contains("<<Email>>"))
                                sEmailTo = sLine.Replace("<<Email>>", "").Trim();
                            else if (sLine.Contains("<<eMail.CC>>"))
                                sEmailCC = sLine.Replace("<<eMail.CC>>", "").Trim();

                            else if (sLine.Contains("<<eMail.Attachments>>"))
                            {
                                sAttachmentList = sLine.Replace("<<eMail.Attachments>>", "").Trim().Split(';');
                                foreach (string sAttachmentFile in sAttachmentList)
                                {
                                    if (File.Exists(GlobalClass.sApplicationStartupPath + "\\" + GlobalClass.sApplicationTemplatesFolder + sAttachmentFile.Trim()))
                                        sAttachments = sAttachments + GlobalClass.sApplicationStartupPath + "\\" + GlobalClass.sApplicationTemplatesFolder + sAttachmentFile.Trim() + ";";
                                }
                            }
                            else if (sLine.Contains("<<Signature>>"))
                                sSignature = sLine.Replace("<<Signature>>", "").Trim();
                            else
                                sBody = sBody + sLine + System.Environment.NewLine;
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Generating e-mail", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    // The file could not be read
                }
            }
            else
                MessageBox.Show("Template " + sTemplate + " does not exist.", "Generating e-mail", MessageBoxButtons.OK, MessageBoxIcon.Error);


            // Generate E-mail
            GlobalClass.GenerateEmailNotification(sEmailFrom, sEmailTo, sEmailCC, sSubject, sBody, sAttachments, sSaveAs, sSignature);
            return sReturn;
        }

        public static int GenerateEmailNotification(string sEmailFrom, string sEmailTo, string sEmailCC, string sSubject, string sBody, string sAttachments, string sSaveAs, string sSignature)
        {
            int iReturn = 0;

            //string sLine = string.Empty;

            Outlook.Application app = new Outlook.ApplicationClass();
            Outlook.NameSpace NS = app.GetNamespace("MAPI");
            Outlook.MAPIFolder inboxFld = NS.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);

            Outlook.Application oApp;
            Outlook._NameSpace oNameSpace;
            Outlook.MAPIFolder oOutboxFolder;


            //Return a reference to the MAPI layer 
            oApp = new Outlook.Application();

            oNameSpace = oApp.GetNamespace("MAPI");
            oNameSpace.Logon(null, null, true, true);

            //gets defaultfolder for my Outlook Outbox 
            oOutboxFolder = oNameSpace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderSentMail);

            Outlook._MailItem oMailItem = (Outlook._MailItem)oApp.CreateItem(Outlook.OlItemType.olMailItem);

            if (sEmailFrom.Trim() != string.Empty)
                oMailItem.SentOnBehalfOfName = sEmailFrom;

            oMailItem.To = sEmailTo;

            if (sEmailCC != string.Empty)
                oMailItem.CC = sEmailCC;

            oMailItem.Subject = sSubject;

            //oMailItem.BodyFormat = Outlook.OlBodyFormat.olFormatPlain;
            //oMailItem.Body = sBody;
            
            sBody = sBody.Replace("\r\n", "<br/>");
            oMailItem.BodyFormat = Outlook.OlBodyFormat.olFormatHTML;

            if (sSignature != string.Empty)
                oMailItem.HTMLBody = "<HTML><BODY>" + sBody + ReadSignature(sSignature) + "</BODY></HTML>";
                //oMailItem.HTMLBody = sBody + "<div>" + ReadSignature(sSignature) + "</div>";
            else
                oMailItem.HTMLBody = "<HTML><BODY>" + sBody + "</BODY></HTML>";
            
            string[] Attachments = sAttachments.Split(';');

            foreach (string sAttachment in Attachments)
            {                
                if (sAttachment.Trim() != string.Empty)
                {
                    if (File.Exists(sAttachment))
                    {
                        string sDisplayName = System.IO.Path.GetFileName(sAttachment.Trim()).Trim();
                        int iPosition = (int)oMailItem.HTMLBody.Length + 1;
                        int iAttachType = (int)Outlook.OlAttachmentType.olByValue;
                        Outlook.Attachment oAttach = oMailItem.Attachments.Add(sAttachment.Trim(), iAttachType, iPosition, sDisplayName);
                    }
                }
            }
            oMailItem.SaveSentMessageFolder = oOutboxFolder;
            //uncomment this to also save this in your draft 
            oMailItem.Save();

            if (sSaveAs.Trim() != string.Empty)
                oMailItem.SaveAs(sSaveAs, Outlook.OlSaveAsType.olMSGUnicode);

            //adds it to the Drafts Folde
            oMailItem.Display(false); //Display false = no modal
         
            return iReturn;
        }

        public static string ReadSignature(string sSignature)
        {
            string sApplicationDataFolder = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\Microsoft\\Signatures";
            string sReturn = string.Empty;
            DirectoryInfo diInfo = new DirectoryInfo(sApplicationDataFolder);

            if (diInfo.Exists)
            {
                FileInfo[] fiSignature = diInfo.GetFiles("*.htm");

                if (sSignature == string.Empty)
                {
                    if (File.Exists(fiSignature[0].FullName))
                    {
                        string sFileName = fiSignature[0].Name.Replace(fiSignature[0].Extension, string.Empty);
                        StreamReader sr = new StreamReader(fiSignature[0].FullName, Encoding.Default);
                        string sSignatureFile = sr.ReadToEnd();

                        if (!string.IsNullOrEmpty(sSignatureFile))
                            sReturn = sReturn + sSignatureFile.Replace(sFileName + "_files/", sApplicationDataFolder + "/" + sFileName + "_files/") + "<br/>";
                    }
                }
                else if (sSignature != string.Empty)
                {
                    for (int i = 0; i < fiSignature.Length; i++)
                    {
                        if (File.Exists(fiSignature[i].FullName))
                        {
                            string sFileName = fiSignature[i].Name.Replace(fiSignature[i].Extension, string.Empty);
                            if (sFileName == sSignature)
                            {
                                StreamReader sr = new StreamReader(fiSignature[i].FullName, Encoding.Default);
                                string sSignatureFile = sr.ReadToEnd();

                                if (!string.IsNullOrEmpty(sSignatureFile))
                                    sReturn = sReturn + sSignatureFile.Replace(sFileName + "_files/", sApplicationDataFolder + "/" + sFileName + "_files/") + "<br/>";
                            }
                        }
                    }
                }
            }
            return sReturn;
        }

        public static bool ColumnExists(DataTable Table, string sColumnName)
        {
            bool bReturn = false;
            foreach (DataColumn Column in Table.Columns)
            {
                if (Column.ColumnName == sColumnName)
                {
                    bReturn = true;
                    break;
                }
            }

            return bReturn;
        }

        public static bool ColumnExists(DataRow SetDr, string sColumnName)
        {
            bool bReturn = false;
            foreach (DataColumn Column in SetDr.Table.Columns)
            {
                if (Column.ColumnName == sColumnName)
                {
                    bReturn = true;
                    break;
                }
            }

            return bReturn;
        }

        public static string GetDataRowValue(DataRow dr, string sColumnName, string sFormat)
        {
            string sReturn = string.Empty;

            if (ColumnExists(dr.Table, sColumnName))
            {
                if (dr[sColumnName] != DBNull.Value)
                {
                    Type myType = dr[sColumnName].GetType();
                    if (myType.FullName == "System.String")
                    {
                        sReturn = (string)dr[sColumnName];
                        if (sReturn == "Off")
                            sReturn = string.Empty;
                    }
                    else if (myType.FullName == "System.Int32")
                    {
                        if (sFormat == string.Empty)
                            sReturn = ((Int32)dr[sColumnName]).ToString();
                        else
                            sReturn = String.Format(sFormat, (Int32)dr[sColumnName]);
                    }
                    else if (myType.FullName == "System.Int64")
                    {
                        if (sFormat == string.Empty)
                            sReturn = ((Int64)dr[sColumnName]).ToString();
                        else
                            sReturn = String.Format(sFormat, (Int64)dr[sColumnName]);
                    }
                    else if (myType.FullName == "System.Decimal")
                    {
                        if (sFormat == string.Empty)
                            sReturn = ((Decimal)dr[sColumnName]).ToString();
                        else
                            sReturn = String.Format(sFormat, (Decimal)dr[sColumnName]);
                    }
                    else if (myType.FullName == "System.Double")
                    {
                        if (sFormat == string.Empty)
                            sReturn = ((Double)dr[sColumnName]).ToString();
                        else
                            sReturn = String.Format(sFormat, (Double)dr[sColumnName]);
                    }

                    else if (myType.FullName == "System.DateTime")
                        sReturn = GlobalClass.FormatDateTimeValue((DateTime)dr[sColumnName]);
                    else if (myType.FullName == "System.Boolean")
                        sReturn = GlobalClass.FormatBooleanValue(GlobalClass.FormatBooleanValue((bool)dr[sColumnName]));
                }
            }

            return sReturn;
        }
        public static string SaveRecord(string sTableView, Object ObjectType)
        {
            string sReturn = string.Empty;

            if (sTableView != string.Empty)
            {
                int iOperatorID = -1;
                string sTable = sTableView.Replace("dbo.", "dbo.tbl").Trim();
                string sPrimaryKeyColumn = string.Empty;
                string sPrimaryKeyValue = string.Empty;

                DataTable ColumnInfoDataTable = GlobalClass.GetDataTable(sTableView, "SELECT * FROM dbo.fn_TableColumnInfo('" + sTable + "') WHERE PrimaryKeyColumn = 1 ORDER BY 2");
                if (ColumnInfoDataTable != null)
                {
                    if (ColumnInfoDataTable.Rows.Count == 1)
                    {
                        string sCurrentFieldType = (string)ColumnInfoDataTable.Rows[0]["TypeName"];
                        string sCurrentField = (string)ColumnInfoDataTable.Rows[0]["ColumnName"];
                        int iCurrentFieldLength = 0;
                        int iCurrentFieldScale = 0;

                        if (ColumnInfoDataTable.Rows[0]["ColumnPrecision"] != DBNull.Value)
                            iCurrentFieldLength = (Int16)ColumnInfoDataTable.Rows[0]["ColumnPrecision"];

                        if (ColumnInfoDataTable.Rows[0]["ColumnScale"] != DBNull.Value)
                            iCurrentFieldScale = (int)ColumnInfoDataTable.Rows[0]["ColumnScale"];

                        string sCurrentValue = string.Empty;

                        Type myType = ObjectType.GetType();
                        PropertyInfo Property = myType.GetProperty(sCurrentField);

                        if (Property != null)
                        {
                            try
                            {
                                if (sCurrentFieldType == "nvarchar")
                                    sCurrentValue = "'" + GlobalClass.FormatStringValue(Convert.ToString(Property.GetValue(ObjectType, null)), iCurrentFieldLength) + "'";
                                else if (sCurrentFieldType == "int")
                                    sCurrentValue = Convert.ToInt32(Property.GetValue(ObjectType, null)).ToString().Trim();
                                else if (sCurrentFieldType == "numeric")
                                    sCurrentValue = Math.Round(Convert.ToDecimal(Property.GetValue(ObjectType, null)), iCurrentFieldScale).ToString().Trim();
                                else if (sCurrentFieldType == "decimal")
                                    sCurrentValue = Math.Round(Convert.ToDecimal(Property.GetValue(ObjectType, null)), iCurrentFieldScale).ToString().Trim();
                                else if (sCurrentFieldType == "float")
                                {
                                    iCurrentFieldScale = 7;
                                    sCurrentValue = Math.Round(Convert.ToDecimal(Property.GetValue(ObjectType, null)), iCurrentFieldScale).ToString().Trim();
                                }
                                else if (sCurrentFieldType == "bit")
                                    if (Convert.ToBoolean(Property.GetValue(ObjectType, null)) == true) sCurrentValue = "1"; else sCurrentValue = "0";
                                else if (sCurrentFieldType == "datetime")
                                    sCurrentValue = "'" + Convert.ToDateTime(Property.GetValue(ObjectType, null)).ToString().Trim() + "'";
                            }
                            catch
                            {
                                sReturn = sReturn + "Can not convert value " + sCurrentField + " Table " + sTableView + System.Environment.NewLine;
                                MessageBox.Show("Can not convert value " + sCurrentField + " Table " + sTableView, "Conversion Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }

                            if ((bool)ColumnInfoDataTable.Rows[0]["PrimaryKeyColumn"] == true)
                            {
                                sPrimaryKeyColumn = sCurrentField;
                                sPrimaryKeyValue = sCurrentValue;
                            }
                        }

                        // Get OperatorID value
                        PropertyInfo pinfoOperatorID = myType.GetProperty("OperatorID");
                        if (pinfoOperatorID != null)
                            iOperatorID = Convert.ToInt32(pinfoOperatorID.GetValue(ObjectType, null));
                        else
                            sReturn = sReturn + "SQL statement is empty. Can not identify OperatorID value." + System.Environment.NewLine;


                        if ((sPrimaryKeyColumn != string.Empty) && (sPrimaryKeyValue != string.Empty))
                        {
                            string sCheckRecordSql = "SELECT * FROM " + sTableView + " WHERE [" + sPrimaryKeyColumn + "] = " + sPrimaryKeyValue;
                            string sOperation = string.Empty;
                            string sSqlStatement = string.Empty;

                            DataTable CheckRecordDataTable = GlobalClass.GetDataTable("CheckRecordTable", sCheckRecordSql);                            
                            if (CheckRecordDataTable != null)
                            {
                                List<ParameterClass> ParameterList = new List<ParameterClass>();
                                if (CheckRecordDataTable.Rows.Count == 0)
                                {
                                    sOperation = "INSERT";
                                    sSqlStatement = GlobalClass.GenerateInsertStatement(sTableView, ObjectType, ParameterList);
                                }
                                else
                                {
                                    sOperation = "UPDATE";
                                    sSqlStatement = GlobalClass.GenerateUpdateStatement(sTableView, ObjectType, ParameterList);
                                }

                                if (sSqlStatement != string.Empty)
                                {
                                    try
                                    {
                                        if (ParameterList.Count > 0)
                                            GlobalClass.ExecuteSQL(sSqlStatement, ParameterList);
                                        else
                                            GlobalClass.ExecuteSQL(sSqlStatement);

                                        GlobalClass.LogUserAction(2, iOperatorID, "User [#UserName#] Execute " + sOperation + " on the table " + sTable, sSqlStatement);
                                    }
                                    catch
                                    {
                                        sReturn = "Incorrect SQL statement." + System.Environment.NewLine;
                                        GlobalClass.LogUserAction(-2, iOperatorID, "ERROR! User [#UserName#] Execute " + sOperation + " on the table " + sTable, sSqlStatement);
                                        
                                    }
                                }
                                else
                                    sReturn = sReturn + "SQL statement is empty" + System.Environment.NewLine;
                            }       
                        }
                        else
                            sReturn = sReturn + "SQL statement is empty. Can not identify Primary Key." + System.Environment.NewLine;
                    }
                }
            }
            return sReturn;
        }
        
        public static string GenerateInsertStatement(string sTableView, Object ObjectType, List<ParameterClass> ParameterList)
        {
            string sReturn = string.Empty;

            if (sTableView != string.Empty)
            {
                string sTable = sTableView.Replace("dbo.", "dbo.tbl").Trim();
                string sInsertFieldsSql = string.Empty;
                string sInsertValuesSql = string.Empty;

                string sPrimaryKeyColumn = string.Empty;
                string sPrimaryKeyValue = string.Empty;
                //List<ParameterClass> ParameterList = new List<ParameterClass>();

                string sSQL = "SELECT * FROM dbo.fn_TableColumnInfo('" + sTable + "') ORDER BY 2";
                DataTable ColumnInfoDataTable = GlobalClass.GetDataTable(sTableView, sSQL);

                foreach (DataRow dr in ColumnInfoDataTable.Rows)
                {
                    string sCurrentFieldType = (string)dr["TypeName"];
                    string sCurrentField = (string)dr["ColumnName"];
                    int iCurrentFieldLength = 0;
                    int iCurrentFieldScale = 0;

                    if (dr["ColumnPrecision"] != DBNull.Value)
                        iCurrentFieldLength = (Int16)dr["ColumnPrecision"];

                    if (dr["ColumnScale"] != DBNull.Value)
                        iCurrentFieldScale = (int)dr["ColumnScale"];

                    string sCurrentValue = string.Empty;

                    Type myType = ObjectType.GetType();
                    PropertyInfo Property = myType.GetProperty(sCurrentField);

                    if (Property != null)
                    {
                        try
                        {
                            if (sCurrentFieldType == "nvarchar")
                            {
                                sCurrentValue =  GlobalClass.FormatStringValue(Convert.ToString(Property.GetValue(ObjectType, null)), iCurrentFieldLength);
                                if ((sCurrentValue == "Off") || (sCurrentValue == string.Empty))
                                    sCurrentValue = "NULL";
                                else
                                    sCurrentValue = "N'" + sCurrentValue + "'";
                            }
                            else if (sCurrentFieldType == "int")
                            {
                                int iValue = Convert.ToInt32(Property.GetValue(ObjectType, null));
                                sCurrentValue = GlobalClass.FormatIntegerValue(iValue);
                            }
                            else if (sCurrentFieldType == "numeric")
                            {
                                decimal dValue = Convert.ToDecimal(Property.GetValue(ObjectType, null));
                                sCurrentValue = Math.Round(dValue, iCurrentFieldScale).ToString().Trim();
                            }
                            else if (sCurrentFieldType == "decimal")
                            {
                                decimal dValue = Convert.ToDecimal(Property.GetValue(ObjectType, null));
                                sCurrentValue = Math.Round(dValue, iCurrentFieldScale).ToString().Trim();
                            }
                            else if (sCurrentFieldType == "float")
                            {
                                iCurrentFieldScale = 7;
                                decimal dValue = Convert.ToDecimal(Property.GetValue(ObjectType, null));
                                if (dValue == 0)
                                    sCurrentValue = "NULL";
                                else
                                    sCurrentValue = Math.Round(dValue, iCurrentFieldScale).ToString().Trim();
                            }
                            else if (sCurrentFieldType == "bit")
                            {
                                if (Convert.ToBoolean(Property.GetValue(ObjectType, null)) == true)
                                    sCurrentValue = "1";
                                else
                                    sCurrentValue = "0";
                            }
                            else if (sCurrentFieldType == "datetime")
                                sCurrentValue = "'" + Convert.ToDateTime(Property.GetValue(ObjectType, null)).ToString().Trim() + "'";
                            else if (sCurrentFieldType == "image")
                            {
                                ParameterClass Parameter = new ParameterClass();
                                Parameter.Parameter = "@" + sCurrentField;
                                Parameter.Value = (byte[])Property.GetValue(ObjectType, null);
                                if (Parameter.Value != null)
                                {
                                    sCurrentValue = "@" + sCurrentField;
                                    ParameterList.Add(Parameter);
                                }
                            }
                        }
                        catch
                        {
                            MessageBox.Show("Can not convert value " + sCurrentField + " Table " + sTableView, "Conversion Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }

                        if ((bool)dr["PrimaryKeyColumn"] == true)
                        {
                            sPrimaryKeyColumn = sCurrentField;
                            sPrimaryKeyValue = sCurrentValue;
                        }

                        if ((int)dr["IsIdentity"] == 0)
                        {
                            if ((sCurrentField != string.Empty) && (sCurrentValue != string.Empty))
                            {
                                sInsertFieldsSql = sInsertFieldsSql + sCurrentField + ",";
                                sInsertValuesSql = sInsertValuesSql + sCurrentValue + ",";
                            }
                        }
                    }
                }

                sInsertFieldsSql = sInsertFieldsSql.TrimEnd(',');
                sInsertValuesSql = sInsertValuesSql.TrimEnd(',');

                if ((sInsertFieldsSql != string.Empty) && (sInsertValuesSql != string.Empty) &&
                    (sPrimaryKeyColumn != string.Empty) && (sPrimaryKeyValue != string.Empty))
                {
                    string sInsertSql = "INSERT INTO " + sTableView + " (" + sInsertFieldsSql + ") VALUES (" + sInsertValuesSql + ")";
                    sInsertSql = sInsertSql.Replace("'0001-01-01 00:00:00'", "NULL") + ";" + System.Environment.NewLine;

                    DataTable dataTable = GlobalClass.GetDataTable(sTableView, "SELECT * FROM " + sTableView + " WHERE " + sPrimaryKeyColumn + " = " + sPrimaryKeyValue);
                    if (dataTable != null)
                        if (dataTable.Rows.Count == 0)
                            sReturn = sInsertSql.Trim();
                }
            }

            return sReturn;
        }

        public static string GenerateUpdateStatement(string sTableView, Object ObjectType, List<ParameterClass> ParameterList)
        {
            string sReturn = string.Empty;

            if (sTableView != string.Empty)
            {
                string sTable = sTableView.Replace("dbo.", "dbo.tbl").Trim();
                string sPrimaryKeyColumn = string.Empty;
                string sPrimaryKeyValue = string.Empty;

                string sUpdateSql = string.Empty;
                DataTable ColumnInfoDataTable = GlobalClass.GetDataTable(sTableView, "SELECT * FROM dbo.fn_TableColumnInfo('" + sTable + "') ORDER BY 2");

                foreach (DataRow dr in ColumnInfoDataTable.Rows)
                {
                    string sCurrentFieldType = (string)dr["TypeName"];
                    string sCurrentField = (string)dr["ColumnName"];
                    int iCurrentFieldLength = 0;
                    int iCurrentFieldScale = 0;

                    if (dr["ColumnPrecision"] != DBNull.Value)
                        iCurrentFieldLength = (Int16)dr["ColumnPrecision"];
                    if (dr["ColumnScale"] != DBNull.Value)
                        iCurrentFieldScale = (int)dr["ColumnScale"];

                    string sCurrentValue = string.Empty;

                    Type myType = ObjectType.GetType();
                    PropertyInfo Property = myType.GetProperty(sCurrentField);

                    if (Property != null)
                    {
                        try
                        {
                            if (sCurrentFieldType == "nvarchar")
                            {
                                //sCurrentValue = "'" + GlobalClass.FormatStringValue(Convert.ToString(Property.GetValue(ObjectType, null)), iCurrentFieldLength) + "'";

                                sCurrentValue = GlobalClass.FormatStringValue(Convert.ToString(Property.GetValue(ObjectType, null)), iCurrentFieldLength);
                                if ((sCurrentValue == "Off") || (sCurrentValue == string.Empty))
                                    sCurrentValue = "NULL";
                                else
                                    sCurrentValue = "N'" + sCurrentValue + "'";
                            }
                            else if (sCurrentFieldType == "int")
                            {
                                //sCurrentValue = Convert.ToInt32(Property.GetValue(ObjectType, null)).ToString().Trim();

                                int iValue = Convert.ToInt32(Property.GetValue(ObjectType, null));
                                sCurrentValue = GlobalClass.FormatIntegerValue(iValue);

                            }
                            else if (sCurrentFieldType == "numeric")
                            {
                                decimal dValue = Convert.ToDecimal(Property.GetValue(ObjectType, null));
                                sCurrentValue = Math.Round(dValue, iCurrentFieldScale).ToString().Trim();
                            }
                            else if (sCurrentFieldType == "decimal")
                            {
                                decimal dValue = Convert.ToDecimal(Property.GetValue(ObjectType, null));
                                sCurrentValue = Math.Round(dValue, iCurrentFieldScale).ToString().Trim();
                            }
                            else if (sCurrentFieldType == "float")
                            {
                                iCurrentFieldScale = 7;
                                decimal dValue = Convert.ToDecimal(Property.GetValue(ObjectType, null));
                                if (dValue == 0)
                                    sCurrentValue = "NULL";
                                else
                                    sCurrentValue = Math.Round(dValue, iCurrentFieldScale).ToString().Trim();
                            }

                            else if (sCurrentFieldType == "bit")
                            {
                                if (Convert.ToBoolean(Property.GetValue(ObjectType, null)) == true)
                                    sCurrentValue = "1";
                                else
                                    sCurrentValue = "0";
                            }
                            else if (sCurrentFieldType == "datetime")
                                sCurrentValue = "'" + Convert.ToDateTime(Property.GetValue(ObjectType, null)).ToString().Trim() + "'";
                            else if (sCurrentFieldType == "image")
                            {
                                ParameterClass Parameter = new ParameterClass();
                                Parameter.Parameter = "@" + sCurrentField;
                                Parameter.Value = (byte[])Property.GetValue(ObjectType, null);
                                if (Parameter.Value != null)
                                {
                                    sCurrentValue = "@" + sCurrentField;
                                    ParameterList.Add(Parameter);
                                }
                            }
                        }
                        catch
                        {
                            MessageBox.Show("Can not convert value " + sCurrentField + " Table " + sTableView, "Conversion Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }

                        if ((bool)dr["PrimaryKeyColumn"] == true)
                        {
                            sPrimaryKeyColumn = sCurrentField;
                            sPrimaryKeyValue = sCurrentValue;
                        }
                        else
                        {
                            if ((sCurrentField != string.Empty) && (sCurrentValue != string.Empty))
                                sUpdateSql = sUpdateSql + sCurrentField + " = " + sCurrentValue + ",";
                        }
                    }
                }
                
                sUpdateSql = sUpdateSql.TrimEnd(',');

                if ((sUpdateSql != string.Empty) && 
                    (sPrimaryKeyColumn != string.Empty) && (sPrimaryKeyValue != string.Empty))
                {
                    sUpdateSql = "UPDATE " + sTableView + " SET " + sUpdateSql + " WHERE " + sPrimaryKeyColumn + " = " + sPrimaryKeyValue;
                    sUpdateSql = sUpdateSql.Replace("'0001-01-01 00:00:00'", "NULL") + ";" + System.Environment.NewLine;

                    DataTable dataTable = GlobalClass.GetDataTable(sTableView, "SELECT * FROM " + sTableView + " WHERE " + sPrimaryKeyColumn + " = " + sPrimaryKeyValue);
                    if (dataTable != null)
                        if (dataTable.Rows.Count == 1)
                            sReturn = sUpdateSql.Trim();
                }
            }
            return sReturn;
        }

        public static void CopyListViewToClipboard(ListView lv)
        {
            string sTab = "\t";
            string sNewLine = "\n";
            string sCarriageReturn = "\r";

            StringBuilder buffer = new StringBuilder();
            for (int i = 0; i < lv.Columns.Count; i++)
            {
                string sLine = lv.Columns[i].Text;
                sLine = sLine.Replace(sTab, string.Empty);
                sLine = sLine.Replace(sNewLine, string.Empty);
                sLine = sLine.Replace(sCarriageReturn, string.Empty);

                buffer.Append(sLine);
                buffer.Append(sTab);
            }
            buffer.Append(sNewLine);

            ListView.SelectedListViewItemCollection SelectedRecords = lv.SelectedItems;

            if (SelectedRecords != null)
            {
                foreach (ListViewItem RecordItem in SelectedRecords)
                {
                    for (int j = 0; j < lv.Columns.Count; j++)
                    {
                        string sLine = RecordItem.SubItems[j].Text;
                        sLine = sLine.Replace(sTab, string.Empty);
                        sLine = sLine.Replace(sNewLine, string.Empty);
                        sLine = sLine.Replace(sCarriageReturn, string.Empty);

                        buffer.Append(sLine);
                        buffer.Append(sTab);
                    }
                    buffer.Append(sNewLine);
                }
            }

            Clipboard.SetText(buffer.ToString());
        }

        public static string FormatMergeString(string sString)
        {
            sString = sString.Trim();

            if (sString == string.Empty)
                sString = " ";

            return sString;
        }

        public static string FormatStringForFileName(string sString)
        {
            sString = sString.Trim().Replace("/", "_").Replace("%20", "").Replace("%26", "").Replace(" ", "").Replace("&", "");

            return sString;
        }

        public static string FormatStringWithSuperScript(string sString)
        {
            string sGamma = ((char)947).ToString();
            string sSS = ((char)8319).ToString();

            sString = sString.Trim().Replace("&gamma;", sGamma);
            sString = sString.Trim().Replace("&^;", sSS);

            
            return sString;
        }

        public static string GetIDEAYear(string sYearCode)
        {
            string sReturn = string.Empty;
                
            int iYearCode = -1;

            if (int.TryParse(sYearCode, out iYearCode) == true)
                sReturn = (1960 + iYearCode).ToString();

            return sReturn;
        }

        public static string GetCellValue(Excel.Worksheet o, string sRange)
        {
            string sReturn = string.Empty;

            Excel.Range r = (Excel.Range)o.get_Range(sRange, sRange);
            if (r.Value2 != null)
                sReturn = r.Value2.ToString().Trim().Replace("\r", "\r\n");

            return sReturn;
        }

        public static void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Unable to release the Object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }

        public static bool IsNumeric(string s)
        {    
            double Result;    
            return double.TryParse(s, out Result);  // TryParse routines were added in Framework version 2.0.
        }

        public static string ColumnDistinctValues(ListView lv, int iColumnIndex)
        {
            string sReturn = string.Empty;

            foreach (ListViewItem lvItem in lv.Items)
            {
                string sValue = lvItem.SubItems[iColumnIndex].Text;
                if (sValue.Trim() != string.Empty)
                    if (sReturn.Contains(sValue + "|") == false)
                        sReturn = sReturn + lvItem.SubItems[iColumnIndex].Text + "|";
            }

            return sReturn;
        }

        public static bool PrinterExists(string sPrinterName)
        {
            bool bReturn = false;

            foreach (String Printer in PrinterSettings.InstalledPrinters)
            {
                if (Printer == sPrinterName)
                {
                    bReturn = true;
                    break;
                }
            }
            return bReturn;
        }

        public static void PrintDocument(string sFileName)
        {
            if (File.Exists(sFileName))
            {
                Process printJob = new Process();
                printJob.StartInfo.FileName = sFileName;
                printJob.StartInfo.UseShellExecute = true;
                printJob.StartInfo.CreateNoWindow = true;
                printJob.StartInfo.Verb = "print";
                printJob.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
                printJob.Start();
            }
        }

        public static void PrintDocument(string sFileName, string sPrinterName)
        {
            if (PrinterExists(sPrinterName))
            {
                if (String.IsNullOrEmpty(sFileName)) throw new ArgumentNullException("sFileName");
                if (String.IsNullOrEmpty(sPrinterName)) throw new ArgumentNullException("sPrinterName");
                if (File.Exists(sFileName))
                {
                    Process p = new Process();
                    p.StartInfo.FileName = sFileName;
                    p.StartInfo.Verb = "PrintTo";
                    p.StartInfo.CreateNoWindow = true;
                    p.StartInfo.Arguments = '"' + sPrinterName + '"';
                    p.StartInfo.UseShellExecute = true;
                    p.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
                    p.Start();
                }
                else
                    throw new FileNotFoundException("File does not exist!", sFileName);
            }
            else
                throw new FileNotFoundException("Printer does not exist!", sPrinterName);
        }

        public static DateTime DateValueFromObject(DateTimePicker dObject)
        {
            DateTime dReturn = DateTime.MinValue;

            if (dObject.CustomFormat != " ")
            {
                if (DateTime.TryParse(dObject.Value.ToString(), out dReturn) == false)
                    dReturn = DateTime.MinValue;
            }

            return dReturn;
        }

        public static void DateValueToObject(DateTimePicker dObject, DateTime dValue)
        {
            if (dValue != DateTime.MinValue)
            {
                dObject.ValueChanged -= null ;
                dObject.Value = dValue;
                dObject.CustomFormat = "yyyy-MM-dd"; // Set Format to yyyy-MM-dd
            }
            else
            {
                dObject.Value = DateTime.Now;
                dObject.CustomFormat = " "; // One blank string.
            }
        }

        public static List<string[]> ParseCSV(string sFilePath, string sSection, char sSeparator)
        {
            List<string[]> parsedData = new List<string[]>();

            if (File.Exists(sFilePath))
            {
                try
                {
                    using (StreamReader readFile = new StreamReader(sFilePath))
                    {
                        string sLine;

                        while ((sLine = readFile.ReadLine()) != null)
                        {
                            if (sLine.Trim() != string.Empty)
                            {
                                string[] sRowLine = sLine.Trim().Split(sSeparator);
                                if (sRowLine.Length >= 1)
                                {
                                    for (int i = 0; i < sRowLine.Length; i++)
                                        sRowLine[i] = sRowLine[i].Trim();

                                    if (sRowLine[0].Trim() == sSection.Trim())
                                        parsedData.Add(sRowLine);
                                }
                            }
                        }
                    }
                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message);
                }
            }
            return parsedData;
        }

        public static string Translation(List<string[]> parsedData, int iItemID)
        {
            string sReturn = string.Empty;

            foreach (string[] sRowLine in parsedData)
            {
                if (sRowLine.Length == 4)
                {
                    int iRowItemID = -1;
                    if (int.TryParse(sRowLine[1], out iRowItemID) == true)
                    {
                        if (iRowItemID == iItemID)
                        {
                            sReturn = sRowLine[3];
                            break;
                        }
                    }
                }
            }
            return sReturn;
        }

        public static string Translation(List<string[]> parsedData, string sItemCode)
        {
            string sReturn = string.Empty;

            foreach (string[] sRowLine in parsedData)
            {
                if (sRowLine.Length == 4)
                {
                    if (sRowLine[3].Trim() == sItemCode.Trim())
                    {
                        sReturn = sRowLine[3];
                        break;
                    }
                }
            }
            return sReturn;
        }

        public static string SetPropertyValue(Object ObjectType, PropertyInfo Property, string sCurrentValue)
        {
            string sReturn = string.Empty;

            if (Property != null)
            {
                if (Property.CanWrite == true)
                {
                    try
                    {
                        // Use the SetValue method to change the caption.
                        if (Property.PropertyType.Name == "Int32")
                        {
                            if (sCurrentValue == "")
                                sCurrentValue = "0";
                            if (sCurrentValue == "Off")
                                sCurrentValue = "-1";

                            Property.SetValue(ObjectType, Convert.ToInt32(sCurrentValue), null);
                        }
                        else if (Property.PropertyType.Name == "Int16")
                        {
                            if (sCurrentValue == "")
                                sCurrentValue = "0";
                            if (sCurrentValue == "Off")
                                sCurrentValue = "-1";

                            Property.SetValue(ObjectType, Convert.ToInt16(sCurrentValue), null);
                        }
                        else if (Property.PropertyType.Name == "Decimal")
                        {
                            if (sCurrentValue == "")
                                sCurrentValue = "0.00";
                            if (sCurrentValue == "Off")
                                sCurrentValue = "-1";

                            Property.SetValue(ObjectType, Convert.ToDecimal(sCurrentValue), null);
                        }
                        else if (Property.PropertyType.Name == "Double")
                        {
                            if (sCurrentValue == "")
                                sCurrentValue = "0.00";
                            if (sCurrentValue == "Off")
                                sCurrentValue = "-1";
                            Double dValue = 0.00;

                            if (Double.TryParse(sCurrentValue, System.Globalization.NumberStyles.Any, null, out dValue))
                            {
                                Property.SetValue(ObjectType, dValue, null);
                                dValue = Convert.ToDouble(Property.GetValue(ObjectType, null));
                            }
                        }
                        else if (Property.PropertyType.Name == "String")
                        {
                            //if (sCurrentValue == "Off")
                            //    sCurrentValue = string.Empty;
                            sCurrentValue = sCurrentValue.Replace("\r", "\r\n");
                            //sCurrentValue = sCurrentValue.Replace("'", "`");
                            Property.SetValue(ObjectType, Convert.ToString(sCurrentValue.Trim()), null);
                        }
                        else if (Property.PropertyType.Name == "Boolean")
                        {
                            if ((sCurrentValue == "Off") || ((sCurrentValue == "")))
                                sCurrentValue = "False";
                            else
                            {
                                if (sCurrentValue == "Yes")
                                    sCurrentValue = "True";
                                else if (sCurrentValue == "No")
                                    sCurrentValue = "False";
                            }

                            Property.SetValue(ObjectType, Convert.ToBoolean(sCurrentValue.Trim()), null);
                        }
                        else if (Property.PropertyType.Name == "DateTime")
                        {
                            DateTime dCurrentValue = DateTime.MinValue;

                            if (sCurrentValue == "")
                                dCurrentValue = DateTime.MinValue;
                            else
                            {
                                double excelDate;
                                
                                if (double.TryParse(sCurrentValue, out excelDate) == true)
                                    dCurrentValue = GlobalClass.ConvertToDateTime(excelDate);
                                else
                                {
                                    if (sCurrentValue.Substring(4, 1) == "-")
                                        dCurrentValue = GlobalClass.FormatIAEAStringDateTimeValue(sCurrentValue); // yyyy-mm-dd >= 1800
                                    else
                                        dCurrentValue = GlobalClass.FormatTLDStringDateTimeValue(sCurrentValue); // dd/mm/yyyy >= 1800
                                }
                            }

                            Property.SetValue(ObjectType, dCurrentValue, null);
                        }
                    }
                    catch
                    {
                        sReturn = sReturn + "Can not set Property Value [" + Property.Name + " = " + sCurrentValue + "]" + System.Environment.NewLine;
                    }
                }
                else
                    sReturn = sReturn + "Property [" + Property.Name + "] is Read Only" + System.Environment.NewLine;
            }
            else
                sReturn = sReturn + "Property is Empty" + System.Environment.NewLine;

            return sReturn;
        }

        public static string GetMemberType(string sParticipationType)
        {
            string sReturn = string.Empty;

            if (sParticipationType == GlobalClass.sParticipationTypeSSDL)
                sReturn = "MN";
            else if (sParticipationType == GlobalClass.sParticipationTypeReference)
                sReturn = "R";
            else if (sParticipationType == GlobalClass.sParticipationTypePrimary)
                sReturn = "A";
            else if (sParticipationType == GlobalClass.sParticipationTypeHospitals)
                sReturn = "H";
            else
                sReturn = sParticipationType;

            return sReturn;
        }

        public static string GetParticipationType(string sMemberType)
        {
            string sReturn = string.Empty;

            if ("MN".Contains(sMemberType))
                sReturn = GlobalClass.sParticipationTypeSSDL;
            else if(sMemberType == "H")
                sReturn = GlobalClass.sParticipationTypeHospitals;
            else if(sMemberType == "R")
                sReturn = GlobalClass.sParticipationTypeReference;
            else if(sMemberType == "A")
                sReturn = GlobalClass.sParticipationTypePrimary;

            return sReturn;
        }


        public static string GetFirstElemets(int iElementNo, string sSelectedItems, string sSelectedType, char sSeparator)
        {
            string sReturn = string.Empty;
            int iCount = 0;
            foreach (string sItem in sSelectedItems.Split(sSeparator))
            {
                if (sItem != string.Empty)
                {
                    if (iCount < iElementNo)
                    {
                        string sItemCode = string.Empty;
                        if (sSelectedType == "Country")
                            sItemCode = sItem;
                        else if (sSelectedType == "Region")
                            sItemCode = sItem;
                        else if (sSelectedType == "RegionCode")
                            sItemCode = sItem;
                        else if (sSelectedType == "Batch")
                        {
                            BatchClass Batch = GlobalClass.Manager.GetBatch(Convert.ToInt32(sItem));
                            if (Batch != null)
                                sItemCode = Batch.BatchNo;
                        }
                        else if (sSelectedType == "AuditType")
                            sItemCode = sItem;
                        else if (sSelectedType == "ParticipationType")
                            sItemCode = sItem;
                        else if (sSelectedType == "Variable")
                            sItemCode = sItem;
                        else
                            sItemCode = sItem;

                        if (sItemCode != string.Empty)
                            sReturn = sReturn + sItemCode + ",";
                    }
                    iCount = iCount + 1;
                }
            }
            sReturn = sReturn.TrimEnd(',');

            if (iCount > iElementNo)
                sReturn = sReturn + "...";

            return sReturn;
        }

        public static void RefreshControlValues(Control BaseControl, SSDLBaseClass BaseObject)
        {
            foreach (Control c in BaseControl.Controls)
            {
                if ((c is TextBox) || (c is CheckBox) || (c is ComboBox) || (c is RadioButton) || (c is DateTimePicker))
                {
                    if (c.Tag != null)
                    {
                        string sFieldName = string.Empty;
                        string sFieldValue = string.Empty;

                        string sTag = c.Tag.ToString();
                        int iIndex = sTag.IndexOf(":");

                        if (iIndex > -1)
                        {
                            sFieldName = sTag.Substring(0, sTag.IndexOf(":"));
                            sFieldValue = sTag.Substring(sTag.IndexOf(":") + 1);
                        }
                        else
                            sFieldName = sTag;

                        if (BaseObject is TLDSetClass)
                            sFieldName = sFieldName.Replace("TLDSet.", string.Empty);
                        else if (BaseObject is TLDDataSheetClass)
                            sFieldName = sFieldName.Replace("TLDSet.TLDDataSheet.", string.Empty);
                        else if (BaseObject is TLDEvaluationClass)
                            sFieldName = sFieldName.Replace("TLDSet.Evaluation.", string.Empty);
                        else if (BaseObject is TLDCertificateClass)
                            sFieldName = sFieldName.Replace("TLDSet.TLDCertificate.", string.Empty);


                        // Get the type and PropertyInfo.
                        Type myType = BaseObject.GetType();
                        PropertyInfo Property = myType.GetProperty(sFieldName.Trim());

                        try
                        {
                            if (Property != null)
                            {
                                if (c is TextBox)
                                {
                                    (c as TextBox).Text = Convert.ToString(Property.GetValue(BaseObject, null));
                                    if ((c as TextBox).Text == "N/A")
                                        (c as TextBox).Text = string.Empty;


                                    if (sFieldValue != string.Empty) // realy Format
                                        (c as TextBox).Text = String.Format(sFieldValue, Property.GetValue(BaseObject, null));
                                    else
                                    {
                                        if (Property.PropertyType.Name == "Double")
                                            (c as TextBox).Text = String.Format("{0:0.00}", Property.GetValue(BaseObject, null));
                                        else if (Property.PropertyType.Name == "Decimal")
                                            (c as TextBox).Text = String.Format("{0:0.00}", Property.GetValue(BaseObject, null));
                                        else if (Property.PropertyType.Name == "Int32")
                                            (c as TextBox).Text = String.Format("{0:0}", Property.GetValue(BaseObject, null));
                                        else if (Property.PropertyType.Name == "Int64")
                                            (c as TextBox).Text = String.Format("{0:0}", Property.GetValue(BaseObject, null));
                                        else
                                            (c as TextBox).Text = Convert.ToString(Property.GetValue(BaseObject, null));
                                    }

                                }
                                else if (c is ComboBox)
                                {
                                    List<DictionaryTypeClass> DictionaryList = GlobalClass.Dictionary.GetDictionary(sFieldName);
                                    if (DictionaryList.Count > 0)
                                    {
                                        (c as ComboBox).DataSource = null;
                                        (c as ComboBox).DataSource = DictionaryList;
                                        (c as ComboBox).DisplayMember = "Description";

                                        if (Property.PropertyType.Name == "Int16")
                                        {
                                            (c as ComboBox).ValueMember = "ItemID";
                                            (c as ComboBox).SelectedValue = Convert.ToInt16(Property.GetValue(BaseObject, null));
                                        }
                                        else if (Property.PropertyType.Name == "Int32")
                                        {
                                            (c as ComboBox).ValueMember = "ItemID";
                                            (c as ComboBox).SelectedValue = Convert.ToInt32(Property.GetValue(BaseObject, null));
                                        }
                                        else if (Property.PropertyType.Name == "Int64")
                                        {
                                            (c as ComboBox).ValueMember = "ItemID";
                                            (c as ComboBox).SelectedValue = Convert.ToInt64(Property.GetValue(BaseObject, null));
                                        }
                                        else if (Property.PropertyType.Name == "String")
                                        {
                                            (c as ComboBox).ValueMember = "ItemCode";
                                            (c as ComboBox).SelectedValue = Convert.ToString(Property.GetValue(BaseObject, null));
                                        }
                                    }
                                    else
                                    {
                                        List<EquipmentTypeClass> EquipmentList = GlobalClass.Dictionary.GetEquipment(sFieldName);
                                        if (EquipmentList.Count > 0)
                                        {
                                            (c as ComboBox).DataSource = null;
                                            (c as ComboBox).DataSource = EquipmentList;
                                            (c as ComboBox).DisplayMember = "Description";
                                            (c as ComboBox).ValueMember = "ItemCode";
                                            (c as ComboBox).SelectedValue = Convert.ToString(Property.GetValue(BaseObject, null));
                                        }
                                    }
                                }
                                else if (c is RadioButton)
                                {
                                    if (Property.PropertyType.Name == "Boolean")
                                    {
                                        if (GlobalClass.FormatBooleanValue(Convert.ToInt16(sFieldValue)) == Convert.ToString(Property.GetValue(BaseObject, null)))
                                            (c as RadioButton).Checked = true;
                                    }
                                    else
                                    {
                                        if (sFieldValue == Convert.ToString(Property.GetValue(BaseObject, null)))
                                            (c as RadioButton).Checked = true;
                                    }
                                }

                                else if (c is CheckBox)
                                {
                                    if (Property.PropertyType.Name == "String")
                                        (c as CheckBox).Checked = Convert.ToString(Property.GetValue(BaseObject, null)) == sFieldValue;
                                    else if (Property.PropertyType.Name == "Boolean")
                                        (c as CheckBox).Checked = Convert.ToBoolean(Property.GetValue(BaseObject, null));
                                    else
                                        (c as CheckBox).Checked = Convert.ToBoolean(Property.GetValue(BaseObject, null));
                                }
                                else if (c is DateTimePicker)
                                {
                                    if (Convert.ToDateTime(Property.GetValue(BaseObject, null)) != DateTime.MinValue)
                                    {
                                        (c as DateTimePicker).CustomFormat = "yyyy-MM-dd"; // Set Format to yyyy-MM-dd
                                        (c as DateTimePicker).Value = Convert.ToDateTime(Property.GetValue(BaseObject, null));
                                    }
                                    else
                                        (c as DateTimePicker).CustomFormat = " "; // One blank string.
                                }
                            }
                        }
                        catch
                        { }
                    }
                }
                else
                {
                    RefreshControlValues(c, BaseObject);
                }
            }
        }

        public static void SaveControlValues(Control BaseControl, SSDLBaseClass BaseObject)
        {
            foreach (Control c in BaseControl.Controls)
            {
                if ((c is TextBox) || (c is CheckBox) || (c is ComboBox) || (c is RadioButton) || (c is DateTimePicker))
                {
                    if (c.Tag != null)
                    {
                        string sFieldName = string.Empty;
                        string sFieldValue = string.Empty;

                        string sTag = c.Tag.ToString();
                        int iIndex = sTag.IndexOf(":");

                        if (iIndex > -1)
                        {
                            sFieldName = sTag.Substring(0, sTag.IndexOf(":"));
                            //sFieldValue = sTag.Substring(sTag.IndexOf(":") + 1);
                        }
                        else
                            sFieldName = sTag;

                        if (BaseObject is TLDSetClass)
                            sFieldName = sFieldName.Replace("TLDSet.", string.Empty);
                        else if (BaseObject is TLDDataSheetClass)
                            sFieldName = sFieldName.Replace("TLDSet.TLDDataSheet.", string.Empty);
                        else if (BaseObject is TLDEvaluationClass)
                            sFieldName = sFieldName.Replace("TLDSet.Evaluation.", string.Empty);
                        else if (BaseObject is TLDCertificateClass)
                            sFieldName = sFieldName.Replace("TLDSet.TLDCertificate.", string.Empty);

                        // Get the type and PropertyInfo.
                        Type BaseObjectType = BaseObject.GetType();
                        PropertyInfo Property = BaseObjectType.GetProperty(sFieldName.Trim());
                        
                        if (Property != null)
                        {
                            try
                            {
                                if (Property.CanWrite == true)
                                {
                                    if (c is TextBox)
                                    {
                                        sFieldValue = (c as TextBox).Text.Trim();
                                    }
                                    else if (c is ComboBox)
                                    {
                                        if ((c as ComboBox).SelectedValue != null)
                                            sFieldValue = (c as ComboBox).SelectedValue.ToString();
                                    }
                                    else if (c is RadioButton)
                                    {
                                        if ((c as RadioButton).Checked)
                                            sFieldValue = sTag.Substring(sTag.IndexOf(":") + 1);
                                        else
                                            continue;
                                    }
                                    else if (c is CheckBox)
                                    {
                                        sFieldValue = (c as CheckBox).Checked.ToString();
                                    }
                                    else if (c is DateTimePicker)
                                    {
                                        if ((c as DateTimePicker).CustomFormat != " ")
                                            sFieldValue = (c as DateTimePicker).Value.ToString();
                                        else
                                            sFieldValue = DateTime.MinValue.ToString();
                                    }

                                    string sErrorString = GlobalClass.SetPropertyValue(BaseObject, Property, sFieldValue.Trim());

                                }
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show("Error " + sFieldName + " - " + sFieldValue + ": " + ex.Message, "Saving results...", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                            }
                        }
                    }
                }
                else
                {
                    SaveControlValues(c, BaseObject);
                }
            }
        }


        public static string CompareObjects(SSDLBaseClass BaseObject1, SSDLBaseClass BaseObject2)
        {
            string sReturn = string.Empty;

            if (BaseObject1 != null)
            {
                if (BaseObject2 != null)
                {
                    Type BaseObject1Type = BaseObject1.GetType();
                    if (BaseObject2.GetType() == BaseObject1Type)
                    {
                        // This will only use public properties. Is that enough?
                        foreach (PropertyInfo Property in BaseObject1Type.GetProperties())
                        {
                            if (Property.CanRead)
                            {
                                object Value1 = Property.GetValue(BaseObject1, null);
                                object Value2 = Property.GetValue(BaseObject2, null);
                                if (!object.Equals(Value1, Value2))
                                {
                                    sReturn = sReturn + Property.Name + ",";
                                }
                            }
                        }
                        sReturn = sReturn.TrimEnd(',');
                    }
                }
            }
            return sReturn;
        }



        public static TypeSelection2Class EvaluationVariable(Excel.Worksheet ThisSheet, string VariableName, string sRange)
        {
            TypeSelection2Class Variable = new TypeSelection2Class();

            Variable.Type = VariableName;
            Variable.TypeDescription = GlobalClass.GetCellValue(ThisSheet, sRange);

            return Variable;
        }

        public static TypeSelection2Class EvaluationVariableDate(Excel.Worksheet ThisSheet, string VariableName, string sRange)
        {
            TypeSelection2Class Variable = new TypeSelection2Class();

            Variable.Type = VariableName;
            string sCurrentValue = GlobalClass.GetCellValue(ThisSheet, sRange);

            double excelDate;

            if (double.TryParse(sCurrentValue, out excelDate) == true)
            {
                sCurrentValue = GlobalClass.ConvertToDateTime(excelDate).ToString();
                Variable.TypeDescription = sCurrentValue;
            }

            return Variable;
        }

        public static List<TypeSelection2Class> LoadVariableListTLDDataSheetHospital(string sFileName)
        {
            List<TypeSelection2Class> VariableList = new List<TypeSelection2Class>();
            VariableList.Clear();

            if (File.Exists(sFileName))
            {
                bool isCobalt = false;

                string sBatchNo = string.Empty;
                string sSetNo = string.Empty;

                int iIndex = -1;
                double dValue = -1;
                string sValue = string.Empty;
                object oMissing = System.Reflection.Missing.Value;

                string sTranslationFilePath = GlobalClass.sApplicationStartupPath + "\\Translation.csv";

                if (File.Exists(sTranslationFilePath))
                {
                    List<string[]> TranslationEquipmentCo60 = GlobalClass.ParseCSV(sTranslationFilePath, "EquipmentCo60", ',');
                    List<string[]> TranslationEquipmentLinac = GlobalClass.ParseCSV(sTranslationFilePath, "EquipmentLinac", ',');

                    List<string[]> TranslationIonisationChamber = GlobalClass.ParseCSV(sTranslationFilePath, "IonisationChamber", ',');
                    List<string[]> TranslationElectrometer = GlobalClass.ParseCSV(sTranslationFilePath, "Electrometer", ',');


                    if (File.Exists(sFileName))
                    {
                        try
                        {
                            Excel.Application ThisApplication = new Excel.ApplicationClass();
                            ThisApplication.Visible = false;
                            Excel.Workbook ThisWorkBook = (Excel.Workbook)ThisApplication.Workbooks.Open(sFileName, oMissing, oMissing, 5, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
                            //Excel.Worksheet ThisSheet = (Excel.Worksheet)ThisWorkBook.Sheets[1]; //"Sheet1"

                            ThisWorkBook.Unprotect("dmrp");


                            try
                            {
                                foreach (Excel.Worksheet ThisSheet in ThisWorkBook.Application.Worksheets)
                                {
                                    ThisSheet.Visible = Microsoft.Office.Interop.Excel.XlSheetVisibility.xlSheetVisible;

                                    if ((ThisSheet.Name == "Datasheet Part1") || (ThisSheet.Name == "Hoja de Datos Seccion I"))
                                    {
                                        // Linac - "The treatment unit used for this audit is of the type"
                                        // Cobalt - "The Co-60 treatment unit used for this audit is of the type"

                                        isCobalt = GlobalClass.GetCellValue(ThisSheet, "B5").Contains("Co-60");

                                        if (isCobalt == true)
                                        {
                                            string sConditions = GlobalClass.GetCellValue(ThisSheet, "B56");

                                            if ((GlobalClass.GetCellValue(ThisSheet, "B57").Trim() != "0") && (GlobalClass.GetCellValue(ThisSheet, "B57").Trim() != ""))
                                                sConditions = sConditions + System.Environment.NewLine + GlobalClass.GetCellValue(ThisSheet, "B57");
                                            if ((GlobalClass.GetCellValue(ThisSheet, "B58").Trim() != "0") && (GlobalClass.GetCellValue(ThisSheet, "B58").Trim() != ""))
                                                sConditions = sConditions + System.Environment.NewLine + GlobalClass.GetCellValue(ThisSheet, "B58");
                                            if ((GlobalClass.GetCellValue(ThisSheet, "B59").Trim() != "0") && (GlobalClass.GetCellValue(ThisSheet, "B59").Trim() != ""))
                                                sConditions = sConditions + System.Environment.NewLine + GlobalClass.GetCellValue(ThisSheet, "B59");
                                            if ((GlobalClass.GetCellValue(ThisSheet, "B60").Trim() != "0") && (GlobalClass.GetCellValue(ThisSheet, "B60").Trim() != ""))
                                                sConditions = sConditions + System.Environment.NewLine + GlobalClass.GetCellValue(ThisSheet, "B60");

                                            VariableList.Add(new TypeSelection2Class("Conditions", sConditions));
                                        }
                                        else
                                        {
                                            string sConditions = GlobalClass.GetCellValue(ThisSheet, "B60");

                                            if ((GlobalClass.GetCellValue(ThisSheet, "B61").Trim() != "0") && (GlobalClass.GetCellValue(ThisSheet, "B61").Trim() != ""))
                                                sConditions = sConditions + System.Environment.NewLine + GlobalClass.GetCellValue(ThisSheet, "B61");
                                            if ((GlobalClass.GetCellValue(ThisSheet, "B62").Trim() != "0") && (GlobalClass.GetCellValue(ThisSheet, "B62").Trim() != ""))
                                                sConditions = sConditions + System.Environment.NewLine + GlobalClass.GetCellValue(ThisSheet, "B62");
                                            if ((GlobalClass.GetCellValue(ThisSheet, "B63").Trim() != "0") && (GlobalClass.GetCellValue(ThisSheet, "B63").Trim() != ""))
                                                sConditions = sConditions + System.Environment.NewLine + GlobalClass.GetCellValue(ThisSheet, "B63");

                                            VariableList.Add(new TypeSelection2Class("Conditions", sConditions));
                                        }
                                    }
                                    else if ((ThisSheet.Name == "Datasheet Part2") || (ThisSheet.Name == "Hoja de Datos Seccion II"))
                                    {
                                        if (isCobalt == true)
                                        {
                                            string sDetailedExplanations = GlobalClass.GetCellValue(ThisSheet, "B64");

                                            if ((GlobalClass.GetCellValue(ThisSheet, "B65").Trim() != "0") && (GlobalClass.GetCellValue(ThisSheet, "B65").Trim() != ""))
                                                sDetailedExplanations = sDetailedExplanations + System.Environment.NewLine + GlobalClass.GetCellValue(ThisSheet, "B65");
                                            if ((GlobalClass.GetCellValue(ThisSheet, "B66").Trim() != "0") && (GlobalClass.GetCellValue(ThisSheet, "B66").Trim() != ""))
                                                sDetailedExplanations = sDetailedExplanations + System.Environment.NewLine + GlobalClass.GetCellValue(ThisSheet, "B66");
                                            if ((GlobalClass.GetCellValue(ThisSheet, "B67").Trim() != "0") && (GlobalClass.GetCellValue(ThisSheet, "B67").Trim() != ""))
                                                sDetailedExplanations = sDetailedExplanations + System.Environment.NewLine + GlobalClass.GetCellValue(ThisSheet, "B67");
                                            if ((GlobalClass.GetCellValue(ThisSheet, "B68").Trim() != "0") && (GlobalClass.GetCellValue(ThisSheet, "B68").Trim() != ""))
                                                sDetailedExplanations = sDetailedExplanations + System.Environment.NewLine + GlobalClass.GetCellValue(ThisSheet, "B68");
                                            if ((GlobalClass.GetCellValue(ThisSheet, "B69").Trim() != "0") && (GlobalClass.GetCellValue(ThisSheet, "B69").Trim() != ""))
                                                sDetailedExplanations = sDetailedExplanations + System.Environment.NewLine + GlobalClass.GetCellValue(ThisSheet, "B69");
                                            if ((GlobalClass.GetCellValue(ThisSheet, "B70").Trim() != "0") && (GlobalClass.GetCellValue(ThisSheet, "B70").Trim() != ""))
                                                sDetailedExplanations = sDetailedExplanations + System.Environment.NewLine + GlobalClass.GetCellValue(ThisSheet, "B70");

                                            VariableList.Add(new TypeSelection2Class("DetailedExplanations", sDetailedExplanations));
                                        }
                                        else
                                        {
                                            string sDetailedExplanations = GlobalClass.GetCellValue(ThisSheet, "B57");

                                            if ((GlobalClass.GetCellValue(ThisSheet, "B58").Trim() != "0") && (GlobalClass.GetCellValue(ThisSheet, "B58").Trim() != ""))
                                                sDetailedExplanations = sDetailedExplanations + System.Environment.NewLine + GlobalClass.GetCellValue(ThisSheet, "B58");
                                            if ((GlobalClass.GetCellValue(ThisSheet, "B59").Trim() != "0") && (GlobalClass.GetCellValue(ThisSheet, "B59").Trim() != ""))
                                                sDetailedExplanations = sDetailedExplanations + System.Environment.NewLine + GlobalClass.GetCellValue(ThisSheet, "B59");
                                            if ((GlobalClass.GetCellValue(ThisSheet, "B60").Trim() != "0") && (GlobalClass.GetCellValue(ThisSheet, "B60").Trim() != ""))
                                                sDetailedExplanations = sDetailedExplanations + System.Environment.NewLine + GlobalClass.GetCellValue(ThisSheet, "B60");
                                            if ((GlobalClass.GetCellValue(ThisSheet, "B61").Trim() != "0") && (GlobalClass.GetCellValue(ThisSheet, "B61").Trim() != ""))
                                                sDetailedExplanations = sDetailedExplanations + System.Environment.NewLine + GlobalClass.GetCellValue(ThisSheet, "B61");
                                            if ((GlobalClass.GetCellValue(ThisSheet, "B62").Trim() != "0") && (GlobalClass.GetCellValue(ThisSheet, "B62").Trim() != ""))
                                                sDetailedExplanations = sDetailedExplanations + System.Environment.NewLine + GlobalClass.GetCellValue(ThisSheet, "B62");
                                            if ((GlobalClass.GetCellValue(ThisSheet, "B63").Trim() != "0") && (GlobalClass.GetCellValue(ThisSheet, "B63").Trim() != ""))
                                                sDetailedExplanations = sDetailedExplanations + System.Environment.NewLine + GlobalClass.GetCellValue(ThisSheet, "B63");

                                            VariableList.Add(new TypeSelection2Class("DetailedExplanations", sDetailedExplanations));
                                        }
                                    }
                                    else if (ThisSheet.Name == "Tables")
                                    {
                                        if (isCobalt == true)
                                            VariableList.Add(new TypeSelection2Class("SetBeamType", sBeamTypeCo60));
                                        else
                                            VariableList.Add(new TypeSelection2Class("SetBeamType", sBeamTypePhoton));

                                        sBatchNo = GlobalClass.GetCellValue(ThisSheet, "B94").Trim();
                                        sSetNo = GlobalClass.GetCellValue(ThisSheet, "B95").Trim();

                                        VariableList.Add(new TypeSelection2Class("BatchNo", sBatchNo));
                                        VariableList.Add(new TypeSelection2Class("SetNo", sSetNo));

                                        VariableList.Add(new TypeSelection2Class("AuditType", sAuditTypeRT));

                                        if (sSetNo.Substring(sSetNo.Length - 1, 1) != "R")
                                        {
                                            VariableList.Add(new TypeSelection2Class("SetType", "1")); // 1-FirstIrradiation | 2-FollowUp 

                                            if (sSetNo.Contains("P"))
                                                VariableList.Add(new TypeSelection2Class("ParticipationType", sParticipationTypePrimary));
                                            else if (sSetNo.Contains("R"))
                                                VariableList.Add(new TypeSelection2Class("ParticipationType", sParticipationTypeReference));
                                            else if (sSetNo.Contains("DL"))
                                                VariableList.Add(new TypeSelection2Class("ParticipationType", sParticipationTypeSSDL));
                                            else
                                                VariableList.Add(new TypeSelection2Class("ParticipationType", sParticipationTypeHospitals));
                                        }
                                        else
                                        {
                                            VariableList.Add(new TypeSelection2Class("SetType", "2")); // 1-FirstIrradiation | 2-FollowUp 
                                            if (sSetNo.Substring(0, sSetNo.Length - 1).Contains("P"))
                                                VariableList.Add(new TypeSelection2Class("ParticipationType", sParticipationTypePrimary));
                                            else if (sSetNo.Substring(0, sSetNo.Length - 1).Contains("R"))
                                                VariableList.Add(new TypeSelection2Class("ParticipationType", sParticipationTypeReference));
                                            else if (sSetNo.Substring(0, sSetNo.Length - 1).Contains("DL"))
                                                VariableList.Add(new TypeSelection2Class("ParticipationType", sParticipationTypeSSDL));
                                            else
                                                VariableList.Add(new TypeSelection2Class("ParticipationType", sParticipationTypeHospitals));
                                        }


                                        if (int.TryParse(GlobalClass.GetCellValue(ThisSheet, "B21"), out iIndex))
                                        {
                                            if (iIndex == 1)
                                                sValue = "Yes";
                                            else if (iIndex == 2)
                                                sValue = "No";
                                            else
                                                sValue = "Off";

                                            VariableList.Add(new TypeSelection2Class("PreviousParticipation", sValue));
                                        }

                                        if (isCobalt == true)
                                        {
                                            if (int.TryParse(GlobalClass.GetCellValue(ThisSheet, "B26"), out iIndex))
                                            {
                                                string sItemCode = GlobalClass.Translation(TranslationEquipmentCo60, iIndex);

                                                EquipmentTypeClass Equipment = GlobalClass.Dictionary.GetEquipmentItem(GlobalClass.Dictionary.GetEquipmentByGroup("RadionuclideTherapy"), sItemCode);

                                                if (Equipment != null)
                                                {
                                                    VariableList.Add(new TypeSelection2Class("Equipment", Equipment.ItemCode));
                                                    VariableList.Add(new TypeSelection2Class("EquipmentCo60", Equipment.ItemCode));
                                                }
                                            }
                                        }
                                        else
                                        {
                                            if (int.TryParse(GlobalClass.GetCellValue(ThisSheet, "B26"), out iIndex))
                                            {
                                                string sItemCode = GlobalClass.Translation(TranslationEquipmentLinac, iIndex);

                                                EquipmentTypeClass Equipment = GlobalClass.Dictionary.GetEquipmentItem(GlobalClass.Dictionary.GetEquipmentByGroup("LinearAccelerator"), sItemCode);

                                                if (Equipment != null)
                                                {
                                                    VariableList.Add(new TypeSelection2Class("Equipment", Equipment.ItemCode));
                                                    VariableList.Add(new TypeSelection2Class("EquipmentLinac", Equipment.ItemCode));
                                                }
                                            }
                                        }

                                        VariableList.Add(GlobalClass.EvaluationVariable(ThisSheet, "EquipmentOther", "B27"));

                                        if (int.TryParse(GlobalClass.GetCellValue(ThisSheet, "B61"), out iIndex))
                                        {
                                            string sItemCode = GlobalClass.Translation(TranslationIonisationChamber, iIndex);

                                            EquipmentTypeClass Equipment = GlobalClass.Dictionary.GetEquipmentItem(GlobalClass.Dictionary.GetEquipment("IonisationChamber"), sItemCode);

                                            if (Equipment != null)
                                                VariableList.Add(new TypeSelection2Class("IonisationChamber", Equipment.ItemCode));
                                        }

                                        if (int.TryParse(GlobalClass.GetCellValue(ThisSheet, "B62"), out iIndex))
                                        {
                                            string sItemCode = GlobalClass.Translation(TranslationElectrometer, iIndex);

                                            EquipmentTypeClass Equipment = GlobalClass.Dictionary.GetEquipmentItem(GlobalClass.Dictionary.GetEquipment("Electrometer"), sItemCode);

                                            if (Equipment != null)
                                                VariableList.Add(new TypeSelection2Class("Electrometer", Equipment.ItemCode));
                                        }

                                        if (int.TryParse(GlobalClass.GetCellValue(ThisSheet, "B90"), out iIndex))
                                        {
                                            EquipmentTypeClass Equipment = GlobalClass.Dictionary.GetEquipmentItem(GlobalClass.Dictionary.GetEquipment("DosimetryProtocol"), iIndex - 1);

                                            if (Equipment != null)
                                                VariableList.Add(new TypeSelection2Class("DosimetryProtocol", Equipment.ItemCode));
                                        }

                                        string sContactFamilyName = GlobalClass.GetCellValue(ThisSheet, "B3").Trim();
                                        string sContactPosition = GlobalClass.GetCellValue(ThisSheet, "B4").Trim();

                                        if (sContactFamilyName == "0") sContactFamilyName = string.Empty;
                                        if (sContactPosition == "0") sContactPosition = string.Empty;

                                        if ((sContactFamilyName == string.Empty) && (sContactPosition == string.Empty))
                                        {
                                            sContactFamilyName = GlobalClass.GetCellValue(ThisSheet, "B1").Trim();
                                            sContactPosition = GlobalClass.GetCellValue(ThisSheet, "B2").Trim();

                                            if (sContactFamilyName == "0") sContactFamilyName = string.Empty;
                                            if (sContactPosition == "0") sContactPosition = string.Empty;

                                            if ((sContactFamilyName != string.Empty) && (sContactPosition == string.Empty))
                                                sContactPosition = "Radiation oncologist";
                                        }
                                        else if ((sContactFamilyName != string.Empty) && (sContactPosition == string.Empty))
                                            sContactPosition = "Medical physicist";

                                        VariableList.Add(new TypeSelection2Class("ContactFamilyName", sContactFamilyName));
                                        VariableList.Add(new TypeSelection2Class("ContactPosition", sContactPosition));

                                        VariableList.Add(GlobalClass.EvaluationVariable(ThisSheet, "ContactDepartment", "B6"));
                                        VariableList.Add(GlobalClass.EvaluationVariable(ThisSheet, "ContactTelephone1", "B11"));
                                        VariableList.Add(GlobalClass.EvaluationVariable(ThisSheet, "ContactTelephone2", "B12"));
                                        VariableList.Add(GlobalClass.EvaluationVariable(ThisSheet, "ContactEmail", "B13"));


                                        VariableList.Add(GlobalClass.EvaluationVariable(ThisSheet, "OperatorName", "B5"));
                                        VariableList.Add(GlobalClass.EvaluationVariable(ThisSheet, "Country", "B7"));
                                        VariableList.Add(GlobalClass.EvaluationVariable(ThisSheet, "Street", "B8"));
                                        VariableList.Add(GlobalClass.EvaluationVariable(ThisSheet, "City", "B9"));
                                        VariableList.Add(GlobalClass.EvaluationVariable(ThisSheet, "State", "B10"));

                                        VariableList.Add(GlobalClass.EvaluationVariable(ThisSheet, "CompletedByPersonFamilyName", "B14"));
                                        VariableList.Add(new TypeSelection2Class("CompletedByPersonFirstName", string.Empty));

                                        DateTime dCompletedDate = GlobalClass.FormatIAEAStringDateTimeValue(
                                                             GlobalClass.GetIDEAYear(GlobalClass.GetCellValue(ThisSheet, "B18")) + "-" +
                                                                                     GlobalClass.GetCellValue(ThisSheet, "B17").PadLeft(2, '0') + "-" +
                                                                                     GlobalClass.GetCellValue(ThisSheet, "B16").PadLeft(2, '0'));

                                        VariableList.Add(new TypeSelection2Class("CompletedDate", dCompletedDate.ToString()));

                                        VariableList.Add(GlobalClass.EvaluationVariable(ThisSheet, "IrradiatedByPersonFamilyName", "B19"));
                                        VariableList.Add(new TypeSelection2Class("IrradiatedByPersonFirstName", string.Empty));

                                        if (int.TryParse(GlobalClass.GetCellValue(ThisSheet, "B20"), out iIndex))
                                        {
                                            DictionaryTypeClass DictionaryType = GlobalClass.Dictionary.GetDictionaryType(GlobalClass.Dictionary.GetDictionary("IrradiatedByPersonPosition"), iIndex - 1);
                                            if (DictionaryType != null)
                                                VariableList.Add(new TypeSelection2Class("IrradiatedByPersonPosition", DictionaryType.ItemCode));
                                        }


                                        if (int.TryParse(GlobalClass.GetCellValue(ThisSheet, "B25"), out iIndex))
                                        {
                                            DictionaryTypeClass DictionaryType = GlobalClass.Dictionary.GetDictionaryType(GlobalClass.Dictionary.GetDictionary("ParticipationOrganiser"), iIndex - 1);
                                            if (DictionaryType != null)
                                                VariableList.Add(new TypeSelection2Class("ParticipationOrganiser", DictionaryType.ItemCode));
                                        }

                                        if (int.TryParse(GlobalClass.GetCellValue(ThisSheet, "B24"), out iIndex))
                                        {
                                            DictionaryTypeClass DictionaryType = GlobalClass.Dictionary.GetDictionaryType(GlobalClass.Dictionary.GetDictionary("ParticipationYear"), iIndex);
                                            if (DictionaryType != null)
                                                VariableList.Add(new TypeSelection2Class("ParticipationYear", DictionaryType.ItemCode));
                                        }


                                        if (int.TryParse(GlobalClass.GetCellValue(ThisSheet, "B28"), out iIndex))
                                        {
                                            DictionaryTypeClass DictionaryType = GlobalClass.Dictionary.GetDictionaryType(GlobalClass.Dictionary.GetDictionary("EquipmentProductionYear"), iIndex);
                                            if (DictionaryType != null)
                                                VariableList.Add(new TypeSelection2Class("EquipmentProductionYear", DictionaryType.ItemCode));
                                        }

                                        if (int.TryParse(GlobalClass.GetCellValue(ThisSheet, "B29"), out iIndex))
                                        {
                                            DictionaryTypeClass DictionaryType = GlobalClass.Dictionary.GetDictionaryType(GlobalClass.Dictionary.GetDictionary("EquipmentInstallationYear"), iIndex);
                                            if (DictionaryType != null)
                                                VariableList.Add(new TypeSelection2Class("EquipmentInstallationYear", DictionaryType.ItemCode));
                                        }


                                        if (isCobalt == true)
                                        {
                                            //DateTime dEquipmentLastSourceReplacementDate = GlobalClass.FormatIAEAStringDateTimeValue(
                                            //                                               GlobalClass.GetIDEAYear(GlobalClass.GetCellValue(ThisSheet, "B32")) + "-" +
                                            //                                                                       GlobalClass.GetCellValue(ThisSheet, "B31").PadLeft(2, '0') + "-" +
                                            //                                                                       GlobalClass.GetCellValue(ThisSheet, "B30").PadLeft(2, '0'));

                                            VariableList.Add(new TypeSelection2Class("EquipmentLastSourceReplacementYear", GlobalClass.GetCellValue(ThisSheet, "B32").ToString()));
                                        }

                                        DateTime dIrradiationDate = GlobalClass.FormatIAEAStringDateTimeValue(
                                                                    GlobalClass.GetIDEAYear(GlobalClass.GetCellValue(ThisSheet, "B35")) + "-" +
                                                                                            GlobalClass.GetCellValue(ThisSheet, "B34").PadLeft(2, '0') + "-" +
                                                                                            GlobalClass.GetCellValue(ThisSheet, "B33").PadLeft(2, '0'));

                                        VariableList.Add(new TypeSelection2Class("IrradiationDate", dIrradiationDate.ToString()));



                                        if (int.TryParse(GlobalClass.GetCellValue(ThisSheet, "B36"), out iIndex))
                                            VariableList.Add(new TypeSelection2Class("IrradiationDepth", iIndex.ToString()));

                                        if (int.TryParse(GlobalClass.GetCellValue(ThisSheet, "B37"), out iIndex))
                                            VariableList.Add(new TypeSelection2Class("IrradiationFieldSize1", iIndex.ToString()));

                                        if (int.TryParse(GlobalClass.GetCellValue(ThisSheet, "B38"), out iIndex))
                                            VariableList.Add(new TypeSelection2Class("IrradiationFieldSize2", iIndex.ToString()));


                                        //int.TryParse(GlobalClass.GetCellValue(ThisSheet, "B39"), out iIndex);
                                        //this.BeamQualityTPR20Distance = iIndex;

                                        if (int.TryParse(GlobalClass.GetCellValue(ThisSheet, "B39"), out iIndex))
                                            VariableList.Add(new TypeSelection2Class("IrradiationDistance", iIndex.ToString()));


                                        if (int.TryParse(GlobalClass.GetCellValue(ThisSheet, "B40"), out iIndex))
                                        {
                                            if (iIndex == 1)
                                                VariableList.Add(new TypeSelection2Class("IrradiationDistanceType", "SSD"));
                                            else if (iIndex == 2)
                                                VariableList.Add(new TypeSelection2Class("IrradiationDistanceType", "SAD"));
                                            else
                                                VariableList.Add(new TypeSelection2Class("IrradiationDistanceType", "Off"));
                                        }

                                        VariableList.Add(new TypeSelection2Class("BeamGeometry", "Vertical"));

                                        if (double.TryParse(GlobalClass.GetCellValue(ThisSheet, "B41"), out dValue))
                                            VariableList.Add(new TypeSelection2Class("IrradiationSetting1", dValue.ToString()));

                                        if (double.TryParse(GlobalClass.GetCellValue(ThisSheet, "B42"), out dValue))
                                            VariableList.Add(new TypeSelection2Class("IrradiationSetting2", dValue.ToString()));

                                        if (double.TryParse(GlobalClass.GetCellValue(ThisSheet, "B43"), out dValue))
                                            VariableList.Add(new TypeSelection2Class("UserDose1", dValue.ToString()));

                                        if (double.TryParse(GlobalClass.GetCellValue(ThisSheet, "B44"), out dValue))
                                            VariableList.Add(new TypeSelection2Class("UserDose2", dValue.ToString()));

                                        string sFactors = GlobalClass.GetCellValue(ThisSheet, "B45");
                                        if ((GlobalClass.GetCellValue(ThisSheet, "B46").Trim() != "0") && (GlobalClass.GetCellValue(ThisSheet, "B46").Trim() != ""))
                                            sFactors = sFactors + System.Environment.NewLine + GlobalClass.GetCellValue(ThisSheet, "B46");
                                        if ((GlobalClass.GetCellValue(ThisSheet, "B47").Trim() != "0") && (GlobalClass.GetCellValue(ThisSheet, "B47").Trim() != ""))
                                            sFactors = sFactors + System.Environment.NewLine + GlobalClass.GetCellValue(ThisSheet, "B47");
                                        if ((GlobalClass.GetCellValue(ThisSheet, "B48").Trim() != "0") && (GlobalClass.GetCellValue(ThisSheet, "B48").Trim() != ""))
                                            sFactors = sFactors + System.Environment.NewLine + GlobalClass.GetCellValue(ThisSheet, "B48");
                                        if ((GlobalClass.GetCellValue(ThisSheet, "B49").Trim() != "0") && (GlobalClass.GetCellValue(ThisSheet, "B49").Trim() != ""))
                                            sFactors = sFactors + System.Environment.NewLine + GlobalClass.GetCellValue(ThisSheet, "B49");

                                        VariableList.Add(new TypeSelection2Class("Factors", sFactors));

                                        if (double.TryParse(GlobalClass.GetCellValue(ThisSheet, "B50"), out dValue))
                                            VariableList.Add(new TypeSelection2Class("BeamOutput", dValue.ToString()));

                                        if (int.TryParse(GlobalClass.GetCellValue(ThisSheet, "B51"), out iIndex))
                                        {
                                            if (isCobalt == true)
                                            {
                                                DictionaryTypeClass DictionaryType = GlobalClass.Dictionary.GetDictionaryType(GlobalClass.Dictionary.GetDictionary("BeamUnits"), iIndex + 1);
                                                if (DictionaryType != null)
                                                    VariableList.Add(new TypeSelection2Class("BeamUnits", DictionaryType.ItemCode));
                                            }
                                            else
                                            {
                                                DictionaryTypeClass DictionaryType = GlobalClass.Dictionary.GetDictionaryType(GlobalClass.Dictionary.GetDictionary("BeamUnits"), iIndex);
                                                if (DictionaryType != null)
                                                    VariableList.Add(new TypeSelection2Class("BeamUnits", DictionaryType.ItemCode));
                                            }
                                        }

                                        DateTime dBeamOutputDate = GlobalClass.FormatIAEAStringDateTimeValue(
                                                                   GlobalClass.GetIDEAYear(GlobalClass.GetCellValue(ThisSheet, "B54")) + "-" +
                                                                                           GlobalClass.GetCellValue(ThisSheet, "B53").PadLeft(2, '0') + "-" +
                                                                                           GlobalClass.GetCellValue(ThisSheet, "B52").PadLeft(2, '0'));
                                        VariableList.Add(new TypeSelection2Class("BeamOutputDate", dBeamOutputDate.ToString()));

                                        VariableList.Add(GlobalClass.EvaluationVariable(ThisSheet, "MeasuredByPersonFamilyName", "B56"));

                                        if (int.TryParse(GlobalClass.GetCellValue(ThisSheet, "B57"), out iIndex))
                                        {
                                            DictionaryTypeClass DictionaryType = GlobalClass.Dictionary.GetDictionaryType(GlobalClass.Dictionary.GetDictionary("MeasuredByPosition"), iIndex - 1);
                                            if (DictionaryType != null)
                                                VariableList.Add(new TypeSelection2Class("MeasuredByPosition", DictionaryType.ItemCode));
                                        }

                                        DateTime dMeasuredDate = GlobalClass.FormatIAEAStringDateTimeValue(
                                                                 GlobalClass.GetIDEAYear(GlobalClass.GetCellValue(ThisSheet, "B60")) + "-" +
                                                                                         GlobalClass.GetCellValue(ThisSheet, "B59").PadLeft(2, '0') + "-" +
                                                                                         GlobalClass.GetCellValue(ThisSheet, "B58").PadLeft(2, '0'));
                                        VariableList.Add(new TypeSelection2Class("MeasuredDate", dMeasuredDate.ToString()));

                                        if (int.TryParse(GlobalClass.GetCellValue(ThisSheet, "B63"), out iIndex))
                                        {
                                            DictionaryTypeClass DictionaryType = GlobalClass.Dictionary.GetDictionaryType(GlobalClass.Dictionary.GetDictionary("CalibrationType"), iIndex - 1);
                                            if (DictionaryType != null)
                                                VariableList.Add(new TypeSelection2Class("CalibrationType", DictionaryType.ItemCode));
                                        }

                                        if (double.TryParse(GlobalClass.GetCellValue(ThisSheet, "B64"), out dValue))
                                            VariableList.Add(new TypeSelection2Class("CalibrationValue", dValue.ToString()));

                                        if (int.TryParse(GlobalClass.GetCellValue(ThisSheet, "B65"), out iIndex))
                                        {
                                            DictionaryTypeClass DictionaryType = GlobalClass.Dictionary.GetDictionaryType(GlobalClass.Dictionary.GetDictionary("CalibrationUnit"), iIndex);
                                            if (DictionaryType != null)
                                                VariableList.Add(new TypeSelection2Class("CalibrationUnit", DictionaryType.ItemCode));
                                        }

                                        VariableList.Add(GlobalClass.EvaluationVariable(ThisSheet, "CalibrationLaboratory", "B66"));


                                        DateTime dCalibrationDate = GlobalClass.FormatIAEAStringDateTimeValue(
                                                                    GlobalClass.GetIDEAYear(GlobalClass.GetCellValue(ThisSheet, "B69")) + "-" +
                                                                                            GlobalClass.GetCellValue(ThisSheet, "B67").PadLeft(2, '0') + "-" +
                                                                                            GlobalClass.GetCellValue(ThisSheet, "B68").PadLeft(2, '0'));
                                        VariableList.Add(new TypeSelection2Class("CalibrationDate", dCalibrationDate.ToString()));

                                        if (double.TryParse(GlobalClass.GetCellValue(ThisSheet, "B70"), out dValue))
                                            VariableList.Add(new TypeSelection2Class("Temperature", dValue.ToString()));

                                        if (double.TryParse(GlobalClass.GetCellValue(ThisSheet, "B71"), out dValue))
                                            VariableList.Add(new TypeSelection2Class("Pressure", dValue.ToString()));

                                        if (int.TryParse(GlobalClass.GetCellValue(ThisSheet, "B72"), out iIndex))
                                        {
                                            DictionaryTypeClass DictionaryType = GlobalClass.Dictionary.GetDictionaryType(GlobalClass.Dictionary.GetDictionary("PressureUnit"), iIndex - 1);
                                            if (DictionaryType != null)
                                                VariableList.Add(new TypeSelection2Class("PressureUnit", DictionaryType.ItemCode));
                                        }

                                        if (int.TryParse(GlobalClass.GetCellValue(ThisSheet, "B73"), out iIndex))
                                        {
                                            if (iIndex == 1)
                                                VariableList.Add(new TypeSelection2Class("PhantomType", "Water"));
                                            else if (iIndex == 2)
                                                VariableList.Add(new TypeSelection2Class("PhantomType", "Plastic"));
                                            else
                                                VariableList.Add(new TypeSelection2Class("PhantomType", "Off"));
                                        }

                                        if (int.TryParse(GlobalClass.GetCellValue(ThisSheet, "B74"), out iIndex))
                                        {
                                            DictionaryTypeClass DictionaryType = GlobalClass.Dictionary.GetDictionaryType(GlobalClass.Dictionary.GetDictionary("PhantomMaterial"), iIndex - 1);
                                            if (DictionaryType != null)
                                                VariableList.Add(new TypeSelection2Class("PhantomMaterial", DictionaryType.ItemCode));
                                        }

                                        if (int.TryParse(GlobalClass.GetCellValue(ThisSheet, "B75"), out iIndex))
                                            VariableList.Add(new TypeSelection2Class("ChamberIrradiationFieldSize1", iIndex.ToString()));

                                        if (int.TryParse(GlobalClass.GetCellValue(ThisSheet, "B76"), out iIndex))
                                            VariableList.Add(new TypeSelection2Class("ChamberIrradiationFieldSize2", iIndex.ToString()));

                                        if (int.TryParse(GlobalClass.GetCellValue(ThisSheet, "B77"), out iIndex))
                                            VariableList.Add(new TypeSelection2Class("ChamberIrradiationDistance", iIndex.ToString()));


                                        if (int.TryParse(GlobalClass.GetCellValue(ThisSheet, "B78"), out iIndex))
                                        {
                                            if (iIndex == 1)
                                                VariableList.Add(new TypeSelection2Class("ChamberIrradiationDistanceType", "SSD"));
                                            else if (iIndex == 2)
                                                VariableList.Add(new TypeSelection2Class("ChamberIrradiationDistanceType", "SAD"));
                                            else
                                                VariableList.Add(new TypeSelection2Class("ChamberIrradiationDistanceType", "Off"));
                                        }

                                        if (int.TryParse(GlobalClass.GetCellValue(ThisSheet, "B79"), out iIndex))
                                        {
                                            DictionaryTypeClass DictionaryType = GlobalClass.Dictionary.GetDictionaryType(GlobalClass.Dictionary.GetDictionary("ChamberIrradiationMeasuringPoint"), iIndex - 1);
                                            if (DictionaryType != null)
                                                VariableList.Add(new TypeSelection2Class("ChamberIrradiationMeasuringPoint", DictionaryType.ItemCode));
                                        }

                                        if (int.TryParse(GlobalClass.GetCellValue(ThisSheet, "B80"), out iIndex))
                                            VariableList.Add(new TypeSelection2Class("ChamberIrradiationDepth", iIndex.ToString()));

                                        if (isCobalt == true)
                                        {
                                            if (int.TryParse(GlobalClass.GetCellValue(ThisSheet, "B81"), out iIndex))
                                            {
                                                DictionaryTypeClass DictionaryType = GlobalClass.Dictionary.GetDictionaryType(GlobalClass.Dictionary.GetDictionary("CapMaterial"), iIndex - 1);
                                                if (DictionaryType != null)
                                                    VariableList.Add(new TypeSelection2Class("CapMaterial", DictionaryType.ItemCode));
                                            }

                                            if (double.TryParse(GlobalClass.GetCellValue(ThisSheet, "B82"), out dValue))
                                                VariableList.Add(new TypeSelection2Class("CapThickness", dValue.ToString()));
                                        }

                                        if (double.TryParse(GlobalClass.GetCellValue(ThisSheet, "B83").Replace("nC", "").Trim(), out dValue))
                                            VariableList.Add(new TypeSelection2Class("ReadingUncorrected", dValue.ToString()));

                                        if (double.TryParse(GlobalClass.GetCellValue(ThisSheet, "B84"), out dValue))
                                            VariableList.Add(new TypeSelection2Class("ReadingMeasurementSetting", dValue.ToString()));

                                        if (double.TryParse(GlobalClass.GetCellValue(ThisSheet, "B85"), out dValue))
                                            VariableList.Add(new TypeSelection2Class("ReadingTemperature", dValue.ToString()));

                                        if (double.TryParse(GlobalClass.GetCellValue(ThisSheet, "B86"), out dValue))
                                            VariableList.Add(new TypeSelection2Class("ReadingPressure", dValue.ToString()));

                                        if (int.TryParse(GlobalClass.GetCellValue(ThisSheet, "B87"), out iIndex))
                                        {
                                            DictionaryTypeClass DictionaryType = GlobalClass.Dictionary.GetDictionaryType(GlobalClass.Dictionary.GetDictionary("ReadingPressureUnit"), iIndex - 1);
                                            if (DictionaryType != null)
                                                VariableList.Add(new TypeSelection2Class("ReadingPressureUnit", DictionaryType.ItemCode));
                                        }


                                        if (isCobalt == true)
                                        {
                                            VariableList.Add(new TypeSelection2Class("IrradiationUnits1", "min"));
                                            VariableList.Add(new TypeSelection2Class("IrradiationUnits2", "min"));
                                            VariableList.Add(new TypeSelection2Class("IrradiationUnits3", "min"));
                                            VariableList.Add(new TypeSelection2Class("CorrectionUnits", "s"));
                                            VariableList.Add(new TypeSelection2Class("ReadingMeasurementSettingUnits", "min"));
                                        }
                                        else
                                        {
                                            VariableList.Add(new TypeSelection2Class("IrradiationUnits1", "MU"));
                                            VariableList.Add(new TypeSelection2Class("IrradiationUnits2", "MU"));
                                            VariableList.Add(new TypeSelection2Class("IrradiationUnits3", "MU"));
                                            VariableList.Add(new TypeSelection2Class("CorrectionUnits", "MU"));
                                            VariableList.Add(new TypeSelection2Class("ReadingMeasurementSettingUnits", "MU"));
                                        }

                                        if (double.TryParse(GlobalClass.GetCellValue(ThisSheet, "B92"), out dValue))
                                            VariableList.Add(new TypeSelection2Class("CorrectionSetting", dValue.ToString()));

                                        //this.DetailedExplanations = GlobalClass.GetCellValue(ThisSheet, "B93");

                                        if (isCobalt == false)
                                        {
                                            if (int.TryParse(GlobalClass.GetCellValue(ThisSheet, "B96"), out iIndex))
                                                VariableList.Add(new TypeSelection2Class("EquipmentEnergy", iIndex.ToString()));

                                            if (double.TryParse(GlobalClass.GetCellValue(ThisSheet, "B97"), out dValue))
                                            {
                                                VariableList.Add(new TypeSelection2Class("BeamQualityD20D10", dValue.ToString()));
                                                VariableList.Add(new TypeSelection2Class("BeamQuality", "D20/D10"));
                                            }

                                            if (double.TryParse(GlobalClass.GetCellValue(ThisSheet, "B98"), out dValue))
                                            {
                                                VariableList.Add(new TypeSelection2Class("BeamQualityTPR20", dValue.ToString()));
                                                VariableList.Add(new TypeSelection2Class("BeamQuality", "TRP20/10"));
                                            }

                                            if (int.TryParse(GlobalClass.GetCellValue(ThisSheet, "B99"), out iIndex))
                                                VariableList.Add(new TypeSelection2Class("BeamQualityTPR20Distance", iIndex.ToString()));

                                            if (double.TryParse(GlobalClass.GetCellValue(ThisSheet, "B100"), out dValue))
                                            {
                                                VariableList.Add(new TypeSelection2Class("BeamQualityOther", dValue.ToString()));
                                                VariableList.Add(new TypeSelection2Class("BeamQuality", "Other"));
                                            }

                                            VariableList.Add(GlobalClass.EvaluationVariable(ThisSheet, "BeamQualityOtherConditions", "B101"));
                                        }
                                    }
                                }
                            }
                            catch (Exception oEx)
                            {
                                MessageBox.Show("Import TLD Data Sheet Error. " + oEx.Message, "Import TLD Data Sheet Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                            finally
                            {
                                ThisWorkBook.Close(false, oMissing, oMissing);

                                ThisApplication.Quit();

                                //GlobalClass.releaseObject(ThisSheet);
                                GlobalClass.releaseObject(ThisWorkBook);
                                GlobalClass.releaseObject(ThisApplication);
                            }
                        }
                        catch (Exception oEx)
                        {
                            MessageBox.Show("Import TLD Data Sheet Error. " + oEx.Message, "Import TLD Data Sheet Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }
            }
            return VariableList;
        }

        public static List<TypeSelection2Class> LoadVariableListTLDDataSheetSSDL(string sFileName)
        {
            List<TypeSelection2Class> VariableList = new List<TypeSelection2Class>();
            VariableList.Clear();

            if (File.Exists(sFileName))
            {
                bool isCobalt = false;
                
                string sBatchNo = string.Empty;
                string sSetNo = string.Empty;

                int iIndex = -1;
                double dValue = -1;
                string sValue = string.Empty;
                object oMissing = System.Reflection.Missing.Value;

                string sTranslationFilePath = GlobalClass.sApplicationStartupPath + "\\Translation.csv";

                if (File.Exists(sTranslationFilePath))
                {
                    List<string[]> TranslationEquipmentCo60 = GlobalClass.ParseCSV(sTranslationFilePath, "EquipmentCo60", ',');
                    List<string[]> TranslationEquipmentLinac = GlobalClass.ParseCSV(sTranslationFilePath, "EquipmentLinac", ',');

                    List<string[]> TranslationIonisationChamber = GlobalClass.ParseCSV(sTranslationFilePath, "IonisationChamber", ',');
                    List<string[]> TranslationElectrometer = GlobalClass.ParseCSV(sTranslationFilePath, "Electrometer", ',');


                    if (File.Exists(sFileName))
                    {
                        try
                        {
                            Excel.Application ThisApplication = new Excel.ApplicationClass();
                            ThisApplication.Visible = false;
                            Excel.Workbook ThisWorkBook = (Excel.Workbook)ThisApplication.Workbooks.Open(sFileName, oMissing, oMissing, 5, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
                            //Excel.Worksheet ThisSheet = (Excel.Worksheet)ThisWorkBook.Sheets[1]; //"Sheet1"

                            ThisWorkBook.Unprotect("dmrp");


                            try
                            {
                                foreach (Excel.Worksheet ThisSheet in ThisWorkBook.Application.Worksheets)
                                {
                                    ThisSheet.Visible = Microsoft.Office.Interop.Excel.XlSheetVisibility.xlSheetVisible;

                                    if (ThisSheet.Name == "Datasheet Part1")
                                    {
                                        // Linac - "The treatment unit used for this audit is of the type"
                                        // Cobalt - "The Co-60 treatment unit used for this audit is of the type"

                                        isCobalt = GlobalClass.GetCellValue(ThisSheet, "B5").Contains("Co-60");

                                        if (isCobalt == true)
                                        {
                                            string sConditions = GlobalClass.GetCellValue(ThisSheet, "B59");

                                            if ((GlobalClass.GetCellValue(ThisSheet, "B60").Trim() != "0") && (GlobalClass.GetCellValue(ThisSheet, "B60").Trim() != ""))
                                                sConditions = sConditions + System.Environment.NewLine + GlobalClass.GetCellValue(ThisSheet, "B60");
                                            if ((GlobalClass.GetCellValue(ThisSheet, "B61").Trim() != "0") && (GlobalClass.GetCellValue(ThisSheet, "B61").Trim() != ""))
                                                sConditions = sConditions + System.Environment.NewLine + GlobalClass.GetCellValue(ThisSheet, "B61");
                                            if ((GlobalClass.GetCellValue(ThisSheet, "B62").Trim() != "0") && (GlobalClass.GetCellValue(ThisSheet, "B62").Trim() != ""))
                                                sConditions = sConditions + System.Environment.NewLine + GlobalClass.GetCellValue(ThisSheet, "B62");
                                            if ((GlobalClass.GetCellValue(ThisSheet, "B63").Trim() != "0") && (GlobalClass.GetCellValue(ThisSheet, "B63").Trim() != ""))
                                                sConditions = sConditions + System.Environment.NewLine + GlobalClass.GetCellValue(ThisSheet, "B63");

                                            VariableList.Add(new TypeSelection2Class("Conditions", sConditions));
                                        }
                                        else
                                        {
                                            string sConditions = GlobalClass.GetCellValue(ThisSheet, "B60");

                                            if ((GlobalClass.GetCellValue(ThisSheet, "B61").Trim() != "0") && (GlobalClass.GetCellValue(ThisSheet, "B61").Trim() != ""))
                                                sConditions = sConditions + System.Environment.NewLine + GlobalClass.GetCellValue(ThisSheet, "B61");
                                            if ((GlobalClass.GetCellValue(ThisSheet, "B62").Trim() != "0") && (GlobalClass.GetCellValue(ThisSheet, "B62").Trim() != ""))
                                                sConditions = sConditions + System.Environment.NewLine + GlobalClass.GetCellValue(ThisSheet, "B62");
                                            if ((GlobalClass.GetCellValue(ThisSheet, "B63").Trim() != "0") && (GlobalClass.GetCellValue(ThisSheet, "B63").Trim() != ""))
                                                sConditions = sConditions + System.Environment.NewLine + GlobalClass.GetCellValue(ThisSheet, "B63");

                                            VariableList.Add(new TypeSelection2Class("Conditions", sConditions));
                                        }
                                    }
                                    else if (ThisSheet.Name == "Datasheet Part2")
                                    {
                                        if (isCobalt == true)
                                        {
                                            string sDetailedExplanations = GlobalClass.GetCellValue(ThisSheet, "B61");

                                            if ((GlobalClass.GetCellValue(ThisSheet, "B62").Trim() != "0") && (GlobalClass.GetCellValue(ThisSheet, "B62").Trim() != ""))
                                                sDetailedExplanations = sDetailedExplanations + System.Environment.NewLine + GlobalClass.GetCellValue(ThisSheet, "B62");
                                            if ((GlobalClass.GetCellValue(ThisSheet, "B63").Trim() != "0") && (GlobalClass.GetCellValue(ThisSheet, "B63").Trim() != ""))
                                                sDetailedExplanations = sDetailedExplanations + System.Environment.NewLine + GlobalClass.GetCellValue(ThisSheet, "B63");
                                            if ((GlobalClass.GetCellValue(ThisSheet, "B64").Trim() != "0") && (GlobalClass.GetCellValue(ThisSheet, "B64").Trim() != ""))
                                                sDetailedExplanations = sDetailedExplanations + System.Environment.NewLine + GlobalClass.GetCellValue(ThisSheet, "B64");
                                            if ((GlobalClass.GetCellValue(ThisSheet, "B65").Trim() != "0") && (GlobalClass.GetCellValue(ThisSheet, "B65").Trim() != ""))
                                                sDetailedExplanations = sDetailedExplanations + System.Environment.NewLine + GlobalClass.GetCellValue(ThisSheet, "B65");
                                            if ((GlobalClass.GetCellValue(ThisSheet, "B66").Trim() != "0") && (GlobalClass.GetCellValue(ThisSheet, "B66").Trim() != ""))
                                                sDetailedExplanations = sDetailedExplanations + System.Environment.NewLine + GlobalClass.GetCellValue(ThisSheet, "B66");
                                            if ((GlobalClass.GetCellValue(ThisSheet, "B67").Trim() != "0") && (GlobalClass.GetCellValue(ThisSheet, "B67").Trim() != ""))
                                                sDetailedExplanations = sDetailedExplanations + System.Environment.NewLine + GlobalClass.GetCellValue(ThisSheet, "B67");

                                            VariableList.Add(new TypeSelection2Class("DetailedExplanations", sDetailedExplanations));
                                        }
                                        else
                                        {
                                            string sDetailedExplanations = GlobalClass.GetCellValue(ThisSheet, "B57");

                                            if ((GlobalClass.GetCellValue(ThisSheet, "B58").Trim() != "0") && (GlobalClass.GetCellValue(ThisSheet, "B58").Trim() != ""))
                                                sDetailedExplanations = sDetailedExplanations + System.Environment.NewLine + GlobalClass.GetCellValue(ThisSheet, "B58");
                                            if ((GlobalClass.GetCellValue(ThisSheet, "B59").Trim() != "0") && (GlobalClass.GetCellValue(ThisSheet, "B59").Trim() != ""))
                                                sDetailedExplanations = sDetailedExplanations + System.Environment.NewLine + GlobalClass.GetCellValue(ThisSheet, "B59");
                                            if ((GlobalClass.GetCellValue(ThisSheet, "B60").Trim() != "0") && (GlobalClass.GetCellValue(ThisSheet, "B60").Trim() != ""))
                                                sDetailedExplanations = sDetailedExplanations + System.Environment.NewLine + GlobalClass.GetCellValue(ThisSheet, "B60");
                                            if ((GlobalClass.GetCellValue(ThisSheet, "B61").Trim() != "0") && (GlobalClass.GetCellValue(ThisSheet, "B61").Trim() != ""))
                                                sDetailedExplanations = sDetailedExplanations + System.Environment.NewLine + GlobalClass.GetCellValue(ThisSheet, "B61");
                                            if ((GlobalClass.GetCellValue(ThisSheet, "B62").Trim() != "0") && (GlobalClass.GetCellValue(ThisSheet, "B62").Trim() != ""))
                                                sDetailedExplanations = sDetailedExplanations + System.Environment.NewLine + GlobalClass.GetCellValue(ThisSheet, "B62");
                                            if ((GlobalClass.GetCellValue(ThisSheet, "B63").Trim() != "0") && (GlobalClass.GetCellValue(ThisSheet, "B63").Trim() != ""))
                                                sDetailedExplanations = sDetailedExplanations + System.Environment.NewLine + GlobalClass.GetCellValue(ThisSheet, "B63");

                                            VariableList.Add(new TypeSelection2Class("DetailedExplanations", sDetailedExplanations));
                                        }
                                    }
                                    else if (ThisSheet.Name == "Tables")
                                    {
                                        if (isCobalt == true)
                                            VariableList.Add(new TypeSelection2Class("SetBeamType", sBeamTypeCo60));
                                        else
                                            VariableList.Add(new TypeSelection2Class("SetBeamType", sBeamTypePhoton));

                                        sBatchNo = GlobalClass.GetCellValue(ThisSheet, "B97").Trim();
                                        sSetNo = GlobalClass.GetCellValue(ThisSheet, "B98").Trim();

                                        VariableList.Add(new TypeSelection2Class("BatchNo", sBatchNo));
                                        VariableList.Add(new TypeSelection2Class("SetNo", sSetNo));

                                        VariableList.Add(new TypeSelection2Class("AuditType", sAuditTypeRT));

                                        if (sSetNo.Substring(sSetNo.Length - 1, 1) != "R")
                                        {
                                            VariableList.Add(new TypeSelection2Class("SetType", "1")); // 1-FirstIrradiation | 2-FollowUp 

                                            if (sSetNo.Contains("P"))
                                                VariableList.Add(new TypeSelection2Class("ParticipationType", sParticipationTypePrimary));
                                            else if (sSetNo.Contains("R"))
                                                VariableList.Add(new TypeSelection2Class("ParticipationType", sParticipationTypeReference));
                                            else if (sSetNo.Contains("DL"))
                                                VariableList.Add(new TypeSelection2Class("ParticipationType", sParticipationTypeSSDL));
                                            else
                                                VariableList.Add(new TypeSelection2Class("ParticipationType", sParticipationTypeHospitals));
                                        }
                                        else 
                                        {
                                            VariableList.Add(new TypeSelection2Class("SetType", "2")); // 1-FirstIrradiation | 2-FollowUp 
                                            if (sSetNo.Substring(0, sSetNo.Length - 1).Contains("P"))
                                                VariableList.Add(new TypeSelection2Class("ParticipationType", sParticipationTypePrimary));
                                            else if (sSetNo.Substring(0, sSetNo.Length - 1).Contains("R"))
                                                VariableList.Add(new TypeSelection2Class("ParticipationType", sParticipationTypeReference));
                                            else if (sSetNo.Substring(0, sSetNo.Length - 1).Contains("DL"))
                                                VariableList.Add(new TypeSelection2Class("ParticipationType", sParticipationTypeSSDL));
                                            else
                                                VariableList.Add(new TypeSelection2Class("ParticipationType", sParticipationTypeHospitals));
                                        }

                                        if (int.TryParse(GlobalClass.GetCellValue(ThisSheet, "B21"), out iIndex))
                                        {
                                            if (iIndex == 1)
                                                sValue = "Yes";
                                            else if (iIndex == 2)
                                                sValue = "No";
                                            else
                                                sValue = "Off";

                                            VariableList.Add(new TypeSelection2Class("PreviousParticipation", sValue));
                                        }

                                        if (isCobalt == true)
                                        {
                                            if (int.TryParse(GlobalClass.GetCellValue(ThisSheet, "B26"), out iIndex))
                                            {
                                                string sItemCode = GlobalClass.Translation(TranslationEquipmentCo60, iIndex);

                                                EquipmentTypeClass Equipment = GlobalClass.Dictionary.GetEquipmentItem(GlobalClass.Dictionary.GetEquipmentByGroup("RadionuclideTherapy"), sItemCode);

                                                if (Equipment != null)
                                                {
                                                    VariableList.Add(new TypeSelection2Class("Equipment", Equipment.ItemCode));
                                                    VariableList.Add(new TypeSelection2Class("EquipmentCo60", Equipment.ItemCode));
                                                }
                                            }
                                        }
                                        else
                                        {
                                            if (int.TryParse(GlobalClass.GetCellValue(ThisSheet, "B26"), out iIndex))
                                            {
                                                string sItemCode = GlobalClass.Translation(TranslationEquipmentLinac, iIndex);

                                                EquipmentTypeClass Equipment = GlobalClass.Dictionary.GetEquipmentItem(GlobalClass.Dictionary.GetEquipmentByGroup("LinearAccelerator"), sItemCode);

                                                if (Equipment != null)
                                                {
                                                    VariableList.Add(new TypeSelection2Class("Equipment", Equipment.ItemCode));
                                                    VariableList.Add(new TypeSelection2Class("EquipmentLinac", Equipment.ItemCode));
                                                }
                                            }
                                        }

                                        VariableList.Add(GlobalClass.EvaluationVariable(ThisSheet, "EquipmentOther", "B27"));

                                        if (int.TryParse(GlobalClass.GetCellValue(ThisSheet, "B64"), out iIndex))
                                        {
                                            string sItemCode = GlobalClass.Translation(TranslationIonisationChamber, iIndex);

                                            EquipmentTypeClass Equipment = GlobalClass.Dictionary.GetEquipmentItem(GlobalClass.Dictionary.GetEquipment("IonisationChamber"), sItemCode);

                                            if (Equipment != null)
                                                VariableList.Add(new TypeSelection2Class("IonisationChamber", Equipment.ItemCode));
                                        }

                                        if (int.TryParse(GlobalClass.GetCellValue(ThisSheet, "B65"), out iIndex))
                                        {
                                            string sItemCode = GlobalClass.Translation(TranslationElectrometer, iIndex);

                                            EquipmentTypeClass Equipment = GlobalClass.Dictionary.GetEquipmentItem(GlobalClass.Dictionary.GetEquipment("Electrometer"), sItemCode);

                                            if (Equipment != null)
                                                VariableList.Add(new TypeSelection2Class("Electrometer", Equipment.ItemCode));
                                        }

                                        if (int.TryParse(GlobalClass.GetCellValue(ThisSheet, "B93"), out iIndex))
                                        {
                                            EquipmentTypeClass Equipment = GlobalClass.Dictionary.GetEquipmentItem(GlobalClass.Dictionary.GetEquipment("DosimetryProtocol"), iIndex - 1);

                                            if (Equipment != null)
                                                VariableList.Add(new TypeSelection2Class("DosimetryProtocol", Equipment.ItemCode));
                                        }

                                        string sContactFamilyName = GlobalClass.GetCellValue(ThisSheet, "B1").Trim();
                                        string sContactPosition = GlobalClass.GetCellValue(ThisSheet, "B2").Trim();

                                        if (sContactFamilyName == "0") sContactFamilyName = string.Empty;
                                        if (sContactPosition == "0") sContactPosition = string.Empty;

                                        VariableList.Add(new TypeSelection2Class("ContactFamilyName", sContactFamilyName));
                                        VariableList.Add(new TypeSelection2Class("ContactPosition", sContactPosition));

                                        VariableList.Add(GlobalClass.EvaluationVariable(ThisSheet, "ContactDepartment", "B4"));
                                        VariableList.Add(GlobalClass.EvaluationVariable(ThisSheet, "ContactTelephone1", "B10"));
                                        VariableList.Add(GlobalClass.EvaluationVariable(ThisSheet, "ContactTelephone2", "B11"));
                                        VariableList.Add(GlobalClass.EvaluationVariable(ThisSheet, "ContactEmail", "B13"));


                                        VariableList.Add(GlobalClass.EvaluationVariable(ThisSheet, "OperatorName", "B3"));
                                        VariableList.Add(GlobalClass.EvaluationVariable(ThisSheet, "Country", "B5"));
                                        VariableList.Add(GlobalClass.EvaluationVariable(ThisSheet, "Street", "B6"));
                                        VariableList.Add(GlobalClass.EvaluationVariable(ThisSheet, "City", "B7"));
                                        VariableList.Add(GlobalClass.EvaluationVariable(ThisSheet, "State", "B9"));

                                        VariableList.Add(GlobalClass.EvaluationVariable(ThisSheet, "CompletedByPersonFamilyName", "B14"));
                                        VariableList.Add(new TypeSelection2Class("CompletedByPersonFirstName", string.Empty));

                                        DateTime dCompletedDate = GlobalClass.FormatIAEAStringDateTimeValue(
                                                             GlobalClass.GetIDEAYear(GlobalClass.GetCellValue(ThisSheet, "B18")) + "-" +
                                                                                     GlobalClass.GetCellValue(ThisSheet, "B17").PadLeft(2, '0') + "-" +
                                                                                     GlobalClass.GetCellValue(ThisSheet, "B16").PadLeft(2, '0'));

                                        VariableList.Add(new TypeSelection2Class("CompletedDate", dCompletedDate.ToString()));

                                        VariableList.Add(GlobalClass.EvaluationVariable(ThisSheet, "IrradiatedByPersonFamilyName", "B19"));
                                        VariableList.Add(new TypeSelection2Class("IrradiatedByPersonFirstName", string.Empty));
                                        VariableList.Add(new TypeSelection2Class("IrradiatedByPersonPosition", string.Empty));

                                        if (int.TryParse(GlobalClass.GetCellValue(ThisSheet, "B25"), out iIndex))
                                        {
                                            DictionaryTypeClass DictionaryType = GlobalClass.Dictionary.GetDictionaryType(GlobalClass.Dictionary.GetDictionary("ParticipationOrganiser"), iIndex - 1);
                                            if (DictionaryType != null)
                                                VariableList.Add(new TypeSelection2Class("ParticipationOrganiser", DictionaryType.ItemCode));
                                        }

                                        if (int.TryParse(GlobalClass.GetCellValue(ThisSheet, "B24"), out iIndex))
                                        {
                                            DictionaryTypeClass DictionaryType = GlobalClass.Dictionary.GetDictionaryType(GlobalClass.Dictionary.GetDictionary("ParticipationYear"), iIndex);
                                            if (DictionaryType != null)
                                                VariableList.Add(new TypeSelection2Class("ParticipationYear", DictionaryType.ItemCode));
                                        }

                                        

                                        if (int.TryParse(GlobalClass.GetCellValue(ThisSheet, "B28"), out iIndex))
                                        {
                                            DictionaryTypeClass DictionaryType = GlobalClass.Dictionary.GetDictionaryType(GlobalClass.Dictionary.GetDictionary("EquipmentProductionYear"), iIndex);
                                            if (DictionaryType != null)
                                                VariableList.Add(new TypeSelection2Class("EquipmentProductionYear", DictionaryType.ItemCode));
                                        }

                                        if (int.TryParse(GlobalClass.GetCellValue(ThisSheet, "B29"), out iIndex))
                                        {
                                            DictionaryTypeClass DictionaryType = GlobalClass.Dictionary.GetDictionaryType(GlobalClass.Dictionary.GetDictionary("EquipmentInstallationYear"), iIndex);
                                            if (DictionaryType != null)
                                                VariableList.Add(new TypeSelection2Class("EquipmentInstallationYear", DictionaryType.ItemCode));
                                        }


                                        if (isCobalt == true)
                                        {
                                            //DateTime dEquipmentLastSourceReplacementDate = GlobalClass.FormatIAEAStringDateTimeValue(
                                            //                                               GlobalClass.GetIDEAYear(GlobalClass.GetCellValue(ThisSheet, "B32")) + "-" +
                                            //                                                                       GlobalClass.GetCellValue(ThisSheet, "B31").PadLeft(2, '0') + "-" +
                                            //                                                                       GlobalClass.GetCellValue(ThisSheet, "B30").PadLeft(2, '0'));

                                            VariableList.Add(new TypeSelection2Class("EquipmentLastSourceReplacementYear", GlobalClass.GetCellValue(ThisSheet, "B32").ToString()));
                                        }
                                        else
                                        {
                                            if (int.TryParse(GlobalClass.GetCellValue(ThisSheet, "B99"), out iIndex))
                                                VariableList.Add(new TypeSelection2Class("EquipmentEnergy", iIndex.ToString()));

                                            if (double.TryParse(GlobalClass.GetCellValue(ThisSheet, "B100"), out dValue))
                                            {
                                                VariableList.Add(new TypeSelection2Class("BeamQuality", "D20/D10"));
                                                VariableList.Add(new TypeSelection2Class("BeamQualityD20D10", dValue.ToString()));
                                            }

                                            if (double.TryParse(GlobalClass.GetCellValue(ThisSheet, "B101"), out dValue))
                                            {
                                                VariableList.Add(new TypeSelection2Class("BeamQuality", "TRP20/10"));
                                                VariableList.Add(new TypeSelection2Class("BeamQualityTPR20", dValue.ToString()));
                                                if (int.TryParse(GlobalClass.GetCellValue(ThisSheet, "B102"), out iIndex))
                                                    VariableList.Add(new TypeSelection2Class("BeamQualityTPR20Distance", iIndex.ToString()));
                                            }


                                            if (double.TryParse(GlobalClass.GetCellValue(ThisSheet, "B103"), out dValue))
                                            {
                                                VariableList.Add(new TypeSelection2Class("BeamQuality", "Other"));
                                                VariableList.Add(new TypeSelection2Class("BeamQualityOther", dValue.ToString()));
                                                VariableList.Add(GlobalClass.EvaluationVariable(ThisSheet, "BeamQualityOtherConditions", "B104"));
                                            }
                                        }

                                        DateTime dIrradiationDate = GlobalClass.FormatIAEAStringDateTimeValue(
                                                                    GlobalClass.GetIDEAYear(GlobalClass.GetCellValue(ThisSheet, "B35")) + "-" +
                                                                                            GlobalClass.GetCellValue(ThisSheet, "B34").PadLeft(2, '0') + "-" +
                                                                                            GlobalClass.GetCellValue(ThisSheet, "B33").PadLeft(2, '0'));

                                        VariableList.Add(new TypeSelection2Class("IrradiationDate", dIrradiationDate.ToString()));

                                        if (int.TryParse(GlobalClass.GetCellValue(ThisSheet, "B36"), out iIndex))
                                            VariableList.Add(new TypeSelection2Class("IrradiationDepth", iIndex.ToString()));

                                        if (int.TryParse(GlobalClass.GetCellValue(ThisSheet, "B37"), out iIndex))
                                            VariableList.Add(new TypeSelection2Class("IrradiationFieldSize1", iIndex.ToString()));

                                        if (int.TryParse(GlobalClass.GetCellValue(ThisSheet, "B38"), out iIndex))
                                            VariableList.Add(new TypeSelection2Class("IrradiationFieldSize2", iIndex.ToString()));


                                        //int.TryParse(GlobalClass.GetCellValue(ThisSheet, "B39"), out iIndex);
                                        //this.BeamQualityTPR20Distance = iIndex;

                                        if (int.TryParse(GlobalClass.GetCellValue(ThisSheet, "B39"), out iIndex))
                                            VariableList.Add(new TypeSelection2Class("IrradiationDistance", iIndex.ToString()));


                                        if (int.TryParse(GlobalClass.GetCellValue(ThisSheet, "B40"), out iIndex))
                                        {
                                            if (iIndex == 1)
                                                VariableList.Add(new TypeSelection2Class("IrradiationDistanceType", "SSD"));
                                            else if (iIndex == 2)
                                                VariableList.Add(new TypeSelection2Class("IrradiationDistanceType", "SAD"));
                                            else
                                                VariableList.Add(new TypeSelection2Class("IrradiationDistanceType", "Off"));
                                        }

                                        if (int.TryParse(GlobalClass.GetCellValue(ThisSheet, "B41"), out iIndex))
                                        {
                                            if (iIndex == 1)
                                                VariableList.Add(new TypeSelection2Class("BeamGeometry", "Horizontal"));
                                            else if (iIndex == 2)
                                                VariableList.Add(new TypeSelection2Class("BeamGeometry", "Vertical"));
                                        }

                                        if (double.TryParse(GlobalClass.GetCellValue(ThisSheet, "B42"), out dValue))
                                            VariableList.Add(new TypeSelection2Class("IrradiationSetting1", dValue.ToString()));

                                        if (double.TryParse(GlobalClass.GetCellValue(ThisSheet, "B43"), out dValue))
                                            VariableList.Add(new TypeSelection2Class("IrradiationSetting2", dValue.ToString()));

                                        if (double.TryParse(GlobalClass.GetCellValue(ThisSheet, "B44"), out dValue))
                                            VariableList.Add(new TypeSelection2Class("IrradiationSetting3", dValue.ToString()));


                                        if (double.TryParse(GlobalClass.GetCellValue(ThisSheet, "B45"), out dValue))
                                            VariableList.Add(new TypeSelection2Class("UserDose1", dValue.ToString()));

                                        if (double.TryParse(GlobalClass.GetCellValue(ThisSheet, "B46"), out dValue))
                                            VariableList.Add(new TypeSelection2Class("UserDose2", dValue.ToString()));

                                        if (double.TryParse(GlobalClass.GetCellValue(ThisSheet, "B47"), out dValue))
                                            VariableList.Add(new TypeSelection2Class("UserDose3", dValue.ToString()));


                                        string sFactors = GlobalClass.GetCellValue(ThisSheet, "B48");
                                        if ((GlobalClass.GetCellValue(ThisSheet, "B49").Trim() != "0") && (GlobalClass.GetCellValue(ThisSheet, "B49").Trim() != ""))
                                            sFactors = sFactors + System.Environment.NewLine + GlobalClass.GetCellValue(ThisSheet, "B49");
                                        if ((GlobalClass.GetCellValue(ThisSheet, "B50").Trim() != "0") && (GlobalClass.GetCellValue(ThisSheet, "B50").Trim() != ""))
                                            sFactors = sFactors + System.Environment.NewLine + GlobalClass.GetCellValue(ThisSheet, "B50");
                                        if ((GlobalClass.GetCellValue(ThisSheet, "B51").Trim() != "0") && (GlobalClass.GetCellValue(ThisSheet, "B51").Trim() != ""))
                                            sFactors = sFactors + System.Environment.NewLine + GlobalClass.GetCellValue(ThisSheet, "B51");
                                        if ((GlobalClass.GetCellValue(ThisSheet, "B52").Trim() != "0") && (GlobalClass.GetCellValue(ThisSheet, "B52").Trim() != ""))
                                            sFactors = sFactors + System.Environment.NewLine + GlobalClass.GetCellValue(ThisSheet, "B52");
                                        VariableList.Add(new TypeSelection2Class("Factors", sFactors));

                                        if (double.TryParse(GlobalClass.GetCellValue(ThisSheet, "B53"), out dValue))
                                            VariableList.Add(new TypeSelection2Class("BeamOutput", dValue.ToString()));

                                        if (int.TryParse(GlobalClass.GetCellValue(ThisSheet, "B54"), out iIndex))
                                        {
                                            if (isCobalt == true)
                                            {
                                                DictionaryTypeClass DictionaryType = GlobalClass.Dictionary.GetDictionaryType(GlobalClass.Dictionary.GetDictionary("BeamUnits"), iIndex + 1);
                                                if (DictionaryType != null)
                                                    VariableList.Add(new TypeSelection2Class("BeamUnits", DictionaryType.ItemCode));
                                            }
                                            else
                                            {
                                                DictionaryTypeClass DictionaryType = GlobalClass.Dictionary.GetDictionaryType(GlobalClass.Dictionary.GetDictionary("BeamUnits"), iIndex);
                                                if (DictionaryType != null)
                                                    VariableList.Add(new TypeSelection2Class("BeamUnits", DictionaryType.ItemCode));
                                            }
                                        }

                                        DateTime dBeamOutputDate = GlobalClass.FormatIAEAStringDateTimeValue(
                                                                   GlobalClass.GetIDEAYear(GlobalClass.GetCellValue(ThisSheet, "B57")) + "-" +
                                                                                           GlobalClass.GetCellValue(ThisSheet, "B56").PadLeft(2, '0') + "-" +
                                                                                           GlobalClass.GetCellValue(ThisSheet, "B55").PadLeft(2, '0'));
                                        VariableList.Add(new TypeSelection2Class("BeamOutputDate", dBeamOutputDate.ToString()));

                                        VariableList.Add(GlobalClass.EvaluationVariable(ThisSheet, "MeasuredByPersonFamilyName", "B59"));

                                        if (int.TryParse(GlobalClass.GetCellValue(ThisSheet, "B60"), out iIndex))
                                        {
                                            DictionaryTypeClass DictionaryType = GlobalClass.Dictionary.GetDictionaryType(GlobalClass.Dictionary.GetDictionary("MeasuredByPosition"), iIndex - 1);
                                            if (DictionaryType != null)
                                                VariableList.Add(new TypeSelection2Class("MeasuredByPosition", DictionaryType.ItemCode));
                                        }

                                        DateTime dMeasuredDate = GlobalClass.FormatIAEAStringDateTimeValue(
                                                                 GlobalClass.GetIDEAYear(GlobalClass.GetCellValue(ThisSheet, "B63")) + "-" +
                                                                                         GlobalClass.GetCellValue(ThisSheet, "B62").PadLeft(2, '0') + "-" +
                                                                                         GlobalClass.GetCellValue(ThisSheet, "B61").PadLeft(2, '0'));
                                        VariableList.Add(new TypeSelection2Class("MeasuredDate", dMeasuredDate.ToString()));

                                        if (int.TryParse(GlobalClass.GetCellValue(ThisSheet, "B66"), out iIndex))
                                        {
                                            DictionaryTypeClass DictionaryType = GlobalClass.Dictionary.GetDictionaryType(GlobalClass.Dictionary.GetDictionary("CalibrationType"), iIndex - 1);
                                            if (DictionaryType != null)
                                                VariableList.Add(new TypeSelection2Class("CalibrationType", DictionaryType.ItemCode));
                                        }

                                        if (double.TryParse(GlobalClass.GetCellValue(ThisSheet, "B67"), out dValue))
                                            VariableList.Add(new TypeSelection2Class("CalibrationValue", dValue.ToString()));

                                        if (int.TryParse(GlobalClass.GetCellValue(ThisSheet, "B68"), out iIndex))
                                        {
                                            DictionaryTypeClass DictionaryType = GlobalClass.Dictionary.GetDictionaryType(GlobalClass.Dictionary.GetDictionary("CalibrationUnit"), iIndex);
                                            if (DictionaryType != null)
                                                VariableList.Add(new TypeSelection2Class("CalibrationUnit", DictionaryType.ItemCode));
                                        }

                                        VariableList.Add(GlobalClass.EvaluationVariable(ThisSheet, "CalibrationLaboratory", "B69"));


                                        DateTime dCalibrationDate = GlobalClass.FormatIAEAStringDateTimeValue(
                                                                    GlobalClass.GetIDEAYear(GlobalClass.GetCellValue(ThisSheet, "B72")) + "-" +
                                                                                            GlobalClass.GetCellValue(ThisSheet, "B71").PadLeft(2, '0') + "-" +
                                                                                            GlobalClass.GetCellValue(ThisSheet, "B70").PadLeft(2, '0'));
                                        VariableList.Add(new TypeSelection2Class("CalibrationDate", dCalibrationDate.ToString()));

                                        if (double.TryParse(GlobalClass.GetCellValue(ThisSheet, "B73"), out dValue))
                                            VariableList.Add(new TypeSelection2Class("Temperature", dValue.ToString()));

                                        if (double.TryParse(GlobalClass.GetCellValue(ThisSheet, "B74"), out dValue))
                                            VariableList.Add(new TypeSelection2Class("Pressure", dValue.ToString()));

                                        if (int.TryParse(GlobalClass.GetCellValue(ThisSheet, "B75"), out iIndex))
                                        {
                                            DictionaryTypeClass DictionaryType = GlobalClass.Dictionary.GetDictionaryType(GlobalClass.Dictionary.GetDictionary("PressureUnit"), iIndex - 1);
                                            if (DictionaryType != null)
                                                VariableList.Add(new TypeSelection2Class("PressureUnit", DictionaryType.ItemCode));
                                        }

                                        if (int.TryParse(GlobalClass.GetCellValue(ThisSheet, "B76"), out iIndex))
                                        {
                                            if (iIndex == 1)
                                                VariableList.Add(new TypeSelection2Class("PhantomType", "Water"));
                                            else if (iIndex == 2)
                                                VariableList.Add(new TypeSelection2Class("PhantomType", "Plastic"));
                                            else
                                                VariableList.Add(new TypeSelection2Class("PhantomType", "Off"));
                                        }

                                        if (int.TryParse(GlobalClass.GetCellValue(ThisSheet, "B77"), out iIndex))
                                        {
                                            DictionaryTypeClass DictionaryType = GlobalClass.Dictionary.GetDictionaryType(GlobalClass.Dictionary.GetDictionary("PhantomMaterial"), iIndex - 1);
                                            if (DictionaryType != null)
                                                VariableList.Add(new TypeSelection2Class("PhantomMaterial", DictionaryType.ItemCode));
                                        }

                                        if (int.TryParse(GlobalClass.GetCellValue(ThisSheet, "B78"), out iIndex))
                                            VariableList.Add(new TypeSelection2Class("ChamberIrradiationFieldSize1", iIndex.ToString()));

                                        if (int.TryParse(GlobalClass.GetCellValue(ThisSheet, "B79"), out iIndex))
                                            VariableList.Add(new TypeSelection2Class("ChamberIrradiationFieldSize2", iIndex.ToString()));

                                        if (int.TryParse(GlobalClass.GetCellValue(ThisSheet, "B80"), out iIndex))
                                            VariableList.Add(new TypeSelection2Class("ChamberIrradiationDistance", iIndex.ToString()));


                                        if (int.TryParse(GlobalClass.GetCellValue(ThisSheet, "B81"), out iIndex))
                                        {
                                            if (iIndex == 1)
                                                VariableList.Add(new TypeSelection2Class("ChamberIrradiationDistanceType", "SSD"));
                                            else if (iIndex == 2)
                                                VariableList.Add(new TypeSelection2Class("ChamberIrradiationDistanceType", "SAD"));
                                            else
                                                VariableList.Add(new TypeSelection2Class("ChamberIrradiationDistanceType", "Off"));
                                        }

                                        if (int.TryParse(GlobalClass.GetCellValue(ThisSheet, "B82"), out iIndex))
                                        {
                                            DictionaryTypeClass DictionaryType = GlobalClass.Dictionary.GetDictionaryType(GlobalClass.Dictionary.GetDictionary("ChamberIrradiationMeasuringPoint"), iIndex - 1);
                                            if (DictionaryType != null)
                                                VariableList.Add(new TypeSelection2Class("ChamberIrradiationMeasuringPoint", DictionaryType.ItemCode));
                                        }

                                        if (int.TryParse(GlobalClass.GetCellValue(ThisSheet, "B83"), out iIndex))
                                            VariableList.Add(new TypeSelection2Class("ChamberIrradiationDepth", iIndex.ToString()));

                                        if (isCobalt == true)
                                        {
                                            if (int.TryParse(GlobalClass.GetCellValue(ThisSheet, "B84"), out iIndex))
                                            {
                                                DictionaryTypeClass DictionaryType = GlobalClass.Dictionary.GetDictionaryType(GlobalClass.Dictionary.GetDictionary("CapMaterial"), iIndex - 1);
                                                if (DictionaryType != null)
                                                    VariableList.Add(new TypeSelection2Class("CapMaterial", DictionaryType.ItemCode));
                                            }

                                            if (double.TryParse(GlobalClass.GetCellValue(ThisSheet, "B85"), out dValue))
                                                VariableList.Add(new TypeSelection2Class("CapThickness", dValue.ToString()));
                                        }

                                        if (double.TryParse(GlobalClass.GetCellValue(ThisSheet, "B86").Replace("nC", "").Trim(), System.Globalization.NumberStyles.Any, null, out dValue))
                                            VariableList.Add(new TypeSelection2Class("ReadingUncorrected", String.Format("{0:F20}", dValue)));



                                        if (double.TryParse(GlobalClass.GetCellValue(ThisSheet, "B87"), out dValue))
                                            VariableList.Add(new TypeSelection2Class("ReadingMeasurementSetting", dValue.ToString()));

                                        if (double.TryParse(GlobalClass.GetCellValue(ThisSheet, "B88"), out dValue))
                                            VariableList.Add(new TypeSelection2Class("ReadingTemperature", dValue.ToString()));

                                        if (double.TryParse(GlobalClass.GetCellValue(ThisSheet, "B89"), out dValue))
                                            VariableList.Add(new TypeSelection2Class("ReadingPressure", dValue.ToString()));

                                        if (int.TryParse(GlobalClass.GetCellValue(ThisSheet, "B90"), out iIndex))
                                        {
                                            DictionaryTypeClass DictionaryType = GlobalClass.Dictionary.GetDictionaryType(GlobalClass.Dictionary.GetDictionary("ReadingPressureUnit"), iIndex - 1);
                                            if (DictionaryType != null)
                                                VariableList.Add(new TypeSelection2Class("ReadingPressureUnit", DictionaryType.ItemCode));
                                        }


                                        if (isCobalt == true)
                                        {
                                            VariableList.Add(new TypeSelection2Class("IrradiationUnits1", "min"));
                                            VariableList.Add(new TypeSelection2Class("IrradiationUnits2", "min"));
                                            VariableList.Add(new TypeSelection2Class("IrradiationUnits3", "min"));
                                            VariableList.Add(new TypeSelection2Class("CorrectionUnits", "s"));
                                            VariableList.Add(new TypeSelection2Class("ReadingMeasurementSettingUnits", "min"));
                                        }
                                        else
                                        {
                                            VariableList.Add(new TypeSelection2Class("IrradiationUnits1", "MU"));
                                            VariableList.Add(new TypeSelection2Class("IrradiationUnits2", "MU"));
                                            VariableList.Add(new TypeSelection2Class("IrradiationUnits3", "MU"));
                                            VariableList.Add(new TypeSelection2Class("CorrectionUnits", "MU"));
                                            VariableList.Add(new TypeSelection2Class("ReadingMeasurementSettingUnits", "MU"));
                                        }
                                    }
                                }
                            }
                            catch (Exception oEx)
                            {
                                MessageBox.Show("Import TLD Data Sheet Error. " + oEx.Message, "Import TLD Data Sheet Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                            finally
                            {
                                ThisWorkBook.Close(false, oMissing, oMissing);

                                ThisApplication.Quit();

                                //GlobalClass.releaseObject(ThisSheet);
                                GlobalClass.releaseObject(ThisWorkBook);
                                GlobalClass.releaseObject(ThisApplication);
                            }

                        }
                        catch (Exception oEx)
                        {
                            MessageBox.Show("Open TLD Data Sheet Error. " + oEx.Message, "Import TLD Data Sheet Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }
            }
            return VariableList;
        }

        public static List<TypeSelection2Class> LoadVariableListCertificateIDEAAutoDetect(string sFileName)
        {
            List<TypeSelection2Class> VariableList = new List<TypeSelection2Class>();
            VariableList.Clear();

            if (File.Exists(sFileName))
            {
                Excel.Application ThisApplication = new Excel.ApplicationClass();
                Excel.Workbook ThisWorkBook = ThisApplication.Workbooks.Open(sFileName, 0, false, 5, System.Reflection.Missing.Value, System.Reflection.Missing.Value, false, System.Reflection.Missing.Value, System.Reflection.Missing.Value, true, false, System.Reflection.Missing.Value, false, false, false);//Open the excel sheet                

                Excel.Worksheet ThisSheet = (Excel.Worksheet)ThisWorkBook.Sheets["Sheet1"]; //Select first sheet

                if (ThisSheet != null)
                {
                    //Set section
                    VariableList.Add(GlobalClass.EvaluationVariable(ThisSheet, "BatchNo", "B3"));
                    VariableList.Add(GlobalClass.EvaluationVariable(ThisSheet, "SetNo", "B5"));

                    // Evaluation section
                    VariableList.Add(GlobalClass.EvaluationVariable(ThisSheet, "ReaderFileName", "B15"));
                    VariableList.Add(GlobalClass.EvaluationVariable(ThisSheet, "ReaderFilePath", "B16"));
                    VariableList.Add(GlobalClass.EvaluationVariable(ThisSheet, "NoOfSetsInLoader", "B17"));
                    VariableList.Add(GlobalClass.EvaluationVariable(ThisSheet, "PositionInLoader", "B18"));

                    TypeSelection2Class VariableReferenceCapsule1 = GlobalClass.EvaluationVariable(ThisSheet, "ReferenceCapsule1", "F2");

                    if (VariableReferenceCapsule1.Type.Trim() != string.Empty)
                    {
                        // SSDL
                        VariableList.Add(new TypeSelection2Class("HolderCorrectionID", "3"));

                        VariableList.Add(GlobalClass.EvaluationVariable(ThisSheet, "ReferenceCapsule1TCBNo", "G2"));
                        VariableList.Add(GlobalClass.EvaluationVariable(ThisSheet, "ReferenceCapsule2TCBNo", "H2"));

                        VariableList.Add(GlobalClass.EvaluationVariable(ThisSheet, "ReferenceCapsule1TCBValue", "G3"));
                        VariableList.Add(GlobalClass.EvaluationVariable(ThisSheet, "ReferenceCapsule2TCBValue", "H3"));
                    }
                    else
                    {
                        // REST
                        VariableList.Add(new TypeSelection2Class("HolderCorrectionID", "2"));

                        VariableList.Add(GlobalClass.EvaluationVariable(ThisSheet, "ReferenceCapsule1TCBNo", "F2"));
                        VariableList.Add(GlobalClass.EvaluationVariable(ThisSheet, "ReferenceCapsule2TCBNo", "G2"));
                        VariableList.Add(GlobalClass.EvaluationVariable(ThisSheet, "ReferenceCapsule3TCBNo", "H2"));

                        VariableList.Add(GlobalClass.EvaluationVariable(ThisSheet, "ReferenceCapsule1TCBValue", "F3"));
                        VariableList.Add(GlobalClass.EvaluationVariable(ThisSheet, "ReferenceCapsule2TCBValue", "G3"));
                        VariableList.Add(GlobalClass.EvaluationVariable(ThisSheet, "ReferenceCapsule3TCBValue", "H3"));
                    }

                    VariableList.Add(GlobalClass.EvaluationVariableDate(ThisSheet, "ReadingDate", "H5"));
                    VariableList.Add(GlobalClass.EvaluationVariableDate(ThisSheet, "RCIrradiationDate", "H4"));
                    VariableList.Add(GlobalClass.EvaluationVariableDate(ThisSheet, "IrradiationDate", "H11"));

                    VariableList.Add(GlobalClass.EvaluationVariable(ThisSheet, "CertificateNo", "B5"));
                    VariableList.Add(GlobalClass.EvaluationVariable(ThisSheet, "RadiationUnit", "H7"));
                    //VariableList.Add(GlobalClass.EvaluationVariable(ThisSheet, "OperatorName", "B8"));
                    //VariableList.Add(GlobalClass.EvaluationVariable(ThisSheet, "Street", "B10"));
                    //VariableList.Add(GlobalClass.EvaluationVariable(ThisSheet, "City", "B11"));
                    //VariableList.Add(GlobalClass.EvaluationVariable(ThisSheet, "Country", "C12"));

                    VariableList.Add(GlobalClass.EvaluationVariable(ThisSheet, "IrradiatedByPersonFamilyName", "B12"));

                    VariableList.Add(GlobalClass.EvaluationVariable(ThisSheet, "UserDose1", "H12"));
                    VariableList.Add(GlobalClass.EvaluationVariable(ThisSheet, "UserDose2", "H13"));
                    VariableList.Add(GlobalClass.EvaluationVariable(ThisSheet, "UserDose3", "H14"));

                    VariableList.Add(GlobalClass.EvaluationVariable(ThisSheet, "DoseMeasured1", "C96"));
                    VariableList.Add(GlobalClass.EvaluationVariable(ThisSheet, "DoseMeasured2", "C97"));
                    VariableList.Add(GlobalClass.EvaluationVariable(ThisSheet, "DoseMeasured3", "C98"));
                    VariableList.Add(GlobalClass.EvaluationVariable(ThisSheet, "DoseMeasuredAvarage", "C99"));

                    VariableList.Add(GlobalClass.EvaluationVariable(ThisSheet, "DoseStated1", "D96"));
                    VariableList.Add(GlobalClass.EvaluationVariable(ThisSheet, "DoseStated2", "D97"));
                    VariableList.Add(GlobalClass.EvaluationVariable(ThisSheet, "DoseStated3", "D98"));

                    if (Path.GetFileName(sFileName).Contains("P"))
                    {
                        //VariableList.Add(GlobalClass.EvaluationVariable(ThisSheet, "DeviationPrimary1", "E96"));
                        //VariableList.Add(GlobalClass.EvaluationVariable(ThisSheet, "DeviationPrimary2", "E97"));
                        //VariableList.Add(GlobalClass.EvaluationVariable(ThisSheet, "DeviationPrimary3", "E98"));
                        //VariableList.Add(GlobalClass.EvaluationVariable(ThisSheet, "DeviationAvaragePrimary", "E99"));
                    }
                    else
                    {
                        VariableList.Add(GlobalClass.EvaluationVariable(ThisSheet, "DeviationSSDL1", "E96"));
                        VariableList.Add(GlobalClass.EvaluationVariable(ThisSheet, "DeviationSSDL2", "E97"));
                        VariableList.Add(GlobalClass.EvaluationVariable(ThisSheet, "DeviationSSDL3", "E98"));
                        VariableList.Add(GlobalClass.EvaluationVariable(ThisSheet, "DeviationAvarageSSDL", "E99"));
                    }
                    VariableList.Add(GlobalClass.EvaluationVariable(ThisSheet, "MeasuredStatedRatio1", "F96"));
                    VariableList.Add(GlobalClass.EvaluationVariable(ThisSheet, "MeasuredStatedRatio2", "F97"));
                    VariableList.Add(GlobalClass.EvaluationVariable(ThisSheet, "MeasuredStatedRatio3", "F98"));

                    VariableList.Add(GlobalClass.EvaluationVariable(ThisSheet, "Avarage", "F99"));



                    VariableList.Add(GlobalClass.EvaluationVariable(ThisSheet, "Background", "F53"));
                    VariableList.Add(GlobalClass.EvaluationVariable(ThisSheet, "BackgroundSD", "G53"));

                    VariableList.Add(GlobalClass.EvaluationVariable(ThisSheet, "MeanControl", "F56"));
                    VariableList.Add(GlobalClass.EvaluationVariable(ThisSheet, "MeanControlSD", "G56"));

                    VariableList.Add(GlobalClass.EvaluationVariable(ThisSheet, "SentControl", "F67"));
                    VariableList.Add(GlobalClass.EvaluationVariable(ThisSheet, "SentControlSD", "G67"));

                    VariableList.Add(GlobalClass.EvaluationVariable(ThisSheet, "RefDoseReadout1", "F58"));
                    VariableList.Add(GlobalClass.EvaluationVariable(ThisSheet, "RefDoseReadout2", "F59"));
                    //VariableList.Add(GlobalClass.EvaluationVariable(ThisSheet, "RefDoseReadout3", "XXX"));
                    VariableList.Add(GlobalClass.EvaluationVariable(ThisSheet, "RefDoseReadoutSD1", "G58"));
                    VariableList.Add(GlobalClass.EvaluationVariable(ThisSheet, "RefDoseReadoutSD2", "G59"));
                    //VariableList.Add(GlobalClass.EvaluationVariable(ThisSheet, "RefDoseReadoutSD3", "XXX"));

                    VariableList.Add(GlobalClass.EvaluationVariable(ThisSheet, "UserCapsuleReadout1", "F64"));
                    VariableList.Add(GlobalClass.EvaluationVariable(ThisSheet, "UserCapsuleReadout2", "F65"));
                    VariableList.Add(GlobalClass.EvaluationVariable(ThisSheet, "UserCapsuleReadout3", "F66"));

                    VariableList.Add(GlobalClass.EvaluationVariable(ThisSheet, "UserCapsuleReadoutSD1", "G64"));
                    VariableList.Add(GlobalClass.EvaluationVariable(ThisSheet, "UserCapsuleReadoutSD2", "G65"));
                    VariableList.Add(GlobalClass.EvaluationVariable(ThisSheet, "UserCapsuleReadoutSD3", "G66"));

                    VariableList.Add(GlobalClass.EvaluationVariable(ThisSheet, "MeanRefMinBckgDivDR", "F62"));
                    VariableList.Add(GlobalClass.EvaluationVariable(ThisSheet, "MeanRefMinBckgDivDRSD", "G62"));


                    //Variable.Type = "UserDose3"; Variable.TypeDescription = GlobalClass.GetCellValue(ThisSheet, "XX");
                    //Variable.Type = "AirKerma1"; Variable.TypeDescription = GlobalClass.GetCellValue(ThisSheet, "XX");
                    //Variable.Type = "AirKerma2"; Variable.TypeDescription = GlobalClass.GetCellValue(ThisSheet, "XX");
                    //Variable.Type = "AirKerma3"; Variable.TypeDescription = GlobalClass.GetCellValue(ThisSheet, "XX");
                    //Variable.Type = "DoseMeasured3"; Variable.TypeDescription = GlobalClass.GetCellValue(ThisSheet, "XX");
                    //Variable.Type = "DeviationPrimary1"; Variable.TypeDescription = GlobalClass.GetCellValue(ThisSheet, "XX");
                    //Variable.Type = "DeviationPrimary2"; Variable.TypeDescription = GlobalClass.GetCellValue(ThisSheet, "XX");
                    //Variable.Type = "DeviationPrimary3"; Variable.TypeDescription = GlobalClass.GetCellValue(ThisSheet, "XX");
                    //Variable.Type = "DeviationAvaragePrimary"; Variable.TypeDescription = GlobalClass.GetCellValue(ThisSheet, "XX");

                    //Variable.Type = "MeasuredStatedRatio1"; Variable.TypeDescription = GlobalClass.GetCellValue(ThisSheet, "XX");
                    //Variable.Type = "MeasuredStatedRatio2"; Variable.TypeDescription = GlobalClass.GetCellValue(ThisSheet, "XX");
                    //Variable.Type = "MeasuredStatedRatio3"; Variable.TypeDescription = GlobalClass.GetCellValue(ThisSheet, "XX");
                    //Variable.Type = "Background"; Variable.TypeDescription = GlobalClass.GetCellValue(ThisSheet, "XX");
                    //Variable.Type = "BackgroundSD"; Variable.TypeDescription = GlobalClass.GetCellValue(ThisSheet, "XX");
                    //Variable.Type = "MeanControl"; Variable.TypeDescription = GlobalClass.GetCellValue(ThisSheet, "XX");
                    //Variable.Type = "MeanControlSD"; Variable.TypeDescription = GlobalClass.GetCellValue(ThisSheet, "XX");
                    //Variable.Type = "SentControl"; Variable.TypeDescription = GlobalClass.GetCellValue(ThisSheet, "XX");
                    //Variable.Type = "SentControlSD"; Variable.TypeDescription = GlobalClass.GetCellValue(ThisSheet, "XX");
                    //Variable.Type = "RefDoseReadout1"; Variable.TypeDescription = GlobalClass.GetCellValue(ThisSheet, "XX");
                    //Variable.Type = "RefDoseReadout2"; Variable.TypeDescription = GlobalClass.GetCellValue(ThisSheet, "XX");
                    //Variable.Type = "RefDoseReadout3"; Variable.TypeDescription = GlobalClass.GetCellValue(ThisSheet, "XX");
                    //Variable.Type = "RefDoseReadoutSD1"; Variable.TypeDescription = GlobalClass.GetCellValue(ThisSheet, "XX");
                    //Variable.Type = "RefDoseReadoutSD2"; Variable.TypeDescription = GlobalClass.GetCellValue(ThisSheet, "XX");
                    //Variable.Type = "RefDoseReadoutSD3"; Variable.TypeDescription = GlobalClass.GetCellValue(ThisSheet, "XX");
                    //Variable.Type = "UserCapsuleReadout1"; Variable.TypeDescription = GlobalClass.GetCellValue(ThisSheet, "XX");
                    //Variable.Type = "UserCapsuleReadout2"; Variable.TypeDescription = GlobalClass.GetCellValue(ThisSheet, "XX");
                    //Variable.Type = "UserCapsuleReadout3"; Variable.TypeDescription = GlobalClass.GetCellValue(ThisSheet, "XX");
                    //Variable.Type = "UserCapsuleReadoutSD1"; Variable.TypeDescription = GlobalClass.GetCellValue(ThisSheet, "XX");
                    //Variable.Type = "UserCapsuleReadoutSD2"; Variable.TypeDescription = GlobalClass.GetCellValue(ThisSheet, "XX");
                    //Variable.Type = "MeanRefMinBckgDivDR"; Variable.TypeDescription = GlobalClass.GetCellValue(ThisSheet, "XX");
                    //Variable.Type = "MeanRefMinBckgDivDRSD"; Variable.TypeDescription = GlobalClass.GetCellValue(ThisSheet, "XX");
                    //Variable.Type = "CertificateComment"; Variable.TypeDescription = GlobalClass.GetCellValue(ThisSheet, "XX");
                    //Variable.Type = "CreatedOn"; Variable.TypeDescription = GlobalClass.GetCellValue(ThisSheet, "XX");
                    //Variable.Type = "CreatedBy"; Variable.TypeDescription = GlobalClass.GetCellValue(ThisSheet, "XX");
                    //Variable.Type = "LastUpdate"; Variable.TypeDescription = GlobalClass.GetCellValue(ThisSheet, "XX");
                    //Variable.Type = "UpdateComment"; Variable.TypeDescription = GlobalClass.GetCellValue(ThisSheet, "XX");
 
                    
                }

                ThisWorkBook.Close(false, System.Reflection.Missing.Value, System.Reflection.Missing.Value);

                ThisApplication.Quit();

                GlobalClass.releaseObject(ThisSheet);
                GlobalClass.releaseObject(ThisWorkBook);
                GlobalClass.releaseObject(ThisApplication);
            }

            return VariableList;
        }

        public static List<TypeSelection2Class> LoadVariableListEvaluation(string sFileName)
        {
            List<TypeSelection2Class> VariableList = new List<TypeSelection2Class>();
            VariableList.Clear();

            if (File.Exists(sFileName))
            {
                Excel.Application ThisApplication = new Excel.ApplicationClass();
                Excel.Workbook ThisWorkBook = ThisApplication.Workbooks.Open(sFileName, 0, false, 5, System.Reflection.Missing.Value, System.Reflection.Missing.Value, false, System.Reflection.Missing.Value, System.Reflection.Missing.Value, true, false, System.Reflection.Missing.Value, false, false, false);//Open the excel sheet

                //Excel.Sheets ThisSheet = ThisWorkBook.Sheets; //Get the sheets from workbook

                foreach (Excel.Worksheet ThisSheet in ThisWorkBook.Application.Worksheets)
                {
                    if (ThisSheet != null)
                    {
                        if ((ThisSheet.Name == "InputData") || (ThisSheet.Name == "input data"))
                        {
                            for (int iPosition = 1; iPosition <= GlobalClass.MaxVariableCount; iPosition++)
                            {
                                if (ThisSheet.get_Range("B" + iPosition + ":B" + iPosition, Type.Missing).Cells.Value2 != null)
                                {
                                    TypeSelection2Class Variable = new TypeSelection2Class();
                                    Variable.Type = ThisSheet.get_Range("B" + iPosition + ":B" + iPosition, Type.Missing).Cells.Value2.ToString().Trim();
                                    if (ThisSheet.get_Range("C" + iPosition + ":C" + iPosition, Type.Missing).Cells.Value2 != null)
                                        Variable.TypeDescription = ThisSheet.get_Range("C" + iPosition + ":C" + iPosition, Type.Missing).Cells.Value2.ToString().Trim();
                                    VariableList.Add(Variable);
                                }
                            }

                        }
                        else if ((ThisSheet.Name == "OutputData") || (ThisSheet.Name == "output data"))
                        {
                            for (int iPosition = 1; iPosition <= GlobalClass.MaxVariableCount; iPosition++)
                            {
                                if (ThisSheet.get_Range("B" + iPosition + ":B" + iPosition, Type.Missing).Cells.Value2 != null)
                                {
                                    TypeSelection2Class Variable = new TypeSelection2Class();
                                    Variable.Type = ThisSheet.get_Range("B" + iPosition + ":B" + iPosition, Type.Missing).Cells.Value2.ToString().Trim();
                                    if (ThisSheet.get_Range("C" + iPosition + ":C" + iPosition, Type.Missing).Cells.Value2 != null)
                                        Variable.TypeDescription = ThisSheet.get_Range("C" + iPosition + ":C" + iPosition, Type.Missing).Cells.Value2.ToString().Trim();
                                    VariableList.Add(Variable);
                                }
                            }
                        }
                    }
                }


                ThisWorkBook.Close(false, System.Reflection.Missing.Value, System.Reflection.Missing.Value);

                ThisApplication.Quit();

                //GlobalClass.releaseObject(ThisSheet);
                GlobalClass.releaseObject(ThisWorkBook);
                GlobalClass.releaseObject(ThisApplication);

            }

            return VariableList;
        }

        public static List<TypeSelection2Class> LoadVariableListPDF(string sFileName)
        {
            List<TypeSelection2Class> VariableList = new List<TypeSelection2Class>();
            VariableList.Clear();

            if (File.Exists(sFileName))
            {
                try
                {
                    PdfReader pdfReader = new PdfReader(sFileName);

                    foreach (DictionaryEntry de in pdfReader.AcroFields.Fields)
                    {
                        string sCurrentField = de.Key.ToString().Trim();
                        string sCurrentValue = pdfReader.AcroFields.GetField(de.Key.ToString().Trim());

                        // Ignore these fields
                        if ((sCurrentField == "OperatorID") || (sCurrentField == "TLDDataID") ||
                            (sCurrentField == "LastUpdate") || (sCurrentField == "UpdateComment"))
                            continue;

                        if (sCurrentField == "EmailInstitutional")
                            sCurrentField = "InstitutionalEmail";
                        else if (sCurrentField == "FaxInstitutional")
                            sCurrentField = "InstitutionalFax";
                        else if (sCurrentField == "TelephoneInstitutional")
                            sCurrentField = "InstitutionalTelephone1";

                        else if (sCurrentField == "BeamType")
                        {
                            sCurrentField = "SetBeamType";
                            if (sCurrentValue == GlobalClass.sUnitTypeAccelerator)
                                sCurrentValue = GlobalClass.sBeamTypePhoton;
                        }
                        else if ((sCurrentField == "BeamType1") || (sCurrentField == "BeamType2") || (sCurrentField == "BeamType3") || (sCurrentField == "BeamType4") || (sCurrentField == "BeamType5"))
                        {
                            if (sCurrentValue == GlobalClass.sUnitTypeAccelerator)
                                sCurrentValue = GlobalClass.sBeamTypePhoton;
                        }

                        VariableList.Add(new TypeSelection2Class(sCurrentField, sCurrentValue));
                    }
                    pdfReader.Close();
                }
                catch
                {
                }
            }
            return VariableList;
        }

        public static string LoadSignatureListDatabaseIDEA(TLDSetClass TLDSet)
        {
            string sReturn = string.Empty;

            if (TLDSet != null)
            {
                if (TLDSet.Certificate != null)
                {
                    SignatureClass Signature1 = TLDSet.Certificate.GetSignature(GlobalClass.iSignatureTLDCertificateSignByOfficer); // 41
                    SignatureClass Signature2 = TLDSet.Certificate.GetSignature(GlobalClass.iSignatureTLDCertificateSignBySectionHead); // 42
                    //SignatureClass Signature3 = TLDSet.Certificate.GetSignature(GlobalClass.iSignatureTLDCertificateDispatched); // 43

                    string sTable = "IDEA.dbo.FullReportTLDCertificates";
                    DataTable IDEATLDCertificatesTable = GlobalClass.GetYPDataTable("TLDCertificates", "SELECT * FROM " + sTable + " WHERE CertificateNo = '" + TLDSet.SetNo + "'");

                    foreach (DataRow dr in IDEATLDCertificatesTable.Rows)
                    {
                        int iSignatureType = (int)dr["SignatureType"];
                        int iUserID = (int)dr["UserID"];

                        if (iUserID == 116)
                            iUserID = 117; // Signed by user: [flory] 
                        else if (iUserID == 128)
                            iUserID = 147; // Signed by user: [PaduaS] 
                        else if (iUserID == 131)
                            iUserID = 150; // Signed by user: [vandermerd] 
                        else if (iUserID == 120)
                            iUserID = 146; // Signed by user: [pirkfellna] 
                                    
                        if (iSignatureType == 1)
                        {
                            if (Signature1 == null)
                            {
                                Signature1 = new SignatureClass();
                                Signature1.SignatureType = GlobalClass.iSignatureTLDCertificateSignByOfficer;
                                Signature1.SignatureDate = (DateTime)dr["SignatureDate"];
                                Signature1.OperatorID = TLDSet.OperatorID;
                                Signature1.DocumentID = TLDSet.Certificate.CertificateID;
                                Signature1.UserID = iUserID;
                                Signature1.SignatureByUser = (string)dr["SignatureByUser"];
                                Signature1.SignatureDetails = (string)dr["SignatureDetails"];
                                //Signature1.SignatureImage = (byte[])dr["SignatureImage"];

                                TLDSet.Certificate.SignatureList.Add(Signature1);
                                Signature1.SaveSignature();
                            }
                        }
                        else if (iSignatureType == 2)
                        {
                            if (Signature2 == null)
                            {
                                Signature2 = new SignatureClass();
                                Signature2.SignatureType = GlobalClass.iSignatureTLDCertificateSignBySectionHead;
                                Signature2.SignatureDate = (DateTime)dr["SignatureDate"];
                                Signature2.OperatorID = TLDSet.OperatorID;
                                Signature2.DocumentID = TLDSet.Certificate.CertificateID;
                                Signature2.UserID = iUserID;
                                Signature2.SignatureByUser = (string)dr["SignatureByUser"];
                                Signature2.SignatureDetails = (string)dr["SignatureDetails"];
                                //Signature2.SignatureImage = (byte[])dr["SignatureImage"];

                                TLDSet.Certificate.SignatureList.Add(Signature2);
                                Signature2.SaveSignature();
                            }
                        }
                        else if (iSignatureType == 3)
                        {
                            /*
                            if (Signature3 == null)
                            {
                                Signature3 = new SignatureClass();
                                Signature3.SignatureType = GlobalClass.iSignatureTLDCertificateDispatched;
                                Signature3.SignatureDate = (DateTime)dr["SignatureDate"];
                                Signature3.OperatorID = TLDSet.OperatorID;
                                Signature3.DocumentID = TLDSet.Certificate.CertificateID;
                                Signature3.UserID = iUserID;
                                Signature3.SignatureByUser = (string)dr["SignatureByUser"];
                                Signature3.SignatureDetails = (string)dr["SignatureDetails"];
                                //Signature3.SignatureImage = (byte[])dr["SignatureImage"];

                                TLDSet.Certificate.SignatureList.Add(Signature3);
                                Signature3.SaveSignature();
                            }
                             * */
                        }
                    }
                }
            }

            return sReturn;
        }

        public static string LoadAttachmentListDatabaseIDEA(TLDSetClass TLDSet)
        {
            string sReturn = string.Empty;

            if (TLDSet != null)
            {
                if (TLDSet.Certificate != null)
                {

                    string sTable = "IDEA.dbo.FullReportTLDCertificates";
                    DataTable IDEATLDCertificatesTable = GlobalClass.GetYPDataTable("TLDCertificates", "SELECT * FROM " + sTable + " WHERE CertificateNo = '" + TLDSet.SetNo + "'");

                    foreach (DataRow dr in IDEATLDCertificatesTable.Rows)
                    {
                        // Check Attachments
                        AttachmentClass Attachment1 = TLDSet.Certificate.GetAttachment(GlobalClass.iAttachmentTLDCertificate);
                        AttachmentClass Attachment2 = TLDSet.Certificate.GetAttachment(GlobalClass.iAttachmentEvaluationSheetVeryfied);

                        if (Attachment1 == null)
                        {
                            Attachment1 = new AttachmentClass(GlobalClass.iAttachmentTLDCertificate);
                            Attachment1.AttachmentType = GlobalClass.iAttachmentTLDCertificate;
                            Attachment1.AttachmentDate = (DateTime)dr["PackageSignatureDate"];
                            Attachment1.OperatorID = TLDSet.OperatorID;
                            Attachment1.DocumentID = TLDSet.Certificate.CertificateID;
                            Attachment1.AttachmentDetails = (string)dr["PackageSignatureByUser"];
                            Attachment1.AttachmentFileName = (string)dr["CertificateTLDFileName"].ToString().Replace(".xls", ".pdf");
                            Attachment1.AttachmentForm = (byte[])dr["CertificateTLDPDF"];

                            Attachment1.CreatedOn = (DateTime)dr["PackageSignatureDate"];
                            Attachment1.CreatedBy = (string)dr["PackageSignatureByUser"];
                            Attachment1.LastUpdate = (DateTime)dr["PackageSignatureDate"];
                            Attachment1.UpdateComment = (string)dr["PackageSignatureByUser"];

                            TLDSet.Certificate.AttachmentList.Add(Attachment1);
                            Attachment1.SaveAttachment();
                        }

                        if (Attachment2 == null)
                        {
                            Attachment2 = new AttachmentClass(GlobalClass.iAttachmentEvaluationSheetVeryfied);
                            Attachment2.AttachmentType = GlobalClass.iAttachmentEvaluationSheetVeryfied;
                            Attachment2.AttachmentDate = (DateTime)dr["PackageSignatureDate"];
                            Attachment2.OperatorID = TLDSet.OperatorID;
                            Attachment2.DocumentID = TLDSet.Certificate.CertificateID;
                            Attachment2.AttachmentDetails = (string)dr["PackageSignatureByUser"];
                            Attachment2.AttachmentFileName = (string)dr["CertificateTLDFileName"];
                            Attachment2.AttachmentForm = (byte[])dr["CertificateTLDPDF"];

                            Attachment2.CreatedOn = (DateTime)dr["PackageSignatureDate"];
                            Attachment2.CreatedBy = (string)dr["PackageSignatureByUser"];
                            Attachment2.LastUpdate = (DateTime)dr["PackageSignatureDate"];
                            Attachment2.UpdateComment = (string)dr["PackageSignatureByUser"];

                            TLDSet.Certificate.AttachmentList.Add(Attachment2);
                            Attachment2.SaveAttachment();
                        }

                    }
                }
            }

            return sReturn;
        }

        public static List<TypeSelection2Class> LoadVariableListDatabaseIDEA(TLDSetClass TLDSet)
        {
            List<TypeSelection2Class> VariableList = new List<TypeSelection2Class>();
            VariableList.Clear();
            if (TLDSet != null)
            {
                string sTable = "dbo.FullReportIDEATLDSets";

                //if (TLDSet.SetType == 1)
                //    sTable = "dbo.FullReportIDEATLDSets";
                //else
                //    sTable = "dbo.FullReportIDEATLDFollowUpSets";

                DataTable ColumnInfoDataTable = GlobalClass.GetYPDataTable(sTable, "SELECT * FROM dbo.fn_TableColumnInfo('" + sTable + "') ORDER BY 2");
                DataTable IDEATLDSetDataTable = GlobalClass.GetYPDataTable("IDEASetNo", "SELECT * FROM " + sTable + " WHERE SetNo = '" + TLDSet.SetNo + "'");

                foreach (DataRow idr in IDEATLDSetDataTable.Rows)
                {
                    foreach (DataRow dr in ColumnInfoDataTable.Rows)
                    {
                        string sCurrentFieldType = (string)dr["TypeName"];
                        string sCurrentField = (string)dr["ColumnName"];
                        string sCurrentValue = string.Empty;

                        int iCurrentFieldLength = 0;
                        int iCurrentFieldScale = 0;


                        if (dr["ColumnPrecision"] != DBNull.Value)
                            iCurrentFieldLength = (Int16)dr["ColumnPrecision"];

                        if (dr["ColumnScale"] != DBNull.Value)
                            iCurrentFieldScale = (int)dr["ColumnScale"];

                        if (idr[sCurrentField] != DBNull.Value)
                        {
                            try
                            {
                                if (sCurrentFieldType == "nvarchar")
                                    sCurrentValue = GlobalClass.FormatStringValue((string)idr[sCurrentField], iCurrentFieldLength);
                                if (sCurrentFieldType == "varchar")
                                    sCurrentValue = GlobalClass.FormatStringValue((string)idr[sCurrentField], iCurrentFieldLength);
                                else if (sCurrentFieldType == "int")
                                    sCurrentValue = GlobalClass.FormatIntegerValue((int)idr[sCurrentField]);
                                else if (sCurrentFieldType == "numeric")
                                    sCurrentValue = Math.Round((decimal)idr[sCurrentField], iCurrentFieldScale).ToString().Trim();
                                else if (sCurrentFieldType == "decimal")
                                    sCurrentValue = Math.Round((decimal)idr[sCurrentField], iCurrentFieldScale).ToString().Trim();
                                else if (sCurrentFieldType == "float")
                                {
                                    iCurrentFieldScale = 7;
                                    decimal dValue = Convert.ToDecimal(idr[sCurrentField]);
                                    sCurrentValue = Math.Round(dValue, iCurrentFieldScale).ToString().Trim();

                                    //sCurrentValue = Math.Round((decimal)idr[sCurrentField], iCurrentFieldScale).ToString().Trim();
                                }
                                else if (sCurrentFieldType == "bit")
                                {
                                    if (Convert.ToBoolean((bool)idr[sCurrentField]) == true)
                                        sCurrentValue = "1";
                                    else
                                        sCurrentValue = "0";
                                }
                                else if (sCurrentFieldType == "datetime")
                                    sCurrentValue = Convert.ToDateTime((DateTime)idr[sCurrentField]).ToString().Trim();
                            }
                            catch
                            {
                                MessageBox.Show("Can not convert value " + sCurrentField + " Table " + sTable, "Conversion Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }

                            VariableList.Add(new TypeSelection2Class(sCurrentField, sCurrentValue));
                        }
                    }
                }
            }
            return VariableList;
        }

        public static TreeNode FindChildNode(TreeNode tvSelection, string sMatchNodeName)
        {
            foreach (TreeNode node in tvSelection.Nodes)
            {
                if (node.Name == sMatchNodeName)
                {
                    return node;
                }
                else
                {
                    TreeNode nodeChild = FindChildNode(node, sMatchNodeName);
                    if (nodeChild != null) return nodeChild;
                }
            }
            return (TreeNode)null;
        }

        public static Form IsFormAlreadyOpen(Type FormType)
        {
            foreach (Form OpenForm in Application.OpenForms)
            {
                if (OpenForm.GetType() == FormType)
                    return OpenForm;
            }

            return null;
        }

        public static string GenerateLabels(string sFileNameExcel, string sFileNameWord, string sExcelSql)
        {
            string sReturn = string.Empty;


            if (File.Exists(sFileNameExcel))
            {
                if (File.Exists(sFileNameWord))
                {
                    Object oMissing = System.Reflection.Missing.Value;
                    Object oFileNameWord = sFileNameWord;
                    Object oFalse = false;

                    Word._Application oWordApp = new Word.Application();
                    Word._Document oWordDoc = null;

                    oWordApp.Visible = false;

                    //ADDING A NEW DOCUMENT FROM A TEMPLATE
                    oWordDoc = oWordApp.Documents.Add(
                        /* ref object Template */ ref oFileNameWord,
                        /* ref object NewTemplate */ ref oMissing,
                        /* ref object DocumentType */ ref oMissing,
                        /* ref object Visible */ ref oMissing);

                    //SETTING THE FOCUES ON THE PAGE FOOTER
                    //oWordApp.ActiveWindow.ActivePane.View.SeekView = Word.WdSeekView.wdSeekCurrentPageFooter;
                    //oWordApp.Selection.HeaderFooter.Range.Select();
                    //oWordApp.Selection.HeaderFooter.Range.Delete(ref oMissing, ref oMissing);


                    //SETTING THE FOCUES ON THE EVEN PAGE HEADER
                    //oWordApp.ActiveWindow.ActivePane.View.SeekView = Word.WdSeekView.wdSeekCurrentPageHeader;
                    //oWordApp.Selection.HeaderFooter.Range.Select();
                    //oWordApp.Selection.HeaderFooter.Range.Delete(ref oMissing, ref oMissing);

                    //oWordApp.ActiveWindow.ActivePane.View.SeekView = Word.WdSeekView.wdSeekMainDocument;

                    oWordDoc.Select();

                    //Insert the mail merge fields temporarily so that you can use the range that contains the merge fields 
                    //as a layout for your labels -- to use this as a layout, you can add it as an AutoText entry.
                    // Justify the rest of the document.

                    //oWordApp.Selection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    //oWordApp.Selection.Font.Name = "Arial";
                    //oWordApp.Selection.Font.Size = 14;
                    //oWordApp.Selection.Font.Bold = 1;
                    //oWordApp.Selection.TypeParagraph();
                    //oWordDoc.MailMerge.Fields.Add(oWordApp.Selection.Range, "OperatorName");
                    //oWordApp.Selection.TypeParagraph();
                    //oWordApp.Selection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    //oWordApp.Selection.Font.Name = "Arial";
                    //oWordApp.Selection.Font.Size = 14;
                    //oWordDoc.MailMerge.Fields.Add(oWordApp.Selection.Range, "SetNos");
                    //oWordApp.Selection.TypeText(" ");

                    Word.AutoTextEntry oAutoText = oWordApp.NormalTemplate.AutoTextEntries.Add("MyLabelLayout", oWordDoc.Content);
                    //oWordDoc.Content.Delete();
                    //Merge fields in document no longer needed now that the AutoText entry for the label layout has been added so delete it.

                    //Set up the mail merge type as mailing labels and use a tab-delimited text file as the data source.
                    //oWordDoc.MailMerge.MainDocumentType = Word.WdMailMergeMainDocType.wdMailingLabels;
                    oWordDoc.MailMerge.OpenDataSource(sFileNameExcel, Word.WdOpenFormat.wdOpenFormatAuto,
                    oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, sExcelSql);

                    //Create the new document for the labels using the AutoText entry you added -- 5160 is the label number to use for this sample.
                    //You can specify the label number you want to use for the output in the Name argument.
                    /*
                    oWordApp.MailingLabel.CreateNewDocument("3653", //"5160", //ref Object Name
                                                            "", //ref Object Address,
                                                            "MyLabelLayout", //ref Object AutoText,
                                                            oMissing, //ref Object ExtractAddress,
                                                            Word.WdPaperTray.wdPrinterManualFeed,//ref Object LaserTray,
                                                            oMissing,//ref Object PrintEPostageLabel,
                                                            oMissing); //ref Object Vertical

                    */

                    //Execute the mail merge to generate the labels.
                    oWordDoc.MailMerge.Destination = Word.WdMailMergeDestination.wdSendToNewDocument;
                    oWordDoc.MailMerge.Execute(ref oFalse);

                    //Delete the AutoText entry you added
                    oAutoText.Delete();

                    oWordDoc.Close(ref oFalse, ref oMissing, ref oMissing);
                    oWordApp.Visible = true;

                    //Prevent save to Normal template when user exits Word
                    oWordApp.NormalTemplate.Saved = true;
                }
            }

            return sReturn;
        }

        public static string FormatExcelReport(string sFileNameExcel, string sSheetName, string sColumnsToDelete, string sTitle)
        {
            string sReturn = string.Empty;

            if (File.Exists(sFileNameExcel))
            {
                try
                {
                    object oMissing = System.Reflection.Missing.Value;
                    Excel.Application ThisApplication = new Excel.ApplicationClass();
                    ThisApplication.Visible = false;
                    Excel.Workbook ThisWorkBook = (Excel.Workbook)ThisApplication.Workbooks.Open(sFileNameExcel, oMissing, oMissing, 5, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
                    //Excel.Worksheet ThisSheet = (Excel.Worksheet)ThisWorkBook.Sheets[1]; //"Sheet1"

                    Excel.Worksheet ThisSheet = (Excel.Worksheet)ThisWorkBook.Sheets[sSheetName]; //Select first sheet
                    if (ThisSheet != null)
                    {
                        Excel.Range SheetColumns = ThisSheet.Columns;

                        string[] sColumns = sColumnsToDelete.Split(',');

                        foreach (string sColumnindex in sColumns)
                        {
                            int iColumnindex = -1;

                            if (int.TryParse(sColumnindex, out iColumnindex) == true)
                            {

                                Excel.Range Column = (Excel.Range)SheetColumns[iColumnindex, Missing.Value];
                                Column.Delete(Missing.Value);
                                GlobalClass.releaseObject(Column);
                            }
                        }

                        Excel.Range r = ThisSheet.get_Range("A1", "A1").EntireRow;
                        r.Insert(Excel.XlInsertShiftDirection.xlShiftDown);

                        ThisSheet.get_Range("A1" + ":A1", Type.Missing).Cells.Value2 = sTitle;
                        ThisSheet.get_Range("A1" + ":A1", Type.Missing).Font.Size = 14;
                        ThisSheet.get_Range("A1" + ":A1", Type.Missing).Font.Bold = true;
                        //ThisSheet.get_Range("A1:A1,C1:C1", Type.Missing).Merge(Type.Missing);
                        //ThisApplication.get_Range("A1:C1,A1:C1", Type.Missing).Merge(Type.Missing);


                        //Sets page setup properties
                        ThisSheet.PageSetup.Orientation = Microsoft.Office.Interop.Excel.XlPageOrientation.xlLandscape;
                        ThisSheet.PageSetup.PaperSize = Microsoft.Office.Interop.Excel.XlPaperSize.xlPaperA4;

                        ThisWorkBook.Save();
                        ThisWorkBook.Close(false, sFileNameExcel, Missing.Value);

                        ThisApplication.Quit();

                        
                        GlobalClass.releaseObject(SheetColumns);
                        GlobalClass.releaseObject(ThisSheet);
                        GlobalClass.releaseObject(ThisWorkBook);
                        GlobalClass.releaseObject(ThisApplication);
                    }
                }

                catch (Exception oEx)
                {
                    MessageBox.Show(oEx.Message, "Report Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            return sReturn;
        }

        public static string[] ImportValuesFromPDF(string sPDFFileName, string[] sFieldNames)
        {
            string[] sReturn = new string[sFieldNames.Length];

            if (File.Exists(sPDFFileName))
            {
                try
                {
                    PdfReader pdfReader = new PdfReader(sPDFFileName);

                    int iCounter = 0;
                    foreach (string sFieldName in sFieldNames)
                    {
                        sReturn[iCounter] = pdfReader.AcroFields.GetField(sFieldName).ToString();
                        iCounter = iCounter + 1;
                    }

                    pdfReader.Close();
                }
                catch
                {
                    sReturn = new string[0];
                }
            }
            return sReturn;
        }

        public static void DisablePDFField(AcroFields pdfFormFields, string sFieldNames)
        {
            /*
            Sets a field property. Valid property names are:
            <p>
            <ul>
            <li>textfont - sets the text font. The value for this entry is a <CODE>BaseFont</CODE>.<br>
            <li>textcolor - sets the text color. The value for this entry is a <CODE>java.awt.Color</CODE>.<br>
            <li>textsize - sets the text size. The value for this entry is a <CODE>Float</CODE>.
            <li>bgcolor - sets the background color. The value for this entry is a <CODE>java.awt.Color</CODE>.
                If <code>null</code> removes the background.<br>
            <li>bordercolor - sets the border color. The value for this entry is a <CODE>java.awt.Color</CODE>.
                If <code>null</code> removes the border.<br>
            </ul>
            @param field the field name
            @param name the property name
            @param value the property value
            @param inst an array of <CODE>int</CODE> indexing into <CODE>AcroField.Item.merged</CODE> elements to process.
            Set to <CODE>null</CODE> to process all
            @return <CODE>true</CODE> if the property exists, <CODE>false</CODE> otherwise
            */
            string sErrorString = string.Empty;

            if (pdfFormFields.SetFieldProperty(sFieldNames, "setfflags", PdfFormField.FF_READ_ONLY, null) == false)
                sErrorString = sErrorString + sFieldNames + "setfflags" + ",";
            if (pdfFormFields.SetFieldProperty(sFieldNames, "bgcolor", iTextSharp.text.Color.LIGHT_GRAY, null) == false)
                sErrorString = sErrorString+ sFieldNames + "bgcolor" + ",";
            //if (pdfFormFields.SetFieldProperty(sFieldNames, "textcolor", iTextSharp.text.Color.LIGHT_GRAY, null) == false)
            //    sErrorString = sErrorString + sFieldNames + "textcolor" + ",";

            pdfFormFields.RegenerateField(sFieldNames);
        }

        public static void SummaryCalibrationReport(string sOutputFile, DateTime dStartDate, DateTime dEndDate)
        {
            StatisticsClass Statistics = new StatisticsClass();

            Statistics.StatisticsID = 0;
            Statistics.StartDate = dStartDate;
            Statistics.EndDate = dEndDate;
            Statistics.StatisticsDescription = "Summary of calibrations performed from " + GlobalClass.FormatDateTimeValue(dStartDate) + " to " + GlobalClass.FormatDateTimeValue(dEndDate);

            Statistics.PopulateSummaryCalibrationTotals();
            Statistics.PopulateSummaryCalibrationIAEA();
            Statistics.PopulateSummaryCalibrationSSDL();

            string sLine1 = "Total of " + Statistics.TotalChambers.ToString() + " chambers calibrated from " + GlobalClass.FormatDateTimeValue(dStartDate) + " to " + GlobalClass.FormatDateTimeValue(dEndDate);
            string sLine2 = Statistics.TotalNonIAEAChambers.ToString() + " non-IAEA chambers from " + Statistics.TotalNonIAEASSDL.ToString() + " labs ";
            string sLine3 = Statistics.TotalIAEAChambers.ToString() + " IAEA chambers ( " + Statistics.TotalIAEAQuatroChambers.ToString() + " for Quatro)";

            // step 1: creating the document
            // Creates an instance of the iTextSharp.text.Document-object:
            Document pdfDocument = new Document(PageSize.A4);
            Cell CellItem;

            try
            {
                // step 2: creating the writer
                // Creates a Writer that listens to this document and writes the document to the Stream of your choice:
                PdfWriter pdfWriter = PdfWriter.GetInstance(pdfDocument, new FileStream(sOutputFile, FileMode.Create));

                // we Add a Header that will show up on PAGE 1
                HeaderFooter pdfHeader = new HeaderFooter(new Phrase("International Atomic Energy Agency \n The IAEA/WHO SSDL Network \n  " + Statistics.StatisticsDescription, GlobalClass.iReportFont14Blue), false);
                pdfHeader.Alignment = Element.ALIGN_CENTER;
                pdfHeader.Border = iTextSharp.text.Rectangle.BOTTOM_BORDER;
                pdfDocument.Header = pdfHeader;

                // we Add a Footer that will show up on PAGE 1
                HeaderFooter pdfFooter = new HeaderFooter(new Phrase(Statistics.StatisticsDescription + "\n Printed on " + DateTime.Now.ToString() + "\n Page ", GlobalClass.iReportFont8), true);
                pdfFooter.Alignment = Element.ALIGN_RIGHT;
                pdfFooter.Border = iTextSharp.text.Rectangle.TOP_BORDER;
                pdfDocument.Footer = pdfFooter;

                // step 3: initialisations + opening the document
                // Opens the document:
                pdfDocument.Open();

                pdfDocument.Add(new Phrase(sLine1, GlobalClass.iReportFont10)); pdfDocument.Add(new Phrase("\n", GlobalClass.iReportFont10));
                pdfDocument.Add(new Phrase(sLine2, GlobalClass.iReportFont10)); pdfDocument.Add(new Phrase("\n", GlobalClass.iReportFont10));
                pdfDocument.Add(new Phrase(sLine3, GlobalClass.iReportFont10)); pdfDocument.Add(new Phrase("\n", GlobalClass.iReportFont10));

                // Create a table and add it to the document
                Table InfoTable = new Table(4);
                InfoTable.Padding = 1;
                InfoTable.Spacing = 0;
                InfoTable.Width = 100;

                //datatable.setBorder(Rectangle.NO_BORDER);
                float[] InfoTableHeaderWidths = { 40, 20, 20, 20 };
                InfoTable.Widths = InfoTableHeaderWidths;
                InfoTable.BorderWidth = 0;

                InfoTable.DefaultHorizontalAlignment = Element.ALIGN_CENTER;
                InfoTable.DefaultVerticalAlignment = Element.ALIGN_MIDDLE;

                InfoTable.DefaultCellBackgroundColor = iTextSharp.text.Color.WHITE;
                InfoTable.DefaultCellBorderWidth = 1;

                InfoTable.DefaultCellBackgroundColor = iTextSharp.text.Color.WHITE;
                CellItem = new Cell(new Phrase("Summary of SSDL members calibrations", GlobalClass.iReportFont14Blue));
                CellItem.Colspan = 4;
                CellItem.BorderWidth = 1;
                InfoTable.AddCell(CellItem);

                InfoTable.DefaultCellBackgroundColor = new iTextSharp.text.Color(255, 255, 192);
                InfoTable.AddCell(new Phrase("", GlobalClass.iReportFont12));
                InfoTable.DefaultHorizontalAlignment = Element.ALIGN_CENTER;
                InfoTable.AddCell(new Phrase("Chambers", GlobalClass.iReportFont12));
                InfoTable.AddCell(new Phrase("cal. points", GlobalClass.iReportFont12));
                InfoTable.AddCell(new Phrase("Calibrations", GlobalClass.iReportFont12));
                InfoTable.EndHeaders();

                foreach (StatisticsDetailsClass StatisticsDetailsItem in Statistics.StatisticsDetailsSSDL)
                {
                    InfoTable.DefaultHorizontalAlignment = Element.ALIGN_LEFT;
                    if (StatisticsDetailsItem.CalibrationType == 0)
                    {
                        InfoTable.DefaultCellBackgroundColor = new iTextSharp.text.Color(192, 255, 255);
                        InfoTable.AddCell(new Phrase(StatisticsDetailsItem.ChamberTypeDescription, GlobalClass.iReportFont12));
                    }
                    else
                    {
                        if ((StatisticsDetailsItem.ChamberType == 3) || (StatisticsDetailsItem.ChamberType == 4))
                            continue;

                        InfoTable.DefaultHorizontalAlignment = Element.ALIGN_RIGHT;
                        InfoTable.DefaultCellBackgroundColor = iTextSharp.text.Color.WHITE;
                        InfoTable.AddCell(new Phrase(StatisticsDetailsItem.CalibrationTypeDescription, GlobalClass.iReportFont12));
                    }

                    InfoTable.DefaultHorizontalAlignment = Element.ALIGN_CENTER;
                    InfoTable.AddCell(new Phrase(StatisticsDetailsItem.TotalChambers.ToString(), GlobalClass.iReportFont12));
                    InfoTable.AddCell(new Phrase(StatisticsDetailsItem.TotalCalibrationPoints.ToString(), GlobalClass.iReportFont12));
                    InfoTable.AddCell(new Phrase(StatisticsDetailsItem.TotalCalibrations.ToString(), GlobalClass.iReportFont12));
                }
                // Totals
                InfoTable.DefaultCellBackgroundColor = new iTextSharp.text.Color(255, 255, 192);
                InfoTable.DefaultHorizontalAlignment = Element.ALIGN_RIGHT;
                InfoTable.AddCell(new Phrase("Total", GlobalClass.iReportFont12));
                InfoTable.DefaultHorizontalAlignment = Element.ALIGN_CENTER;
                InfoTable.AddCell(new Phrase(Statistics.TotalNonIAEAChambers.ToString(), GlobalClass.iReportFont12));
                InfoTable.AddCell(new Phrase(Statistics.TotalNonIAEASSDL.ToString(), GlobalClass.iReportFont12));
                InfoTable.AddCell(new Phrase(Statistics.TotalNonIAEACalibrations.ToString(), GlobalClass.iReportFont12));


                InfoTable.DefaultCellBackgroundColor = iTextSharp.text.Color.WHITE;
                CellItem = new Cell(new Phrase("Summary of IAEA calibrations", GlobalClass.iReportFont14Blue));
                CellItem.Colspan = 4;
                CellItem.BorderWidth = 1;
                InfoTable.AddCell(CellItem);

                InfoTable.DefaultCellBackgroundColor = new iTextSharp.text.Color(255, 255, 192);
                InfoTable.AddCell(new Phrase("", GlobalClass.iReportFont12));
                InfoTable.DefaultHorizontalAlignment = Element.ALIGN_CENTER;
                InfoTable.AddCell(new Phrase("Chambers", GlobalClass.iReportFont12));
                InfoTable.AddCell(new Phrase("cal. points", GlobalClass.iReportFont12));
                InfoTable.AddCell(new Phrase("Calibrations", GlobalClass.iReportFont12));

                foreach (StatisticsDetailsClass StatisticsDetailsItem in Statistics.StatisticsDetailsIAEA)
                {
                    InfoTable.DefaultHorizontalAlignment = Element.ALIGN_LEFT;
                    if (StatisticsDetailsItem.CalibrationType == 0)
                    {
                        InfoTable.DefaultCellBackgroundColor = new iTextSharp.text.Color(192, 255, 255);
                        InfoTable.AddCell(new Phrase(StatisticsDetailsItem.ChamberTypeDescription, GlobalClass.iReportFont12));
                    }
                    else
                    {
                        if ((StatisticsDetailsItem.ChamberType == 3) || (StatisticsDetailsItem.ChamberType == 4))
                            continue;

                        InfoTable.DefaultHorizontalAlignment = Element.ALIGN_RIGHT;
                        InfoTable.DefaultCellBackgroundColor = iTextSharp.text.Color.WHITE;
                        InfoTable.AddCell(new Phrase(StatisticsDetailsItem.CalibrationTypeDescription, GlobalClass.iReportFont12));
                    }

                    InfoTable.DefaultHorizontalAlignment = Element.ALIGN_CENTER;
                    InfoTable.AddCell(new Phrase(StatisticsDetailsItem.TotalChambers.ToString(), GlobalClass.iReportFont12));
                    InfoTable.AddCell(new Phrase(StatisticsDetailsItem.TotalCalibrationPoints.ToString(), GlobalClass.iReportFont12));
                    InfoTable.AddCell(new Phrase(StatisticsDetailsItem.TotalCalibrations.ToString(), GlobalClass.iReportFont12));
                }
                // Totals
                InfoTable.DefaultCellBackgroundColor = new iTextSharp.text.Color(255, 255, 192);
                InfoTable.DefaultHorizontalAlignment = Element.ALIGN_RIGHT;
                InfoTable.AddCell(new Phrase("Total", GlobalClass.iReportFont12));
                InfoTable.DefaultHorizontalAlignment = Element.ALIGN_CENTER;
                InfoTable.AddCell(new Phrase(Statistics.TotalIAEAChambers.ToString(), GlobalClass.iReportFont12));
                InfoTable.AddCell(new Phrase(Statistics.TotalIAEASSDL.ToString(), GlobalClass.iReportFont12));
                InfoTable.AddCell(new Phrase((Statistics.TotalIAEACalibrations + Statistics.TotalIAEAQuatroCalibrations).ToString(), GlobalClass.iReportFont12));

                pdfDocument.Add(InfoTable);
            }
            catch (Exception ex)
            {
                //this.Message = ex.Message;
                MessageBox.Show("Error " + ex.Message);
            }
            finally
            {
                pdfDocument.Close();
            }
        }

        public static void SummaryCertificateReport(string sOutputFile, DateTime dStartDate, DateTime dEndDate)
        {
            StatisticsClass Statistics = new StatisticsClass();

            Statistics.StatisticsID = 0;
            Statistics.StartDate = dStartDate;
            Statistics.EndDate = dEndDate;
            Statistics.StatisticsDescription = "Summary of certificate performed from " + GlobalClass.FormatDateTimeValue(dStartDate) + " to " + GlobalClass.FormatDateTimeValue(dEndDate);

            Statistics.PopulateSummaryCertificateTotals();
            Statistics.PopulateSummaryCertificateSSDL();

            string sLine1 = "Total of " + Statistics.TotalChambers.ToString() + " chambers calibrated from " + GlobalClass.FormatDateTimeValue(dStartDate) + " to " + GlobalClass.FormatDateTimeValue(dEndDate);
            string sLine2 = Statistics.TotalNonIAEAChambers.ToString() + " non-IAEA chambers from " + Statistics.TotalNonIAEASSDL.ToString() + " labs ";
            string sLine3 = Statistics.TotalIAEAChambers.ToString() + " IAEA chambers ( " + Statistics.TotalIAEAQuatroChambers.ToString() + " for Quatro)";

            Document pdfDocument = new Document(PageSize.A4);

            try
            {
                // step 2: creating the writer
                // Creates a Writer that listens to this document and writes the document to the Stream of your choice:
                PdfWriter pdfWriter = PdfWriter.GetInstance(pdfDocument, new FileStream(sOutputFile, FileMode.Create));

                // we Add a Header that will show up on PAGE 1
                HeaderFooter pdfHeader = new HeaderFooter(new Phrase("International Atomic Energy Agency \n The IAEA/WHO SSDL Network \n  " + Statistics.StatisticsDescription, GlobalClass.iReportFont14Blue), false);
                pdfHeader.Alignment = Element.ALIGN_CENTER;
                pdfHeader.Border = iTextSharp.text.Rectangle.BOTTOM_BORDER;
                pdfDocument.Header = pdfHeader;

                // we Add a Footer that will show up on PAGE 1
                HeaderFooter pdfFooter = new HeaderFooter(new Phrase(Statistics.StatisticsDescription + "\n Printed on " + DateTime.Now.ToString() + "\n Page ", GlobalClass.iReportFont8), true);
                pdfFooter.Alignment = Element.ALIGN_RIGHT;
                pdfFooter.Border = iTextSharp.text.Rectangle.TOP_BORDER;
                pdfDocument.Footer = pdfFooter;

                // step 3: initialisations + opening the document
                // Opens the document:
                pdfDocument.Open();

                pdfDocument.Add(new Phrase(sLine1, GlobalClass.iReportFont10)); pdfDocument.Add(new Phrase("\n", GlobalClass.iReportFont10));
                pdfDocument.Add(new Phrase(sLine2, GlobalClass.iReportFont10)); pdfDocument.Add(new Phrase("\n", GlobalClass.iReportFont10));
                pdfDocument.Add(new Phrase(sLine3, GlobalClass.iReportFont10)); pdfDocument.Add(new Phrase("\n", GlobalClass.iReportFont10));

                // Create a table and add it to the document
                Table InfoTable = new Table(3);
                InfoTable.Padding = 1;
                InfoTable.Spacing = 0;
                InfoTable.Width = 100;

                //datatable.setBorder(Rectangle.NO_BORDER);
                float[] InfoTableHeaderWidths = { 60, 20, 20 };
                InfoTable.Widths = InfoTableHeaderWidths;
                InfoTable.BorderWidth = 0;

                InfoTable.DefaultHorizontalAlignment = Element.ALIGN_CENTER;
                InfoTable.DefaultVerticalAlignment = Element.ALIGN_MIDDLE;

                InfoTable.DefaultCellBackgroundColor = iTextSharp.text.Color.WHITE;
                InfoTable.DefaultCellBorderWidth = 1;

                InfoTable.DefaultCellBackgroundColor = new iTextSharp.text.Color(255, 255, 192);
                InfoTable.AddCell(new Phrase("", GlobalClass.iReportFont12));
                InfoTable.DefaultHorizontalAlignment = Element.ALIGN_CENTER;
                InfoTable.AddCell(new Phrase("Chambers", GlobalClass.iReportFont12));
                InfoTable.AddCell(new Phrase("Calibrations", GlobalClass.iReportFont12));
                InfoTable.EndHeaders();

                foreach (StatisticsDetailsClass StatisticsDetailsItem in Statistics.StatisticsDetailsSSDL)
                {
                    InfoTable.DefaultHorizontalAlignment = Element.ALIGN_LEFT;
                    if (StatisticsDetailsItem.CalibrationType == 0)
                    {
                        InfoTable.DefaultCellBackgroundColor = new iTextSharp.text.Color(192, 255, 255);
                        InfoTable.AddCell(new Phrase(StatisticsDetailsItem.ChamberTypeDescription, GlobalClass.iReportFont12));
                    }
                    else
                    {
                        if ((StatisticsDetailsItem.ChamberType == 3) || (StatisticsDetailsItem.ChamberType == 4))
                            continue;

                        InfoTable.DefaultHorizontalAlignment = Element.ALIGN_RIGHT;
                        InfoTable.DefaultCellBackgroundColor = iTextSharp.text.Color.WHITE;
                        InfoTable.AddCell(new Phrase(StatisticsDetailsItem.CalibrationTypeDescription, GlobalClass.iReportFont12));
                    }

                    InfoTable.DefaultHorizontalAlignment = Element.ALIGN_CENTER;
                    InfoTable.AddCell(new Phrase(StatisticsDetailsItem.TotalChambers.ToString(), GlobalClass.iReportFont12));
                    InfoTable.AddCell(new Phrase(StatisticsDetailsItem.TotalCalibrations.ToString(), GlobalClass.iReportFont12));
                }
                // Totals
                InfoTable.DefaultCellBackgroundColor = new iTextSharp.text.Color(255, 255, 192);
                InfoTable.DefaultHorizontalAlignment = Element.ALIGN_RIGHT;
                InfoTable.AddCell(new Phrase("Total", GlobalClass.iReportFont12));
                InfoTable.DefaultHorizontalAlignment = Element.ALIGN_CENTER;
                InfoTable.AddCell(new Phrase(Statistics.TotalNonIAEAChambers.ToString(), GlobalClass.iReportFont12));
                InfoTable.AddCell(new Phrase(Statistics.TotalNonIAEACalibrations.ToString(), GlobalClass.iReportFont12));

                pdfDocument.Add(InfoTable);
            }
            catch (Exception ex)
            {
                //this.Message = ex.Message;
                MessageBox.Show("Error " + ex.Message);
            }
            finally
            {
                pdfDocument.Close();
            }
        }

        /*
        public static string GetJson(DataTable dt)
        {
            System.Web.Script.Serialization.JavaScriptSerializer serializer = new

            System.Web.Script.Serialization.JavaScriptSerializer();
            List<Dictionary<string, object>> rows =
              new List<Dictionary<string, object>>();
            Dictionary<string, object> row = null;

            foreach (DataRow dr in dt.Rows)
            {
                row = new Dictionary<string, object>();
                foreach (DataColumn col in dt.Columns)
                {
                    row.Add(col.ColumnName.Trim(), dr[col]);
                }
                rows.Add(row);
            }
            return serializer.Serialize(rows);
        }
        */
        public static string GenerateVisualDataHTML(object IncomingObject, string sHTMLTemplate, string sAuditType, string sParticipationType)
        {
            string sReturn = string.Empty;

            if (File.Exists(sHTMLTemplate))
            {
                if (IncomingObject != null)
                {
                    string sHTMLFileName = GlobalClass.sApplicationStartupPath + @"\Preview\TreeRadial.html";
                    string sNodes = string.Empty;
                    //string sNodes2 = string.Empty;
                    string sHTMLLines = string.Empty;

                    if (IncomingObject is YearClass)
                    {
                        YearClass SelectedBatchYear = IncomingObject as YearClass;
                        if (SelectedBatchYear != null)
                        {
                            foreach (YearClass BatchYear in GlobalClass.Dictionary.BatchYearList)
                            {
                                if (BatchYear.Year == SelectedBatchYear.Year)
                                {
                                    sNodes = sNodes + "{'name': '" + BatchYear.Year.ToString() + "'";

                                    string sNodesLevel2 = string.Empty;
                                    foreach (BatchClass Batch in BatchYear.GetBatchList(sAuditType, sParticipationType))
                                    {
                                        sNodesLevel2 = sNodesLevel2 + "{'name': '" + Batch.BatchNo + "'";

                                        string sNodesLevel3 = string.Empty; //Batch.GetVisualDataBatchPlanningChildrenNodes(sParticipationType);

                                        foreach (BatchPlanningClass BatchPlanning in Batch.BatchPlanningList)
                                        {
                                            if (BatchPlanning.ParticipationType == sParticipationType)
                                            {
                                                if (BatchPlanning.Country != null)
                                                {
                                                    if (BatchPlanning.PendingActivities != null)
                                                    {
                                                        int iTotal = BatchPlanning.PendingActivities.SetSentTotalFirstRun + BatchPlanning.PendingActivities.SetSentTotalFollowUp;
                                                        if (iTotal > 0)
                                                            sNodesLevel3 = sNodesLevel3 + "{'name': '" + BatchPlanning.Country.CountryName.Replace("'", "") + "', 'size': " + iTotal.ToString() + "}|";
                                                    }
                                                }
                                            }
                                        }

                                        if (sNodesLevel3 != string.Empty)
                                            sNodesLevel2 = sNodesLevel2 + "," + System.Environment.NewLine + "'children': [" + System.Environment.NewLine + sNodesLevel3 + "]" + System.Environment.NewLine; ;

                                        sNodesLevel2 = sNodesLevel2 + "}|";
                                    }

                                    if (sNodesLevel2 != string.Empty)
                                    {
                                        sNodesLevel2 = sNodesLevel2.TrimEnd('|').Replace("|", "," + System.Environment.NewLine);
                                        sNodes = sNodes + "," + System.Environment.NewLine + "'children': [" + System.Environment.NewLine + sNodesLevel2 + "]" + System.Environment.NewLine;
                                    }

                                    sNodes = sNodes + "}" + System.Environment.NewLine;
                                }
                            }

                            string sLine = string.Empty;
                            StreamReader sr = new StreamReader(sHTMLTemplate, System.Text.Encoding.Default, true);
                            while ((sLine = sr.ReadLine()) != null)
                            {
                                sLine = sLine.Replace("<<DataVariableName>>", "Year" + SelectedBatchYear.Year.ToString());
                                sLine = sLine.Replace("<<DataTitle>>", "Year " + SelectedBatchYear.Year.ToString());
                                sLine = sLine.Replace("<<DataJSON>>", sNodes);

                                sHTMLLines = sHTMLLines + sLine + System.Environment.NewLine;
                            }

                            sr.Close();
                        }
                    }
                    else if (IncomingObject is BatchClass)
                    {
                        BatchClass SelectedBatch = (IncomingObject as BatchClass);
                        if (SelectedBatch != null)
                        {
                            foreach (YearClass BatchYear in GlobalClass.Dictionary.BatchYearList)
                            {
                                foreach (BatchClass Batch in BatchYear.GetBatchList(sAuditType, sParticipationType))
                                {
                                    if (SelectedBatch.BatchID == Batch.BatchID)
                                    {
                                        string sSql = "EXECUTE dbo.prcQueryTLDTaskResults " + GlobalClass.sTaskInvitedDataSheets;

                                        if (sAuditType != string.Empty)
                                            sSql = sSql + ", " + sAuditType;
                                        else
                                            sSql = sSql + ", NULL";

                                        if (sParticipationType != string.Empty)
                                            sSql = sSql + ", '" + sParticipationType + "'";
                                        else
                                            sSql = sSql + ", NULL";

                                        sSql = sSql + ", 1, '" + Batch.BatchID.ToString() + "', NULL, NULL"; // Show All Certificates

                                        DataTable dataTable = GlobalClass.GetDataTable("QueryTLDTaskResults", sSql);
                                        /*
                                        if (dataTable != null)
                                        {
                                            // create the results datatable
                                            DataTable dtResults = new DataTable();
                                            dtResults.Columns.Add("BatchNo", typeof(string));
                                            dtResults.Columns.Add("Country", typeof(string));
                                            dtResults.Columns.Add("LabCode", typeof(string));
                                            dtResults.Columns.Add("SetNo", typeof(string));

                                            foreach (DataRow odr in dataTable.Rows)
                                            {
                                                dtResults.ImportRow(odr);
                                            }

                                            //sNodes2 = GlobalClass.GetJson(dtResults);
                                        }
                                        */
                                        sNodes = sNodes + "{'name': '" + Batch.BatchNo + "'";

                                        string sNodesLevel2 = string.Empty;
                                        foreach (BatchPlanningClass BatchPlanning in Batch.BatchPlanningList)
                                        {
                                            if (BatchPlanning.ParticipationType == sParticipationType)
                                            {
                                                if (BatchPlanning.Country != null)
                                                {
                                                    int iTotal = BatchPlanning.PendingActivities.SetSentTotalFirstRun + BatchPlanning.PendingActivities.SetSentTotalFollowUp;

                                                    if (sHTMLTemplate == GlobalClass.TLDGraphCollapsibleTreeLayout)
                                                        sNodesLevel2 = sNodesLevel2 + "{'name': '" + BatchPlanning.Country.CountryName.Replace("'", "") + "'";
                                                    else
                                                        sNodesLevel2 = sNodesLevel2 + "{'name': '" + BatchPlanning.Country.CCode + "'";

                                                    string sNodesLevel3 = string.Empty; 
                                                    if (dataTable != null)
                                                    {
                                                        foreach (OperatorClass Operator in BatchPlanning.Country.OperatorList)
                                                        {
                                                            DataRow[] OperatorSets = dataTable.Select("OperatorID = '" + Operator.OperatorID.ToString() + "'");
                                                            if (OperatorSets.Length > 0)
                                                            {
                                                                if (sHTMLTemplate == GlobalClass.TLDGraphCollapsibleTreeLayout)
                                                                    sNodesLevel3 = sNodesLevel3 + "{'name': '" + Operator.OperatorName.Replace("'", "") + "'";
                                                                else
                                                                    sNodesLevel3 = sNodesLevel3 + "{'name': '" + Operator.LabCode.Replace("'", "") + "'";
                                                                string sNodesLevel4 = string.Empty;

                                                                foreach (DataRow odr in OperatorSets)
                                                                {
                                                                    SetIDClass SetID = new SetIDClass();

                                                                    SetID.AuditType = sAuditType;
                                                                    SetID.AssignValue(odr, "OperatorID");
                                                                    SetID.AssignValue(odr, "SetID");
                                                                    SetID.AssignValue(odr, "SetNo");
                                                                    SetID.AssignValue(odr, "FollowUpSetNo");
                                                                    SetID.AssignValue(odr, "SetType");

                                                                    if (SetID.SetType == 1)
                                                                    {
                                                                        if (SetID.FollowUpSetNo == string.Empty)
                                                                            sNodesLevel4 = sNodesLevel4 + "{'name': '" + SetID.SetNo + "', 'size': " + SetID.SetType + "}|";
                                                                        else
                                                                            sNodesLevel4 = sNodesLevel4 + "{'name': '" + SetID.SetNo + " - " + SetID.FollowUpSetNo + "', 'size': " + SetID.SetType + "}|";
                                                                    }
                                                                }
                                                                if (sNodesLevel4 != string.Empty)
                                                                {
                                                                    sNodesLevel4 = sNodesLevel4.TrimEnd('|').Replace("|", "," + System.Environment.NewLine);
                                                                    sNodesLevel3 = sNodesLevel3 + "," + System.Environment.NewLine + "'children': [" + System.Environment.NewLine + sNodesLevel4 + "]" + System.Environment.NewLine; 
                                                                    sNodesLevel3 = sNodesLevel3 + "}|";
                                                                }
                                                            }
                                                        }
                                                    }
                                                    if (sNodesLevel3 != string.Empty)
                                                    {
                                                        sNodesLevel3 = sNodesLevel3.TrimEnd('|').Replace("|", "," + System.Environment.NewLine);
                                                        sNodesLevel2 = sNodesLevel2 + "," + System.Environment.NewLine + "'children': [" + System.Environment.NewLine + sNodesLevel3 + "]" + System.Environment.NewLine; 
                                                    }
                                                    else
                                                        sNodesLevel2 = sNodesLevel2 + ", 'size': " + iTotal.ToString();
                                                    sNodesLevel2 = sNodesLevel2 + "}|";
                                                }
                                            }
                                        }

                                        if (sNodesLevel2 != string.Empty)
                                        {
                                            sNodesLevel2 = sNodesLevel2.TrimEnd('|').Replace("|", "," + System.Environment.NewLine);
                                            sNodes = sNodes + "," + System.Environment.NewLine + "'children': [" + System.Environment.NewLine + sNodesLevel2 + "]" + System.Environment.NewLine;
                                        }

                                        sNodes = sNodes + "}" + System.Environment.NewLine;
                                    }
                                }
                            }

                            string sLine = string.Empty;
                            StreamReader sr = new StreamReader(sHTMLTemplate, System.Text.Encoding.Default, true);
                            while ((sLine = sr.ReadLine()) != null)
                            {
                                sLine = sLine.Replace("<<DataVariableName>>", "Batch" + SelectedBatch.BatchNo);
                                sLine = sLine.Replace("<<DataTitle>>", "Batch " + SelectedBatch.BatchNo);
                                sLine = sLine.Replace("<<DataJSON>>", sNodes);
                                sHTMLLines = sHTMLLines + sLine + System.Environment.NewLine;
                            }

                            sr.Close();
                        }
                    }
                    else if (IncomingObject is BatchPlanningClass)
                    {
                        BatchPlanningClass SelectedBatchPlanning = (IncomingObject as BatchPlanningClass);
                        if (SelectedBatchPlanning != null)
                        {
                            if (SelectedBatchPlanning.Country != null)
                            {
                                string sOperatorID = string.Empty;
                                foreach (OperatorClass O in SelectedBatchPlanning.Country.GetOperatorList(sParticipationType, string.Empty))
                                    sOperatorID = sOperatorID + O.OperatorID.ToString() + "|";
                                sOperatorID = sOperatorID.TrimEnd('|');

                                string sSql = "EXECUTE dbo.prcQueryTLDTaskResults " + GlobalClass.sTaskInvitedDataSheets;

                                if (sAuditType != string.Empty)
                                    sSql = sSql + ", " + sAuditType;
                                else
                                    sSql = sSql + ", NULL";

                                if (sParticipationType != string.Empty)
                                    sSql = sSql + ", '" + sParticipationType + "'";
                                else
                                    sSql = sSql + ", NULL";

                                sSql = sSql + ", 1, NULL, '" + sOperatorID + "', NULL"; // Show All Certificates

                                DataTable dataTable = GlobalClass.GetDataTable("QueryTLDTaskResults", sSql);

                                sNodes = sNodes + "{'name': '" + SelectedBatchPlanning.Country.CCode + "'";
                                string sNodesLevel1 = string.Empty;
                                foreach (YearClass BatchYear in GlobalClass.Dictionary.BatchYearList)
                                {
                                    if (BatchYear.Year >= GlobalClass.TLDActiveYear)
                                    {
                                        DataRow[] BatchYearSets = dataTable.Select("BatchYear = " + BatchYear.Year.ToString() + "");
                                        if (BatchYearSets.Length > 0)
                                        {
                                            sNodesLevel1 = sNodesLevel1 + "{'name': '" + BatchYear.Year.ToString() + "'";
                                            string sNodesLevel2 = string.Empty;
                                            foreach (BatchClass Batch in BatchYear.GetBatchList(sAuditType, sParticipationType))
                                            {
                                                DataRow[] BatchSets = dataTable.Select("BatchID = " + Batch.BatchID.ToString() + "");

                                                if (BatchSets.Length > 0)
                                                {
                                                    sNodesLevel2 = sNodesLevel2 + "{'name': '" + Batch.BatchNo + "'";
                                                    string sNodesLevel3 = string.Empty; 
                                                    foreach (BatchPlanningClass BatchPlanning in Batch.BatchPlanningList)
                                                    {
                                                        if (BatchPlanning.ParticipationType == sParticipationType)
                                                        {
                                                            if (BatchPlanning.Country != null)
                                                            {
                                                                DataRow[] BatchPlanningSets = dataTable.Select("BatchPlanningID = '" + BatchPlanning.BatchPlanningID.ToString() + "'");
                                                                
                                                                foreach (DataRow odr in BatchPlanningSets)
                                                                {
                                                                    SetIDClass SetID = new SetIDClass();

                                                                    SetID.AuditType = sAuditType;
                                                                    SetID.AssignValue(odr, "OperatorID");
                                                                    SetID.AssignValue(odr, "SetID");
                                                                    SetID.AssignValue(odr, "SetNo");
                                                                    SetID.AssignValue(odr, "FollowUpSetNo");
                                                                    SetID.AssignValue(odr, "SetType");

                                                                    if (SetID.SetType == 1)
                                                                    {
                                                                        if (SetID.FollowUpSetNo == string.Empty)
                                                                            sNodesLevel3 = sNodesLevel3 + "{'name': '" + SetID.SetNo + "', 'size': " + SetID.SetType + "}|";
                                                                        else
                                                                            sNodesLevel3 = sNodesLevel3 + "{'name': '" + SetID.SetNo + " - " + SetID.FollowUpSetNo + "', 'size': " + SetID.SetType + "}|";
                                                                    }
                                                                }
                                                            }
                                                        }
                                                    }
                                                    if (sNodesLevel3 != string.Empty)
                                                    {
                                                        sNodesLevel3 = sNodesLevel3.TrimEnd('|').Replace("|", "," + System.Environment.NewLine);
                                                        sNodesLevel2 = sNodesLevel2 + "," + System.Environment.NewLine + "'children': [" + System.Environment.NewLine + sNodesLevel3 + "]" + System.Environment.NewLine; ;
                                                    }
                                                    sNodesLevel2 = sNodesLevel2 + "}|";
                                                }
                                            }
                                            if (sNodesLevel2 != string.Empty)
                                            {
                                                sNodesLevel2 = sNodesLevel2.TrimEnd('|').Replace("|", "," + System.Environment.NewLine);
                                                sNodesLevel1 = sNodesLevel1 + "," + System.Environment.NewLine + "'children': [" + System.Environment.NewLine + sNodesLevel2 + "]" + System.Environment.NewLine; ;
                                            }
                                            sNodesLevel1 = sNodesLevel1 + "}|";
                                        }
                                    }
                                }

                                if (sNodesLevel1 != string.Empty)
                                {
                                    sNodesLevel1 = sNodesLevel1.TrimEnd('|').Replace("|", "," + System.Environment.NewLine);
                                    sNodes = sNodes + "," + System.Environment.NewLine + "'children': [" + System.Environment.NewLine + sNodesLevel1 + "]" + System.Environment.NewLine;
                                }

                                sNodes = sNodes + "}" + System.Environment.NewLine;

                                string sLine = string.Empty;
                                StreamReader sr = new StreamReader(sHTMLTemplate, System.Text.Encoding.Default, true);
                                while ((sLine = sr.ReadLine()) != null)
                                {
                                    sLine = sLine.Replace("<<DataVariableName>>", "Country" + SelectedBatchPlanning.Country.CCode);
                                    sLine = sLine.Replace("<<DataTitle>>", "" + SelectedBatchPlanning.Country.Country);
                                    sLine = sLine.Replace("<<DataJSON>>", sNodes);
                                    sHTMLLines = sHTMLLines + sLine + System.Environment.NewLine;
                                }

                                sr.Close();
                            }
                        }
                    }
                    else if (IncomingObject is OperatorClass)
                    {
                        OperatorClass SelectedOperator = (IncomingObject as OperatorClass);
                        if (SelectedOperator != null)
                        {
                            string sSql = "EXECUTE dbo.prcQueryTLDTaskResults " + GlobalClass.sTaskInvitedDataSheets;

                            if (sAuditType != string.Empty)
                                sSql = sSql + ", " + sAuditType;
                            else
                                sSql = sSql + ", NULL";

                            if (sParticipationType != string.Empty)
                                sSql = sSql + ", '" + sParticipationType + "'";
                            else
                                sSql = sSql + ", NULL";

                            sSql = sSql + ", 1, NULL, '" + SelectedOperator.OperatorID.ToString() + "', NULL"; // Show All Certificates

                            DataTable dataTable = GlobalClass.GetDataTable("QueryTLDTaskResults", sSql);
                            if (dataTable != null)
                            {
                                if (sHTMLTemplate == GlobalClass.TLDGraphCollapsibleTreeLayout)
                                    sNodes = sNodes + "{'name': '" + SelectedOperator.OperatorName.Replace("'", "") + "'";
                                else
                                    sNodes = sNodes + "{'name': '" + SelectedOperator.LabCode + "'";
                                string sNodesLevel1 = string.Empty;

                                foreach (UnitClass Unit in SelectedOperator.UnitList)
                                {
                                    if (sHTMLTemplate == GlobalClass.TLDGraphCollapsibleTreeLayout)
                                        sNodesLevel1 = sNodesLevel1 + "{'name': '" + Unit.Title + "'";
                                    else
                                        sNodesLevel1 = sNodesLevel1 + "{'name': '" + Unit.UnitCode + "'";
                                    string sNodesLevel2 = string.Empty;

                                    foreach (BeamClass Beam in Unit.BeamList)
                                    {
                                        if (sHTMLTemplate == GlobalClass.TLDGraphCollapsibleTreeLayout)
                                            sNodesLevel2 = sNodesLevel2 + "{'name': '" + Beam.Energy + "'";
                                        else
                                            sNodesLevel2 = sNodesLevel2 + "{'name': '" + Beam.BeamCode + "'";
                                        string sNodesLevel3 = string.Empty;

                                        foreach (YearClass BatchYear in GlobalClass.Dictionary.BatchYearList)
                                        {
                                            if (BatchYear.Year >= GlobalClass.TLDActiveYear)
                                            {
                                                foreach (BatchClass Batch in BatchYear.GetBatchList(sAuditType, sParticipationType))
                                                {
                                                    DataRow[] BatchSets = dataTable.Select("BatchID = " + Batch.BatchID.ToString() + " AND BeamID = " + Beam.BeamID.ToString());

                                                    if (BatchSets.Length > 0)
                                                    {
                                                        sNodesLevel3 = sNodesLevel3 + "{'name': '" + Batch.BatchYear.ToString() + " " + Batch.BatchNo + "'";
                                                        string sNodesLevel4 = string.Empty;

                                                        foreach (DataRow odr in BatchSets)
                                                        {
                                                            SetIDClass SetID = new SetIDClass();

                                                            SetID.AuditType = sAuditType;
                                                            SetID.AssignValue(odr, "OperatorID");
                                                            SetID.AssignValue(odr, "UnitID");
                                                            SetID.AssignValue(odr, "BeamID");
                                                            SetID.AssignValue(odr, "SetID");
                                                            SetID.AssignValue(odr, "SetNo");
                                                            SetID.AssignValue(odr, "FollowUpSetNo");
                                                            SetID.AssignValue(odr, "SetType");

                                                            if (SetID.BeamID == Beam.BeamID)
                                                            {
                                                                if (SetID.SetType == 1)
                                                                {
                                                                    if (SetID.FollowUpSetNo == string.Empty)
                                                                        sNodesLevel4 = sNodesLevel4 + "{'name': '" + SetID.SetNo + "', 'size': " + SetID.SetType + "}|";
                                                                    else
                                                                        sNodesLevel4 = sNodesLevel4 + "{'name': '" + SetID.SetNo + " - " + SetID.FollowUpSetNo + "', 'size': " + SetID.SetType + "}|";
                                                                }
                                                            }
                                                        }
                                                        if (sNodesLevel4 != string.Empty)
                                                        {
                                                            sNodesLevel4 = sNodesLevel4.TrimEnd('|').Replace("|", "," + System.Environment.NewLine);
                                                            sNodesLevel3 = sNodesLevel3 + "," + System.Environment.NewLine + "'children': [" + System.Environment.NewLine + sNodesLevel4 + "]" + System.Environment.NewLine;
                                                            sNodesLevel3 = sNodesLevel3 + "}|";
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                        if (sNodesLevel3 != string.Empty)
                                        {
                                            sNodesLevel3 = sNodesLevel3.TrimEnd('|').Replace("|", "," + System.Environment.NewLine);
                                            sNodesLevel2 = sNodesLevel2 + "," + System.Environment.NewLine + "'children': [" + System.Environment.NewLine + sNodesLevel3 + "]" + System.Environment.NewLine;
                                        }
                                        sNodesLevel2 = sNodesLevel2 + "}|";
                                    }
                                    if (sNodesLevel2 != string.Empty)
                                    {
                                        sNodesLevel2 = sNodesLevel2.TrimEnd('|').Replace("|", "," + System.Environment.NewLine);
                                        sNodesLevel1 = sNodesLevel1 + "," + System.Environment.NewLine + "'children': [" + System.Environment.NewLine + sNodesLevel2 + "]" + System.Environment.NewLine; ;
                                    }
                                    sNodesLevel1 = sNodesLevel1 + "}|";
                                }

                                if (sNodesLevel1 != string.Empty)
                                {
                                    sNodesLevel1 = sNodesLevel1.TrimEnd('|').Replace("|", "," + System.Environment.NewLine);
                                    sNodes = sNodes + "," + System.Environment.NewLine + "'children': [" + System.Environment.NewLine + sNodesLevel1 + "]" + System.Environment.NewLine;
                                }

                                sNodes = sNodes + "}" + System.Environment.NewLine;

                                string sLine = string.Empty;
                                StreamReader sr = new StreamReader(sHTMLTemplate, System.Text.Encoding.Default, true);
                                while ((sLine = sr.ReadLine()) != null)
                                {
                                    sLine = sLine.Replace("<<DataVariableName>>", "Country" + SelectedOperator.CCode);
                                    sLine = sLine.Replace("<<DataTitle>>", "" + SelectedOperator.Country);
                                    sLine = sLine.Replace("<<DataJSON>>", sNodes);
                                    sHTMLLines = sHTMLLines + sLine + System.Environment.NewLine;
                                }

                                sr.Close();
                            }
                        }
                    }

                    if (File.Exists(sHTMLFileName))
                        File.Delete(sHTMLFileName);
                    StreamWriter swHTML = new StreamWriter(sHTMLFileName, true);
                    swHTML.WriteLine(sHTMLLines);
                    swHTML.Close();

                    if (File.Exists(sHTMLFileName))
                        sReturn = sHTMLFileName;
                    //System.Diagnostics.Process.Start(sHTMLFile);
                }

            }

            return sReturn;
        }
    }
}
