using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using System.Data;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using System.Runtime.InteropServices;
using System.Reflection;
using System.IO;
using iTextSharp.text;
using iTextSharp.text.pdf;


namespace SSDLAdmin
{
    public class TLDPackageClass : SSDLBaseAttachmentClass
    {
        private int _PackageID = -1;
        
        private string _AuditType = string.Empty;
        private string _ParticipationType = string.Empty;
        private string _ParticipationCategory = "RegularParticipation"; //RegularParticipation, SpecialRequest, CRP, QUATRO


        private int _BatchPlanningID = -1;
        private int _BatchID = -1;
        private int _PackageType = -1; // 1-FirstIrradiation | 2-FollowUp 

        private int _ContactID = -1; // Main Contact ID to sent certificates

        private string _CoordinationNetwork = "Off"; // From TLDPackage.Country // WHO | PAHO // Only for countries with National Coordinator
        private string _SendToDataSheetsOption = "Off"; // From TLDPackage.Country
        private string _ReturnDataSheetsOption = "Off"; // From TLDPackage.Country
        private string _CommunicationLanguage = string.Empty; // From TLDPackage.Country // "English", "Spanish", "Russian";

        private string _PackageDescription = string.Empty;
        private string _PackageComment = string.Empty;
        private string _PackageStatus = "Off";


        //private bool _Archived = false;

        //private DateTime _SignatureDate = DateTime.MinValue;
        //private string _SignatureByUser = string.Empty;

        private DateTime _DispatchedOn = DateTime.MinValue;
        private string _DispatchedBy = string.Empty;

        private DateTime _ArchivedOn = DateTime.MinValue;
        private string _ArchivedBy = string.Empty;

        private DateTime _CreatedOn = DateTime.MinValue;
        private string _CreatedBy = string.Empty;

        private string _UpdateComment = string.Empty;
        private DateTime _LastUpdate;

        public TLDApplicationFormClass TLDApplicationForm = null;
        //public List<TLDCertificateClass> TLDCertificateList = new List<TLDCertificateClass>();

        public TLDPackageClass(SSDLBaseClass ParentObject)
            : base(ParentObject)           // Call base-class constructor first
        {
            if (ParentObject is OperatorClass)
            {
                this.OperatorID = (ParentObject as OperatorClass).OperatorID;
                this.CCode = (ParentObject as OperatorClass).CCode;
                this.LabCode = (ParentObject as OperatorClass).LabCode;
            }
        }

        public int PackageID
        {
            get { return _PackageID; }
            set
            {
                if (this._PackageID != value)
                {
                    this._PackageID = value;
                    _StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public override int DocumentID
        {
            get { return _PackageID; }
        }

        public string AuditType
        {
            get { return this._AuditType; }
            set
            {
                if (this._AuditType.Trim() != value.Trim())
                {
                    this._AuditType = value.Trim();
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public string ParticipationType
        {
            get { return this._ParticipationType; }
            set
            {
                if (this._ParticipationType.Trim() != value.Trim())
                {
                    this._ParticipationType = value.Trim();
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public string ParticipationCategory
        {
            get
            {
                string sReturn = this._ParticipationCategory;

                if (sReturn == string.Empty)
                    sReturn = "Off";

                return sReturn;
            }
            set
            {
                if (value.Trim() == string.Empty)
                    value = "Off";

                if (this._ParticipationCategory.Trim() != value.Trim())
                {
                    this._ParticipationCategory = value.Trim();
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public int BatchPlanningID
        {
            get { return _BatchPlanningID; }
            set
            {
                if (this._BatchPlanningID != value)
                {
                    this._BatchPlanningID = value;
                    _StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public int BatchID
        {
            set
            {
                if (this._BatchID != value)
                {
                    this._BatchID = value;
                    _StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public BatchClass Batch
        {
            get
            {
                return GlobalClass.Manager.GetBatch(this._BatchID);
            }
        }

        public BatchPlanningClass BatchPlanning
        {
            get
            {
                BatchPlanningClass BatchPlanning = null;

                if (Batch != null)
                    BatchPlanning = Batch.GetBatchPlanning(this._BatchPlanningID);

                return BatchPlanning;
            }
        }

        public CountryClass Country
        {
            get
            {
                return GlobalClass.Manager.GetCountry(this._CCode);
            }
        }

        public ContactClass Contact
        {
            get
            {
                ContactClass ReturnContact = null;

                if (this.Operator != null)
                    ReturnContact = this.Operator.GetContact(this.ContactID);                    

                return ReturnContact;
            }
        }

        public int PackageType
        {
            get { return _PackageType; }
            set
            {
                if (this._PackageType != value)
                {
                    this._PackageType = value;
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public int ContactID
        {
            get 
            {
                int iContactID = this._ContactID;

                if (iContactID == -1)
                {
                    if (this.TLDApplicationForm != null)
                        if (this.TLDApplicationForm.MedicalPhysicist != null)
                            iContactID = this.TLDApplicationForm.MedicalPhysicist.ContactID;
                }

                return iContactID; 
            }
            set
            {
                if (this._ContactID != value)
                {
                    this._ContactID = value;
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public string CoordinationNetwork
        {
            get
            {
                string sReturn = this._CoordinationNetwork;

                if (sReturn == string.Empty)
                    sReturn = "Off";

                return sReturn;
            }
            set
            {
                if (value.Trim() == string.Empty)
                    value = "Off";

                if (this._CoordinationNetwork.Trim() != value.Trim())
                {
                    this._CoordinationNetwork = value.Trim();
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public string SendToDataSheetsOption
        {
            get
            {
                string sReturn = this._SendToDataSheetsOption;

                if (sReturn == string.Empty)
                    sReturn = "Off";

                return sReturn;
            }
            set
            {
                if (value.Trim() == string.Empty)
                    value = "Off";

                if (this._SendToDataSheetsOption.Trim() != value.Trim())
                {
                    this._SendToDataSheetsOption = value.Trim();
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public string ReturnDataSheetsOption
        {
            get
            {
                string sReturn = this._ReturnDataSheetsOption;

                if (sReturn == string.Empty)
                    sReturn = "Off";

                return sReturn;
            }
            set
            {
                if (value.Trim() == string.Empty)
                    value = "Off";

                if (this._ReturnDataSheetsOption.Trim() != value.Trim())
                {
                    this._ReturnDataSheetsOption = value.Trim();
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public string CommunicationLanguage
        {
            get
            {
                string sReturn = this._CommunicationLanguage;

                if (sReturn == "Off")
                    sReturn = "English";
                else if (sReturn == string.Empty)
                    sReturn = "English";

                return sReturn;
            }
            set
            {
                if (this._CommunicationLanguage.Trim() != value.Trim())
                {
                    this._CommunicationLanguage = value.Trim();
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public string PackageDescription
        {
            get { return _PackageDescription; }
            set
            {
                if (this._PackageDescription.Trim() != value.Trim())
                {
                    this._PackageDescription = value.Trim();
                    _StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public string PackageComment
        {
            get { return _PackageComment; }
            set
            {
                if (this._PackageComment.Trim() != value.Trim())
                {
                    this._PackageComment = value.Trim();
                    _StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public string PackageStatus
        {
            get
            {
                string sReturn = this._PackageStatus;

                if (sReturn == string.Empty)
                    sReturn = "Off";

                return sReturn;
            }
            set
            {
                if (value.Trim() == string.Empty)
                    value = "Off";

                if (this._PackageStatus.Trim() != value.Trim())
                {
                    this._PackageStatus = value.Trim();
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

       /*
        public bool Archived
        {
            get { return this._Archived; }
            set
            {
                if (this._Archived != value)
                {
                    this._Archived = value;
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }
        
        public DateTime SignatureDate
        {
            get { return _SignatureDate; }
            set
            {
                if (this._SignatureDate != value)
                {
                    this._SignatureDate = value;
                    _StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public string SignatureByUser
        {
            get { return _SignatureByUser; }
            set
            {
                if (this._SignatureByUser.Trim() != value.Trim())
                {
                    this._SignatureByUser = value.Trim();
                    _StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }
        */

        public DateTime DispatchedOn
        {
            get { return this._DispatchedOn; }
            set
            {
                if (this._DispatchedOn != value)
                {
                    this._DispatchedOn = value;
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public string DispatchedBy
        {
            get { return this._DispatchedBy; }
            set
            {
                if (this._DispatchedBy.Trim() != value.Trim())
                {
                    this._DispatchedBy = value.Trim();
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public DateTime ArchivedOn
        {
            get { return this._ArchivedOn; }
            set
            {
                if (this._ArchivedOn != value)
                {
                    this._ArchivedOn = value;
                    this._StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public string ArchivedBy
        {
            get { return this._ArchivedBy; }
            set
            {
                if (this._ArchivedBy.Trim() != value.Trim())
                {
                    this._ArchivedBy = value.Trim();
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

        public DateTime LastUpdate
        {
            get { return _LastUpdate; }
            set
            {
                if (this._LastUpdate != value)
                {
                    this._LastUpdate = value;
                    _StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public string UpdateComment
        {
            get { return _UpdateComment; }
            set
            {
                if (this._UpdateComment.Trim() != value.Trim())
                {
                    this._UpdateComment = value.Trim();
                    _StateStatus = GlobalClass.sStateStatusDirty;
                }
            }
        }

        public List<TLDSetClass> GetTLDSetList()
        {
            List<TLDSetClass> ReturnTLDSetList = new List<TLDSetClass>();

            if (this.Operator != null)
            {
                foreach (TLDSetClass TLDSet in this.Operator.TLDSetList)
                {
                    if (TLDSet.PackageID == this.PackageID)
                    {
                        ReturnTLDSetList.Add(TLDSet);
                    }
                }
            }

            ReturnTLDSetList.Sort(delegate(TLDSetClass p1, TLDSetClass p2) { return p1.SetNo.CompareTo(p2.SetNo); });
            return ReturnTLDSetList;
        }

        public List<ContactClass> GetTLDContactList()
        {
            List<ContactClass> ReturnTLDContactList = new List<ContactClass>();
            string sContactIDs = string.Empty;

            if (this.Operator != null)
            {
                foreach (TLDSetClass TLDSet in this.Operator.TLDSetList)
                {
                    if (TLDSet.PackageID == this.PackageID)
                    {
                        if (TLDSet.Contact != null)
                        {
                            if (sContactIDs.Contains(TLDSet.Contact.ContactID.ToString() + ",") == false)
                            {
                                ReturnTLDContactList.Add(TLDSet.Contact);
                                sContactIDs = sContactIDs + TLDSet.Contact.ContactID.ToString() + ",";
                            }
                        }
                    }
                }
            }

            return ReturnTLDContactList;
        }

        public string SetNos
        {
            get 
            {
                string sReturn = string.Empty;

                foreach (TLDSetClass TLDSet in this.GetTLDSetList())
                    sReturn = sReturn + TLDSet.SetNo + ",";

                sReturn = sReturn.TrimEnd(',');
                int iLastIndex = sReturn.LastIndexOf(',');
                if (iLastIndex > -1)
                    sReturn = sReturn.Substring(0, iLastIndex) + " and " + sReturn.Substring(iLastIndex + 1);

                sReturn = sReturn.Replace(",", ", ");

                return sReturn; 
            }
        }

        public string OriginalSetNos
        {
            get
            {
                string sReturn = string.Empty;

                if (this.PackageType == 2)
                {
                    foreach (TLDSetClass TLDSet in this.GetTLDSetList())
                    {
                        string sOriginalSetNo = TLDSet.OriginalSetNo;
                        if (sOriginalSetNo != string.Empty)
                            sReturn = sReturn + sOriginalSetNo + ",";
                    }

                    sReturn = sReturn.TrimEnd(',');
                    int iLastIndex = sReturn.LastIndexOf(',');
                    if (iLastIndex > -1)
                        sReturn = sReturn.Substring(0, iLastIndex) + " and " + sReturn.Substring(iLastIndex + 1);

                    sReturn = sReturn.Replace(",", ", ");
                }

                return sReturn;
            }
        }

        public string FollowUpSetNos
        {
            get
            {
                string sReturn = string.Empty;

                if (this.PackageType == 1)
                {
                    foreach (TLDSetClass TLDSet in this.GetTLDSetList())
                    {
                        string sFollowUpSetNo = TLDSet.FollowUpSetNo;
                        if (sFollowUpSetNo != string.Empty)
                            sReturn = sReturn + TLDSet.FollowUpSetNo + ",";
                    }

                    sReturn = sReturn.TrimEnd(',');
                    int iLastIndex = sReturn.LastIndexOf(',');
                    if (iLastIndex > -1)
                        sReturn = sReturn.Substring(0, iLastIndex) + " and " + sReturn.Substring(iLastIndex + 1);

                    sReturn = sReturn.Replace(",", ", ");
                }

                return sReturn;
            }
        }

        public void PopulateRecord(DataRow SetDr, DataTable TLDAttachmentsTable)
        {
            if (SetDr != null)
            {
                if (SetDr["PackageID"] != DBNull.Value) { this.PackageID = (int)SetDr["PackageID"]; };
                if (SetDr["BatchID"] != DBNull.Value) { this.BatchID = (int)SetDr["BatchID"]; };
                if (SetDr["BatchPlanningID"] != DBNull.Value) { this.BatchPlanningID = (int)SetDr["BatchPlanningID"]; };
                if (SetDr["ContactID"] != DBNull.Value) { this.ContactID = (int)SetDr["ContactID"]; };

                if (SetDr["PackageType"] != DBNull.Value) { this.PackageType = (int)SetDr["PackageType"]; };
                if (SetDr["AuditType"] != DBNull.Value) { this.AuditType = (string)SetDr["AuditType"]; };
                if (SetDr["ParticipationType"] != DBNull.Value) { this.ParticipationType = (string)SetDr["ParticipationType"]; };
                if (SetDr["ParticipationCategory"] != DBNull.Value) { this.ParticipationCategory = (string)SetDr["ParticipationCategory"]; };

                if (SetDr["CoordinationNetwork"] != DBNull.Value) { this.CoordinationNetwork = (string)SetDr["CoordinationNetwork"]; }
                if (SetDr["CommunicationLanguage"] != DBNull.Value) { this.CommunicationLanguage = (string)SetDr["CommunicationLanguage"]; };
                if (SetDr["SendToDataSheetsOption"] != DBNull.Value) { this.SendToDataSheetsOption = (string)SetDr["SendToDataSheetsOption"]; };
                if (SetDr["ReturnDataSheetsOption"] != DBNull.Value) { this.ReturnDataSheetsOption = (string)SetDr["ReturnDataSheetsOption"]; };

                if (SetDr["PackageDescription"] != DBNull.Value) { this.PackageDescription = (string)SetDr["PackageDescription"]; };
                if (SetDr["PackageComment"] != DBNull.Value) { this.PackageComment = (string)SetDr["PackageComment"]; };
                if (SetDr["PackageStatus"] != DBNull.Value) { this.PackageStatus = (string)SetDr["PackageStatus"]; };
                
                //if (SetDr["Archived"] != DBNull.Value) { this.Archived = (bool)SetDr["Archived"]; };
                if (SetDr["DispatchedOn"] != DBNull.Value) { this.DispatchedOn = (DateTime)SetDr["DispatchedOn"]; };
                if (SetDr["DispatchedBy"] != DBNull.Value) { this.DispatchedBy = (string)SetDr["DispatchedBy"]; };

                if (SetDr["ArchivedOn"] != DBNull.Value) { this.ArchivedOn = (DateTime)SetDr["ArchivedOn"]; };
                if (SetDr["ArchivedBy"] != DBNull.Value) { this.ArchivedBy = (string)SetDr["ArchivedBy"]; };

                if (SetDr["LastUpdate"] != DBNull.Value) { this.LastUpdate = (DateTime)SetDr["LastUpdate"]; };
                if (SetDr["UpdateComment"] != DBNull.Value) { this.UpdateComment = (string)SetDr["UpdateComment"]; };

                if (SetDr["CreatedOn"] != DBNull.Value) { this.CreatedOn = (DateTime)SetDr["CreatedOn"]; };
                if (SetDr["CreatedBy"] != DBNull.Value) { this.CreatedBy = (string)SetDr["CreatedBy"]; };

                this.AttachmentList.Clear();
                if (TLDAttachmentsTable != null)
                {
                    foreach (DataRow AttDr in TLDAttachmentsTable.Rows)
                    {
                        if ((int)AttDr["AttachmentType"] == GlobalClass.iAttachmentApplicationForm)                            
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

        public void PopulateTLDApplicationForm()
        {
            string sSql = "SELECT * FROM dbo.TLDApplicationForm WHERE PackageID = " + this._PackageID.ToString();
            DataTable TLDApplicationFormDataTable = GlobalClass.GetDataTable("TLDApplicationForm", sSql);
            this.TLDApplicationForm = null;

            if (TLDApplicationFormDataTable != null)
            {
                if (TLDApplicationFormDataTable.Rows.Count == 1)
                {
                    foreach (DataRow dr in TLDApplicationFormDataTable.Rows)
                    {
                        this.TLDApplicationForm = new TLDApplicationFormClass(this);

                        Type TLDApplicationFormType = this.TLDApplicationForm.GetType();
                        PropertyInfo[] TLDApplicationFormProperties = TLDApplicationFormType.GetProperties();
                        foreach (PropertyInfo Property in TLDApplicationFormProperties)
                        {
                            if (Property != null)
                            {
                                string sFieldName = Property.Name.Trim();
                                string sFieldType = Property.PropertyType.Name;

                                if (GlobalClass.ColumnExists(TLDApplicationFormDataTable, sFieldName) == true)
                                {
                                    if (dr[sFieldName] != DBNull.Value)
                                    {
                                        try
                                        {
                                            if (Property.CanWrite == true)
                                                Property.SetValue(this.TLDApplicationForm, dr[sFieldName], null);
                                        }
                                        catch
                                        {
                                            MessageBox.Show("Incorrect value in the field TLDPackage.TLDApplicationForm." + sFieldName, "Value Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                        }
                                    }
                                }
                            }
                        }
                    }

                    this.TLDApplicationForm.StateStatus = GlobalClass.sStateStatusClean;
                }
            }
        }

        public int DeletePackage()
        {
            int iReturn = 0;
            string sSql = string.Empty;
            int iRowCount = 0;

            // Check TLDSet Table
            sSql = "SELECT PackageID FROM dbo.TLDSet WHERE PackageID = " + this._PackageID.ToString();
            DataTable dataTable = GlobalClass.GetDataTable("CheckRecord", sSql);
            if (dataTable != null)
                iRowCount = iRowCount + dataTable.Rows.Count;

            if (this._PackageID > 0)
            {
                if (iRowCount == 0)
                {
                    sSql = "DELETE FROM dbo.TLDApplicationForm WHERE PackageID = " + this._PackageID.ToString() + ";";
                    sSql = sSql + "DELETE FROM dbo.TLDPackages WHERE PackageID = " + this._PackageID.ToString();

                    try
                    {
                        iReturn = GlobalClass.ExecuteSQL(sSql);
                    }
                    catch
                    {
                        iReturn = 0;

                        MessageBox.Show("Incorrect SQL statement.", "SQL Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                else
                {
                    MessageBox.Show("Can not delete selected Package [" + this._PackageID.ToString() + "]", "Check integraty Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

            }

            return iReturn;
        }

        public string CheckBeforeSave()
        {
            string sReturn = string.Empty;

            if (GlobalClass.User != null)
            {
                if ((this._PackageID == -1) && (GlobalClass.User.isUserPermissionValid("TLDPackageCreate") == false))
                    sReturn = sReturn + "User [" + GlobalClass.User.UserName + "] does not have permissions to create TLD Package." + System.Environment.NewLine;
                if ((this._PackageID != -1) && (GlobalClass.User.isUserPermissionValid("TLDPackageEdit") == false))
                    sReturn = sReturn + "User [" + GlobalClass.User.UserName + "] does not have permissions to edit TLD Package." + System.Environment.NewLine;
            }
            else
                sReturn = sReturn + "Unknown user does not have permissions to perform this operation." + System.Environment.NewLine;

            if (this.OperatorID == -1)
            {
                if (this._ParticipationType == GlobalClass.sParticipationTypeHospitals)
                    sReturn = sReturn + "Please select Hospital" + System.Environment.NewLine;
                else
                    sReturn = sReturn + "Please select " + this._ParticipationType + " Laboratory" + System.Environment.NewLine;
            }

            if (this.Operator == null)
                sReturn = sReturn + "Unknown Operator" + System.Environment.NewLine;

            if (this._AuditType == string.Empty)
                sReturn = sReturn + "Please select Audit Type" + System.Environment.NewLine;

            if (this._BatchPlanningID == -1)
                sReturn = sReturn + "Please select Batch Planning item" + System.Environment.NewLine;


            if (this._ContactID == -1)
                if (this.TLDApplicationForm != null)
                    if (this.TLDApplicationForm.MedicalPhysicist != null)
                        this._ContactID = this.TLDApplicationForm.MedicalPhysicist.ContactID;

            // DictionaryType --------------------------------------------------------------------------
            DictionaryTypeClass DictionaryType = GlobalClass.Dictionary.GetDictionaryType(GlobalClass.Dictionary.GetDictionary("ParticipationType"), this._ParticipationType);
            if (DictionaryType == null)
                sReturn = sReturn + "Unknown Participation Type = [" + this._ParticipationType + "]" + System.Environment.NewLine;

            DictionaryType = GlobalClass.Dictionary.GetDictionaryType(GlobalClass.Dictionary.GetDictionary("SetType"), this._PackageType);
            if (DictionaryType == null)
                sReturn = sReturn + "Unknown Set Type = [" + this._PackageType.ToString() + "]" + System.Environment.NewLine;

            // Batch  --------------------------------------------------------------------------
            if (this._BatchID == -1)
                sReturn = sReturn + "Please select Batch" + System.Environment.NewLine;
            else
            {
                //BatchClass Batch = GlobalClass.Dictionary.GetBatch(this._BatchID);
                if (this.Batch == null)
                    sReturn = sReturn + "Unknown Batch ID [" + this._BatchID.ToString() + "]" + System.Environment.NewLine;
            }

            return sReturn;
        }

        public int SavePackage()
        {
            int iReturn = 0;

            this._LastUpdate = DateTime.Now;
            this._UpdateComment = "Updated by user: [" + GlobalClass.User.UserName + "] IP Address: [" + GlobalClass.GetIPAddress() + "] Date: [" + this._LastUpdate.ToString() + "]";

            string sSql = string.Empty;
            string sFields = string.Empty;
            string sValues = string.Empty;

            int iPackageID = -1;

            // Get PaperDocumentID
            sSql = "SELECT PackageID FROM dbo.TLDPackages WHERE PackageID = " + this._PackageID.ToString();
            DataTable dataTable = GlobalClass.GetDataTable("TLDPackages", sSql);
            if (dataTable != null)
            {
                if (dataTable.Rows.Count == 1)
                {
                    if (dataTable.Rows[0]["PackageID"] != DBNull.Value)
                    {
                        iPackageID = (int)dataTable.Rows[0]["PackageID"];
                    }
                }
            }
            dataTable = null;

            try
            {
                this._LastUpdate = DateTime.Now;

                if (iPackageID == -1)
                {
                    List<ParameterClass> ParameterList = new List<ParameterClass>();
                    this._UpdateComment = "Inserted by user: [" + GlobalClass.User.UserName + "] IP Address: [" + GlobalClass.GetIPAddress() + "] Date: [" + this._LastUpdate.ToString() + "]";

                    // Insert Record                        
                    sFields = "PackageType, AuditType, BatchID, BatchPlanningID, OperatorID, ContactID, ParticipationType, ParticipationCategory, CCode, SendToDataSheetsOption, ReturnDataSheetsOption, CommunicationLanguage, CoordinationNetwork, PackageDescription, PackageComment, PackageStatus, DispatchedOn, DispatchedBy, ArchivedOn, ArchivedBy, CreatedBy, CreatedOn, UpdateComment, LastUpdate";
                    sValues = sValues + "" + this.PackageType.ToString().Trim() + ", ";
                    sValues = sValues + "'" + this.AuditType.ToString().Trim() + "', ";
                    sValues = sValues + "" + GlobalClass.FormatIntegerValue(this._BatchID) + ", ";
                    sValues = sValues + "" + GlobalClass.FormatIntegerValue(this.BatchPlanningID) + ", ";
                    sValues = sValues + "" + GlobalClass.FormatIntegerValue(this.OperatorID) + ", ";
                    sValues = sValues + "" + GlobalClass.FormatIntegerValue(this.ContactID) + ", ";
                    //sValues = sValues + "" + this.ContactID.ToString() + ", ";
                    sValues = sValues + "'" + GlobalClass.FormatStringValue(this.ParticipationType, 50) + "', ";
                    sValues = sValues + "'" + GlobalClass.FormatStringValue(this.ParticipationCategory, 50) + "', ";
                    sValues = sValues + "'" + this.CCode.Trim() + "', ";
                    sValues = sValues + "'" + GlobalClass.FormatStringValue(this.SendToDataSheetsOption.Trim(), 50) + "', ";
                    sValues = sValues + "'" + GlobalClass.FormatStringValue(this.ReturnDataSheetsOption.Trim(), 50) + "', ";
                    sValues = sValues + "'" + GlobalClass.FormatStringValue(this.CommunicationLanguage.Trim(), 50) + "', ";
                    sValues = sValues + "'" + GlobalClass.FormatStringValue(this.CoordinationNetwork.Trim(), 50) + "', ";

                    sValues = sValues + "'" + GlobalClass.FormatStringValue(this.PackageDescription.Trim(), 2000) + "', ";
                    sValues = sValues + "'" + GlobalClass.FormatStringValue(this.PackageComment.Trim(), 2000) + "', ";
                    sValues = sValues + "'" + GlobalClass.FormatStringValue(this.PackageStatus.Trim(), 100) + "', ";
                    
                    sValues = sValues + "'" + this.DispatchedOn + "',";
                    sValues = sValues + "'" + GlobalClass.FormatStringValue(this.DispatchedBy.Trim(), 250) + "', ";

                    sValues = sValues + "'" + this.ArchivedOn + "',";
                    sValues = sValues + "'" + GlobalClass.FormatStringValue(this.ArchivedBy.Trim(), 250) + "', ";


                    sValues = sValues + "'" + GlobalClass.FormatStringValue(this.CreatedBy.Trim(), 250) + "', ";
                    sValues = sValues + "'" + this.CreatedOn + "',";
                    sValues = sValues + "'" + GlobalClass.FormatStringValue(this.UpdateComment.Trim(), 250) + "', ";
                    sValues = sValues + "'" + this.LastUpdate + "'";

                    sValues = sValues.Replace("'0001-01-01 00:00:00'", "NULL");
                    sSql = "INSERT INTO dbo.TLDPackages (" + sFields + ") VALUES (" + sValues + ")";
                    if (ParameterList.Count == 0)
                        iReturn = GlobalClass.ExecuteSQL(sSql);
                    else
                        iReturn = GlobalClass.ExecuteSQL(sSql, ParameterList);

                    GlobalClass.LogUserAction(2, this.OperatorID, "User [#UserName#] Execute INSERT on dbo.TLDPackages table", sSql);

                    // Get Max PackageID
                    sSql = "SELECT MAX(PackageID) as PackageID FROM dbo.TLDPackages";
                    dataTable = GlobalClass.GetDataTable("MaxPackageID", sSql);
                    if (dataTable != null)
                    {
                        if (dataTable.Rows.Count == 1)
                        {
                            if (dataTable.Rows[0]["PackageID"] != DBNull.Value)
                            {
                                this._PackageID = (int)dataTable.Rows[0]["PackageID"];
                            }
                        }
                    }
                    dataTable = null;

                    if (this._PackageID != -1)
                    {
                        if (TLDApplicationForm != null)
                        {
                            TLDApplicationForm.PackageID = this._PackageID;
                            TLDApplicationForm.SaveApplicationForm();
                        }

                        foreach (AttachmentClass Attachment in this.AttachmentList)
                        {
                            if (Attachment.DocumentID != this._PackageID)
                                Attachment.DocumentID = this._PackageID;
                        }
                        this.SaveAttachments();
                    }
                }
                else
                {
                    List<ParameterClass> ParameterList = new List<ParameterClass>();
                    this._UpdateComment = "Updated by user: [" + GlobalClass.User.UserName + "] IP Address: [" + GlobalClass.GetIPAddress() + "] Date: [" + this._LastUpdate.ToString() + "]";

                    // Update Record
                    sSql = "UPDATE dbo.TLDPackages SET ";
                    sSql = sSql + "PackageType = " + this.PackageType.ToString() + ", ";
                    sSql = sSql + "AuditType = '" + this.AuditType.ToString() + "', ";
                    sSql = sSql + "BatchID = " + GlobalClass.FormatIntegerValue(this._BatchID) + ", ";
                    sSql = sSql + "BatchPlanningID = " + GlobalClass.FormatIntegerValue(this.BatchPlanningID) + ", ";
                    sSql = sSql + "OperatorID = " + GlobalClass.FormatIntegerValue(this.OperatorID) + ", ";
                    sSql = sSql + "ContactID = " + GlobalClass.FormatIntegerValue(this.ContactID) + ", ";
                    
                    sSql = sSql + "ParticipationType = '" + GlobalClass.FormatStringValue(this.ParticipationType, 50) + "', ";
                    sSql = sSql + "ParticipationCategory = '" + GlobalClass.FormatStringValue(this.ParticipationCategory, 50) + "', ";                    
                    sSql = sSql + "CCode = '" + this._CCode.Trim() + "', ";
                    sSql = sSql + "SendToDataSheetsOption = '" + GlobalClass.FormatStringValue(this.SendToDataSheetsOption.Trim(), 50) + "', ";
                    sSql = sSql + "ReturnDataSheetsOption = '" + GlobalClass.FormatStringValue(this.ReturnDataSheetsOption.Trim(), 50) + "', ";
                    sSql = sSql + "CommunicationLanguage = '" + GlobalClass.FormatStringValue(this.CommunicationLanguage.Trim(), 50) + "', ";
                    sSql = sSql + "CoordinationNetwork = '" + GlobalClass.FormatStringValue(this.CoordinationNetwork.Trim(), 50) + "', ";
                    sSql = sSql + "PackageDescription = '" + GlobalClass.FormatStringValue(this.PackageDescription.Trim(), 2000) + "', ";
                    sSql = sSql + "PackageComment = '" +  GlobalClass.FormatStringValue(this.PackageComment.Trim(), 2000) + "', ";
                    sSql = sSql + "PackageStatus = '" + GlobalClass.FormatStringValue(this.PackageStatus.Trim(), 2000) + "', ";
                    
                    sSql = sSql + "DispatchedOn = " + "'" + this.DispatchedOn + "', ";
                    sSql = sSql + "DispatchedBy = " + "'" + GlobalClass.FormatStringValue(this.DispatchedBy.Trim(), 250) + "', ";
                    sSql = sSql + "ArchivedOn = " + "'" + this.ArchivedOn + "', ";
                    sSql = sSql + "ArchivedBy = " + "'" + GlobalClass.FormatStringValue(this.ArchivedBy.Trim(), 250) + "', ";

                    sSql = sSql + "UpdateComment = " + "'" + GlobalClass.FormatStringValue(this.UpdateComment.Trim(), 250) + "', ";
                    sSql = sSql + "LastUpdate = " + "'" + this.LastUpdate + "' ";
                    sSql = sSql + "WHERE PackageID = " + this.PackageID.ToString() + "";

                    sSql = sSql.Replace("'0001-01-01 00:00:00'", "NULL");

                    if (ParameterList.Count == 0)
                        iReturn = GlobalClass.ExecuteSQL(sSql);
                    else
                        iReturn = GlobalClass.ExecuteSQL(sSql, ParameterList);

                    if (this.TLDApplicationForm != null)
                    {
                        if (TLDApplicationForm.PackageID != this._PackageID)
                            TLDApplicationForm.PackageID = this._PackageID;

                        TLDApplicationForm.SaveApplicationForm();
                    }

                    this.SaveAttachments();

                    GlobalClass.LogUserAction(2, this.OperatorID, "User [#UserName#] Execute UPDATE on dbo.TLDPackages table", sSql);
                }
            }
            catch
            {
                iReturn = 0;
                GlobalClass.LogUserAction(-2, this.OperatorID, "ERROR! User [#UserName#] Execute INSERT/UPDATE on table dbo.TLDPackages", sSql);
                MessageBox.Show("Incorrect SQL statement.", "SQL Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            this.StateStatus = GlobalClass.sStateStatusClean;
            return iReturn;
        }

        public string PrepopulateApplicationForm()
        {
            string sReturn = string.Empty;

            this.TLDApplicationForm = new TLDApplicationFormClass(this);
            if (this.Batch != null)
                this.TLDApplicationForm.BatchNo = this.Batch.BatchNo;

            if (this.Operator != null)
            {
                this.TLDApplicationForm.OperatorName = this.Operator.OperatorName;
                this.TLDApplicationForm.Street = this.Operator.Street;
                this.TLDApplicationForm.City = this.Operator.City;
                this.TLDApplicationForm.Zip = this.Operator.Zip;
                this.TLDApplicationForm.State = this.Operator.State;
                this.TLDApplicationForm.Country = this.Operator.Country;

                this.TLDApplicationForm.InstitutionalEmail = this.Operator.InstitutionalEmail;
                this.TLDApplicationForm.InstitutionalTelephone1 = this.Operator.InstitutionalTelephone1;
                this.TLDApplicationForm.InstitutionalFax = this.Operator.InstitutionalFax;

                //pdfFormFields.SetField("InstitutionalTelephone2", this.InstitutionalTelephone2);
                //pdfFormFields.SetField("InstitutionalWebPage", this.InstitutionalWebPage);
            }

            ContactClass NewContact = null;
            if (this.AuditType == GlobalClass.sAuditTypeRT)
            {
                if (this.ParticipationType == GlobalClass.sParticipationTypeHospitals)
                    NewContact = this.Operator.GetRoleContact("AuditRT", "MedicalPhysicist");
                else
                    NewContact = this.Operator.GetRoleContact("AuditRT");
            }
            else if (this.AuditType == GlobalClass.sAuditTypeRP)
            {
                if (this.ParticipationType == GlobalClass.sParticipationTypeHospitals)
                    NewContact = this.Operator.GetRoleContact("AuditRP", "MedicalPhysicist");
                else
                    NewContact = this.Operator.GetRoleContact("AuditRP");
            }

            if (NewContact != null)
            {
                if (this._ContactID == -1)
                    this._ContactID = NewContact.ContactID;

                this.TLDApplicationForm.MedicalPhysicistContactID = NewContact.ContactID;
                this.TLDApplicationForm.MedicalPhysicistTitle = NewContact.Title;
                this.TLDApplicationForm.MedicalPhysicistFamilyName = NewContact.FamilyName;
                this.TLDApplicationForm.MedicalPhysicistFirstName = NewContact.FirstName;
                this.TLDApplicationForm.MedicalPhysicistPosition = NewContact.Position;
                this.TLDApplicationForm.MedicalPhysicistDepartment = NewContact.Department;
                this.TLDApplicationForm.MedicalPhysicistEmail = NewContact.Email;
                this.TLDApplicationForm.MedicalPhysicistTelephone1 = NewContact.Telephone1;
                this.TLDApplicationForm.MedicalPhysicistTelephone2 = NewContact.Telephone2;
            }

            this.TLDApplicationForm.CreatedOn = DateTime.Now;
            this.TLDApplicationForm.CreatedBy = "Created by user: [" + GlobalClass.User.UserName + "] IP Address: [" + GlobalClass.GetIPAddress() + "] Date: [" + this.TLDApplicationForm.CreatedOn.ToString() + "]";

            return sReturn;
        }

        public string ExportApplicationForm()
        {
            string sReturn = string.Empty;

            if (this.ParticipationType == GlobalClass.sParticipationTypeHospitals)
            {
                if (this.Batch != null)
                {
                    if (this.TLDApplicationForm != null)
                    {
                        string sPdfTemplate = GlobalClass.sApplicationStartupPath + "\\" + GlobalClass.TLDApplicatioForm;
                        if (this.CommunicationLanguage == GlobalClass.sCommunicationLanguageSpanish)
                            sPdfTemplate = GlobalClass.sApplicationStartupPath + "\\" + GlobalClass.TLDApplicatioFormSpanish;
                        else if (this.CommunicationLanguage == GlobalClass.sCommunicationLanguageRussian)
                            sPdfTemplate = GlobalClass.sApplicationStartupPath + "\\" + GlobalClass.TLDApplicatioFormRussian;

                        string sFilePath = GlobalClass.sApplicationTempPath;
                        string sFileName = "ApplicatioForm-" + Batch.BatchNo + "-" + this._CCode + "-" + this.OperatorID.ToString() + ".pdf";

                        if (!Directory.Exists(sFilePath))
                            Directory.CreateDirectory(sFilePath);

                        if (File.Exists(sPdfTemplate))
                        {
                            PdfReader pdfReader = new PdfReader(sPdfTemplate);
                            PdfStamper pdfStamper = new PdfStamper(pdfReader, new FileStream(sFilePath + sFileName, FileMode.Create), '\0', true);
                            AcroFields pdfFormFields = pdfStamper.AcroFields;

                            if (Batch.Powder != null)
                                pdfFormFields.SetField("AuditType", Batch.Powder.AuditType);

                            pdfFormFields.SetField("FormStatus", "PrePopulated");

                            pdfFormFields.SetField("PackageID", this.PackageID.ToString());
                            pdfFormFields.SetField("BatchNo", Batch.BatchNo);
                            pdfFormFields.SetFieldProperty("BatchNo", "setfflags", PdfFormField.FF_READ_ONLY, null);
                            pdfFormFields.RegenerateField("BatchNo");

                            pdfFormFields.SetField("BatchStartDate", GlobalClass.FormatTLDStringDateTimeValue(Batch.BatchStartDate));
                            pdfFormFields.SetFieldProperty("BatchStartDate", "setfflags", PdfFormField.FF_READ_ONLY, null);
                            pdfFormFields.RegenerateField("BatchStartDate");

                            pdfFormFields.SetField("BatchEndDate", GlobalClass.FormatTLDStringDateTimeValue(Batch.BatchEndDate));
                            pdfFormFields.SetFieldProperty("BatchEndDate", "setfflags", PdfFormField.FF_READ_ONLY, null);
                            pdfFormFields.RegenerateField("BatchEndDate");

                            if (this.Operator != null)
                            {
                                pdfFormFields.SetField("CCode", this.Operator.CCode);
                                pdfFormFields.SetField("OperatorID", this.Operator.OperatorID.ToString());
                                pdfFormFields.SetField("LabCode", this.Operator.LabCode.Trim());
                            }

                            if (this.TLDApplicationForm != null)
                            {
                                pdfFormFields.SetField("OperatorName", this.TLDApplicationForm.OperatorName);
                                pdfFormFields.SetField("Street", this.TLDApplicationForm.Street);
                                pdfFormFields.SetField("City", this.TLDApplicationForm.City);
                                pdfFormFields.SetField("Zip", this.TLDApplicationForm.Zip);
                                pdfFormFields.SetField("State", this.TLDApplicationForm.State);
                                pdfFormFields.SetField("Country", this.TLDApplicationForm.CCode);

                                pdfFormFields.SetField("InstitutionalEmail", this.TLDApplicationForm.InstitutionalEmail);
                                pdfFormFields.SetField("InstitutionalTelephone1", this.TLDApplicationForm.InstitutionalTelephone1);
                                pdfFormFields.SetField("InstitutionalFax", this.TLDApplicationForm.InstitutionalFax);


                                pdfFormFields.SetField("MedicalPhysicistContactID", this.TLDApplicationForm.MedicalPhysicistContactID.ToString());
                                pdfFormFields.SetField("MedicalPhysicistTitle", this.TLDApplicationForm.MedicalPhysicistTitle);
                                pdfFormFields.SetField("MedicalPhysicistFamilyName", this.TLDApplicationForm.MedicalPhysicistFamilyName);
                                pdfFormFields.SetField("MedicalPhysicistFirstName", this.TLDApplicationForm.MedicalPhysicistFirstName);
                                pdfFormFields.SetField("MedicalPhysicistPosition", this.TLDApplicationForm.MedicalPhysicistPosition);
                                pdfFormFields.SetField("MedicalPhysicistDepartment", this.TLDApplicationForm.MedicalPhysicistDepartment);
                                pdfFormFields.SetField("MedicalPhysicistEmail", this.TLDApplicationForm.MedicalPhysicistEmail);
                                pdfFormFields.SetField("MedicalPhysicistTelephone1", this.TLDApplicationForm.MedicalPhysicistTelephone1);
                                pdfFormFields.SetField("MedicalPhysicistTelephone2", this.TLDApplicationForm.MedicalPhysicistTelephone2);


                                pdfFormFields.SetField("Beam1", "Yes");
                                if (this.TLDApplicationForm.BeamType2 != "Off")
                                    pdfFormFields.SetField("Beam2", "Yes");
                                if (this.TLDApplicationForm.BeamType3 != "Off")
                                    pdfFormFields.SetField("Beam3", "Yes");

                                pdfFormFields.SetField("BeamType1", this.TLDApplicationForm.BeamType1);
                                pdfFormFields.SetField("BeamType2", this.TLDApplicationForm.BeamType2);
                                pdfFormFields.SetField("BeamType3", this.TLDApplicationForm.BeamType3);
                                //pdfFormFields.SetField("BeamType4", this.TLDApplicationForm.BeamType4);
                                //pdfFormFields.SetField("BeamType5", this.TLDApplicationForm.BeamType5);

                                pdfFormFields.SetField("ParticipationYear1", this.TLDApplicationForm.ParticipationYear1);
                                pdfFormFields.SetField("ParticipationYear2", this.TLDApplicationForm.ParticipationYear2);
                                pdfFormFields.SetField("ParticipationYear3", this.TLDApplicationForm.ParticipationYear3);
                                //pdfFormFields.SetField("ParticipationYear4", this.TLDApplicationForm.ParticipationYear4);
                                //pdfFormFields.SetField("ParticipationYear5", this.TLDApplicationForm.ParticipationYear5);

                                pdfFormFields.SetField("HolderStand", this.TLDApplicationForm.HolderStand);
                            }

                            // flatten the form to remove editting options, set it to false
                            // to leave the form open to subsequent manual edits
                            pdfStamper.FormFlattening = false;

                            pdfReader.Close();
                            // close the pdf
                            pdfStamper.Close();

                            if (File.Exists(sFilePath + sFileName))
                                sReturn = sFilePath + sFileName;
                        }
                        else
                            MessageBox.Show("Template " + sPdfTemplate + " does not exist.", "File does not exist.", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
            return sReturn;
        }

        public string ImportApplicationForm(string sPDFFileName)
        {
            string sReturn = string.Empty; // Ok

            if (File.Exists(sPDFFileName))
            {
                if (this.TLDApplicationForm != null)
                {
                    try
                    {
                        PdfReader pdfReader = new PdfReader(sPDFFileName);
                        if (pdfReader.AcroFields.Fields.Count > 0)
                        {
                            this.TLDApplicationForm.LastUpdate = DateTime.Now;
                            this.TLDApplicationForm.UpdateComment = "Application Form imported by user: [" + GlobalClass.User.UserName + "] IP Address: [" + GlobalClass.GetIPAddress() + "] Date: [" + this._LastUpdate + "]";

                            foreach (DictionaryEntry de in pdfReader.AcroFields.Fields)
                            {
                                string sCurrentField = de.Key.ToString().Trim();
                                string sCurrentValue = pdfReader.AcroFields.GetField(de.Key.ToString());

                                // Ignore these fields
                                if ((sCurrentField == "OperatorID") || (sCurrentField == "ReportID") || (sCurrentField == "ContactID") ||
                                    (sCurrentField == "CCode") || (sCurrentField == "LastUpdate") || (sCurrentField == "UpdateComment"))
                                    continue;

                                Type myType = this.TLDApplicationForm.GetType();
                                PropertyInfo Property = myType.GetProperty(sCurrentField);

                                if (Property != null)
                                    sReturn = sReturn + GlobalClass.SetPropertyValue(this.TLDApplicationForm, Property, sCurrentValue);
                            }


                            ContactClass MatchContact = this.Operator.MatchContact(this.TLDApplicationForm.MedicalPhysicistEmail);
                            if (MatchContact != null)
                            {
                                if (MatchContact.ContactID != this.TLDApplicationForm.MedicalPhysicistContactID)
                                    this.TLDApplicationForm.MedicalPhysicistContactID = MatchContact.ContactID;
                                else
                                    this.TLDApplicationForm.MedicalPhysicistContactID = -1;
                            }

                            MatchContact = this.Operator.MatchContact(this.TLDApplicationForm.RadiationOncologistEmail);
                            if (MatchContact != null)
                            {
                                if (MatchContact.ContactID != this.TLDApplicationForm.RadiationOncologistContactID)
                                    this.TLDApplicationForm.RadiationOncologistContactID = MatchContact.ContactID;
                                else
                                    this.TLDApplicationForm.RadiationOncologistContactID = -1;
                            }

                            if (this.TLDApplicationForm.MedicalPhysicistTitle == "Off")
                                this.TLDApplicationForm.MedicalPhysicistTitle = string.Empty;

                            if (this.TLDApplicationForm.RadiationOncologistTitle == "Off")
                                this.TLDApplicationForm.RadiationOncologistTitle = string.Empty;


                            if (this.TLDApplicationForm.BeamType1 == "Accelerator")
                                this.TLDApplicationForm.BeamType1 = GlobalClass.sBeamTypePhoton;

                            if (this.TLDApplicationForm.BeamType2 == "Accelerator")
                                this.TLDApplicationForm.BeamType2 = GlobalClass.sBeamTypePhoton;

                            if (this.TLDApplicationForm.BeamType3 == "Accelerator")
                                this.TLDApplicationForm.BeamType3 = GlobalClass.sBeamTypePhoton;

                            if (this.TLDApplicationForm.BeamType4 == "Accelerator")
                                this.TLDApplicationForm.BeamType4 = GlobalClass.sBeamTypePhoton;

                            if (this.TLDApplicationForm.BeamType5 == "Accelerator")
                                this.TLDApplicationForm.BeamType5 = GlobalClass.sBeamTypePhoton;

                            GlobalClass.LogUserAction(6, this.OperatorID, "Application Form import data for operator [" + this.OperatorID.ToString() + "]", "Import Data from [" + sPDFFileName + "] file.");
                        }
                        pdfReader.Close();
                    }
                    catch
                    {
                        sReturn = sReturn + System.Environment.NewLine + "ERROR! Application Form import data for operator [" + this.OperatorID.ToString() + "]. Import Value from [" + sPDFFileName + "] file";
                        GlobalClass.LogUserAction(-6, this.OperatorID, "ERROR! Application Form import data for operator [" + this.OperatorID.ToString() + "]", "Import Data from [" + sPDFFileName + "] file");
                    }
                }
            }
            else
                sReturn = sReturn + System.Environment.NewLine + "File " + sPDFFileName + " does not exists.";

            return sReturn;
        }

        public string GenerateDipatchTLDCoveringLetter(string sSaveAsFormat)
        {
            string sReturn = string.Empty;

            string sTemplate = string.Empty;

            if (this._AuditType == GlobalClass.sAuditTypeRT)
            {
                if (this.ParticipationType == GlobalClass.sParticipationTypeHospitals)
                {
                    if (this.PackageType == 1)
                    {
                        if (this.SendToDataSheetsOption == "HospitalCCNationalCoordinator")
                            sTemplate = GlobalClass.sApplicationStartupPath + "\\" + GlobalClass.TLDCoveringLetterDispatchTLDTemplateRTCountryCoordinatorCC;
                        else
                            sTemplate = GlobalClass.sApplicationStartupPath + "\\" + GlobalClass.TLDCoveringLetterDispatchTLDTemplateRTHospital;
                    }
                    else if (this.PackageType == 2)
                        sTemplate = GlobalClass.sApplicationStartupPath + "\\" + GlobalClass.TLDCoveringLetterDispatchTLDTemplateRTHospitalFolowUp;
                }
                else if (this.ParticipationType == GlobalClass.sParticipationTypeSSDL)
                {
                    if (this.PackageType == 1)
                        sTemplate = GlobalClass.sApplicationStartupPath + "\\" + GlobalClass.TLDCoveringLetterDispatchTLDTemplateRTSSDL;
                    else if (this.PackageType == 2)
                        sTemplate = GlobalClass.sApplicationStartupPath + "\\" + GlobalClass.TLDCoveringLetterDispatchTLDTemplateRTSSDLFolowUp;
                }
                else if (this.ParticipationType == GlobalClass.sParticipationTypeReference)
                    sTemplate = GlobalClass.sApplicationStartupPath + "\\" + GlobalClass.TLDCoveringLetterDispatchTLDTemplateRTPrimary;
                else if (this.ParticipationType == GlobalClass.sParticipationTypePrimary)
                    sTemplate = GlobalClass.sApplicationStartupPath + "\\" + GlobalClass.TLDCoveringLetterDispatchTLDTemplateRTPrimary;
            }

            if (this.TLDApplicationForm != null)
                if (this.TLDApplicationForm.MedicalPhysicist != null)
                    sReturn = GenerateCoveringLetter(this.TLDApplicationForm.MedicalPhysicist, sTemplate, "TLDCoveringLetter", sSaveAsFormat);

            return sReturn;
        }

        public string GenerateDipatchTLDProformaInvoice(string sSaveAsFormat)
        {
            string sReturn = string.Empty;

            string sTemplate = GlobalClass.sApplicationStartupPath + "\\" + GlobalClass.TLDCoveringLetterProformaInvoiceHospital;

            if (this.TLDApplicationForm != null)
                if (this.TLDApplicationForm.MedicalPhysicist != null)
                    sReturn = GenerateCoveringLetter(this.TLDApplicationForm.MedicalPhysicist, sTemplate, "ProformaInvoice", sSaveAsFormat);

            return sReturn;
        }

        public string GenerateDipatchCertificateCoveringLetter(string sSaveAsFormat)
        {
            string sReturn = string.Empty;

            string sTemplate = string.Empty;

            if (this._AuditType == GlobalClass.sAuditTypeRT)
            {
                if (this.Batch.BatchType == 1) //Hospital
                {
                    if (this.ParticipationType == GlobalClass.sParticipationTypeHospitals)
                        sTemplate = GlobalClass.sApplicationStartupPath + "\\" + GlobalClass.TLDCoveringLetterTemplateRTHospital;
                    else if (this.ParticipationType == GlobalClass.sParticipationTypeReference)
                        sTemplate = GlobalClass.sApplicationStartupPath + "\\" + GlobalClass.TLDCoveringLetterTemplateRTHospitalPrimary;
                    else if (this.ParticipationType == GlobalClass.sParticipationTypePrimary)
                        sTemplate = GlobalClass.sApplicationStartupPath + "\\" + GlobalClass.TLDCoveringLetterTemplateRTHospitalPrimary;

                }
                else if (this.Batch.BatchType == 2) //SSDL
                {
                    if (this.PackageType == 1)
                    {
                        if (this.ParticipationType == GlobalClass.sParticipationTypeSSDL)
                            sTemplate = GlobalClass.sApplicationStartupPath + "\\" + GlobalClass.TLDCoveringLetterTemplateRTSSDL;
                        else if (this.ParticipationType == GlobalClass.sParticipationTypeReference)
                            sTemplate = GlobalClass.sApplicationStartupPath + "\\" + GlobalClass.TLDCoveringLetterTemplateRTSSDLPrimary;
                        else if (this.ParticipationType == GlobalClass.sParticipationTypePrimary)
                            sTemplate = GlobalClass.sApplicationStartupPath + "\\" + GlobalClass.TLDCoveringLetterTemplateRTSSDLPrimary;
                    }
                }
            }
            else if (this._AuditType == GlobalClass.sAuditTypeRP)
            {
                if (this.Batch.BatchType == 2) //SSDL
                {
                    if (this.PackageType == 1)
                    {
                        if (this.ParticipationType == GlobalClass.sParticipationTypeSSDL)
                            sTemplate = GlobalClass.sApplicationStartupPath + "\\" + GlobalClass.TLDCoveringLetterTemplateRPSSDL;
                        else if (this.ParticipationType == GlobalClass.sParticipationTypeReference)
                            sTemplate = GlobalClass.sApplicationStartupPath + "\\" + GlobalClass.TLDCoveringLetterTemplateRPSSDLPrimary;
                        else if (this.ParticipationType == GlobalClass.sParticipationTypePrimary)
                            sTemplate = GlobalClass.sApplicationStartupPath + "\\" + GlobalClass.TLDCoveringLetterTemplateRPSSDLPrimary;
                    }
                }
            }

            if (this.Contact != null)
                sReturn = GenerateCoveringLetter(this.Contact, sTemplate, "CertificateCoveringLetter", sSaveAsFormat);

            return sReturn;
        }

        public string GenerateCoveringLetter(ContactClass Contact, string sTemplate, string sSaveAsName, string sSaveAsFormat)
        {
            string sReturn = string.Empty;

            string sFilePath = GlobalClass.sApplicationTempPath;

            Directory.SetCurrentDirectory(GlobalClass.sApplicationStartupPath);
            if (!Directory.Exists(sFilePath))
                Directory.CreateDirectory(sFilePath);


            if (File.Exists(sTemplate))
            {
                if (sSaveAsName == string.Empty)
                    sSaveAsName = "CoveringLetter";

                string sSaveFileName = string.Empty;
                string sDispatchedOn = GlobalClass.FormatDateTimeValue(this.DispatchedOn);
                //string sSetsNo = string.Empty;

                string sPackageComment = string.Empty;
                if (this.PackageComment.Trim() != string.Empty)
                    sPackageComment = " \nAdditional comments: " + this.PackageComment.Trim() + " \n";
                else
                    sPackageComment = " ";

                string sResults = string.Empty;
                //sResults = "TLD set #" + "\t" + "Within 5% acceptance limit?" + "\n";
                //sResults = sResults + "---------" + "\t\t\t" + "---------------------------" + "\n";

                int iCountCertificates = 0;
                int iCountResults = 0;
                int iCountCapsuleResults = 0;

                foreach (TLDSetClass TLDSet in this.GetTLDSetList())
                {
                    if (TLDSet.Certificate != null)
                    {
                        iCountCertificates = iCountCertificates + 1;
                        //iContactID = TLDSet.ContactID;
                        //sSetsNo = sSetsNo + TLDSet.SetNo + ",";

                        if (TLDSet.Certificate.DRatio > 0.0)
                        {
                            if (TLDSet.Certificate.DRatioInRange)
                            {
                                iCountResults = iCountResults + 1;
                                sResults = sResults + TLDSet.SetNo.ToString() + "\t\t\t\t\tYes\n";

                                if ((TLDSet.Certificate.Capsule1InRange) && (TLDSet.Certificate.Capsule2InRange) && (TLDSet.Certificate.Capsule3InRange))
                                    iCountCapsuleResults = iCountCapsuleResults + 0;
                                else
                                    iCountCapsuleResults = iCountCapsuleResults + 1;
                            }
                            else
                                sResults = sResults + TLDSet.SetNo.ToString() + "\t\t\t\t\tNo\n";
                        }
                    }
                }

                //if (sSetsNo.Length > 0)
                //    sSetsNo = sSetsNo.TrimEnd(',');

                // Check results and create summary of the results

                string sResultSummary = string.Empty;
                string sDOLF = string.Empty;

                if (this.ParticipationType != GlobalClass.sParticipationTypePrimary) //Primary
                {
                    if (this.PackageType == 1)
                    {
                        if (this.ParticipationType == GlobalClass.sParticipationTypeHospitals) //Hospitals
                        {
                            if (iCountResults == iCountCertificates)
                                sResultSummary = GlobalClass.Dictionary.GetExternalVariable("ResultSummary" + GlobalClass.sAuditTypeRT + "Hospital1");
                            else if ((iCountResults > 0) && (iCountResults != iCountCertificates))
                                sResultSummary = GlobalClass.Dictionary.GetExternalVariable("ResultSummary" + GlobalClass.sAuditTypeRT + "Hospital2");
                            else if ((iCountResults == 0) && (iCountCertificates > 0))
                                sResultSummary = GlobalClass.Dictionary.GetExternalVariable("ResultSummary" + GlobalClass.sAuditTypeRT + "Hospital3");
                        }
                        else
                        {
                            if (iCountResults == iCountCertificates)
                            {
                                sResultSummary = GlobalClass.Dictionary.GetExternalVariable("ResultSummary" + GlobalClass.sAuditTypeRT + "1");
                                sDOLF = GlobalClass.Dictionary.GetExternalVariable("DOLF.1604");
                            }
                            else
                            {
                                sResultSummary = GlobalClass.Dictionary.GetExternalVariable("ResultSummary" + GlobalClass.sAuditTypeRT + "2");
                                sDOLF = GlobalClass.Dictionary.GetExternalVariable("DOLF.1605");
                            }
                        }
                    }
                    else if (this.PackageType == 2)
                    {
                        if (this.ParticipationType == GlobalClass.sParticipationTypeHospitals) //Hospitals
                        {
                            if (iCountResults == iCountCertificates)
                                sResultSummary = GlobalClass.Dictionary.GetExternalVariable("ResultSummary" + GlobalClass.sAuditTypeRT + "HospitalFollowUp1");
                            else if ((iCountResults > 0) && (iCountResults != iCountCertificates))
                                sResultSummary = GlobalClass.Dictionary.GetExternalVariable("ResultSummary" + GlobalClass.sAuditTypeRT + "HospitalFollowUp2");
                            else if ((iCountResults == 0) && (iCountCertificates > 0))
                                sResultSummary = GlobalClass.Dictionary.GetExternalVariable("ResultSummary" + GlobalClass.sAuditTypeRT + "HospitalFollowUp3");
                        }
                        else
                        {
                            if (iCountResults == iCountCertificates)
                                sResultSummary = GlobalClass.Dictionary.GetExternalVariable("ResultSummaryFollowUp" + GlobalClass.sAuditTypeRT + "1");
                            else
                                sResultSummary = GlobalClass.Dictionary.GetExternalVariable("ResultSummaryFollowUp" + GlobalClass.sAuditTypeRT + "2");
                        }
                    }
                }
                else
                    sDOLF = GlobalClass.Dictionary.GetExternalVariable("DOLF.1603");


                //OperatorClass Operator = GlobalClass.Manager.GetOperator(this.OperatorID);
                if (this.Operator != null)
                {
                    
                    NationalCoordinatorClass NationalCoordinator = null;
                    NationalCoordinatorClass PAHOCoordinator = null;


                    if (Contact == null)
                    {
                        if (this.TLDApplicationForm != null)
                        {
                            if (this.TLDApplicationForm.MedicalPhysicist != null)
                                Contact = this.TLDApplicationForm.MedicalPhysicist;
                        }
                    }

                    if (this.Country != null)
                    {
                        NationalCoordinator = this.Country.GetNationalCoordinator();
                        PAHOCoordinator = this.Country.GetPAHOCoordinator();
                    }


                    //OBJECT OF MISSING "NULL VALUE"
                    Object oMissing = System.Reflection.Missing.Value;

                    //OBJECTS OF FALSE AND TRUE
                    Object oTrue = true;
                    Object oFalse = false;

                    //CREATING OBJECTS OF WORD AND DOCUMENT
                    Word._Application oWord = null;
                    Word._Document oWordDoc = null;

                    try
                    {
                        oWord = new Word.Application();

                        //SETTING THE VISIBILITY TO TRUE
                        oWord.Visible = false;

                        //THE LOCATION OF THE TEMPLATE FILE ON THE MACHINE
                        Object oTemplatePath = sTemplate;

                        //ADDING A NEW DOCUMENT FROM A TEMPLATE
                        oWordDoc = oWord.Documents.Add(
                            /* ref object Template */ ref oTemplatePath,
                            /* ref object NewTemplate */ ref oMissing,
                            /* ref object DocumentType */ ref oMissing,
                            /* ref object Visible */ ref oMissing);

                        int iTotalFields = 0;
                        foreach (Word.Field myMergeField in oWordDoc.Fields)
                        {
                            iTotalFields = iTotalFields + 1;
                            Word.Range rngFieldCode = myMergeField.Code;
                            string sFieldText = rngFieldCode.Text;

                            // ONLY GETTING THE MAILMERGE FIELDS
                            if (sFieldText.StartsWith(" MERGEFIELD"))
                            {
                                // THE TEXT COMES IN THE FORMAT OF 
                                // MERGEFIELD  MyFieldName  \\* MERGEFORMAT
                                // THIS HAS TO BE EDITED TO GET ONLY THE FIELDNAME "MyFieldName"
                                Int32 iEndMerge = sFieldText.IndexOf("\\");
                                //Int32 iFieldNameLength = sFieldText.Length - iEndMerge;
                                //string sFieldName = sFieldText.Substring(11, iEndMerge - 11);
                                string sFieldName = sFieldText.Replace("MERGEFIELD", "").Trim();
                                // GIVES THE FIELDNAMES AS THE USER HAD ENTERED IN .dot FILE
                                sFieldName = sFieldName.Trim();

                                // **** FIELD REPLACEMENT IMPLEMENTATION GOES HERE ****//
                                // THE PROGRAMMER CAN HAVE HIS OWN IMPLEMENTATIONS HERE
                                myMergeField.Select();
                                if (sFieldName == "OperatorName")
                                {
                                    if (this.Operator != null)
                                        oWord.Selection.TypeText(GlobalClass.FormatMergeString(this.Operator.OperatorName));
                                    else
                                        oWord.Selection.Delete(ref oMissing, ref oMissing);
                                }
                                else if (sFieldName == "Street")
                                {
                                    if (this.Operator != null)
                                    {
                                        string sAddress = string.Empty;
                                        if (this.Operator.Street.Trim() != string.Empty)
                                            sAddress = this.Operator.Street.Trim();

                                        if (this.Operator.POBox.Trim() != string.Empty)
                                        {
                                            if (sAddress != string.Empty)
                                                sAddress = sAddress + ", " + this.Operator.POBox.Trim();
                                            else
                                                sAddress = this.Operator.POBox.Trim();
                                        }

                                        oWord.Selection.TypeText(GlobalClass.FormatMergeString(sAddress));
                                    }
                                    else
                                        oWord.Selection.Delete(ref oMissing, ref oMissing);
                                }
                                else if (sFieldName == "City")
                                {
                                    if (this.Operator != null)
                                        oWord.Selection.TypeText(GlobalClass.FormatMergeString(this.Operator.City));
                                    else
                                        oWord.Selection.Delete(ref oMissing, ref oMissing);
                                }
                                else if (sFieldName == "CITY")
                                {
                                    if (this.Operator != null)
                                        oWord.Selection.TypeText(GlobalClass.FormatMergeString(this.Operator.City.ToUpper()));
                                    else
                                        oWord.Selection.Delete(ref oMissing, ref oMissing);
                                }
                                else if (sFieldName == "Zip")
                                {
                                    if (this.Operator != null)
                                        oWord.Selection.TypeText(GlobalClass.FormatMergeString(this.Operator.Zip));
                                    else
                                        oWord.Selection.Delete(ref oMissing, ref oMissing);
                                }
                                else if (sFieldName == "Country")
                                {
                                    if (this.Operator != null)
                                        oWord.Selection.TypeText(GlobalClass.FormatMergeString(this.Operator.Country));
                                    else
                                        oWord.Selection.Delete(ref oMissing, ref oMissing);
                                }
                                else if (sFieldName == "COUNTRY")
                                {
                                    if (this.Operator != null)
                                        oWord.Selection.TypeText(GlobalClass.FormatMergeString(this.Operator.Country.ToUpper()));
                                    else
                                        oWord.Selection.Delete(ref oMissing, ref oMissing);
                                }
                                else if (sFieldName == "ContactTitle")
                                {
                                    if (Contact != null)
                                        oWord.Selection.TypeText(GlobalClass.FormatMergeString(Contact.Title));
                                    else
                                        oWord.Selection.Delete(ref oMissing, ref oMissing);
                                }
                                else if (sFieldName == "ContactFirstName")
                                {
                                    if (Contact != null)
                                        oWord.Selection.TypeText(GlobalClass.FormatMergeString(Contact.FirstName));
                                    else
                                        oWord.Selection.Delete(ref oMissing, ref oMissing);
                                }
                                else if (sFieldName == "ContactFamilyName")
                                {
                                    if (Contact != null)
                                        oWord.Selection.TypeText(GlobalClass.FormatMergeString(Contact.FamilyName));
                                    else
                                        oWord.Selection.Delete(ref oMissing, ref oMissing);
                                }
                                else if (sFieldName == "PhysicistName")
                                {
                                    if (Contact != null)
                                    {
                                        string sPhysicistName = Contact.Title + " " + Contact.FamilyName + " " + Contact.FirstName;
                                        oWord.Selection.TypeText(GlobalClass.FormatMergeString(sPhysicistName.Trim()));
                                    }
                                    else
                                        oWord.Selection.Delete(ref oMissing, ref oMissing);
                                }
                                else if (sFieldName == "ContactEmail")
                                {
                                    if (Contact != null)
                                        oWord.Selection.TypeText(GlobalClass.FormatMergeString(Contact.Email));
                                    else
                                        oWord.Selection.Delete(ref oMissing, ref oMissing);
                                }
                                else if (sFieldName == "ContactTelephone1")
                                {
                                    if (Contact != null)
                                        oWord.Selection.TypeText(GlobalClass.FormatMergeString(Contact.Telephone1));
                                    else
                                        oWord.Selection.Delete(ref oMissing, ref oMissing);
                                }
                                else if (sFieldName == "ContactTelephone2")
                                {
                                    if (Contact != null)
                                        oWord.Selection.TypeText(GlobalClass.FormatMergeString(Contact.Telephone2));
                                    else
                                        oWord.Selection.Delete(ref oMissing, ref oMissing);
                                }
                                else if (sFieldName == "BatchNo")
                                {
                                    if (this.Batch != null)
                                        oWord.Selection.TypeText(GlobalClass.FormatMergeString(this.Batch.BatchNo));
                                    else
                                        oWord.Selection.Delete(ref oMissing, ref oMissing);
                                }
                                else if (sFieldName == "BatchYear")
                                {
                                    if (this.Batch != null)
                                        oWord.Selection.TypeText(GlobalClass.FormatMergeString(this.Batch.BatchYear));
                                    else
                                        oWord.Selection.Delete(ref oMissing, ref oMissing);
                                }
                                else if (sFieldName == "BatchWindow")
                                {
                                    if (this.Batch != null)
                                        oWord.Selection.TypeText(GlobalClass.FormatMergeString(this.Batch.BatchWindow));
                                    else
                                        oWord.Selection.Delete(ref oMissing, ref oMissing);
                                }
                                else if (sFieldName == "NextBatchYear")
                                {
                                    if (this.Batch != null)
                                    {
                                        int iNextBatchYear = int.Parse(this.Batch.BatchYear) + 1;
                                        oWord.Selection.TypeText(GlobalClass.FormatMergeString(iNextBatchYear.ToString()));
                                    }
                                    else
                                        oWord.Selection.Delete(ref oMissing, ref oMissing);
                                }
                                else if (sFieldName == "TLDSetCount")
                                {
                                    int iTLDSetCount = 0;
                                    foreach (TLDSetClass TLDSet in this.GetTLDSetList())
                                        iTLDSetCount = iTLDSetCount + 1;

                                    if (iTLDSetCount > 0)
                                        oWord.Selection.TypeText(GlobalClass.FormatMergeString(iTLDSetCount.ToString()));
                                    else
                                        oWord.Selection.Delete(ref oMissing, ref oMissing);
                                }
                                else if ((sFieldName == "SetNos") || (sFieldName == "SetsNo"))
                                    oWord.Selection.TypeText(GlobalClass.FormatMergeString(this.SetNos));
                                else if (sFieldName == "OriginalSetNos")
                                {
                                    if (this.PackageType == 2)
                                        oWord.Selection.TypeText(GlobalClass.FormatMergeString(this.OriginalSetNos));
                                    else
                                        oWord.Selection.Delete(ref oMissing, ref oMissing);
                                }
                                else if (sFieldName == "PackageComment")
                                    oWord.Selection.TypeText(sPackageComment);
                                else if (sFieldName == "Results")
                                    oWord.Selection.TypeText(GlobalClass.FormatMergeString(sResults));
                                else if (sFieldName == "ResultSummary")
                                    oWord.Selection.TypeText(GlobalClass.FormatMergeString(sResultSummary));
                                else if (sFieldName == "DOLF")
                                    oWord.Selection.TypeText(GlobalClass.FormatMergeString(sDOLF));
                                else if (sFieldName == "SignatureDate")
                                    oWord.Selection.TypeText(GlobalClass.FormatMergeString(sDispatchedOn));
                                else if (sFieldName == "PreparationDate")
                                    oWord.Selection.TypeText(GlobalClass.FormatMergeString(sDispatchedOn));
                                else if (sFieldName == "Date")
                                    oWord.Selection.TypeText(GlobalClass.FormatMergeString(GlobalClass.FormatDateTimeValue(DateTime.Now)));
                                else if (sFieldName == "PlasticHolder")
                                {
                                    if (this.TLDApplicationForm != null)
                                    {
                                        if ((this.TLDApplicationForm.HolderStand == "Yes") || (this.TLDApplicationForm.HolderStand == "Horisontal") || (this.TLDApplicationForm.HolderStand == "Vertical"))
                                            oWord.Selection.TypeText(GlobalClass.FormatMergeString(", a plastic holder"));
                                        else
                                            oWord.Selection.Delete(ref oMissing, ref oMissing);
                                    }
                                    else
                                        oWord.Selection.Delete(ref oMissing, ref oMissing);
                                }
                                else if (sFieldName == "CoordinatorName")
                                {
                                    if (NationalCoordinator != null)
                                        oWord.Selection.TypeText(GlobalClass.FormatMergeString(NationalCoordinator.CoordinatorName));
                                    else
                                        oWord.Selection.Delete(ref oMissing, ref oMissing);
                                }
                                else if (sFieldName == "CoordinatorInstitutionName")
                                {
                                    if (NationalCoordinator != null)
                                        oWord.Selection.TypeText(GlobalClass.FormatMergeString(NationalCoordinator.InstitutionName));
                                    else
                                        oWord.Selection.Delete(ref oMissing, ref oMissing);
                                }
                                else if (sFieldName == "CoordinatorDepartment")
                                {
                                    if (NationalCoordinator != null)
                                        oWord.Selection.TypeText(GlobalClass.FormatMergeString(NationalCoordinator.Department));
                                    else
                                        oWord.Selection.Delete(ref oMissing, ref oMissing);
                                }
                                else if (sFieldName == "CoordinatorEmail")
                                {
                                    if (NationalCoordinator != null)
                                        oWord.Selection.TypeText(GlobalClass.FormatMergeString(NationalCoordinator.Email));
                                    else
                                        oWord.Selection.Delete(ref oMissing, ref oMissing);
                                }
                                else if (sFieldName == "CoordinatorPhone")
                                {
                                    if (NationalCoordinator != null)
                                        oWord.Selection.TypeText(GlobalClass.FormatMergeString(NationalCoordinator.Phone));
                                    else
                                        oWord.Selection.Delete(ref oMissing, ref oMissing);
                                }
                                else if (sFieldName == "CoordinatorStreet")
                                {
                                    if (NationalCoordinator != null)
                                        oWord.Selection.TypeText(GlobalClass.FormatMergeString(NationalCoordinator.Street));
                                    else
                                        oWord.Selection.Delete(ref oMissing, ref oMissing);
                                }
                                else if (sFieldName == "CoordinatorCity")
                                {
                                    if (NationalCoordinator != null)
                                        oWord.Selection.TypeText(GlobalClass.FormatMergeString(NationalCoordinator.City));
                                    else
                                        oWord.Selection.Delete(ref oMissing, ref oMissing);
                                }
                                else if (sFieldName == "CoordinatorZip")
                                {
                                    if (NationalCoordinator != null)
                                        oWord.Selection.TypeText(GlobalClass.FormatMergeString(NationalCoordinator.Zip));
                                    else
                                        oWord.Selection.Delete(ref oMissing, ref oMissing);
                                }
                                else if (sFieldName == "CoordinatorPOBox")
                                {
                                    if (NationalCoordinator != null)
                                        oWord.Selection.TypeText(GlobalClass.FormatMergeString(NationalCoordinator.POBox));
                                    else
                                        oWord.Selection.Delete(ref oMissing, ref oMissing);
                                }
                                else if (sFieldName == "CoordinatorState")
                                {
                                    if (NationalCoordinator != null)
                                        oWord.Selection.TypeText(GlobalClass.FormatMergeString(NationalCoordinator.State));
                                    else
                                        oWord.Selection.Delete(ref oMissing, ref oMissing);
                                }
                                else if (sFieldName == "CoordinatorCountry")
                                {
                                    if (NationalCoordinator != null)
                                        oWord.Selection.TypeText(GlobalClass.FormatMergeString(NationalCoordinator.Country));
                                    else
                                        oWord.Selection.Delete(ref oMissing, ref oMissing);
                                }
                                else if (sFieldName == "COORDINATORCOUNTRY")
                                {
                                    if (NationalCoordinator != null)
                                        oWord.Selection.TypeText(GlobalClass.FormatMergeString(NationalCoordinator.Country.ToUpper()));
                                    else
                                        oWord.Selection.Delete(ref oMissing, ref oMissing);
                                }
                                else if (sFieldName == "PAHOCoordinatorName")
                                {
                                    if (NationalCoordinator != null)
                                        oWord.Selection.TypeText(GlobalClass.FormatMergeString(PAHOCoordinator.CoordinatorName));
                                    else
                                        oWord.Selection.Delete(ref oMissing, ref oMissing);
                                }
                                else if (sFieldName == "PAHOCoordinatorEmail")
                                {
                                    if (NationalCoordinator != null)
                                        oWord.Selection.TypeText(GlobalClass.FormatMergeString(PAHOCoordinator.Email));
                                    else
                                        oWord.Selection.Delete(ref oMissing, ref oMissing);
                                }
                                else if (sFieldName == "PAHOCoordinatorPhone")
                                {
                                    if (NationalCoordinator != null)
                                        oWord.Selection.TypeText(GlobalClass.FormatMergeString(PAHOCoordinator.Phone));
                                    else
                                        oWord.Selection.Delete(ref oMissing, ref oMissing);
                                }
                                else if (sFieldName == "PAHOCoordinatorInstitutionName")
                                {
                                    if (PAHOCoordinator != null)
                                        oWord.Selection.TypeText(GlobalClass.FormatMergeString(PAHOCoordinator.InstitutionName));
                                    else
                                        oWord.Selection.Delete(ref oMissing, ref oMissing);
                                }
                                else if (sFieldName == "PAHOCoordinatorDepartment")
                                {
                                    if (PAHOCoordinator != null)
                                        oWord.Selection.TypeText(GlobalClass.FormatMergeString(PAHOCoordinator.Department));
                                    else
                                        oWord.Selection.Delete(ref oMissing, ref oMissing);
                                }
                                else if (sFieldName == "PAHOCoordinatorStreet")
                                {
                                    if (NationalCoordinator != null)
                                        oWord.Selection.TypeText(GlobalClass.FormatMergeString(PAHOCoordinator.Street));
                                    else
                                        oWord.Selection.Delete(ref oMissing, ref oMissing);
                                }
                                else if (sFieldName == "PAHOCoordinatorCity")
                                {
                                    if (NationalCoordinator != null)
                                        oWord.Selection.TypeText(GlobalClass.FormatMergeString(PAHOCoordinator.City));
                                    else
                                        oWord.Selection.Delete(ref oMissing, ref oMissing);
                                }
                                else if (sFieldName == "PAHOCoordinatorZip")
                                {
                                    if (NationalCoordinator != null)
                                        oWord.Selection.TypeText(GlobalClass.FormatMergeString(PAHOCoordinator.Zip));
                                    else
                                        oWord.Selection.Delete(ref oMissing, ref oMissing);
                                }
                                else if (sFieldName == "PAHOCoordinatorPOBox")
                                {
                                    if (NationalCoordinator != null)
                                        oWord.Selection.TypeText(GlobalClass.FormatMergeString(PAHOCoordinator.POBox));
                                    else
                                        oWord.Selection.Delete(ref oMissing, ref oMissing);
                                }
                                else if (sFieldName == "PAHOCoordinatorState")
                                {
                                    if (NationalCoordinator != null)
                                        oWord.Selection.TypeText(GlobalClass.FormatMergeString(PAHOCoordinator.State));
                                    else
                                        oWord.Selection.Delete(ref oMissing, ref oMissing);
                                }
                                else if (sFieldName == "PAHOCoordinatorCountry")
                                {
                                    if (NationalCoordinator != null)
                                        oWord.Selection.TypeText(GlobalClass.FormatMergeString(PAHOCoordinator.Country));
                                    else
                                        oWord.Selection.Delete(ref oMissing, ref oMissing);
                                }
                                else if (sFieldName == "PAHOCOORDINATORCOUNTRY")
                                {
                                    if (NationalCoordinator != null)
                                        oWord.Selection.TypeText(GlobalClass.FormatMergeString(PAHOCoordinator.Country.ToUpper()));
                                    else
                                        oWord.Selection.Delete(ref oMissing, ref oMissing);
                                }
                            }
                        }

                        //THE LOCATION WHERE THE FILE NEEDS TO BE SAVED

                        Object oFileFormat = (Object)Word.WdSaveFormat.wdFormatDocumentDefault;
                        Object oSaveAsFile = (Object)sSaveFileName;

                        if (sSaveAsFormat == GlobalClass.sFormatPDF)
                        {
                            oFileFormat = (Object)Word.WdSaveFormat.wdFormatPDF;
                            sSaveFileName = sFilePath + sSaveAsName + this.Batch.BatchNo + "-" + this.OperatorID.ToString() + ".pdf";

                            oWordDoc.ExportAsFixedFormat(sSaveFileName, Word.WdExportFormat.wdExportFormatPDF, false, Word.WdExportOptimizeFor.wdExportOptimizeForPrint);
                        }
                        else if (sSaveAsFormat == GlobalClass.sFormatWord97)
                        {
                            oFileFormat = (Object)Word.WdSaveFormat.wdFormatDocument97; // Word 97 format
                            sSaveFileName = sFilePath + sSaveAsName + this.Batch.BatchNo + "-" + this.OperatorID.ToString() + ".doc";
                            oSaveAsFile = (Object)sSaveFileName;

                            oWordDoc.SaveAs(
                                /* ref object FileName */ ref oSaveAsFile,
                                /* ref object FileFormat */ ref oFileFormat,
                                /* ref object LockComments */ ref oMissing,
                                /* ref object Password */ ref oMissing,
                                /* ref object AddToRecentFiles */ ref oMissing,
                                /* ref object WritePassword */ ref oMissing,
                                /* ref object ReadOnlyRecommended */ ref oMissing,
                                /* ref object EmbedTrueTypeFonts */ ref oMissing,
                                /* ref object SaveNativePictureFormat */ ref oMissing,
                                /* ref object SaveFormsData */ ref oMissing,
                                /* ref object SaveAsAOCELetter */ ref oMissing,
                                /* ref object Encoding */ ref oMissing,
                                /* ref object InsertLineBreaks */ ref oMissing,
                                /* ref object AllowSubstitutions */ ref oMissing,
                                /* ref object LineEncoding */ ref oMissing,
                                /* ref object AddBiDiMarks */ ref oMissing);
                        }
                        else if (sSaveAsFormat == GlobalClass.sFormatWord2010)
                        {
                            oFileFormat = (Object)Word.WdSaveFormat.wdFormatDocumentDefault; // Word 2010 format                            
                            sSaveFileName = sFilePath + sSaveAsName + this.Batch.BatchNo + "-" + this.OperatorID.ToString() + ".docx";
                            oSaveAsFile = (Object)sSaveFileName;
                            oWordDoc.SaveAs(
                                /* ref object FileName */ ref oSaveAsFile,
                                /* ref object FileFormat */ ref oFileFormat,
                                /* ref object LockComments */ ref oMissing,
                                /* ref object Password */ ref oMissing,
                                /* ref object AddToRecentFiles */ ref oMissing,
                                /* ref object WritePassword */ ref oMissing,
                                /* ref object ReadOnlyRecommended */ ref oMissing,
                                /* ref object EmbedTrueTypeFonts */ ref oMissing,
                                /* ref object SaveNativePictureFormat */ ref oMissing,
                                /* ref object SaveFormsData */ ref oMissing,
                                /* ref object SaveAsAOCELetter */ ref oMissing,
                                /* ref object Encoding */ ref oMissing,
                                /* ref object InsertLineBreaks */ ref oMissing,
                                /* ref object AllowSubstitutions */ ref oMissing,
                                /* ref object LineEncoding */ ref oMissing,
                                /* ref object AddBiDiMarks */ ref oMissing);
                        }
                    }
                    catch (COMException)
                    {
                    }
                    finally
                    {
                        // Closing the Word Instance
                        if (oWordDoc != null)
                        {
                            oWordDoc.Close(
                                /* ref object SaveChanges */ false,//ref oMissing,
                                /* ref object OriginalFormat */ ref oMissing,
                                /* ref object RouteDocument */ ref oMissing);
                            oWordDoc = null;
                        }
                        if (oWord != null)
                        {
                            oWord.Quit(
                                /* ref object SaveChanges */ ref oMissing,
                                /* ref object OriginalFormat */ ref oMissing,
                                /* ref object RouteDocument */ ref oMissing);

                            System.Runtime.InteropServices.Marshal.ReleaseComObject(oWord);
                        }
                    }

                    GC.Collect();
                    GC.WaitForPendingFinalizers();

                    GC.Collect();
                    GC.WaitForPendingFinalizers();

                    sReturn = sSaveFileName;
                }
            }

            return sReturn;
        }

        public void ArchivePackage(DateTime dtpArchivedOn)
        {
            List<TLDSetClass> TLDSetList = this.GetTLDSetList();
            if (TLDSetList != null)
            {
                if (TLDSetList.Count == 0)
                {
                    this._ArchivedOn = dtpArchivedOn;
                    this._ArchivedBy = "Package archived by user [" + GlobalClass.User.UserName.Trim() + "]; IP Address: [" + GlobalClass.GetIPAddress() + "] Date: [" + dtpArchivedOn.ToString() + "]";
                    string sSql = string.Empty;

                    sSql = sSql + "UPDATE dbo.TLDPackages SET ";
                    sSql = sSql + "ArchivedOn = '" + this._ArchivedOn + "', ";
                    sSql = sSql + "ArchivedBy = '" + GlobalClass.FormatStringValue(this._ArchivedBy, 250) + "' ";
                    sSql = sSql + "WHERE PackageID = " + this._PackageID.ToString();

                    sSql = sSql.Replace("'0001-01-01 00:00:00'", "NULL");

                    GlobalClass.ExecuteSQL(sSql);
                }
            }
        }

        public string DispatchPackage(DateTime dtpDispatchedOn)
        {
            int iCountCertificates = 0;
            int iCountSignatures = 0;
            string sErrorString = string.Empty;

            string sContactIDs = string.Empty;

            foreach (TLDSetClass TLDSet in this.GetTLDSetList())
            {
                if (TLDSet.Certificate != null)
                {
                    iCountCertificates = iCountCertificates + 1;

                    SignatureClass Signature1 = TLDSet.Certificate.GetSignature(GlobalClass.iSignatureTLDCertificateSignByOfficer);
                    SignatureClass Signature2 = TLDSet.Certificate.GetSignature(GlobalClass.iSignatureTLDCertificateSignBySectionHead);

                    if ((Signature1 != null) && (Signature2 != null))
                        iCountSignatures = iCountSignatures + 1;

                    if (TLDSet.Contact != null)
                        if (sContactIDs.Contains(TLDSet.Contact.ContactID.ToString() + ",") == false)
                            sContactIDs = sContactIDs + TLDSet.Contact.ContactID.ToString() + ",";

                }
            }
            sContactIDs = sContactIDs.TrimEnd(',');

            if ((iCountSignatures != iCountCertificates) && (iCountSignatures > 0))
                sErrorString = sErrorString + "Some certificates are not signed. Only signed certificate(s) could be dispatched." + System.Environment.NewLine;
            else if ((iCountSignatures != iCountCertificates) && (iCountSignatures == 0))
                sErrorString = sErrorString + "There no certificates to be dispatched." + System.Environment.NewLine;

            if (sContactIDs.Contains("-1"))
                sErrorString = sErrorString + "There are one or more undefined contact person to dispatch certificates. Chack contact details for each Set No." + System.Environment.NewLine;
            else if ((sContactIDs.Split(',').Length > 1) && (this._ContactID == -1))
                sErrorString = sErrorString + "There are more than one contact person to dispatch certificates. Chack contact details for each Set No." + System.Environment.NewLine;

            // All Certificates have been signed
            if (sErrorString == string.Empty)
            {
                foreach (TLDSetClass TLDSet in this.GetTLDSetList())
                {
                    if (TLDSet.Certificate != null)
                    {
                        AttachmentClass CertificateAttachment = TLDSet.Certificate.GetAttachment(GlobalClass.iAttachmentTLDCertificate);

                        // Regenerate certificates if they are not archived
                        if ((CertificateAttachment != null) && (TLDSet.SetArchived == false))
                        {
                            CertificateAttachment.RemoveAttachment();
                            TLDSet.Certificate.AttachmentList.Remove(CertificateAttachment);
                            // Should be null
                            CertificateAttachment = TLDSet.Certificate.GetAttachment(GlobalClass.iAttachmentTLDCertificate);
                        }

                        if (CertificateAttachment == null)
                        {
                            string sFileName = TLDSet.Certificate.GenerateCertificate(GlobalClass.sFormatPDF);
                            if (File.Exists(sFileName))
                            {
                                string sError = TLDSet.Certificate.AttachDocument(GlobalClass.iAttachmentTLDCertificate, DateTime.Now, sFileName, null);
                                if (sError == string.Empty)
                                {
                                    TLDSet.Certificate.SaveAttachments();
                                    CertificateAttachment = TLDSet.Certificate.GetAttachment(GlobalClass.iAttachmentTLDCertificate);
                                }
                                else
                                    sErrorString = sErrorString + sError + System.Environment.NewLine;
                            }
                        }

                        if (sErrorString == string.Empty)
                        {
                            if (CertificateAttachment != null)
                            {
                                //SignatureClass Signature3 = TLDSet.Certificate.GetSignature(GlobalClass.iSignatureTLDCertificateDispatched);
                                //if (Signature3 == null)
                                TLDSet.Certificate.SignCertificate(GlobalClass.iSignatureTLDCertificateDispatched, dtpDispatchedOn);
                            }
                        }
                    }
                }

                // Dispatching the Package
                this.DispatchedOn = dtpDispatchedOn;
                this.DispatchedBy = "Package dispatched by user [" + GlobalClass.User.UserName.Trim() + "]; IP Address: [" + GlobalClass.GetIPAddress() + "] Date: [" + dtpDispatchedOn.ToString() + "]";

                if (this._ContactID == -1)
                {
                    if (sContactIDs.Split(',').Length == 1)
                    {
                        int iContactID = -1;
                        if (int.TryParse(sContactIDs.Split(',')[0].Trim(), out iContactID))
                            this.ContactID = iContactID;
                    }
                }

                // Save Package
                string sCheckBeforeSave = this.CheckBeforeSave();

                if (sCheckBeforeSave == string.Empty)
                {
                    if (this.TLDApplicationForm != null)
                    {
                        sCheckBeforeSave = this.TLDApplicationForm.CheckBeforeSave();
                        if (sCheckBeforeSave == string.Empty)
                            this.SavePackage(); // will also save TLDApplicationForm
                        else
                            sErrorString = "TLD Application Form validation" + System.Environment.NewLine + sCheckBeforeSave;
                    }
                    else
                        this.SavePackage(); // will also save TLDApplicationForm
                }
                else
                    sErrorString = "TLD Package validation" + System.Environment.NewLine + sCheckBeforeSave;
            }
            else
                sErrorString = "Dispatching error for package No " + this.PackageID.ToString() + " - " + this.Country.CountryName +
                    System.Environment.NewLine + sErrorString;

            return sErrorString;
        }

        public TLDSetClass InitiateTLDRun(string sSetNo, string sSetBeamType)
        {
            TLDSetClass TLDSet = null;

            if (this.Batch != null)
            {
                if (this.Batch.Powder != null)
                {
                    if (this.Operator != null)
                    {
                        TLDSet = new TLDSetClass(this.Operator);

                        TLDSet.AuditType = Batch.Powder.AuditType;
                        TLDSet.ParticipationType = this._ParticipationType;
                        TLDSet.OperatorID = this.Operator.OperatorID;
                        TLDSet.CCode = this.Operator.CCode;
                        TLDSet.PackageID = this._PackageID;

                        if ((sSetBeamType != string.Empty) || (sSetBeamType != "Off"))
                        {
                            DictionaryTypeClass DictionaryType = GlobalClass.Dictionary.GetDictionaryType(GlobalClass.Dictionary.GetDictionary("BeamType"), sSetBeamType);
                            if (DictionaryType != null)
                                TLDSet.SetBeamType = DictionaryType.ItemCode;
                        }

                        TLDSet.UnitID = -1;
                        TLDSet.BeamID = -1;
                        TLDSet.SetType = this.PackageType; // 1-FirstIrradiation 

                        TLDSet.BatchID = Batch.BatchID;
                        TLDSet.SetNo = sSetNo;

                        TLDSet.LastParticipation = TLDSet.GetLastParticipation();

                        if (this.TLDApplicationForm != null)
                        {
                            if (this.TLDApplicationForm.MedicalPhysicist != null)
                                TLDSet.ContactID = this.TLDApplicationForm.MedicalPhysicist.ContactID;
                            else if (this.TLDApplicationForm.RadiationOncologist != null)
                                TLDSet.ContactID = this.TLDApplicationForm.RadiationOncologist.ContactID;
                        }

                        if (Batch.Powder.AuditType == GlobalClass.sAuditTypeRP)
                            TLDSet.SetBeamType = "Cs137";

                        TLDSet.CreatedOn = DateTime.Now;
                        TLDSet.CreatedBy = "Created by user: [" + GlobalClass.User.UserName + "] IP Address: [" + GlobalClass.GetIPAddress() + "] Date: [" + TLDSet.CreatedOn.ToString() + "]";

                        TLDSet.LastUpdate = DateTime.Now;
                        TLDSet.UpdateComment = "Updated by user: [" + GlobalClass.User.UserName + "] IP Address: [" + GlobalClass.GetIPAddress() + "] Date: [" + TLDSet.LastUpdate.ToString() + "]";

                        // Create DataSheet
                        string sErrorMessage = TLDSet.CreateDataSheet();
                        if (sErrorMessage != string.Empty)
                            MessageBox.Show(sErrorMessage, "TLD Data Sheet validation", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        /*
                        string sCheckBeforeSave = TLDSet.CheckBeforeSave();
                        if (sCheckBeforeSave == string.Empty)
                        {
                            if (TLDSet.SaveTLDSet() == 1)
                            {
                                //    this.Operator.PendingTLDSetList.Add(TLDSet);

                                if (TLDSet.TLDDataSheet != null)
                                {
                                    sCheckBeforeSave = TLDSet.TLDDataSheet.CheckBeforeSave();

                                    if (sCheckBeforeSave == string.Empty)
                                    {
                                        if (TLDSet.TLDDataSheet.SaveTLDDataSheet() == 1)
                                        { }
                                    }
                                    else
                                        MessageBox.Show(sCheckBeforeSave, "TLD Data Sheet validation", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                }
                            }
                        }
                        else
                            MessageBox.Show(sCheckBeforeSave, "TLD Set validation", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        */
                    }
                }
            }

            return TLDSet;
        }

        public string SendApplicationFormByEmail()
        {
            string sReturn = string.Empty;

            string sSetNo = string.Empty;
            string sUnitCode = string.Empty;
            string sBeamCode = string.Empty;

            string sEmailFrom = string.Empty;
            string sEmailTo = string.Empty;
            string sEmailCC = string.Empty;
            string sSubject = string.Empty;
            string sBody = string.Empty;
            string sLine = string.Empty;
            string sAttachments = string.Empty;
            string sTemplate = string.Empty;
            string sSaveAs = string.Empty;
            string sSignature = string.Empty;

            //Set the current directory.
            Directory.SetCurrentDirectory(GlobalClass.sApplicationStartupPath);

            if (this.Operator != null)
            {
                if (this.ParticipationType == GlobalClass.sParticipationTypeHospitals)
                {
                    sTemplate = GlobalClass.sApplicationStartupPath + "\\" + GlobalClass.TLDEmailTemplateApplicationFormRTHospital;

                    string sOutputFile = this.ExportApplicationForm();
                    if (File.Exists(sOutputFile))
                        sAttachments = sOutputFile + ";";

                    if (this.CommunicationLanguage == GlobalClass.sCommunicationLanguageEnglish)
                        sAttachments = sAttachments + GlobalClass.sApplicationStartupPath + "\\" + GlobalClass.TLDPrinciplesOfOperation + ";";
                    else if (this.CommunicationLanguage == GlobalClass.sCommunicationLanguageSpanish)
                        sAttachments = sAttachments + GlobalClass.sApplicationStartupPath + "\\" + GlobalClass.TLDPrinciplesOfOperationSpanish + ";";
                    else if (this.CommunicationLanguage == GlobalClass.sCommunicationLanguageRussian)
                        sAttachments = sAttachments + GlobalClass.sApplicationStartupPath + "\\" + GlobalClass.TLDPrinciplesOfOperationRussian + ";";
                }
                else if (this.ParticipationType == GlobalClass.sParticipationTypeSSDL)
                {
                    if (this.AuditType == GlobalClass.sAuditTypeRT)
                        sTemplate = GlobalClass.sApplicationStartupPath + "\\" + GlobalClass.TLDEmailTemplateApplicationFormRTSSDL;
                    else if (this.AuditType == GlobalClass.sAuditTypeRP)
                        sTemplate = GlobalClass.sApplicationStartupPath + "\\" + GlobalClass.TLDEmailTemplateApplicationFormRPSSDL;
                }


                if (File.Exists(sTemplate))
                {
                    foreach (TLDSetClass TLDSet in this.GetTLDSetList())
                    {
                        sSetNo = sSetNo + TLDSet.SetNo + ",";

                        if (TLDSet.Beam != null)
                            sBeamCode = sBeamCode + TLDSet.Beam.BeamCode + ",";
                        if (TLDSet.Unit != null)
                            sUnitCode = sUnitCode + TLDSet.Unit.UnitCode + ","; ;
                    }
                    sSetNo = sSetNo.TrimEnd(',');
                    sBeamCode = sBeamCode.TrimEnd(',');
                    sUnitCode = sUnitCode.TrimEnd(',');

                    try
                    {
                        // Create an instance of StreamReader to read from a file.
                        // The using statement also closes the StreamReader.
                        using (StreamReader sr = new StreamReader(sTemplate, System.Text.Encoding.Default, true))
                        {
                            // Read and display lines from the file until the end of the file is reached.
                            while ((sLine = sr.ReadLine()) != null)
                            {
                                sLine = sLine.Replace("<<OperatorID>>", this.Operator.OperatorID.ToString());
                                sLine = sLine.Replace("<<CCode>>", this.Operator.CCode);
                                sLine = sLine.Replace("<<AuditType>>", this.AuditType);

                                if (this.Batch != null)
                                {
                                    sLine = sLine.Replace("<<BatchNo>>", this.Batch.BatchNo);
                                    sLine = sLine.Replace("<<BatchYear>>", this.Batch.BatchYear);
                                    sLine = sLine.Replace("<<BatchMonth>>", this.Batch.BatchMonth);
                                    if (sLine.Contains("<<BatchWindow>>"))
                                    {
                                        if (this.PackageType == 1)
                                        {
                                            if (this.CommunicationLanguage == GlobalClass.sCommunicationLanguageSpanish)
                                                sLine = sLine.Replace("<<BatchWindow>>", this.Batch.BatchWindowSpanish);
                                            else if (this.CommunicationLanguage == GlobalClass.sCommunicationLanguageRussian)
                                                sLine = sLine.Replace("<<BatchWindow>>", this.Batch.BatchWindowRussian);
                                            else
                                                sLine = sLine.Replace("<<BatchWindow>>", this.Batch.BatchWindow);
                                        }
                                        else if (this.PackageType == 2)
                                        {
                                            if (this.CommunicationLanguage == GlobalClass.sCommunicationLanguageSpanish)
                                                sLine = sLine.Replace("<<BatchWindow>>", "lo antes posible");
                                            else if (this.CommunicationLanguage == GlobalClass.sCommunicationLanguageRussian)
                                                sLine = sLine.Replace("<<BatchWindow>>", "как можно скорее по получении");
                                            else
                                                sLine = sLine.Replace("<<BatchWindow>>", "as soon as possible");
                                        }
                                    }

                                    //sLine = sLine.Replace("<<ApplicationEndDate>>",  GlobalClass.FormatDateTimeValue(this.Batch.ApplicationEndDate));
                                    //sLine = sLine.Replace("<<TLDPackageSendDate>>", GlobalClass.FormatDateTimeValue(this.Batch.TLDPackageSendDate));

                                    sLine = sLine.Replace("<<ApplicationEndDate>>", this.Batch.ApplicationEndDate.ToLongDateString());
                                    sLine = sLine.Replace("<<TLDPackageSendDate>>", this.Batch.TLDPackageSendDate.ToLongDateString());
                                }

                                sLine = sLine.Replace("<<SetNo>>", sSetNo);
                                sLine = sLine.Replace("<<UnitCode>>", sUnitCode);
                                sLine = sLine.Replace("<<BeamCode>>", sBeamCode);

                                if (this.Operator.LastParticipation != null)
                                    sLine = sLine.Replace("<<LastParticipationYear>>", this.Operator.LastParticipation.LastIrradiationYear.ToString());

                                sLine = sLine.Replace("<<LabCode>>", this.Operator.LabCode);
                                sLine = sLine.Replace("<<OperatorName>>", this.Operator.OperatorName);
                                sLine = sLine.Replace("<<Street>>", this.Operator.Street);
                                sLine = sLine.Replace("<<City>>", this.Operator.City);
                                sLine = sLine.Replace("<<POBox>>", this.Operator.POBox);
                                sLine = sLine.Replace("<<State>>", this.Operator.State);
                                sLine = sLine.Replace("<<Country>>", this.Operator.Country);

                                sLine = sLine.Replace("<<InstitutionalEmail>>", this.Operator.InstitutionalEmail);
                                sLine = sLine.Replace("<<InstitutionalTelephone1>>", this.Operator.InstitutionalTelephone1);
                                sLine = sLine.Replace("<<InstitutionalFax>>", this.Operator.InstitutionalFax);

                                sLine = sLine.Replace("<<PackageComment>>", this.PackageComment);
                                sLine = sLine.Replace("<<PackageDescription>>", this.PackageDescription);


                                if (this.TLDApplicationForm != null)
                                {
                                    if (this.TLDApplicationForm.MedicalPhysicist != null)
                                    {
                                        sLine = sLine.Replace("<<ContactTitle>>", this.TLDApplicationForm.MedicalPhysicist.Title.Trim());
                                        sLine = sLine.Replace("<<ContactFamilyName>>", this.TLDApplicationForm.MedicalPhysicist.FamilyName.Trim());
                                        sLine = sLine.Replace("<<ContactFirstName>>", this.TLDApplicationForm.MedicalPhysicist.FirstName.Trim());
                                        sLine = sLine.Replace("<<ContactPosition>>", this.TLDApplicationForm.MedicalPhysicist.Position.Trim());
                                        sLine = sLine.Replace("<<ContactDepartment>>", this.TLDApplicationForm.MedicalPhysicist.Department.Trim());
                                        sLine = sLine.Replace("<<ContactEmail>>", this.TLDApplicationForm.MedicalPhysicist.Email.Trim());
                                        sLine = sLine.Replace("<<ContactTelephone1>>", this.TLDApplicationForm.MedicalPhysicist.Telephone1.Trim());
                                        sLine = sLine.Replace("<<ContactTelephone2>>", this.TLDApplicationForm.MedicalPhysicist.Telephone2.Trim());
                                    }
                                }

                                if (sLine.Contains("<<eMail.Subject>>"))
                                    sSubject = sLine.Replace("<<eMail.Subject>>", "").Trim();
                                else if (sLine.Contains("<<eMail.TemplateObject>>"))
                                { }
                                else if (sLine.Contains("<<eMail.From>>"))
                                    sEmailFrom = sLine.Replace("<<eMail.From>>", string.Empty).Trim();
                                else if (sLine.Contains("<<eMail.To>>"))
                                {
                                    sEmailTo = sLine.Replace("<<eMail.To>>", "").Trim();
                                    if (sEmailTo == "<<ContactEmail>>")
                                    {
                                        sEmailTo = string.Empty;
                                        if (this.TLDApplicationForm != null)
                                            if (this.TLDApplicationForm.MedicalPhysicist != null)
                                                sEmailTo = this.TLDApplicationForm.MedicalPhysicist.Email;
                                    }
                                    else if (sEmailTo == "<<InstitutionalEmail>>")
                                    {
                                        sEmailTo = this.Operator.InstitutionalEmail;
                                    }

                                    sEmailTo = sEmailTo.Replace(" ", "").Replace(",", ";");
                                }
                                else if (sLine.Contains("<<eMail.CC>>"))
                                {
                                    sEmailCC = sLine.Replace("<<eMail.CC>>", "").Trim();
                                    if (sEmailCC == "<<ContactEmail>>")
                                    {
                                        sEmailCC = string.Empty;
                                        if (this.TLDApplicationForm != null)
                                            if (this.TLDApplicationForm.MedicalPhysicist != null)
                                                sEmailCC = this.TLDApplicationForm.MedicalPhysicist.Email;
                                    }
                                    else if (sEmailCC == "<<InstitutionalEmail>>")
                                    {
                                        sEmailCC = this.Operator.InstitutionalEmail;
                                    }
                                    else if (sEmailCC == "<<NationalCoordinatorEmail>>")
                                    {
                                        sEmailCC = sEmailCC.Replace("<<NationalCoordinatorEmail>>", string.Empty).Trim();
                                        List<NationalCoordinatorClass> NationalCoordinatorList = GlobalClass.Manager.GetCountryCoordinatorList(this.CCode, 0);
                                        foreach (NationalCoordinatorClass NationalCoordinator in NationalCoordinatorList)
                                        {
                                            if (NationalCoordinator != null)
                                                if (NationalCoordinator.Email != string.Empty)
                                                    sEmailCC = sEmailCC + NationalCoordinator.Email + ";";
                                        }
                                    }
                                    sEmailCC = sEmailCC.Replace(" ", "").Replace(",", ";");
                                }
                                else if (sLine.Contains("<<eMail.Attachments>>"))
                                {
                                    string[] sAttachmentLineList = sLine.Replace("<<eMail.Attachments>>", "").Trim().Split(';');
                                    foreach (string sAttachmentFile in sAttachmentLineList)
                                    {
                                        if (File.Exists(GlobalClass.sApplicationStartupPath + "\\" + GlobalClass.sApplicationTemplatesFolder + sAttachmentFile.Trim()))
                                            sAttachments = sAttachments + GlobalClass.sApplicationStartupPath + "\\" + GlobalClass.sApplicationTemplatesFolder + sAttachmentFile.Trim() + ";";
                                    }
                                }
                                else if (sLine.Contains("<<eMail.Signature>>"))
                                    sSignature = sLine.Replace("<<eMail.Signature>>", "").Trim();

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

                    sAttachments = sAttachments.TrimEnd(';');
                }
                else
                    MessageBox.Show("Template " + sTemplate + " does not exist.", "Generating e-mail", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            //sSaveAs = GlobalClass.sApplicationStartupPath + @"\Reports\Request" + Operator.CCode + Operator.OperatorID.ToString() + "_" + this._RequestID.ToString() + ".msg";

            // Generate E-mail
            GlobalClass.GenerateEmailNotification(sEmailFrom, sEmailTo, sEmailCC, sSubject, sBody, sAttachments, sSaveAs, sSignature);
            return sReturn;
        }

        public string SendApplicationFormReminderByEmail()
        {
            string sReturn = string.Empty;

            string sSetNo = string.Empty;
            string sUnitCode = string.Empty;
            string sBeamCode = string.Empty;

            string sEmailFrom = string.Empty;
            string sEmailTo = string.Empty;
            string sEmailCC = string.Empty;
            string sSubject = string.Empty;
            string sBody = string.Empty;
            string sLine = string.Empty;
            string sAttachments = string.Empty;
            string sTemplate = string.Empty;
            string sSaveAs = string.Empty;
            string sSignature = string.Empty;

            //Set the current directory.
            Directory.SetCurrentDirectory(GlobalClass.sApplicationStartupPath);

            if (this.Operator != null)
            {
                if (this.AuditType == GlobalClass.sAuditTypeRT)
                {
                    if (this.ParticipationType == GlobalClass.sParticipationTypeSSDL)
                        sTemplate = GlobalClass.sApplicationStartupPath + "\\" + GlobalClass.TLDEmailTemplateApplicationFormRTSSDLReminder;
                    else if (this.ParticipationType == GlobalClass.sParticipationTypeHospitals)
                        sTemplate = GlobalClass.sApplicationStartupPath + "\\" + GlobalClass.TLDEmailTemplateApplicationFormRTHospitalReminder;
                }
                else if (this.AuditType == GlobalClass.sAuditTypeRP)
                {
                    if (this.ParticipationType == GlobalClass.sParticipationTypeSSDL)
                        sTemplate = GlobalClass.sApplicationStartupPath + "\\" + GlobalClass.TLDEmailTemplateApplicationFormRTSSDLReminder;
                }

                if (File.Exists(sTemplate))
                {
                    foreach (TLDSetClass TLDSet in this.GetTLDSetList())
                    {
                        sSetNo = sSetNo + TLDSet.SetNo + ",";

                        if (TLDSet.Beam != null)
                            sBeamCode = sBeamCode + TLDSet.Beam.BeamCode + ",";
                        if (TLDSet.Unit != null)
                            sUnitCode = sUnitCode + TLDSet.Unit.UnitCode + ","; ;
                    }
                    sSetNo = sSetNo.TrimEnd(',');
                    sBeamCode = sBeamCode.TrimEnd(',');
                    sUnitCode = sUnitCode.TrimEnd(',');

                    string sOutputFile = this.ExportApplicationForm();
                    if (File.Exists(sOutputFile))
                        sAttachments = sOutputFile + ";";

                    if (this.CommunicationLanguage == GlobalClass.sCommunicationLanguageEnglish)
                        sAttachments = sAttachments + GlobalClass.sApplicationStartupPath + "\\" + GlobalClass.TLDPrinciplesOfOperation + ";";
                    else if (this.CommunicationLanguage == GlobalClass.sCommunicationLanguageSpanish)
                        sAttachments = sAttachments + GlobalClass.sApplicationStartupPath + "\\" + GlobalClass.TLDPrinciplesOfOperationSpanish + ";";
                    else if (this.CommunicationLanguage == GlobalClass.sCommunicationLanguageRussian)
                        sAttachments = sAttachments + GlobalClass.sApplicationStartupPath + "\\" + GlobalClass.TLDPrinciplesOfOperationRussian + ";";

                    try
                    {
                        // Create an instance of StreamReader to read from a file.
                        // The using statement also closes the StreamReader.
                        using (StreamReader sr = new StreamReader(sTemplate, System.Text.Encoding.Default, true))
                        {
                            // Read and display lines from the file until the end of the file is reached.
                            while ((sLine = sr.ReadLine()) != null)
                            {
                                sLine = sLine.Replace("<<OperatorID>>", this.Operator.OperatorID.ToString());
                                sLine = sLine.Replace("<<CCode>>", this.Operator.CCode);
                                sLine = sLine.Replace("<<AuditType>>", this.AuditType);

                                if (this.Batch != null)
                                {
                                    sLine = sLine.Replace("<<BatchNo>>", this.Batch.BatchNo);
                                    sLine = sLine.Replace("<<BatchYear>>", this.Batch.BatchYear);
                                    sLine = sLine.Replace("<<BatchMonth>>", this.Batch.BatchMonth);
                                    if (sLine.Contains("<<BatchWindow>>"))
                                    {
                                        if (this.PackageType == 1)
                                        {
                                            if (this.CommunicationLanguage == GlobalClass.sCommunicationLanguageSpanish)
                                                sLine = sLine.Replace("<<BatchWindow>>", this.Batch.BatchWindowSpanish);
                                            else if (this.CommunicationLanguage == GlobalClass.sCommunicationLanguageRussian)
                                                sLine = sLine.Replace("<<BatchWindow>>", this.Batch.BatchWindowRussian);
                                            else
                                                sLine = sLine.Replace("<<BatchWindow>>", this.Batch.BatchWindow);
                                        }
                                        else if (this.PackageType == 2)
                                        {
                                            if (this.CommunicationLanguage == GlobalClass.sCommunicationLanguageSpanish)
                                                sLine = sLine.Replace("<<BatchWindow>>", "lo antes posible");
                                            else if (this.CommunicationLanguage == GlobalClass.sCommunicationLanguageRussian)
                                                sLine = sLine.Replace("<<BatchWindow>>", "как можно скорее по получении");
                                            else
                                                sLine = sLine.Replace("<<BatchWindow>>", "as soon as possible");
                                        }
                                    }

                                    //sLine = sLine.Replace("<<ApplicationEndDate>>",  GlobalClass.FormatDateTimeValue(this.Batch.ApplicationEndDate));
                                    //sLine = sLine.Replace("<<TLDPackageSendDate>>", GlobalClass.FormatDateTimeValue(this.Batch.TLDPackageSendDate));

                                    sLine = sLine.Replace("<<ApplicationEndDate>>", this.Batch.ApplicationEndDate.ToLongDateString());
                                    sLine = sLine.Replace("<<TLDPackageSendDate>>", this.Batch.TLDPackageSendDate.ToLongDateString());
                                }

                                sLine = sLine.Replace("<<SetNo>>", sSetNo);
                                sLine = sLine.Replace("<<UnitCode>>", sUnitCode);
                                sLine = sLine.Replace("<<BeamCode>>", sBeamCode);

                                if (this.Operator.LastParticipation != null)
                                    sLine = sLine.Replace("<<LastParticipationYear>>", this.Operator.LastParticipation.LastIrradiationYear.ToString());


                                sLine = sLine.Replace("<<LabCode>>", this.Operator.LabCode);
                                sLine = sLine.Replace("<<OperatorName>>", this.Operator.OperatorName);
                                sLine = sLine.Replace("<<Street>>", this.Operator.Street);
                                sLine = sLine.Replace("<<City>>", this.Operator.City);
                                sLine = sLine.Replace("<<POBox>>", this.Operator.POBox);
                                sLine = sLine.Replace("<<State>>", this.Operator.State);
                                sLine = sLine.Replace("<<Country>>", this.Operator.Country);

                                sLine = sLine.Replace("<<InstitutionalEmail>>", this.Operator.InstitutionalEmail);
                                sLine = sLine.Replace("<<InstitutionalTelephone1>>", this.Operator.InstitutionalTelephone1);
                                sLine = sLine.Replace("<<InstitutionalFax>>", this.Operator.InstitutionalFax);

                                sLine = sLine.Replace("<<PackageComment>>", this.PackageComment);
                                sLine = sLine.Replace("<<PackageDescription>>", this.PackageDescription);


                                if (this.TLDApplicationForm != null)
                                {
                                    if (this.TLDApplicationForm.MedicalPhysicist != null)
                                    {
                                        sLine = sLine.Replace("<<ContactTitle>>", this.TLDApplicationForm.MedicalPhysicist.Title.Trim());
                                        sLine = sLine.Replace("<<ContactFamilyName>>", this.TLDApplicationForm.MedicalPhysicist.FamilyName.Trim());
                                        sLine = sLine.Replace("<<ContactFirstName>>", this.TLDApplicationForm.MedicalPhysicist.FirstName.Trim());
                                        sLine = sLine.Replace("<<ContactPosition>>", this.TLDApplicationForm.MedicalPhysicist.Position.Trim());
                                        sLine = sLine.Replace("<<ContactDepartment>>", this.TLDApplicationForm.MedicalPhysicist.Department.Trim());
                                        sLine = sLine.Replace("<<ContactEmail>>", this.TLDApplicationForm.MedicalPhysicist.Email.Trim());
                                        sLine = sLine.Replace("<<ContactTelephone1>>", this.TLDApplicationForm.MedicalPhysicist.Telephone1.Trim());
                                        sLine = sLine.Replace("<<ContactTelephone2>>", this.TLDApplicationForm.MedicalPhysicist.Telephone2.Trim());
                                    }
                                }

                                if (sLine.Contains("<<eMail.Subject>>"))
                                    sSubject = sLine.Replace("<<eMail.Subject>>", "").Trim();
                                else if (sLine.Contains("<<eMail.TemplateObject>>"))
                                { }
                                else if (sLine.Contains("<<eMail.From>>"))
                                    sEmailFrom = sLine.Replace("<<eMail.From>>", string.Empty).Trim();
                                else if (sLine.Contains("<<eMail.To>>"))
                                {
                                    sEmailTo = sLine.Replace("<<eMail.To>>", "").Trim();
                                    if (sEmailTo == "<<ContactEmail>>")
                                    {
                                        sEmailTo = string.Empty;
                                        if (this.TLDApplicationForm != null)
                                            if (this.TLDApplicationForm.MedicalPhysicist != null)
                                                sEmailTo = this.TLDApplicationForm.MedicalPhysicist.Email;
                                    }
                                    else if (sEmailTo == "<<InstitutionalEmail>>")
                                    {
                                        sEmailTo = this.Operator.InstitutionalEmail;
                                    }

                                    sEmailTo = sEmailTo.Replace(" ", "").Replace(",", ";");
                                }
                                else if (sLine.Contains("<<eMail.CC>>"))
                                {
                                    sEmailCC = sLine.Replace("<<eMail.CC>>", "").Trim();
                                    if (sEmailCC == "<<ContactEmail>>")
                                    {
                                        sEmailCC = string.Empty;
                                        if (this.TLDApplicationForm != null)
                                            if (this.TLDApplicationForm.MedicalPhysicist != null)
                                                sEmailCC = this.TLDApplicationForm.MedicalPhysicist.Email;
                                    }
                                    else if (sEmailCC == "<<InstitutionalEmail>>")
                                    {
                                        sEmailCC = this.Operator.InstitutionalEmail;
                                    }
                                    else if (sEmailCC == "<<NationalCoordinatorEmail>>")
                                    {
                                        sEmailCC = sEmailCC.Replace("<<NationalCoordinatorEmail>>", string.Empty).Trim();
                                        List<NationalCoordinatorClass> NationalCoordinatorList = GlobalClass.Manager.GetCountryCoordinatorList(this.CCode, 0);
                                        foreach (NationalCoordinatorClass NationalCoordinator in NationalCoordinatorList)
                                        {
                                            if (NationalCoordinator != null)
                                                if (NationalCoordinator.Email != string.Empty)
                                                    sEmailCC = sEmailCC + NationalCoordinator.Email + ";";
                                        }
                                    }
                                    sEmailCC = sEmailCC.Replace(" ", "").Replace(",", ";");
                                }
                                else if (sLine.Contains("<<eMail.Attachments>>"))
                                {
                                    string[] sAttachmentLineList = sLine.Replace("<<eMail.Attachments>>", "").Trim().Split(';');
                                    foreach (string sAttachmentFile in sAttachmentLineList)
                                    {
                                        if (File.Exists(GlobalClass.sApplicationStartupPath + "\\" + GlobalClass.sApplicationTemplatesFolder + sAttachmentFile.Trim()))
                                            sAttachments = sAttachments + GlobalClass.sApplicationStartupPath + "\\" + GlobalClass.sApplicationTemplatesFolder + sAttachmentFile.Trim() + ";";
                                    }
                                }
                                else if (sLine.Contains("<<eMail.Signature>>"))
                                    sSignature = sLine.Replace("<<eMail.Signature>>", "").Trim();

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

                    sAttachments = sAttachments.TrimEnd(';');
                }
                else
                    MessageBox.Show("Template " + sTemplate + " does not exist.", "Generating e-mail", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            //sSaveAs = GlobalClass.sApplicationStartupPath + @"\Reports\Request" + Operator.CCode + Operator.OperatorID.ToString() + "_" + this._RequestID.ToString() + ".msg";

            // Generate E-mail
            GlobalClass.GenerateEmailNotification(sEmailFrom, sEmailTo, sEmailCC, sSubject, sBody, sAttachments, sSaveAs, sSignature);
            return sReturn;
        }

        public string SendDataSheetByEmail()
        {
            string sReturn = string.Empty;

            string sSetNo = string.Empty;
            string sUnitCode = string.Empty;
            string sBeamCode = string.Empty;

            string sEmailFrom = string.Empty;
            string sEmailTo = string.Empty;
            string sEmailCC = string.Empty;
            string sSubject = string.Empty;
            string sBody = string.Empty;
            string sLine = string.Empty;
            string sAttachments = string.Empty;
            string sTemplate = string.Empty;
            string sSaveAs = string.Empty;
            string sSignature = string.Empty;

            //string sTemplateType = string.Empty;
            string sReturnDataSheetsOption = string.Empty;
           
            string sCoordinatorName = string.Empty;
            string sCoordinatorEmail = string.Empty;

            //Set the current directory.
            Directory.SetCurrentDirectory(GlobalClass.sApplicationStartupPath);

            List<NationalCoordinatorClass> NationalCoordinatorList = GlobalClass.Manager.GetCountryCoordinatorList(this.CCode, 0);
            foreach (NationalCoordinatorClass NationalCoordinator in NationalCoordinatorList)
            {
                if (NationalCoordinator.Email != string.Empty)
                {
                    sCoordinatorName = sCoordinatorName + NationalCoordinator.CoordinatorName + ", ";
                    sCoordinatorEmail = sCoordinatorEmail + NationalCoordinator.Email.Replace(",", ";").Trim() + ";";
                }
            }

            sCoordinatorName = sCoordinatorName.Trim().TrimEnd(',').Trim();
            sCoordinatorEmail = sCoordinatorEmail.Replace(" ", "").Replace(",", ";").TrimEnd(';');


            // In both letters : to co-ordinators and to hospitals there should be instructions how to return TLDs and how to return data sheets marked A, B, C below.
            // These are options to select by Sharon (1 per letter) depending on specific arrangement with the co-ordinator. 
            if (this.ReturnDataSheetsOption == "OptionA")
                sReturnDataSheetsOption = GlobalClass.Dictionary.GetExternalVariable("ReturnDataSheetsOptionA");
            else if (this.ReturnDataSheetsOption == "OptionB")
                sReturnDataSheetsOption = GlobalClass.Dictionary.GetExternalVariable("ReturnDataSheetsOptionB");
            else if (this.ReturnDataSheetsOption == "OptionC")
                sReturnDataSheetsOption = GlobalClass.Dictionary.GetExternalVariable("ReturnDataSheetsOptionC");           

            if (this.Operator != null)
            {
                //sEmail = this.ContactEmail.Trim();
                //sCCEmail = this.EmailInstitutional.Trim();

                if (this.AuditType == GlobalClass.sAuditTypeRT)
                {
                    if (this.ParticipationType == GlobalClass.sParticipationTypeHospitals)
                    {
                        if (this.SendToDataSheetsOption == "DirectHospital")
                        {
                            if (this.PackageType == 1)
                                sTemplate = GlobalClass.sApplicationStartupPath + "\\" + GlobalClass.TLDEmailTemplateSendDataSheetRTHospital;
                            else if (this.PackageType == 2)
                                sTemplate = GlobalClass.sApplicationStartupPath + "\\" + GlobalClass.TLDEmailTemplateSendDataSheetRTHospitalFollowUp;
                        }
                        else if (this.SendToDataSheetsOption == "HospitalCCNationalCoordinator")
                        {
                            if (this.PackageType == 1)
                                sTemplate = GlobalClass.sApplicationStartupPath + "\\" + GlobalClass.TLDEmailTemplateSendDataSheetRTHospitalCountryCoordinatorCC;
                            else if (this.PackageType == 2)
                                sTemplate = GlobalClass.sApplicationStartupPath + "\\" + GlobalClass.TLDEmailTemplateSendDataSheetRTHospitalCountryCoordinatorCCFollowUp;
                        }
                        else if (this.SendToDataSheetsOption == "NationalCoordinator")
                        {
                            if (this.PackageType == 1)
                                sTemplate = GlobalClass.sApplicationStartupPath + "\\" + GlobalClass.TLDEmailTemplateSendDataSheetRTHospitalCountryCoordinator;
                            else if (this.PackageType == 2)
                                sTemplate = GlobalClass.sApplicationStartupPath + "\\" + GlobalClass.TLDEmailTemplateSendDataSheetRTHospitalCountryCoordinatorFollowUp;
                        }
                        else if (this.SendToDataSheetsOption == "NationalCoordinatorPAHO")
                        {
                            if (this.PackageType == 1)
                                sTemplate = GlobalClass.sApplicationStartupPath + "\\" + GlobalClass.TLDEmailTemplateSendDataSheetRTHospitalPAHOEnglish;
                            else if (this.PackageType == 2)
                                sTemplate = GlobalClass.sApplicationStartupPath + "\\" + GlobalClass.TLDEmailTemplateSendDataSheetRTHospitalPAHOEnglishFollowUp;

                            if (this.CommunicationLanguage == GlobalClass.sCommunicationLanguageSpanish)
                            {
                                if (this.PackageType == 1)
                                    sTemplate = GlobalClass.sApplicationStartupPath + "\\" + GlobalClass.TLDEmailTemplateSendDataSheetRTHospitalPAHOSpanish;
                                else if (this.PackageType == 2)
                                    sTemplate = GlobalClass.sApplicationStartupPath + "\\" + GlobalClass.TLDEmailTemplateSendDataSheetRTHospitalPAHOSpanishFollowUp;
                            }
                        }
                    }
                    else if (this.ParticipationType == GlobalClass.sParticipationTypeSSDL)
                    {
                        if (this.PackageType == 1)
                            sTemplate = GlobalClass.sApplicationStartupPath + "\\" + GlobalClass.TLDEmailTemplateSendDataSheetRTSSDL;
                        else if (this.PackageType == 2)
                            sTemplate = GlobalClass.sApplicationStartupPath + "\\" + GlobalClass.TLDEmailTemplateSendDataSheetRTSSDLFollowUp;
                    }
                    else if (this.ParticipationType == GlobalClass.sParticipationTypeReference)
                        sTemplate = GlobalClass.sApplicationStartupPath + "\\" + GlobalClass.TLDEmailTemplateSendDataSheetRTReference;
                    else if (this.ParticipationType == GlobalClass.sParticipationTypePrimary)
                        sTemplate = GlobalClass.sApplicationStartupPath + "\\" + GlobalClass.TLDEmailTemplateSendDataSheetRTPrimary;
                }
                else if (this.AuditType == GlobalClass.sAuditTypeRP)
                {
                    if (this.ParticipationType == GlobalClass.sParticipationTypeSSDL)
                        sTemplate = GlobalClass.sApplicationStartupPath + "\\" + GlobalClass.TLDEmailTemplateSendDataSheetRPSSDL;
                    else if (this.ParticipationType == GlobalClass.sParticipationTypeReference)
                        sTemplate = GlobalClass.sApplicationStartupPath + "\\" + GlobalClass.TLDEmailTemplateSendDataSheetRPReference;
                    else if (this.ParticipationType == GlobalClass.sParticipationTypePrimary)
                        sTemplate = GlobalClass.sApplicationStartupPath + "\\" + GlobalClass.TLDEmailTemplateSendDataSheetRPPrimary;
                }

                if (File.Exists(sTemplate))
                {
                    string sBeamTypeList = string.Empty;
                    foreach (TLDSetClass TLDSet in this.GetTLDSetList())
                    {
                        sSetNo = sSetNo + TLDSet.SetNo + ",";

                        if (TLDSet.Beam != null)
                            sBeamCode = sBeamCode + TLDSet.Beam.BeamCode + ",";
                        if (TLDSet.Unit != null)
                            sUnitCode = sUnitCode + TLDSet.Unit.UnitCode + ",";

                        if (sBeamTypeList.Contains(TLDSet.SetBeamType) == false)
                            sBeamTypeList = sBeamTypeList + TLDSet.SetBeamType + ",";

                        if (TLDSet.TLDDataSheet != null)
                        {
                            // Add TLD DataSheet
                            string sOutputFile = TLDSet.TLDDataSheet.ExportTLDDataSheet("PrePopulated");
                            if (File.Exists(sOutputFile))
                                sAttachments = sAttachments + sOutputFile + ";";                                
                        }
                    }
                    sSetNo = sSetNo.TrimEnd(',');
                    sBeamCode = sBeamCode.TrimEnd(',');
                    sUnitCode = sUnitCode.TrimEnd(',');
                    sBeamTypeList = sBeamTypeList.TrimEnd(',');

                    // Add TLD InstructionSheet
                    if (this.Batch != null)
                    {
                        string sTLDInstructionSheet = this.Batch.PrepopulateTLDInstructionSheet(this.ParticipationType, this.CommunicationLanguage, this.PackageType, sBeamTypeList);
                        if (sTLDInstructionSheet != string.Empty)
                            //if (File.Exists(sTLDInstructionSheet))
                                sAttachments = sAttachments + sTLDInstructionSheet + ";";                            
                    }
       
                    try
                    {
                        // Create an instance of StreamReader to read from a file.
                        // The using statement also closes the StreamReader.
                        using (StreamReader sr = new StreamReader(sTemplate, System.Text.Encoding.Default, true))
                        {
                            // Read and display lines from the file until the end of the file is reached.
                            while ((sLine = sr.ReadLine()) != null)
                            {
                                if (sLine.Contains("<<ReturnDataSheetsOption>>"))
                                    sLine = sLine.Replace("<<ReturnDataSheetsOption>>", sReturnDataSheetsOption);

                                sLine = sLine.Replace("<<OperatorID>>", this.Operator.OperatorID.ToString());
                                sLine = sLine.Replace("<<CCode>>", this.Operator.CCode);
                                sLine = sLine.Replace("<<AuditType>>", this.AuditType);

                                sLine = sLine.Replace("<<CoordinatorName>>", sCoordinatorName);
                                sLine = sLine.Replace("<<CoordinatorEmail>>", sCoordinatorEmail);

                                if (this.Batch != null)
                                {
                                    sLine = sLine.Replace("<<BatchNo>>", this.Batch.BatchNo);
                                    sLine = sLine.Replace("<<BatchYear>>", this.Batch.BatchYear);
                                    sLine = sLine.Replace("<<BatchMonth>>", this.Batch.BatchMonth);

                                    if (sLine.Contains("<<BatchWindow>>"))
                                    {
                                        if (this.PackageType == 1)
                                        {
                                            if (this.CommunicationLanguage == GlobalClass.sCommunicationLanguageSpanish)
                                                sLine = sLine.Replace("<<BatchWindow>>", this.Batch.BatchWindowSpanish);
                                            else if (this.CommunicationLanguage == GlobalClass.sCommunicationLanguageRussian)
                                                sLine = sLine.Replace("<<BatchWindow>>", this.Batch.BatchWindowRussian);
                                            else
                                                sLine = sLine.Replace("<<BatchWindow>>", this.Batch.BatchWindow);
                                        }
                                        else if (this.PackageType == 2)
                                        {
                                            if (this.CommunicationLanguage == GlobalClass.sCommunicationLanguageSpanish)
                                                sLine = sLine.Replace("<<BatchWindow>>", "lo antes posible");
                                            else if (this.CommunicationLanguage == GlobalClass.sCommunicationLanguageRussian)
                                                sLine = sLine.Replace("<<BatchWindow>>", "как можно скорее по получении");
                                            else
                                                sLine = sLine.Replace("<<BatchWindow>>", "as soon as possible");
                                        }

                                    }

                                    sLine = sLine.Replace("<<ApplicationEndDate>>", this.Batch.ApplicationEndDate.ToLongDateString());
                                    sLine = sLine.Replace("<<TLDPackageSendDate>>", this.Batch.TLDPackageSendDate.ToLongDateString());
                                }

                                sLine = sLine.Replace("<<SetNo>>", sSetNo);
                                sLine = sLine.Replace("<<SetNos>>", this.SetNos);
                                sLine = sLine.Replace("<<OriginalSetNos>>", this.OriginalSetNos);
                                sLine = sLine.Replace("<<FollowUpSetNos>>", this.FollowUpSetNos);
                                
                                sLine = sLine.Replace("<<UnitCode>>", sUnitCode);
                                sLine = sLine.Replace("<<BeamCode>>", sBeamCode);
                                if (this.Operator.LastParticipation != null)
                                    sLine = sLine.Replace("<<LastParticipationYear>>", this.Operator.LastParticipation.LastIrradiationYear.ToString());

                                sLine = sLine.Replace("<<LabCode>>", this.Operator.LabCode);
                                sLine = sLine.Replace("<<OperatorName>>", this.Operator.OperatorName);
                                sLine = sLine.Replace("<<Street>>", this.Operator.Street);
                                sLine = sLine.Replace("<<City>>", this.Operator.City);
                                sLine = sLine.Replace("<<POBox>>", this.Operator.POBox);
                                sLine = sLine.Replace("<<State>>", this.Operator.State);
                                sLine = sLine.Replace("<<Country>>", this.Operator.Country);

                                sLine = sLine.Replace("<<InstitutionalEmail>>", this.Operator.InstitutionalEmail);
                                sLine = sLine.Replace("<<InstitutionalTelephone1>>", this.Operator.InstitutionalTelephone1);
                                sLine = sLine.Replace("<<InstitutionalFax>>", this.Operator.InstitutionalFax);

                                sLine = sLine.Replace("<<PackageComment>>", this.PackageComment);
                                sLine = sLine.Replace("<<PackageDescription>>", this.PackageDescription);

                                if (this.TLDApplicationForm != null)
                                {
                                    if (this.TLDApplicationForm.MedicalPhysicist != null)
                                    {
                                        sLine = sLine.Replace("<<ContactTitle>>", this.TLDApplicationForm.MedicalPhysicist.Title.Trim());
                                        sLine = sLine.Replace("<<ContactFamilyName>>", this.TLDApplicationForm.MedicalPhysicist.FamilyName.Trim());
                                        sLine = sLine.Replace("<<ContactFirstName>>", this.TLDApplicationForm.MedicalPhysicist.FirstName.Trim());
                                        sLine = sLine.Replace("<<ContactPosition>>", this.TLDApplicationForm.MedicalPhysicist.Position.Trim());
                                        sLine = sLine.Replace("<<ContactDepartment>>", this.TLDApplicationForm.MedicalPhysicist.Department.Trim());
                                        sLine = sLine.Replace("<<ContactEmail>>", this.TLDApplicationForm.MedicalPhysicist.Email.Trim());
                                        sLine = sLine.Replace("<<ContactTelephone1>>", this.TLDApplicationForm.MedicalPhysicist.Telephone1.Trim());
                                        sLine = sLine.Replace("<<ContactTelephone2>>", this.TLDApplicationForm.MedicalPhysicist.Telephone2.Trim());
                                    }
                                }

                                if (sLine.Contains("<<eMail.Subject>>"))
                                    sSubject = sLine.Replace("<<eMail.Subject>>", "").Trim();
                                else if (sLine.Contains("<<eMail.TemplateObject>>"))
                                { }
                                else if (sLine.Contains("<<eMail.From>>"))
                                    sEmailFrom = sLine.Replace("<<eMail.From>>", string.Empty).Trim();
                                else if (sLine.Contains("<<eMail.To>>"))
                                {
                                    sEmailTo = sLine.Replace("<<eMail.To>>", "").Trim();
                                    if (sEmailTo == "<<ContactEmail>>")
                                    {
                                        sEmailTo = string.Empty;
                                        if (this.TLDApplicationForm != null)
                                            if (this.TLDApplicationForm.MedicalPhysicist != null)
                                                sEmailTo = this.TLDApplicationForm.MedicalPhysicist.Email;
                                    }
                                    else if (sEmailTo == "<<InstitutionalEmail>>")
                                    {
                                        sEmailTo = this.Operator.InstitutionalEmail;
                                    }

                                    sEmailTo = sEmailTo.Replace(" ", "").Replace(",", ";");
                                }
                                else if (sLine.Contains("<<eMail.CC>>"))
                                {
                                    sEmailCC = sLine.Replace("<<eMail.CC>>", "").Trim();
                                    if (sEmailCC.Contains("<<ContactEmail>>"))
                                    {
                                        string sContactEmail = string.Empty;
                                        
                                        if (this.TLDApplicationForm != null)
                                            if (this.TLDApplicationForm.MedicalPhysicist != null)
                                                sContactEmail = this.TLDApplicationForm.MedicalPhysicist.Email;

                                        sEmailCC = sEmailCC.Replace("<<ContactEmail>>", sContactEmail).Trim();
                                    }
                                    else if (sEmailCC.Contains("<<InstitutionalEmail>>"))
                                    {
                                        string sInstitutionalEmail = string.Empty;
                                        if (this.Operator.InstitutionalEmail.Trim() != string.Empty)
                                            sInstitutionalEmail = this.Operator.InstitutionalEmail;

                                        sEmailCC = sEmailCC.Replace("<<ContactEmail>>", sInstitutionalEmail).Trim();
                                    }
                                    else if (sEmailCC.Contains("<<NationalCoordinatorEmail>>"))
                                    {
                                        sEmailCC = sEmailCC.Replace("<<NationalCoordinatorEmail>>", sCoordinatorEmail).Trim();
                                    }
                                    sEmailCC = sEmailCC.Replace(" ", "").Replace(",", ";");
                                }
                                else if (sLine.Contains("<<eMail.Attachments>>"))
                                {
                                    string[] sAttachmentLineList = sLine.Replace("<<eMail.Attachments>>", "").Trim().Split(';');
                                    foreach (string sAttachmentFile in sAttachmentLineList)
                                    {
                                        if (File.Exists(GlobalClass.sApplicationStartupPath + "\\" + GlobalClass.sApplicationTemplatesFolder + sAttachmentFile.Trim()))
                                            sAttachments = sAttachments + GlobalClass.sApplicationStartupPath + "\\" + GlobalClass.sApplicationTemplatesFolder + sAttachmentFile.Trim() + ";";
                                    }
                                }
                                else if (sLine.Contains("<<eMail.Signature>>"))
                                    sSignature = sLine.Replace("<<eMail.Signature>>", "").Trim();

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

                    sAttachments = sAttachments.TrimEnd(';');
                }
                else
                    MessageBox.Show("Template " + sTemplate + " does not exist.", "Generating e-mail", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            //sSaveAs = GlobalClass.sApplicationStartupPath + @"\Reports\Request" + Operator.CCode + Operator.OperatorID.ToString() + "_" + this._RequestID.ToString() + ".msg";

            // Generate E-mail
            GlobalClass.GenerateEmailNotification(sEmailFrom, sEmailTo, sEmailCC, sSubject, sBody, sAttachments, sSaveAs, sSignature);
            
            return sReturn;
        }

        public string SendDataSheetReminderByEmail()
        {
            string sReturn = string.Empty;

            string sSetNo = string.Empty;
            string sUnitCode = string.Empty;
            string sBeamCode = string.Empty;

            string sEmailFrom = string.Empty;
            string sEmailTo = string.Empty;
            string sEmailCC = string.Empty;
            string sSubject = string.Empty;
            string sBody = string.Empty;
            string sLine = string.Empty;
            string sAttachments = string.Empty;
            string sTemplate = string.Empty;
            string sSaveAs = string.Empty;
            string sSignature = string.Empty;

            //Set the current directory.
            Directory.SetCurrentDirectory(GlobalClass.sApplicationStartupPath);

            if (this.Operator != null)
            {
                //sEmail = this.ContactEmail.Trim();
                //sCCEmail = this.EmailInstitutional.Trim();

                if (this.AuditType == GlobalClass.sAuditTypeRT)
                {
                    if (this.ParticipationType == GlobalClass.sParticipationTypeSSDL)
                        sTemplate = GlobalClass.sApplicationStartupPath + "\\" + GlobalClass.TLDEmailTemplateSendDataSheetRTReminder;
                    else if (this.ParticipationType == GlobalClass.sParticipationTypeHospitals)
                        sTemplate = GlobalClass.sApplicationStartupPath + "\\" + GlobalClass.TLDEmailTemplateSendDataSheetRTReminder;
                    else if (this.ParticipationType == GlobalClass.sParticipationTypeReference)
                        sTemplate = GlobalClass.sApplicationStartupPath + "\\" + GlobalClass.TLDEmailTemplateSendDataSheetRTReminder;
                    else if (this.ParticipationType == GlobalClass.sParticipationTypePrimary)
                        sTemplate = GlobalClass.sApplicationStartupPath + "\\" + GlobalClass.TLDEmailTemplateSendDataSheetRTReminder;
                }
                else if (this.AuditType == GlobalClass.sAuditTypeRP)
                {
                    if (this.ParticipationType == GlobalClass.sParticipationTypeSSDL)
                        sTemplate = GlobalClass.sApplicationStartupPath + "\\" + GlobalClass.TLDEmailTemplateSendDataSheetRPReminder;
                    else if (this.ParticipationType == GlobalClass.sParticipationTypeReference)
                        sTemplate = GlobalClass.sApplicationStartupPath + "\\" + GlobalClass.TLDEmailTemplateSendDataSheetRPReminder;
                    else if (this.ParticipationType == GlobalClass.sParticipationTypePrimary)
                        sTemplate = GlobalClass.sApplicationStartupPath + "\\" + GlobalClass.TLDEmailTemplateSendDataSheetRPReminder;
                }


                if (File.Exists(sTemplate))
                {
                    foreach (TLDSetClass TLDSet in this.GetTLDSetList())
                    {
                        if (TLDSet.Beam != null)
                            sBeamCode = sBeamCode + TLDSet.Beam.BeamCode + ",";
                        else
                        {
                            sSetNo = sSetNo + TLDSet.SetNo + ","; // Only missing Sets

                            if (TLDSet.TLDDataSheet != null)
                            {
                                // Add TLD DataSheet
                                string sOutputFile = TLDSet.TLDDataSheet.ExportTLDDataSheet("PrePopulated");
                                if (File.Exists(sOutputFile))
                                    sAttachments = sAttachments + sOutputFile + ";";
                            }
                        }


                        if (TLDSet.Unit != null)
                            sUnitCode = sUnitCode + TLDSet.Unit.UnitCode + ",";
                    }
                    sSetNo = sSetNo.TrimEnd(',');
                    if (sSetNo == string.Empty)
                        sSetNo = "[None]";

                    sBeamCode = sBeamCode.TrimEnd(',');
                    sUnitCode = sUnitCode.TrimEnd(',');

                    try
                    {
                        // Create an instance of StreamReader to read from a file.
                        // The using statement also closes the StreamReader.
                        using (StreamReader sr = new StreamReader(sTemplate, System.Text.Encoding.Default, true))
                        {
                            // Read and display lines from the file until the end of the file is reached.
                            while ((sLine = sr.ReadLine()) != null)
                            {
                                sLine = sLine.Replace("<<OperatorID>>", this.Operator.OperatorID.ToString());
                                sLine = sLine.Replace("<<CCode>>", this.Operator.CCode);
                                sLine = sLine.Replace("<<AuditType>>", this.AuditType);

                                if (this.Batch != null)
                                {
                                    sLine = sLine.Replace("<<BatchNo>>", this.Batch.BatchNo);
                                    sLine = sLine.Replace("<<BatchYear>>", this.Batch.BatchYear);
                                    sLine = sLine.Replace("<<BatchMonth>>", this.Batch.BatchMonth);

                                    if (sLine.Contains("<<BatchWindow>>"))
                                    {
                                        if (this.PackageType == 1)
                                        {
                                            if (this.CommunicationLanguage == GlobalClass.sCommunicationLanguageSpanish)
                                                sLine = sLine.Replace("<<BatchWindow>>", this.Batch.BatchWindowSpanish);
                                            else if (this.CommunicationLanguage == GlobalClass.sCommunicationLanguageRussian)
                                                sLine = sLine.Replace("<<BatchWindow>>", this.Batch.BatchWindowRussian);
                                            else
                                                sLine = sLine.Replace("<<BatchWindow>>", this.Batch.BatchWindow);
                                        }
                                        else if (this.PackageType == 2)
                                        {
                                            if (this.CommunicationLanguage == GlobalClass.sCommunicationLanguageSpanish)
                                                sLine = sLine.Replace("<<BatchWindow>>", "lo antes posible");
                                            else if (this.CommunicationLanguage == GlobalClass.sCommunicationLanguageRussian)
                                                sLine = sLine.Replace("<<BatchWindow>>", "как можно скорее по получении");
                                            else
                                                sLine = sLine.Replace("<<BatchWindow>>", "as soon as possible");
                                        }
                                    }
                                }


                                if (sLine.Contains("<<SetNo>>"))
                                    sLine = sLine.Replace("<<SetNo>>", sSetNo);
                                if (sLine.Contains("<<SetNos>>"))
                                    sLine = sLine.Replace("<<SetNos>>", this.SetNos);

                                sLine = sLine.Replace("<<UnitCode>>", sUnitCode);
                                sLine = sLine.Replace("<<BeamCode>>", sBeamCode);

                                if (this.Operator.LastParticipation != null)
                                    sLine = sLine.Replace("<<LastParticipationYear>>", this.Operator.LastParticipation.LastIrradiationYear.ToString());

                                sLine = sLine.Replace("<<LabCode>>", this.Operator.LabCode);
                                sLine = sLine.Replace("<<OperatorName>>", this.Operator.OperatorName);
                                sLine = sLine.Replace("<<Street>>", this.Operator.Street);
                                sLine = sLine.Replace("<<City>>", this.Operator.City);
                                sLine = sLine.Replace("<<POBox>>", this.Operator.POBox);
                                sLine = sLine.Replace("<<State>>", this.Operator.State);
                                sLine = sLine.Replace("<<Country>>", this.Operator.Country);

                                sLine = sLine.Replace("<<InstitutionalEmail>>", this.Operator.InstitutionalEmail);
                                sLine = sLine.Replace("<<InstitutionalTelephone1>>", this.Operator.InstitutionalTelephone1);
                                sLine = sLine.Replace("<<InstitutionalFax>>", this.Operator.InstitutionalFax);

                                sLine = sLine.Replace("<<PackageComment>>", this.PackageComment);
                                sLine = sLine.Replace("<<PackageDescription>>", this.PackageDescription);

                                if (this.TLDApplicationForm != null)
                                {
                                    if (this.TLDApplicationForm.MedicalPhysicist != null)
                                    {
                                        sLine = sLine.Replace("<<ContactTitle>>", this.TLDApplicationForm.MedicalPhysicist.Title.Trim());
                                        sLine = sLine.Replace("<<ContactFamilyName>>", this.TLDApplicationForm.MedicalPhysicist.FamilyName.Trim());
                                        sLine = sLine.Replace("<<ContactFirstName>>", this.TLDApplicationForm.MedicalPhysicist.FirstName.Trim());
                                        sLine = sLine.Replace("<<ContactPosition>>", this.TLDApplicationForm.MedicalPhysicist.Position.Trim());
                                        sLine = sLine.Replace("<<ContactDepartment>>", this.TLDApplicationForm.MedicalPhysicist.Department.Trim());
                                        sLine = sLine.Replace("<<ContactEmail>>", this.TLDApplicationForm.MedicalPhysicist.Email.Trim());
                                        sLine = sLine.Replace("<<ContactTelephone1>>", this.TLDApplicationForm.MedicalPhysicist.Telephone1.Trim());
                                        sLine = sLine.Replace("<<ContactTelephone2>>", this.TLDApplicationForm.MedicalPhysicist.Telephone2.Trim());
                                    }
                                }

                                if (sLine.Contains("<<eMail.Subject>>"))
                                    sSubject = sLine.Replace("<<eMail.Subject>>", "").Trim();
                                else if (sLine.Contains("<<eMail.TemplateObject>>"))
                                { }
                                else if (sLine.Contains("<<eMail.From>>"))
                                    sEmailFrom = sLine.Replace("<<eMail.From>>", string.Empty).Trim();

                                else if (sLine.Contains("<<eMail.To>>"))
                                {
                                    sEmailTo = sLine.Replace("<<eMail.To>>", "").Trim();
                                    if (sEmailTo == "<<ContactEmail>>")
                                    {
                                        sEmailTo = string.Empty;
                                        if (this.TLDApplicationForm != null)
                                            if (this.TLDApplicationForm.MedicalPhysicist != null)
                                                sEmailTo = this.TLDApplicationForm.MedicalPhysicist.Email;
                                    }
                                    else if (sEmailTo == "<<InstitutionalEmail>>")
                                    {
                                        sEmailTo = this.Operator.InstitutionalEmail;
                                    }

                                    sEmailTo = sEmailTo.Replace(" ", "").Replace(",", ";");
                                }
                                else if (sLine.Contains("<<eMail.CC>>"))
                                {
                                    sEmailCC = sLine.Replace("<<eMail.CC>>", string.Empty).Trim();
                                    if (sEmailCC == "<<ContactEmail>>")
                                    {
                                        sEmailCC = string.Empty;
                                        if (this.TLDApplicationForm != null)
                                            if (this.TLDApplicationForm.MedicalPhysicist != null)
                                                sEmailCC = this.TLDApplicationForm.MedicalPhysicist.Email;
                                    }
                                    else if (sEmailCC == "<<InstitutionalEmail>>")
                                    {
                                        sEmailCC = this.Operator.InstitutionalEmail;
                                    }
                                    else if (sEmailCC == "<<NationalCoordinatorEmail>>")
                                    {
                                        sEmailCC = sEmailCC.Replace("<<NationalCoordinatorEmail>>", string.Empty).Trim();
                                        List<NationalCoordinatorClass> NationalCoordinatorList = GlobalClass.Manager.GetCountryCoordinatorList(this.CCode, 0);
                                        foreach (NationalCoordinatorClass NationalCoordinator in NationalCoordinatorList)
                                        {
                                            if (NationalCoordinator != null)
                                                if (NationalCoordinator.Email != string.Empty)
                                                    sEmailCC = sEmailCC + NationalCoordinator.Email + ";";
                                        }
                                    }
                                    sEmailCC = sEmailCC.Replace(" ", "").Replace(",", ";");
                                }
                                else if (sLine.Contains("<<eMail.Attachments>>"))
                                {
                                    string[] sAttachmentLineList = sLine.Replace("<<eMail.Attachments>>", "").Trim().Split(';');
                                    foreach (string sAttachmentFile in sAttachmentLineList)
                                    {
                                        if (File.Exists(GlobalClass.sApplicationStartupPath + "\\" + GlobalClass.sApplicationTemplatesFolder + sAttachmentFile.Trim()))
                                            sAttachments = sAttachments + GlobalClass.sApplicationStartupPath + "\\" + GlobalClass.sApplicationTemplatesFolder + sAttachmentFile.Trim() + ";";
                                    }
                                }
                                else if (sLine.Contains("<<eMail.Signature>>"))
                                    sSignature = sLine.Replace("<<eMail.Signature>>", "").Trim();

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

                    sAttachments = sAttachments.TrimEnd(';');
                }
                else
                    MessageBox.Show("Template " + sTemplate + " does not exist.", "Generating e-mail", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            //sSaveAs = GlobalClass.sApplicationStartupPath + @"\Reports\Request" + Operator.CCode + Operator.OperatorID.ToString() + "_" + this._RequestID.ToString() + ".msg";

            // Generate E-mail
            GlobalClass.GenerateEmailNotification(sEmailFrom, sEmailTo, sEmailCC, sSubject, sBody, sAttachments, sSaveAs, sSignature);
            return sReturn;
        }

        public string SendCertificateByEmail()
        {
            string sReturn = string.Empty;

            string sSetNo = string.Empty;
            string sUnitCode = string.Empty;
            string sBeamCode = string.Empty;

            string sTemplate = string.Empty;
            string sSubject = string.Empty;
            string sBody = string.Empty;
            string sLine = string.Empty;
            string sAttachments = string.Empty;
            string sSaveAs = string.Empty;
            string sEmailFrom = string.Empty;
            string sEmailTo = string.Empty;
            string sEmailCC = string.Empty;
            string sSignature = string.Empty;

            Directory.SetCurrentDirectory(GlobalClass.sApplicationStartupPath);

            if (this.Operator != null)
            {
                if (this.PackageType == 1)
                {
                    if (this._AuditType == GlobalClass.sAuditTypeRT)
                    {
                        if (this.ParticipationType == GlobalClass.sParticipationTypeSSDL)
                            sTemplate = GlobalClass.sApplicationStartupPath + "\\" + GlobalClass.TLDEmailTemplateDispatchCertificateRTSSDL;
                        else if (this.ParticipationType == GlobalClass.sParticipationTypeHospitals)
                            sTemplate = GlobalClass.sApplicationStartupPath + "\\" + GlobalClass.TLDEmailTemplateDispatchCertificateRTHospital;
                        else if (this.ParticipationType == GlobalClass.sParticipationTypeReference)
                            sTemplate = GlobalClass.sApplicationStartupPath + "\\" + GlobalClass.TLDEmailTemplateDispatchCertificateRTPrimary;
                        else if (this.ParticipationType == GlobalClass.sParticipationTypePrimary)
                            sTemplate = GlobalClass.sApplicationStartupPath + "\\" + GlobalClass.TLDEmailTemplateDispatchCertificateRTPrimary;
                    }
                    else if (this._AuditType == GlobalClass.sAuditTypeRP)
                    {
                        if (this.ParticipationType == GlobalClass.sParticipationTypeSSDL)
                            sTemplate = GlobalClass.sApplicationStartupPath + "\\" + GlobalClass.TLDEmailTemplateDispatchCertificateRPSSDL;
                        else if (this.ParticipationType == GlobalClass.sParticipationTypeReference)
                            sTemplate = GlobalClass.sApplicationStartupPath + "\\" + GlobalClass.TLDEmailTemplateDispatchCertificateRPPrimary;
                        else if (this.ParticipationType == GlobalClass.sParticipationTypePrimary)
                            sTemplate = GlobalClass.sApplicationStartupPath + "\\" + GlobalClass.TLDEmailTemplateDispatchCertificateRPPrimary;
                    }
                }
                else if (this.PackageType == 2)
                {
                    if (this._AuditType == GlobalClass.sAuditTypeRT)
                    {
                        if (this.ParticipationType == GlobalClass.sParticipationTypeSSDL)
                            sTemplate = GlobalClass.sApplicationStartupPath + "\\" + GlobalClass.TLDEmailTemplateDispatchCertificateRTSSDLFollowUp;
                        else if (this.ParticipationType == GlobalClass.sParticipationTypeHospitals)
                            sTemplate = GlobalClass.sApplicationStartupPath + "\\" + GlobalClass.TLDEmailTemplateDispatchCertificateRTHospitalFollowUp;
                    }
                    else if (this._AuditType == GlobalClass.sAuditTypeRP)
                    {
                        if (this.ParticipationType == GlobalClass.sParticipationTypeSSDL)
                            sTemplate = GlobalClass.sApplicationStartupPath + "\\" + GlobalClass.TLDEmailTemplateDispatchCertificateRPSSDLFollowUp;
                    }
                }

                if (File.Exists(sTemplate))
                {
                    string sFollowUpNeeded = string.Empty;
                    int iCountCertificates = 0;
                    int iCountGoodResults = 0;


                    foreach (TLDSetClass TLDSet in this.GetTLDSetList())
                    {
                        if (TLDSet.Certificate != null)
                        {
                            iCountCertificates = iCountCertificates + 1;
                            sSetNo = sSetNo + TLDSet.SetNo + ",";

                            if (TLDSet.Beam != null)
                                sBeamCode = sBeamCode + TLDSet.Beam.BeamCode + ",";
                            if (TLDSet.Unit != null)
                                sUnitCode = sUnitCode + TLDSet.Unit.UnitCode + ",";

                            if (TLDSet.Certificate.DRatio > 0.0)
                                if (TLDSet.Certificate.DRatioInRange)
                                    iCountGoodResults = iCountGoodResults + 1;
                        }
                    }

                    sSetNo = sSetNo.TrimEnd(',');
                    sBeamCode = sBeamCode.TrimEnd(',');
                    sUnitCode = sUnitCode.TrimEnd(',');
                    if (iCountCertificates > iCountGoodResults)
                        sFollowUpNeeded = "Follow-up needed";

                    try
                    {
                        // Create an instance of StreamReader to read from a file.
                        // The using statement also closes the StreamReader.
                        using (StreamReader sr = new StreamReader(sTemplate, System.Text.Encoding.Default, true))
                        {
                            // Read and display lines from the file until the end of 
                            // the file is reached.
                            while ((sLine = sr.ReadLine()) != null)
                            {
                                sLine = sLine.Replace("<<OperatorID>>", this.Operator.OperatorID.ToString());
                                sLine = sLine.Replace("<<CCode>>", this.Operator.CCode);
                                sLine = sLine.Replace("<<AuditType>>", this.AuditType);

                                if (this.Batch != null)
                                {
                                    sLine = sLine.Replace("<<BatchNo>>", this.Batch.BatchNo);
                                    sLine = sLine.Replace("<<BatchYear>>", this.Batch.BatchYear);
                                    sLine = sLine.Replace("<<BatchMonth>>", this.Batch.BatchMonth);

                                    if (sLine.Contains("<<BatchWindow>>"))
                                    {
                                        if (this.PackageType == 1)
                                        {
                                            if (this.CommunicationLanguage == GlobalClass.sCommunicationLanguageSpanish)
                                                sLine = sLine.Replace("<<BatchWindow>>", this.Batch.BatchWindowSpanish);
                                            else if (this.CommunicationLanguage == GlobalClass.sCommunicationLanguageRussian)
                                                sLine = sLine.Replace("<<BatchWindow>>", this.Batch.BatchWindowRussian);
                                            else
                                                sLine = sLine.Replace("<<BatchWindow>>", this.Batch.BatchWindow);
                                        }
                                        else if (this.PackageType == 2)
                                        {
                                            if (this.CommunicationLanguage == GlobalClass.sCommunicationLanguageSpanish)
                                                sLine = sLine.Replace("<<BatchWindow>>", "lo antes posible");
                                            else if (this.CommunicationLanguage == GlobalClass.sCommunicationLanguageRussian)
                                                sLine = sLine.Replace("<<BatchWindow>>", "как можно скорее по получении");
                                            else
                                                sLine = sLine.Replace("<<BatchWindow>>", "as soon as possible");
                                        }
                                    }
                                }

                                sLine = sLine.Replace("<<SetNo>>", sSetNo);
                                sLine = sLine.Replace("<<SetNos>>", this.SetNos);
                                sLine = sLine.Replace("<<UnitCode>>", sUnitCode);
                                sLine = sLine.Replace("<<BeamCode>>", sBeamCode);
                                if (this.Operator.LastParticipation != null)
                                    sLine = sLine.Replace("<<LastParticipationYear>>", this.Operator.LastParticipation.LastIrradiationYear.ToString());

                                sLine = sLine.Replace("<<FollowUpNeeded>>", sFollowUpNeeded);
                                

                                sLine = sLine.Replace("<<LabCode>>", this.Operator.LabCode);
                                sLine = sLine.Replace("<<OperatorName>>", this.Operator.OperatorName);
                                sLine = sLine.Replace("<<Street>>", this.Operator.Street);
                                sLine = sLine.Replace("<<City>>", this.Operator.City);
                                sLine = sLine.Replace("<<POBox>>", this.Operator.POBox);
                                sLine = sLine.Replace("<<State>>", this.Operator.State);
                                sLine = sLine.Replace("<<Country>>", this.Operator.Country);

                                sLine = sLine.Replace("<<InstitutionalEmail>>", this.Operator.InstitutionalEmail);
                                sLine = sLine.Replace("<<InstitutionalTelephone1>>", this.Operator.InstitutionalTelephone1);
                                sLine = sLine.Replace("<<InstitutionalFax>>", this.Operator.InstitutionalFax);

                                sLine = sLine.Replace("<<PackageComment>>", this.PackageComment);
                                sLine = sLine.Replace("<<PackageDescription>>", this.PackageDescription);


                                if (this.Contact != null)
                                {
                                    sLine = sLine.Replace("<<ContactTitle>>", this.Contact.Title.Trim());
                                    sLine = sLine.Replace("<<ContactFamilyName>>", this.Contact.FamilyName.Trim());
                                    sLine = sLine.Replace("<<ContactFirstName>>", this.Contact.FirstName.Trim());
                                    sLine = sLine.Replace("<<ContactPosition>>", this.Contact.Position.Trim());
                                    sLine = sLine.Replace("<<ContactDepartment>>", this.Contact.Department.Trim());
                                    sLine = sLine.Replace("<<ContactEmail>>", this.Contact.Email.Trim());
                                    sLine = sLine.Replace("<<ContactTelephone1>>", this.Contact.Telephone1.Trim());
                                    sLine = sLine.Replace("<<ContactTelephone2>>", this.Contact.Telephone2.Trim());
                                }
                            

                                if (sLine.Contains("<<eMail.Subject>>"))
                                    sSubject = sLine.Replace("<<eMail.Subject>>", "").Trim();
                                else if (sLine.Contains("<<eMail.TemplateObject>>"))
                                { }
                                else if (sLine.Contains("<<eMail.From>>"))
                                    sEmailFrom = sLine.Replace("<<eMail.From>>", string.Empty).Trim();
                                else if (sLine.Contains("<<eMail.To>>"))
                                {
                                    sEmailTo = sLine.Replace("<<eMail.To>>", "").Trim();
                                    if (sEmailTo == "<<ContactEmail>>")
                                    {
                                        sEmailTo = string.Empty;
                                        if (this.Contact != null)
                                            sEmailTo = this.Contact.Email;
                                    }
                                    else if (sEmailTo == "<<InstitutionalEmail>>")
                                    {
                                        sEmailTo = this.Operator.InstitutionalEmail;
                                    }

                                    sEmailTo = sEmailTo.Replace(" ", "").Replace(",", ";");
                                }
                                else if (sLine.Contains("<<eMail.CC>>"))
                                {
                                    sEmailCC = sLine.Replace("<<eMail.CC>>", "").Trim();
                                    if (sEmailCC == "<<ContactEmail>>")
                                    {
                                        sEmailCC = string.Empty;
                                        if (this.Contact != null)
                                            sEmailCC = this.Contact.Email;
                                    }
                                    else if (sEmailCC == "<<InstitutionalEmail>>")
                                    {
                                        sEmailCC = this.Operator.InstitutionalEmail;
                                    }
                                    else if (sEmailCC == "<<NationalCoordinatorEmail>>")
                                    {
                                        sEmailCC = sEmailCC.Replace("<<NationalCoordinatorEmail>>", string.Empty).Trim();
                                        List<NationalCoordinatorClass> NationalCoordinatorList = GlobalClass.Manager.GetCountryCoordinatorList(this.CCode, 0);
                                        foreach (NationalCoordinatorClass NationalCoordinator in NationalCoordinatorList)
                                        {
                                            if (NationalCoordinator != null)
                                                if (NationalCoordinator.Email != string.Empty)
                                                    sEmailCC = sEmailCC + NationalCoordinator.Email + ";";
                                        }
                                    }
                                    sEmailCC = sEmailCC.Replace(" ", "").Replace(",", ";");
                                }
                                else if (sLine.Contains("<<eMail.Attachments>>"))
                                {
                                    string[] sAttachmentList = sLine.Replace("<<eMail.Attachments>>", "").Trim().Split(';');
                                    foreach (string sAttachmentFile in sAttachmentList)
                                    {
                                        if (File.Exists(GlobalClass.sApplicationStartupPath + "\\" + sAttachmentFile.Trim()))
                                            sAttachments = sAttachments + GlobalClass.sApplicationStartupPath + "\\" + sAttachmentFile.Trim() + ";";
                                    }
                                }
                                else if (sLine.Contains("<<eMail.Signature>>"))
                                    sSignature = sLine.Replace("<<eMail.Signature>>", "").Trim();

                                else
                                {
                                    sBody = sBody + sLine + System.Environment.NewLine;
                                }
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

                string sFilePath = GlobalClass.sApplicationTempPath;

                Directory.SetCurrentDirectory(GlobalClass.sApplicationStartupPath);
                if (!Directory.Exists(sFilePath))
                    Directory.CreateDirectory(sFilePath);

                string sCoveringLetter = this.GenerateDipatchCertificateCoveringLetter(GlobalClass.sFormatPDF);
                if (File.Exists(sCoveringLetter))
                    sAttachments = sAttachments + sCoveringLetter + ";";

                /*
                AttachmentClass Attachment = this.GetAttachment(GlobalClass.iAttachmentTLDCertificateCoverLetter);
                if (Attachment != null)
                {

                    string sFileName = Path.GetFileName(Attachment.AttachmentFileName);
                    string sCoveringLetter = sFilePath + sFileName;

                    sCoveringLetter = Attachment.SaveToFile(sCoveringLetter);
                    if (File.Exists(sCoveringLetter))
                        sAttachments = sAttachments + sCoveringLetter + ";";
                }
                */

                foreach (TLDSetClass TLDSet in this.GetTLDSetList())
                {
                    if (TLDSet.Certificate != null)
                    {
                        AttachmentClass CertificateAttachment = TLDSet.Certificate.GetAttachment(GlobalClass.iAttachmentTLDCertificate);
                        if (CertificateAttachment != null)
                        {
                            string sFileName = Path.GetFileName(CertificateAttachment.AttachmentFileName);
                            string sFullFileName = sFilePath + sFileName;

                            sFullFileName = CertificateAttachment.SaveToFile(sFullFileName);
                            if (File.Exists(sFullFileName))
                                sAttachments = sAttachments + sFullFileName + ";";
                        }
                    }
                }

                GlobalClass.GenerateEmailNotification(sEmailFrom, sEmailTo, sEmailCC, sSubject, sBody, sAttachments, sSaveAs, sSignature);
            }
            return sReturn;
        }
    }
}
