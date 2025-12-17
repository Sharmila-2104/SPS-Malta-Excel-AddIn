namespace SPS_Malta_AaddIn
{
    partial class MaltaRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public MaltaRibbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.btn_GivenName = this.Factory.CreateRibbonButton();
            this.btn_Surname = this.Factory.CreateRibbonButton();
            this.btn_Place = this.Factory.CreateRibbonButton();
            this.btn_FileNameImageNumber = this.Factory.CreateRibbonButton();
            this.btn_FourDigit = this.Factory.CreateRibbonButton();
            this.btn_Gender = this.Factory.CreateRibbonButton();
            this.btn_MaritalStatus = this.Factory.CreateRibbonButton();
            this.btn_ExcelMerge = this.Factory.CreateRibbonButton();
            this.btn_Split = this.Factory.CreateRibbonButton();
            this.btn_BirthToBaptisms = this.Factory.CreateRibbonButton();
            this.btn_MarriageToBanns = this.Factory.CreateRibbonButton();
            this.btn_DeathToBurial = this.Factory.CreateRibbonButton();
            this.btn_Headercheck = this.Factory.CreateRibbonButton();
            this.btn_Trim = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Label = "SPS_MaltaAddIns";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.btn_GivenName);
            this.group1.Items.Add(this.btn_Surname);
            this.group1.Items.Add(this.btn_Place);
            this.group1.Items.Add(this.btn_FileNameImageNumber);
            this.group1.Items.Add(this.btn_FourDigit);
            this.group1.Items.Add(this.btn_Gender);
            this.group1.Items.Add(this.btn_MaritalStatus);
            this.group1.Items.Add(this.btn_ExcelMerge);
            this.group1.Items.Add(this.btn_Split);
            this.group1.Items.Add(this.btn_BirthToBaptisms);
            this.group1.Items.Add(this.btn_MarriageToBanns);
            this.group1.Items.Add(this.btn_DeathToBurial);
            this.group1.Items.Add(this.btn_Headercheck);
            this.group1.Items.Add(this.btn_Trim);
            this.group1.Label = "Malta";
            this.group1.Name = "group1";
            // 
            // btn_GivenName
            // 
            this.btn_GivenName.Image = global::SPS_Malta_AaddIn.Properties.Resources.database__2_;
            this.btn_GivenName.Label = "GivenName";
            this.btn_GivenName.Name = "btn_GivenName";
            this.btn_GivenName.ShowImage = true;
            this.btn_GivenName.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_GivenName_Click);
            // 
            // btn_Surname
            // 
            this.btn_Surname.Image = global::SPS_Malta_AaddIn.Properties.Resources.database__2_1;
            this.btn_Surname.Label = "SurNames";
            this.btn_Surname.Name = "btn_Surname";
            this.btn_Surname.ShowImage = true;
            this.btn_Surname.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_Surname_Click);
            // 
            // btn_Place
            // 
            this.btn_Place.Image = global::SPS_Malta_AaddIn.Properties.Resources.database__2_2;
            this.btn_Place.Label = "Residence";
            this.btn_Place.Name = "btn_Place";
            this.btn_Place.ShowImage = true;
            this.btn_Place.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_Place_Click);
            // 
            // btn_FileNameImageNumber
            // 
            this.btn_FileNameImageNumber.Image = global::SPS_Malta_AaddIn.Properties.Resources.database__2_3;
            this.btn_FileNameImageNumber.Label = "FileNameImageNumber";
            this.btn_FileNameImageNumber.Name = "btn_FileNameImageNumber";
            this.btn_FileNameImageNumber.ShowImage = true;
            this.btn_FileNameImageNumber.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_FileNameImageNumber_Click);
            // 
            // btn_FourDigit
            // 
            this.btn_FourDigit.Image = global::SPS_Malta_AaddIn.Properties.Resources.ten1;
            this.btn_FourDigit.Label = "TenDigit";
            this.btn_FourDigit.Name = "btn_FourDigit";
            this.btn_FourDigit.ShowImage = true;
            this.btn_FourDigit.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_FourDigit_Click);
            // 
            // btn_Gender
            // 
            this.btn_Gender.Image = global::SPS_Malta_AaddIn.Properties.Resources.Gender;
            this.btn_Gender.Label = "Gender";
            this.btn_Gender.Name = "btn_Gender";
            this.btn_Gender.ShowImage = true;
            this.btn_Gender.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_Gender_Click);
            // 
            // btn_MaritalStatus
            // 
            this.btn_MaritalStatus.Image = global::SPS_Malta_AaddIn.Properties.Resources.Marital;
            this.btn_MaritalStatus.Label = "MaritalStatus";
            this.btn_MaritalStatus.Name = "btn_MaritalStatus";
            this.btn_MaritalStatus.ShowImage = true;
            this.btn_MaritalStatus.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_MaritalStatus_Click);
            // 
            // btn_ExcelMerge
            // 
            this.btn_ExcelMerge.Image = global::SPS_Malta_AaddIn.Properties.Resources.Merge;
            this.btn_ExcelMerge.Label = "ExcelMerge";
            this.btn_ExcelMerge.Name = "btn_ExcelMerge";
            this.btn_ExcelMerge.ShowImage = true;
            this.btn_ExcelMerge.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_ExcelMerge_Click);
            // 
            // btn_Split
            // 
            this.btn_Split.Image = global::SPS_Malta_AaddIn.Properties.Resources.Split;
            this.btn_Split.Label = "SplitExcel";
            this.btn_Split.Name = "btn_Split";
            this.btn_Split.ShowImage = true;
            this.btn_Split.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_Split_Click);
            // 
            // btn_BirthToBaptisms
            // 
            this.btn_BirthToBaptisms.Image = global::SPS_Malta_AaddIn.Properties.Resources.Compare;
            this.btn_BirthToBaptisms.Label = "BirthToBaptisms";
            this.btn_BirthToBaptisms.Name = "btn_BirthToBaptisms";
            this.btn_BirthToBaptisms.ShowImage = true;
            this.btn_BirthToBaptisms.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_BirthToBaptisms_Click);
            // 
            // btn_MarriageToBanns
            // 
            this.btn_MarriageToBanns.Image = global::SPS_Malta_AaddIn.Properties.Resources.Compare1;
            this.btn_MarriageToBanns.Label = "MarriageToBanns";
            this.btn_MarriageToBanns.Name = "btn_MarriageToBanns";
            this.btn_MarriageToBanns.ShowImage = true;
            this.btn_MarriageToBanns.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_MarriageToBanns_Click);
            // 
            // btn_DeathToBurial
            // 
            this.btn_DeathToBurial.Image = global::SPS_Malta_AaddIn.Properties.Resources.Compare2;
            this.btn_DeathToBurial.Label = "DeathToBurial";
            this.btn_DeathToBurial.Name = "btn_DeathToBurial";
            this.btn_DeathToBurial.ShowImage = true;
            this.btn_DeathToBurial.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_DeathToBurial_Click);
            // 
            // btn_Headercheck
            // 
            this.btn_Headercheck.Image = global::SPS_Malta_AaddIn.Properties.Resources.Header;
            this.btn_Headercheck.Label = "Header check";
            this.btn_Headercheck.Name = "btn_Headercheck";
            this.btn_Headercheck.ShowImage = true;
            this.btn_Headercheck.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_Headercheck_Click);
            // 
            // btn_Trim
            // 
            this.btn_Trim.Image = global::SPS_Malta_AaddIn.Properties.Resources.Trim;
            this.btn_Trim.Label = "Trim";
            this.btn_Trim.Name = "btn_Trim";
            this.btn_Trim.ShowImage = true;
            this.btn_Trim.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_Trim_Click);
            // 
            // MaltaRibbon
            // 
            this.Name = "MaltaRibbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.MaltaRibbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_Surname;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_Place;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_GivenName;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_FileNameImageNumber;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_FourDigit;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_Gender;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_MaritalStatus;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_Split;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_ExcelMerge;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_BirthToBaptisms;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_MarriageToBanns;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_DeathToBurial;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_Headercheck;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_Trim;
    }

    partial class ThisRibbonCollection
    {
        internal MaltaRibbon MaltaRibbon
        {
            get { return this.GetRibbon<MaltaRibbon>(); }
        }
    }
}
