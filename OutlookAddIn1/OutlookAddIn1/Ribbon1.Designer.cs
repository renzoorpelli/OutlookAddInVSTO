namespace OutlookAddIn1
{
    partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Variable del diseñador necesaria.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon1()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Limpiar los recursos que se estén usando.
        /// </summary>
        /// <param name="disposing">true si los recursos administrados se deben desechar; false en caso contrario.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Código generado por el Diseñador de componentes

        /// <summary>
        /// Método necesario para admitir el Diseñador. No se puede modificar
        /// el contenido de este método con el editor de código.
        /// </summary>
        private void InitializeComponent()
        {
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.cmbProyectos = this.Factory.CreateRibbonComboBox();
            this.cmbEspecialidad = this.Factory.CreateRibbonComboBox();
            this.cmbRFQ = this.Factory.CreateRibbonComboBox();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.cmbRevision = this.Factory.CreateRibbonComboBox();
            this.btnDatosEmail = this.Factory.CreateRibbonButton();
            this.group3 = this.Factory.CreateRibbonGroup();
            this.lblProveedor = this.Factory.CreateRibbonLabel();
            this.btnContactos = this.Factory.CreateRibbonButton();
            this.btnFechaComprometida = this.Factory.CreateRibbonButton();
            this.lblFecha = this.Factory.CreateRibbonLabel();
            this.group4 = this.Factory.CreateRibbonGroup();
            this.edtSearchBox = this.Factory.CreateRibbonEditBox();
            this.group5 = this.Factory.CreateRibbonGroup();
            this.btnRutaIT = this.Factory.CreateRibbonButton();
            this.btnRutaIC = this.Factory.CreateRibbonButton();
            this.btnRefTec = this.Factory.CreateRibbonButton();
            this.group6 = this.Factory.CreateRibbonGroup();
            this.cmbCategoria = this.Factory.CreateRibbonComboBox();
            this.cmbAnalisis = this.Factory.CreateRibbonComboBox();
            this.group7 = this.Factory.CreateRibbonGroup();
            this.lblAcciones = this.Factory.CreateRibbonLabel();
            this.buttonGroup1 = this.Factory.CreateRibbonButtonGroup();
            this.btnAplicar = this.Factory.CreateRibbonButton();
            this.btnGuardarMail = this.Factory.CreateRibbonButton();
            this.buttonGroup2 = this.Factory.CreateRibbonButtonGroup();
            this.btnGenerarTemplate = this.Factory.CreateRibbonButton();
            this.group8 = this.Factory.CreateRibbonGroup();
            this.edtRemitente = this.Factory.CreateRibbonEditBox();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.group2.SuspendLayout();
            this.group3.SuspendLayout();
            this.group4.SuspendLayout();
            this.group5.SuspendLayout();
            this.group6.SuspendLayout();
            this.group7.SuspendLayout();
            this.buttonGroup1.SuspendLayout();
            this.buttonGroup2.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.group2);
            this.tab1.Groups.Add(this.group3);
            this.tab1.Groups.Add(this.group4);
            this.tab1.Groups.Add(this.group5);
            this.tab1.Groups.Add(this.group6);
            this.tab1.Groups.Add(this.group7);
            this.tab1.Groups.Add(this.group8);
            this.tab1.Label = "PREMO";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.cmbProyectos);
            this.group1.Items.Add(this.cmbEspecialidad);
            this.group1.Items.Add(this.cmbRFQ);
            this.group1.Label = "Operaciones";
            this.group1.Name = "group1";
            // 
            // cmbProyectos
            // 
            this.cmbProyectos.Label = "PROYECTOS";
            this.cmbProyectos.Name = "cmbProyectos";
            this.cmbProyectos.Text = null;
            // 
            // cmbEspecialidad
            // 
            this.cmbEspecialidad.Label = "ESPECIALIDAD";
            this.cmbEspecialidad.Name = "cmbEspecialidad";
            this.cmbEspecialidad.Text = null;
            // 
            // cmbRFQ
            // 
            this.cmbRFQ.Label = "RFQ";
            this.cmbRFQ.Name = "cmbRFQ";
            this.cmbRFQ.Text = null;
            // 
            // group2
            // 
            this.group2.Items.Add(this.cmbRevision);
            this.group2.Items.Add(this.btnDatosEmail);
            this.group2.Name = "group2";
            // 
            // cmbRevision
            // 
            this.cmbRevision.Label = "REVISIÓN";
            this.cmbRevision.Name = "cmbRevision";
            this.cmbRevision.Text = null;
            // 
            // btnDatosEmail
            // 
            this.btnDatosEmail.Label = "Obtener datos email";
            this.btnDatosEmail.Name = "btnDatosEmail";
            this.btnDatosEmail.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnDatosEmail_Click);
            // 
            // group3
            // 
            this.group3.Items.Add(this.lblProveedor);
            this.group3.Items.Add(this.btnContactos);
            this.group3.Items.Add(this.btnFechaComprometida);
            this.group3.Items.Add(this.lblFecha);
            this.group3.Name = "group3";
            // 
            // lblProveedor
            // 
            this.lblProveedor.Label = "PROVEEDOR";
            this.lblProveedor.Name = "lblProveedor";
            // 
            // btnContactos
            // 
            this.btnContactos.Label = "Surgerir Contactos";
            this.btnContactos.Name = "btnContactos";
            this.btnContactos.OfficeImageId = "ContactPictureMenu";
            this.btnContactos.ShowImage = true;
            // 
            // btnFechaComprometida
            // 
            this.btnFechaComprometida.Label = "Fecha Comprometida";
            this.btnFechaComprometida.Name = "btnFechaComprometida";
            this.btnFechaComprometida.OfficeImageId = "DataTypeDateTime";
            this.btnFechaComprometida.ShowImage = true;
            // 
            // lblFecha
            // 
            this.lblFecha.Label = "dd/MM/YYYY";
            this.lblFecha.Name = "lblFecha";
            // 
            // group4
            // 
            this.group4.Items.Add(this.edtSearchBox);
            this.group4.Items.Add(this.edtRemitente);
            this.group4.Name = "group4";
            // 
            // edtSearchBox
            // 
            this.edtSearchBox.Label = "Buscar Proyecto";
            this.edtSearchBox.Name = "edtSearchBox";
            this.edtSearchBox.OfficeImageId = "InstantSearch";
            this.edtSearchBox.ShowImage = true;
            this.edtSearchBox.Text = null;
            // 
            // group5
            // 
            this.group5.Items.Add(this.btnRutaIT);
            this.group5.Items.Add(this.btnRutaIC);
            this.group5.Items.Add(this.btnRefTec);
            this.group5.Name = "group5";
            // 
            // btnRutaIT
            // 
            this.btnRutaIT.Label = "RUTA IT";
            this.btnRutaIT.Name = "btnRutaIT";
            this.btnRutaIT.OfficeImageId = "FileCloseDatabase";
            this.btnRutaIT.ShowImage = true;
            // 
            // btnRutaIC
            // 
            this.btnRutaIC.Label = "RUTA IC";
            this.btnRutaIC.Name = "btnRutaIC";
            this.btnRutaIC.OfficeImageId = "FileCloseDatabase";
            this.btnRutaIC.ShowImage = true;
            // 
            // btnRefTec
            // 
            this.btnRefTec.Label = "Referente Técnico";
            this.btnRefTec.Name = "btnRefTec";
            this.btnRefTec.OfficeImageId = "InstantSearch";
            this.btnRefTec.ShowImage = true;
            // 
            // group6
            // 
            this.group6.Items.Add(this.cmbCategoria);
            this.group6.Items.Add(this.cmbAnalisis);
            this.group6.Name = "group6";
            // 
            // cmbCategoria
            // 
            this.cmbCategoria.Label = "CATEGORIA / FLAG";
            this.cmbCategoria.Name = "cmbCategoria";
            this.cmbCategoria.Text = null;
            // 
            // cmbAnalisis
            // 
            this.cmbAnalisis.Label = "ANÁLISIS TÉCNICO";
            this.cmbAnalisis.Name = "cmbAnalisis";
            this.cmbAnalisis.Text = null;
            // 
            // group7
            // 
            this.group7.Items.Add(this.lblAcciones);
            this.group7.Items.Add(this.buttonGroup1);
            this.group7.Items.Add(this.buttonGroup2);
            this.group7.Name = "group7";
            // 
            // lblAcciones
            // 
            this.lblAcciones.Label = "ACCIONES RÁPIDAS";
            this.lblAcciones.Name = "lblAcciones";
            // 
            // buttonGroup1
            // 
            this.buttonGroup1.Items.Add(this.btnAplicar);
            this.buttonGroup1.Items.Add(this.btnGuardarMail);
            this.buttonGroup1.Name = "buttonGroup1";
            // 
            // btnAplicar
            // 
            this.btnAplicar.Label = "Aplicar en PREMO";
            this.btnAplicar.Name = "btnAplicar";
            // 
            // btnGuardarMail
            // 
            this.btnGuardarMail.Label = "Guardar Mail";
            this.btnGuardarMail.Name = "btnGuardarMail";
            // 
            // buttonGroup2
            // 
            this.buttonGroup2.Items.Add(this.btnGenerarTemplate);
            this.buttonGroup2.Name = "buttonGroup2";
            // 
            // btnGenerarTemplate
            // 
            this.btnGenerarTemplate.Label = "Generar Template";
            this.btnGenerarTemplate.Name = "btnGenerarTemplate";
            this.btnGenerarTemplate.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnGenerarTemplate_Click);
            // 
            // group8
            // 
            this.group8.Name = "group8";
            // 
            // edtRemitente
            // 
            this.edtRemitente.Enabled = false;
            this.edtRemitente.Label = "Remitente";
            this.edtRemitente.Name = "edtRemitente";
            this.edtRemitente.OfficeImageId = "ContactProperties";
            this.edtRemitente.ShowImage = true;
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Outlook.Explorer";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.group3.ResumeLayout(false);
            this.group3.PerformLayout();
            this.group4.ResumeLayout(false);
            this.group4.PerformLayout();
            this.group5.ResumeLayout(false);
            this.group5.PerformLayout();
            this.group6.ResumeLayout(false);
            this.group6.PerformLayout();
            this.group7.ResumeLayout(false);
            this.group7.PerformLayout();
            this.buttonGroup1.ResumeLayout(false);
            this.buttonGroup1.PerformLayout();
            this.buttonGroup2.ResumeLayout(false);
            this.buttonGroup2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonComboBox cmbProyectos;
        internal Microsoft.Office.Tools.Ribbon.RibbonComboBox cmbEspecialidad;
        internal Microsoft.Office.Tools.Ribbon.RibbonComboBox cmbRFQ;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonComboBox cmbRevision;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDatosEmail;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnContactos;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group4;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel lblProveedor;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox edtSearchBox;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group5;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnRutaIT;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnRutaIC;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnRefTec;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group6;
        internal Microsoft.Office.Tools.Ribbon.RibbonComboBox cmbCategoria;
        internal Microsoft.Office.Tools.Ribbon.RibbonComboBox cmbAnalisis;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group7;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel lblAcciones;
        internal Microsoft.Office.Tools.Ribbon.RibbonButtonGroup buttonGroup1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAplicar;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnGuardarMail;
        internal Microsoft.Office.Tools.Ribbon.RibbonButtonGroup buttonGroup2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnGenerarTemplate;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group8;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnFechaComprometida;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel lblFecha;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox edtRemitente;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
