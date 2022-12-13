using Ionic.Zip;
using Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Mime;
using System.Text;
using System.Windows.Forms;

namespace OutlookAddIn1
{
    public partial class Ribbon1
    {
        private string carpetaSeleccionadaPath; //path carpeta seleccionada generar Template
        private string fileFullPath;
        private string logFileFullPath;



        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            //obtengo la lista de nombres de los proyectos y las cargo en el cmb
            var listaProyectos = ObtenerNombresProyectos();
            this.CargarElementosCmbProyectos(listaProyectos);
        }

        #region PUNTO 2
        private void btnGenerarTemplate_Click(object sender, RibbonControlEventArgs e)
        {
            FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog();


            if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
            {
                carpetaSeleccionadaPath = folderBrowserDialog.SelectedPath; // obtengo el directorio 
            }

            if (VerificarTamanioDirectorio(carpetaSeleccionadaPath))
            {
                if (ConvertirDirectorioZip(carpetaSeleccionadaPath))
                {
                    MessageBox.Show($"El directorio fue comprimido exitosamente y guardado en {carpetaSeleccionadaPath}", "Atención", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    GenerarMail();
                }
            }
            else
            {
                MessageBox.Show("Error, el directorio excede los 15MB", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        #endregion

        #region PUNTO 4.1
        private void btnRegistrarEmail_Click(object sender, RibbonControlEventArgs e)
        {
            if (cmbProyectos.Text != "" && lblFecha.Label != "dd/MM/YYYY")
            {
                if (edtSearchBox.Text != "")
                {
                    string path = "C:\\Users\\Renzo\\Documents\\develop\\TRABAJO\\OoutlookAddIn\\OutlookAddIn1\\OutlookAddIn1\\bin\\Debug";
                    StringBuilder stringBuilder = new StringBuilder();
                    stringBuilder.Append(cmbProyectos.Text + DateTime.Now.ToShortDateString().Replace("/", "_") + ".txt");
                    string fileDirectory = Path.Combine(path, stringBuilder.ToString());

                    string datos = ObtenerDatosMailSeleccionado(edtSearchBox.Text);
                    GenerarCuerpoArchivo(datos, fileDirectory);
                    MessageBox.Show($"Archivo guardado corretamente en ${fileDirectory}", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show("El proyecto o Archivo del mail no existe", "Atención", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            else
            {
                MessageBox.Show("Debe seleccionar un proyecto y un timestamp", "Atención", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        /// <summary>
        /// método encargado de generar el cuerpo del archivo el cual persistirá los datos del mail
        /// </summary>
        /// <param name="texto">los datos que contendrá el txt</param>
        /// <param name="rutaArchivo"></param>
        private void GenerarCuerpoArchivo(string texto, string rutaArchivo)
        {
            GenerarLogsMail(texto, rutaArchivo);
        }

        private string ObtenerDatosMailSeleccionado(string proyecto)
        {
            StringBuilder stringBuilder = new StringBuilder();
            var thisAddIn = Globals.ThisAddIn;

            Microsoft.Office.Interop.Outlook.MailItem mailSeleccionado = (MailItem)thisAddIn.Application.ActiveExplorer().Selection[1];

            if (proyecto.Contains("PROYECTO"))
            {
                if (!(mailSeleccionado is null))
                {
                    Microsoft.Office.Interop.Outlook.Inspector inspector = mailSeleccionado.GetInspector;

                    if (inspector.IsWordMail())
                    {
                        var outlookWordDocument = inspector.WordEditor as Microsoft.Office.Interop.Word.Document;

                        Microsoft.Office.Interop.Word.Find find = outlookWordDocument.Content.Find;

                        if (find.HitHighlight(proyecto, Microsoft.Office.Interop.Word.WdColor.wdColorAqua))
                        {
                            string nombreProyecto = proyecto.Split('-')[0];
                            stringBuilder.AppendLine("NOMBRE PROYECTO: " + nombreProyecto);
                            stringBuilder.AppendLine("EMAIL REMITENTE: " + mailSeleccionado.SendUsingAccount.SmtpAddress);
                            stringBuilder.AppendLine("FECHA ACTUAL: " + DateTime.Now.ToShortDateString());
                        }
                        else
                        {
                            MessageBox.Show("No se encontró el proyecto", "Atención", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }

                    }
                }
            }
            else
            {
                MessageBox.Show("Ingrese un nombre válido", "Atención", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }


            return stringBuilder.ToString();
        }

        #endregion

        #region PUNTO 4 BUSCAR EMAIL

        /// <summary>
        /// metodo encargado del evento click, se encarga de obtener
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnDatosEmail_Click(object sender, RibbonControlEventArgs e)
        {
            if (edtSearchBox.Text != "")
            {
                BuscarProyectoCuerpoMail(edtSearchBox.Text);
            }
            else
            {
                MessageBox.Show("Ingrese el TimeStamp del proyecto", "Atención", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        public void BuscarProyectoCuerpoMail(string proyecto)
        {
            //obtengo referencia al AddIn
            var thisAddIn = Globals.ThisAddIn;

            Microsoft.Office.Interop.Outlook.MailItem mailSeleccionado = (MailItem)thisAddIn.Application.ActiveExplorer().Selection[1];

            if (proyecto.Contains("PROYECTO"))
            {
                if (!(mailSeleccionado is null))
                {
                    Microsoft.Office.Interop.Outlook.Inspector inspector = mailSeleccionado.GetInspector;

                    if (inspector.IsWordMail())
                    {
                        var outlookWordDocument = inspector.WordEditor as Microsoft.Office.Interop.Word.Document;

                        Microsoft.Office.Interop.Word.Find find = outlookWordDocument.Content.Find;

                        if (find.HitHighlight(proyecto, Microsoft.Office.Interop.Word.WdColor.wdColorAqua))
                        {
                            MessageBox.Show("Se encontró el proyecto", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);

                            SeleccionarProyectoCMB(proyecto);
                            EstablecerFechaLabel(proyecto);
                            EstablecerRemitenteLabel(mailSeleccionado.SendUsingAccount.SmtpAddress);
                        }
                        else
                        {
                            MessageBox.Show("No se encontró el proyecto", "Atención", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }

                    }
                }

            }
            else
            {
                MessageBox.Show("Ingrese un nombre válido", "Atención", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        public void SeleccionarProyectoCMB(string proyecto)
        {
            string nombre = proyecto.Split('-')[0];
            cmbProyectos.Text = nombre;
        }

        public void EstablecerFechaLabel(string proyecto)
        {
            string[] nombre = proyecto.Split('-');

            StringBuilder stringBuilder = new StringBuilder();
            stringBuilder.Append(nombre[1] + "/");
            stringBuilder.Append(nombre[2] + "/");
            stringBuilder.Append(nombre[3]);

            lblFecha.Label = stringBuilder.ToString();
        }

        public void EstablecerRemitenteLabel(string remitente)
        {
            if (remitente.Contains("@"))
            {
                edtRemitente.Text = remitente;
            }
            else
            {
                edtRemitente.Text = "Remitente invalido";
            }
        }

        #endregion

        #region Carga y Obtencion elementos ComboBox punto 1
        private void CargarElementosCmbProyectos(List<string> datos)
        {

            foreach (var item in datos)
            {
                RibbonDropDownItem newItem = Factory.CreateRibbonDropDownItem();
                newItem.Tag = item.ToString();
                newItem.Label = item.ToString();
                this.cmbProyectos.Items.Add(newItem);

            }
        }


        private List<string> ObtenerNombresProyectos()
        {
            List<string> nombres = new List<string>();

            var lista = Entidades2.Services.ProyectoService.ObtenerProyectos();
            foreach (var proyecto in lista)
            {
                nombres.Add(proyecto.Nombre);
            }

            return nombres;

        }
        #endregion

        #region Comprimir Carpeta Punto 2

        private bool VerificarTamanioDirectorio(string path)
        {
            if (!(path is null))
            {
                DirectoryInfo directory = new DirectoryInfo(path);
                FileInfo[] archivos = directory.GetFiles();
                decimal tamanio = 0; // tamaño total del directorio

                foreach (var archivo in archivos)
                {
                    tamanio += archivo.Length;
                }

                if (tamanio > 15000000)
                {
                    return false;
                }

            }
            return true;
        }

        private bool ConvertirDirectorioZip(string path)
        {
            if (!(path is null))
            {
                using (ZipFile zip = new ZipFile())
                {
                    zip.UseUnicodeAsNecessary = true;
                    zip.AddDirectory(path);
                    zip.CompressionLevel = Ionic.Zlib.CompressionLevel.BestCompression;
                    zip.Comment = "El archivo zip fue creado" + System.DateTime.Now.ToString("G");
                    this.fileFullPath = path + "/Directorio.zip";//seteo el path completo del archivo
                    zip.Save(fileFullPath);
                    return true;
                }
            }
            return false;
        }

        #endregion

        #region Envio de Mail punto 2.B

        //El método estará asociado al evento Inspectors
        private void GenerarMail()
        {

            var ol = new Microsoft.Office.Interop.Outlook.Application();

            Microsoft.Office.Interop.Outlook.MailItem mailItem = ol.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olMailItem) as Microsoft.Office.Interop.Outlook.MailItem;

            if (mailItem != null)
            {
                if (mailItem.EntryID == null && this.cmbProyectos.Text != "")
                {
                    mailItem.Subject = this.cmbProyectos.Text;
                    string cuerpoMail = LeerTxt("C:\\Users\\Renzo\\Documents\\develop\\TRABAJO\\OoutlookAddIn\\OutlookAddIn1\\OutlookAddIn1\\bin\\Debug\\CartaCliente.txt");

                    StringBuilder stringBuilder = new StringBuilder();
                    stringBuilder.Append(cuerpoMail);
                    stringBuilder.Replace("NOMBREPROYECTO", this.cmbProyectos.Text);
                    stringBuilder.Append('\n' + $"{this.cmbProyectos.Text}-{DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss")}");

                    if (!File.Exists(this.fileFullPath))
                    {
                        MessageBox.Show("El arhivo no existe", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    else
                    {
                        mailItem.Attachments.Add(fileFullPath, Microsoft.Office.Interop.Outlook.OlAttachmentType.olByValue, Type.Missing, Type.Missing);
                    }
                    mailItem.Body = stringBuilder.ToString();

                    //obtengo referencia al AddIn
                    var thisAddIn = Globals.ThisAddIn;


                    mailItem.DeleteAfterSubmit = false;

                    mailItem.Display();


                    Microsoft.Office.Interop.Outlook.Folder carpetaEnviados = (Folder)thisAddIn.Application.Session.GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderSentMail);

                    if (carpetaEnviados != null)
                    {
                        mailItem.SaveSentMessageFolder = carpetaEnviados;
                        mailItem.Save();
                        string nombreArchivo = $"{this.cmbProyectos.Text}-{DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss")}".Replace('/', '_').Replace(':', '_');

                        logFileFullPath = "C:\\Users\\Renzo\\Documents\\develop\\TRABAJO\\OoutlookAddIn\\OutlookAddIn1\\OutlookAddIn1\\bin\\Debug\\" + nombreArchivo + ".txt";

                        this.GenerarLogsMail($"{this.cmbProyectos.Text}-{DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss")}", logFileFullPath);
                    }

                }
                else
                {
                    MessageBox.Show("Por favor Seleccione un proyecto", "Atención", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }

            }
        }

        /// <summary>
        /// metodo encargado de leer el contenido de un archivo txt
        /// </summary>
        /// <param name="fileFullName"></param>
        /// <returns></returns>
        public string LeerTxt(string fileFullName)
        {
            string cuerpoMail = "";
            if (!(fileFullName is null))
            {
                try
                {
                    using (StreamReader streamReader = new StreamReader(fileFullName))
                    {
                        cuerpoMail = streamReader.ReadToEnd();
                    }
                }
                catch (System.Exception)
                {
                    throw;
                }
            }
            return cuerpoMail;
        }
        #endregion

        #region PUNTO 3 Generar Logs Envio Mail

        /// <summary>
        /// metodo encargado de generar un log de un mail enviado en un arhivo txt
        /// </summary>
        /// <param name="mailInfo"></param>
        /// <param name="filename"></param>
        private void GenerarLogsMail(string mailInfo, string filename)
        {
            using (StreamWriter streamWriter = new StreamWriter(filename))
            {
                streamWriter.Write(mailInfo);
            }

        }




        #endregion

        #region PUNTO 5
        private void btnGuardarMail_Click(object sender, RibbonControlEventArgs e)
        {
            var thisAddIn = Globals.ThisAddIn;
            Microsoft.Office.Interop.Outlook.MailItem mailSeleccionado = (MailItem)thisAddIn.Application.ActiveExplorer().Selection[1];

            if (!(mailSeleccionado is null) && mailSeleccionado.Subject.Contains("PROYECTO"))
            {
                FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog();
                if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
                {
                    carpetaSeleccionadaPath = folderBrowserDialog.SelectedPath; // obtengo el directorio

                    GuardarMailMSG(carpetaSeleccionadaPath, mailSeleccionado);
                }
            }
            else
            {
                MessageBox.Show($"Por favor seleccione un mail válido", "Atención", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }

        }

        private void GuardarMailMSG(string pathGuardar, Microsoft.Office.Interop.Outlook.MailItem mailSeleccionado)
        {
            try
            {
                string date = mailSeleccionado.ReceivedTime.ToString().Replace("/", "_").Replace(":", "_");
                string filename = mailSeleccionado.Subject + date + ".msg";
                string fileFullPath = Path.Combine(pathGuardar, filename);
                mailSeleccionado.SaveAs(fileFullPath, Microsoft.Office.Interop.Outlook.OlSaveAsType.olMSG);
                MessageBox.Show($"Mail Guardado Correctamente en {fileFullPath}", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message, "Atención", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }
        }

    }

    #endregion
}

