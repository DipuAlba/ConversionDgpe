namespace DipuAlba.ConversionDGPE
{
    public partial class Principal : Form
    {
        public Principal()
        {
            InitializeComponent();
        }

        private void btnOpenFile_Click(object sender, EventArgs e)
        {
            openFileDialogExcel.ShowDialog();

            if (!string.IsNullOrWhiteSpace(openFileDialogExcel.FileName))
            {
                tbSourceFile.Text = openFileDialogExcel.FileName;
            }
        }

        private void btnConvertir_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(tbSourceFile.Text))
            {
                MessageBox.Show("No ha indicado ningún fichero");
                return;
            }

            try
            {
                Cursor.Current = Cursors.WaitCursor;
                var fileInfo = new FileInfo(tbSourceFile.Text);
                if (!fileInfo.Exists)
                {
                    MessageBox.Show("El fichero no existe");
                    return;
                }

                var declaracion = ExcelReader.Convertir(tbSourceFile.Text);
                var resul = Serialize(declaracion);

                var rutaDestino = Path.Combine(fileInfo.DirectoryName, $"{DateTime.Now.Ticks}.xml");
                File.WriteAllBytes(rutaDestino, resul);
                MessageBox.Show($"El resultado se ha guardado en {rutaDestino}");

            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error procesando el fichero: {ex}");
            }
            finally
            {
                Cursor.Current = Cursors.Default;
            }

        }

        private static byte[] Serialize(dgp_declaracion declaracion)
        {
            var rutaTemporal = Path.GetTempFileName() + ".xml";
            XmlSerializer.Serialize(declaracion, rutaTemporal);
            try
            {
                return File.ReadAllBytes(rutaTemporal);
            }
            finally
            {
                File.Delete(rutaTemporal);
            }
        }
    }
    
}