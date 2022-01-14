using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;

namespace DipuAlba.ConversionDGPE
{
    internal class ExcelReader
    {
        internal static dgp_declaracion Convertir(string rutaOrigen)
        {
            var declaracion = new dgp_declaracion();
            Debug.WriteLine("Llamada GetDeclaracion");

            var file = new FileInfo(rutaOrigen);
            using var package = new ExcelPackage(file);
            var sheet = package.Workbook.Worksheets["Hoja1"];
            //int colCount = sheet.Dimension.End.Column;  //get Column Count
            int rowCount = sheet.Dimension.End.Row;     //get row count
            var idx = sheet.Cells["1:1"];
            var listadoContratosOrgando = new List<dgp_declaracionEnteContratanteDepartamentoOrganoContratanteContrato>();
            var listadoOrganos = new List<dgp_declaracionEnteContratanteDepartamentoOrganoContratante>();
            for (int row = 2; row <= rowCount; row++)
            {


                var col = idx.First(c => c.Value.ToString() == nameof(declaracion.anio)).Start.Column;
                declaracion.anio = sheet.Cells[row, col].Value.ToString();
                declaracion.cabecera ??= new dgp_declaracionCabecera();

                declaracion.cabecera.RegistrosEnviados = rowCount.ToString();
                col = GetIndexColumn(idx, nameof(declaracion.cabecera.tipoAdmin));
                declaracion.cabecera.tipoAdmin = sheet.Cells[row, col].Value.ToString();
                col = GetIndexColumn(idx, nameof(declaracion.cabecera.tipoAdminLocal));
                declaracion.cabecera.tipoAdminLocal = sheet.Cells[row, col].Value.ToString();



                declaracion.cabecera.usuario = new dgp_declaracionCabeceraUsuario();
                col = GetIndexColumn(idx, nameof(declaracion.cabecera.usuario.apellido1));
                declaracion.cabecera.usuario.apellido1 = sheet.Cells[row, col].Value.ToString();
                col = GetIndexColumn(idx, nameof(declaracion.cabecera.usuario.apellido2));
                declaracion.cabecera.usuario.apellido2 = sheet.Cells[row, col].Value.ToString();
                col = GetIndexColumn(idx, nameof(declaracion.cabecera.usuario.cargo));
                declaracion.cabecera.usuario.cargo = sheet.Cells[row, col].Value.ToString();
                col = GetIndexColumn(idx, nameof(declaracion.cabecera.usuario.codPostal));
                declaracion.cabecera.usuario.codPostal = sheet.Cells[row, col].Value.ToString();
                col = GetIndexColumn(idx, nameof(declaracion.cabecera.usuario.direccion));
                declaracion.cabecera.usuario.direccion = sheet.Cells[row, col].Value.ToString();
                col = GetIndexColumn(idx, nameof(declaracion.cabecera.usuario.dni));
                declaracion.cabecera.usuario.dni = sheet.Cells[row, col].Value.ToString();
                col = GetIndexColumn(idx, nameof(declaracion.cabecera.usuario.mail));
                declaracion.cabecera.usuario.mail = sheet.Cells[row, col].Value.ToString();
                col = GetIndexColumn(idx, nameof(declaracion.cabecera.usuario.municipio));
                declaracion.cabecera.usuario.municipio = sheet.Cells[row, col].Value.ToString();
                col = GetIndexColumn(idx, nameof(declaracion.cabecera.usuario.nombre));
                declaracion.cabecera.usuario.nombre = sheet.Cells[row, col].Value.ToString();
                col = GetIndexColumn(idx, nameof(declaracion.cabecera.usuario.provincia));
                declaracion.cabecera.usuario.provincia = sheet.Cells[row, col].Value.ToString();
                col = GetIndexColumn(idx, nameof(declaracion.cabecera.usuario.telefono));
                declaracion.cabecera.usuario.telefono = sheet.Cells[row, col].Value.ToString();
                declaracion.enteContratante ??= new dgp_declaracionEnteContratante();

                col = GetIndexColumn(idx, nameof(declaracion.enteContratante.NIF));
                declaracion.enteContratante.NIF = sheet.Cells[row, col].Value.ToString();
                col = GetIndexColumn(idx, nameof(declaracion.enteContratante.nombreEnteContratante));
                declaracion.enteContratante.nombreEnteContratante = sheet.Cells[row, col].Value.ToString();
                col = GetIndexColumn(idx, nameof(declaracion.enteContratante.codEnteContratante));
                declaracion.enteContratante.codEnteContratante = sheet.Cells[row, col].Value.ToString();

                var organo = new dgp_declaracionEnteContratanteDepartamentoOrganoContratante
                {
                    codigoOrganoContratante = sheet.Cells[row, GetIndexColumn(idx, nameof(dgp_declaracionEnteContratanteDepartamentoOrganoContratante.codigoOrganoContratante))].Value?.ToString(),
                    NIF = sheet.Cells[row, GetIndexColumn(idx, nameof(dgp_declaracionEnteContratanteDepartamentoOrganoContratante.NIF))].Value.ToString()
                };
                organo.NIF = sheet.Cells[row, GetIndexColumn(idx, nameof(dgp_declaracionEnteContratanteDepartamentoOrganoContratante.nombreOrganoContratante))].Value?.ToString();

                var contratoOrgano = new dgp_declaracionEnteContratanteDepartamentoOrganoContratanteContrato
                {
                    numero = sheet.Cells[row, GetIndexColumn(idx, nameof(dgp_declaracionEnteContratanteDepartamentoOrganoContratanteContrato.numero))].Value.ToString(),
                    tipoContrato = sheet.Cells[row, GetIndexColumn(idx, nameof(dgp_declaracionEnteContratanteDepartamentoOrganoContratanteContrato.tipoContrato))].Value.ToString(),
                    provincia = sheet.Cells[row, GetIndexColumn(idx, nameof(dgp_declaracionEnteContratanteDepartamentoOrganoContratanteContrato.provincia))].Value.ToString(),
                    objeto = sheet.Cells[row, GetIndexColumn(idx, nameof(dgp_declaracionEnteContratanteDepartamentoOrganoContratanteContrato.objeto))].Value.ToString(),
                    conPorLotes = dgp_declaracionEnteContratanteDepartamentoOrganoContratanteContratoConPorLotes.no,
                    tramite = sheet.Cells[row, GetIndexColumn(idx, nameof(dgp_declaracionEnteContratanteDepartamentoOrganoContratanteContrato.tramite))].Value.ToString(),
                    procedimientoAdjud = sheet.Cells[row, GetIndexColumn(idx, nameof(dgp_declaracionEnteContratanteDepartamentoOrganoContratanteContrato.procedimientoAdjud))].Value.ToString(),
                    formaAdjud = sheet.Cells[row, GetIndexColumn(idx, nameof(dgp_declaracionEnteContratanteDepartamentoOrganoContratanteContrato.formaAdjud))].Value.ToString(),
                    importePresupuesto = sheet.Cells[row, GetIndexColumn(idx, nameof(dgp_declaracionEnteContratanteDepartamentoOrganoContratanteContrato.importePresupuesto))].Value.ToString(),
                    revisionPrecios = dgp_declaracionEnteContratanteDepartamentoOrganoContratanteContratoRevisionPrecios.no,
                    CPV = new dgp_declaracionEnteContratanteDepartamentoOrganoContratanteContratoCPV()
                    {
                        codigoCPV = sheet.Cells[row, GetIndexColumn(idx, nameof(dgp_declaracionEnteContratanteDepartamentoOrganoContratanteContratoCPV.codigoCPV))].Value.ToString(),
                        version = "2008"
                    },
                    //contratoOrgano.caracteristicaBienes = sheet.Cells[row, GetIndexColumn(idx, nameof(dgp_declaracionEnteContratanteDepartamentoOrganoContratanteContrato.caracteristicaBienes))].Value.ToString();
                    plazo = sheet.Cells[row, GetIndexColumn(idx, nameof(dgp_declaracionEnteContratanteDepartamentoOrganoContratanteContrato.plazo))].Value.ToString()
                };


                contratoOrgano.plurianual = Convert.ToInt32(contratoOrgano.plazo) > 12 ? dgp_declaracionEnteContratanteDepartamentoOrganoContratanteContratoPlurianual.si : dgp_declaracionEnteContratanteDepartamentoOrganoContratanteContratoPlurianual.no;

                var contratista = new dgp_declaracionEnteContratanteDepartamentoOrganoContratanteContratoContratista
                {
                    descripcion = sheet.Cells[row, GetIndexColumn(idx, nameof(dgp_declaracionEnteContratanteDepartamentoOrganoContratanteContratoContratista.descripcion))].Value.ToString(),
                    nif = sheet.Cells[row, GetIndexColumn(idx, nameof(dgp_declaracionEnteContratanteDepartamentoOrganoContratanteContratoContratista.nif))].Value.ToString(),
                    nacionalidad = sheet.Cells[row, GetIndexColumn(idx, nameof(dgp_declaracionEnteContratanteDepartamentoOrganoContratanteContratoContratista.nacionalidad))].Value.ToString()
                };

                contratoOrgano.contratista = new dgp_declaracionEnteContratanteDepartamentoOrganoContratanteContratoContratista[] { contratista };

                var fechaAdju = sheet.Cells[row, GetIndexColumn(idx, nameof(dgp_declaracionEnteContratanteDepartamentoOrganoContratanteContrato.fechaAdjudicacion))].Value.ToString();
                //if (!DateTime.TryParse(fechaAdju, out var fechaA)) throw new Exception(fechaAdju);
                contratoOrgano.fechaAdjudicacion = DateTime.FromOADate(Convert.ToInt32(fechaAdju));
                contratoOrgano.importeAdjudicacion = sheet.Cells[row, GetIndexColumn(idx, nameof(dgp_declaracionEnteContratanteDepartamentoOrganoContratanteContrato.importeAdjudicacion))].Value.ToString();
                //contratoOrgano.modalidadDeterminacionPrecio = sheet.Cells[row, GetIndexColumn(idx, nameof(dgp_declaracionEnteContratanteDepartamentoOrganoContratanteContrato.modalidadDeterminacionPrecio))].Value.ToString();

                var fechaForm = sheet.Cells[row, GetIndexColumn(idx, nameof(dgp_declaracionEnteContratanteDepartamentoOrganoContratanteContrato.fechaFormalizacion))].Value.ToString();
                //if (!DateTime.TryParse(fechaForm, out var fechaF)) throw new Exception(fechaForm);
                contratoOrgano.fechaFormalizacion = DateTime.FromOADate(Convert.ToInt32(fechaForm));
                listadoContratosOrgando.Add(contratoOrgano);



                organo.contrato = listadoContratosOrgando.ToArray();

                if (row == rowCount)
                {
                    listadoOrganos.Add(organo);


                    var dpto = new dgp_declaracionEnteContratanteDepartamento()
                    {
                        organoContratante = listadoOrganos.ToArray()
                    };

                    declaracion.enteContratante.departamento = new dgp_declaracionEnteContratanteDepartamento[] { dpto };

                }

            }

            return declaracion;
        }

        private static int GetIndexColumn(ExcelRange idx, string colName)
        {
            var rango = idx.FirstOrDefault(c => c.Value.ToString() == "ns1:" + colName);
            return rango == null ? 
                idx.First(c => c.Value.ToString() == colName).Start.Column : 
                rango.Start.Column;
        }
    }
}
