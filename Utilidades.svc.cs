using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.ServiceModel;
using System.Text;
using OSIsoft.AF;
using OSIsoft.AF.Asset;
using OSIsoft.AF.EventFrame;
using OSIsoft.AF.PI;
using OSIsoft.AF.Time;
using OSIsoft.AF.UnitsOfMeasure;
using OSIsoft.AF.Data;
using System.Web;
using System.Web.Configuration;
using System.Net;
using OSIsoft.AF.Search;
using System.Data;
using System.IO;
using UNACEM.PI.DAL;
using common;
using Newtonsoft.Json.Linq;
using OfficeOpenXml;
using Microsoft.SharePoint.Client;
using System.Security;
using OSIsoft.AF.Analysis;
using System.Web.Script.Serialization;

namespace wsUNACEMPI
{
    // NOTE: You can use the "Rename" command on the "Refactor" menu to change the class name "Utilidades" in code, svc and config file together.
    // NOTE: In order to launch WCF Test Client for testing this service, please select Utilidades.svc or Utilidades.svc.cs at the Solution Explorer and start debugging.
    public class Utilidades : IUtilidades
    {
        //metodos utilidad migrar tags

        public Tag ObtenerTag(string servidor, string tag)
        {
            try
            {
                Tag oTag = new Tag();
                oTag.Nombre = tag;

                //PIServer myPIServer = new PIServers().DefaultPIServer;                                

                PIServers oPIServers = new PIServers();
                PIServer myPIServer = oPIServers[servidor];
                myPIServer.Connect();

                PIPoint lo_PiPoint;

                try
                {
                    lo_PiPoint = PIPoint.FindPIPoint(myPIServer, tag);
                    oTag.Nombre = lo_PiPoint.Name;
                    oTag.Tipo = lo_PiPoint.PointType.ToString();
                    oTag.Existe = true;
                }
                catch (Exception ex1)
                {
                    oTag.Existe = false;
                    Funciones.CapturarError(ex1, "Utilidades.svc - ObtenerTag", "svcOperacionesPI", false);
                    return oTag;
                }

                return oTag;
            }
            catch (Exception ex)
            {
                Funciones.CapturarError(ex, "Utilidades.svc - ObtenerTag", "svcOperacionesPI", false);
                return null;
            }

        }


        public List<TagValor> ListarTagValores(string servidor, string tag, string fechaini, string fechafin)
        {
            List<TagValor> oValores = new List<TagValor>();
            TagValor oValor = new TagValor();
            DateTime ldt_DateIni;
            DateTime ldt_DateFin;

            try
            {
                if (fechaini.Length == 16)
                {
                    ldt_DateIni = new DateTime(Convert.ToInt32(fechaini.Substring(0, 4)), Convert.ToInt32(fechaini.Substring(5, 2)), Convert.ToInt32(fechaini.Substring(8, 2)),
                                               Convert.ToInt32(fechaini.Substring(11, 2)), Convert.ToInt32(fechaini.Substring(14, 2)), 0);
                }
                else
                {
                    ldt_DateIni = new DateTime(Convert.ToInt32(fechaini.Substring(0, 4)), Convert.ToInt32(fechaini.Substring(5, 2)), Convert.ToInt32(fechaini.Substring(8, 2)),
                                               Convert.ToInt32(fechaini.Substring(11, 2)), Convert.ToInt32(fechaini.Substring(14, 2)), Convert.ToInt32(fechaini.Substring(17, 2)));
                }
            }
            catch (Exception ex)
            {
                Funciones.CapturarError(ex, "Operaciones.svc - ListarTagValores", "svcOperacionesPI", false);
                return oValores;
            }

            try
            {
                if (fechafin.Length == 16)
                {
                    ldt_DateFin = new DateTime(Convert.ToInt32(fechafin.Substring(0, 4)), Convert.ToInt32(fechafin.Substring(5, 2)), Convert.ToInt32(fechafin.Substring(8, 2)),
                                           Convert.ToInt32(fechafin.Substring(11, 2)), Convert.ToInt32(fechafin.Substring(14, 2)), 0);
                }
                else
                {
                    ldt_DateFin = new DateTime(Convert.ToInt32(fechafin.Substring(0, 4)), Convert.ToInt32(fechafin.Substring(5, 2)), Convert.ToInt32(fechafin.Substring(8, 2)),
                                           Convert.ToInt32(fechafin.Substring(11, 2)), Convert.ToInt32(fechafin.Substring(14, 2)), Convert.ToInt32(fechafin.Substring(17, 2)));
                }
            }
            catch (Exception ex)
            {
                Funciones.CapturarError(ex, "Operaciones.svc - ListarTagValores", "svcOperacionesPI", false);
                return oValores;
            }


            try
            {
                //PIServer myPIServer = new PIServers().DefaultPIServer;

                PIServers oPIServers = new PIServers();
                PIServer myPIServer = oPIServers[servidor];
                myPIServer.Connect();

                PIPoint lo_PiPoint;
                lo_PiPoint = PIPoint.FindPIPoint(myPIServer, tag);

                AFTime oAFTimeIni = new AFTime(ldt_DateIni.ToUniversalTime());
                AFTime oAFTimeFin = new AFTime(ldt_DateFin.ToUniversalTime());
                AFTimeRange myRange = new AFTimeRange(oAFTimeIni, oAFTimeFin);

                AFValues lo_AFValues = lo_PiPoint.RecordedValues(myRange, AFBoundaryType.Inside, "", false);

                foreach (AFValue oAFValue in lo_AFValues)
                {
                    oValor = new TagValor();
                    oValor.Nombre = oAFValue.PIPoint.Name;
                    oValor.Valor = oAFValue.Value.ToString();
                    oValor.Fecha = oAFValue.Timestamp.LocalTime.ToString("yyyy-MM-dd HH:mm:ss");

                    oValores.Add(oValor);
                }

                return oValores;

            }
            catch (Exception ex)
            {
                Funciones.CapturarError(ex, "Utilidades.svc - ListarTagValores", "svcOperacionesPI", false);
                return oValores;
            }
        }


        public Resultado MigrarTag(EntradaMigrarTag EntradaMigrarTag)
        {
            DateTime ldt_DateIni;
            DateTime ldt_DateFin;

            List<AFValue> oAFValoresOrigen = new List<AFValue>();

            try
            {
                if (EntradaMigrarTag.fechaini.Length == 16)
                {
                    ldt_DateIni = new DateTime(Convert.ToInt32(EntradaMigrarTag.fechaini.Substring(0, 4)), Convert.ToInt32(EntradaMigrarTag.fechaini.Substring(5, 2)), Convert.ToInt32(EntradaMigrarTag.fechaini.Substring(8, 2)),
                                               Convert.ToInt32(EntradaMigrarTag.fechaini.Substring(11, 2)), Convert.ToInt32(EntradaMigrarTag.fechaini.Substring(14, 2)), 0);
                }
                else
                {
                    ldt_DateIni = new DateTime(Convert.ToInt32(EntradaMigrarTag.fechaini.Substring(0, 4)), Convert.ToInt32(EntradaMigrarTag.fechaini.Substring(5, 2)), Convert.ToInt32(EntradaMigrarTag.fechaini.Substring(8, 2)),
                                               Convert.ToInt32(EntradaMigrarTag.fechaini.Substring(11, 2)), Convert.ToInt32(EntradaMigrarTag.fechaini.Substring(14, 2)), Convert.ToInt32(EntradaMigrarTag.fechaini.Substring(17, 2)));
                }       
            }
            catch (Exception ex)
            {
                Funciones.CapturarError(ex, "Utilidades.svc - MigrarTag", "svcOperacionesPI", false);
                return new Resultado(-1, ex.Message);
            }

            try
            {
                if (EntradaMigrarTag.fechafin.Length == 16)
                {
                    ldt_DateFin = new DateTime(Convert.ToInt32(EntradaMigrarTag.fechafin.Substring(0, 4)), Convert.ToInt32(EntradaMigrarTag.fechafin.Substring(5, 2)), Convert.ToInt32(EntradaMigrarTag.fechafin.Substring(8, 2)),
                                               Convert.ToInt32(EntradaMigrarTag.fechafin.Substring(11, 2)), Convert.ToInt32(EntradaMigrarTag.fechafin.Substring(14, 2)), 0);
                }
                else
                {
                    ldt_DateFin = new DateTime(Convert.ToInt32(EntradaMigrarTag.fechafin.Substring(0, 4)), Convert.ToInt32(EntradaMigrarTag.fechafin.Substring(5, 2)), Convert.ToInt32(EntradaMigrarTag.fechafin.Substring(8, 2)),
                                               Convert.ToInt32(EntradaMigrarTag.fechafin.Substring(11, 2)), Convert.ToInt32(EntradaMigrarTag.fechafin.Substring(14, 2)), Convert.ToInt32(EntradaMigrarTag.fechafin.Substring(17, 2)));
                }
            }
            catch (Exception ex)
            {
                Funciones.CapturarError(ex, "Utilidades.svc - MigrarTag", "svcOperacionesPI", false);
                return new Resultado(-1, ex.Message);
            }

            try
            {
                PIServers oPIServers = new PIServers();

                //obtenemos servidor origen
                PIServer myPIServerOrigen = oPIServers[EntradaMigrarTag.ServidorOrigen];
                myPIServerOrigen.Connect();

                //obtenemos el PiPoint de origen
                PIPoint lo_PiPointOrigen;
                lo_PiPointOrigen = PIPoint.FindPIPoint(myPIServerOrigen, EntradaMigrarTag.TagOrigen);

                //obtenemos la lista de AFValues del origen a migrar
                AFTime oAFTimeIni = new AFTime(ldt_DateIni.ToUniversalTime());
                AFTime oAFTimeFin = new AFTime(ldt_DateFin.ToUniversalTime());
                AFTimeRange myRange = new AFTimeRange(oAFTimeIni, oAFTimeFin);

                AFValues lo_AFValuesOrigen = lo_PiPointOrigen.RecordedValues(myRange, AFBoundaryType.Inside, "", false);
                oAFValoresOrigen = lo_AFValuesOrigen.ToList();


                //obtenemos servidor destino
                PIServer myPIServerDestino = oPIServers[EntradaMigrarTag.ServidorDestino];
                myPIServerDestino.Connect();

                //Dictionary<String, Object> attributes  = new Dictionary<String, Object>();
                //attributes.Add(PICommonPointAttributes.PointType, "Float64");
                //attributes.Add(PICommonPointAttributes.Compressing, 0);
                //myPIServerDestino.CreatePIPoint(EntradaMigrarTag.TagDestino, attributes);

                //obtenemos el PiPoint de destino
                PIPoint lo_PiPointDestino;
                lo_PiPointDestino = PIPoint.FindPIPoint(myPIServerDestino, EntradaMigrarTag.TagDestino);

                //transferimos los valores al servidor/tag destino
                lo_PiPointDestino.UpdateValues(oAFValoresOrigen, OSIsoft.AF.Data.AFUpdateOption.Replace);

                return new Resultado(0, "Operación Exitosa");


            }
            catch (Exception ex)
            {
                Funciones.CapturarError(ex, "Utilidades.svc - MigrarTag", "svcOperacionesPI", false);
                return new Resultado(-1, ex.Message);
            }
        }


        public List<Tag> BuscarTags(string servidor, string query)
        {
            List<Tag> oTags = new List<Tag>();
            Tag oTag = new Tag();

            try
            {
                PIServers oPIServers = new PIServers();
                PIServer myPIServer = oPIServers[servidor];
                myPIServer.Connect();

                PIPointQuery oQuery = new PIPointQuery("tag", AFSearchOperator.Equal, query);

                List<PIPointQuery> oQuerys = new List<PIPointQuery>();
                oQuerys.Add(oQuery);

                List<PIPoint> oPoints = PIPoint.FindPIPoints(myPIServer, oQuerys).ToList();

                foreach (PIPoint oPoint in oPoints)
                {
                    oTag = new Tag();
                    oTag.Nombre = oPoint.Name;
                    oPoint.LoadAttributes(PICommonPointAttributes.Descriptor);
                    oTag.Descripcion = oPoint.GetAttribute(PICommonPointAttributes.Descriptor).ToString();
                    oTag.Tipo = oPoint.PointType.ToString();
                    oTag.Existe = true;

                    oTags.Add(oTag);
                }

                return oTags;
            }
            catch (Exception ex)
            {
                Funciones.CapturarError(ex, "Utilidades.svc - BuscarTags", "svcOperacionesPI", false);
                return oTags;
            }
        }



        //metodos utilidad importacion datos de excel a PI
        public List<HojaExcel> ListarHojasDeExcel(string url)
        {
            List<HojaExcel> oHojas = new List<HojaExcel>();
            HojaExcel oHoja = new HojaExcel();

            try
            {
                FileInfo existingFile = new FileInfo(url);

                using (ExcelPackage package = new ExcelPackage(existingFile))
                {
                    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                    foreach (ExcelWorksheet ews in package.Workbook.Worksheets)
                    {
                        oHoja = new HojaExcel();
                        oHoja.Indice = ews.Index;
                        oHoja.Nombre = ews.Name;

                        oHojas.Add(oHoja);
                    }
                }

                return oHojas;
            }
            catch (Exception ex)
            {
                Funciones.CapturarError(ex, "Utilidades.svc - ListarHojasDeExcel", "wsUNACEMPI", true);
                return new List<HojaExcel>();
            }
        }


        public Resultado ImportarExcelAPI(EntradaImportarExcelAPI EntradaImportarExcelAPI)
        {
            try
            {
                string cServidorPIAF = WebConfigurationManager.AppSettings["ServidorPIAF"];
                string cServidorPIData = WebConfigurationManager.AppSettings["ServidorPIData"];
                FileInfo existingFile = new FileInfo(EntradaImportarExcelAPI.UrlExcel);

                using (ExcelPackage package = new ExcelPackage(existingFile))
                {
                    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                    ExcelWorksheet ews = package.Workbook.Worksheets[EntradaImportarExcelAPI.HojaExcel];

                    foreach (ConfiguracionExcel oConfiguracion in EntradaImportarExcelAPI.Configuraciones)
                    {

                        //obtenemos el tag
                        string tag = oConfiguracion.Tag;
                        if (oConfiguracion.CeldaTag.Trim() != "")
                        {
                            tag = ews.Cells[oConfiguracion.CeldaTag.Trim()].Value.ToString();
                        }

                        //obtenemos servidor pi
                        PIServers oPIServers = new PIServers();
                        PIServer myPIServer = oPIServers[cServidorPIData];
                        myPIServer.Connect();

                        //obtenemos el PiPoint
                        PIPoint lo_PiPoint;
                        lo_PiPoint = PIPoint.FindPIPoint(myPIServer, tag);

                        int nFilaFin = oConfiguracion.FilaFin;
                        if (nFilaFin <= 0)
                        {
                            nFilaFin = ews.Dimension.Rows;

                        }


                        for (int i = oConfiguracion.FilaIni; i <= nFilaFin; i++)
                        {

                            DateTime dFecha;
                            try
                            {
                                dFecha = Convert.ToDateTime(ews.Cells[oConfiguracion.Fecha + i.ToString()].Value);
                                if (dFecha > DateTime.Now)
                                {
                                    //no grabamos datos futuros - salimos del bucle
                                    break;
                                }
                            }
                            catch (Exception ex2)
                            {
                                try
                                {
                                    dFecha = DateTime.FromOADate(Convert.ToDouble(ews.Cells[oConfiguracion.Fecha + i.ToString()].Value));
                                    if (dFecha > DateTime.Now)
                                    {
                                        //no grabamos datos futuros - salimos del bucle
                                        break;
                                    }
                                }
                                catch (Exception ex3)
                                {
                                    //no es fecha valida - salimos del bucle
                                    break;
                                }
                            }

                            string cFormula = ObtenerFormula(oConfiguracion.Data, i);



                            ews.Cells["XFD1"].Formula = cFormula;
                            ews.Cells["XFD1"].Calculate();
                            //package.Workbook.Calculate();
                            object oValor = ews.Cells["XFD1"].Value;


                            AFTime lo_AFTime = new AFTime(dFecha.ToUniversalTime());
                            AFValue lo_AFValue = new AFValue(oValor, lo_AFTime);

                            if (!EsValorNulo(ews, cFormula))
                            {
                                lo_PiPoint.UpdateValue(lo_AFValue, OSIsoft.AF.Data.AFUpdateOption.Replace);
                            }
                            else
                            {
                                //el valor resultado es nulo
                                if (oConfiguracion.FilaFin <= 0)
                                {
                                    //no se definio fila fin por lo que es automatico / salimos del bucle
                                    break;
                                }
                            }



                            ////verificamos si es un numero                            
                            //double nValor=0;
                            //if (Double.TryParse(oValor.ToString(), out nValor))
                            //{
                            //    AFTime lo_AFTime = new AFTime(dFecha.ToUniversalTime());
                            //    AFValue lo_AFValue = new AFValue(nValor, lo_AFTime);

                            //    lo_PiPoint.UpdateValue(lo_AFValue, OSIsoft.AF.Data.AFUpdateOption.Replace);
                            //}


                        }
                    }
                }

                return new Resultado(0, "Operación Exitosa");
            }
            catch (Exception ex)
            {
                Funciones.CapturarError(ex, "Utilidades.svc - ImportarExcelAPI", "wsUNACEMPI", true);
                return new Resultado(-1, ex.Message);
            }
        }


        public Resultado ImportarExcelAPI2(EntradaImportarExcelAPI2 EntradaImportarExcelAPI2)
        {
            try
            {
                string cServidorPIAF = WebConfigurationManager.AppSettings["ServidorPIAF"];
                string cServidorPIData = WebConfigurationManager.AppSettings["ServidorPIData"];
                FileInfo existingFile = new FileInfo(EntradaImportarExcelAPI2.UrlExcel);

                using (ExcelPackage package = new ExcelPackage(existingFile))
                {
                    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                    //validamos si tiene archivos externos asociados
                    //if (EntradaImportarExcelAPI2.ArchivosExternos != null)
                    //{
                    //    for (int i = 0; i < EntradaImportarExcelAPI2.ArchivosExternos.Count; i++)
                    //    {
                    //Add a reference to an external file.
                    //var externalLinkFile = new FileInfo(EntradaImportarExcelAPI2.ArchivosExternos[i]);
                    //var externalWorkbook = package.Workbook.  ExternalLinks.AddExternalWorkbook(externalLinkFile);
                    //    }
                    //}


                    ExcelWorksheet ews = package.Workbook.Worksheets[EntradaImportarExcelAPI2.HojaExcel];

                    foreach (ConfiguracionExcel oConfiguracion in EntradaImportarExcelAPI2.Configuraciones)
                    {

                        //obtenemos el tag
                        string tag = oConfiguracion.Tag;
                        if (oConfiguracion.CeldaTag.Trim() != "")
                        {
                            tag = ews.Cells[oConfiguracion.CeldaTag.Trim()].Value.ToString();
                        }
                   
                        //obtenemos servidor pi
                        PIServers oPIServers = new PIServers();
                        PIServer myPIServer = oPIServers[cServidorPIData];
                        myPIServer.Connect();

                        //obtenemos el PiPoint
                        PIPoint lo_PiPoint;
                        lo_PiPoint = PIPoint.FindPIPoint(myPIServer, tag);

                        int nFilaFin = oConfiguracion.FilaFin;
                        if (nFilaFin <= 0)
                        {
                            nFilaFin = ews.Dimension.Rows;
                        }


                        for (int i = oConfiguracion.FilaIni; i <= nFilaFin; i++)
                        {

                            DateTime dFecha;
                            try
                            {
                                dFecha = Convert.ToDateTime(ews.Cells[oConfiguracion.Fecha + i.ToString()].Value);
                                if (dFecha > DateTime.Now)
                                {
                                    //no grabamos datos futuros - salimos del bucle
                                    break;
                                }
                            }
                            catch (Exception ex2)
                            {
                                try
                                {
                                    dFecha = DateTime.FromOADate(Convert.ToDouble(ews.Cells[oConfiguracion.Fecha + i.ToString()].Value));
                                    if (dFecha > DateTime.Now)
                                    {
                                        //no grabamos datos futuros - salimos del bucle
                                        break;
                                    }
                                }
                                catch (Exception ex3)
                                {
                                    //no es fecha valida - salimos del bucle
                                    break;
                                }
                            }

                            string cFormula = ObtenerFormula(oConfiguracion.Data, i);



                            ews.Cells["XFD1"].Formula = cFormula;
                            ews.Cells["XFD1"].Calculate();
                            //package.Workbook.Calculate();
                            object oValor = ews.Cells["XFD1"].Value;


                            AFTime lo_AFTime = new AFTime(dFecha.ToUniversalTime());
                            AFValue lo_AFValue = new AFValue(oValor, lo_AFTime);

                            if (!EsValorNulo(ews, cFormula))
                            {
                                try
                                {
                                    lo_PiPoint.UpdateValue(lo_AFValue, OSIsoft.AF.Data.AFUpdateOption.Replace);
                                }
                                catch (PIException ExceptionD)
                                {
                                    //errores por intentar grabar datos futuros no detienen el proceso
                                    if (ExceptionD.StatusCode != -11046)
                                    {
                                        Funciones.CapturarError(ExceptionD, "Utilidades.svc - ImportarExcelAPI2", "svcOperacionesPI", true);
                                        return new Resultado(-1, ExceptionD.Message);
                                    }
                                }

                            }
                            else
                            {
                                //el valor resultado es nulo
                                if (oConfiguracion.FilaFin <= 0)
                                {
                                    //no se definio fila fin por lo que es automatico / salimos del bucle
                                    break;
                                }
                            }



                            ////verificamos si es un numero                            
                            //double nValor=0;
                            //if (Double.TryParse(oValor.ToString(), out nValor))
                            //{
                            //    AFTime lo_AFTime = new AFTime(dFecha.ToUniversalTime());
                            //    AFValue lo_AFValue = new AFValue(nValor, lo_AFTime);

                            //    lo_PiPoint.UpdateValue(lo_AFValue, OSIsoft.AF.Data.AFUpdateOption.Replace);
                            //}


                        }
                    }
                }


                //validamos si tiene analisis por ejecutar
                if (EntradaImportarExcelAPI2.Analisis != null)
                {
                    foreach (AnalisisPorEjecutar oAnalisis in EntradaImportarExcelAPI2.Analisis)
                    {
                        DateTime dFechaIni;
                        dFechaIni = new DateTime(Convert.ToInt32(oAnalisis.FechaIni.Substring(0, 4)), Convert.ToInt32(oAnalisis.FechaIni.Substring(5, 2)), Convert.ToInt32(oAnalisis.FechaIni.Substring(8, 2)));

                        DateTime dFechaFin;
                        dFechaFin = new DateTime(Convert.ToInt32(oAnalisis.FechaFin.Substring(0, 4)), Convert.ToInt32(oAnalisis.FechaFin.Substring(5, 2)), Convert.ToInt32(oAnalisis.FechaFin.Substring(8, 2)));

                        PISystems oPIAF = new PISystems();
                        PISystem oPI = oPIAF.DefaultPISystem;

                        oPI.Connect();

                        string cDatabase = oAnalisis.Database;
                        AFDatabase oDB = oPI.Databases[cDatabase];

                        AFAnalysisService aAnalysisService = null;
                        IEnumerable<AFAnalysis> elemAnalyses = null;
                        List<AFAnalysis> foundAnalyses = new List<AFAnalysis>();
                        Object response = null;

                        aAnalysisService = oPI.AnalysisService;
                        string user_analysisfilter = oAnalisis.NombreAnalisis;

                        AFTime oTiempoIni = new AFTime(dFechaIni.ToUniversalTime());
                        AFTime oTiempoFin = new AFTime(dFechaFin.ToUniversalTime());
                        AFTimeRange backfillPeriod = new AFTimeRange(oTiempoIni, oTiempoFin);

                        List<string> rutas = new List<string>();
                        rutas.Add(oAnalisis.RutaElemento);
                        AFKeyedResults<string, AFElement> oResultados = AFElement.FindElementsByPath(rutas, oDB);

                        if (oResultados.Count > 0)
                        {
                            AFElement oElementoTemp = oResultados[oAnalisis.RutaElemento];

                            String analysisfilter = "Target:=\"" + oElementoTemp.GetPath(oDB) + "\" Name:=\"" + user_analysisfilter + "\"";
                            AFAnalysisSearch analysisSearch = new AFAnalysisSearch(oDB, "analysisSearch", AFAnalysisSearch.ParseQuery(analysisfilter));

                            elemAnalyses = analysisSearch.FindAnalyses(0, true).ToList();
                            foundAnalyses.AddRange(elemAnalyses);

                            foreach (var analysis_n in foundAnalyses)
                            {
                                //response = aAnalysisService.QueueCalculation(new List<AFAnalysis> { analysis_n }, backfillPeriod, AFAnalysisService.CalculationMode.FillDataGaps);
                                response = aAnalysisService.QueueCalculation(new List<AFAnalysis> { analysis_n }, backfillPeriod, AFAnalysisService.CalculationMode.DeleteExistingData);
                            }
                        }





                    }
                }

                return new Resultado(0, "Operación Exitosa");
            }
            catch (Exception ex)
            {
                Funciones.CapturarError(ex, "Utilidades.svc - ImportarExcelAPI2", "svcOperacionesPI", true);
                return new Resultado(-1, ex.Message);
            }
        }





        //metodos utilidad importacion de excel configurados en AF
        public Resultado ImportarExcelDeAF(string nombre, string fecha, string usuario)
        {
            //SE REQUIERE INSTALAR EN EL SERVIDOR sharepointclientcomponents_16-6518-1200_x64-en-us
            //SE REQUIERE HAYA UNA CARPETA "DATA"

            string cRespuestas = "";

            string cServidorPIAF = WebConfigurationManager.AppSettings["ServidorPIAF"];
            string cServidorPIData = WebConfigurationManager.AppSettings["ServidorPIData"];

            string cURLSharepoint = WebConfigurationManager.AppSettings["URLSharepoint"];
            string cUsuarioSharepoint = WebConfigurationManager.AppSettings["UsuarioSharepoint"];
            string cClaveSharepoint = WebConfigurationManager.AppSettings["ClaveSharepoint"];

            string cURLExcel = "";
            string cTituloFecha = "";

            DateTime dFechaImportacion;
            try
            {
                dFechaImportacion = new DateTime(Convert.ToInt32(fecha.Substring(0, 4)), Convert.ToInt32(fecha.Substring(5, 2)), Convert.ToInt32(fecha.Substring(8, 2)));
            }
            catch (Exception ex)
            {
                Funciones.CapturarError(ex, "Utilidades.svc - ImportarExcelDeAF - La fecha ingresada no es valida", "wsUNACEMPI", false);
                return new Resultado(-1, "La fecha ingresada no es valida");
            }

            Funciones.CapturarMensaje("Mensaje: El usuario " + usuario + " inició el proceso de importación de datos del excel " + nombre + " a PI para el mes de " + ObtenerNombreMes(dFechaImportacion.Month) + "-" + dFechaImportacion.Year.ToString());

            try
            {
                // 1. CONECTAMOS CON SHAREPOINT

                System.Net.ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12;

                //Namespace: It belongs to Microsoft.SharePoint.Client
                ClientContext ctx = new ClientContext(cURLSharepoint);

                // Namespace: It belongs to System.Security
                SecureString secureString = new SecureString();
                cClaveSharepoint.ToList().ForEach(secureString.AppendChar);

                // Namespace: It belongs to Microsoft.SharePoint.Client
                ctx.Credentials = new SharePointOnlineCredentials(cUsuarioSharepoint, secureString);

                // Namespace: It belongs to Microsoft.SharePoint.Client
                Web web = ctx.Web;

                ctx.Load(web);
                ctx.ExecuteQuery();
                cRespuestas = cRespuestas + DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt") + " - Se conectó a Sharepoint \n";
                Funciones.CapturarMensaje("Mensaje: Se conectó a Sharepoint");




                // 2. BAJAMOS EL EXCEL

                string cDatabase = WebConfigurationManager.AppSettings["databaseExternalData"];

                PISystems oPIAF = new PISystems();
                PISystem oPI = oPIAF[cServidorPIAF];

                oPI.Connect();
                AFDatabase oDB = oPI.Databases[cDatabase];

                //buscamos el elemento excel
                var oSearch = new AFElementSearch(oDB, "Buscar", @"Root:'' Name:='" + nombre + "' TemplateName:='Excel-Anual'");
                List<AFElement> oElementosExcel = oSearch.FindElements(fullLoad: true).ToList();
                if (oElementosExcel.Count == 0)
                {
                    return new Resultado(-1, "No se encontro el elemento de template Excel con nombre " + nombre);
                }

                AFElement oAFExcel = oElementosExcel[0];
                cURLExcel = oAFExcel.Attributes["Link"].GetValue().Value.ToString().Trim();
                cTituloFecha = oAFExcel.Attributes["Fecha"].GetValue().Value.ToString().Trim();

                Uri filename = new Uri(cURLExcel);
                string server = filename.AbsoluteUri.Replace(filename.AbsolutePath, "");
                string serverrelative = filename.AbsolutePath;

                Microsoft.SharePoint.Client.FileInformation f = Microsoft.SharePoint.Client.File.OpenBinaryDirect(ctx, serverrelative);

                ctx.ExecuteQuery();

                using (var fileStream = new FileStream(AppDomain.CurrentDomain.BaseDirectory + @"\data\" + nombre + ".xlsx", FileMode.Create))
                    f.Stream.CopyTo(fileStream);

                cRespuestas = cRespuestas + DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt") + " - Se bajó el excel " + nombre + " desde Sharepoint  \n";
                Funciones.CapturarMensaje("Mensaje: Se bajó el excel " + nombre + " desde Sharepoint");



                // 3. IMPORTAMOS DATOS A PI DESDE EL EXCEL
                string cMes = ObtenerNombreMes(dFechaImportacion.Month) + "_" + dFechaImportacion.Year.ToString();
                FileInfo existingFile = new FileInfo(AppDomain.CurrentDomain.BaseDirectory + @"\data\" + nombre + ".xlsx");

                using (ExcelPackage package = new ExcelPackage(existingFile))
                {
                    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[cMes];

                    int filas = worksheet.Dimension.Rows;
                    int columnas = worksheet.Dimension.Columns;

                    int filaFecha = 0;
                    int columnaFecha = 0;


                    //buscamos columna fecha en el excel
                    for (int i = 1; i <= filas; i++)
                    {
                        for (int j = 1; j <= columnas; j++)
                        {
                            object valor = worksheet.Cells[i, j].Value;
                            if (valor != null && valor.ToString().Trim().ToUpper() == cTituloFecha.ToUpper())
                            {
                                filaFecha = i;
                                columnaFecha = j;
                                break;
                            }
                        }

                        if (filaFecha != 0)
                        {
                            break;
                        }
                    }
                    if (filaFecha == 0)
                    {
                        return new Resultado(-1, "No se encontro la columna de fecha con título: " + cTituloFecha);
                    }


                    //buscamos elementos columnas AF del excel
                    string cRutaAFExcel = oAFExcel.GetPath();
                    cRutaAFExcel = cRutaAFExcel.Replace(@"\\" + oPI.Name + @"\" + cDatabase + @"\", "");

                    var oSearch2 = new AFElementSearch(oDB, "Buscar", @"Root:'" + cRutaAFExcel + "' TemplateName:='Material excel-Anual'");
                    foreach (AFElement oElementoTemp in oSearch2.FindElements(fullLoad: true))
                    {
                        int columnaCabecera = 0;
                        int columnaSubCabecera = 0;
                        string cTituloCabecera = "";
                        string cTituloSubCabecera = "";
                        int nMetodoBusqueda = 0;

                        try
                        {
                            cTituloCabecera = oElementoTemp.Attributes["Cabecera"].GetValue().Value.ToString().Trim();
                            cTituloSubCabecera = oElementoTemp.Attributes["TITULO"].GetValue().Value.ToString().Trim();
                            nMetodoBusqueda = oElementoTemp.Attributes["Metodo busqueda"].GetValue().ValueAsInt32();

                            string cTag = "";
                            if (oElementoTemp.Attributes["XLS"].DataReference != null)
                            {
                                cTag = oElementoTemp.Attributes["XLS"].DataReference.ToString();
                                cTag = cTag.Substring(cTag.LastIndexOf(@"\") + 1);
                            }

                            //buscamos la columna de la cabecera en el excel
                            for (int j = 1; j <= columnas; j++)
                            {
                                object valor = worksheet.Cells[filaFecha, j].Value;
                                if (valor != null && valor.ToString().Trim().Replace("\n", " ").ToUpper() == cTituloCabecera.ToUpper())
                                {
                                    columnaCabecera = j;
                                    break;
                                }
                            }
                            if (columnaCabecera == 0)
                            {
                                cRespuestas = cRespuestas + DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt") + " - No se encontro la columna de la cabecera: " + cTituloCabecera + "  \n";
                                Funciones.CapturarMensaje("No se encontro la columna de la cabecera: " + cTituloCabecera);
                                continue;
                            }

                            // buscamos la columna de la sub-cabecera de en el excel
                            for (int j = columnaCabecera; j <= columnas; j++)
                            {
                                object valor = worksheet.Cells[filaFecha + 1, j].Value;
                                if (nMetodoBusqueda == 0)
                                {
                                    if (valor != null && valor.ToString().Trim().Replace("\n", " ").ToUpper() == cTituloSubCabecera.ToUpper())
                                    {
                                        columnaSubCabecera = j;
                                        break;
                                    }
                                }
                                else
                                {
                                    if (valor != null && valor.ToString().Trim().Replace("\n", " ").ToUpper().StartsWith(cTituloSubCabecera.ToUpper()))
                                    {
                                        columnaSubCabecera = j;
                                        break;
                                    }
                                }
                            }
                            if (columnaSubCabecera == 0)
                            {
                                if (nMetodoBusqueda == 0)
                                {
                                    cRespuestas = cRespuestas + DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt") + " - No se encontro la columna de la sub-cabecera: " + cTituloSubCabecera + "  \n";
                                    Funciones.CapturarMensaje("No se encontro la columna de la sub-cabecera: " + cTituloSubCabecera);
                                    continue;
                                }
                                else
                                {
                                    cRespuestas = cRespuestas + DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt") + " - No se encontro la columna de la sub-cabecera que empieza con: " + cTituloSubCabecera + "  \n";
                                    Funciones.CapturarMensaje("No se encontro la columna de la sub-cabecera que empieza con: " + cTituloSubCabecera);
                                    continue;
                                }
                            }

                            //importamos datos
                            for (int i = filaFecha + 2; i <= filas; i++)
                            {
                                object oValorFecha = worksheet.Cells[i, columnaFecha].Value;

                                if (oValorFecha != null && oValorFecha.ToString().Trim().ToUpper() != "TOTALES")
                                {
                                    DateTime dFecha = new DateTime();
                                    bool bFechaValida = false;
                                    try
                                    {
                                        dFecha = DateTime.FromOADate(Convert.ToDouble(oValorFecha));
                                        if (dFechaImportacion.Year == dFecha.Year && dFechaImportacion.Month == dFecha.Month && dFecha < DateTime.Now)
                                        {
                                            bFechaValida = true;
                                        }
                                    }
                                    catch (Exception ex3)
                                    {

                                    }

                                    if (bFechaValida)
                                    {
                                        object oValorDato = worksheet.Cells[i, columnaSubCabecera].Value;

                                        double nDato = 0;
                                        if (oValorDato != null)
                                        {
                                            if (Double.TryParse(oValorDato.ToString(), out nDato))
                                            {
                                                bool rpta = ModPIExtFunctions.Write_Value_PI(cServidorPIData, cTag, dFecha, nDato, true);
                                                if (rpta == false)
                                                {
                                                    ModPIExtFunctions.Write_Value_PI(cServidorPIData, cTag, dFecha, nDato, true);
                                                }
                                            }
                                        }
                                    }

                                }
                                else
                                {
                                    break;
                                }
                            }

                            cRespuestas = cRespuestas + DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt") + " - Se importó datos de " + cTituloCabecera + " - " + cTituloSubCabecera + "  \n";
                            Funciones.CapturarMensaje("Mensaje: Se importó datos de " + cTituloCabecera + " - " + cTituloSubCabecera);
                        }
                        catch (Exception ex2)
                        {
                            cRespuestas = cRespuestas + DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt") + " - No se pudo importar datos de " + cTituloCabecera + " - " + cTituloSubCabecera + ". " + ex2.Message + "  \n";
                            Funciones.CapturarError(ex2, "Utilidades.svc - ImportarExcelDeAF", "wsUNACEMPI", false);
                        }



                    }


                }





                return new Resultado(0, cRespuestas);
            }
            catch (Exception ex)
            {
                Funciones.CapturarError(ex, "Utilidades.svc - ImportarExcelDeAF", "wsUNACEMPI", false);
                return new Resultado(-1, ex.Message);
            }
        }


        public Resultado ImportarTodoExcelDeAF(string fecha, string usuario)
        {
            string cRespuestas = "";

            string cServidorPIAF = WebConfigurationManager.AppSettings["ServidorPIAF"];
            string cServidorPIData = WebConfigurationManager.AppSettings["ServidorPIData"];

            DateTime dFechaImportacion;
            try
            {
                dFechaImportacion = new DateTime(Convert.ToInt32(fecha.Substring(0, 4)), Convert.ToInt32(fecha.Substring(5, 2)), Convert.ToInt32(fecha.Substring(8, 2)));
            }
            catch (Exception ex)
            {
                Funciones.CapturarError(ex, "Utilidades.svc - ImportarTodoExcelDeAF - La fecha ingresada no es valida", "wsUNACEMPI", false);
                return new Resultado(-1, "La fecha ingresada no es valida");
            }

            Funciones.CapturarMensaje("Mensaje: El usuario " + usuario + " inició el proceso de importación de archivos de excel de tipo \"Anual\" a PI para el mes de " + ObtenerNombreMes(dFechaImportacion.Month) + "-" + dFechaImportacion.Year.ToString());

            try
            {
                string cDatabase = WebConfigurationManager.AppSettings["databaseExternalData"];

                PISystems oPIAF = new PISystems();
                PISystem oPI = oPIAF[cServidorPIAF];

                oPI.Connect();
                AFDatabase oDB = oPI.Databases[cDatabase];

                //buscamos el elemento excel
                var oSearch = new AFElementSearch(oDB, "Buscar", @"Root:'' TemplateName:='Excel-Anual'");
                List<AFElement> oElementosExcel = oSearch.FindElements(fullLoad: true).ToList();
                foreach (AFElement oElementoExcel in oElementosExcel)
                {
                    string cNombre = oElementoExcel.Name;
                    Resultado oResultadoExcel = ImportarExcelDeAF(cNombre, fecha, usuario);
                    cRespuestas = cRespuestas + oResultadoExcel.descripcion;
                }

                return new Resultado(0, cRespuestas);
            }
            catch (Exception ex)
            {
                Funciones.CapturarError(ex, "Utilidades.svc - ImportarTodoExcelDeAF", "wsUNACEMPI", false);
                return new Resultado(-1, ex.Message);
            }
        }


        public List<ExcelAF> ListarExcelDeAF(string template)
        {
            List<ExcelAF> oDatos = new List<ExcelAF>();
            ExcelAF oDato = new ExcelAF();

            try
            {
                string cServidorPIAF = WebConfigurationManager.AppSettings["ServidorPIAF"];
                string cServidorPIData = WebConfigurationManager.AppSettings["ServidorPIData"];
                string cDatabase = WebConfigurationManager.AppSettings["databaseExternalData"];

                PISystems oPIAF = new PISystems();
                PISystem oPI = oPIAF[cServidorPIAF];

                oPI.Connect();
                AFDatabase oDB = oPI.Databases[cDatabase];

                //buscamos el elemento excel
                var oSearch = new AFElementSearch(oDB, "Buscar", @"Root:'' TemplateName:='" + template + "'");
                List<AFElement> oElementosExcel = oSearch.FindElements(fullLoad: true).ToList();
                foreach (AFElement oElementoExcel in oElementosExcel)
                {
                    oDato = new ExcelAF();
                    oDato.nombre = oElementoExcel.Name;
                    oDatos.Add(oDato);
                }

                return oDatos;
            }
            catch (Exception ex)
            {
                Funciones.CapturarError(ex, "Utilidades.svc - ListarExcelDeAF", "wsUNACEMPI", false);
                return new List<ExcelAF>();
            }
        }


        public Resultado ImportarExcelMensualDeAF(string nombre, string fecha, string usuario)
        {
            //SE REQUIERE INSTALAR EN EL SERVIDOR sharepointclientcomponents_16-6518-1200_x64-en-us
            //SE REQUIERE HAYA UNA CARPETA "DATA"

            string cRespuestas = "";

            string cServidorPIAF = WebConfigurationManager.AppSettings["ServidorPIAF"];
            string cServidorPIData = WebConfigurationManager.AppSettings["ServidorPIData"];

            string cURLSharepoint = WebConfigurationManager.AppSettings["URLSharepoint"];
            string cUsuarioSharepoint = WebConfigurationManager.AppSettings["UsuarioSharepoint"];
            string cClaveSharepoint = WebConfigurationManager.AppSettings["ClaveSharepoint"];

            string cURLExcel = "";
            string cTituloHora = "";
            int nFilasEntreCabeceraYTitulo = 0;
            int nFilaInicioDatos = 0;
            int nTipoUbicacion = 0;

            DateTime dFechaImportacion;
            try
            {
                dFechaImportacion = new DateTime(Convert.ToInt32(fecha.Substring(0, 4)), Convert.ToInt32(fecha.Substring(5, 2)), Convert.ToInt32(fecha.Substring(8, 2)));
            }
            catch (Exception ex)
            {
                Funciones.CapturarError(ex, "Utilidades.svc - ImportarExcelMensualDeAF - La fecha ingresada no es valida", "wsUNACEMPI", false);
                return new Resultado(-1, "La fecha ingresada no es valida");
            }

            Funciones.CapturarMensaje("Mensaje: El usuario " + usuario + " inició el proceso de importación de datos del excel " + nombre + " a PI para el dia " + fecha);

            try
            {

                // 1. BUSCAMOS CONFIGURACION DE EXCEL EN AF

                string cDatabase = WebConfigurationManager.AppSettings["databaseExternalData"];

                PISystems oPIAF = new PISystems();
                PISystem oPI = oPIAF[cServidorPIAF];

                oPI.Connect();
                AFDatabase oDB = oPI.Databases[cDatabase];

                //buscamos el elemento excel
                var oSearch = new AFElementSearch(oDB, "Buscar", @"Root:'' Name:='" + nombre + "' TemplateName:='Excel-Mensual'");
                List<AFElement> oElementosExcel = oSearch.FindElements(fullLoad: true).ToList();
                if (oElementosExcel.Count == 0)
                {
                    return new Resultado(-1, "No se encontro el elemento de template Excel-Mensual con nombre " + nombre);
                }

                AFElement oAFExcel = oElementosExcel[0];
                nTipoUbicacion = oAFExcel.Attributes["Tipo Ubicacion"].GetValue().ValueAsInt32();

                string cLinkCarpeta = oAFExcel.Attributes["Link Carpeta"].GetValue().Value.ToString().Trim();


                string cNombreArchivo = oAFExcel.Attributes["Nombre Archivo"].GetValue().Value.ToString().Trim();
                cNombreArchivo = cNombreArchivo + " " + ObtenerNombreMesCompleto(dFechaImportacion.Month) + dFechaImportacion.Year + ".xlsx";

                string cNombreArchivo2 = oAFExcel.Attributes["Nombre Archivo"].GetValue().Value.ToString().Trim();
                cNombreArchivo2 = cNombreArchivo2 + "-" + ObtenerNombreMesCompleto(dFechaImportacion.Month) + dFechaImportacion.Year + ".xlsx";

                string cNombreArchivo3 = oAFExcel.Attributes["Nombre Archivo"].GetValue().Value.ToString().Trim();
                cNombreArchivo3 = cNombreArchivo3 + "" + ObtenerNombreMesCompleto(dFechaImportacion.Month) + dFechaImportacion.Year + ".xlsx";


                string cNombreArchivo4 = oAFExcel.Attributes["Nombre Archivo"].GetValue().Value.ToString().Trim();
                cNombreArchivo4 = cNombreArchivo4 + " " + ObtenerNombreMesCompleto(dFechaImportacion.Month) + "-" + dFechaImportacion.Year + ".xlsx";

                string cNombreArchivo5 = oAFExcel.Attributes["Nombre Archivo"].GetValue().Value.ToString().Trim();
                cNombreArchivo5 = cNombreArchivo5 + "-" + ObtenerNombreMesCompleto(dFechaImportacion.Month) + "-" + dFechaImportacion.Year + ".xlsx";

                string cNombreArchivo6 = oAFExcel.Attributes["Nombre Archivo"].GetValue().Value.ToString().Trim();
                cNombreArchivo6 = cNombreArchivo6 + "" + ObtenerNombreMesCompleto(dFechaImportacion.Month) + "-" + dFechaImportacion.Year + ".xlsx";


                string cNombreArchivo7 = oAFExcel.Attributes["Nombre Archivo"].GetValue().Value.ToString().Trim();
                cNombreArchivo7 = cNombreArchivo7 + " " + ObtenerNombreMesCompleto(dFechaImportacion.Month) + " " + dFechaImportacion.Year + ".xlsx";

                string cNombreArchivo8 = oAFExcel.Attributes["Nombre Archivo"].GetValue().Value.ToString().Trim();
                cNombreArchivo8 = cNombreArchivo8 + "-" + ObtenerNombreMesCompleto(dFechaImportacion.Month) + " " + dFechaImportacion.Year + ".xlsx";

                string cNombreArchivo9 = oAFExcel.Attributes["Nombre Archivo"].GetValue().Value.ToString().Trim();
                cNombreArchivo9 = cNombreArchivo9 + "" + ObtenerNombreMesCompleto(dFechaImportacion.Month) + " " + dFechaImportacion.Year + ".xlsx";




                if (nTipoUbicacion == 0) //sharepoint
                    cURLExcel = cLinkCarpeta + "/" + cNombreArchivo;
                else  //red
                    cURLExcel = cLinkCarpeta + @"\" + cNombreArchivo;

                cTituloHora = oAFExcel.Attributes["Hora"].GetValue().Value.ToString().Trim();
                nFilaInicioDatos = oAFExcel.Attributes["Fila Inicio Datos"].GetValue().ValueAsInt32();



                if (nTipoUbicacion == 0)    //sharepoint
                {
                    // 1. CONECTAMOS CON SHAREPOINT

                    System.Net.ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12;

                    //Namespace: It belongs to Microsoft.SharePoint.Client
                    ClientContext ctx = new ClientContext(cURLSharepoint);

                    // Namespace: It belongs to System.Security
                    SecureString secureString = new SecureString();
                    cClaveSharepoint.ToList().ForEach(secureString.AppendChar);

                    // Namespace: It belongs to Microsoft.SharePoint.Client
                    ctx.Credentials = new SharePointOnlineCredentials(cUsuarioSharepoint, secureString);

                    // Namespace: It belongs to Microsoft.SharePoint.Client
                    Web web = ctx.Web;

                    ctx.Load(web);
                    ctx.ExecuteQuery();
                    cRespuestas = cRespuestas + DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt") + " - Se conectó a Sharepoint \n";
                    Funciones.CapturarMensaje("Mensaje: Se conectó a Sharepoint");



                    // 2. BAJAMOS EL EXCEL
                    Uri filename = new Uri(cURLExcel);
                    string server = filename.AbsoluteUri.Replace(filename.AbsolutePath, "");
                    string serverrelative = filename.AbsolutePath;

                    Microsoft.SharePoint.Client.FileInformation f = Microsoft.SharePoint.Client.File.OpenBinaryDirect(ctx, serverrelative);

                    ctx.ExecuteQuery();

                    using (var fileStream = new FileStream(AppDomain.CurrentDomain.BaseDirectory + @"\data\" + nombre + ".xlsx", FileMode.Create))
                        f.Stream.CopyTo(fileStream);

                    cRespuestas = cRespuestas + DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt") + " - Se bajó el excel " + nombre + " desde Sharepoint  \n";
                    Funciones.CapturarMensaje("Mensaje: Se bajó el excel " + nombre + " desde Sharepoint");

                }
                else    //red
                {
                    // 2. BAJAMOS EL EXCEL
                    FileInfo originalFile = new FileInfo(cURLExcel);
                    if (originalFile.Exists)
                    {
                        originalFile.CopyTo(AppDomain.CurrentDomain.BaseDirectory + @"\data\" + nombre + ".xlsx", true);

                        cRespuestas = cRespuestas + DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt") + " - Se bajó el excel " + nombre + " desde la Red  \n";
                        Funciones.CapturarMensaje("Mensaje: Se bajó el excel " + nombre + " desde la Red");
                    }
                    else
                    {
                        FileInfo originalFile2 = new FileInfo(cLinkCarpeta + @"\" + cNombreArchivo2);
                        if (originalFile2.Exists)
                        {
                            originalFile2.CopyTo(AppDomain.CurrentDomain.BaseDirectory + @"\data\" + nombre + ".xlsx", true);

                            cRespuestas = cRespuestas + DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt") + " - Se bajó el excel " + nombre + " desde la Red  \n";
                            Funciones.CapturarMensaje("Mensaje: Se bajó el excel " + nombre + " desde la Red");
                        }
                        else
                        {
                            FileInfo originalFile3 = new FileInfo(cLinkCarpeta + @"\" + cNombreArchivo3);
                            if (originalFile3.Exists)
                            {
                                originalFile3.CopyTo(AppDomain.CurrentDomain.BaseDirectory + @"\data\" + nombre + ".xlsx", true);

                                cRespuestas = cRespuestas + DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt") + " - Se bajó el excel " + nombre + " desde la Red  \n";
                                Funciones.CapturarMensaje("Mensaje: Se bajó el excel " + nombre + " desde la Red");
                            }
                            else
                            {
                                FileInfo originalFile4 = new FileInfo(cLinkCarpeta + @"\" + cNombreArchivo4);
                                if (originalFile4.Exists)
                                {
                                    originalFile4.CopyTo(AppDomain.CurrentDomain.BaseDirectory + @"\data\" + nombre + ".xlsx", true);

                                    cRespuestas = cRespuestas + DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt") + " - Se bajó el excel " + nombre + " desde la Red  \n";
                                    Funciones.CapturarMensaje("Mensaje: Se bajó el excel " + nombre + " desde la Red");
                                }
                                else
                                {
                                    FileInfo originalFile5 = new FileInfo(cLinkCarpeta + @"\" + cNombreArchivo5);
                                    if (originalFile5.Exists)
                                    {
                                        originalFile5.CopyTo(AppDomain.CurrentDomain.BaseDirectory + @"\data\" + nombre + ".xlsx", true);

                                        cRespuestas = cRespuestas + DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt") + " - Se bajó el excel " + nombre + " desde la Red  \n";
                                        Funciones.CapturarMensaje("Mensaje: Se bajó el excel " + nombre + " desde la Red");
                                    }
                                    else
                                    {
                                        FileInfo originalFile6 = new FileInfo(cLinkCarpeta + @"\" + cNombreArchivo6);
                                        if (originalFile6.Exists)
                                        {
                                            originalFile6.CopyTo(AppDomain.CurrentDomain.BaseDirectory + @"\data\" + nombre + ".xlsx", true);

                                            cRespuestas = cRespuestas + DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt") + " - Se bajó el excel " + nombre + " desde la Red  \n";
                                            Funciones.CapturarMensaje("Mensaje: Se bajó el excel " + nombre + " desde la Red");
                                        }
                                        else
                                        {
                                            FileInfo originalFile7 = new FileInfo(cLinkCarpeta + @"\" + cNombreArchivo7);
                                            if (originalFile7.Exists)
                                            {
                                                originalFile7.CopyTo(AppDomain.CurrentDomain.BaseDirectory + @"\data\" + nombre + ".xlsx", true);

                                                cRespuestas = cRespuestas + DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt") + " - Se bajó el excel " + nombre + " desde la Red  \n";
                                                Funciones.CapturarMensaje("Mensaje: Se bajó el excel " + nombre + " desde la Red");
                                            }
                                            else
                                            {
                                                FileInfo originalFile8 = new FileInfo(cLinkCarpeta + @"\" + cNombreArchivo8);
                                                if (originalFile8.Exists)
                                                {
                                                    originalFile8.CopyTo(AppDomain.CurrentDomain.BaseDirectory + @"\data\" + nombre + ".xlsx", true);

                                                    cRespuestas = cRespuestas + DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt") + " - Se bajó el excel " + nombre + " desde la Red  \n";
                                                    Funciones.CapturarMensaje("Mensaje: Se bajó el excel " + nombre + " desde la Red");
                                                }
                                                else
                                                {
                                                    FileInfo originalFile9 = new FileInfo(cLinkCarpeta + @"\" + cNombreArchivo9);
                                                    if (originalFile9.Exists)
                                                    {
                                                        originalFile9.CopyTo(AppDomain.CurrentDomain.BaseDirectory + @"\data\" + nombre + ".xlsx", true);

                                                        cRespuestas = cRespuestas + DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt") + " - Se bajó el excel " + nombre + " desde la Red  \n";
                                                        Funciones.CapturarMensaje("Mensaje: Se bajó el excel " + nombre + " desde la Red");
                                                    }
                                                    else
                                                    {
                                                        Funciones.CapturarMensaje("El archivo excel para \"" + nombre + "\" no existe o no tiene permisos suficientes para acceder a el");
                                                        return new Resultado(-1, "El archivo excel para \"" + nombre + "\" no existe o no tiene permisos suficientes para acceder a el");
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }

                }






                // 3. IMPORTAMOS DATOS A PI DESDE EL EXCEL
                string cDia = dFechaImportacion.Day < 10 ? "0" + dFechaImportacion.Day.ToString() : dFechaImportacion.Day.ToString();
                FileInfo existingFile = new FileInfo(AppDomain.CurrentDomain.BaseDirectory + @"\data\" + nombre + ".xlsx");

                using (ExcelPackage package = new ExcelPackage(existingFile))
                {
                    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[cDia];

                    int filas = worksheet.Dimension.Rows;
                    int columnas = worksheet.Dimension.Columns;

                    int filaHora = 0;
                    int columnaHora = 0;


                    //buscamos columna fecha en el excel
                    for (int i = 1; i <= filas; i++)
                    {
                        for (int j = 1; j <= columnas; j++)
                        {
                            object valor = worksheet.Cells[i, j].Value;
                            if (valor != null && valor.ToString().Trim().ToUpper() == cTituloHora.ToUpper())
                            {
                                filaHora = i;
                                columnaHora = j;
                                break;
                            }
                        }

                        if (filaHora != 0)
                        {
                            break;
                        }
                    }
                    if (filaHora == 0)
                    {
                        return new Resultado(-1, "No se encontro la columna de hora con título: " + cTituloHora);
                    }


                    //buscamos elementos columnas AF del excel
                    string cRutaAFExcel = oAFExcel.GetPath();
                    cRutaAFExcel = cRutaAFExcel.Replace(@"\\" + oPI.Name + @"\" + cDatabase + @"\", "");

                    var oSearch2 = new AFElementSearch(oDB, "Buscar", @"Root:'" + cRutaAFExcel + "' TemplateName:='Material excel-Mensual'");
                    foreach (AFElement oElementoTemp in oSearch2.FindElements(fullLoad: true))
                    {
                        int columnaCabecera = 0;
                        int columnaSubCabecera = 0;
                        string cTituloCabecera = "";
                        string cTituloSubCabecera = "";
                        int nMetodoBusqueda = 0;

                        try
                        {
                            cTituloCabecera = oElementoTemp.Attributes["Cabecera"].GetValue().Value.ToString().Trim();
                            cTituloSubCabecera = oElementoTemp.Attributes["TITULO"].GetValue().Value.ToString().Trim();
                            nMetodoBusqueda = oElementoTemp.Attributes["Metodo busqueda"].GetValue().ValueAsInt32();
                            nFilasEntreCabeceraYTitulo = oElementoTemp.Attributes["FilasEntreCabeceraYTitulo"].GetValue().ValueAsInt32();


                            string cTag = "";
                            if (oElementoTemp.Attributes["XLS"].DataReference != null)
                            {
                                cTag = oElementoTemp.Attributes["XLS"].DataReference.ToString();
                                cTag = cTag.Substring(cTag.LastIndexOf(@"\") + 1);
                            }

                            //buscamos la columna de la cabecera en el excel
                            for (int j = 1; j <= columnas; j++)
                            {
                                object valor = worksheet.Cells[filaHora, j].Value;
                                if (valor != null && valor.ToString().Trim().Replace("\n", " ").ToUpper() == cTituloCabecera.ToUpper())
                                {
                                    columnaCabecera = j;
                                    break;
                                }
                            }
                            if (columnaCabecera == 0)
                            {
                                cRespuestas = cRespuestas + DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt") + " - No se encontro la columna de la cabecera: " + cTituloCabecera + "  \n";
                                Funciones.CapturarMensaje("No se encontro la columna de la cabecera: " + cTituloCabecera);
                                continue;
                            }

                            // buscamos la columna de la sub-cabecera de en el excel
                            for (int j = columnaCabecera; j <= columnas; j++)
                            {
                                object valor = worksheet.Cells[filaHora + nFilasEntreCabeceraYTitulo + 1, j].Value;
                                if (nMetodoBusqueda == 0)
                                {
                                    if (valor != null && valor.ToString().Trim().Replace("\n", " ").ToUpper() == cTituloSubCabecera.ToUpper())
                                    {
                                        columnaSubCabecera = j;
                                        break;
                                    }
                                }
                                else
                                {
                                    if (valor != null && valor.ToString().Trim().Replace("\n", " ").ToUpper().StartsWith(cTituloSubCabecera.ToUpper()))
                                    {
                                        columnaSubCabecera = j;
                                        break;
                                    }
                                }
                            }
                            if (columnaSubCabecera == 0)
                            {
                                if (nMetodoBusqueda == 0)
                                {
                                    cRespuestas = cRespuestas + DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt") + " - No se encontro la columna de la sub-cabecera: " + cTituloSubCabecera + "  \n";
                                    Funciones.CapturarMensaje("No se encontro la columna de la sub-cabecera: " + cTituloSubCabecera);
                                    continue;
                                }
                                else
                                {
                                    cRespuestas = cRespuestas + DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt") + " - No se encontro la columna de la sub-cabecera que empieza con: " + cTituloSubCabecera + "  \n";
                                    Funciones.CapturarMensaje("No se encontro la columna de la sub-cabecera que empieza con: " + cTituloSubCabecera);
                                    continue;
                                }
                            }

                            //importamos datos
                            for (int i = nFilaInicioDatos; i <= filas; i++)
                            {
                                object oValorHora = worksheet.Cells[i, columnaHora].Value;

                                if (oValorHora != null && oValorHora.ToString().Trim().ToUpper() != "TOTALES")
                                {
                                    DateTime dFecha = new DateTime();
                                    bool bFechaValida = false;
                                    try
                                    {
                                        dFecha = Convert.ToDateTime(oValorHora);
                                        dFecha = dFechaImportacion.AddHours(dFecha.Hour).AddMinutes(dFecha.Minute);

                                        if (dFecha < DateTime.Now)
                                        {
                                            bFechaValida = true;
                                        }

                                        //dFecha = Convert.ToDateTime(dFechaImportacion.ToString("yyyy-MM-dd") + " " + oValorHora.ToString());
                                        //bFechaValida = true;
                                    }
                                    catch (Exception ex3)
                                    {

                                    }

                                    if (bFechaValida)
                                    {
                                        object oValorDato = worksheet.Cells[i, columnaSubCabecera].Value;

                                        double nDato = 0;
                                        if (oValorDato != null)
                                        {
                                            if (Double.TryParse(oValorDato.ToString(), out nDato))
                                            {
                                                bool rpta = ModPIExtFunctions.Write_Value_PI(cServidorPIData, cTag, dFecha, nDato, true);
                                                if (rpta == false)
                                                {
                                                    ModPIExtFunctions.Write_Value_PI(cServidorPIData, cTag, dFecha, nDato, true);
                                                }
                                            }
                                        }
                                    }

                                }
                                else
                                {
                                    break;
                                }
                            }

                            cRespuestas = cRespuestas + DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt") + " - Se importó datos de " + cTituloCabecera + " - " + cTituloSubCabecera + "  \n";
                            Funciones.CapturarMensaje("Mensaje: Se importó datos de " + cTituloCabecera + " - " + cTituloSubCabecera);
                        }
                        catch (Exception ex2)
                        {
                            cRespuestas = cRespuestas + DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt") + " - No se pudo importar datos de " + cTituloCabecera + " - " + cTituloSubCabecera + ". " + ex2.Message + "  \n";
                            Funciones.CapturarError(ex2, "Utilidades.svc - ImportarExcelMensualDeAF", "wsUNACEMPI", false);
                        }



                    }


                }





                return new Resultado(0, cRespuestas);
            }
            catch (Exception ex)
            {
                Funciones.CapturarError(ex, "Utilidades.svc - ImportarExcelMensualDeAF", "wsUNACEMPI", false);
                return new Resultado(-1, ex.Message);
            }
        }


        public Resultado ImportarExcelMensualDeAFCompleto(string nombre, string fecha, string usuario)
        {
            //SE REQUIERE INSTALAR EN EL SERVIDOR sharepointclientcomponents_16-6518-1200_x64-en-us
            //SE REQUIERE HAYA UNA CARPETA "DATA"

            string cRespuestas = "";

            string cServidorPIAF = WebConfigurationManager.AppSettings["ServidorPIAF"];
            string cServidorPIData = WebConfigurationManager.AppSettings["ServidorPIData"];

            string cURLSharepoint = WebConfigurationManager.AppSettings["URLSharepoint"];
            string cUsuarioSharepoint = WebConfigurationManager.AppSettings["UsuarioSharepoint"];
            string cClaveSharepoint = WebConfigurationManager.AppSettings["ClaveSharepoint"];

            string cURLExcel = "";
            string cTituloHora = "";
            int nFilasEntreCabeceraYTitulo = 0;
            int nFilaInicioDatos = 0;
            int nTipoUbicacion = 0;

            DateTime dFechaImportacion;
            try
            {
                dFechaImportacion = new DateTime(Convert.ToInt32(fecha.Substring(0, 4)), Convert.ToInt32(fecha.Substring(5, 2)), Convert.ToInt32(fecha.Substring(8, 2)));
            }
            catch (Exception ex)
            {
                Funciones.CapturarError(ex, "Utilidades.svc - ImportarExcelMensualDeAFCompleto - La fecha ingresada no es valida", "wsUNACEMPI", false);
                return new Resultado(-1, "La fecha ingresada no es valida");
            }

            Funciones.CapturarMensaje("Mensaje: El usuario " + usuario + " inició el proceso de importación de datos del excel " + nombre + " a PI para el mes de " + ObtenerNombreMesCompleto(dFechaImportacion.Month));

            try
            {

                // 1. BUSCAMOS CONFIGURACION DE EXCEL EN AF

                string cDatabase = WebConfigurationManager.AppSettings["databaseExternalData"];

                PISystems oPIAF = new PISystems();
                PISystem oPI = oPIAF[cServidorPIAF];

                oPI.Connect();
                AFDatabase oDB = oPI.Databases[cDatabase];

                //buscamos el elemento excel
                var oSearch = new AFElementSearch(oDB, "Buscar", @"Root:'' Name:='" + nombre + "' TemplateName:='Excel-Mensual'");
                List<AFElement> oElementosExcel = oSearch.FindElements(fullLoad: true).ToList();
                if (oElementosExcel.Count == 0)
                {
                    return new Resultado(-1, "No se encontro el elemento de template Excel-Mensual con nombre " + nombre);
                }

                AFElement oAFExcel = oElementosExcel[0];
                nTipoUbicacion = oAFExcel.Attributes["Tipo Ubicacion"].GetValue().ValueAsInt32();

                string cLinkCarpeta = oAFExcel.Attributes["Link Carpeta"].GetValue().Value.ToString().Trim();


                string cNombreArchivo = oAFExcel.Attributes["Nombre Archivo"].GetValue().Value.ToString().Trim();
                cNombreArchivo = cNombreArchivo + " " + ObtenerNombreMesCompleto(dFechaImportacion.Month) + dFechaImportacion.Year + ".xlsx";

                string cNombreArchivo2 = oAFExcel.Attributes["Nombre Archivo"].GetValue().Value.ToString().Trim();
                cNombreArchivo2 = cNombreArchivo2 + "-" + ObtenerNombreMesCompleto(dFechaImportacion.Month) + dFechaImportacion.Year + ".xlsx";

                string cNombreArchivo3 = oAFExcel.Attributes["Nombre Archivo"].GetValue().Value.ToString().Trim();
                cNombreArchivo3 = cNombreArchivo3 + "" + ObtenerNombreMesCompleto(dFechaImportacion.Month) + dFechaImportacion.Year + ".xlsx";


                string cNombreArchivo4 = oAFExcel.Attributes["Nombre Archivo"].GetValue().Value.ToString().Trim();
                cNombreArchivo4 = cNombreArchivo4 + " " + ObtenerNombreMesCompleto(dFechaImportacion.Month) + "-" + dFechaImportacion.Year + ".xlsx";

                string cNombreArchivo5 = oAFExcel.Attributes["Nombre Archivo"].GetValue().Value.ToString().Trim();
                cNombreArchivo5 = cNombreArchivo5 + "-" + ObtenerNombreMesCompleto(dFechaImportacion.Month) + "-" + dFechaImportacion.Year + ".xlsx";

                string cNombreArchivo6 = oAFExcel.Attributes["Nombre Archivo"].GetValue().Value.ToString().Trim();
                cNombreArchivo6 = cNombreArchivo6 + "" + ObtenerNombreMesCompleto(dFechaImportacion.Month) + "-" + dFechaImportacion.Year + ".xlsx";


                string cNombreArchivo7 = oAFExcel.Attributes["Nombre Archivo"].GetValue().Value.ToString().Trim();
                cNombreArchivo7 = cNombreArchivo7 + " " + ObtenerNombreMesCompleto(dFechaImportacion.Month) + " " + dFechaImportacion.Year + ".xlsx";

                string cNombreArchivo8 = oAFExcel.Attributes["Nombre Archivo"].GetValue().Value.ToString().Trim();
                cNombreArchivo8 = cNombreArchivo8 + "-" + ObtenerNombreMesCompleto(dFechaImportacion.Month) + " " + dFechaImportacion.Year + ".xlsx";

                string cNombreArchivo9 = oAFExcel.Attributes["Nombre Archivo"].GetValue().Value.ToString().Trim();
                cNombreArchivo9 = cNombreArchivo9 + "" + ObtenerNombreMesCompleto(dFechaImportacion.Month) + " " + dFechaImportacion.Year + ".xlsx";



                if (nTipoUbicacion == 0) //sharepoint
                    cURLExcel = cLinkCarpeta + "/" + cNombreArchivo;
                else  //red
                    cURLExcel = cLinkCarpeta + @"\" + cNombreArchivo;

                cTituloHora = oAFExcel.Attributes["Hora"].GetValue().Value.ToString().Trim();
                nFilaInicioDatos = oAFExcel.Attributes["Fila Inicio Datos"].GetValue().ValueAsInt32();



                if (nTipoUbicacion == 0)    //sharepoint
                {
                    // 1. CONECTAMOS CON SHAREPOINT

                    System.Net.ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12;

                    //Namespace: It belongs to Microsoft.SharePoint.Client
                    ClientContext ctx = new ClientContext(cURLSharepoint);

                    // Namespace: It belongs to System.Security
                    SecureString secureString = new SecureString();
                    cClaveSharepoint.ToList().ForEach(secureString.AppendChar);

                    // Namespace: It belongs to Microsoft.SharePoint.Client
                    ctx.Credentials = new SharePointOnlineCredentials(cUsuarioSharepoint, secureString);

                    // Namespace: It belongs to Microsoft.SharePoint.Client
                    Web web = ctx.Web;

                    ctx.Load(web);
                    ctx.ExecuteQuery();
                    cRespuestas = cRespuestas + DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt") + " - Se conectó a Sharepoint \n";
                    Funciones.CapturarMensaje("Mensaje: Se conectó a Sharepoint");



                    // 2. BAJAMOS EL EXCEL
                    Uri filename = new Uri(cURLExcel);
                    string server = filename.AbsoluteUri.Replace(filename.AbsolutePath, "");
                    string serverrelative = filename.AbsolutePath;

                    Microsoft.SharePoint.Client.FileInformation f = Microsoft.SharePoint.Client.File.OpenBinaryDirect(ctx, serverrelative);

                    ctx.ExecuteQuery();

                    using (var fileStream = new FileStream(AppDomain.CurrentDomain.BaseDirectory + @"\data\" + nombre + ".xlsx", FileMode.Create))
                        f.Stream.CopyTo(fileStream);

                    cRespuestas = cRespuestas + DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt") + " - Se bajó el excel " + nombre + " desde Sharepoint  \n";
                    Funciones.CapturarMensaje("Mensaje: Se bajó el excel " + nombre + " desde Sharepoint");

                }
                else    //red
                {
                    // 2. BAJAMOS EL EXCEL
                    FileInfo originalFile = new FileInfo(cURLExcel);
                    if (originalFile.Exists)
                    {
                        originalFile.CopyTo(AppDomain.CurrentDomain.BaseDirectory + @"\data\" + nombre + ".xlsx", true);

                        cRespuestas = cRespuestas + DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt") + " - Se bajó el excel " + nombre + " desde la Red  \n";
                        Funciones.CapturarMensaje("Mensaje: Se bajó el excel " + nombre + " desde la Red");
                    }
                    else
                    {
                        FileInfo originalFile2 = new FileInfo(cLinkCarpeta + @"\" + cNombreArchivo2);
                        if (originalFile2.Exists)
                        {
                            originalFile2.CopyTo(AppDomain.CurrentDomain.BaseDirectory + @"\data\" + nombre + ".xlsx", true);

                            cRespuestas = cRespuestas + DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt") + " - Se bajó el excel " + nombre + " desde la Red  \n";
                            Funciones.CapturarMensaje("Mensaje: Se bajó el excel " + nombre + " desde la Red");
                        }
                        else
                        {
                            FileInfo originalFile3 = new FileInfo(cLinkCarpeta + @"\" + cNombreArchivo3);
                            if (originalFile3.Exists)
                            {
                                originalFile3.CopyTo(AppDomain.CurrentDomain.BaseDirectory + @"\data\" + nombre + ".xlsx", true);

                                cRespuestas = cRespuestas + DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt") + " - Se bajó el excel " + nombre + " desde la Red  \n";
                                Funciones.CapturarMensaje("Mensaje: Se bajó el excel " + nombre + " desde la Red");
                            }
                            else
                            {
                                FileInfo originalFile4 = new FileInfo(cLinkCarpeta + @"\" + cNombreArchivo4);
                                if (originalFile4.Exists)
                                {
                                    originalFile4.CopyTo(AppDomain.CurrentDomain.BaseDirectory + @"\data\" + nombre + ".xlsx", true);

                                    cRespuestas = cRespuestas + DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt") + " - Se bajó el excel " + nombre + " desde la Red  \n";
                                    Funciones.CapturarMensaje("Mensaje: Se bajó el excel " + nombre + " desde la Red");
                                }
                                else
                                {
                                    FileInfo originalFile5 = new FileInfo(cLinkCarpeta + @"\" + cNombreArchivo5);
                                    if (originalFile5.Exists)
                                    {
                                        originalFile5.CopyTo(AppDomain.CurrentDomain.BaseDirectory + @"\data\" + nombre + ".xlsx", true);

                                        cRespuestas = cRespuestas + DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt") + " - Se bajó el excel " + nombre + " desde la Red  \n";
                                        Funciones.CapturarMensaje("Mensaje: Se bajó el excel " + nombre + " desde la Red");
                                    }
                                    else
                                    {
                                        FileInfo originalFile6 = new FileInfo(cLinkCarpeta + @"\" + cNombreArchivo6);
                                        if (originalFile6.Exists)
                                        {
                                            originalFile6.CopyTo(AppDomain.CurrentDomain.BaseDirectory + @"\data\" + nombre + ".xlsx", true);

                                            cRespuestas = cRespuestas + DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt") + " - Se bajó el excel " + nombre + " desde la Red  \n";
                                            Funciones.CapturarMensaje("Mensaje: Se bajó el excel " + nombre + " desde la Red");
                                        }
                                        else
                                        {
                                            FileInfo originalFile7 = new FileInfo(cLinkCarpeta + @"\" + cNombreArchivo7);
                                            if (originalFile7.Exists)
                                            {
                                                originalFile7.CopyTo(AppDomain.CurrentDomain.BaseDirectory + @"\data\" + nombre + ".xlsx", true);

                                                cRespuestas = cRespuestas + DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt") + " - Se bajó el excel " + nombre + " desde la Red  \n";
                                                Funciones.CapturarMensaje("Mensaje: Se bajó el excel " + nombre + " desde la Red");
                                            }
                                            else
                                            {
                                                FileInfo originalFile8 = new FileInfo(cLinkCarpeta + @"\" + cNombreArchivo8);
                                                if (originalFile8.Exists)
                                                {
                                                    originalFile8.CopyTo(AppDomain.CurrentDomain.BaseDirectory + @"\data\" + nombre + ".xlsx", true);

                                                    cRespuestas = cRespuestas + DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt") + " - Se bajó el excel " + nombre + " desde la Red  \n";
                                                    Funciones.CapturarMensaje("Mensaje: Se bajó el excel " + nombre + " desde la Red");
                                                }
                                                else
                                                {
                                                    FileInfo originalFile9 = new FileInfo(cLinkCarpeta + @"\" + cNombreArchivo9);
                                                    if (originalFile9.Exists)
                                                    {
                                                        originalFile9.CopyTo(AppDomain.CurrentDomain.BaseDirectory + @"\data\" + nombre + ".xlsx", true);

                                                        cRespuestas = cRespuestas + DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt") + " - Se bajó el excel " + nombre + " desde la Red  \n";
                                                        Funciones.CapturarMensaje("Mensaje: Se bajó el excel " + nombre + " desde la Red");
                                                    }
                                                    else
                                                    {
                                                        Funciones.CapturarMensaje("El archivo excel para \"" + nombre + "\" no existe o no tiene permisos suficientes para acceder a el");
                                                        return new Resultado(-1, "El archivo excel para \"" + nombre + "\" no existe o no tiene permisos suficientes para acceder a el");
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }






                // 3. IMPORTAMOS DATOS A PI DESDE EL EXCEL
                DateTime dFechaIni = new DateTime(dFechaImportacion.Year, dFechaImportacion.Month, 1);
                DateTime dFechaFin = dFechaIni.AddMonths(1).AddDays(-1);
                DateTime dFechaTemp = dFechaIni;

                while (dFechaTemp <= dFechaFin)
                {

                    string cDia = dFechaTemp.Day < 10 ? "0" + dFechaTemp.Day.ToString() : dFechaTemp.Day.ToString();
                    FileInfo existingFile = new FileInfo(AppDomain.CurrentDomain.BaseDirectory + @"\data\" + nombre + ".xlsx");

                    using (ExcelPackage package = new ExcelPackage(existingFile))
                    {
                        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                        ExcelWorksheet worksheet = package.Workbook.Worksheets[cDia];

                        int filas = worksheet.Dimension.Rows;
                        int columnas = worksheet.Dimension.Columns;

                        int filaHora = 0;
                        int columnaHora = 0;


                        //buscamos columna fecha en el excel
                        for (int i = 1; i <= filas; i++)
                        {
                            for (int j = 1; j <= columnas; j++)
                            {
                                object valor = worksheet.Cells[i, j].Value;
                                if (valor != null && valor.ToString().Trim().ToUpper() == cTituloHora.ToUpper())
                                {
                                    filaHora = i;
                                    columnaHora = j;
                                    break;
                                }
                            }

                            if (filaHora != 0)
                            {
                                break;
                            }
                        }
                        if (filaHora == 0)
                        {
                            return new Resultado(-1, "No se encontro la columna de hora con título: " + cTituloHora);
                        }


                        //buscamos elementos columnas AF del excel
                        string cRutaAFExcel = oAFExcel.GetPath();
                        cRutaAFExcel = cRutaAFExcel.Replace(@"\\" + oPI.Name + @"\" + cDatabase + @"\", "");

                        var oSearch2 = new AFElementSearch(oDB, "Buscar", @"Root:'" + cRutaAFExcel + "' TemplateName:='Material excel-Mensual'");
                        foreach (AFElement oElementoTemp in oSearch2.FindElements(fullLoad: true))
                        {
                            int columnaCabecera = 0;
                            int columnaSubCabecera = 0;
                            string cTituloCabecera = "";
                            string cTituloSubCabecera = "";
                            int nMetodoBusqueda = 0;

                            try
                            {
                                cTituloCabecera = oElementoTemp.Attributes["Cabecera"].GetValue().Value.ToString().Trim();
                                cTituloSubCabecera = oElementoTemp.Attributes["TITULO"].GetValue().Value.ToString().Trim();
                                nMetodoBusqueda = oElementoTemp.Attributes["Metodo busqueda"].GetValue().ValueAsInt32();
                                nFilasEntreCabeceraYTitulo = oElementoTemp.Attributes["FilasEntreCabeceraYTitulo"].GetValue().ValueAsInt32();


                                string cTag = "";
                                if (oElementoTemp.Attributes["XLS"].DataReference != null)
                                {
                                    cTag = oElementoTemp.Attributes["XLS"].DataReference.ToString();
                                    cTag = cTag.Substring(cTag.LastIndexOf(@"\") + 1);
                                }

                                //buscamos la columna de la cabecera en el excel
                                for (int j = 1; j <= columnas; j++)
                                {
                                    object valor = worksheet.Cells[filaHora, j].Value;
                                    if (valor != null && valor.ToString().Trim().Replace("\n", " ").ToUpper() == cTituloCabecera.ToUpper())
                                    {
                                        columnaCabecera = j;
                                        break;
                                    }
                                }
                                if (columnaCabecera == 0)
                                {
                                    cRespuestas = cRespuestas + DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt") + " - No se encontro la columna de la cabecera: " + cTituloCabecera + "  \n";
                                    Funciones.CapturarMensaje("No se encontro la columna de la cabecera: " + cTituloCabecera);
                                    continue;
                                }

                                // buscamos la columna de la sub-cabecera de en el excel
                                for (int j = columnaCabecera; j <= columnas; j++)
                                {
                                    object valor = worksheet.Cells[filaHora + nFilasEntreCabeceraYTitulo + 1, j].Value;
                                    if (nMetodoBusqueda == 0)
                                    {
                                        if (valor != null && valor.ToString().Trim().Replace("\n", " ").ToUpper() == cTituloSubCabecera.ToUpper())
                                        {
                                            columnaSubCabecera = j;
                                            break;
                                        }
                                    }
                                    else
                                    {
                                        if (valor != null && valor.ToString().Trim().Replace("\n", " ").ToUpper().StartsWith(cTituloSubCabecera.ToUpper()))
                                        {
                                            columnaSubCabecera = j;
                                            break;
                                        }
                                    }
                                }
                                if (columnaSubCabecera == 0)
                                {
                                    if (nMetodoBusqueda == 0)
                                    {
                                        cRespuestas = cRespuestas + DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt") + " - No se encontro la columna de la sub-cabecera: " + cTituloSubCabecera + "  \n";
                                        Funciones.CapturarMensaje("No se encontro la columna de la sub-cabecera: " + cTituloSubCabecera);
                                        continue;
                                    }
                                    else
                                    {
                                        cRespuestas = cRespuestas + DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt") + " - No se encontro la columna de la sub-cabecera que empieza con: " + cTituloSubCabecera + "  \n";
                                        Funciones.CapturarMensaje("No se encontro la columna de la sub-cabecera que empieza con: " + cTituloSubCabecera);
                                        continue;
                                    }
                                }

                                //importamos datos
                                for (int i = nFilaInicioDatos; i <= filas; i++)
                                {
                                    object oValorHora = worksheet.Cells[i, columnaHora].Value;

                                    if (oValorHora != null && oValorHora.ToString().Trim().ToUpper() != "TOTALES")
                                    {
                                        DateTime dFecha = new DateTime();
                                        bool bFechaValida = false;
                                        try
                                        {
                                            dFecha = Convert.ToDateTime(oValorHora);
                                            dFecha = dFechaTemp.AddHours(dFecha.Hour).AddMinutes(dFecha.Minute);

                                            if (dFecha < DateTime.Now)
                                            {
                                                bFechaValida = true;
                                            }

                                            //dFecha = Convert.ToDateTime(dFechaImportacion.ToString("yyyy-MM-dd") + " " + oValorHora.ToString());
                                            //bFechaValida = true;
                                        }
                                        catch (Exception ex3)
                                        {

                                        }

                                        if (bFechaValida)
                                        {
                                            object oValorDato = worksheet.Cells[i, columnaSubCabecera].Value;

                                            double nDato = 0;
                                            if (oValorDato != null)
                                            {
                                                if (Double.TryParse(oValorDato.ToString(), out nDato))
                                                {
                                                    bool rpta = ModPIExtFunctions.Write_Value_PI(cServidorPIData, cTag, dFecha, nDato, true);
                                                    if (rpta == false)
                                                    {
                                                        ModPIExtFunctions.Write_Value_PI(cServidorPIData, cTag, dFecha, nDato, true);
                                                    }
                                                }
                                            }
                                        }

                                    }
                                    else
                                    {
                                        break;
                                    }
                                }

                                cRespuestas = cRespuestas + DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt") + " - Se importó datos de " + cTituloCabecera + " - " + cTituloSubCabecera + " para el " + dFechaTemp.ToString("yyyy-MM-dd") + "  \n";
                                Funciones.CapturarMensaje("Mensaje: Se importó datos de " + cTituloCabecera + " - " + cTituloSubCabecera);
                            }
                            catch (Exception ex2)
                            {
                                cRespuestas = cRespuestas + DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt") + " - No se pudo importar datos de " + cTituloCabecera + " - " + cTituloSubCabecera + " para el " + dFechaTemp.ToString("yyyy-MM-dd") + ". " + ex2.Message + "  \n";
                                Funciones.CapturarError(ex2, "Utilidades.svc - ImportarExcelMensualDeAF", "wsUNACEMPI", false);
                            }



                        }


                    }






                    dFechaTemp = dFechaTemp.AddDays(1);
                }








                return new Resultado(0, cRespuestas);
            }
            catch (Exception ex)
            {
                Funciones.CapturarError(ex, "Utilidades.svc - ImportarExcelMensualDeAF", "wsUNACEMPI", false);
                return new Resultado(-1, ex.Message);
            }
        }


        public Resultado ImportarTodoExcelDeAFMensual(string fecha, string usuario)
        {
            string cRespuestas = "";

            string cServidorPIAF = WebConfigurationManager.AppSettings["ServidorPIAF"];
            string cServidorPIData = WebConfigurationManager.AppSettings["ServidorPIData"];

            DateTime dFechaImportacion;
            try
            {
                dFechaImportacion = new DateTime(Convert.ToInt32(fecha.Substring(0, 4)), Convert.ToInt32(fecha.Substring(5, 2)), Convert.ToInt32(fecha.Substring(8, 2)));
            }
            catch (Exception ex)
            {
                Funciones.CapturarError(ex, "Utilidades.svc - ImportarTodoExcelDeAFMensual - La fecha ingresada no es valida", "wsUNACEMPI", false);
                return new Resultado(-1, "La fecha ingresada no es valida");
            }

            Funciones.CapturarMensaje("Mensaje: El usuario " + usuario + " inició el proceso de importación de archivos de excel de tipo \"Mensual\" a PI para el mes de " + ObtenerNombreMes(dFechaImportacion.Month) + "-" + dFechaImportacion.Year.ToString());

            try
            {
                string cDatabase = WebConfigurationManager.AppSettings["databaseExternalData"];

                PISystems oPIAF = new PISystems();
                PISystem oPI = oPIAF[cServidorPIAF];

                oPI.Connect();
                AFDatabase oDB = oPI.Databases[cDatabase];

                //buscamos el elemento excel
                var oSearch = new AFElementSearch(oDB, "Buscar", @"Root:'' TemplateName:='Excel-Mensual'");
                List<AFElement> oElementosExcel = oSearch.FindElements(fullLoad: true).ToList();
                foreach (AFElement oElementoExcel in oElementosExcel)
                {
                    string cNombre = oElementoExcel.Name;
                    Resultado oResultadoExcel = ImportarExcelMensualDeAFCompleto(cNombre, fecha, usuario);
                    cRespuestas = cRespuestas + oResultadoExcel.descripcion;
                }

                return new Resultado(0, cRespuestas);
            }
            catch (Exception ex)
            {
                Funciones.CapturarError(ex, "Utilidades.svc - ImportarTodoExcelDeAFMensual", "wsUNACEMPI", false);
                return new Resultado(-1, ex.Message);
            }
        }


        public Resultado ImportarExcelMensualGenericoDeAF(string nombre, string fecha, string usuario)
        {
            //SE REQUIERE INSTALAR EN EL SERVIDOR sharepointclientcomponents_16-6518-1200_x64-en-us
            //SE REQUIERE HAYA UNA CARPETA "DATA"

            string cRespuestas = "";

            string cServidorPIAF = WebConfigurationManager.AppSettings["ServidorPIAF"];
            string cServidorPIData = WebConfigurationManager.AppSettings["ServidorPIData"];

            string cURLSharepoint = WebConfigurationManager.AppSettings["URLSharepoint"];
            string cUsuarioSharepoint = WebConfigurationManager.AppSettings["UsuarioSharepoint"];
            string cClaveSharepoint = WebConfigurationManager.AppSettings["ClaveSharepoint"];

            string cURLExcel = "";
            int nTipoUbicacion = 0;
            string cColumnaHora = "";
            string cExtension = "";

            DateTime dFechaImportacion;
            try
            {
                dFechaImportacion = new DateTime(Convert.ToInt32(fecha.Substring(0, 4)), Convert.ToInt32(fecha.Substring(5, 2)), Convert.ToInt32(fecha.Substring(8, 2)));
            }
            catch (Exception ex)
            {
                Funciones.CapturarError(ex, "Utilidades.svc - ImportarExcelMensualGenericoDeAF - La fecha ingresada no es valida", "wsUNACEMPI", false);
                return new Resultado(-1, "La fecha ingresada no es valida");
            }

            Funciones.CapturarMensaje("Mensaje: El usuario " + usuario + " inició el proceso de importación de datos del excel mensual genérico " + nombre + " a PI para el dia " + fecha);

            try
            {
                // 1. BUSCAMOS CONFIGURACION DE EXCEL EN AF

                string cDatabase = WebConfigurationManager.AppSettings["databaseExternalData"];

                PISystems oPIAF = new PISystems();
                PISystem oPI = oPIAF[cServidorPIAF];

                oPI.Connect();
                AFDatabase oDB = oPI.Databases[cDatabase];

                //buscamos el elemento excel
                var oSearch = new AFElementSearch(oDB, "Buscar", @"Root:'' Name:='" + nombre + "' TemplateName:='Excel-Mensual-Generico'");
                List<AFElement> oElementosExcel = oSearch.FindElements(fullLoad: true).ToList();
                if (oElementosExcel.Count == 0)
                {
                    return new Resultado(-1, "No se encontro el elemento de template Excel-Mensual-Generico con nombre " + nombre);
                }

                AFElement oAFExcel = oElementosExcel[0];
                nTipoUbicacion = oAFExcel.Attributes["Tipo Ubicacion"].GetValue().ValueAsInt32();

                string cLinkCarpeta = oAFExcel.Attributes["Link Carpeta"].GetValue().Value.ToString().Trim();
                cExtension = oAFExcel.Attributes["Extension"].GetValue().Value.ToString().Trim();

                string cNombreArchivo = oAFExcel.Attributes["Nombre Archivo"].GetValue().Value.ToString().Trim();
                cNombreArchivo = cNombreArchivo + " " + ObtenerNombreMesCompleto(dFechaImportacion.Month) + dFechaImportacion.Year + "." + cExtension;

                string cNombreArchivo2 = oAFExcel.Attributes["Nombre Archivo"].GetValue().Value.ToString().Trim();
                cNombreArchivo2 = cNombreArchivo2 + "-" + ObtenerNombreMesCompleto(dFechaImportacion.Month) + dFechaImportacion.Year + "." + cExtension;

                string cNombreArchivo3 = oAFExcel.Attributes["Nombre Archivo"].GetValue().Value.ToString().Trim();
                cNombreArchivo3 = cNombreArchivo3 + "" + ObtenerNombreMesCompleto(dFechaImportacion.Month) + dFechaImportacion.Year + "." + cExtension;


                string cNombreArchivo4 = oAFExcel.Attributes["Nombre Archivo"].GetValue().Value.ToString().Trim();
                cNombreArchivo4 = cNombreArchivo4 + " " + ObtenerNombreMesCompleto(dFechaImportacion.Month) + "-" + dFechaImportacion.Year + "." + cExtension;

                string cNombreArchivo5 = oAFExcel.Attributes["Nombre Archivo"].GetValue().Value.ToString().Trim();
                cNombreArchivo5 = cNombreArchivo5 + "-" + ObtenerNombreMesCompleto(dFechaImportacion.Month) + "-" + dFechaImportacion.Year + "." + cExtension;

                string cNombreArchivo6 = oAFExcel.Attributes["Nombre Archivo"].GetValue().Value.ToString().Trim();
                cNombreArchivo6 = cNombreArchivo6 + "" + ObtenerNombreMesCompleto(dFechaImportacion.Month) + "-" + dFechaImportacion.Year + "." + cExtension;


                string cNombreArchivo7 = oAFExcel.Attributes["Nombre Archivo"].GetValue().Value.ToString().Trim();
                cNombreArchivo7 = cNombreArchivo7 + " " + ObtenerNombreMesCompleto(dFechaImportacion.Month) + " " + dFechaImportacion.Year + "." + cExtension;

                string cNombreArchivo8 = oAFExcel.Attributes["Nombre Archivo"].GetValue().Value.ToString().Trim();
                cNombreArchivo8 = cNombreArchivo8 + "-" + ObtenerNombreMesCompleto(dFechaImportacion.Month) + " " + dFechaImportacion.Year + "." + cExtension;

                string cNombreArchivo9 = oAFExcel.Attributes["Nombre Archivo"].GetValue().Value.ToString().Trim();
                cNombreArchivo9 = cNombreArchivo9 + "" + ObtenerNombreMesCompleto(dFechaImportacion.Month) + " " + dFechaImportacion.Year + "." + cExtension;




                if (nTipoUbicacion == 0) //sharepoint
                    cURLExcel = cLinkCarpeta + "/" + cNombreArchivo;
                else  //red
                    cURLExcel = cLinkCarpeta + @"\" + cNombreArchivo;

                cColumnaHora = oAFExcel.Attributes["Columna Hora"].GetValue().Value.ToString().Trim();





                if (nTipoUbicacion == 0)    //sharepoint
                {
                    // 1. CONECTAMOS CON SHAREPOINT

                    System.Net.ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12;

                    //Namespace: It belongs to Microsoft.SharePoint.Client
                    ClientContext ctx = new ClientContext(cURLSharepoint);

                    // Namespace: It belongs to System.Security
                    SecureString secureString = new SecureString();
                    cClaveSharepoint.ToList().ForEach(secureString.AppendChar);

                    // Namespace: It belongs to Microsoft.SharePoint.Client
                    ctx.Credentials = new SharePointOnlineCredentials(cUsuarioSharepoint, secureString);

                    // Namespace: It belongs to Microsoft.SharePoint.Client
                    Web web = ctx.Web;

                    ctx.Load(web);
                    ctx.ExecuteQuery();
                    cRespuestas = cRespuestas + DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt") + " - Se conectó a Sharepoint \n";
                    Funciones.CapturarMensaje("Mensaje: Se conectó a Sharepoint");



                    // 2. BAJAMOS EL EXCEL
                    Uri filename = new Uri(cURLExcel);
                    string server = filename.AbsoluteUri.Replace(filename.AbsolutePath, "");
                    string serverrelative = filename.AbsolutePath;

                    Microsoft.SharePoint.Client.FileInformation f = Microsoft.SharePoint.Client.File.OpenBinaryDirect(ctx, serverrelative);

                    ctx.ExecuteQuery();

                    using (var fileStream = new FileStream(AppDomain.CurrentDomain.BaseDirectory + @"\data\" + nombre + "." + cExtension, FileMode.Create))
                        f.Stream.CopyTo(fileStream);

                    cRespuestas = cRespuestas + DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt") + " - Se bajó el excel " + nombre + " desde Sharepoint  \n";
                    Funciones.CapturarMensaje("Mensaje: Se bajó el excel " + nombre + " desde Sharepoint");

                }
                else    //red
                {
                    // 2. BAJAMOS EL EXCEL
                    FileInfo originalFile = new FileInfo(cURLExcel);
                    if (originalFile.Exists)
                    {
                        originalFile.CopyTo(AppDomain.CurrentDomain.BaseDirectory + @"\data\" + nombre + "." + cExtension, true);

                        cRespuestas = cRespuestas + DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt") + " - Se bajó el excel " + nombre + " desde la Red  \n";
                        Funciones.CapturarMensaje("Mensaje: Se bajó el excel " + nombre + " desde la Red");
                    }
                    else
                    {
                        FileInfo originalFile2 = new FileInfo(cLinkCarpeta + @"\" + cNombreArchivo2);
                        if (originalFile2.Exists)
                        {
                            originalFile2.CopyTo(AppDomain.CurrentDomain.BaseDirectory + @"\data\" + nombre + "." + cExtension, true);

                            cRespuestas = cRespuestas + DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt") + " - Se bajó el excel " + nombre + " desde la Red  \n";
                            Funciones.CapturarMensaje("Mensaje: Se bajó el excel " + nombre + " desde la Red");
                        }
                        else
                        {
                            FileInfo originalFile3 = new FileInfo(cLinkCarpeta + @"\" + cNombreArchivo3);
                            if (originalFile3.Exists)
                            {
                                originalFile3.CopyTo(AppDomain.CurrentDomain.BaseDirectory + @"\data\" + nombre + "." + cExtension, true);

                                cRespuestas = cRespuestas + DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt") + " - Se bajó el excel " + nombre + " desde la Red  \n";
                                Funciones.CapturarMensaje("Mensaje: Se bajó el excel " + nombre + " desde la Red");
                            }
                            else
                            {
                                FileInfo originalFile4 = new FileInfo(cLinkCarpeta + @"\" + cNombreArchivo4);
                                if (originalFile4.Exists)
                                {
                                    originalFile4.CopyTo(AppDomain.CurrentDomain.BaseDirectory + @"\data\" + nombre + "." + cExtension, true);

                                    cRespuestas = cRespuestas + DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt") + " - Se bajó el excel " + nombre + " desde la Red  \n";
                                    Funciones.CapturarMensaje("Mensaje: Se bajó el excel " + nombre + " desde la Red");
                                }
                                else
                                {
                                    FileInfo originalFile5 = new FileInfo(cLinkCarpeta + @"\" + cNombreArchivo5);
                                    if (originalFile5.Exists)
                                    {
                                        originalFile5.CopyTo(AppDomain.CurrentDomain.BaseDirectory + @"\data\" + nombre + "." + cExtension, true);

                                        cRespuestas = cRespuestas + DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt") + " - Se bajó el excel " + nombre + " desde la Red  \n";
                                        Funciones.CapturarMensaje("Mensaje: Se bajó el excel " + nombre + " desde la Red");
                                    }
                                    else
                                    {
                                        FileInfo originalFile6 = new FileInfo(cLinkCarpeta + @"\" + cNombreArchivo6);
                                        if (originalFile6.Exists)
                                        {
                                            originalFile6.CopyTo(AppDomain.CurrentDomain.BaseDirectory + @"\data\" + nombre + "." + cExtension, true);

                                            cRespuestas = cRespuestas + DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt") + " - Se bajó el excel " + nombre + " desde la Red  \n";
                                            Funciones.CapturarMensaje("Mensaje: Se bajó el excel " + nombre + " desde la Red");
                                        }
                                        else
                                        {
                                            FileInfo originalFile7 = new FileInfo(cLinkCarpeta + @"\" + cNombreArchivo7);
                                            if (originalFile7.Exists)
                                            {
                                                originalFile7.CopyTo(AppDomain.CurrentDomain.BaseDirectory + @"\data\" + nombre + "." + cExtension, true);

                                                cRespuestas = cRespuestas + DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt") + " - Se bajó el excel " + nombre + " desde la Red  \n";
                                                Funciones.CapturarMensaje("Mensaje: Se bajó el excel " + nombre + " desde la Red");
                                            }
                                            else
                                            {
                                                FileInfo originalFile8 = new FileInfo(cLinkCarpeta + @"\" + cNombreArchivo8);
                                                if (originalFile8.Exists)
                                                {
                                                    originalFile8.CopyTo(AppDomain.CurrentDomain.BaseDirectory + @"\data\" + nombre + "." + cExtension, true);

                                                    cRespuestas = cRespuestas + DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt") + " - Se bajó el excel " + nombre + " desde la Red  \n";
                                                    Funciones.CapturarMensaje("Mensaje: Se bajó el excel " + nombre + " desde la Red");
                                                }
                                                else
                                                {
                                                    FileInfo originalFile9 = new FileInfo(cLinkCarpeta + @"\" + cNombreArchivo9);
                                                    if (originalFile9.Exists)
                                                    {
                                                        originalFile9.CopyTo(AppDomain.CurrentDomain.BaseDirectory + @"\data\" + nombre + "." + cExtension, true);

                                                        cRespuestas = cRespuestas + DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt") + " - Se bajó el excel " + nombre + " desde la Red  \n";
                                                        Funciones.CapturarMensaje("Mensaje: Se bajó el excel " + nombre + " desde la Red");
                                                    }
                                                    else
                                                    {
                                                        Funciones.CapturarMensaje("El archivo excel para \"" + nombre + "\" no existe o no tiene permisos suficientes para acceder a el");
                                                        return new Resultado(-1, "El archivo excel para \"" + nombre + "\" no existe o no tiene permisos suficientes para acceder a el");
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }


                // 3. IMPORTAMOS DATOS A PI DESDE EL EXCEL
                string cDia = dFechaImportacion.Day < 10 ? "0" + dFechaImportacion.Day.ToString() : dFechaImportacion.Day.ToString();
                FileInfo existingFile = new FileInfo(AppDomain.CurrentDomain.BaseDirectory + @"\data\" + nombre + "." + cExtension);

                using (ExcelPackage package = new ExcelPackage(existingFile))
                {
                    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[cDia];

                    int filas = worksheet.Dimension.Rows;
                    int columnas = worksheet.Dimension.Columns;

                    //buscamos elementos columnas AF del excel
                    string cRutaAFExcel = oAFExcel.GetPath();
                    cRutaAFExcel = cRutaAFExcel.Replace(@"\\" + oPI.Name + @"\" + cDatabase + @"\", "");

                    var oSearch2 = new AFElementSearch(oDB, "Buscar", @"Root:'" + cRutaAFExcel + "' TemplateName:='Columna-Excel-Generico'");
                    foreach (AFElement oElementoTemp in oSearch2.FindElements(fullLoad: true))
                    {
                        int nFilaInicioDatos = 0;
                        int nFilaFinDatos = 0;
                        string cColumnaDato = "";
                        string cTag = "";

                        try
                        {
                            nFilaInicioDatos = oElementoTemp.Attributes["Fila Inicio Datos"].GetValue().ValueAsInt32();
                            nFilaFinDatos = worksheet.Dimension.Rows;
                            cColumnaDato = oElementoTemp.Attributes["Columna Dato"].GetValue().ToString();

                            if (oElementoTemp.Attributes["XLS"].DataReference != null)
                            {
                                cTag = oElementoTemp.Attributes["XLS"].DataReference.ToString();
                                cTag = cTag.Substring(cTag.LastIndexOf(@"\") + 1);
                            }

                            //importamos datos
                            for (int i = nFilaInicioDatos; i <= nFilaFinDatos; i++)
                            {
                                object oValorHora = worksheet.Cells[cColumnaHora + i.ToString()].Value;

                                if (oValorHora != null && oValorHora.ToString().Trim().ToUpper() != "TOTALES")
                                {
                                    DateTime dFecha = new DateTime();
                                    bool bFechaValida = false;
                                    try
                                    {
                                        dFecha = Convert.ToDateTime(oValorHora);
                                        dFecha = dFechaImportacion.AddHours(dFecha.Hour).AddMinutes(dFecha.Minute);

                                        if (dFecha < DateTime.Now)
                                        {
                                            bFechaValida = true;
                                        }
                                    }
                                    catch (Exception ex3)
                                    {
                                        int nHora = 0;
                                            if (Int32.TryParse(oValorHora.ToString(), out nHora))
                                            {
                                                dFecha = dFechaImportacion.AddHours(nHora);
                                                bFechaValida = true;
                                            }
                                    }

                                    if (bFechaValida)
                                    {
                                        object oValorDato = worksheet.Cells[cColumnaDato + i.ToString()].Value;

                                        double nDato = 0;
                                        if (oValorDato != null)
                                        {
                                            if (Double.TryParse(oValorDato.ToString(), out nDato))
                                            {
                                                ModPIExtFunctions.Write_Value_PI(cServidorPIData, cTag, dFecha, nDato, true);
                                            }
                                        }
                                    }

                                }
                                else
                                {
                                    break;
                                }
                            }

                            cRespuestas = cRespuestas + DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt") + " - Se importó datos de la columna " + cColumnaDato + " para el " + dFechaImportacion.ToString("yyyy-MM-dd") + "  \n";
                            Funciones.CapturarMensaje("Mensaje: Se importó datos de la columna " + cColumnaDato + " para el " + dFechaImportacion.ToString("yyyy-MM-dd"));

                        }
                        catch (Exception ex2)
                        {
                            cRespuestas = cRespuestas + DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt") + " - No se pudo importar datos de la columna " + cColumnaDato + " para el " + dFechaImportacion.ToString("yyyy-MM-dd") + ". " + ex2.Message + "  \n";
                            Funciones.CapturarError(ex2, "Utilidades.svc - ImportarExcelMensualGenericoDeAF", "wsUNACEMPI", false);
                        }
                    }




                }



                return new Resultado(0, cRespuestas);
            }
            catch (Exception ex)
            {
                Funciones.CapturarError(ex, "Utilidades.svc - ImportarExcelMensualGenericoDeAF", "wsUNACEMPI", false);
                return new Resultado(-1, ex.Message);
            }

        }


        public Resultado ImportarExcelMensualGenericoDeAFCompleto(string nombre, string fecha, string usuario)
        {
            //SE REQUIERE INSTALAR EN EL SERVIDOR sharepointclientcomponents_16-6518-1200_x64-en-us
            //SE REQUIERE HAYA UNA CARPETA "DATA"

            string cRespuestas = "";

            string cServidorPIAF = WebConfigurationManager.AppSettings["ServidorPIAF"];
            string cServidorPIData = WebConfigurationManager.AppSettings["ServidorPIData"];

            string cURLSharepoint = WebConfigurationManager.AppSettings["URLSharepoint"];
            string cUsuarioSharepoint = WebConfigurationManager.AppSettings["UsuarioSharepoint"];
            string cClaveSharepoint = WebConfigurationManager.AppSettings["ClaveSharepoint"];

            string cURLExcel = "";
            int nTipoUbicacion = 0;
            string cColumnaHora = "";
            string cExtension = "";

            DateTime dFechaImportacion;
            try
            {
                dFechaImportacion = new DateTime(Convert.ToInt32(fecha.Substring(0, 4)), Convert.ToInt32(fecha.Substring(5, 2)), Convert.ToInt32(fecha.Substring(8, 2)));
            }
            catch (Exception ex)
            {
                Funciones.CapturarError(ex, "Utilidades.svc - ImportarExcelMensualGenericoDeAFCompleto - La fecha ingresada no es valida", "wsUNACEMPI", false);
                return new Resultado(-1, "La fecha ingresada no es valida");
            }

            Funciones.CapturarMensaje("Mensaje: El usuario " + usuario + " inició el proceso de importación de datos del excel mensual genérico " + nombre + " a PI para el mes de " + ObtenerNombreMesCompleto(dFechaImportacion.Month));

            try
            {
                // 1. BUSCAMOS CONFIGURACION DE EXCEL EN AF

                string cDatabase = WebConfigurationManager.AppSettings["databaseExternalData"];

                PISystems oPIAF = new PISystems();
                PISystem oPI = oPIAF[cServidorPIAF];

                oPI.Connect();
                AFDatabase oDB = oPI.Databases[cDatabase];

                //buscamos el elemento excel
                var oSearch = new AFElementSearch(oDB, "Buscar", @"Root:'' Name:='" + nombre + "' TemplateName:='Excel-Mensual-Generico'");
                List<AFElement> oElementosExcel = oSearch.FindElements(fullLoad: true).ToList();
                if (oElementosExcel.Count == 0)
                {
                    return new Resultado(-1, "No se encontro el elemento de template Excel-Mensual-Generico con nombre " + nombre);
                }

                AFElement oAFExcel = oElementosExcel[0];
                nTipoUbicacion = oAFExcel.Attributes["Tipo Ubicacion"].GetValue().ValueAsInt32();

                string cLinkCarpeta = oAFExcel.Attributes["Link Carpeta"].GetValue().Value.ToString().Trim();
                cExtension = oAFExcel.Attributes["Extension"].GetValue().Value.ToString().Trim();

                string cNombreArchivo = oAFExcel.Attributes["Nombre Archivo"].GetValue().Value.ToString().Trim();
                cNombreArchivo = cNombreArchivo + " " + ObtenerNombreMesCompleto(dFechaImportacion.Month) + dFechaImportacion.Year + "." + cExtension;

                string cNombreArchivo2 = oAFExcel.Attributes["Nombre Archivo"].GetValue().Value.ToString().Trim();
                cNombreArchivo2 = cNombreArchivo2 + "-" + ObtenerNombreMesCompleto(dFechaImportacion.Month) + dFechaImportacion.Year + "." + cExtension;

                string cNombreArchivo3 = oAFExcel.Attributes["Nombre Archivo"].GetValue().Value.ToString().Trim();
                cNombreArchivo3 = cNombreArchivo3 + "" + ObtenerNombreMesCompleto(dFechaImportacion.Month) + dFechaImportacion.Year + "." + cExtension;


                string cNombreArchivo4 = oAFExcel.Attributes["Nombre Archivo"].GetValue().Value.ToString().Trim();
                cNombreArchivo4 = cNombreArchivo4 + " " + ObtenerNombreMesCompleto(dFechaImportacion.Month) + "-" + dFechaImportacion.Year + "." + cExtension;

                string cNombreArchivo5 = oAFExcel.Attributes["Nombre Archivo"].GetValue().Value.ToString().Trim();
                cNombreArchivo5 = cNombreArchivo5 + "-" + ObtenerNombreMesCompleto(dFechaImportacion.Month) + "-" + dFechaImportacion.Year + "." + cExtension;

                string cNombreArchivo6 = oAFExcel.Attributes["Nombre Archivo"].GetValue().Value.ToString().Trim();
                cNombreArchivo6 = cNombreArchivo6 + "" + ObtenerNombreMesCompleto(dFechaImportacion.Month) + "-" + dFechaImportacion.Year + "." + cExtension;


                string cNombreArchivo7 = oAFExcel.Attributes["Nombre Archivo"].GetValue().Value.ToString().Trim();
                cNombreArchivo7 = cNombreArchivo7 + " " + ObtenerNombreMesCompleto(dFechaImportacion.Month) + " " + dFechaImportacion.Year + "." + cExtension;

                string cNombreArchivo8 = oAFExcel.Attributes["Nombre Archivo"].GetValue().Value.ToString().Trim();
                cNombreArchivo8 = cNombreArchivo8 + "-" + ObtenerNombreMesCompleto(dFechaImportacion.Month) + " " + dFechaImportacion.Year + "." + cExtension;

                string cNombreArchivo9 = oAFExcel.Attributes["Nombre Archivo"].GetValue().Value.ToString().Trim();
                cNombreArchivo9 = cNombreArchivo9 + "" + ObtenerNombreMesCompleto(dFechaImportacion.Month) + " " + dFechaImportacion.Year + "." + cExtension;




                if (nTipoUbicacion == 0) //sharepoint
                    cURLExcel = cLinkCarpeta + "/" + cNombreArchivo;
                else  //red
                    cURLExcel = cLinkCarpeta + @"\" + cNombreArchivo;

                cColumnaHora = oAFExcel.Attributes["Columna Hora"].GetValue().Value.ToString().Trim();





                if (nTipoUbicacion == 0)    //sharepoint
                {
                    // 1. CONECTAMOS CON SHAREPOINT

                    System.Net.ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12;

                    //Namespace: It belongs to Microsoft.SharePoint.Client
                    ClientContext ctx = new ClientContext(cURLSharepoint);

                    // Namespace: It belongs to System.Security
                    SecureString secureString = new SecureString();
                    cClaveSharepoint.ToList().ForEach(secureString.AppendChar);

                    // Namespace: It belongs to Microsoft.SharePoint.Client
                    ctx.Credentials = new SharePointOnlineCredentials(cUsuarioSharepoint, secureString);

                    // Namespace: It belongs to Microsoft.SharePoint.Client
                    Web web = ctx.Web;

                    ctx.Load(web);
                    ctx.ExecuteQuery();
                    cRespuestas = cRespuestas + DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt") + " - Se conectó a Sharepoint \n";
                    Funciones.CapturarMensaje("Mensaje: Se conectó a Sharepoint");



                    // 2. BAJAMOS EL EXCEL
                    Uri filename = new Uri(cURLExcel);
                    string server = filename.AbsoluteUri.Replace(filename.AbsolutePath, "");
                    string serverrelative = filename.AbsolutePath;

                    Microsoft.SharePoint.Client.FileInformation f = Microsoft.SharePoint.Client.File.OpenBinaryDirect(ctx, serverrelative);

                    ctx.ExecuteQuery();

                    using (var fileStream = new FileStream(AppDomain.CurrentDomain.BaseDirectory + @"\data\" + nombre + "." + cExtension, FileMode.Create))
                        f.Stream.CopyTo(fileStream);

                    cRespuestas = cRespuestas + DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt") + " - Se bajó el excel " + nombre + " desde Sharepoint  \n";
                    Funciones.CapturarMensaje("Mensaje: Se bajó el excel " + nombre + " desde Sharepoint");

                }
                else    //red
                {
                    // 2. BAJAMOS EL EXCEL
                    FileInfo originalFile = new FileInfo(cURLExcel);
                    if (originalFile.Exists)
                    {
                        originalFile.CopyTo(AppDomain.CurrentDomain.BaseDirectory + @"\data\" + nombre + "." + cExtension, true);

                        cRespuestas = cRespuestas + DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt") + " - Se bajó el excel " + nombre + " desde la Red  \n";
                        Funciones.CapturarMensaje("Mensaje: Se bajó el excel " + nombre + " desde la Red");
                    }
                    else
                    {
                        FileInfo originalFile2 = new FileInfo(cLinkCarpeta + @"\" + cNombreArchivo2);
                        if (originalFile2.Exists)
                        {
                            originalFile2.CopyTo(AppDomain.CurrentDomain.BaseDirectory + @"\data\" + nombre + "." + cExtension, true);

                            cRespuestas = cRespuestas + DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt") + " - Se bajó el excel " + nombre + " desde la Red  \n";
                            Funciones.CapturarMensaje("Mensaje: Se bajó el excel " + nombre + " desde la Red");
                        }
                        else
                        {
                            FileInfo originalFile3 = new FileInfo(cLinkCarpeta + @"\" + cNombreArchivo3);
                            if (originalFile3.Exists)
                            {
                                originalFile3.CopyTo(AppDomain.CurrentDomain.BaseDirectory + @"\data\" + nombre + "." + cExtension, true);

                                cRespuestas = cRespuestas + DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt") + " - Se bajó el excel " + nombre + " desde la Red  \n";
                                Funciones.CapturarMensaje("Mensaje: Se bajó el excel " + nombre + " desde la Red");
                            }
                            else
                            {
                                FileInfo originalFile4 = new FileInfo(cLinkCarpeta + @"\" + cNombreArchivo4);
                                if (originalFile4.Exists)
                                {
                                    originalFile4.CopyTo(AppDomain.CurrentDomain.BaseDirectory + @"\data\" + nombre + "." + cExtension, true);

                                    cRespuestas = cRespuestas + DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt") + " - Se bajó el excel " + nombre + " desde la Red  \n";
                                    Funciones.CapturarMensaje("Mensaje: Se bajó el excel " + nombre + " desde la Red");
                                }
                                else
                                {
                                    FileInfo originalFile5 = new FileInfo(cLinkCarpeta + @"\" + cNombreArchivo5);
                                    if (originalFile5.Exists)
                                    {
                                        originalFile5.CopyTo(AppDomain.CurrentDomain.BaseDirectory + @"\data\" + nombre + "." + cExtension, true);

                                        cRespuestas = cRespuestas + DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt") + " - Se bajó el excel " + nombre + " desde la Red  \n";
                                        Funciones.CapturarMensaje("Mensaje: Se bajó el excel " + nombre + " desde la Red");
                                    }
                                    else
                                    {
                                        FileInfo originalFile6 = new FileInfo(cLinkCarpeta + @"\" + cNombreArchivo6);
                                        if (originalFile6.Exists)
                                        {
                                            originalFile6.CopyTo(AppDomain.CurrentDomain.BaseDirectory + @"\data\" + nombre + "." + cExtension, true);

                                            cRespuestas = cRespuestas + DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt") + " - Se bajó el excel " + nombre + " desde la Red  \n";
                                            Funciones.CapturarMensaje("Mensaje: Se bajó el excel " + nombre + " desde la Red");
                                        }
                                        else
                                        {
                                            FileInfo originalFile7 = new FileInfo(cLinkCarpeta + @"\" + cNombreArchivo7);
                                            if (originalFile7.Exists)
                                            {
                                                originalFile7.CopyTo(AppDomain.CurrentDomain.BaseDirectory + @"\data\" + nombre + "." + cExtension, true);

                                                cRespuestas = cRespuestas + DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt") + " - Se bajó el excel " + nombre + " desde la Red  \n";
                                                Funciones.CapturarMensaje("Mensaje: Se bajó el excel " + nombre + " desde la Red");
                                            }
                                            else
                                            {
                                                FileInfo originalFile8 = new FileInfo(cLinkCarpeta + @"\" + cNombreArchivo8);
                                                if (originalFile8.Exists)
                                                {
                                                    originalFile8.CopyTo(AppDomain.CurrentDomain.BaseDirectory + @"\data\" + nombre + "." + cExtension, true);

                                                    cRespuestas = cRespuestas + DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt") + " - Se bajó el excel " + nombre + " desde la Red  \n";
                                                    Funciones.CapturarMensaje("Mensaje: Se bajó el excel " + nombre + " desde la Red");
                                                }
                                                else
                                                {
                                                    FileInfo originalFile9 = new FileInfo(cLinkCarpeta + @"\" + cNombreArchivo9);
                                                    if (originalFile9.Exists)
                                                    {
                                                        originalFile9.CopyTo(AppDomain.CurrentDomain.BaseDirectory + @"\data\" + nombre + "." + cExtension, true);

                                                        cRespuestas = cRespuestas + DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt") + " - Se bajó el excel " + nombre + " desde la Red  \n";
                                                        Funciones.CapturarMensaje("Mensaje: Se bajó el excel " + nombre + " desde la Red");
                                                    }
                                                    else
                                                    {
                                                        Funciones.CapturarMensaje("El archivo excel para \"" + nombre + "\" no existe o no tiene permisos suficientes para acceder a el");
                                                        return new Resultado(-1, "El archivo excel para \"" + nombre + "\" no existe o no tiene permisos suficientes para acceder a el");
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }


                // 3. IMPORTAMOS DATOS A PI DESDE EL EXCEL
                DateTime dFechaIni = new DateTime(dFechaImportacion.Year, dFechaImportacion.Month, 1);
                DateTime dFechaFin = dFechaIni.AddMonths(1).AddDays(-1);
                DateTime dFechaTemp = dFechaIni;

                while (dFechaTemp <= dFechaFin)
                {
                    string cDia = dFechaTemp.Day < 10 ? "0" + dFechaTemp.Day.ToString() : dFechaTemp.Day.ToString();
                    FileInfo existingFile = new FileInfo(AppDomain.CurrentDomain.BaseDirectory + @"\data\" + nombre + "." + cExtension);

                    using (ExcelPackage package = new ExcelPackage(existingFile))
                    {
                        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                        ExcelWorksheet worksheet = package.Workbook.Worksheets[cDia];

                        int filas = worksheet.Dimension.Rows;
                        int columnas = worksheet.Dimension.Columns;

                        //buscamos elementos columnas AF del excel
                        string cRutaAFExcel = oAFExcel.GetPath();
                        cRutaAFExcel = cRutaAFExcel.Replace(@"\\" + oPI.Name + @"\" + cDatabase + @"\", "");

                        var oSearch2 = new AFElementSearch(oDB, "Buscar", @"Root:'" + cRutaAFExcel + "' TemplateName:='Columna-Excel-Generico'");
                        foreach (AFElement oElementoTemp in oSearch2.FindElements(fullLoad: true))
                        {
                            int nFilaInicioDatos = 0;
                            int nFilaFinDatos = 0;
                            string cColumnaDato = "";
                            string cTag = "";

                            try
                            {
                                nFilaInicioDatos = oElementoTemp.Attributes["Fila Inicio Datos"].GetValue().ValueAsInt32();
                                nFilaFinDatos = worksheet.Dimension.Rows;
                                cColumnaDato = oElementoTemp.Attributes["Columna Dato"].GetValue().ToString();

                                if (oElementoTemp.Attributes["XLS"].DataReference != null)
                                {
                                    cTag = oElementoTemp.Attributes["XLS"].DataReference.ToString();
                                    cTag = cTag.Substring(cTag.LastIndexOf(@"\") + 1);
                                }

                                //importamos datos
                                for (int i = nFilaInicioDatos; i <= nFilaFinDatos; i++)
                                {
                                    object oValorHora = worksheet.Cells[cColumnaHora + i.ToString()].Value;
                                    //Funciones.CapturarMensaje("Columna: " + cColumnaDato + ", Fecha: " + dFechaTemp.ToString("yyyy-MM-dd") + ", oHora: " + oValorHora.ToString());

                                    if (oValorHora != null && oValorHora.ToString().Trim().ToUpper() != "TOTALES")
                                    {
                                        DateTime dFecha = new DateTime();
                                        bool bFechaValida = false;
                                        try
                                        {
                                            dFecha = Convert.ToDateTime(oValorHora);
                                            dFecha = dFechaTemp.AddHours(dFecha.Hour).AddMinutes(dFecha.Minute);

                                            if (dFecha < DateTime.Now)
                                            {
                                                bFechaValida = true;
                                            }
                                        }
                                        catch (Exception ex3)
                                        {
                                            // Funciones.CapturarMensaje("Error: " + ex3.Message + " ---- " + "Columna: " + cColumnaDato + ", Fecha: " + dFechaTemp.ToString("yyyy-MM-dd") + ", oHora: " + oValorHora.ToString());
                                            int nHora = 0;
                                            if (Int32.TryParse(oValorHora.ToString(), out nHora))
                                            {
                                                dFecha = dFechaTemp.AddHours(nHora);
                                                bFechaValida = true;
                                            }
                                            else
                                            {
                                                //Funciones.CapturarMensaje("No se pudo convertir a numero la hora: " + oValorHora.ToString());
                                            }
                                            
                                        }

                                        //Funciones.CapturarMensaje("Fecha valida: " + bFechaValida.ToString() + " --- " + "Columna: " + cColumnaDato + ", Fecha: " + dFechaTemp.ToString("yyyy-MM-dd") + ", oHora: " + oValorHora.ToString());
                                        if (bFechaValida)
                                        {
                                            object oValorDato = worksheet.Cells[cColumnaDato + i.ToString()].Value;
                                            //Funciones.CapturarMensaje("valor dato: " + oValorDato.ToString() + " ---- " + "Columna: " + cColumnaDato + ", Fecha: " + dFechaTemp.ToString("yyyy-MM-dd") + ", oHora: " + oValorHora.ToString());
                                            

                                            double nDato = 0;
                                            if (oValorDato != null)
                                            {
                                                if (Double.TryParse(oValorDato.ToString(), out nDato))
                                                {
                                                    bool rpta = ModPIExtFunctions.Write_Value_PI(cServidorPIData, cTag, dFecha, nDato, true);
                                                    //Funciones.CapturarMensaje("Write_Value_PI: " + rpta.ToString() + " ----- Columna: " + cColumnaDato + ", Fecha: " + dFechaTemp.ToString("yyyy-MM-dd") + ", oHora: " + oValorHora.ToString());
                                                }
                                                else {
                                                    //Funciones.CapturarMensaje("no se pudo convertr a numero el valor: " + oValorDato.ToString() + " ----- Columna: " + cColumnaDato + ", Fecha: " + dFechaTemp.ToString("yyyy-MM-dd") + ", oHora: " + oValorHora.ToString());
                                          
                                                }
                                            }
                                            else
                                            {
                                                //Funciones.CapturarMensaje("valor dato nulo --- " + "Columna: " + cColumnaDato + ", Fecha: " + dFechaTemp.ToString("yyyy-MM-dd") + ", oHora: " + oValorHora.ToString());
                                            }
                                        }

                                    }
                                    else
                                    {
                                        break;
                                    }
                                }

                                cRespuestas = cRespuestas + DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt") + " - Se importó datos de la columna " + cColumnaDato + " para el " + dFechaTemp.ToString("yyyy-MM-dd") + "  \n";
                                Funciones.CapturarMensaje("Mensaje: Se importó datos de la columna " + cColumnaDato + " para el " + dFechaTemp.ToString("yyyy-MM-dd"));

                            }
                            catch (Exception ex2)
                            {
                                cRespuestas = cRespuestas + DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt") + " - No se pudo importar datos de la columna " + cColumnaDato + " para el " + dFechaTemp.ToString("yyyy-MM-dd") + ". " + ex2.Message + "  \n";
                                Funciones.CapturarError(ex2, "Utilidades.svc - ImportarExcelMensualGenericoDeAFCompleto", "wsUNACEMPI", false);
                            }
                        }

                    }

                    dFechaTemp = dFechaTemp.AddDays(1);
                }


                return new Resultado(0, cRespuestas);
            }
            catch (Exception ex)
            {
                Funciones.CapturarError(ex, "Utilidades.svc - ImportarExcelMensualGenericoDeAFCompleto", "wsUNACEMPI", false);
                return new Resultado(-1, ex.Message);
            }

        }


        public Resultado ImportarTodoExcelDeAFMensualGenerico(string fecha, string usuario)
        {
            string cRespuestas = "";

            string cServidorPIAF = WebConfigurationManager.AppSettings["ServidorPIAF"];
            string cServidorPIData = WebConfigurationManager.AppSettings["ServidorPIData"];

            DateTime dFechaImportacion;
            try
            {
                dFechaImportacion = new DateTime(Convert.ToInt32(fecha.Substring(0, 4)), Convert.ToInt32(fecha.Substring(5, 2)), Convert.ToInt32(fecha.Substring(8, 2)));
            }
            catch (Exception ex)
            {
                Funciones.CapturarError(ex, "Utilidades.svc - ImportarTodoExcelDeAFMensualGenerico - La fecha ingresada no es valida", "wsUNACEMPI", false);
                return new Resultado(-1, "La fecha ingresada no es valida");
            }

            Funciones.CapturarMensaje("Mensaje: El usuario " + usuario + " inició el proceso de importación de archivos de excel genérico de tipo \"Mensual\" a PI para el mes de " + ObtenerNombreMes(dFechaImportacion.Month) + "-" + dFechaImportacion.Year.ToString());

            try
            {
                string cDatabase = WebConfigurationManager.AppSettings["databaseExternalData"];

                PISystems oPIAF = new PISystems();
                PISystem oPI = oPIAF[cServidorPIAF];

                oPI.Connect();
                AFDatabase oDB = oPI.Databases[cDatabase];

                //buscamos el elemento excel
                var oSearch = new AFElementSearch(oDB, "Buscar", @"Root:'' TemplateName:='Excel-Mensual-Generico'");
                List<AFElement> oElementosExcel = oSearch.FindElements(fullLoad: true).ToList();
                foreach (AFElement oElementoExcel in oElementosExcel)
                {
                    string cNombre = oElementoExcel.Name;
                    Resultado oResultadoExcel = ImportarExcelMensualGenericoDeAFCompleto(cNombre, fecha, usuario);
                    cRespuestas = cRespuestas + oResultadoExcel.descripcion;
                }

                return new Resultado(0, cRespuestas);
            }
            catch (Exception ex)
            {
                Funciones.CapturarError(ex, "Utilidades.svc - ImportarTodoExcelDeAFMensualGenerico", "wsUNACEMPI", false);
                return new Resultado(-1, ex.Message);
            }
        }


        public Resultado ImportarExcelAnualGenericoDeAF(string nombre, string fecha, string usuario)
        {
            //SE REQUIERE INSTALAR EN EL SERVIDOR sharepointclientcomponents_16-6518-1200_x64-en-us
            //SE REQUIERE HAYA UNA CARPETA "DATA"

            string cRespuestas = "";

            string cServidorPIAF = WebConfigurationManager.AppSettings["ServidorPIAF"];
            string cServidorPIData = WebConfigurationManager.AppSettings["ServidorPIData"];

            string cURLSharepoint = WebConfigurationManager.AppSettings["URLSharepoint"];
            string cUsuarioSharepoint = WebConfigurationManager.AppSettings["UsuarioSharepoint"];
            string cClaveSharepoint = WebConfigurationManager.AppSettings["ClaveSharepoint"];

            string cURLExcel = "";
            int nTipoUbicacion = 0;
            string cColumnaFecha = "";

            DateTime dFechaImportacion;
            try
            {
                dFechaImportacion = new DateTime(Convert.ToInt32(fecha.Substring(0, 4)), Convert.ToInt32(fecha.Substring(5, 2)), Convert.ToInt32(fecha.Substring(8, 2)));
            }
            catch (Exception ex)
            {
                Funciones.CapturarError(ex, "Utilidades.svc - ImportarExcelAnualGenericoDeAF - La fecha ingresada no es valida", "wsUNACEMPI", false);
                return new Resultado(-1, "La fecha ingresada no es valida");
            }

            Funciones.CapturarMensaje("Mensaje: El usuario " + usuario + " inició el proceso de importación de datos del excel anual genérico " + nombre + " a PI para el mes de " + ObtenerNombreMes(dFechaImportacion.Month) + "-" + dFechaImportacion.Year.ToString());

            try
            {
                // 1. BUSCAMOS CONFIGURACION DE EXCEL EN AF

                string cDatabase = WebConfigurationManager.AppSettings["databaseExternalData"];

                PISystems oPIAF = new PISystems();
                PISystem oPI = oPIAF[cServidorPIAF];

                oPI.Connect();
                AFDatabase oDB = oPI.Databases[cDatabase];

                //buscamos el elemento excel
                var oSearch = new AFElementSearch(oDB, "Buscar", @"Root:'' Name:='" + nombre + "' TemplateName:='Excel-Anual-Generico'");
                List<AFElement> oElementosExcel = oSearch.FindElements(fullLoad: true).ToList();
                Funciones.CapturarMensaje(oElementosExcel.Count.ToString());
                if (oElementosExcel.Count == 0)
                {
                    Funciones.CapturarMensaje("No se encontro el elemento de template Excel-Anual-Generico con nombre " + nombre);
                    return new Resultado(-1, "No se encontro el elemento de template Excel-Anual-Generico con nombre " + nombre);
                }

                AFElement oAFExcel = oElementosExcel[0];
                nTipoUbicacion = oAFExcel.Attributes["Tipo Ubicacion"].GetValue().ValueAsInt32();
                Funciones.CapturarMensaje(nTipoUbicacion.ToString());
                cURLExcel = oAFExcel.Attributes["Link"].GetValue().Value.ToString().Trim();
                Funciones.CapturarMensaje(cURLExcel);
                cColumnaFecha = oAFExcel.Attributes["Columna Fecha"].GetValue().Value.ToString().Trim();
                Funciones.CapturarMensaje(cColumnaFecha);


                if (nTipoUbicacion == 0)    //sharepoint
                {
                    // 1. CONECTAMOS CON SHAREPOINT

                    System.Net.ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12;

                    //Namespace: It belongs to Microsoft.SharePoint.Client
                    ClientContext ctx = new ClientContext(cURLSharepoint);

                    // Namespace: It belongs to System.Security
                    SecureString secureString = new SecureString();
                    cClaveSharepoint.ToList().ForEach(secureString.AppendChar);

                    // Namespace: It belongs to Microsoft.SharePoint.Client
                    ctx.Credentials = new SharePointOnlineCredentials(cUsuarioSharepoint, secureString);

                    // Namespace: It belongs to Microsoft.SharePoint.Client
                    Web web = ctx.Web;

                    ctx.Load(web);
                    ctx.ExecuteQuery();
                    cRespuestas = cRespuestas + DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt") + " - Se conectó a Sharepoint \n";
                    Funciones.CapturarMensaje("Mensaje: Se conectó a Sharepoint");



                    // 2. BAJAMOS EL EXCEL
                    Uri filename = new Uri(cURLExcel);
                    string server = filename.AbsoluteUri.Replace(filename.AbsolutePath, "");
                    string serverrelative = filename.AbsolutePath;

                    Microsoft.SharePoint.Client.FileInformation f = Microsoft.SharePoint.Client.File.OpenBinaryDirect(ctx, serverrelative);

                    ctx.ExecuteQuery();

                    using (var fileStream = new FileStream(AppDomain.CurrentDomain.BaseDirectory + @"\data\" + nombre + ".xlsx", FileMode.Create))
                        f.Stream.CopyTo(fileStream);

                    cRespuestas = cRespuestas + DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt") + " - Se bajó el excel " + nombre + " desde Sharepoint  \n";
                    Funciones.CapturarMensaje("Mensaje: Se bajó el excel " + nombre + " desde Sharepoint");

                }
                else    //red
                {
                    // 2. BAJAMOS EL EXCEL
                    FileInfo originalFile = new FileInfo(cURLExcel);
                    
                    Funciones.CapturarMensaje(originalFile.Exists.ToString());
                    if (originalFile.Exists)
                    {
                        Funciones.CapturarMensaje(AppDomain.CurrentDomain.BaseDirectory);
                        Funciones.CapturarMensaje(AppDomain.CurrentDomain.BaseDirectory + @"\data\" + nombre + ".xlsx");
                        FileInfo destinoFile = originalFile.CopyTo(AppDomain.CurrentDomain.BaseDirectory + @"\data\" + nombre + ".xlsx", true);

                        cRespuestas = cRespuestas + DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt") + " - Se bajó el excel " + nombre + " desde la Red  \n";
                        Funciones.CapturarMensaje("Mensaje: Se bajó el excel " + nombre + " desde la Red");
                    }
                    else
                    {
                        Funciones.CapturarMensaje("El archivo en " + cURLExcel + " no existe o no tiene permisos suficientes para acceder a el");
                        return new Resultado(-1, "El archivo en " + cURLExcel + " no existe o no tiene permisos suficientes para acceder a el");
                    }
                }


                // 3. IMPORTAMOS DATOS A PI DESDE EL EXCEL
                string cMes = ObtenerNombreMes(dFechaImportacion.Month) + "_" + dFechaImportacion.Year.ToString();
                FileInfo existingFile = new FileInfo(AppDomain.CurrentDomain.BaseDirectory + @"\data\" + nombre + ".xlsx");

                using (ExcelPackage package = new ExcelPackage(existingFile))
                {
                    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[cMes];

                    int filas = worksheet.Dimension.Rows;
                    int columnas = worksheet.Dimension.Columns;

                    //buscamos elementos columnas AF del excel
                    string cRutaAFExcel = oAFExcel.GetPath();
                    cRutaAFExcel = cRutaAFExcel.Replace(@"\\" + oPI.Name + @"\" + cDatabase + @"\", "");

                    var oSearch2 = new AFElementSearch(oDB, "Buscar", @"Root:'" + cRutaAFExcel + "' TemplateName:='Columna-Excel-Generico'");
                    foreach (AFElement oElementoTemp in oSearch2.FindElements(fullLoad: true))
                    {
                        int nFilaInicioDatos = 0;
                        int nFilaFinDatos = 0;
                        string cColumnaDato = "";
                        string cTag = "";

                        try
                        {
                            nFilaInicioDatos = oElementoTemp.Attributes["Fila Inicio Datos"].GetValue().ValueAsInt32();
                            nFilaFinDatos = worksheet.Dimension.Rows;
                            cColumnaDato = oElementoTemp.Attributes["Columna Dato"].GetValue().ToString();

                            if (oElementoTemp.Attributes["XLS"].DataReference != null)
                            {
                                cTag = oElementoTemp.Attributes["XLS"].DataReference.ToString();
                                cTag = cTag.Substring(cTag.LastIndexOf(@"\") + 1);
                            }

                            //importamos datos
                            for (int i = nFilaInicioDatos; i <= nFilaFinDatos; i++)
                            {
                                object oValorFecha = worksheet.Cells[cColumnaFecha + i.ToString()].Value;

                                if (oValorFecha != null)
                                {
                                    DateTime dFecha = new DateTime();
                                    bool bFechaValida = false;
                                    try
                                    {
                                        try
                                        {
                                            dFecha = Convert.ToDateTime(oValorFecha);
                                            if (dFechaImportacion.Year == dFecha.Year && dFechaImportacion.Month == dFecha.Month && dFecha < DateTime.Now)
                                            {
                                                bFechaValida = true;
                                            }
                                        }
                                        catch (Exception ex4)
                                        {
                                            dFecha = DateTime.FromOADate(Convert.ToDouble(oValorFecha));
                                            if (dFechaImportacion.Year == dFecha.Year && dFechaImportacion.Month == dFecha.Month && dFecha < DateTime.Now)
                                            {
                                                bFechaValida = true;
                                            }
                                        }
                                    }
                                    catch (Exception ex3)
                                    {

                                    }

                                    if (bFechaValida)
                                    {
                                        object oValorDato = worksheet.Cells[cColumnaDato + i.ToString()].Value;

                                        double nDato = 0;
                                        if (oValorDato != null)
                                        {
                                            if (Double.TryParse(oValorDato.ToString(), out nDato))
                                            {
                                                AFAnnotations oAnotaciones = new AFAnnotations();
                                                AFAnnotation oAnotacion = oAnotaciones.Add("UsuarioModificacion", usuario);
                                                oAnotacion = oAnotaciones.Add("Comentarios", "Excel Load");

                                                ModPIExtFunctions.Write_Value_PI(cServidorPIData, cTag, dFecha, nDato, true, oAnotaciones);
                                            }
                                        }
                                    }

                                }
                                else
                                {
                                    break;
                                }
                            }

                            cRespuestas = cRespuestas + DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt") + " - Se importó datos de la columna " + cColumnaDato + " para el mes de " + cMes + "  \n";
                            Funciones.CapturarMensaje("Mensaje: Se importó datos de la columna " + cColumnaDato + " para el mes de " + cMes);

                        }
                        catch (Exception ex2)
                        {
                            cRespuestas = cRespuestas + DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt") + " - No se pudo importar datos de la columna " + cColumnaDato + " para el mes de " + cMes + ". " + ex2.Message + "  \n";
                            Funciones.CapturarError(ex2, "Utilidades.svc - ImportarExcelAnualGenericoDeAF", "wsUNACEMPI", false);
                        }
                    }




                }



                return new Resultado(0, cRespuestas);
            }
            catch (Exception ex)
            {
                Funciones.CapturarError(ex, "Utilidades.svc - ImportarExcelMensualGenericoDeAF", "wsUNACEMPI", false);
                return new Resultado(-1, ex.Message);
            }

        }


        public Resultado ImportarTodoExcelDeAFAnualGenerico(string fecha, string usuario)
        {
            string cRespuestas = "";

            string cServidorPIAF = WebConfigurationManager.AppSettings["ServidorPIAF"];
            string cServidorPIData = WebConfigurationManager.AppSettings["ServidorPIData"];

            DateTime dFechaImportacion;
            try
            {
                dFechaImportacion = new DateTime(Convert.ToInt32(fecha.Substring(0, 4)), Convert.ToInt32(fecha.Substring(5, 2)), Convert.ToInt32(fecha.Substring(8, 2)));
            }
            catch (Exception ex)
            {
                Funciones.CapturarError(ex, "Utilidades.svc - ImportarTodoExcelDeAFAnualGenerico - La fecha ingresada no es valida", "wsUNACEMPI", false);
                return new Resultado(-1, "La fecha ingresada no es valida");
            }

            Funciones.CapturarMensaje("Mensaje: El usuario " + usuario + " inició el proceso de importación de archivos de excel genérico de tipo \"Anual\" a PI para el mes de " + ObtenerNombreMes(dFechaImportacion.Month) + "-" + dFechaImportacion.Year.ToString());

            try
            {
                string cDatabase = WebConfigurationManager.AppSettings["databaseExternalData"];

                PISystems oPIAF = new PISystems();
                PISystem oPI = oPIAF[cServidorPIAF];

                oPI.Connect();
                AFDatabase oDB = oPI.Databases[cDatabase];

                //buscamos el elemento excel
                var oSearch = new AFElementSearch(oDB, "Buscar", @"Root:'' TemplateName:='Excel-Anual-Generico'");
                List<AFElement> oElementosExcel = oSearch.FindElements(fullLoad: true).ToList();
                foreach (AFElement oElementoExcel in oElementosExcel)
                {
                    string cNombre = oElementoExcel.Name;
                    Resultado oResultadoExcel = ImportarExcelAnualGenericoDeAF(cNombre, fecha, usuario);
                    cRespuestas = cRespuestas + oResultadoExcel.descripcion;
                }

                return new Resultado(0, cRespuestas);
            }
            catch (Exception ex)
            {
                Funciones.CapturarError(ex, "Utilidades.svc - ImportarTodoExcelDeAFAnualGenerico", "wsUNACEMPI", false);
                return new Resultado(-1, ex.Message);
            }
        }


        public Resultado ImportarExcelGenericoDeAF(string nombre, string usuario)
        {
            //SE REQUIERE INSTALAR EN EL SERVIDOR sharepointclientcomponents_16-6518-1200_x64-en-us
            //SE REQUIERE HAYA UNA CARPETA "DATA"

            string cRespuestas = "";

            string cServidorPIAF = WebConfigurationManager.AppSettings["ServidorPIAF"];
            string cServidorPIData = WebConfigurationManager.AppSettings["ServidorPIData"];

            string cURLSharepoint = WebConfigurationManager.AppSettings["URLSharepoint"];
            string cUsuarioSharepoint = WebConfigurationManager.AppSettings["UsuarioSharepoint"];
            string cClaveSharepoint = WebConfigurationManager.AppSettings["ClaveSharepoint"];

            string cURLExcel = "";
            int nTipoUbicacion = 0;
            string cColumnaFecha = "";
            string cHoja = "";

            Funciones.CapturarMensaje("Mensaje: El usuario " + usuario + " inició el proceso de importación de datos del excel genérico " + nombre + " a PI");

            try
            {
                // 1. BUSCAMOS CONFIGURACION DE EXCEL EN AF

                string cDatabase = WebConfigurationManager.AppSettings["databaseExternalData"];

                PISystems oPIAF = new PISystems();
                PISystem oPI = oPIAF[cServidorPIAF];

                oPI.Connect();
                AFDatabase oDB = oPI.Databases[cDatabase];

                //buscamos el elemento excel
                var oSearch = new AFElementSearch(oDB, "Buscar", @"Root:'' Name:='" + nombre + "' TemplateName:='Excel-Generico'");
                List<AFElement> oElementosExcel = oSearch.FindElements(fullLoad: true).ToList();
                if (oElementosExcel.Count == 0)
                {
                    return new Resultado(-1, "No se encontro el elemento de template Excel-Generico con nombre " + nombre);
                }

                AFElement oAFExcel = oElementosExcel[0];
                nTipoUbicacion = oAFExcel.Attributes["Tipo Ubicacion"].GetValue().ValueAsInt32();
                cURLExcel = oAFExcel.Attributes["Link"].GetValue().Value.ToString().Trim();
                cColumnaFecha = oAFExcel.Attributes["Columna Fecha"].GetValue().Value.ToString().Trim();
                cHoja = oAFExcel.Attributes["Hoja"].GetValue().Value.ToString().Trim();

                if (nTipoUbicacion == 0)    //sharepoint
                {
                    // 1. CONECTAMOS CON SHAREPOINT

                    System.Net.ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12;

                    //Namespace: It belongs to Microsoft.SharePoint.Client
                    ClientContext ctx = new ClientContext(cURLSharepoint);

                    // Namespace: It belongs to System.Security
                    SecureString secureString = new SecureString();
                    cClaveSharepoint.ToList().ForEach(secureString.AppendChar);

                    // Namespace: It belongs to Microsoft.SharePoint.Client
                    ctx.Credentials = new SharePointOnlineCredentials(cUsuarioSharepoint, secureString);

                    // Namespace: It belongs to Microsoft.SharePoint.Client
                    Web web = ctx.Web;

                    ctx.Load(web);
                    ctx.ExecuteQuery();
                    cRespuestas = cRespuestas + DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt") + " - Se conectó a Sharepoint \n";
                    Funciones.CapturarMensaje("Mensaje: Se conectó a Sharepoint");



                    // 2. BAJAMOS EL EXCEL
                    Uri filename = new Uri(cURLExcel);
                    string server = filename.AbsoluteUri.Replace(filename.AbsolutePath, "");
                    string serverrelative = filename.AbsolutePath;

                    Microsoft.SharePoint.Client.FileInformation f = Microsoft.SharePoint.Client.File.OpenBinaryDirect(ctx, serverrelative);

                    ctx.ExecuteQuery();

                    using (var fileStream = new FileStream(AppDomain.CurrentDomain.BaseDirectory + @"\data\" + nombre + ".xlsx", FileMode.Create))
                        f.Stream.CopyTo(fileStream);

                    cRespuestas = cRespuestas + DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt") + " - Se bajó el excel " + nombre + " desde Sharepoint  \n";
                    Funciones.CapturarMensaje("Mensaje: Se bajó el excel " + nombre + " desde Sharepoint");

                }
                else    //red
                {
                    // 2. BAJAMOS EL EXCEL
                    FileInfo originalFile = new FileInfo(cURLExcel);
                    if (originalFile.Exists)
                    {
                        originalFile.CopyTo(AppDomain.CurrentDomain.BaseDirectory + @"\data\" + nombre + ".xlsx", true);

                        cRespuestas = cRespuestas + DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt") + " - Se bajó el excel " + nombre + " desde la Red  \n";
                        Funciones.CapturarMensaje("Mensaje: Se bajó el excel " + nombre + " desde la Red");
                    }
                    else
                    {
                        Funciones.CapturarMensaje("El archivo en " + cURLExcel + " no existe o no tiene permisos suficientes para acceder a el");
                        return new Resultado(-1, "El archivo en " + cURLExcel + " no existe o no tiene permisos suficientes para acceder a el");
                    }
                }


                // 3. IMPORTAMOS DATOS A PI DESDE EL EXCEL
                FileInfo existingFile = new FileInfo(AppDomain.CurrentDomain.BaseDirectory + @"\data\" + nombre + ".xlsx");

                using (ExcelPackage package = new ExcelPackage(existingFile))
                {
                    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[cHoja];

                    int filas = worksheet.Dimension.Rows;
                    int columnas = worksheet.Dimension.Columns;

                    //buscamos elementos columnas AF del excel
                    string cRutaAFExcel = oAFExcel.GetPath();
                    cRutaAFExcel = cRutaAFExcel.Replace(@"\\" + oPI.Name + @"\" + cDatabase + @"\", "");

                    var oSearch2 = new AFElementSearch(oDB, "Buscar", @"Root:'" + cRutaAFExcel + "' TemplateName:='Columna-Excel-Generico'");
                    foreach (AFElement oElementoTemp in oSearch2.FindElements(fullLoad: true))
                    {
                        int nFilaInicioDatos = 0;
                        int nFilaFinDatos = 0;
                        string cColumnaDato = "";
                        string cTag = "";
                        int nModoImportacion = 0;
                        int nNumeroDiasHaciaAtras = 0;

                        try
                        {
                            nFilaInicioDatos = oElementoTemp.Attributes["Fila Inicio Datos"].GetValue().ValueAsInt32();
                            nFilaFinDatos = worksheet.Dimension.Rows;
                            cColumnaDato = oElementoTemp.Attributes["Columna Dato"].GetValue().ToString();
                            nModoImportacion = oElementoTemp.Attributes["Modo Importacion"].GetValue().ValueAsInt32();
                            nNumeroDiasHaciaAtras = oElementoTemp.Attributes["Numero Dias Hacia Atras"].GetValue().ValueAsInt32();

                            if (oElementoTemp.Attributes["XLS"].DataReference != null)
                            {
                                cTag = oElementoTemp.Attributes["XLS"].DataReference.ToString();
                                cTag = cTag.Substring(cTag.LastIndexOf(@"\") + 1);
                            }

                            //importamos datos
                            for (int i = nFilaInicioDatos; i <= nFilaFinDatos; i++)
                            {
                                object oValorFecha = worksheet.Cells[cColumnaFecha + i.ToString()].Value;

                                if (oValorFecha != null)
                                {
                                    DateTime dFecha = new DateTime();
                                    bool bFechaValida = false;
                                    try
                                    {
                                        try
                                        {
                                            dFecha = Convert.ToDateTime(oValorFecha);
                                            if (nModoImportacion == 1)
                                            {
                                                if (dFecha >= DateTime.Today.AddDays(-(nNumeroDiasHaciaAtras - 1)) && dFecha < DateTime.Now)
                                                {
                                                    bFechaValida = true;
                                                }
                                            }
                                            else
                                            {
                                                if (dFecha < DateTime.Now)
                                                {
                                                    bFechaValida = true;
                                                }
                                            }

                                        }
                                        catch (Exception ex4)
                                        {
                                            dFecha = DateTime.FromOADate(Convert.ToDouble(oValorFecha));
                                            if (nModoImportacion == 1)
                                            {
                                                if (dFecha >= DateTime.Today.AddDays(-(nNumeroDiasHaciaAtras - 1)) && dFecha < DateTime.Now)
                                                {
                                                    bFechaValida = true;
                                                }
                                            }
                                            else
                                            {
                                                if (dFecha < DateTime.Now)
                                                {
                                                    bFechaValida = true;
                                                }
                                            }

                                        }
                                    }
                                    catch (Exception ex3)
                                    {

                                    }

                                    if (bFechaValida)
                                    {
                                        object oValorDato = worksheet.Cells[cColumnaDato + i.ToString()].Value;

                                        double nDato = 0;
                                        if (oValorDato != null)
                                        {
                                            if (Double.TryParse(oValorDato.ToString(), out nDato))
                                            {
                                                ModPIExtFunctions.Write_Value_PI(cServidorPIData, cTag, dFecha, nDato, true);
                                            }
                                        }
                                    }

                                }
                                else
                                {
                                    break;
                                }
                            }

                            cRespuestas = cRespuestas + DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt") + " - Se importó datos de la columna " + cColumnaDato + "  \n";
                            Funciones.CapturarMensaje("Mensaje: Se importó datos de la columna " + cColumnaDato);

                        }
                        catch (Exception ex2)
                        {
                            cRespuestas = cRespuestas + DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt") + " - No se pudo importar datos de la columna " + cColumnaDato + ". " + ex2.Message + "  \n";
                            Funciones.CapturarError(ex2, "Utilidades.svc - ImportarExcelGenericoDeAF", "wsUNACEMPI", false);
                        }
                    }

                }

                return new Resultado(0, cRespuestas);

            }
            catch (Exception ex)
            {
                Funciones.CapturarError(ex, "Utilidades.svc - ImportarExcelGenericoDeAF", "wsUNACEMPI", false);
                return new Resultado(-1, ex.Message);
            }

        }


        public Resultado ImportarTodoExcelDeAFGenerico(string usuario)
        {
            string cRespuestas = "";

            string cServidorPIAF = WebConfigurationManager.AppSettings["ServidorPIAF"];
            string cServidorPIData = WebConfigurationManager.AppSettings["ServidorPIData"];

            Funciones.CapturarMensaje("Mensaje: El usuario " + usuario + " inició el proceso de importación de archivos de excel genérico a PI");

            try
            {
                string cDatabase = WebConfigurationManager.AppSettings["databaseExternalData"];

                PISystems oPIAF = new PISystems();
                PISystem oPI = oPIAF[cServidorPIAF];

                oPI.Connect();
                AFDatabase oDB = oPI.Databases[cDatabase];

                //buscamos el elemento excel
                var oSearch = new AFElementSearch(oDB, "Buscar", @"Root:'' TemplateName:='Excel-Generico'");
                List<AFElement> oElementosExcel = oSearch.FindElements(fullLoad: true).ToList();
                foreach (AFElement oElementoExcel in oElementosExcel)
                {
                    string cNombre = oElementoExcel.Name;
                    Resultado oResultadoExcel = ImportarExcelGenericoDeAF(cNombre, usuario);
                    cRespuestas = cRespuestas + oResultadoExcel.descripcion;
                }

                return new Resultado(0, cRespuestas);
            }
            catch (Exception ex)
            {
                Funciones.CapturarError(ex, "Utilidades.svc - ImportarTodoExcelDeAFGenerico", "wsUNACEMPI", false);
                return new Resultado(-1, ex.Message);
            }
        }



        //metodos de apoyo

        private string ObtenerFormula(string FormatoData, int Fila)
        {
            FormatoData = FormatoData.Replace(" ", "");

            List<char> oCaracteres = FormatoData.ToList();
            List<int> oIndicesColumnas = new List<int>();
            int nIndiceTemp = -1;

            for (int i = 0; i < oCaracteres.Count; i++)
            {
                if (
                    (oCaracteres[i] == '+' || oCaracteres[i] == '-' || oCaracteres[i] == '*' || oCaracteres[i] == '/' ||
                    oCaracteres[i] == '(' || oCaracteres[i] == ')')
                    &&
                    nIndiceTemp != -1
                    )
                {
                    oIndicesColumnas.Add(nIndiceTemp);
                    nIndiceTemp = -1;
                }
                else
                {
                    nIndiceTemp = i;
                }
            }

            if (nIndiceTemp != -1)
            {
                oIndicesColumnas.Add(nIndiceTemp);
            }

            for (int i = oIndicesColumnas.Count - 1; i >= 0; i--)
            {
                FormatoData = FormatoData.Insert(oIndicesColumnas[i] + 1, Fila.ToString());
            }

            return FormatoData;

        }

        private bool EsValorNulo(ExcelWorksheet ews, string cFormula)
        {
            bool bEsNulo = true;

            string[] aCeldas = cFormula.Split(new char[] { '+', '-', '*', '/', '(', ')' }, StringSplitOptions.RemoveEmptyEntries);

            foreach (string celda in aCeldas)
            {
                if (ews.Cells[celda].Value != null)
                {
                    bEsNulo = false;
                    break;
                }
            }

            return bEsNulo;
        }

        private string ObtenerNombreMes(int nMes)
        {
            string cMes = "";
            string cLenguaje = WebConfigurationManager.AppSettings["lenguaje"];

            if (cLenguaje == "en")
            {
                switch (nMes)
                {
                    case 1:
                        cMes = "JAN";
                        break;
                    case 2:
                        cMes = "FEB";
                        break;
                    case 3:
                        cMes = "MAR";
                        break;
                    case 4:
                        cMes = "APR";
                        break;
                    case 5:
                        cMes = "MAY";
                        break;
                    case 6:
                        cMes = "JUN";
                        break;
                    case 7:
                        cMes = "JUL";
                        break;
                    case 8:
                        cMes = "AUG";
                        break;
                    case 9:
                        cMes = "SEP";
                        break;
                    case 10:
                        cMes = "OCT";
                        break;
                    case 11:
                        cMes = "NOV";
                        break;
                    case 12:
                        cMes = "DEC";
                        break;
                }
            }
            else
            {
                switch (nMes)
                {
                    case 1:
                        cMes = "ENE";
                        break;
                    case 2:
                        cMes = "FEB";
                        break;
                    case 3:
                        cMes = "MAR";
                        break;
                    case 4:
                        cMes = "ABR";
                        break;
                    case 5:
                        cMes = "MAY";
                        break;
                    case 6:
                        cMes = "JUN";
                        break;
                    case 7:
                        cMes = "JUL";
                        break;
                    case 8:
                        cMes = "AGO";
                        break;
                    case 9:
                        cMes = "SEP";
                        break;
                    case 10:
                        cMes = "OCT";
                        break;
                    case 11:
                        cMes = "NOV";
                        break;
                    case 12:
                        cMes = "DIC";
                        break;
                }
            }
            

            return cMes;
        }

        private string ObtenerNombreMesCompleto(int nMes)
        {
            string cMes = "";
            string cLenguaje = WebConfigurationManager.AppSettings["lenguaje"];

            if (cLenguaje == "en")
            {
                switch (nMes)
                {
                    case 1:
                        cMes = "JANUARY ";
                        break;
                    case 2:
                        cMes = "FEBRUARY";
                        break;
                    case 3:
                        cMes = "MARCH";
                        break;
                    case 4:
                        cMes = "APRIL";
                        break;
                    case 5:
                        cMes = "MAY";
                        break;
                    case 6:
                        cMes = "JUNE";
                        break;
                    case 7:
                        cMes = "JULY";
                        break;
                    case 8:
                        cMes = "AUGUST";
                        break;
                    case 9:
                        cMes = "SEPTEMBER";
                        break;
                    case 10:
                        cMes = "OCTUBER";
                        break;
                    case 11:
                        cMes = "NOVEMBER";
                        break;
                    case 12:
                        cMes = "DECEMBER";
                        break;
                }
            }
            else
            {
                switch (nMes)
                {
                    case 1:
                        cMes = "ENERO";
                        break;
                    case 2:
                        cMes = "FEBRERO";
                        break;
                    case 3:
                        cMes = "MARZO";
                        break;
                    case 4:
                        cMes = "ABRIL";
                        break;
                    case 5:
                        cMes = "MAYO";
                        break;
                    case 6:
                        cMes = "JUNIO";
                        break;
                    case 7:
                        cMes = "JULIO";
                        break;
                    case 8:
                        cMes = "AGOSTO";
                        break;
                    case 9:
                        cMes = "SETIEMBRE";
                        break;
                    case 10:
                        cMes = "OCTUBRE";
                        break;
                    case 11:
                        cMes = "NOVIEMBRE";
                        break;
                    case 12:
                        cMes = "DICIEMBRE";
                        break;
                }
            }

            

            return cMes;
        }
    }
}
