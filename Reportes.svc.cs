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
using common;
using System.Web.Configuration;
using System.Net;
using OSIsoft.AF.Search;
using System.Data;
using System.IO;
using System.Security.Principal;

namespace wsUNACEMPI
{
    // NOTE: You can use the "Rename" command on the "Refactor" menu to change the class name "Reportes" in code, svc and config file together.
    // NOTE: In order to launch WCF Test Client for testing this service, please select Reportes.svc or Reportes.svc.cs at the Solution Explorer and start debugging.
    public class Reportes : IReportes
    {
        public List<Area> ObtenerAreas()
        {
            try
            {
                int a = 1;
                string cDatabase = WebConfigurationManager.AppSettings["database"];
                string cServidorPIAF = WebConfigurationManager.AppSettings["ServidorPIAF"];
          
                List<Area> oAreas = new List<Area>();

                PISystems oPIAF = new PISystems();
                PISystem oPI = oPIAF[cServidorPIAF];           
                oPI.Connect();

                AFDatabase oDB = oPI.Databases[cDatabase];
                AFTable oAFTableAreas = oDB.Tables["Areas"];
                DataTable dtAreas = oAFTableAreas.Table;
                DataView dvAreas = dtAreas.DefaultView;


                for (int i = 0; i < dvAreas.Count; i++)
                {
                    Area oArea = new Area();
                    oArea.Codigo = dvAreas[0]["COD_AREA"].ToString();
                    oArea.Nombre = dvAreas[0]["DSC_AREA"].ToString();
                    oArea.NombreAbreviado = dvAreas[0]["ABR_AREA"].ToString();
                    oAreas.Add(oArea);
                }

                return oAreas;
            }
            catch (Exception ex)
            {
                Funciones.CapturarError(ex, "Reportes.svc - ObtenerAreas", "wsUNACEMPI", false);
                return new List<Area>();
            }


        }


        public ReporteHorario ObtenerReporteHorario(string nombre, string fechaini, string fechafin, string periodo)
        {
            ReporteHorario oReporteHorario = new ReporteHorario();
            oReporteHorario.Cabeceras = new List<CabeceraReporteHorario>();
            oReporteHorario.Filas = new List<FilaReporteHorario>();

            CabeceraReporteHorario oCabecera = new CabeceraReporteHorario();
            FilaReporteHorario oFila = new FilaReporteHorario();     

            DateTime ldt_Date_Ini;
            try
            {
                ldt_Date_Ini = new DateTime(Convert.ToInt32(fechaini.Substring(0, 4)), Convert.ToInt32(fechaini.Substring(5, 2)), Convert.ToInt32(fechaini.Substring(8, 2)));
            }
            catch (Exception ex)
            {
                Funciones.CapturarError(ex, "Reportes.svc - ObtenerReporteHorario", "wsUNACEMPI", false);
                return oReporteHorario;
            }

            DateTime ldt_Date_Fin;
            try
            {
                ldt_Date_Fin = new DateTime(Convert.ToInt32(fechafin.Substring(0, 4)), Convert.ToInt32(fechafin.Substring(5, 2)), Convert.ToInt32(fechafin.Substring(8, 2)));
            }
            catch (Exception ex)
            {
                Funciones.CapturarError(ex, "Reportes.svc - ObtenerReporteHorario", "wsUNACEMPI", false);
                return oReporteHorario;
            }

            if (nombre.Trim()=="")
            {
                return oReporteHorario;
            }

            //validamos periodo
            string cTipoPeriodo="";
            int nMinutosPeriodo=0;

            periodo = periodo.Trim();
            if (periodo == "")
            {
                return oReporteHorario;
            }
            else if (periodo.Substring(periodo.Length - 1, 1).ToUpper() != "H" && periodo.Substring(periodo.Length - 1, 1).ToUpper() != "M")
            {
                return oReporteHorario;
            }
            else
            {
                try
                {
                    cTipoPeriodo = periodo.Substring(periodo.Length - 1, 1).ToUpper();
                    string cNumeroPeriodo = periodo.Substring(0, periodo.Length - 1);
                    bool r = Int32.TryParse(cNumeroPeriodo, out nMinutosPeriodo);
                    if (r)
                    {
                        if (cTipoPeriodo == "H")
                        {
                            nMinutosPeriodo = nMinutosPeriodo * 60;
                        }
                    }
                    else
                    {
                        return oReporteHorario;
                    }
                }
                catch (Exception exn)
                {
                    return oReporteHorario;
                }
            }

            try
            {
                String cServidorPIAF = WebConfigurationManager.AppSettings["ServidorPIAF"];
                String cDatabase = WebConfigurationManager.AppSettings["databaseReportesNuevo"];

                PISystems oPIAF = new PISystems();
                PISystem oPI = oPIAF[cServidorPIAF];

                oPI.Connect();

                AFDatabase oDB = oPI.Databases[cDatabase];

                //buscamos elemento del reporte
                string cRutaReporte = "";
                bool bReporteExiste = false;
                var oSearchReporte = new AFElementSearch(oDB, "BuscarReporte", @"Root:'' Name:='" + nombre.Trim() + "' TemplateName:='Reporte - Reporte horario'");
                foreach (AFElement oReporte in oSearchReporte.FindElements(fullLoad: true))
                {
                    bReporteExiste = true;
                    cRutaReporte = oReporte.GetPath();
                    cRutaReporte = cRutaReporte.Replace(@"\\" + cServidorPIAF + @"\" + cDatabase + @"\", "");
                    break;
                }

                if (!bReporteExiste)
                {
                    return oReporteHorario;
                }

                //obtenemos las cabeceras
                var oSearch = new AFElementSearch(oDB, "Buscar", @"Root:'" + cRutaReporte + "' TemplateName:='Variable - Reporte horario'");
                IEnumerable<AFElement> oElementos = oSearch.FindElements(fullLoad: true);

                foreach (AFElement oAFElement in oElementos)
                {
                    oCabecera = new CabeceraReporteHorario();
                    oCabecera.Orden = oAFElement.Attributes["Numero orden"].GetValue().ValueAsInt32();
                    oCabecera.Titulo3 = oAFElement.Attributes["Descripcion"].GetValue().ToString();
                    oCabecera.Titulo2 = "";
                    oCabecera.Titulo1 = "";

                    if (oAFElement.Parent != null)
                    {
                        if (oAFElement.Parent.Template.Name == "Fila 2 - Reporte horario")
                        {
                            oCabecera.Titulo2 = oAFElement.Parent.Attributes["Descripcion"].GetValue().ToString();
                        }

                        if (oAFElement.Parent.Template.Name == "Fila 1 - Reporte horario")
                        {
                            oCabecera.Titulo1 = oAFElement.Parent.Attributes["Descripcion"].GetValue().ToString();
                        }

                        if (oAFElement.Parent.Parent != null)
                        {
                            if (oAFElement.Parent.Parent.Template.Name == "Fila 2 - Reporte horario")
                            {
                                oCabecera.Titulo2 = oAFElement.Parent.Parent.Attributes["Descripcion"].GetValue().ToString();
                            }

                            if (oAFElement.Parent.Parent.Template.Name == "Fila 1 - Reporte horario")
                            {
                                oCabecera.Titulo1 = oAFElement.Parent.Parent.Attributes["Descripcion"].GetValue().ToString();
                            }
                        }
                    }

                    oReporteHorario.Cabeceras.Add(oCabecera);
                }

                oReporteHorario.Cabeceras = oReporteHorario.Cabeceras.OrderBy(x => x.Orden).ToList();
                  


                //obtenemos tiempos
                List<AFTime> oTimes = new List<AFTime>();
                DateTime dFechaTemp = ldt_Date_Ini;
               
                while (dFechaTemp <= ldt_Date_Fin){
                    AFTime oTime = new AFTime(dFechaTemp.ToUniversalTime());
                    oTimes.Add(oTime);

                    dFechaTemp = dFechaTemp.AddMinutes(nMinutosPeriodo);
                }
               


                //llenamos datos en blanco
                foreach (AFElement oAFElement in oElementos)
                {
                    foreach (AFTime oTime in oTimes)
                    {
                        oFila = new FilaReporteHorario();
                        oFila.Fecha = oTime.LocalTime.ToString("yyyy-MM-dd HH:mm");
                        oFila.Orden = oAFElement.Attributes["Numero orden"].GetValue().ValueAsInt32();
                        oFila.Nombre = oAFElement.Name;
                        oFila.Dato = "";

                        oReporteHorario.Filas.Add(oFila);
                    }
                }

                oReporteHorario.Filas = oReporteHorario.Filas.OrderBy(x => x.Fecha).ThenBy(x => x.Orden).ToList();


                //obtenemos los datos
                if (oTimes.Count > 0)
                {
                    // Results should be sent back for 100 tags in each page.
                    PIPagingConfiguration config = new PIPagingConfiguration(PIPageType.TagCount, 100);
                    config.OperationTimeoutOverride = new TimeSpan(2, 0, 0);


                    // 1. obtenemos los datos exactos
                    IEnumerable<AFAttribute> attributes = oElementos.SelectMany(elem => elem.Attributes.OfType<AFAttribute>());
                    AFAttributeList attributeList = new AFAttributeList(attributes);

                    IEnumerable<AFAttribute> attributesFiltered = attributeList.Where(a => a.Name.ToUpper() == "Tag".ToUpper() && a.Element.Attributes["Formato muestreo"].GetValue().ValueAsInt32() == 1);
                    AFAttributeList attributeListFiltered = new AFAttributeList(attributesFiltered);

                    AFListData oData = attributeListFiltered.Data;

                    if (oPI.ConnectionInfo.IsConnected == false)
                        oPI.Connect();

                    IEnumerable<AFValues> listResults = oData.RecordedValuesAtTimes(oTimes, AFRetrievalMode.Exact, config);
                    foreach (AFValues oAFValores in listResults)
                    {
                        foreach (AFValue oAFValor in oAFValores) {

                            if (oAFValor.Value.GetType().Name == "AFEnumerationValue" || oAFValor.Value.GetType().BaseType.Name == "PIException" || oAFValor.Value.GetType().BaseType.Name == "SystemException")
                            {
                                continue;
                            }

                            FilaReporteHorario oFilaTemp = oReporteHorario.Filas.FirstOrDefault(x => x.Nombre==oAFValor.Attribute.Element.Name && x.Fecha==oAFValor.Timestamp.LocalTime.ToString("yyyy-MM-dd HH:mm"));
                            if (oFilaTemp != null)
                            {
                                oFilaTemp.Dato = Convert.ToDouble(oAFValor.Value).ToString("###,###,###.##");
                            }
                        }  
                    }


                    // 2. obtenemos los datos interpolados
                    IEnumerable<AFAttribute> attributes2 = oElementos.SelectMany(elem => elem.Attributes.OfType<AFAttribute>());
                    AFAttributeList attributeList2 = new AFAttributeList(attributes2);

                    IEnumerable<AFAttribute> attributesFiltered2 = attributeList2.Where(a => a.Name.ToUpper() == "Tag".ToUpper() && a.Element.Attributes["Formato muestreo"].GetValue().ValueAsInt32() == 0);
                    AFAttributeList attributeListFiltered2 = new AFAttributeList(attributesFiltered2);

                    AFListData oData2 = attributeListFiltered2.Data;

                    if (oPI.ConnectionInfo.IsConnected == false)
                        oPI.Connect();

                    IEnumerable<AFValues> listResults2 = oData2.InterpolatedValuesAtTimes(oTimes,null,true, config);
                    foreach (AFValues oAFValores in listResults2)
                    {
                        foreach (AFValue oAFValor in oAFValores)
                        {

                            if (oAFValor.Value.GetType().Name == "AFEnumerationValue" || oAFValor.Value.GetType().BaseType.Name == "PIException" || oAFValor.Value.GetType().BaseType.Name == "SystemException")
                            {
                                continue;
                            }

                            FilaReporteHorario oFilaTemp = oReporteHorario.Filas.FirstOrDefault(x => x.Nombre == oAFValor.Attribute.Element.Name && x.Fecha == oAFValor.Timestamp.LocalTime.ToString("yyyy-MM-dd HH:mm"));
                            if (oFilaTemp != null)
                            {
                                oFilaTemp.Dato = Convert.ToDouble(oAFValor.Value).ToString("###,###,###.##");
                            }
                        }
                    }


                }
                

            
                return oReporteHorario;

            }
            catch (Exception ex)
            {
                Funciones.CapturarError(ex, "Reportes.svc - ObtenerReporteHorario", "wsUNACEMPI", false);
                oReporteHorario = new ReporteHorario();
                oReporteHorario.Cabeceras = new List<CabeceraReporteHorario>();
                oReporteHorario.Filas = new List<FilaReporteHorario>();
                return oReporteHorario;
            }

        }


        public List<TipoReporte> ObtenerTiposReporte(string template)
        {
            try
            {
                string cDatabase = WebConfigurationManager.AppSettings["databaseReportesNuevo"];
                string cServidorPIAF = WebConfigurationManager.AppSettings["ServidorPIAF"];

                PISystems oPIAF = new PISystems();
                PISystem oPI = oPIAF[cServidorPIAF];

                oPI.Connect();

                AFDatabase lo_Database = oPI.Databases[cDatabase];
                List<TipoReporte> oTiposReporte = new List<TipoReporte>();
                TipoReporte oTipoReporte = new TipoReporte();

                var oSearch = new AFElementSearch(lo_Database, "Buscar", @"Root:'' TemplateName:='" + template + "'");

                foreach (AFElement oElemento in oSearch.FindElements(fullLoad: true))
                {
                    oTipoReporte = new TipoReporte();
                    oTipoReporte.Nombre = oElemento.Name;

                    if (oElemento.Attributes["Descripcion"] != null)
                        oTipoReporte.Descripcion = oElemento.Attributes["Descripcion"].GetValue().ToString();

                    if (oElemento.Attributes["Report Title"] != null)
                        oTipoReporte.Descripcion = oElemento.Attributes["Report Title"].GetValue().ToString();

                    oTiposReporte.Add(oTipoReporte);
                }

                //ordenamos la lista
                List<TipoReporte> oTiposReporteOrdenados = oTiposReporte.OrderBy(o => o.Descripcion).ToList();

                return oTiposReporteOrdenados;
            }
            catch (Exception ex)
            {
                Funciones.CapturarError(ex, "Reportes.svc - ObtenerTiposReporte", "wsUNACEMPI", false);
                return new List<TipoReporte>();
            }
        }


        public Resultado CaLcularClinkerAndCementInventory(string fecha)
        {
            DateTime ldt_Date;
            try
            {
                ldt_Date = new DateTime(Convert.ToInt32(fecha.Substring(0, 4)), Convert.ToInt32(fecha.Substring(5, 2)), Convert.ToInt32(fecha.Substring(8, 2)));
            }
            catch (Exception ex)
            {
                Funciones.CapturarError(ex, "Reportes.svc - CaLcularClinkerAndCementInventory", "wsUNACEMPI", false);
                return new Resultado(-1, ex.Message);
            }

            try
            {
                string cDatabase = WebConfigurationManager.AppSettings["databaseReportesNuevo"];
                string cServidorPIAF = WebConfigurationManager.AppSettings["ServidorPIAF"];
                string cServidorPIData = WebConfigurationManager.AppSettings["ServidorPIData"];

                PISystems oPIAF = new PISystems();
                PISystem oPI = oPIAF[cServidorPIAF];
                oPI.Connect();
                AFDatabase lo_Database = oPI.Databases[cDatabase];

                var oSearch = new AFElementSearch(lo_Database, "Buscar", @"Root:'' TemplateName:='Clinker & Cement Inventory'");

                //procesamos por cada fecha del mes
                DateTime dFechaIniMes = new DateTime(ldt_Date.Year, ldt_Date.Month, 1);
                DateTime dFechaTemp = dFechaIniMes;

                while (dFechaTemp <= ldt_Date)
                {
                    foreach (AFElement oElemento in oSearch.FindElements(fullLoad: true))
                    {
                        try
                        {
                            string cTagStock = "";
                            if (oElemento.Attributes["Stock"].DataReference != null)
                            {
                                cTagStock = oElemento.Attributes["Stock"].DataReference.ToString();
                                cTagStock = cTagStock.Substring(cTagStock.LastIndexOf(@"\") + 1);
                            }

                            string cTagConsumption = "";
                            if (oElemento.Attributes["Consumption"].DataReference != null)
                            {
                                cTagConsumption = oElemento.Attributes["Consumption"].DataReference.ToString();
                                cTagConsumption = cTagConsumption.Substring(cTagConsumption.LastIndexOf(@"\") + 1);
                            }

                            string cTagProduction = "";
                            if (oElemento.Attributes["Production"].DataReference != null)
                            {
                                cTagProduction = oElemento.Attributes["Production"].DataReference.ToString();
                                cTagProduction = cTagProduction.Substring(cTagProduction.LastIndexOf(@"\") + 1);
                            }


                            //obtenemos el valor openning
                            double nValorOpenning = 0;
                            AFValue oAFValorOpenning = ModPIExtFunctions.ObtenerAFValueExacto(cServidorPIData, cTagStock, dFechaTemp.AddDays(-1));
                            if (oAFValorOpenning != null)
                            {
                                if (!(oAFValorOpenning.Value.GetType().Name == "AFEnumerationValue" ||
                                      oAFValorOpenning.Value.GetType().BaseType.Name == "PIException" ||
                                      oAFValorOpenning.Value.GetType().BaseType.Name == "SystemException" ||
                                      oAFValorOpenning.Value.GetType().BaseType.Name == "Exception"))
                                {
                                    Double.TryParse(oAFValorOpenning.Value.ToString(), out nValorOpenning);
                                }
                            }

                            //obtenemos el valor consumption
                            double nValorConsumption = 0;
                            AFValue oAFValorConsumption = ModPIExtFunctions.ObtenerAFValueExacto(cServidorPIData, cTagConsumption, dFechaTemp);
                            if (oAFValorConsumption != null)
                            {
                                if (!(oAFValorConsumption.Value.GetType().Name == "AFEnumerationValue" ||
                                      oAFValorConsumption.Value.GetType().BaseType.Name == "PIException" ||
                                      oAFValorConsumption.Value.GetType().BaseType.Name == "SystemException" ||
                                      oAFValorConsumption.Value.GetType().BaseType.Name == "Exception"))
                                {
                                    Double.TryParse(oAFValorConsumption.Value.ToString(), out nValorConsumption);
                                }
                            }

                            //obtenemos el valor production
                            double nValorProduction = 0;
                            AFValue oAFValorProduction = ModPIExtFunctions.ObtenerAFValueExacto(cServidorPIData, cTagProduction, dFechaTemp);
                            if (oAFValorProduction != null)
                            {
                                if (!(oAFValorProduction.Value.GetType().Name == "AFEnumerationValue" ||
                                      oAFValorProduction.Value.GetType().BaseType.Name == "PIException" ||
                                      oAFValorProduction.Value.GetType().BaseType.Name == "SystemException" ||
                                      oAFValorProduction.Value.GetType().BaseType.Name == "Exception"))
                                {
                                    Double.TryParse(oAFValorProduction.Value.ToString(), out nValorProduction);
                                }
                            }

                            //obtenemos el valor closing
                            double nValorClosing = 0;
                            nValorClosing = nValorOpenning + nValorProduction - nValorConsumption;
                            Funciones.CapturarMensaje("nValorOpenning: " + nValorOpenning.ToString());
                            Funciones.CapturarMensaje("nValorProduction: " + nValorProduction.ToString());
                            Funciones.CapturarMensaje("nValorConsumption: " + nValorConsumption.ToString());
                            Funciones.CapturarMensaje("nValorClosing: " + nValorClosing.ToString());

                            //registramos el valor closing en el tag
                            ModPIExtFunctions.Write_Value_PI(cServidorPIData, cTagStock, dFechaTemp, nValorClosing, true);

                        }
                        catch (Exception ex2)
                        {
                            Funciones.CapturarError(ex2, "Reportes.svc - CaLcularClinkerAndCementInventory", "wsUNACEMPI", false);
                            continue;
                        }
                    }

                    dFechaTemp = dFechaTemp.AddDays(1);
                }

                return new Resultado(0, "Operación Exitosa");

            }
            catch (Exception ex)
            {
                Funciones.CapturarError(ex, "Reportes.svc - CaLcularClinkerAndCementInventory", "wsUNACEMPI", false);
                return new Resultado(-1, ex.Message);
            }
        }


        public List<Inventory> ObtenerReporteInventory(string nombre, string fecha)
        {
            List<Inventory> oInventarios = new List<Inventory>();
            Inventory oInventario = new Inventory();
            InventoryDetalle oDetalle = new InventoryDetalle();

            DateTime ldt_Date;
            try
            {
                ldt_Date = new DateTime(Convert.ToInt32(fecha.Substring(0, 4)), Convert.ToInt32(fecha.Substring(5, 2)), Convert.ToInt32(fecha.Substring(8, 2)));
            }
            catch (Exception ex)
            {
                Funciones.CapturarError(ex, "Reportes.svc - ObtenerReporteInventory", "wsUNACEMPI", false);
                return oInventarios;
            }

            try
            {
                DateTime ldt_Date_Ini = ldt_Date.AddDays(-ldt_Date.Day);
                DateTime ldt_Date_Fin = ldt_Date.AddDays(-ldt_Date.Day + 1).AddMonths(1).AddDays(-1);

                string cDatabase = WebConfigurationManager.AppSettings["databaseReportesNuevo"];
                string cServidorPIAF = WebConfigurationManager.AppSettings["ServidorPIAF"];
                string cServidorPIData = WebConfigurationManager.AppSettings["ServidorPIData"];

                PISystems oPIAF = new PISystems();
                PISystem oPI = oPIAF[cServidorPIAF];
                oPI.Connect();
                AFDatabase lo_Database = oPI.Databases[cDatabase];

                var oSearch = new AFElementSearch(lo_Database, "Buscar", @"Root:'' Name:='" + nombre + "' TemplateName:='Report template'");

                foreach (AFElement oElemento in oSearch.FindElements(fullLoad: true))
                {
                    //obtenemos cabeceras
                    string cRutaAFReporte = oElemento.GetPath();
                    cRutaAFReporte = cRutaAFReporte.Replace(@"\\" + oPI.Name + @"\" + cDatabase + @"\", "");

                    var oSearch2 = new AFElementSearch(lo_Database, "Buscar", @"Root:'" + cRutaAFReporte + "' TemplateName:='Clinker & Cement Inventory'");
                    IEnumerable<AFElement> oElementosInventarios = oSearch2.FindElements(fullLoad: true);
                    foreach (AFElement oElementoBloque in oElementosInventarios)
                    {
                        oInventario = new Inventory();
                        oInventario.Nombre = oElementoBloque.Name;
                        oInventario.Titulo = oElementoBloque.Name;
                        oInventario.TituloInitialStock = oElementoBloque.Attributes["Title Initial Stock"].GetValue().Value.ToString();
                        oInventario.TituloProduction = oElementoBloque.Attributes["Title Production"].GetValue().Value.ToString();
                        oInventario.TituloConsumption = oElementoBloque.Attributes["Title consumption"].GetValue().Value.ToString();
                        oInventario.TituloFinalStock = oElementoBloque.Attributes["Title Final Stock"].GetValue().Value.ToString();
                        oInventario.Orden = oElementoBloque.Attributes["Order"].GetValue().ValueAsInt32();
                        oInventario.Detalles = new List<InventoryDetalle>();

                        oInventarios.Add(oInventario);
                    }


                    //obtenemos detalles
                    IEnumerable<AFAttribute> attributes = oElementosInventarios.SelectMany(elem => elem.Attributes.OfType<AFAttribute>());
                    AFAttributeList attributeList = new AFAttributeList(attributes);

                    IEnumerable<AFAttribute> attributesFiltered = attributeList.Where(a => a.Name.ToUpper() == "Stock".ToUpper() || 
                                                                                           a.Name.ToUpper() == "Production".ToUpper() ||
                                                                                           a.Name.ToUpper() == "Consumption".ToUpper());
                    AFAttributeList attributeListFiltered = new AFAttributeList(attributesFiltered);

                    AFListData oData = attributeListFiltered.Data;

                    // Results should be sent back for 100 tags in each page.
                    PIPagingConfiguration config = new PIPagingConfiguration(PIPageType.TagCount, 100);
                    config.OperationTimeoutOverride = new TimeSpan(2, 0, 0);

                    if (oPI.ConnectionInfo.IsConnected == false)
                        oPI.Connect();
                 
                    AFTime oTiempoIni = new AFTime(ldt_Date_Ini.ToUniversalTime());
                    AFTime oTiempoFin = new AFTime(ldt_Date_Fin.ToUniversalTime());
                    AFTimeRange oTiempoRango = new AFTimeRange(oTiempoIni, oTiempoFin);

                    //traemos los datos
                    IEnumerable<AFValues> listResults = oData.RecordedValues(oTiempoRango, AFBoundaryType.Inside, null, false, config);

                    foreach (AFValues oAFValores in listResults)
                    {
                        foreach (AFValue oAFValor in oAFValores)
                        {
                            if (oAFValor.Value.GetType().Name == "AFEnumerationValue" ||
                                oAFValor.Value.GetType().BaseType.Name == "PIException" ||
                                oAFValor.Value.GetType().BaseType.Name == "SystemException" ||
                                oAFValor.Value.GetType().BaseType.Name == "Exception") {

                                continue;
                            }

                            Inventory oInventarioTemp = oInventarios.FirstOrDefault(x => x.Nombre == oAFValor.Attribute.Element.Name);
                            if (oInventarioTemp != null)
                            {
                                InventoryDetalle oDetalleTemp = oInventarioTemp.Detalles.FirstOrDefault(x => x.Fecha == oAFValor.Timestamp.LocalTime.ToString("yyyy-MM-dd"));
                                if (oDetalleTemp != null)
                                {
                                    if (oAFValor.Attribute.Name.ToUpper() == "Stock".ToUpper())
                                    {
                                        oDetalleTemp.FinalStock = oAFValor.ValueAsDouble().ToString("###,###,##0.00");
                                    }
                                    else if (oAFValor.Attribute.Name.ToUpper() == "Production".ToUpper())
                                    {
                                        oDetalleTemp.Production = oAFValor.ValueAsDouble().ToString("###,###,##0.00");
                                    }
                                    else if (oAFValor.Attribute.Name.ToUpper() == "Consumption".ToUpper())
                                    {
                                        oDetalleTemp.Consumption = oAFValor.ValueAsDouble().ToString("###,###,##0.00");
                                    }
                                }
                                else
                                {
                                    oDetalle = new InventoryDetalle();
                                    oDetalle.Fecha = oAFValor.Timestamp.LocalTime.ToString("yyyy-MM-dd");
                                    oDetalle.InitialStock = "";
                                    oDetalle.Production = "";
                                    oDetalle.Consumption = "";
                                    oDetalle.FinalStock = "";

                                    if (oAFValor.Attribute.Name.ToUpper() == "Stock".ToUpper())
                                    {
                                        oDetalle.FinalStock = oAFValor.ValueAsDouble().ToString("###,###,##0.00");
                                    }
                                    else if (oAFValor.Attribute.Name.ToUpper() == "Production".ToUpper())
                                    {
                                        oDetalle.Production = oAFValor.ValueAsDouble().ToString("###,###,##0.00");
                                    }
                                    else if (oAFValor.Attribute.Name.ToUpper() == "Consumption".ToUpper())
                                    {
                                        oDetalle.Consumption = oAFValor.ValueAsDouble().ToString("###,###,##0.00");
                                    }

                                    oInventarioTemp.Detalles.Add(oDetalle);
                                }


                                //Obtenemos el stock inicial
                                InventoryDetalle oDetalleTemp2 = oInventarioTemp.Detalles.FirstOrDefault(x => x.Fecha == oAFValor.Timestamp.LocalTime.AddDays(1).ToString("yyyy-MM-dd"));
                                if (oDetalleTemp2 != null)
                                {
                                    if (oAFValor.Attribute.Name.ToUpper() == "Stock".ToUpper())
                                    {
                                        oDetalleTemp2.InitialStock = oAFValor.ValueAsDouble().ToString("###,###,##0.00");
                                    }
                                }
                                else
                                {
                                    oDetalle = new InventoryDetalle();
                                    oDetalle.Fecha = oAFValor.Timestamp.LocalTime.AddDays(1).ToString("yyyy-MM-dd");
                                    oDetalle.InitialStock = "";
                                    oDetalle.Production = "";
                                    oDetalle.Consumption = "";
                                    oDetalle.FinalStock = "";

                                    if (oAFValor.Attribute.Name.ToUpper() == "Stock".ToUpper())
                                    {
                                        oDetalle.InitialStock = oAFValor.ValueAsDouble().ToString("###,###,##0.00");
                                    }

                                    oInventarioTemp.Detalles.Add(oDetalle);
                                }


                            }

                        }
                    }



                    break;
                }


                //agregamos fechas faltantea
                foreach (Inventory oInventarioTemp in oInventarios)
                {
                    DateTime dFechaTemp = ldt_Date_Ini.AddDays(1);
                    while (dFechaTemp <= ldt_Date_Fin)
                    {
                        InventoryDetalle oDetalleTemp = oInventarioTemp.Detalles.FirstOrDefault(x => x.Fecha == dFechaTemp.ToString("yyyy-MM-dd"));
                        if (oDetalleTemp == null)
                        {
                            oDetalle = new InventoryDetalle();
                            oDetalle.Fecha = dFechaTemp.ToString("yyyy-MM-dd");
                            oDetalle.InitialStock = "";
                            oDetalle.Production = "";
                            oDetalle.Consumption = "";
                            oDetalle.FinalStock = "";
                            oInventarioTemp.Detalles.Add(oDetalle);
                        }

                        dFechaTemp = dFechaTemp.AddDays(1);
                    }
                }


                //ordenamos
                List<Inventory> oInventariosOrdenados = oInventarios.OrderBy(x => x.Orden).ToList();
                foreach (Inventory oInventarioTemp in oInventariosOrdenados)
                {
                    oInventarioTemp.Detalles = oInventarioTemp.Detalles.OrderBy(x => x.Fecha).ToList();
                    oInventarioTemp.Detalles = oInventarioTemp.Detalles.Where(x => Convert.ToDateTime(x.Fecha) >= ldt_Date_Ini.AddDays(1) && Convert.ToDateTime(x.Fecha) <= ldt_Date_Fin).ToList();
                }

                

                return oInventariosOrdenados;
            }
            catch (Exception ex)
            {
                Funciones.CapturarError(ex, "Reportes.svc - ObtenerReporteInventory", "wsUNACEMPI", false);
                return oInventarios;
            }

        }



    }

    
}
