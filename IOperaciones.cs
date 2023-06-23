using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.ServiceModel;
using System.ServiceModel.Web;
using System.Text;
using System.Data;


namespace wsUNACEMPI
{
    // NOTE: You can use the "Rename" command on the "Refactor" menu to change the interface name "IOperaciones" in both code and config file together.
    [ServiceContract]
    public interface IOperaciones
    {
        [OperationContract]
        [WebGet(UriTemplate = "/ListarEventosYAtributosPorArea?database={database}&query={query}&area={area}", ResponseFormat = WebMessageFormat.Json, BodyStyle = WebMessageBodyStyle.Bare)]
        List<EventoFull> ListarEventosYAtributosPorArea(string database, string query, string area);


        [OperationContract]
        [WebGet(UriTemplate = "/ListarParadas?fechaini={fechaini}&fechafin={fechafin}", ResponseFormat = WebMessageFormat.Json, BodyStyle = WebMessageBodyStyle.Bare)]
        List<Parada> ListarParadas(string fechaini, string fechafin);


        /* INICIO - METODOS PARA DRAKE */

        [OperationContract]
        [WebGet(UriTemplate = "/GenerarRDPDrake?fecha={fecha}", ResponseFormat = WebMessageFormat.Json, BodyStyle = WebMessageBodyStyle.Bare)]
        Resultado GenerarRDPDrake(string fecha);


        [OperationContract]
        [WebGet(UriTemplate = "/ListarDeltasDeFormulacion?template={template}&fechaini={fechaini}&fechafin={fechafin}&maquina={maquina}", ResponseFormat = WebMessageFormat.Json, BodyStyle = WebMessageBodyStyle.Bare)]
        List<Formulacion> ListarDeltasDeFormulacion(string template, string fechaini, string fechafin, string maquina);


        [OperationContract]
        [WebGet(UriTemplate = "/ListarMaquinas", ResponseFormat = WebMessageFormat.Json, BodyStyle = WebMessageBodyStyle.Bare)]
        List<Maquina> ListarMaquinas();


        [OperationContract]
        [WebInvoke(UriTemplate = "/ActualizarDeltaManual", RequestFormat = WebMessageFormat.Json, ResponseFormat = WebMessageFormat.Json,
        BodyStyle = WebMessageBodyStyle.Bare, Method = "POST")]
        Resultado ActualizarDeltaManual(EntradaActualizarDeltaManual oEntradaActualizarDeltaManual);


        [OperationContract]
        [WebGet(UriTemplate = "/ObtenerDOR?fecha={fecha}", ResponseFormat = WebMessageFormat.Json, BodyStyle = WebMessageBodyStyle.Bare)]
        List<BloqueDOR> ObtenerDOR(string fecha);


        [OperationContract]
        [WebGet(UriTemplate = "/ListarDependenciasDOR", ResponseFormat = WebMessageFormat.Json, BodyStyle = WebMessageBodyStyle.Bare)]
        List<DependenciaDOR> ListarDependenciasDOR();


        [OperationContract]
        [WebGet(UriTemplate = "/ListarFactoresDeFormulacion?fecha={fecha}&dependencia={dependencia}", ResponseFormat = WebMessageFormat.Json, BodyStyle = WebMessageBodyStyle.Bare)]
        List<FactorFormulacion> ListarFactoresDeFormulacion(string fecha, string dependencia);


        [OperationContract]
        [WebInvoke(UriTemplate = "/ActualizarFactorDeFormulacion", RequestFormat = WebMessageFormat.Json, ResponseFormat = WebMessageFormat.Json,
        BodyStyle = WebMessageBodyStyle.Bare, Method = "POST")]
        Resultado ActualizarFactorDeFormulacion(EntradaActualizarFactorDeFormulacion oEntradaActualizarFactorDeFormulacion);


        [OperationContract]
        [WebGet(UriTemplate = "/ListarMasterData", ResponseFormat = WebMessageFormat.Json, BodyStyle = WebMessageBodyStyle.Bare)]
        List<MasterData> ListarMasterData();

        [OperationContract]
        [WebGet(UriTemplate = "/ListarMasterDataRD", ResponseFormat = WebMessageFormat.Json, BodyStyle = WebMessageBodyStyle.Bare)]
        List<MasterData> ListarMasterDataRD();


        [OperationContract]
        [WebGet(UriTemplate = "/ListarFactoresDeFormulacionEnRangoFecha?fechaini={fechaini}&fechafin={fechafin}", ResponseFormat = WebMessageFormat.Json, BodyStyle = WebMessageBodyStyle.Bare)]
        List<FactorFormulacion> ListarFactoresDeFormulacionEnRangoFecha(string fechaini, string fechafin);


        [OperationContract]
        [WebGet(UriTemplate = "/ObtenerFAC?fecha={fecha}", ResponseFormat = WebMessageFormat.Json, BodyStyle = WebMessageBodyStyle.Bare)]
        FAC ObtenerFAC(string fecha);


        [OperationContract]
        [WebGet(UriTemplate = "/ListarDeltasDeFormulacionOriginales?template={template}&fechaini={fechaini}&fechafin={fechafin}&maquina={maquina}", ResponseFormat = WebMessageFormat.Json, BodyStyle = WebMessageBodyStyle.Bare)]
        List<Formulacion> ListarDeltasDeFormulacionOriginales(string template, string fechaini, string fechafin, string maquina);


        [OperationContract]
        [WebGet(UriTemplate = "/GenerarRDPDrakeEnRangoFechas?fechaini={fechaini}&fechafin={fechafin}", ResponseFormat = WebMessageFormat.Json, BodyStyle = WebMessageBodyStyle.Bare)]
        Resultado GenerarRDPDrakeEnRangoFechas(string fechaini, string fechafin);


        [OperationContract]
        [WebInvoke(UriTemplate = "/RestaurarDelta", RequestFormat = WebMessageFormat.Json, ResponseFormat = WebMessageFormat.Json,
        BodyStyle = WebMessageBodyStyle.Bare, Method = "POST")]
        Resultado RestaurarDelta(EntradaRestaurarDelta oEntradaRestaurarDelta);


        [OperationContract]
        [WebGet(UriTemplate = "/ObtenerDOROriginal?fecha={fecha}", ResponseFormat = WebMessageFormat.Json, BodyStyle = WebMessageBodyStyle.Bare)]
        List<BloqueDOR> ObtenerDOROriginal(string fecha);


        [OperationContract]
        [WebGet(UriTemplate = "/ObtenerFACOriginal?fecha={fecha}", ResponseFormat = WebMessageFormat.Json, BodyStyle = WebMessageBodyStyle.Bare)]
        FAC ObtenerFACOriginal(string fecha);


        [OperationContract]
        [WebGet(UriTemplate = "/ObtenerFACEnRangoFechaDeMes?fechainimes={fechainimes}&fechafinmes={fechafinmes}", ResponseFormat = WebMessageFormat.Json, BodyStyle = WebMessageBodyStyle.Bare)]
        FAC ObtenerFACEnRangoFechaDeMes(string fechainimes, string fechafinmes);


        [OperationContract]
        [WebGet(UriTemplate = "/ObtenerFACOriginalEnRangoFechaDeMes?fechainimes={fechainimes}&fechafinmes={fechafinmes}", ResponseFormat = WebMessageFormat.Json, BodyStyle = WebMessageBodyStyle.Bare)]
        FAC ObtenerFACOriginalEnRangoFechaDeMes(string fechainimes, string fechafinmes);


        /* FIN - METODOS PARA DRAKE */



        [OperationContract]
        [WebGet(UriTemplate = "/RegistrarSumaDeTagEnRangoFecha?tag={tag}&tagsuma={tagsuma}&fechaini={fechaini}&fechafin={fechafin}&fechasuma={fechasuma}", ResponseFormat = WebMessageFormat.Json, BodyStyle = WebMessageBodyStyle.Bare)]
        Resultado RegistrarSumaDeTagEnRangoFecha(string tag, string tagsuma, string fechaini, string fechafin, string fechasuma);

       
        [OperationContract]
        [WebGet(UriTemplate = "/EjecutarCalculosSumarizadosDeAF?fecha={fecha}", ResponseFormat = WebMessageFormat.Json, BodyStyle = WebMessageBodyStyle.Bare)]
        Resultado EjecutarCalculosSumarizadosDeAF(string fecha);


        [OperationContract]
        [WebGet(UriTemplate = "/CaLcularStockOpenning?fecha={fecha}", ResponseFormat = WebMessageFormat.Json, BodyStyle = WebMessageBodyStyle.Bare)]
        Resultado CaLcularStockOpenning(string fecha);


        [OperationContract]
        [WebGet(UriTemplate = "/RegistrarTotalMMBTUTON?fecha={fecha}", ResponseFormat = WebMessageFormat.Json, BodyStyle = WebMessageBodyStyle.Bare)]
        Resultado RegistrarTotalMMBTUTON(string fecha);


        [OperationContract]
        [WebGet(UriTemplate = "/ListarZtonsbudget", ResponseFormat = WebMessageFormat.Json, BodyStyle = WebMessageBodyStyle.Bare)]
        List<ZtonsbudgetRow> ListarZtonsbudget();


        [OperationContract]
        [WebGet(UriTemplate = "/ListarBudget", ResponseFormat = WebMessageFormat.Json, BodyStyle = WebMessageBodyStyle.Bare)]
        List<Budget> ListarBudget();


        [OperationContract]
        [WebGet(UriTemplate = "/ListarMonthlyBudget?fechaini={fechaini}&fechafin={fechafin}", ResponseFormat = WebMessageFormat.Json, BodyStyle = WebMessageBodyStyle.Bare)]
        List<MonthlyBudget> ListarMonthlyBudget(string fechaini, string fechafin);


        [OperationContract]
        [WebGet(UriTemplate = "/RegistrarMonthlyBudget?fecha={fecha}", ResponseFormat = WebMessageFormat.Json, BodyStyle = WebMessageBodyStyle.Bare)]
        Resultado RegistrarMonthlyBudget(string fecha);


        [OperationContract]
        [WebGet(UriTemplate = "/ListarCRinventory?fechaini={fechaini}&fechafin={fechafin}", ResponseFormat = WebMessageFormat.Json, BodyStyle = WebMessageBodyStyle.Bare)]
        List<CRinventoryItem> ListarCRinventory(string fechaini, string fechafin);


        [OperationContract]
        [WebGet(UriTemplate = "/RegistrarCRinventory?fecha={fecha}", ResponseFormat = WebMessageFormat.Json, BodyStyle = WebMessageBodyStyle.Bare)]
        Resultado RegistrarCRinventory(string fecha);


        [OperationContract]
        [WebGet(UriTemplate = "/ListarAnosMasterData", ResponseFormat = WebMessageFormat.Json, BodyStyle = WebMessageBodyStyle.Bare)]
        List<int> ListarAnosMasterData();


        [OperationContract]
        [WebGet(UriTemplate = "/ListarPeriodosMasterData?ano={ano}", ResponseFormat = WebMessageFormat.Json, BodyStyle = WebMessageBodyStyle.Bare)]
        List<PeriodoMasterData> ListarPeriodosMasterData(int ano);


        [OperationContract]
        [WebInvoke(UriTemplate = "/RegistrarPeriodoMasterData", RequestFormat = WebMessageFormat.Json, ResponseFormat = WebMessageFormat.Json,
        BodyStyle = WebMessageBodyStyle.Bare, Method = "POST")]
        Resultado RegistrarPeriodoMasterData(EntradaRegistrarPeriodoMasterData oEntradaRegistrarPeriodoMasterData);


        [OperationContract]
        [WebGet(UriTemplate = "/ListarPeriodosMasterDataEnRangoFecha?fechaini={fechaini}&fechafin={fechafin}", ResponseFormat = WebMessageFormat.Json, BodyStyle = WebMessageBodyStyle.Bare)]
        List<PeriodoMasterData> ListarPeriodosMasterDataEnRangoFecha(string fechaini, string fechafin);


        [OperationContract]
        [WebGet(UriTemplate = "/RegistrarRawMaterial?fecha={fecha}", ResponseFormat = WebMessageFormat.Json, BodyStyle = WebMessageBodyStyle.Bare)]
        Resultado RegistrarRawMaterial(string fecha);


        [OperationContract]
        [WebGet(UriTemplate = "/EjecutarCalculosSumarizadosDeAFMensual?fecha={fecha}", ResponseFormat = WebMessageFormat.Json, BodyStyle = WebMessageBodyStyle.Bare)]
        Resultado EjecutarCalculosSumarizadosDeAFMensual(string fecha);


        [OperationContract]
        [WebGet(UriTemplate = "/ObtenerRDP?fecha={fecha}", ResponseFormat = WebMessageFormat.Json, BodyStyle = WebMessageBodyStyle.Bare)]
        RDPReporte ObtenerRDP(string fecha);


        [OperationContract]
        [WebInvoke(UriTemplate = "/InsertarNotificacion", RequestFormat = WebMessageFormat.Json, ResponseFormat = WebMessageFormat.Json,
        BodyStyle = WebMessageBodyStyle.Bare, Method = "POST")]
        Resultado InsertarNotificacion(RDPNotificacion oRDPNotificacion);


        [OperationContract]
        [WebGet(UriTemplate = "/EjecutarCalculosBasicosDeAF?fecha={fecha}", ResponseFormat = WebMessageFormat.Json, BodyStyle = WebMessageBodyStyle.Bare)]
        Resultado EjecutarCalculosBasicosDeAF(string fecha);


        [OperationContract]
        [WebGet(UriTemplate = "/ListarCRShift?fechaini={fechaini}&fechafin={fechafin}", ResponseFormat = WebMessageFormat.Json, BodyStyle = WebMessageBodyStyle.Bare)]
        List<CRShift> ListarCRShift(string fechaini, string fechafin);


        [OperationContract]
        [WebGet(UriTemplate = "/RegistrarCRShift?fecha={fecha}", ResponseFormat = WebMessageFormat.Json, BodyStyle = WebMessageBodyStyle.Bare)]
        Resultado RegistrarCRShift(string fecha);


        [OperationContract]
        [WebGet(UriTemplate = "/ObtenerRDPEnergia?fecha={fecha}", ResponseFormat = WebMessageFormat.Json, BodyStyle = WebMessageBodyStyle.Bare)]
        RDPReporte ObtenerRDPEnergia(string fecha);


        [OperationContract]
        [WebGet(UriTemplate = "/ListarDeltasDeFormulacionParaReportesFinales?template={template}&fechaini={fechaini}&fechafin={fechafin}&maquina={maquina}", ResponseFormat = WebMessageFormat.Json, BodyStyle = WebMessageBodyStyle.Bare)]
        List<Formulacion> ListarDeltasDeFormulacionParaReportesFinales(string template, string fechaini, string fechafin, string maquina);


        [OperationContract]
        [WebGet(UriTemplate = "/RegistrarOP?fecha={fecha}", ResponseFormat = WebMessageFormat.Json, BodyStyle = WebMessageBodyStyle.Bare)]
        Resultado RegistrarOP(string fecha);


        [OperationContract]
        [WebGet(UriTemplate = "/ListarOP?fecha={fecha}", ResponseFormat = WebMessageFormat.Json, BodyStyle = WebMessageBodyStyle.Bare)]
        List<OrdenProcesoMaterial> ListarOP(string fecha);


    }




    [DataContract]
    public class OrdenProcesoMaterial
    {
        [DataMember]
        public string Fecha { get; set; }

        [DataMember]
        public string Tag { get; set; }

        [DataMember]
        public string MaterialSAP { get; set; }

        [DataMember]
        public string PuestoTrabajo { get; set; }

        [DataMember]
        public string Centro { get; set; }

        [DataMember]
        public string Comentario { get; set; }

        [DataMember]
        public string NumOP { get; set; }

        [DataMember]
        public string Resultado { get; set; }    

    }


    [DataContract]
    public class ConsultaObtenerOrdenProceso
    {
        [DataMember]
        public string codCentroLogistico { get; set; }

        [DataMember]
        public string codPuestoTrabajo { get; set; }

        [DataMember]
        public string numMaterial { get; set; }

        [DataMember]
        public string fecInicio { get; set; }
    }


    [DataContract]
    public class RespuestaObtenerOrdenProceso
    {
        [DataMember]
        public string CodResultado { get; set; }

        [DataMember]
        public string DscResultado { get; set; }

        [DataMember]
        public string NumOP { get; set; }

    }

    
    [DataContract]
    public class Parada
    {
        [DataMember]
        public string Id;

        [DataMember]
        public string ElementoPrimario;

        [DataMember]
        public string TipoParada;

        [DataMember]
        public bool TransferenciaSAP;

        [DataMember]
        public double Duracion;

        [DataMember]
        public string Inicio;

        [DataMember]
        public string Fin;
    }


    [DataContract]
    public class CRShift
    {
        [DataMember(Order = 1)]
        public string Operator;

        [DataMember(Order = 2)]
        public string Supervisor;

        [DataMember(Order = 3)]
        public string ShiftStart;

        [DataMember(Order = 4)]
        public string ShiftEnd;
    }


    [DataContract]
    public class RDPNotificacion
    {
        [DataMember(Order = 5)]
        public string NOM_ELEMENTO;

        [DataMember(Order = 14)]
        public string FCH_NOTIFICACION;

        [DataMember(Order = 14)]
        public string FCH_PRODUCCION;
       
        [DataMember(Order = 7)]
        public string NUM_NOTIFICACION;

        [DataMember(Order = 8)]
        public string NUM_CONTADOR;

        [DataMember(Order = 10)]
        public string COD_ESTADO;

        [DataMember(Order = 11)]
        public string DSC_RESPUESTA;

        [DataMember(Order = 12)]
        public string COD_RESPUESTA;

        [DataMember(Order = 13)]
        public string USR_NOTIFICACION;    

        [DataMember(Order = 18)]
        public string DSC_RESPUESTA_DETALLE;

        [DataMember(Order = 19)]
        public int NUM_INTENTOS;

        [DataMember(Order = 18)]
        public string DSC_TAG_NOTIFICACION;

        [DataMember(Order = 19)]
        public string NRO_DUPLICADO;
    }


    [DataContract]
    public class ElementoValor
    {
        [DataMember(Order = 1)]
        public string Elemento;

        [DataMember(Order = 2)]
        public double Valor;
    }


    [DataContract]
    public class RDPReporte
    {
        [DataMember(Order = 1)]
        public string Fecha;

        [DataMember(Order = 2)]
        public List<RDPProceso> oProcesos;
    }


    [DataContract]
    public class RDPProceso
    {
        [DataMember(Order = 0)]
        public string FCH_PRODUCCION;

        [DataMember(Order = 1)]
        public string COD_GRUPO;

        //[DataMember(Order = 5)]
        //public string CNT_HORAS;

        //[DataMember(Order = 6)]
        //public string CNT_TONELADAS;

        //[DataMember(Order = 7)]
        //public string CNT_TONELADASNOTIF;

        //[DataMember(Order = 8)]
        //public string TH;

        //[DataMember(Order = 9)]
        //public string DOSIFICACION;

        //[DataMember(Order = 10)]
        //public int TIPO;

        //[DataMember(Order = 11)]
        //public int ESTOTAL;

        //[DataMember(Order = 12)]
        //public bool ConAlerta;

        //[DataMember(Order = 13)]
        //public string MensajeAlerta;

        //[DataMember(Order = 14)]
        //public string IMP_COSTO_UNIT;

        //[DataMember(Order = 15)]
        //public string IMP_COSTO_TOTAL;

        //[DataMember(Order = 16)]
        //public string CNT_FACTOR;

        [DataMember(Order = 17)]
        public int NUM_ORDEN;

        [DataMember(Order = 18)]
        public List<RDPMaquina> MAQUINAS;

    }

    [DataContract]
    public class RDPMaquina
    {
        [DataMember(Order = 0)]
        public string FCH_PRODUCCION;

        [DataMember(Order = 1)]
        public string COD_GRUPO;

        [DataMember(Order = 2)]
        public string COD_MAQUINA;

        //[DataMember(Order = 5)]
        //public string CNT_HORAS;

        //[DataMember(Order = 6)]
        //public string CNT_TONELADAS;

        //[DataMember(Order = 7)]
        //public string CNT_TONELADASNOTIF;

        //[DataMember(Order = 8)]
        //public string TH;

        //[DataMember(Order = 9)]
        //public string DOSIFICACION;

        //[DataMember(Order = 10)]
        //public int TIPO;

        //[DataMember(Order = 11)]
        //public int ESTOTAL;

        //[DataMember(Order = 12)]
        //public bool ConAlerta;

        //[DataMember(Order = 13)]
        //public string MensajeAlerta;

        //[DataMember(Order = 14)]
        //public string IMP_COSTO_UNIT;

        //[DataMember(Order = 15)]
        //public string IMP_COSTO_TOTAL;

        //[DataMember(Order = 16)]
        //public string CNT_FACTOR;

        [DataMember(Order = 17)]
        public int NUM_ORDEN;

        //[DataMember(Order = 18)]
        //public string DSC_COLOR_TH;

        //[DataMember(Order = 19)]
        //public string COD_PUESTO_TRABAJO;

        //[DataMember(Order = 20)]
        //public string CNT_TH_LI;

        //[DataMember(Order = 21)]
        //public string CNT_TH_LS;

        //[DataMember(Order = 22)]
        //public string DSC_COLOR_TH_LI;

        //[DataMember(Order = 23)]
        //public string DSC_COLOR_TH_LC;

        //[DataMember(Order = 24)]
        //public string DSC_COLOR_TH_LS;

        [DataMember(Order = 25)]
        public List<RDPProducto> PRODUCTOS;

    }

    [DataContract]
    public class RDPProducto
    {
        [DataMember(Order = 0)]
        public string FCH_PRODUCCION;

        [DataMember(Order = 1)]
        public string COD_GRUPO;

        [DataMember(Order = 2)]
        public string COD_MAQUINA;

        [DataMember(Order = 3)]
        public string COD_PRODUCTO;

        [DataMember(Order = 5)]
        public string CNT_HORAS;

        [DataMember(Order = 6)]
        public string CNT_TONELADAS;

        //[DataMember(Order = 7)]
        //public string CNT_TONELADASNOTIF;

        //[DataMember(Order = 8)]
        //public string TH;

        //[DataMember(Order = 9)]
        //public string DOSIFICACION;

        //[DataMember(Order = 10)]
        //public int TIPO;

        //[DataMember(Order = 11)]
        //public int ESTOTAL;

        //[DataMember(Order = 12)]
        //public bool ConAlerta;

        //[DataMember(Order = 13)]
        //public string MensajeAlerta;

        //[DataMember(Order = 14)]
        //public string IMP_COSTO_UNIT;

        //[DataMember(Order = 15)]
        //public string IMP_COSTO_TOTAL;

        //[DataMember(Order = 16)]
        //public string CNT_FACTOR;

        [DataMember(Order = 17)]
        public int NUM_ORDEN;

        [DataMember(Order = 18)]
        public string DSC_OP;

        [DataMember(Order = 19)]
        public string COD_MATERIAL_OP;

        [DataMember(Order = 20)]
        public string NUM_OP;




        

        [DataMember(Order = 22)]
        public string NUM_NOTIFICACION;

        [DataMember(Order = 23)]
        public string NUM_CONTADOR;

        [DataMember(Order = 24)]
        public string COD_ESTADO;

        [DataMember(Order = 25)]
        public string DSC_RESPUESTA;

        [DataMember(Order = 26)]
        public string COD_RESPUESTA;

        [DataMember(Order = 27)]
        public string USR_NOTIFICACION;

        [DataMember(Order = 28)]
        public string FCH_NOTIFICACION;

        //[DataMember(Order = 29)]
        //public string CNT_TONELADAS_SAP;


        //[DataMember(Order = 30)]
        //public string CNT_GAS;

        //[DataMember(Order = 31)]
        //public string DSC_UNIDAD_GAS;

        //[DataMember(Order = 32)]
        //public string CNT_TONELADAS_ORIGINAL;

        //[DataMember(Order = 33)]
        //public string CNT_TH_LI;

        //[DataMember(Order = 34)]
        //public string CNT_TH_LS;

        //[DataMember(Order = 35)]
        //public string DSC_COLOR_TH_LI;

        //[DataMember(Order = 36)]
        //public string DSC_COLOR_TH_LC;

        //[DataMember(Order = 37)]
        //public string DSC_COLOR_TH_LS;

        //[DataMember(Order = 38)]
        //public string DSC_COLOR_TH;

        [DataMember(Order = 39)]
        public string COD_ALMACEN;

        [DataMember(Order = 40)]
        public string DSC_RESPUESTA_DETALLE;

        [DataMember(Order = 41)]
        public int NUM_INTENTOS;

        [DataMember(Order = 41)]
        public string COD_PUESTO_TRABAJO;

        [DataMember(Order = 42)]
        public string COD_UNIDAD;

        [DataMember(Order = 43)]
        public string COD_UNIDAD_HORAS;

        [DataMember(Order = 44)]
        public string FLG_GAS;

        [DataMember(Order = 45)]
        public string FLG_COAL;

        [DataMember(Order = 46)]
        public string FLG_HORA;

        [DataMember(Order = 47)]
        public string FLG_ENERGIA;

        [DataMember(Order = 48)]
        public string FLG_SIN_MATERIALES;

        [DataMember(Order = 49)]
        public string DSC_TAG_NOTIFICACION;

        [DataMember(Order = 50)]
        public string VORNR;

        [DataMember(Order = 51)]
        public string DSC_PRODUCTO;

        [DataMember(Order = 52)]
        public string NRO_DUPLICADO;

        [DataMember(Order = 53)]
        public List<RDPMaterial> MATERIALES;

    }

    [DataContract]
    public class RDPMaterial
    {
        [DataMember(Order = 0)]
        public string FCH_PRODUCCION;

        [DataMember(Order = 1)]
        public string COD_GRUPO;

        [DataMember(Order = 2)]
        public string COD_MAQUINA;

        [DataMember(Order = 3)]
        public string COD_PRODUCTO;

        [DataMember(Order = 4)]
        public string COD_MATERIAL;

        //[DataMember(Order = 5)]
        //public string CNT_HORAS;

        //[DataMember(Order = 6)]
        //public string CNT_TONELADAS;

        [DataMember(Order = 7)]
        public string CNT_TONELADASNOTIF;

        //[DataMember(Order = 8)]
        //public string TH;

        //[DataMember(Order = 9)]
        //public string DOSIFICACION;

        //[DataMember(Order = 10)]
        //public int TIPO;

        //[DataMember(Order = 11)]
        //public int ESTOTAL;

        //[DataMember(Order = 12)]
        //public bool ConAlerta;

        //[DataMember(Order = 13)]
        //public string MensajeAlerta;

        //[DataMember(Order = 14)]
        //public string IMP_COSTO_UNIT;

        //[DataMember(Order = 15)]
        //public string IMP_COSTO_TOTAL;

        //[DataMember(Order = 16)]
        //public string CNT_FACTOR;

        //[DataMember(Order = 17)]
        //public int NUM_ORDEN;

        //[DataMember(Order = 18)]
        //public string COD_GRUPO_MATERIAL;

        //[DataMember(Order = 19)]
        //public string CNT_TONELADAS_SAP;

        //[DataMember(Order = 20)]
        //public string CNT_TONELADAS_ORIGINAL;

        //[DataMember(Order = 21)]
        //public string CNT_PORCENTAJE;

        [DataMember(Order = 22)]
        public List<RDPMaterialSAP> MATERIALES_SAP;

        //[DataMember(Order = 23)]
        //public List<RDPMaterial> MATERIALES;

    }

    [DataContract]
    public class RDPMaterialSAP
    {
        [DataMember(Order = 0)]
        public string FCH_PRODUCCION;

        [DataMember(Order = 1)]
        public string COD_GRUPO;

        [DataMember(Order = 2)]
        public string COD_MAQUINA;

        [DataMember(Order = 3)]
        public string COD_PRODUCTO;

        [DataMember(Order = 4)]
        public string COD_MATERIAL;

        [DataMember(Order = 5)]
        public string COD_MATERIAL_SAP;

        [DataMember(Order = 6)]
        public string DSC_MATERIAL_SAP;

        [DataMember(Order = 7)]
        public string CNT_TONELADASNOTIF;

        //[DataMember(Order = 8)]
        //public string CNT_PORCENTAJE;

        //[DataMember(Order = 9)]
        //public int TIPO;

        [DataMember(Order = 10)]
        public int NUM_ORDEN;

        [DataMember(Order = 11)]
        public string COD_UNIDAD;

        //[DataMember(Order = 12)]
        //public string CNT_TONELADAS_SAP;

        //[DataMember(Order = 13)]
        //public string CNT_TONELADAS_ORIGINAL;

        [DataMember(Order = 14)]
        public string COD_ALMACEN;

    }



    [DataContract]
    public class EntradaRegistrarPeriodoMasterData
    {
        [DataMember(Order = 1)]
        public string Periodo;

        [DataMember(Order = 2)]
        public int Estado;

        [DataMember(Order = 3)]
        public string UsuarioModificacion;
    }


    [DataContract]
    public class PeriodoMasterData
    {
        [DataMember(Order = 1)]
        public string Periodo;

        [DataMember(Order = 2)]
        public int Estado;

        [DataMember(Order = 3)]
        public string NombreEstado;

        [DataMember(Order = 4)]
        public string Ano;

        [DataMember(Order = 5)]
        public string Mes;
    }


    [DataContract]
    public class CRinventoryItem
    {
        [DataMember(Order = 1)]
        public string fecha;

        [DataMember(Order = 2)]
        public string nombre;

        [DataMember(Order = 3)]
        public double valor;
    }


    [DataContract]
    public class MonthlyBudget
    {     
        [DataMember(Order = 1)]
        public string PERIO;

        [DataMember(Order = 2)]
        public double TONS;
    }


    [DataContract]
    public class Budget
    {
        [DataMember(Order = 0)]
        public string NAME1;

        [DataMember(Order = 1)]
        public string PERIO;

        [DataMember(Order = 2)]
        public double TONS;

        [DataMember(Order = 3)]
        public string KUNAG;
    }


    [DataContract]
    public class ZtonsbudgetRow
    {
        [DataMember(Order = 0)]
        public string Dato1;

        [DataMember(Order = 1)]
        public string Dato2;
    }


    [DataContract]
    public class FAC
    {
         [DataMember(Order = 0)]
         public List<CabeceraFAC> Cabeceras;

         [DataMember(Order = 1)]
         public List<ElementoFAC> Elementos;
    }


    [DataContract]
    public class CabeceraFAC
     {
         [DataMember(Order = 1)]
         public string Titulo;

         [DataMember(Order = 2)]
         public int Colspan;

         [DataMember(Order = 3)]
         public int Orden;
     }

    [DataContract]
    public class EntradaRestaurarDelta
    {
        [DataMember(Order = 1)]
        public string TagDeltaManual;

        [DataMember(Order = 2)]
        public string TagDeltaOriginal;

        [DataMember(Order = 3)]
        public string Fecha;
    }


    [DataContract]
    public class ElementoFAC
    {
        [DataMember(Order = 1)]
        public string Fecha;

        [DataMember(Order = 2)]
        public string LHV_Pulv_Fuel_Btu;

        [DataMember(Order = 3)]
        public string LHV_NaturalGas_Btu;

        [DataMember(Order = 4)]
        public string _41X_CLISI2_Total;

        [DataMember(Order = 5)]
        public string _41X_CLISI3_Total;

        [DataMember(Order = 6)]
        public string _41X_CLISI100_Total;

        [DataMember(Order = 7)]
        public string _41X_CLISI10_Total;

        [DataMember(Order = 8)]
        public string _41X_CLISI55_Total;

        [DataMember(Order = 9)]
        public string _41X_CLISI5_Total;

        [DataMember(Order = 10)]
        public string _41X_CLIS12_Total;

        [DataMember(Order = 11)]
        public string Clinker_Production_MTD;

        [DataMember(Order = 12)]
        public string Clinker_Production_YTD;

        [DataMember(Order = 13)]
        public string Process_MMBTU_Today;

        [DataMember(Order = 14)]
        public string Process_MMBTU_MTD;

        [DataMember(Order = 15)]
        public string Process_MMBTU_YTD;

        [DataMember(Order = 16)]
        public string Total_MMBTU_Today;

        [DataMember(Order = 17)]
        public string Total_MMBTU_MTD;

        [DataMember(Order = 18)]
        public string Total_MMBTU_YTD;

        [DataMember(Order = 19)]
        public string Process_MMBTU_Ton_Today;

        [DataMember(Order = 20)]
        public string Process_MMBTU_Ton_MTD;

        [DataMember(Order = 21)]
        public string Process_MMBTU_Ton_YTD;

        [DataMember(Order = 22)]
        public string Total_MMBTU_Ton_Today;

        [DataMember(Order = 23)]
        public string Total_MMBTU_Ton_MTD;

        [DataMember(Order = 24)]
        public string Total_MMBTU_Ton_YTD;

    }

    [DataContract]
    public class MasterData
    {
        [DataMember(Order = 1)]
        public string Name;

        [DataMember(Order = 2)]
        public string Title;
    }

    [DataContract]
    public class EntradaActualizarFactorDeFormulacion
    {
        [DataMember(Order = 1)]
        public string Tag;

        [DataMember(Order = 2)]
        public string Fecha;

        [DataMember(Order = 3)]
        public string Valor;

        [DataMember(Order = 4)]
        public string RutaElemento;

        [DataMember(Order = 5)]
        public int Rounding;
    }

    [DataContract]
    public class FactorFormulacion
    {
        [DataMember(Order = 1)]
        public string Nombre;

        [DataMember(Order = 2)]
        public string Titulo;

        [DataMember(Order = 3)]
        public string Tag;

        [DataMember(Order = 4)]
        public string Dependencia;

        [DataMember(Order = 5)]
        public string Fecha;

        [DataMember(Order = 6)]
        public string Valor;

        [DataMember(Order = 7)]
        public string Ruta;

        [DataMember(Order = 8)]
        public int Rounding;
    }


    [DataContract]
    public class DependenciaDOR
    {
        [DataMember(Order = 1)]
        public string Codigo;

        [DataMember(Order = 2)]
        public string Descripcion;
    }


    [DataContract]
    public class ElementoDOR
    {
        [DataMember(Order = 1)]
        public string Padre;

        [DataMember(Order = 2)]
        public string Nombre;

        [DataMember(Order = 3)]
        public string Origen;

        [DataMember(Order = 4)]
        public string ValorDia;

        [DataMember(Order = 5)]
        public string ValorMes;

        [DataMember(Order = 6)]
        public string ValorAno;

        [DataMember(Order = 7)]
        public string Unidad;

        [DataMember(Order = 8)]
        public string ValorDia2;

        [DataMember(Order = 9)]
        public string ValorMes2;

        [DataMember(Order = 10)]
        public string ValorAno2;

        [DataMember(Order = 11)]
        public string Unidad2;

        [DataMember(Order = 12)]
        public string ValorDia3;

        [DataMember(Order = 13)]
        public string ValorMes3;

        [DataMember(Order = 14)]
        public string ValorDia4;

        [DataMember(Order = 15)]
        public int Orden;

        [DataMember(Order = 16)]
        public string CssClass;
    }


    [DataContract]
    public class BloqueDOR
    {
        [DataMember(Order = 1)]
        public string Nombre;

        [DataMember(Order = 2)]
        public int Orden;

        [DataMember(Order = 3)]
        public string CssClass;

        [DataMember(Order = 4)]
        public bool EsUsagePorc;

        [DataMember(Order = 5)]
        public List<ElementoDOR> Elementos;
    }


    [DataContract]
    public class EntradaActualizarDeltaManual
    {
        [DataMember(Order = 1)]
        public string TagDeltaManual;

        [DataMember(Order = 2)]
        public string Fecha;

        [DataMember(Order = 3)]
        public string Valor;

        [DataMember(Order = 4)]
        public int Rounding;

        [DataMember(Order = 5)]
        public string UsuarioModificacion;

        [DataMember(Order = 6)]
        public string Comentarios;
    }


    [DataContract]
    public class Maquina
    {
        [DataMember]
        public string Nombre;

        [DataMember]
        public string Dependencia;

        [DataMember]
        public int Orden;

        [DataMember]
        public string Descripcion;


    }


    [DataContract]
    public class Formulacion
    {
        [DataMember(Order = 1)]
        public string Nombre;

         [DataMember(Order = 2)]
         public string TagContador;

         [DataMember(Order = 3)]
         public string TagDelta;

         [DataMember(Order = 4)]
         public string TagDeltaManual;

         [DataMember(Order = 5)]
         public double Factor;

         [DataMember(Order = 6)]
         public string Titulo;

         [DataMember(Order = 7)]
         public int Orden;

         [DataMember(Order = 8)]
         public string EsFactor;

         [DataMember(Order = 9)]
         public int Rounding;

         [DataMember(Order = 10)]
         public int OrdenExport;

         [DataMember(Order = 11)]
         public List<Delta> Deltas;
    }


    [DataContract]
    public class Delta
    {
        [DataMember(Order = 1)]
        public string Fecha;

        [DataMember(Order = 2)]
        public string Valor;

        [DataMember(Order = 3)]
        public double Factor;

        [DataMember(Order = 4)]
        public string Modificado;

        [DataMember(Order = 5)]
        public string UsuarioModificacion;

        [DataMember(Order = 6)]
        public string Comentarios;

        [DataMember(Order = 7)]
        public string ValorOriginal;

        [DataMember(Order = 8)]
        public bool EsEditable;
    }


    [DataContract]
    public class EventoFull
    {
        [DataMember]
        public string cId;

        [DataMember]
        public string cSeveridad;

        [DataMember]
        public string cNombre;

        [DataMember]
        public string cInicio;

        [DataMember]
        public string cFin;

        [DataMember]
        public double nDuracion;

        [DataMember]
        public string cDescripcion;

        [DataMember]
        public string cCategoria;

        [DataMember]
        public string cCadena;

        [DataMember]
        public string cElementoPrimario;

        [DataMember]
        public string cRutaElementoPrimario;

        [DataMember]
        public string cAnalisis;

        [DataMember]
        public string cPlantilla;

        [DataMember]
        public List<Atributo> oAtributos;

        [DataMember]
        public string cArea;

        [DataMember]
        public string cNumeroAviso;

        [DataMember]
        public int nOrden;

        [DataMember]
        public double nDuracionTurno1;

        [DataMember]
        public double nDuracionTurno2;

        [DataMember]
        public double nDuracionTurno3;

        [DataMember]
        public double nDuracionHorasTurno1;

        [DataMember]
        public double nDuracionHorasTurno2;

        [DataMember]
        public double nDuracionHorasTurno3;

    }


    [DataContract]
    public class Atributo
    {
        [DataMember]
        public string cNombre;

        [DataMember]
        public string cValor;

        [DataMember]
        public string cUOM;

        [DataMember]
        public string cTag;

        [DataMember]
        public string cUrl;



    }


    [DataContract]
    public class Resultado
    {
        [DataMember]
        public int codigo;

        [DataMember]
        public string descripcion;

        public Resultado()
        {
        }

        public Resultado(int codigo, string descripcion)
        {
            this.codigo = codigo;
            this.descripcion = descripcion;
        }
    }
}
