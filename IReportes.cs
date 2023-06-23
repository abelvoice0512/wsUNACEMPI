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
    // NOTE: You can use the "Rename" command on the "Refactor" menu to change the interface name "IReportes" in both code and config file together.
    [ServiceContract]
    public interface IReportes
    {
        [OperationContract]
        [WebGet(UriTemplate = "/ObtenerAreas", ResponseFormat = WebMessageFormat.Json, BodyStyle = WebMessageBodyStyle.Bare)]
        List<Area> ObtenerAreas();


        [OperationContract]
        [WebGet(UriTemplate = "/ObtenerReporteHorario?nombre={nombre}&fechaini={fechaini}&fechafin={fechafin}&periodo={periodo}", ResponseFormat = WebMessageFormat.Json, BodyStyle = WebMessageBodyStyle.Bare)]
        ReporteHorario ObtenerReporteHorario(string nombre, string fechaini, string fechafin, string periodo);


        [OperationContract]
        [WebGet(UriTemplate = "/ObtenerTiposReporte?template={template}", ResponseFormat = WebMessageFormat.Json, BodyStyle = WebMessageBodyStyle.Bare)]
        List<TipoReporte> ObtenerTiposReporte(string template);


        [OperationContract]
        [WebGet(UriTemplate = "/CaLcularClinkerAndCementInventory?fecha={fecha}", ResponseFormat = WebMessageFormat.Json, BodyStyle = WebMessageBodyStyle.Bare)]
        Resultado CaLcularClinkerAndCementInventory(string fecha);


        [OperationContract]
        [WebGet(UriTemplate = "/ObtenerReporteInventory?nombre={nombre}&fecha={fecha}", ResponseFormat = WebMessageFormat.Json, BodyStyle = WebMessageBodyStyle.Bare)]
        List<Inventory> ObtenerReporteInventory(string nombre, string fecha);

    }


    [DataContract]
    public class InventoryDetalle
    {
        [DataMember]
        public string Fecha { get; set; }
       
        [DataMember]
        public string InitialStock { get; set; }

        [DataMember]
        public string Production { get; set; }

        [DataMember]
        public string Consumption { get; set; }

        [DataMember]
        public string FinalStock { get; set; }        
    }

    
    [DataContract]
    public class Inventory
    {    
        [DataMember(Order = 1)]
        public string Nombre { get; set; }

        [DataMember(Order = 2)]
        public string Titulo { get; set; }

        [DataMember(Order = 3)]
        public string TituloInitialStock { get; set; }

        [DataMember(Order = 4)]
        public string TituloProduction { get; set; }

        [DataMember(Order = 5)]
        public string TituloConsumption { get; set; }

        [DataMember(Order = 6)]
        public string TituloFinalStock { get; set; }

        [DataMember(Order = 7)]
        public int Orden { get; set; }

        [DataMember(Order = 8)]
        public List<InventoryDetalle> Detalles { get; set; }
    }


    [DataContract]
    public class TipoReporte
    {
        [DataMember]
        public string Nombre { get; set; }

        [DataMember]
        public string Descripcion { get; set; }
    }

    [DataContract]
    public class ReporteHorario
    {
        [DataMember]
        public List<CabeceraReporteHorario> Cabeceras { get; set; }

        [DataMember]
        public List<FilaReporteHorario> Filas { get; set; }
    }

    [DataContract]
    public class CabeceraReporteHorario
    {
        [DataMember]
        public string Titulo1 { get; set; }

        [DataMember]
        public string Titulo2 { get; set; }

        [DataMember]
        public string Titulo3 { get; set; }

        [DataMember]
        public int Orden { get; set; }
    }

    [DataContract]
    public class FilaReporteHorario
    {
        [DataMember]
        public string Fecha { get; set; }

        [DataMember]
        public string Dato { get; set; }

        [DataMember]
        public int Orden { get; set; }

        [DataMember]
        public string Nombre { get; set; }
    }

    [DataContract]
    public class Area
    {
        [DataMember]
        public string Codigo { get; set; }

        [DataMember]
        public string Nombre { get; set; }

        [DataMember]
        public string NombreAbreviado { get; set; }

    }
}
