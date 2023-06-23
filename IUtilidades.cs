using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.ServiceModel;
using System.ServiceModel.Web;
using System.Text;
using System.Data;
using System.IO; 

namespace wsUNACEMPI
{
    // NOTE: You can use the "Rename" command on the "Refactor" menu to change the interface name "IUtilidades" in both code and config file together.
    [ServiceContract]
    public interface IUtilidades
    {
        //metodos utilidad migrar tags

        [OperationContract]
        [WebGet(UriTemplate = "/ObtenerTag?servidor={servidor}&tag={tag}", ResponseFormat = WebMessageFormat.Json, BodyStyle = WebMessageBodyStyle.Bare)]
        Tag ObtenerTag(string servidor, string tag);


        [OperationContract]
        [WebGet(UriTemplate = "/ListarTagValores?servidor={servidor}&tag={tag}&fechaini={fechaini}&fechafin={fechafin}", ResponseFormat = WebMessageFormat.Json, BodyStyle = WebMessageBodyStyle.Bare)]
        List<TagValor> ListarTagValores(string servidor, string tag, string fechaini, string fechafin);


        [OperationContract]
        [WebInvoke(UriTemplate = "/MigrarTag", RequestFormat = WebMessageFormat.Json, ResponseFormat = WebMessageFormat.Json,
        BodyStyle = WebMessageBodyStyle.Bare, Method = "POST")]
        Resultado MigrarTag(EntradaMigrarTag EntradaMigrarTag);


        [OperationContract]
        [WebGet(UriTemplate = "/BuscarTags?servidor={servidor}&query={query}", ResponseFormat = WebMessageFormat.Json, BodyStyle = WebMessageBodyStyle.Bare)]
        List<Tag> BuscarTags(string servidor, string query);



        //metodos utilidad importacion datos de excel a PI

        [OperationContract]
        [WebGet(UriTemplate = "/ListarHojasDeExcel?url={url}", ResponseFormat = WebMessageFormat.Json, BodyStyle = WebMessageBodyStyle.Bare)]
        List<HojaExcel> ListarHojasDeExcel(string url);


        [OperationContract]
        [WebInvoke(UriTemplate = "/ImportarExcelAPI", RequestFormat = WebMessageFormat.Json, ResponseFormat = WebMessageFormat.Json,
        BodyStyle = WebMessageBodyStyle.Bare, Method = "POST")]
        Resultado ImportarExcelAPI(EntradaImportarExcelAPI EntradaImportarExcelAPI);


        [OperationContract]
        [WebInvoke(UriTemplate = "/ImportarExcelAPI2", RequestFormat = WebMessageFormat.Json, ResponseFormat = WebMessageFormat.Json,
        BodyStyle = WebMessageBodyStyle.Bare, Method = "POST")]
        Resultado ImportarExcelAPI2(EntradaImportarExcelAPI2 EntradaImportarExcelAPI2);



        //metodos utilidad importacion de excel configurados en AF
        [OperationContract]
        [WebGet(UriTemplate = "/ImportarExcelDeAF?nombre={nombre}&fecha={fecha}&usuario={usuario}", ResponseFormat = WebMessageFormat.Json, BodyStyle = WebMessageBodyStyle.Bare)]
        Resultado ImportarExcelDeAF(string nombre, string fecha, string usuario);


        [OperationContract]
        [WebGet(UriTemplate = "/ImportarTodoExcelDeAF?fecha={fecha}&usuario={usuario}", ResponseFormat = WebMessageFormat.Json, BodyStyle = WebMessageBodyStyle.Bare)]
        Resultado ImportarTodoExcelDeAF(string fecha, string usuario);


        [OperationContract]
        [WebGet(UriTemplate = "/ListarExcelDeAF?template={template}", ResponseFormat = WebMessageFormat.Json, BodyStyle = WebMessageBodyStyle.Bare)]
        List<ExcelAF> ListarExcelDeAF(string template);


        [OperationContract]
        [WebGet(UriTemplate = "/ImportarExcelMensualDeAF?nombre={nombre}&fecha={fecha}&usuario={usuario}", ResponseFormat = WebMessageFormat.Json, BodyStyle = WebMessageBodyStyle.Bare)]
        Resultado ImportarExcelMensualDeAF(string nombre, string fecha, string usuario);


        [OperationContract]
        [WebGet(UriTemplate = "/ImportarExcelMensualDeAFCompleto?nombre={nombre}&fecha={fecha}&usuario={usuario}", ResponseFormat = WebMessageFormat.Json, BodyStyle = WebMessageBodyStyle.Bare)]
        Resultado ImportarExcelMensualDeAFCompleto(string nombre, string fecha, string usuario);


        [OperationContract]
        [WebGet(UriTemplate = "/ImportarTodoExcelDeAFMensual?fecha={fecha}&usuario={usuario}", ResponseFormat = WebMessageFormat.Json, BodyStyle = WebMessageBodyStyle.Bare)]
        Resultado ImportarTodoExcelDeAFMensual(string fecha, string usuario);



        [OperationContract]
        [WebGet(UriTemplate = "/ImportarExcelMensualGenericoDeAF?nombre={nombre}&fecha={fecha}&usuario={usuario}", ResponseFormat = WebMessageFormat.Json, BodyStyle = WebMessageBodyStyle.Bare)]
        Resultado ImportarExcelMensualGenericoDeAF(string nombre, string fecha, string usuario);


        [OperationContract]
        [WebGet(UriTemplate = "/ImportarExcelMensualGenericoDeAFCompleto?nombre={nombre}&fecha={fecha}&usuario={usuario}", ResponseFormat = WebMessageFormat.Json, BodyStyle = WebMessageBodyStyle.Bare)]
        Resultado ImportarExcelMensualGenericoDeAFCompleto(string nombre, string fecha, string usuario);


        [OperationContract]
        [WebGet(UriTemplate = "/ImportarTodoExcelDeAFMensualGenerico?fecha={fecha}&usuario={usuario}", ResponseFormat = WebMessageFormat.Json, BodyStyle = WebMessageBodyStyle.Bare)]
        Resultado ImportarTodoExcelDeAFMensualGenerico(string fecha, string usuario);


        [OperationContract]
        [WebGet(UriTemplate = "/ImportarExcelAnualGenericoDeAF?nombre={nombre}&fecha={fecha}&usuario={usuario}", ResponseFormat = WebMessageFormat.Json, BodyStyle = WebMessageBodyStyle.Bare)]
        Resultado ImportarExcelAnualGenericoDeAF(string nombre, string fecha, string usuario);


        [OperationContract]
        [WebGet(UriTemplate = "/ImportarTodoExcelDeAFAnualGenerico?fecha={fecha}&usuario={usuario}", ResponseFormat = WebMessageFormat.Json, BodyStyle = WebMessageBodyStyle.Bare)]
        Resultado ImportarTodoExcelDeAFAnualGenerico(string fecha, string usuario);


        [OperationContract]
        [WebGet(UriTemplate = "/ImportarExcelGenericoDeAF?nombre={nombre}&usuario={usuario}", ResponseFormat = WebMessageFormat.Json, BodyStyle = WebMessageBodyStyle.Bare)]
        Resultado ImportarExcelGenericoDeAF(string nombre, string usuario);


        [OperationContract]
        [WebGet(UriTemplate = "/ImportarTodoExcelDeAFGenerico?usuario={usuario}", ResponseFormat = WebMessageFormat.Json, BodyStyle = WebMessageBodyStyle.Bare)]
        Resultado ImportarTodoExcelDeAFGenerico(string usuario);

    }


    [DataContract]
    public class ExcelAF
    {
        [DataMember]
        public string nombre { get; set; }
    }


    [DataContract]
    public class EntradaImportarExcelAPI2
    {
        [DataMember]
        public List<ConfiguracionExcel> Configuraciones { get; set; }

        [DataMember]
        public string UrlExcel { get; set; }

        [DataMember]
        public string HojaExcel { get; set; }

        [DataMember]
        public List<AnalisisPorEjecutar> Analisis { get; set; }


    }


    [DataContract]
    public class EntradaImportarExcelAPI
    {
        [DataMember]
        public List<ConfiguracionExcel> Configuraciones { get; set; }

        [DataMember]
        public string UrlExcel { get; set; }

        [DataMember]
        public int HojaExcel { get; set; }

        [DataMember]
        public List<AnalisisPorEjecutar> Analisis { get; set; }


    }


    [DataContract]
    public class AnalisisPorEjecutar
    {
        [DataMember]
        public string Database { get; set; }

        [DataMember]
        public string RutaElemento { get; set; }

        [DataMember]
        public string NombreAnalisis { get; set; }

        [DataMember]
        public string FechaIni { get; set; }

        [DataMember]
        public string FechaFin { get; set; }
    }


    [DataContract]
    public class ConfiguracionExcel
    {
        [DataMember]
        public string Tag { get; set; }

        [DataMember]
        public string CeldaTag { get; set; }

        [DataMember]
        public string Data { get; set; }

        [DataMember]
        public string Fecha { get; set; }

        [DataMember]
        public int FilaIni { get; set; }

        [DataMember]
        public int FilaFin { get; set; }

    }



    [DataContract]
    public class HojaExcel
    {
        [DataMember]
        public int Indice { get; set; }

        [DataMember]
        public string Nombre { get; set; }

    }




    [DataContract]
    public class EntradaMigrarTag
    {
        [DataMember]
        public string ServidorOrigen { get; set; }

        [DataMember]
        public string TagOrigen { get; set; }

        [DataMember]
        public string ServidorDestino { get; set; }

        [DataMember]
        public string TagDestino { get; set; }

        [DataMember]
        public string fechaini { get; set; }

        [DataMember]
        public string fechafin { get; set; }

    }


    [DataContract]
    public class Tag
    {
        [DataMember]
        public string Nombre { get; set; }

        [DataMember]
        public string Descripcion { get; set; }

        [DataMember]
        public string Tipo { get; set; }

        [DataMember]
        public bool Existe { get; set; }

    }


    [DataContract]
    public class TagValor
    {
        [DataMember]
        public string Nombre { get; set; }

        [DataMember]
        public string Valor { get; set; }

        [DataMember]
        public string Fecha { get; set; }


    }

}
