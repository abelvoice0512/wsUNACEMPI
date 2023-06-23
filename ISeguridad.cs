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
    // NOTE: You can use the "Rename" command on the "Refactor" menu to change the interface name "ISeguridad" in both code and config file together.
    [ServiceContract]
    public interface ISeguridad
    {
        [OperationContract]
        [WebGet(UriTemplate = "/ListarGruposDeUsuarioEnAD?usuario={usuario}", ResponseFormat = WebMessageFormat.Json, BodyStyle = WebMessageBodyStyle.Bare)]
        List<string> ListarGruposDeUsuarioEnAD(string usuario);

        [OperationContract]
        [WebGet(UriTemplate = "/ListarMiembrosDeGrupoEnServidor?grupo={grupo}", ResponseFormat = WebMessageFormat.Json, BodyStyle = WebMessageBodyStyle.Bare)]
        List<string> ListarMiembrosDeGrupoEnServidor(string grupo);

        [OperationContract]
        [WebGet(UriTemplate = "/UsuarioPerteneceAlGrupo?usuario={usuario}&grupo={grupo}", ResponseFormat = WebMessageFormat.Json, BodyStyle = WebMessageBodyStyle.Bare)]
        bool UsuarioPerteneceAlGrupo(string usuario, string grupo);
    }
}
