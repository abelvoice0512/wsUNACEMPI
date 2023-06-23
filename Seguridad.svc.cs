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
using System.Text.RegularExpressions;
using System.Web.Script.Serialization;
using System.DirectoryServices.AccountManagement;

namespace wsUNACEMPI
{
    // NOTE: You can use the "Rename" command on the "Refactor" menu to change the class name "Seguridad" in code, svc and config file together.
    // NOTE: In order to launch WCF Test Client for testing this service, please select Seguridad.svc or Seguridad.svc.cs at the Solution Explorer and start debugging.
    public class Seguridad : ISeguridad
    {
        public List<string> ListarGruposDeUsuarioEnAD(string usuario)
        {
            List<string> oGrupos = new List<string>();

            try
            {
                String cDominioImpersonation = WebConfigurationManager.AppSettings["DominioImpersonation"];
                String cUsuarioImpersonation = WebConfigurationManager.AppSettings["UsuarioImpersonation"];
                String cPasswordImpersonation = WebConfigurationManager.AppSettings["PasswordImpersonation"];

                WrapperImpersonationContext context = new WrapperImpersonationContext(cDominioImpersonation, cUsuarioImpersonation, cPasswordImpersonation);
                context.Enter();

                //establish domain context
                PrincipalContext oMiDominio = new PrincipalContext(ContextType.Domain);

                //find your user
                UserPrincipal oUserPrincipal = UserPrincipal.FindByIdentity(oMiDominio, usuario);

                if (oUserPrincipal != null)
                {
                    PrincipalSearchResult<Principal> oGruposPSR = oUserPrincipal.GetGroups();

                    for (int i = 0; i < oGruposPSR.Count(); i++)
                    {
                        oGrupos.Add(oGruposPSR.ElementAt(i).Name.ToUpper());
                    }
                }

                context.Leave();

                return oGrupos;
            }
            catch (Exception ex)
            {
                Funciones.CapturarError(ex, "Seguridad.svc - ListarGruposDeUsuarioEnAD", "wsUNACEMPI", false);
                return new List<string>();
            }
        }

        public List<string> ListarMiembrosDeGrupoEnServidor(string grupo)
        {
            List<string> oMiembros = new List<string>();

            try
            {
                PrincipalContext oPrincipalContext = new PrincipalContext(ContextType.Machine);
                GroupPrincipal oGroupPrincipal = GroupPrincipal.FindByIdentity(oPrincipalContext, grupo);
                PrincipalSearchResult<Principal> oPrincipalSearchResult = oGroupPrincipal.GetMembers();

                foreach (Principal oResult in oPrincipalSearchResult)
                {
                    oMiembros.Add(oResult.SamAccountName.ToUpper());
                }

                return oMiembros;
            }
            catch (Exception ex)
            {
                Funciones.CapturarError(ex, "Seguridad.svc - ListarMiembrosDeGrupoEnServidor", "wsUNACEMPI", false);
                return new List<string>();
            }
        }


        public bool UsuarioPerteneceAlGrupo(string usuario, string grupo)
        {
            try
            {
                List<string> oMiembros = ListarMiembrosDeGrupoEnServidor(grupo);
                int res = oMiembros.IndexOf(usuario.ToUpper());
                if (res != -1)
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            catch (Exception ex)
            {
                Funciones.CapturarError(ex, "Seguridad.svc - UsuarioPerteneceAlGrupo", "wsUNACEMPI", false);
                return false;
            }
        }

    }
}
