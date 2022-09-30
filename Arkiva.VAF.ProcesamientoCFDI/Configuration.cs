using MFiles.VAF.Configuration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;

namespace Arkiva.VAF.ProcesamientoCFDI
{
    [DataContract]
    public class Configuration
    {
        [DataMember]
        [JsonConfEditor(
            Label = "RFC Empresa Interna",
            HelpText = "Se debe definir el RFC de Empresa Interna para evaluar si el comprobante es Emitido o Recibido",
            IsRequired = false)]
        [Security(ChangeBy = SecurityAttribute.UserLevel.VaultAdmin)]
        public string sRfcEmpresaInterna { get; set; }
    }
}
