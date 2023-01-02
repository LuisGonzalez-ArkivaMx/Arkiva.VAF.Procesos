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

        [DataMember]
        [JsonConfEditor(HelpText = "Configuracion Exportacion a Chronoscan", Label = "Configuracion Exportacion a Chronoscan", IsRequired = true)]
        public ConfiguracionExportacionAChronoscan ConfiguracionExportacionAChronoscan { get; set; }

        [DataMember]
        [JsonConfEditor(Label = "Idioma", TypeEditor = "options", Options = "{selectOptions:[\"es-MX\",\"en-US\"]}", DefaultValue = "es-MX")]
        public string Idioma = "es-MX";
    }

    [DataContract]
    public class ConfiguracionExportacionAChronoscan
    {
        [DataMember]
        [JsonConfEditor(TypeEditor = "options", Options = "{selectOptions:[\"Yes\",\"No\"]}", HelpText = "Habilitar o deshabilitar la aplicacion", Label = "Enabled", DefaultValue = "No")]
        public string Enabled { get; set; } = "No";

        [DataMember]
        [JsonConfEditor(
            Label = "Directorio Persona Moral",
            HelpText = "Ruta donde se exportaran los documentos de Persona Moral para ser procesados por Chronoscan")]
        [Security(ChangeBy = SecurityAttribute.UserLevel.VaultAdmin)]
        public string DirectorioPersonaMoral { get; set; }

        [DataMember]
        [JsonConfEditor(
            Label = "Directorio Persona Fisica",
            HelpText = "Ruta donde se exportaran los documentos de Persona Fisica para ser procesados por Chronoscan")]
        [Security(ChangeBy = SecurityAttribute.UserLevel.VaultAdmin)]
        public string DirectorioPersonaFisica { get; set; }
    }
}
