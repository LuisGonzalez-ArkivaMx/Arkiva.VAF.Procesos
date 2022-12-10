using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.Serialization;
using System.Threading.Tasks;
using System.Web;
using System.Xml;
using iTextSharp.text.pdf.codec.wmf;
using MFiles.VAF;
using MFiles.VAF.Common;
using MFiles.VAF.Configuration;
using MFiles.VAF.Core;
using MFilesAPI;
using Microsoft.Office.Interop.Excel;
using SpreadsheetLight;

namespace Arkiva.VAF.ProcesamientoCFDI
{
    /// <summary>
    /// The entry point for this Vault Application Framework application.
    /// </summary>
    /// <remarks>Examples and further information available on the developer portal: http://developer.m-files.com/. </remarks>
    public class VaultApplication
        : ConfigurableVaultApplicationBase<Configuration>
    {
        [EventHandler(MFEventHandlerType.MFEventHandlerBeforeCreateNewObjectFinalize, Class = "CL.CargaMasivaDeCfdi")]
        [EventHandler(MFEventHandlerType.MFEventHandlerAfterFileUpload, Class = "CL.CargaMasivaDeCfdi")]
        public void ProcesarMetadataCFDIYCrearClaseCFDINomina(EventHandlerEnvironment env)
        {
            var filesToDelete = new List<string>();

            var wf_FlujoValidaciones = env.Vault
                .WorkflowOperations
                .GetWorkflowIDByAlias("WF.FlujoProcesamientoDeDocumentos");

            var wfs_EstadoValidado = env.Vault
                .WorkflowOperations
                .GetWorkflowStateIDByAlias("WFS.FlujoDeProcesamientoDeDocumentos.Validado");

            var wfs_EstadoNoValidado = env.Vault
                .WorkflowOperations
                .GetWorkflowStateIDByAlias("WFS.FlujoDeProcesamientoDeDocumentos.NoValidado");

            try
            {
                var oObjVerEx = env.ObjVerEx;
                var oObjectFiles = oObjVerEx.Info.Files;
                IEnumerator enumerator = oObjectFiles.GetEnumerator();

                while (enumerator.MoveNext())
                {
                    ObjectFile oFile = (ObjectFile)enumerator.Current;

                    string sFilePath = SysUtils.GetTempFileName(".tmp");

                    filesToDelete.Add(sFilePath);

                    // Obtener la ultima version del archivo especificado
                    FileVer fileVer = oFile.FileVer;

                    // Descargar el archivo en el directorio temporal
                    env.Vault.ObjectFileOperations.DownloadFile(oFile.ID, fileVer.Version, sFilePath);

                    var sFileName = oFile.GetNameForFileSystem();
                    var sDelimitador = ".";
                    int iIndex = sFileName.LastIndexOf(sDelimitador);
                    var sExtension = sFileName.Substring(iIndex);

                    if (sExtension == ".xml") // Validar que los archivos a procesar sean XML
                    {
                        string sCfdiComprobante = "";

                        // Comienza el proceso de extraccion de metadata del archivo (xml)
                        XmlDocument oDocumento = new XmlDocument();
                        XmlNamespaceManager oManager = new XmlNamespaceManager(oDocumento.NameTable);

                        oDocumento.Load(sFilePath);
                        oManager.AddNamespace("cfdi", "http://www.sat.gob.mx/cfd/3");
                        oManager.AddNamespace("nomina12", "http://www.sat.gob.mx/nomina12");
                        oManager.AddNamespace("tfd", "http://www.sat.gob.mx/TimbreFiscalDigital");
                        XmlElement singleNode = oDocumento.DocumentElement;
                        string sVersionCFDI = "";

                        // Validar la version del CFDI
                        if (singleNode.HasAttribute("Version"))
                        {
                            sVersionCFDI = oDocumento
                                .SelectSingleNode("/cfdi:Comprobante/@Version", oManager)
                                .InnerText;
                        }

                        if (sVersionCFDI == "3.3" || sVersionCFDI == "4.0")
                        {
                            string sTipoDeComprobante = "";

                            if (singleNode.HasAttribute("TipoDeComprobante"))
                            {
                                sTipoDeComprobante = oDocumento
                                    .SelectSingleNode("/cfdi:Comprobante/@TipoDeComprobante", oManager)
                                    .InnerText;
                            }

                            // Validar si el Comprobante es Emitido o Recibido
                            var sRfEmisor = oDocumento
                                .SelectSingleNode("/cfdi:Comprobante/cfdi:Emisor/@Rfc", oManager)
                                .InnerText;

                            var sRfReceptor = oDocumento
                                .SelectSingleNode("/cfdi:Comprobante/cfdi:Receptor/@Rfc", oManager)
                                .InnerText;

                            // Si Rfc Empresa Interna es igual a Rfc Emisor el comprobante es emitido
                            if (sRfEmisor == Configuration.sRfcEmpresaInterna)
                                sCfdiComprobante = "Emitido";

                            // Si Rfc Empresa Interna es igual a Rfc Receptor el comprobante es recibido
                            if (sRfReceptor == Configuration.sRfcEmpresaInterna)
                                sCfdiComprobante = "Recibido";

                            switch (sTipoDeComprobante)
                            {
                                case "N":

                                    if (CreateCFDINomina(sFilePath, sFileName) == true)
                                    {
                                        // Cambiar estado de Workflow asignado a la clase Carga Masiva de CFDI
                                        var oWorkflowstate = new ObjectVersionWorkflowState();
                                        var oObjVer = env.Vault.ObjectOperations.GetLatestObjVerEx(env.ObjVerEx.ObjID, true);

                                        oWorkflowstate.Workflow.TypedValue.SetValue(MFDataType.MFDatatypeLookup, wf_FlujoValidaciones);
                                        oWorkflowstate.State.TypedValue.SetValue(MFDataType.MFDatatypeLookup, wfs_EstadoValidado);
                                        env.Vault.ObjectPropertyOperations.SetWorkflowStateEx(oObjVer, oWorkflowstate);

                                        SysUtils.ReportInfoToEventLog("Se genero exitosamente CFDI Nomina. Fin del proceso");
                                    }
                                    else
                                    {
                                        // Mover el estado del workflow a No Valido cuando el proceso termina incorrectamente
                                        var oWorkflowstate = new ObjectVersionWorkflowState();
                                        var oObjVer = env.Vault.ObjectOperations.GetLatestObjVerEx(env.ObjVerEx.ObjID, true);

                                        oWorkflowstate.Workflow.TypedValue.SetValue(MFDataType.MFDatatypeLookup, wf_FlujoValidaciones);
                                        oWorkflowstate.State.TypedValue.SetValue(MFDataType.MFDatatypeLookup, wfs_EstadoNoValidado);
                                        env.Vault.ObjectPropertyOperations.SetWorkflowStateEx(oObjVer, oWorkflowstate);

                                        SysUtils.ReportInfoToEventLog("El proceso no ha terminado correctamente");
                                    }

                                    break;

                                case "E":

                                    if (CreateComprobanteCFDI(sFilePath, sFileName, sCfdiComprobante) == true)
                                    {
                                        // Cambiar estado de Workflow asignado a la clase Carga Masiva de CFDI
                                        var oWorkflowstate = new ObjectVersionWorkflowState();
                                        var oObjVer = env.Vault.ObjectOperations.GetLatestObjVerEx(env.ObjVerEx.ObjID, true);

                                        oWorkflowstate.Workflow.TypedValue.SetValue(MFDataType.MFDatatypeLookup, wf_FlujoValidaciones);
                                        oWorkflowstate.State.TypedValue.SetValue(MFDataType.MFDatatypeLookup, wfs_EstadoValidado);
                                        env.Vault.ObjectPropertyOperations.SetWorkflowStateEx(oObjVer, oWorkflowstate);

                                        SysUtils.ReportInfoToEventLog("Se genero exitosamente Comprobante CFDI. Fin del proceso");
                                    }
                                    else
                                    {
                                        // Mover el estado del workflow a No Valido cuando el proceso termina incorrectamente
                                        var oWorkflowstate = new ObjectVersionWorkflowState();
                                        var oObjVer = env.Vault.ObjectOperations.GetLatestObjVerEx(env.ObjVerEx.ObjID, true);

                                        oWorkflowstate.Workflow.TypedValue.SetValue(MFDataType.MFDatatypeLookup, wf_FlujoValidaciones);
                                        oWorkflowstate.State.TypedValue.SetValue(MFDataType.MFDatatypeLookup, wfs_EstadoNoValidado);
                                        env.Vault.ObjectPropertyOperations.SetWorkflowStateEx(oObjVer, oWorkflowstate);

                                        SysUtils.ReportInfoToEventLog("El proceso no ha terminado correctamente");
                                    }

                                    break;

                                case "I":

                                    if (CreateComprobanteCFDI(sFilePath, sFileName, sCfdiComprobante) == true)
                                    {
                                        // Cambiar estado de Workflow asignado a la clase Carga Masiva de CFDI
                                        var oWorkflowstate = new ObjectVersionWorkflowState();
                                        var oObjVer = env.Vault.ObjectOperations.GetLatestObjVerEx(env.ObjVerEx.ObjID, true);

                                        oWorkflowstate.Workflow.TypedValue.SetValue(MFDataType.MFDatatypeLookup, wf_FlujoValidaciones);
                                        oWorkflowstate.State.TypedValue.SetValue(MFDataType.MFDatatypeLookup, wfs_EstadoValidado);
                                        env.Vault.ObjectPropertyOperations.SetWorkflowStateEx(oObjVer, oWorkflowstate);

                                        SysUtils.ReportInfoToEventLog("Se genero exitosamente Comprobante CFDI. Fin del proceso");
                                    }
                                    else
                                    {
                                        // Mover el estado del workflow a No Valido cuando el proceso termina incorrectamente
                                        var oWorkflowstate = new ObjectVersionWorkflowState();
                                        var oObjVer = env.Vault.ObjectOperations.GetLatestObjVerEx(env.ObjVerEx.ObjID, true);

                                        oWorkflowstate.Workflow.TypedValue.SetValue(MFDataType.MFDatatypeLookup, wf_FlujoValidaciones);
                                        oWorkflowstate.State.TypedValue.SetValue(MFDataType.MFDatatypeLookup, wfs_EstadoNoValidado);
                                        env.Vault.ObjectPropertyOperations.SetWorkflowStateEx(oObjVer, oWorkflowstate);

                                        SysUtils.ReportInfoToEventLog("El proceso no ha terminado correctamente");
                                    }

                                    break;

                                case "P":

                                    if (CreateComprobanteCFDI(sFilePath, sFileName, sCfdiComprobante) == true)
                                    {
                                        // Cambiar estado de Workflow asignado a la clase Carga Masiva de CFDI
                                        var oWorkflowstate = new ObjectVersionWorkflowState();
                                        var oObjVer = env.Vault.ObjectOperations.GetLatestObjVerEx(env.ObjVerEx.ObjID, true);

                                        oWorkflowstate.Workflow.TypedValue.SetValue(MFDataType.MFDatatypeLookup, wf_FlujoValidaciones);
                                        oWorkflowstate.State.TypedValue.SetValue(MFDataType.MFDatatypeLookup, wfs_EstadoValidado);
                                        env.Vault.ObjectPropertyOperations.SetWorkflowStateEx(oObjVer, oWorkflowstate);

                                        SysUtils.ReportInfoToEventLog("Se genero exitosamente Comprobante CFDI. Fin del proceso");
                                    }
                                    else
                                    {
                                        // Mover el estado del workflow a No Valido cuando el proceso termina incorrectamente
                                        var oWorkflowstate = new ObjectVersionWorkflowState();
                                        var oObjVer = env.Vault.ObjectOperations.GetLatestObjVerEx(env.ObjVerEx.ObjID, true);

                                        oWorkflowstate.Workflow.TypedValue.SetValue(MFDataType.MFDatatypeLookup, wf_FlujoValidaciones);
                                        oWorkflowstate.State.TypedValue.SetValue(MFDataType.MFDatatypeLookup, wfs_EstadoNoValidado);
                                        env.Vault.ObjectPropertyOperations.SetWorkflowStateEx(oObjVer, oWorkflowstate);

                                        SysUtils.ReportInfoToEventLog("El proceso no ha terminado correctamente");
                                    }

                                    break;

                                case "T":

                                    if (CreateComprobanteCFDI(sFilePath, sFileName, sCfdiComprobante) == true)
                                    {
                                        // Cambiar estado de Workflow asignado a la clase Carga Masiva de CFDI
                                        var oWorkflowstate = new ObjectVersionWorkflowState();
                                        var oObjVer = env.Vault.ObjectOperations.GetLatestObjVerEx(env.ObjVerEx.ObjID, true);

                                        oWorkflowstate.Workflow.TypedValue.SetValue(MFDataType.MFDatatypeLookup, wf_FlujoValidaciones);
                                        oWorkflowstate.State.TypedValue.SetValue(MFDataType.MFDatatypeLookup, wfs_EstadoValidado);
                                        env.Vault.ObjectPropertyOperations.SetWorkflowStateEx(oObjVer, oWorkflowstate);

                                        SysUtils.ReportInfoToEventLog("Se genero exitosamente Comprobante CFDI. Fin del proceso");
                                    }
                                    else
                                    {
                                        // Mover el estado del workflow a No Valido cuando el proceso termina incorrectamente
                                        var oWorkflowstate = new ObjectVersionWorkflowState();
                                        var oObjVer = env.Vault.ObjectOperations.GetLatestObjVerEx(env.ObjVerEx.ObjID, true);

                                        oWorkflowstate.Workflow.TypedValue.SetValue(MFDataType.MFDatatypeLookup, wf_FlujoValidaciones);
                                        oWorkflowstate.State.TypedValue.SetValue(MFDataType.MFDatatypeLookup, wfs_EstadoNoValidado);
                                        env.Vault.ObjectPropertyOperations.SetWorkflowStateEx(oObjVer, oWorkflowstate);

                                        SysUtils.ReportInfoToEventLog("El proceso no ha terminado correctamente");
                                    }

                                    break;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                SysUtils.ReportErrorMessageToEventLog("Error en ProcesarMetadataCFDIYCrearClaseCFDINomina...", ex);
            }
            finally
            {
                foreach (var sFile in filesToDelete)
                {
                    File.Delete(sFile);
                }
            }
        }

        [EventHandler(MFEventHandlerType.MFEventHandlerBeforeCreateNewObjectFinalize, Class = "Arkiva.Class.CFDIComprobanteEmitido")]
        [EventHandler(MFEventHandlerType.MFEventHandlerBeforeCreateNewObjectFinalize, Class = "Arkiva.Class.CFDIComprobanteRecibido")]
        public void ProcesosComprobantesCFDI(EventHandlerEnvironment env)
        {
            int iOrigenCFDICompulsa = 0;
            bool bIssue = false;
            string sRfcEmisorValue = "";
            string sRfcReceptorValue = "";
            string sNombreValue = "";

            var ot_Proveedor = env.Vault.ObjectTypeOperations.GetObjectTypeIDByAlias("OT.Proveedor");
            var ot_EmpresaInterna = env.Vault.ObjectTypeOperations.GetObjectTypeIDByAlias("OT.EmpresaInterna");
            var ot_Cliente = env.Vault.ObjectTypeOperations.GetObjectTypeIDByAlias("OT.Cliente");
            var cl_Proveedor = env.Vault.ClassOperations.GetObjectClassIDByAlias("CL.Proveedor");
            var cl_EmpresaInterna = env.Vault.ClassOperations.GetObjectClassIDByAlias("CL.EmpresaInterna");
            var cl_Cliente = env.Vault.ClassOperations.GetObjectClassIDByAlias("CL.Cliente");
            var cl_CFDIComprobanteRecibido = env.Vault.ClassOperations.GetObjectClassIDByAlias("Arkiva.Class.CFDIComprobanteRecibido");
            var cl_CFDIComprobanteEmitido = env.Vault.ClassOperations.GetObjectClassIDByAlias("Arkiva.Class.CFDIComprobanteEmitido");
            var cl_ConceptoCFDI = env.Vault.ClassOperations.GetObjectClassIDByAlias("Arkiva.Class.ConceptoDeCFDI");
            var cl_ComplementoPagoEmitido = env.Vault.ClassOperations.GetObjectClassIDByAlias("Arkiva.Class.CFDIComplementoPagoEmitido");
            var cl_ComplementoPagoRecibido = env.Vault.ClassOperations.GetObjectClassIDByAlias("Arkiva.Class.CFDIComplementoPagoRecibido");
            var cl_OrdenCompraEmitidaProveedor = env.Vault.ClassOperations.GetObjectClassIDByAlias("CL.OrdenDeCompraEmitidaProveedor");
            var cl_OrdenCompraRecibidaCliente = env.Vault.ClassOperations.GetObjectClassIDByAlias("CL.OrdenDeCompraRecibidaCliente");
            var cl_EntregableRecibidoProveedor = env.Vault.ClassOperations.GetObjectClassIDByAlias("CL.EntregableRecibidoProveedor");
            var cl_EntregableEmitidoCliente = env.Vault.ClassOperations.GetObjectClassIDByAlias("CL.EntregableEmitidoCliente");
            var cl_ProyectoServicioEspecializado = env.Vault.ClassOperations.GetObjectClassIDByAlias("CL.ProyectoServicioEspecializado");
            var cl_ProyectoCliente = env.Vault.ClassOperations.GetObjectClassIDByAlias("MF.CL.Project");
            var cl_Contrato = env.Vault.ClassOperations.GetObjectClassIDByAlias("CL.Contrato");
            var pd_RfcEmpresa = env.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.RfcEmpresa");
            var pd_Proveedor = env.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("MF.PD.Proveedor");
            var pd_EmpresaInterna = env.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.EmpresaInterna");
            var pd_Cliente = env.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.Cliente");
            var pd_ConceptoCFDI = env.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("CFDI.ConceptodeCFDI.Objeto");
            var pd_ComplementoPagoEmitido = env.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("CFDI.CFDIComplementoEmitido.Texto");
            var pd_ComplementoPagoRecibido = env.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("CFDI.CFDIComplementoRecibido.Texto");
            var pd_Version = env.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("CFDI.VersionCFDI.Texto");
            var pd_Total = env.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("CFDI.Total.Texto");
            var pd_TipoComprobante = env.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("CFDI.TipodeComprobante.TextoCFDInomina");
            var pd_TipoCambio = env.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("CFDI.TipodeCambio.Texto");
            var pd_Subtotal = env.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("CFDI.SubTotal.Texto");
            var pd_Serie = env.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("CFDI.Serie.TextoCFDInomina");
            var pd_Sello = env.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("CFDI.Sello.Texto");
            var pd_NoCertificado = env.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("CFDI.NumeroCertificado.TextoCFDInomina");
            var pd_Moneda = env.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.MonedaTextCFDInomina");
            var pd_MetodoPago = env.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("CFDI.MétododePago.TextoCFDInomina");
            var pd_LugarExpedicion = env.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("CFDI.LugardeExpedición.Texto");
            var pd_FormaPago = env.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("CFDI.FormadePago.Texto");
            var pd_Folio = env.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("CFDI.FoliodeCFDI.TextoCFDInomina");
            var pd_Fecha = env.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("CFDI.FechadeCFDI.Texto");
            var pd_CondicionesPago = env.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("CFDI.CondiciondePago.Texto");
            var pd_Certificado = env.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("CFDI.CertificadoCFDI.TextoCFDInomina");
            var pd_RfcEmisor = env.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("CFDI.RFCEmisor.Texto");
            var pd_NombreEmisor = env.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("CFDI.EmisorCFDI.Texto");
            var pd_RegimenFiscal = env.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("CFDI.RegimenFiscal.Texto");
            var pd_RfcReceptor = env.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("CFDI.RFCReceptor.Texto");
            var pd_NombreReceptor = env.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("CFDI.ReceptorCFDI.Texto");
            var pd_UsoCfdi = env.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("CFDI.UsodeCFDI.Texto");
            var pd_TotalImpuestosTrasladados = env.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("CFDI.TotalImpuestosTrasladados.Texto");
            var pd_Importe = env.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("CFDI.Importe.Texto");
            var pd_TipoFactor = env.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("CFDI.TipodeFactor.Texto");
            var pd_TasaOCuota = env.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("CFDI.TasaoCuota.Texto");
            var pd_Impuesto = env.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("CFDI.Impuestos.Texto");
            var pd_Uuid = env.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("CFDI.UUID.Texto");
            var pd_OrigenCFDI = env.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.OrigenDelCfdi");
            var pd_OrdenesCompraEmitidas = env.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.OrdenesDeCompraEmitidas");
            var pd_OrdenesCompraRecibidas = env.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.OrdenesDeCompraRecibidas");
            var pd_EntregablesRecibidos = env.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.EntregablesRecibidos");
            var pd_EntregablesEmitidos = env.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.EntregablesEmitidos");
            var pd_ProyectosRelacionados = env.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("MF.PD.Project");
            var pd_ContratosRelacionados = env.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("MF.PD.RelatedContract");
            var pd_Class = env.Vault.PropertyDefOperations.GetBuiltInPropertyDef(MFBuiltInPropertyDef.MFBuiltInPropertyDefClass);

            var oPropertyValues = new PropertyValues();
            oPropertyValues = env.Vault.ObjectPropertyOperations.GetProperties(env.ObjVer);

            var iClase = oPropertyValues.SearchForPropertyEx(pd_Class.ID, true).TypedValue.GetLookupID();
            var sUuid = oPropertyValues.SearchForPropertyEx(pd_Uuid, true).TypedValue.Value.ToString();

            // Iniciar proceso para asociar el proveedor o empresa interna al CFDI comprobante 
            if (iClase == cl_CFDIComprobanteRecibido) // Si es comprobante recibido
            {
                // Validaciones RFC Emisor 
                sRfcEmisorValue = oPropertyValues.SearchForPropertyEx(pd_RfcEmisor, true).TypedValue.Value.ToString();
                sNombreValue = oPropertyValues.SearchForPropertyEx(pd_NombreEmisor, true).TypedValue.Value.ToString();

                if (sRfcEmisorValue != "" && sNombreValue != "")
                {
                    // Si existe Rfc/Nombre, se busca Rfc en Organizacion, si no existe se crea y vincula al comprobante
                    if (GetExistingRfc(env, ot_Proveedor, pd_RfcEmpresa, sRfcEmisorValue) == false)
                    {
                        CreateRfcObject(env, ot_Proveedor, cl_Proveedor, pd_RfcEmpresa, sRfcEmisorValue, sNombreValue);
                    }

                    // Vincula Emisor de CFDI Recibido
                    SetBindingProperties(env, cl_Proveedor, pd_RfcEmpresa, pd_Proveedor, sRfcEmisorValue, 2);
                }
                else if (sRfcEmisorValue != "" && sNombreValue == "")
                {
                    // Si no hay nombre emisor en CFDI, Solo se vincula Emisor de CFDI Recibido
                    SetBindingProperties(env, cl_Proveedor, pd_RfcEmpresa, pd_Proveedor, sRfcEmisorValue, 2);
                }

                // Validaciones RFC Receptor
                sRfcReceptorValue = oPropertyValues.SearchForPropertyEx(pd_RfcReceptor, true).TypedValue.Value.ToString();
                sNombreValue = oPropertyValues.SearchForPropertyEx(pd_NombreReceptor, true).TypedValue.Value.ToString();

                if (sRfcReceptorValue != "" && sNombreValue != "")
                {
                    if (GetExistingRfc(env, ot_EmpresaInterna, pd_RfcEmpresa, sRfcReceptorValue) == false)
                    {
                        CreateRfcObject(env, ot_EmpresaInterna, cl_EmpresaInterna, pd_RfcEmpresa, sRfcReceptorValue, sNombreValue);
                    }

                    // Vincula Receptor de CFDI Recibido
                    SetBindingProperties(env, cl_EmpresaInterna, pd_RfcEmpresa, pd_EmpresaInterna, sRfcReceptorValue, 2);
                }
                else if (sRfcReceptorValue != "" && sNombreValue == "")
                {
                    SetBindingProperties(env, cl_EmpresaInterna, pd_RfcEmpresa, pd_EmpresaInterna, sRfcReceptorValue, 2);
                }

                // Vincula los conceptos de CFDI y complementos de pago con el comprobante recibido
                if (sUuid != "")
                {
                    SetBindingProperties(env, cl_ConceptoCFDI, pd_Uuid, pd_ConceptoCFDI, sUuid, 2);
                    SetBindingProperties(env, cl_ComplementoPagoEmitido, pd_Uuid, pd_ComplementoPagoEmitido, sUuid, 2);
                }

                // Obtener el objeto proveedor
                var searchBuilderProveedor = new MFSearchBuilder(env.Vault);
                searchBuilderProveedor.Deleted(false); // No eliminados
                searchBuilderProveedor.ObjType(ot_Proveedor);
                searchBuilderProveedor.Property(pd_RfcEmpresa, MFDataType.MFDatatypeText, sRfcEmisorValue);

                var searchResultsProveedor = searchBuilderProveedor.FindEx();

                if (searchResultsProveedor.Count > 0)
                {
                    var oLookupProveedor = new Lookup
                    {
                        Item = searchResultsProveedor[0].ObjVer.ID
                    };

                    var oLookupsProveedor = new Lookups
                    {
                        { -1, oLookupProveedor }
                    };

                    // Relacionar a ordenes de compra emitidas (proveedor)
                    SetPropertiesCFDIComprobante(env, cl_CFDIComprobanteRecibido, cl_OrdenCompraEmitidaProveedor, pd_Proveedor, pd_OrdenesCompraEmitidas, oLookupsProveedor);

                    // Relacionar a entregable recibidas
                    SetPropertiesCFDIComprobante(env, cl_CFDIComprobanteRecibido, cl_EntregableRecibidoProveedor, pd_Proveedor, pd_EntregablesRecibidos, oLookupsProveedor);

                    // Relacionar a proyecto relacionados
                    SetPropertiesCFDIComprobante(env, cl_CFDIComprobanteRecibido, cl_ProyectoServicioEspecializado, pd_Proveedor, pd_ProyectosRelacionados, oLookupsProveedor);

                    // Relacionar a contrato relacionados
                    SetPropertiesCFDIComprobante(env, cl_CFDIComprobanteRecibido, cl_Contrato, pd_Proveedor, pd_ContratosRelacionados, oLookupsProveedor);
                }
            }

            if (iClase == cl_CFDIComprobanteEmitido) // Si es comprobante emitido
            {
                // Validaciones de RFC Receptor
                sRfcReceptorValue = oPropertyValues.SearchForPropertyEx(pd_RfcReceptor, true).TypedValue.Value.ToString();
                sNombreValue = oPropertyValues.SearchForPropertyEx(pd_NombreReceptor, true).TypedValue.Value.ToString();

                if (sRfcReceptorValue != "" && sNombreValue != "")
                {
                    if (GetExistingRfc(env, ot_Cliente, pd_RfcEmpresa, sRfcReceptorValue) == false)
                    {
                        CreateRfcObject(env, ot_Cliente, cl_Cliente, pd_RfcEmpresa, sRfcReceptorValue, sNombreValue);
                    }

                    SetBindingProperties(env, cl_Cliente, pd_RfcEmpresa, pd_Cliente, sRfcReceptorValue, 2);
                }
                else if (sRfcReceptorValue != "" && sNombreValue == "")
                {
                    SetBindingProperties(env, cl_Cliente, pd_RfcEmpresa, pd_Cliente, sRfcReceptorValue, 2);
                }

                // Validaciones de RFC Emisor
                sRfcEmisorValue = oPropertyValues.SearchForPropertyEx(pd_RfcEmisor, true).TypedValue.Value.ToString();
                sNombreValue = oPropertyValues.SearchForPropertyEx(pd_NombreEmisor, true).TypedValue.Value.ToString();

                if (sRfcEmisorValue != "" && sNombreValue != "")
                {
                    if (GetExistingRfc(env, ot_EmpresaInterna, pd_RfcEmpresa, sRfcEmisorValue) == false)
                    {
                        CreateRfcObject(env, ot_EmpresaInterna, cl_EmpresaInterna, pd_RfcEmpresa, sRfcEmisorValue, sNombreValue);
                    }

                    SetBindingProperties(env, cl_EmpresaInterna, pd_RfcEmpresa, pd_EmpresaInterna, sRfcEmisorValue, 2);
                }
                else if (sRfcEmisorValue != "" && sNombreValue == "")
                {
                    SetBindingProperties(env, cl_EmpresaInterna, pd_RfcEmpresa, pd_EmpresaInterna, sRfcEmisorValue, 2);
                }

                // Vincula los conceptos de CFDI y complementos de pago con el comprobante emitido
                if (sUuid != "")
                {
                    SetBindingProperties(env, cl_ConceptoCFDI, pd_Uuid, pd_ConceptoCFDI, sUuid, 2);
                    SetBindingProperties(env, cl_ComplementoPagoRecibido, pd_Uuid, pd_ComplementoPagoRecibido, sUuid, 2);
                }

                // Obtener el objeto Empresa Interna
                var searchBuilderEmpresaInterna = new MFSearchBuilder(env.Vault);
                searchBuilderEmpresaInterna.Deleted(false); // No eliminados
                searchBuilderEmpresaInterna.Class(cl_EmpresaInterna);
                searchBuilderEmpresaInterna.Property(pd_RfcEmpresa, MFDataType.MFDatatypeText, sRfcEmisorValue);

                var searchResultsEmpresaInterna = searchBuilderEmpresaInterna.FindEx();

                if (searchResultsEmpresaInterna.Count > 0)
                {
                    var oLookupEmpresaInterna = new Lookup
                    {
                        Item = searchResultsEmpresaInterna[0].ObjVer.ID
                    };

                    var oLookupsEmpresaInterna = new Lookups
                    {
                        { -1, oLookupEmpresaInterna }
                    };

                    // Relacionar a Orden de Compra - Recibida (Cliente)
                    SetPropertiesCFDIComprobante(env, cl_CFDIComprobanteEmitido, cl_OrdenCompraRecibidaCliente, pd_EmpresaInterna, pd_OrdenesCompraRecibidas, oLookupsEmpresaInterna);

                    // Relacionar a Entregable Emitido (Cliente)
                    SetPropertiesCFDIComprobante(env, cl_CFDIComprobanteEmitido, cl_EntregableEmitidoCliente, pd_EmpresaInterna, pd_EntregablesEmitidos, oLookupsEmpresaInterna);

                    // Relacionar a proyectos relacionados
                    SetPropertiesCFDIComprobante(env, cl_CFDIComprobanteEmitido, cl_ProyectoCliente, pd_EmpresaInterna, pd_ProyectosRelacionados, oLookupsEmpresaInterna);

                    // Relacionar a contratos relacionados
                    SetPropertiesCFDIComprobante(env, cl_CFDIComprobanteEmitido, cl_Contrato, pd_EmpresaInterna, pd_ContratosRelacionados, oLookupsEmpresaInterna);
                }
            }

            // Inicia proceso de comparacion de CFDI SAT vs CFDI Externo (subido por proveedor)                        
            var iOrigenCFDI = oPropertyValues.SearchForPropertyEx(pd_OrigenCFDI, true).TypedValue.GetLookupID();

            // Agregar a una lista los comprobantes CFDI a comparar
            var oListCFDIComparados = new List<ObjVerEx>
            {
                env.ObjVerEx
            };

            if (iOrigenCFDI == 2)
                iOrigenCFDICompulsa = 1; // Si CFDI es Proveedor, el CFDI a compulsar es SAT
            else
                iOrigenCFDICompulsa = 2; // Si CFDI es SAT, la compulsa es con CFDI Proveedor

            // Obtener el objeto CFDI a compulsar
            var oListCFDICompulsa = GetListCFDICompulsa(env, iClase, pd_OrigenCFDI, iOrigenCFDICompulsa, pd_Uuid, sUuid);

            if (oListCFDICompulsa.Count > 0)
            {
                var sDescripcion = "Propiedades en los que se detectaron inconsistencias: " + Environment.NewLine;
                var oCFDICompulsa = oListCFDICompulsa[0];

                var oPropertyValuesComp = new PropertyValues();
                oPropertyValuesComp = oCFDICompulsa.Vault.ObjectPropertyOperations.GetProperties(oCFDICompulsa.ObjVer);

                // Obtener los campos a validar de los CFDI a compulsar
                // Certificado CFDI
                var sCertificado = oPropertyValues.SearchForPropertyEx(pd_Certificado, true).TypedValue.Value.ToString();
                var sCertificadoComp = oPropertyValuesComp.SearchForPropertyEx(pd_Certificado, true).TypedValue.Value.ToString();

                if (sCertificado != sCertificadoComp)
                {
                    bIssue = true;
                    sDescripcion += "- Certificado CFDI" + Environment.NewLine;
                }

                // Folio de CFDI
                var sFolio = oPropertyValues.SearchForPropertyEx(pd_Folio, true).TypedValue.Value.ToString();
                var sFolioComp = oPropertyValuesComp.SearchForPropertyEx(pd_Folio, true).TypedValue.Value.ToString();

                if (sFolio != sFolioComp)
                {
                    bIssue = true;
                    sDescripcion += "- Folio de CFDI" + Environment.NewLine;
                }

                // Fecha de CFDI
                var sFecha = oPropertyValues.SearchForPropertyEx(pd_Fecha, true).TypedValue.Value.ToString();
                var sFechaComp = oPropertyValuesComp.SearchForPropertyEx(pd_Fecha, true).TypedValue.Value.ToString();

                if (sFecha != sFechaComp)
                {
                    bIssue = true;
                    sDescripcion += "- Fecha de CFDI" + Environment.NewLine;
                }

                // Importe
                var sImporte = oPropertyValues.SearchForPropertyEx(pd_Importe, true).TypedValue.Value.ToString();
                var sImporteComp = oPropertyValuesComp.SearchForPropertyEx(pd_Importe, true).TypedValue.Value.ToString();

                if (sImporte != sImporteComp)
                {
                    bIssue = true;
                    sDescripcion += "- Importe" + Environment.NewLine;
                }

                // Numero Certificado
                var sNumCertificado = oPropertyValues.SearchForPropertyEx(pd_NoCertificado, true).TypedValue.Value.ToString();
                var sNumCertificadoComp = oPropertyValuesComp.SearchForPropertyEx(pd_NoCertificado, true).TypedValue.Value.ToString();

                if (sNumCertificado != sNumCertificadoComp)
                {
                    bIssue = true;
                    sDescripcion += "- Numero Certificado" + Environment.NewLine;
                }

                // Regimen Fiscal
                var sRegimenFiscal = oPropertyValues.SearchForPropertyEx(pd_RegimenFiscal, true).TypedValue.Value.ToString();
                var sRegimenFiscalComp = oPropertyValuesComp.SearchForPropertyEx(pd_RegimenFiscal, true).TypedValue.Value.ToString();

                if (sRegimenFiscal != sRegimenFiscalComp)
                {
                    bIssue = true;
                    sDescripcion += "- Regimen Fiscal" + Environment.NewLine;
                }

                // Sello
                var sSello = oPropertyValues.SearchForPropertyEx(pd_Sello, true).TypedValue.Value.ToString();
                var sSelloComp = oPropertyValuesComp.SearchForPropertyEx(pd_Sello, true).TypedValue.Value.ToString();

                if (sSello != sSelloComp)
                {
                    bIssue = true;
                    sDescripcion += "- Sello" + Environment.NewLine;
                }

                // Serie
                var sSerie = oPropertyValues.SearchForPropertyEx(pd_Serie, true).TypedValue.Value.ToString();
                var sSerieComp = oPropertyValuesComp.SearchForPropertyEx(pd_Serie, true).TypedValue.Value.ToString();

                if (sSerie != sSerieComp)
                {
                    bIssue = true;
                    sDescripcion += "- Serie" + Environment.NewLine;
                }

                // SubTotal
                var sSubtotal = oPropertyValues.SearchForPropertyEx(pd_Subtotal, true).TypedValue.Value.ToString();
                var sSubtotalComp = oPropertyValuesComp.SearchForPropertyEx(pd_Subtotal, true).TypedValue.Value.ToString();

                if (sSubtotal != sSubtotalComp)
                {
                    bIssue = true;
                    sDescripcion += "- SubTotal" + Environment.NewLine;
                }

                // Tasa o Cuota
                var sTasaOCuota = oPropertyValues.SearchForPropertyEx(pd_TasaOCuota, true).TypedValue.Value.ToString();
                var sTasaOCuotaComp = oPropertyValuesComp.SearchForPropertyEx(pd_TasaOCuota, true).TypedValue.Value.ToString();

                if (sTasaOCuota != sTasaOCuotaComp)
                {
                    bIssue = true;
                    sDescripcion += "- Tasa o Cuota" + Environment.NewLine;
                }

                // Tipo de Cambio
                var sTipoCambio = oPropertyValues.SearchForPropertyEx(pd_TipoCambio, true).TypedValue.Value.ToString();
                var sTipoCambioComp = oPropertyValuesComp.SearchForPropertyEx(pd_TipoCambio, true).TypedValue.Value.ToString();

                if (sTipoCambio != sTipoCambioComp)
                {
                    bIssue = true;
                    sDescripcion += "- Tipo de Cambio" + Environment.NewLine;
                }

                // Total
                var sTotal = oPropertyValues.SearchForPropertyEx(pd_Total, true).TypedValue.Value.ToString();
                var sTotalComp = oPropertyValuesComp.SearchForPropertyEx(pd_Total, true).TypedValue.Value.ToString();

                if (sTotal != sTotalComp)
                {
                    bIssue = true;
                    sDescripcion += "- Total" + Environment.NewLine;
                }

                // Total Impuestos Trasladados
                var sTotalImpuestosTrasladados = oPropertyValues.SearchForPropertyEx(pd_TotalImpuestosTrasladados, true).TypedValue.Value.ToString();
                var sTotalImpuestosTrasladadosComp = oPropertyValuesComp.SearchForPropertyEx(pd_TotalImpuestosTrasladados, true).TypedValue.Value.ToString();

                if (sTotalImpuestosTrasladados != sTotalImpuestosTrasladadosComp)
                {
                    bIssue = true;
                    sDescripcion += "- Total Impuestos Trasladados" + Environment.NewLine;
                }

                // Uso de CFDI
                var sUsoCFDI = oPropertyValues.SearchForPropertyEx(pd_UsoCfdi, true).TypedValue.Value.ToString();
                var sUsoCFDIComp = oPropertyValuesComp.SearchForPropertyEx(pd_UsoCfdi, true).TypedValue.Value.ToString();

                if (sUsoCFDI != sUsoCFDIComp)
                {
                    bIssue = true;
                    sDescripcion += "- Uso de CFDI" + Environment.NewLine;
                }

                // Forma de Pago
                var sFormaPago = oPropertyValues.SearchForPropertyEx(pd_FormaPago, true).TypedValue.Value.ToString();
                var sFormaPagoComp = oPropertyValuesComp.SearchForPropertyEx(pd_FormaPago, true).TypedValue.Value.ToString();

                if (sFormaPago != sFormaPagoComp)
                {
                    bIssue = true;
                    sDescripcion += "- Forma de Pago" + Environment.NewLine;
                }

                oListCFDIComparados.Add(oCFDICompulsa);

                // Generar issue si hubo inconcistencias en las comparaciones de la metadata de los CFDI
                if (bIssue == true)
                {
                    var oLookupsDocumentos = new Lookups();
                    var oLookupsProveedorIssue = new Lookups();
                    var oLookupsEmpresaInternaIssue = new Lookups();
                    var oLookupDocumento1 = new Lookup();
                    var oLookupDocumento2 = new Lookup();

                    oLookupDocumento1.Item = env.ObjVer.ID;
                    oLookupDocumento2.Item = oCFDICompulsa.ID;

                    oLookupsDocumentos.Add(-1, oLookupDocumento1);
                    oLookupsDocumentos.Add(-1, oLookupDocumento2);

                    // Generar el numero consecutivo del siguiente issue a crear
                    var issues = GetExistingIssues(env);
                    var issuesCount = issues.Count;
                    var noConsecutivo = issuesCount + 1;
                    var nombreOTitulo = "Issue #" + noConsecutivo;

                    // Obtener Proveedor y Empresa Interna de los documentos comparados
                    var sRfcProveedor = oPropertyValues.SearchForPropertyEx(pd_RfcEmisor, true).TypedValue.Value.ToString();
                    var sRfcEmpresaInterna = oPropertyValues.SearchForPropertyEx(pd_RfcReceptor, true).TypedValue.Value.ToString();

                    var searchBuilderProveedor = new MFSearchBuilder(env.Vault);
                    searchBuilderProveedor.Deleted(false); // No eliminados
                    searchBuilderProveedor.ObjType(ot_Proveedor);
                    searchBuilderProveedor.Property(pd_RfcEmpresa, MFDataType.MFDatatypeText, sRfcProveedor);

                    var searchResultsProveedor = searchBuilderProveedor.FindEx();

                    if (searchResultsProveedor.Count > 0)
                    {
                        var oLookup = new Lookup
                        {
                            Item = searchResultsProveedor[0].ObjVer.ID
                        };

                        oLookupsProveedorIssue.Add(-1, oLookup);
                    }

                    var searchBuilderEmpresaInterna = new MFSearchBuilder(env.Vault);
                    searchBuilderEmpresaInterna.Deleted(false); // No eliminados
                    searchBuilderEmpresaInterna.Class(cl_EmpresaInterna);
                    searchBuilderEmpresaInterna.Property(pd_RfcEmpresa, MFDataType.MFDatatypeText, sRfcEmpresaInterna);

                    var searchResultsEmpresaInterna = searchBuilderEmpresaInterna.FindEx();

                    if (searchResultsEmpresaInterna.Count > 0)
                    {
                        var oLookup = new Lookup
                        {
                            Item = searchResultsEmpresaInterna[0].ObjVer.ID
                        };

                        oLookupsEmpresaInternaIssue.Add(-1, oLookup);
                    }
                    
                    // Crear el issue
                    CreateIssue(env, oLookupsProveedorIssue, oLookupsEmpresaInternaIssue, oLookupsDocumentos, nombreOTitulo, sDescripcion);

                    // Compulsa con inconcistencias
                    UpdateEstatusCFDI(env.ObjVerEx, 2);
                    UpdateEstatusCFDIComp(oCFDICompulsa, 2);
                }
                else
                {
                    // Compulsa exitosa
                    UpdateEstatusCFDI(env.ObjVerEx, 3);
                    UpdateEstatusCFDIComp(oCFDICompulsa, 3);
                }
            }
            else
            {
                // Sin documento "par" para la compulsa del CFDI
                UpdateEstatusCFDI(env.ObjVerEx, 1);
            }
        }

        [EventHandler(MFEventHandlerType.MFEventHandlerBeforeCheckInChanges, Class = "CL.OrdenDeCompraEmitidaProveedor")]
        [EventHandler(MFEventHandlerType.MFEventHandlerBeforeCheckInChanges, Class = "CL.OrdenDeCompraRecibidaCliente")]
        [EventHandler(MFEventHandlerType.MFEventHandlerBeforeCheckInChanges, Class = "CL.FacturaRecibidaProveedor")]
        [EventHandler(MFEventHandlerType.MFEventHandlerBeforeCheckInChanges, Class = "CL.Contrato")]
        public void ProcesosOrdenesDeCompra(EventHandlerEnvironment env)
        {
            var cl_OrdenCompraEmitidaProveedor = env.Vault.ClassOperations.GetObjectClassIDByAlias("CL.OrdenDeCompraEmitidaProveedor");
            var cl_OrdenCompraRecibidaCliente = env.Vault.ClassOperations.GetObjectClassIDByAlias("CL.OrdenDeCompraRecibidaCliente");
            var cl_FacturaRecibidaProveedor = env.Vault.ClassOperations.GetObjectClassIDByAlias("CL.FacturaRecibidaProveedor");
            var cl_EntregableRecibidoProveedor = env.Vault.ClassOperations.GetObjectClassIDByAlias("CL.EntregableRecibidoProveedor");
            var cl_EntregableEmitidoCliente = env.Vault.ClassOperations.GetObjectClassIDByAlias("CL.EntregableEmitidoCliente");
            var cl_ProyectoServicioEspecializado = env.Vault.ClassOperations.GetObjectClassIDByAlias("CL.ProyectoServicioEspecializado");
            var cl_ProyectoCliente = env.Vault.ClassOperations.GetObjectClassIDByAlias("MF.CL.Project");
            var cl_Contrato = env.Vault.ClassOperations.GetObjectClassIDByAlias("CL.Contrato");
            var pd_Proveedor = env.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("MF.PD.Proveedor");
            var pd_EmpresaInterna = env.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.EmpresaInterna");
            var pd_EntregablesRecibidos = env.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.EntregablesRecibidos");
            var pd_EntregablesEmitidos = env.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.EntregablesEmitidos");
            var pd_ProyectosRelacionados = env.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("MF.PD.Project");
            var pd_ContratosRelacionados = env.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("MF.PD.RelatedContract");
            var pd_ContratosRelacionadosSE = env.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.ContratosRelacionados");
            var pd_FacturasRelacionadasRecibidas = env.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.FacturasRelacionadasRecibidas");
            var pd_OrdenesCompraEmitidas = env.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.OrdenesDeCompraEmitidas");
            var pd_EsConvenioModificatorio = env.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.EsConvenioModificatorio");
            var pd_Class = env.Vault.PropertyDefOperations.GetBuiltInPropertyDef(MFBuiltInPropertyDef.MFBuiltInPropertyDefClass);

            var oPropertyValues = new PropertyValues();
            oPropertyValues = env.Vault.ObjectPropertyOperations.GetProperties(env.ObjVer);

            var iClase = oPropertyValues.SearchForPropertyEx(pd_Class.ID, true).TypedValue.GetLookupID();

            if (cl_OrdenCompraEmitidaProveedor == iClase)
            {
                var oLookupsProveedor = oPropertyValues.SearchForPropertyEx(pd_Proveedor, true).TypedValue.GetValueAsLookups();

                // Relacionar a entregables recibidos
                SetPropertiesGenerico(env, cl_EntregableRecibidoProveedor, pd_Proveedor, pd_EntregablesRecibidos, oLookupsProveedor);

                // Relacionar a proyectos relacionados
                SetPropertiesGenerico(env, cl_ProyectoServicioEspecializado, pd_Proveedor, pd_ProyectosRelacionados, oLookupsProveedor);

                // Relacionar a contratos relacionados
                SetPropertiesGenerico(env, cl_Contrato, pd_Proveedor, pd_ContratosRelacionados, oLookupsProveedor, true);

                // Relacionar a facturas recibidas
                SetPropertiesGenerico(env, cl_FacturaRecibidaProveedor, pd_Proveedor, pd_FacturasRelacionadasRecibidas, oLookupsProveedor);
            }

            if (cl_OrdenCompraRecibidaCliente == iClase)
            {
                var oLookupsEmpresaInterna = oPropertyValues.SearchForPropertyEx(pd_EmpresaInterna, true).TypedValue.GetValueAsLookups();

                // Relacionar a entregables emitidos
                SetPropertiesGenerico(env, cl_EntregableEmitidoCliente, pd_EmpresaInterna, pd_EntregablesEmitidos, oLookupsEmpresaInterna);

                // Relacionar a proyectos relacionados
                SetPropertiesGenerico(env, cl_ProyectoCliente, pd_EmpresaInterna, pd_ProyectosRelacionados, oLookupsEmpresaInterna);

                // Relacionar a contratos relacionados
                SetPropertiesGenerico(env, cl_Contrato, pd_EmpresaInterna, pd_ContratosRelacionadosSE, oLookupsEmpresaInterna);
            }

            if (cl_FacturaRecibidaProveedor == iClase)
            {
                var oLookupsProveedor = oPropertyValues.SearchForPropertyEx(pd_Proveedor, true).TypedValue.GetValueAsLookups();

                // Relacionar a ordenes de compra emitidas
                SetPropertiesGenerico(env, cl_OrdenCompraEmitidaProveedor, pd_Proveedor, pd_OrdenesCompraEmitidas, oLookupsProveedor);

                // Relacionar a contratos relacionados
                SetPropertiesGenerico(env, cl_Contrato, pd_Proveedor, pd_ContratosRelacionados, oLookupsProveedor, true);
            }

            if (cl_Contrato == iClase)
            {
                var oLookupsProveedor = oPropertyValues.SearchForPropertyEx(pd_Proveedor, true).TypedValue.GetValueAsLookups();

                // Relacionar a entregables recibidos
                SetPropertiesGenerico(env, cl_EntregableRecibidoProveedor, pd_Proveedor, pd_EntregablesRecibidos, oLookupsProveedor);

                // Si el contrato es un convenio modificatorio, relacionar ordenes de compra emitidas y facturas recibidas
                if (!oPropertyValues.SearchForPropertyEx(pd_EsConvenioModificatorio, true).TypedValue.IsNULL() &&
                    Convert.ToBoolean(oPropertyValues.SearchForPropertyEx(pd_EsConvenioModificatorio, true).TypedValue.Value) == true)
                {
                    // Relacionar a ordenes de compra emitidas
                    SetPropertiesGenerico(env, cl_OrdenCompraEmitidaProveedor, pd_Proveedor, pd_OrdenesCompraEmitidas, oLookupsProveedor);

                    // Relacionar a facturas recibidas
                    SetPropertiesGenerico(env, cl_FacturaRecibidaProveedor, pd_Proveedor, pd_FacturasRelacionadasRecibidas, oLookupsProveedor);
                }

                // Relacionar a proyectos relacionados
                SetPropertiesGenerico(env, cl_ProyectoServicioEspecializado, pd_Proveedor, pd_ProyectosRelacionados, oLookupsProveedor);

                // Actualizar estatus de flujo de excepciones en el proveedor relacionado
                UpdateEstatusDeFirmaFlujoExcepcionesProveedor(env);
            }
        }

        [EventHandler(MFEventHandlerType.MFEventHandlerBeforeCreateNewObjectFinalize, Class = "CL.EntregableRecibidoProveedor")]
        [EventHandler(MFEventHandlerType.MFEventHandlerBeforeCreateNewObjectFinalize, Class = "CL.EntregableEmitidoCliente")]
        [EventHandler(MFEventHandlerType.MFEventHandlerBeforeCheckInChanges, Class = "CL.EntregableRecibidoProveedor")]
        [EventHandler(MFEventHandlerType.MFEventHandlerBeforeCheckInChanges, Class = "CL.EntregableEmitidoCliente")]
        public void ProcesosDocumentoEntregable(EventHandlerEnvironment env)
        {
            var cl_OrdenCompraEmitidaProveedor = env.Vault.ClassOperations.GetObjectClassIDByAlias("CL.OrdenDeCompraEmitidaProveedor");
            var cl_OrdenCompraRecibidaCliente = env.Vault.ClassOperations.GetObjectClassIDByAlias("CL.OrdenDeCompraRecibidaCliente");
            var cl_EntregableRecibidoProveedor = env.Vault.ClassOperations.GetObjectClassIDByAlias("CL.EntregableRecibidoProveedor");
            var cl_EntregableEmitidoCliente = env.Vault.ClassOperations.GetObjectClassIDByAlias("CL.EntregableEmitidoCliente");
            var cl_ProyectoServicioEspecializado = env.Vault.ClassOperations.GetObjectClassIDByAlias("CL.ProyectoServicioEspecializado");
            var cl_ProyectoCliente = env.Vault.ClassOperations.GetObjectClassIDByAlias("MF.CL.Project");
            var cl_Contrato = env.Vault.ClassOperations.GetObjectClassIDByAlias("CL.Contrato");
            var pd_Proveedor = env.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("MF.PD.Proveedor");
            var pd_EmpresaInterna = env.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.EmpresaInterna");
            var pd_OrdenesCompraEmitidas = env.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.OrdenesDeCompraEmitidas");
            var pd_OrdenesCompraRecibidas = env.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.OrdenesDeCompraRecibidas");
            var pd_ProyectosRelacionados = env.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("MF.PD.Project");
            var pd_ContratosRelacionados = env.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("MF.PD.RelatedContract");
            var pd_ContratosRelacionadosSE = env.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.ContratosRelacionados");
            var pd_Class = env.Vault.PropertyDefOperations.GetBuiltInPropertyDef(MFBuiltInPropertyDef.MFBuiltInPropertyDefClass);

            var oPropertyValues = new PropertyValues();
            oPropertyValues = env.Vault.ObjectPropertyOperations.GetProperties(env.ObjVer);

            var iClase = oPropertyValues.SearchForPropertyEx(pd_Class.ID, true).TypedValue.GetLookupID();

            if (cl_EntregableRecibidoProveedor == iClase)
            {
                var oLookupsProveedor = oPropertyValues.SearchForPropertyEx(pd_Proveedor, true).TypedValue.GetValueAsLookups();

                // Relacionar a ordenes de compra emitidas
                SetPropertiesGenerico(env, cl_OrdenCompraEmitidaProveedor, pd_Proveedor, pd_OrdenesCompraEmitidas, oLookupsProveedor);

                // Relacionar a proyectos relacionados
                SetPropertiesGenerico(env, cl_ProyectoServicioEspecializado, pd_Proveedor, pd_ProyectosRelacionados, oLookupsProveedor);

                // Relacionar a contratos relacionados
                SetPropertiesGenerico(env, cl_Contrato, pd_Proveedor, pd_ContratosRelacionados, oLookupsProveedor);
            }

            if (cl_EntregableEmitidoCliente == iClase)
            {
                var oLookupsEmpresaInterna = oPropertyValues.SearchForPropertyEx(pd_EmpresaInterna, true).TypedValue.GetValueAsLookups();

                // Relacionar a ordenes de compra recibidas
                SetPropertiesGenerico(env, cl_OrdenCompraRecibidaCliente, pd_EmpresaInterna, pd_OrdenesCompraRecibidas, oLookupsEmpresaInterna);

                // Relacionar a proyectos relacionados
                SetPropertiesGenerico(env, cl_ProyectoCliente, pd_EmpresaInterna, pd_ProyectosRelacionados, oLookupsEmpresaInterna);

                // Relacionar a contratos relacionados SE
                SetPropertiesGenerico(env, cl_Contrato, pd_EmpresaInterna, pd_ContratosRelacionadosSE, oLookupsEmpresaInterna);
            }
        }

        //[EventHandler(MFEventHandlerType.MFEventHandlerBeforeCreateNewObjectFinalize, Class = "CL.ProyectoServicioEspecializado")]
        //[EventHandler(MFEventHandlerType.MFEventHandlerBeforeCreateNewObjectFinalize, Class = "MF.CL.Project")]        
        //[EventHandler(MFEventHandlerType.MFEventHandlerBeforeCheckInChanges, Class = "CL.ProyectoServicioEspecializado")]
        //[EventHandler(MFEventHandlerType.MFEventHandlerBeforeCheckInChanges, Class = "MF.CL.Project")]        
        //public void ProcesosProyecto(EventHandlerEnvironment env)
        //{
        //    var cl_OrdenCompraEmitidaProveedor = env.Vault.ClassOperations.GetObjectClassIDByAlias("CL.OrdenDeCompraEmitidaProveedor");
        //    var cl_FacturaRecibidaProveedor = env.Vault.ClassOperations.GetObjectClassIDByAlias("CL.FacturaRecibidaProveedor");
        //    var cl_OrdenCompraRecibidaCliente = env.Vault.ClassOperations.GetObjectClassIDByAlias("CL.OrdenDeCompraRecibidaCliente");
        //    var cl_EntregableRecibidoProveedor = env.Vault.ClassOperations.GetObjectClassIDByAlias("CL.EntregableRecibidoProveedor");
        //    var cl_EntregableEmitidoCliente = env.Vault.ClassOperations.GetObjectClassIDByAlias("CL.EntregableEmitidoCliente");
        //    var cl_ProyectoServicioEspecializado = env.Vault.ClassOperations.GetObjectClassIDByAlias("CL.ProyectoServicioEspecializado");
        //    var cl_ProyectoCliente = env.Vault.ClassOperations.GetObjectClassIDByAlias("MF.CL.Project");
        //    var cl_Contrato = env.Vault.ClassOperations.GetObjectClassIDByAlias("CL.Contrato");
        //    var pd_Proveedor = env.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("MF.PD.Proveedor");
        //    var pd_EmpresaInterna = env.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.EmpresaInterna");
        //    var pd_OrdenesCompraEmitidas = env.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.OrdenesDeCompraEmitidas");
        //    var pd_OrdenesCompraRecibidas = env.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.OrdenesDeCompraRecibidas");
        //    var pd_EntregablesRecibidos = env.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.EntregablesRecibidos");
        //    var pd_EntregablesEmitidos = env.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.EntregablesEmitidos");
        //    var pd_ProyectosRelacionados = env.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("MF.PD.Project");
        //    var pd_ContratosRelacionados = env.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("MF.PD.RelatedContract");
        //    var pd_ContratosRelacionadosSE = env.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.ContratosRelacionados");
        //    var pd_EsConvenioModificatorio = env.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.EsConvenioModificatorio");
        //    var pd_FacturasRelacionadasRecibidas = env.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.FacturasRelacionadasRecibidas");
        //    var pd_Class = env.Vault.PropertyDefOperations.GetBuiltInPropertyDef(MFBuiltInPropertyDef.MFBuiltInPropertyDefClass);

        //    var oPropertyValues = new PropertyValues();
        //    oPropertyValues = env.Vault.ObjectPropertyOperations.GetProperties(env.ObjVer);

        //    var iClase = oPropertyValues.SearchForPropertyEx(pd_Class.ID, true).TypedValue.GetLookupID();

        //    if (cl_ProyectoServicioEspecializado == iClase)
        //    {
        //        var oLookupsProveedor = oPropertyValues.SearchForPropertyEx(pd_Proveedor, true).TypedValue.GetValueAsLookups();

        //        // Relacionar a entregables recibidos
        //        SetPropertiesGenerico(env, cl_EntregableRecibidoProveedor, pd_Proveedor, pd_EntregablesRecibidos, oLookupsProveedor, false);

        //        // Relacionar a ordenes de compra emitidas
        //        SetPropertiesGenerico(env, cl_OrdenCompraEmitidaProveedor, pd_Proveedor, pd_OrdenesCompraEmitidas, oLookupsProveedor, false);

        //        // Relacionar a contratos relacionados
        //        SetPropertiesGenerico(env, cl_Contrato, pd_Proveedor, pd_ContratosRelacionados, oLookupsProveedor, false);

        //        // Actualizar estatus de flujo de excepciones en el proveedor relacionado
        //        UpdateEstatusDeFirmaFlujoExcepcionesProveedor(env);
        //    }

        //    if (cl_ProyectoCliente == iClase)
        //    {
        //        var oLookupsEmpresaInterna = oPropertyValues.SearchForPropertyEx(pd_EmpresaInterna, true).TypedValue.GetValueAsLookups();

        //        // Relacionar a entregables emitidos
        //        SetPropertiesGenerico(env, cl_EntregableEmitidoCliente, pd_EmpresaInterna, pd_EntregablesEmitidos, oLookupsEmpresaInterna, false);

        //        // Relacionar a ordenes de compra recibidas
        //        SetPropertiesGenerico(env, cl_OrdenCompraRecibidaCliente, pd_EmpresaInterna, pd_OrdenesCompraRecibidas, oLookupsEmpresaInterna, false);

        //        // Relacionar a contratos relacionados
        //        SetPropertiesGenerico(env, cl_Contrato, pd_EmpresaInterna, pd_ContratosRelacionados, oLookupsEmpresaInterna, false);
        //    }           
        //}       
                                
        [EventHandler(MFEventHandlerType.MFEventHandlerAfterFileUpload, Class = "CL.CreacionMasivaDeProceedoresServiciosEspecializados")]
        public void Asignar_Workflow_CreacionMasivaProveedores(EventHandlerEnvironment env)
        {
            var workflow = PermanentVault
                .WorkflowOperations
                .GetWorkflowIDByAlias("WF.CargaMasivaDeProveedores");

            var estado = PermanentVault
                .WorkflowOperations
                .GetWorkflowStateIDByAlias("WFS.CargaMasivaDeProveedores.Inicio");

            var pd_EstadoProcesamiento = PermanentVault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.EstadoDeProcesamiento");

            try
            {
                var oWorkflowState = new ObjectVersionWorkflowState();
                var oObjVer = env.Vault.ObjectOperations.GetLatestObjVerEx(env.ObjVerEx.ObjID, true);

                oWorkflowState.Workflow.TypedValue.SetValue(MFDataType.MFDatatypeLookup, workflow);
                oWorkflowState.State.TypedValue.SetValue(MFDataType.MFDatatypeLookup, estado);
                env.Vault.ObjectPropertyOperations.SetWorkflowStateEx(oObjVer, oWorkflowState);

                // Actualizacion de estado de procesamiento
                var oLookup = new Lookup();
                var oObjID = new ObjID();

                oObjID.SetIDs
                (
                    ObjType: (int)MFBuiltInObjectType.MFBuiltInObjectTypeDocument,
                    ID: env.ObjVer.ID
                );

                var oPropertyValue = new PropertyValue
                {
                    PropertyDef = pd_EstadoProcesamiento
                };

                oLookup.Item = 3;

                oPropertyValue.TypedValue.SetValueToLookup(oLookup);

                env.Vault.ObjectPropertyOperations.SetProperty
                (
                    ObjVer: env.ObjVer,
                    PropertyValue: oPropertyValue
                );
            }
            catch (Exception ex)
            {
                SysUtils.ReportErrorMessageToEventLog("Ocurrio un error al asignar workflow en Creacion Masiva de Proveedores.", ex);
            }
        }

        [StateAction("WFS.CargaMasivaDeProveedores.Procesar", Class = "CL.CreacionMasivaDeProceedoresServiciosEspecializados")]
        public void Procesar_CreacionMasivaProveedores(StateEnvironment env)
        {
            // Archivos que se deberan limpiar al final del proceso
            var filesToDelete = new List<string>();

            if (env.ObjVerEx.IsEnteringState)
            {
                try
                {
                    var oObjVerEx = env.ObjVerEx;
                    var oObjectFiles = oObjVerEx.Info.Files;
                    IEnumerator enumerator = oObjectFiles.GetEnumerator();

                    while (enumerator.MoveNext())
                    {
                        // Descargar el archivo temporal
                        ObjectFile oFile = (ObjectFile)enumerator.Current;
                        string sPathTempFile = SysUtils.GetTempFileName(".tmp"); //@"C:\temp\CargaMasivaProveedores\temp\tempFile.xlsx"; //
                        filesToDelete.Add(sPathTempFile);

                        FileVer fileVer = oFile.FileVer;

                        env.Vault.ObjectFileOperations.DownloadFile(oFile.ID, fileVer.Version, sPathTempFile);

                        string sNewPathTempFile = @"C:\temp\CargaMasivaProveedores\temp\tempFileCopy.xls";
                        filesToDelete.Add(sNewPathTempFile);

                        File.Copy(sPathTempFile, sNewPathTempFile, true);

                        if (File.Exists(sNewPathTempFile))
                        {
                            CreateNewObjectCreacionMasivaDeProveedores(env, sPathTempFile);
                        }                        
                    }
                }
                catch (Exception ex)
                {
                    SysUtils.ReportErrorMessageToEventLog("Error al intentar procesar el documento.", ex);
                }
                finally
                {
                    foreach (var sFile in filesToDelete)
                    {
                        File.Delete(sFile);
                    }
                }
            }
        }

        //[EventHandler(MFEventHandlerType.MFEventHandlerBeforeCheckInChanges, Class = "CL.CreacionMasivaDeProceedoresServiciosEspecializados")]
        [EventHandler(MFEventHandlerType.MFEventHandlerAfterCreateNewObjectFinalize, Class = "CL.CreacionMasivaDeProveedores")]
        public void EventHandler_CreacionMasivaProveedores(EventHandlerEnvironment env)
        {
            bool bProcesoExitoso = false;
            string sPathExcelFile = "";

            // Inicializar objetos, propiedades, etc.
            var ot_Proveedor = env.Vault.ObjectTypeOperations.GetObjectTypeIDByAlias("OT.Proveedor");
            var pd_CrearListadoProveedores = env.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.CrearListadoDeProveedores");
            var pd_EstadoListadoProveedores = env.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.EstadoListadoProveedores");
            var pd_RfcEmpresa = env.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.RfcEmpresa");
            var pd_EstadoProcesamiento = env.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.EstadoDeProcesamiento");
            var oPropertyValues = new PropertyValues();

            var workflow = PermanentVault
                .WorkflowOperations
                .GetWorkflowIDByAlias("WF.CargaMasivaDeProveedores");

            var estado = PermanentVault
                .WorkflowOperations
                .GetWorkflowStateIDByAlias("WFS.CargaMasivaDeProveedores.Terminado");

            // Archivos que se deberan limpiar al final del proceso
            var filesToDelete = new List<string>();
            
            try
            {
                oPropertyValues = env.ObjVerEx.Properties;

                ////////////////////////////////////////////////////////////////
                ////////// Proceso Crear Listado de Proveedores ///////////////
                //////////////////////////////////////////////////////////////     

                // Si la propiedad crear listado de proveedores es true, generar el archivo excel
                if (!oPropertyValues.SearchForPropertyEx(pd_CrearListadoProveedores, true).TypedValue.IsNULL() &&
                    Convert.ToBoolean(oPropertyValues.SearchForPropertyEx(pd_CrearListadoProveedores, true).TypedValue.Value) == true)
                {
                    // Verificar estado de listado de proveedores
                    if (oPropertyValues.SearchForPropertyEx(pd_EstadoListadoProveedores, true).TypedValue.IsNULL() ||
                        oPropertyValues.SearchForPropertyEx(pd_EstadoListadoProveedores, true).TypedValue.GetLookupID() == 1 ||
                        oPropertyValues.SearchForPropertyEx(pd_EstadoListadoProveedores, true).TypedValue.GetLookupID() == 4)
                    {
                        // Subdirectorio temporal
                        string sTempTempPath = @"C:\temp\CargaMasivaProveedores\temp\";

                        if (!Directory.Exists(sTempTempPath))
                            Directory.CreateDirectory(sTempTempPath);

                        string[] files = Directory.GetFiles(sTempTempPath, "*.xlsx");

                        if (files.Length == 0)
                        {
                            // Abrir la plantilla del listado de proveedores
                            Application oExcelFile = new Application();
                            Workbook wb = oExcelFile.Workbooks.Open(@"C:\temp\CargaMasivaProveedores\ListadoDeProveedores.xlsx");
                            Worksheet ws1 = wb.Sheets[1];
                            ws1.Activate();

                            // Obtener todos los proveedores existentes en la boveda
                            var searchBuilderProveedor = new MFSearchBuilder(env.Vault);
                            searchBuilderProveedor.Deleted(false);
                            searchBuilderProveedor.ObjType(ot_Proveedor);

                            var searchResultsProveedor = searchBuilderProveedor.FindEx();

                            foreach (var proveedor in searchResultsProveedor)
                            {
                                var oPropertyValuesProveedor = proveedor.Properties;

                                var sNombreOTituloProveedor = oPropertyValuesProveedor
                                    .SearchForPropertyEx((int)MFBuiltInPropertyDef.MFBuiltInPropertyDefNameOrTitle, true)
                                    .TypedValue
                                    .Value.ToString();

                                var sRfcProveedor = oPropertyValuesProveedor.SearchForPropertyEx(pd_RfcEmpresa, true).TypedValue.Value.ToString();

                                // Obtener el row en el que se agregar la informacion del proveedor                           
                                Range usedRange = ws1.UsedRange;
                                int rowCount = usedRange.Rows.Count;
                                int rowAdd = rowCount + 1;

                                // Agregar los datos del proveedor a la plantilla
                                ws1.Cells[rowAdd, 1] = sRfcProveedor;
                                ws1.Cells[rowAdd, 2] = sNombreOTituloProveedor;

                                // Liberar el rango actual para obtener el siguiente
                                Marshal.ReleaseComObject(usedRange);
                            }

                            // Guardar el archivo excel despues de obtener la informacion de los proveedores dados de alta en la boveda
                            string sNameExcelFile = string.Format("ListadoDeProveedores_{0:yyyyMMdd_HHmmss}", DateTime.Now);
                            sPathExcelFile = Path.Combine(sTempTempPath, string.Concat(sNameExcelFile, ".xlsx"));
                            filesToDelete.Add(sPathExcelFile);

                            wb.SaveAs(sPathExcelFile);
                            wb.Close(true, Type.Missing, Type.Missing);
                            oExcelFile.Quit();

                            // Limpiar objetos
                            GC.Collect();
                            GC.WaitForPendingFinalizers();

                            // Liberar objetos para matar el proceso Excel que esta corriendo por detras del sistema                        
                            Marshal.ReleaseComObject(ws1);
                            Marshal.ReleaseComObject(wb);
                            Marshal.ReleaseComObject(oExcelFile);

                            // Modificar el documento de carga masiva de empleados
                            var oLookupCrearListado = new Lookup();
                            var objIDCrearListado = new ObjID();

                            objIDCrearListado.SetIDs
                            (
                                ObjType: (int)MFBuiltInObjectType.MFBuiltInObjectTypeDocument,
                                ID: env.ObjVerEx.ID
                            );

                            var oPropertyValueCrearListado = new PropertyValue
                            {
                                PropertyDef = pd_EstadoListadoProveedores
                            };

                            oLookupCrearListado.Item = 2;

                            oPropertyValueCrearListado.TypedValue.SetValueToLookup(oLookupCrearListado);

                            // Si el documento esta establecido como single file, modificar a multi file
                            if (env.ObjVer.Type == (int)MFBuiltInObjectType.MFBuiltInObjectTypeDocument
                                && env.Vault.ObjectOperations.IsSingleFileObject(env.ObjVer) == true)
                            {
                                env.Vault.ObjectOperations.SetSingleFileObject
                                (
                                    ObjVer: env.ObjVer,
                                    SingleFile: false
                                );
                            }

                            // Agregar el archivo al objeto
                            env.Vault.ObjectFileOperations.AddFile(
                                ObjVer: env.ObjVer,
                                Title: sNameExcelFile,
                                Extension: "xlsx",
                                SourcePath: sPathExcelFile);

                            // Actualizar estado de listado de proveedores
                            env.Vault.ObjectPropertyOperations.SetProperty
                            (
                                ObjVer: env.ObjVer,
                                PropertyValue: oPropertyValueCrearListado
                            );
                        }

                        // Establecer el estatus "En Proceso" de la propiedad Estado de Procesamiento
                        var oLookup = new Lookup();
                        var oObjID = new ObjID();

                        oObjID.SetIDs
                        (
                            ObjType: (int)MFBuiltInObjectType.MFBuiltInObjectTypeDocument,
                            ID: env.ObjVer.ID
                        );

                        var oPropertyValue = new PropertyValue
                        {
                            PropertyDef = pd_EstadoProcesamiento
                        };

                        oLookup.Item = 3;

                        oPropertyValue.TypedValue.SetValueToLookup(oLookup);

                        env.Vault.ObjectPropertyOperations.SetProperty
                        (
                            ObjVer: env.ObjVer,
                            PropertyValue: oPropertyValue
                        );

                        SysUtils.ReportInfoToEventLog("Se genero el listado de proveedores");
                    }

                    if (!oPropertyValues.SearchForPropertyEx(pd_EstadoListadoProveedores, true).TypedValue.IsNULL() &&
                        oPropertyValues.SearchForPropertyEx(pd_EstadoListadoProveedores, true).TypedValue.GetLookupID() == 2)
                    {
                        var oObjVerEx = env.ObjVerEx;
                        var oObjectFiles = oObjVerEx.Info.Files;
                        IEnumerator enumerator = oObjectFiles.GetEnumerator();

                        while (enumerator.MoveNext())
                        {
                            // Descargar el excel como un archivo temporal
                            ObjectFile oFile = (ObjectFile)enumerator.Current;
                            string sPathTempFile = SysUtils.GetTempFileName(".tmp");
                            filesToDelete.Add(sPathTempFile);
                            FileVer fileVer = oFile.FileVer;

                            env.Vault.ObjectFileOperations.DownloadFile(oFile.ID, fileVer.Version, sPathTempFile);

                            Application oExcelFile = new Application();
                            Workbook wb = oExcelFile.Workbooks.Open(sPathTempFile);

                            // Leer Sheet1 del excel
                            Worksheet ws1 = wb.Sheets[1];
                            ws1.Activate();
                            Range oRangeColumnsWS1 = ws1.UsedRange;

                            int rowCountWS1 = oRangeColumnsWS1.Rows.Count;

                            for (int i = 3; i <= rowCountWS1; i++)
                            {
                                // Datos del proveedor
                                string sRfcProveedor = "";
                                string sNombreProveedor = "";
                                string sTipoProveedor = "";
                                string sTipoValidacionChecklist = "";
                                string sTipoPersona = "";

                                if (!(oRangeColumnsWS1.Cells[i, 1].Value2 is null))
                                {
                                    sRfcProveedor = oRangeColumnsWS1.Cells[i, 1].Value2.ToString();
                                    sNombreProveedor = oRangeColumnsWS1.Cells[i, 2].Value2.ToString();
                                    sTipoProveedor = oRangeColumnsWS1.Cells[i, 3].Value2.ToString();
                                    sTipoValidacionChecklist = oRangeColumnsWS1.Cells[i, 4].Value2.ToString();
                                    sTipoPersona = oRangeColumnsWS1.Cells[i, 5].Value2.ToString();
                                    var sFechaInicioProveedor = oRangeColumnsWS1.Cells[i, 6].Value2;

                                    // Crear o actualizar el proveedor
                                    if (CreateOrUpdateProveedor(env, sRfcProveedor, sNombreProveedor, sTipoProveedor, sTipoValidacionChecklist, sTipoPersona, sFechaInicioProveedor))
                                    {
                                        // Actualizar estado de listado de proveedores
                                        var oLookupActualizar = new Lookup();
                                        var oObjIDActualizar = new ObjID();

                                        oObjIDActualizar.SetIDs
                                        (
                                            ObjType: (int)MFBuiltInObjectType.MFBuiltInObjectTypeDocument,
                                            ID: env.ObjVer.ID
                                        );

                                        var oPropertyValueActualizar = new PropertyValue
                                        {
                                            PropertyDef = pd_EstadoListadoProveedores
                                        };

                                        oLookupActualizar.Item = 3;

                                        oPropertyValueActualizar.TypedValue.SetValueToLookup(oLookupActualizar);

                                        env.Vault.ObjectPropertyOperations.SetProperty
                                        (
                                            ObjVer: env.ObjVer,
                                            PropertyValue: oPropertyValueActualizar
                                        );

                                        SysUtils.ReportInfoToEventLog("Proveedores actualizados exitosamente");

                                        if (SetPropertiesInProveedorAndContactoExterno(env, sRfcProveedor))
                                        {
                                            SysUtils.ReportInfoToEventLog("Se relacionaron contactos externos al proveedor");
                                        }
                                    }
                                }
                            }

                            // Limpiar objetos
                            GC.Collect();
                            GC.WaitForPendingFinalizers();

                            // Liberar objetos para matar el proceso Excel que esta corriendo por detras del sistema
                            Marshal.ReleaseComObject(oRangeColumnsWS1);
                            Marshal.ReleaseComObject(ws1);

                            // Cerrar y liberar
                            wb.Close();
                            Marshal.ReleaseComObject(wb);

                            // Quitar y liberar
                            oExcelFile.Quit();
                            Marshal.ReleaseComObject(oExcelFile);
                        }

                        // Establecer el estatus "Documento Procesado" de la propiedad Estado de Procesamiento
                        var oLookup = new Lookup();
                        var oObjID = new ObjID();

                        oObjID.SetIDs
                        (
                            ObjType: (int)MFBuiltInObjectType.MFBuiltInObjectTypeDocument,
                            ID: env.ObjVer.ID
                        );

                        var oPropertyValue = new PropertyValue
                        {
                            PropertyDef = pd_EstadoProcesamiento
                        };

                        oLookup.Item = 1;

                        oPropertyValue.TypedValue.SetValueToLookup(oLookup);

                        env.Vault.ObjectPropertyOperations.SetProperty
                        (
                            ObjVer: env.ObjVer,
                            PropertyValue: oPropertyValue
                        );

                        // Actualizacion de estado del workflow
                        var oWorkflowState = new ObjectVersionWorkflowState();
                        var oObjVer = env.Vault.ObjectOperations.GetLatestObjVerEx(env.ObjVerEx.ObjID, true);

                        oWorkflowState.Workflow.TypedValue.SetValue(MFDataType.MFDatatypeLookup, workflow);
                        oWorkflowState.State.TypedValue.SetValue(MFDataType.MFDatatypeLookup, estado);
                        env.Vault.ObjectPropertyOperations.SetWorkflowStateEx(oObjVer, oWorkflowState);
                    }
                }
                else
                {
                    ////////////////////////////////////////////////////////////////
                    ///////// Proceso Creacion Masiva de Proveedores //////////////
                    ////////////////////////////////////////////////////////////// 

                    SysUtils.ReportInfoToEventLog("Inicia proceso de creacion masiva de proveedores");

                    var oObjVerEx = env.ObjVerEx;
                    var oObjectFiles = oObjVerEx.Info.Files;
                    IEnumerator enumerator = oObjectFiles.GetEnumerator();

                    while (enumerator.MoveNext())
                    {
                        // Descargar el excel como un archivo temporal
                        ObjectFile oFile = (ObjectFile)enumerator.Current;
                        string sPathTempFile = SysUtils.GetTempFileName(".tmp");
                        filesToDelete.Add(sPathTempFile);

                        FileVer fileVer = oFile.FileVer;

                        env.Vault.ObjectFileOperations.DownloadFile(oFile.ID, fileVer.Version, sPathTempFile);

                        SysUtils.ReportInfoToEventLog("Ruta de descarga del archivo temporal: " + sPathTempFile);

                        //string sNewPathTempFile = @"C:\temp\CargaMasivaProveedores\temp\tempFile.xlsx";
                        //filesToDelete.Add(sNewPathTempFile);
                        //File.Copy(sPathTempFile, sNewPathTempFile, true);
                        //SysUtils.ReportInfoToEventLog("Ruta de descarga de la copia del archivo temporal: " + sNewPathTempFile);

                        SLDocument slExcel = new SLDocument(sPathTempFile);

                        // Contactos Externos Administradores
                        slExcel.SelectWorksheet("Sheet2");

                        int iRowSheet2 = 2;

                        while (!string.IsNullOrEmpty(slExcel.GetCellValueAsString(iRowSheet2, 1)))
                        {
                            string rfcProveedor = slExcel.GetCellValueAsString(iRowSheet2, 1);
                            string nombre = slExcel.GetCellValueAsString(iRowSheet2, 2);
                            string apellidoP = slExcel.GetCellValueAsString(iRowSheet2, 3);
                            string apellidoM = slExcel.GetCellValueAsString(iRowSheet2, 4);
                            string curp = slExcel.GetCellValueAsString(iRowSheet2, 5);
                            string email = slExcel.GetCellValueAsString(iRowSheet2, 6);

                            SysUtils.ReportInfoToEventLog("RFC del proveedor: " + rfcProveedor);

                            // Actualizar o crear contacto externo admin
                            if (CreateOrUpdateContactoExternoAdmin(env, rfcProveedor, nombre, apellidoP, apellidoM, curp, email))
                            {
                                bProcesoExitoso = true;
                            }

                            iRowSheet2++;
                        }

                        // Proveedores
                        slExcel.SelectWorksheet("Sheet1");

                        int iRowSheet1 = 2;

                        while (!string.IsNullOrEmpty(slExcel.GetCellValueAsString(iRowSheet1, 1)))
                        {
                            string rfcProveedor = slExcel.GetCellValueAsString(iRowSheet1, 1);
                            string nombreProveedor = slExcel.GetCellValueAsString(iRowSheet1, 2);
                            string tipoProveedor = slExcel.GetCellValueAsString(iRowSheet1, 3);
                            string tipoValChecklist = slExcel.GetCellValueAsString(iRowSheet1, 4);
                            string tipoPersona = slExcel.GetCellValueAsString(iRowSheet1, 5);
                            DateTime fechaInicio = slExcel.GetCellValueAsDateTime(iRowSheet1, 6);

                            // Crear o actualizar proveedor
                            if (CreateOrUpdateProveedor(env, rfcProveedor, nombreProveedor, tipoProveedor, tipoValChecklist, tipoPersona, fechaInicio))
                            {
                                if (bProcesoExitoso)
                                {
                                    if (SetPropertiesInProveedorAndContactoExterno(env, rfcProveedor))
                                    {
                                        SysUtils.ReportInfoToEventLog("Los proveedores fueron creados exitosamente");
                                    }
                                }
                            }

                            iRowSheet1++;
                        }

                        //Application oExcelFile = new Application();                       

                        //Workbook wb = oExcelFile.Workbooks.Open(sPathTempFile);

                        //// Leer Sheet1 del excel
                        //Worksheet ws1 = wb.Sheets[1];
                        //ws1.Activate();
                        //Range oRangeColumnsWS1 = ws1.UsedRange;

                        //int rowCountWS1 = oRangeColumnsWS1.Rows.Count;

                        //// Leer Sheet2 del excel
                        //Worksheet ws2 = wb.Sheets[2];
                        //ws2.Activate();
                        //Range oRangeColumnsWS2 = ws2.UsedRange;

                        //int rowCountWS2 = oRangeColumnsWS2.Rows.Count;

                        //for (int ii = 2; ii <= rowCountWS2; ii++)
                        //{
                        //    // Datos del contacto externo administrador
                        //    string sRfcProveedor = "";
                        //    string sNombre = "";
                        //    string sApellidoPaterno = "";
                        //    string sApellidoMaterno = "";
                        //    string sCurp = "";
                        //    string sEmail = "";

                        //    sRfcProveedor = oRangeColumnsWS2.Cells[ii, 1].Value2.ToString();
                        //    sNombre = oRangeColumnsWS2.Cells[ii, 2].Value2.ToString();
                        //    sApellidoPaterno = oRangeColumnsWS2.Cells[ii, 3].Value2.ToString();
                        //    sApellidoMaterno = oRangeColumnsWS2.Cells[ii, 4].Value2.ToString();
                        //    sCurp = oRangeColumnsWS2.Cells[ii, 5].Value2.ToString();
                        //    sEmail = oRangeColumnsWS2.Cells[ii, 6].Value2.ToString();

                        //    SysUtils.ReportInfoToEventLog("RFC del proveedor: " + sRfcProveedor);

                        //    // Crear o actualizar contacto externo administrador del proveedor
                        //    if (CreateOrUpdateContactoExternoAdmin(env, sRfcProveedor, sNombre, sApellidoPaterno, sApellidoMaterno, sCurp, sEmail))
                        //    {
                        //        bProcesoExitoso = true;
                        //    }
                        //}

                        //for (int i = 3; i <= rowCountWS1; i++)
                        //{
                        //    // Datos del proveedor
                        //    string sRfcProveedor = "";
                        //    string sNombreProveedor = "";
                        //    string sTipoProveedor = "";
                        //    string sTipoValidacionChecklist = "";
                        //    string sTipoPersona = "";

                        //    if (!(oRangeColumnsWS1.Cells[i, 1].Value2 is null))
                        //    {
                        //        SysUtils.ReportInfoToEventLog("La celda no es nula");

                        //        sRfcProveedor = oRangeColumnsWS1.Cells[i, 1].Value2.ToString();
                        //        sNombreProveedor = oRangeColumnsWS1.Cells[i, 2].Value2.ToString();
                        //        sTipoProveedor = oRangeColumnsWS1.Cells[i, 3].Value2.ToString();
                        //        sTipoValidacionChecklist = oRangeColumnsWS1.Cells[i, 4].Value2.ToString();
                        //        sTipoPersona = oRangeColumnsWS1.Cells[i, 5].Value2.ToString();
                        //        var sFechaInicioProveedor = oRangeColumnsWS1.Cells[i, 6].Value2;

                        //        // Crear o actualizar el proveedor
                        //        if (CreateOrUpdateProveedor(env, sRfcProveedor, sNombreProveedor, sTipoProveedor, sTipoValidacionChecklist, sTipoPersona, sFechaInicioProveedor))
                        //        {
                        //            if (bProcesoExitoso)
                        //            {
                        //                if (SetPropertiesInProveedorAndContactoExterno(env, sRfcProveedor))
                        //                {
                        //                    SysUtils.ReportInfoToEventLog("Los proveedores fueron creados exitosamente");
                        //                }
                        //            }
                        //        }
                        //    }                            
                        //}

                        //// Limpiar objetos
                        //GC.Collect();
                        //GC.WaitForPendingFinalizers();

                        //// Liberar objetos para matar el proceso Excel que esta corriendo por detras del sistema
                        //Marshal.ReleaseComObject(oRangeColumnsWS1);
                        //Marshal.ReleaseComObject(ws1);
                        //Marshal.ReleaseComObject(oRangeColumnsWS2);
                        //Marshal.ReleaseComObject(ws2);

                        //// Cerrar y liberar
                        //wb.Close();
                        //Marshal.ReleaseComObject(wb);

                        //// Quitar y liberar
                        //oExcelFile.Quit();
                        //Marshal.ReleaseComObject(oExcelFile);
                    }

                    // Establecer el estatus "Documento Procesado" de la propiedad Estado de Procesamiento
                    var oLookup = new Lookup();
                    var oObjID = new ObjID();

                    oObjID.SetIDs
                    (
                        ObjType: (int)MFBuiltInObjectType.MFBuiltInObjectTypeDocument,
                        ID: env.ObjVer.ID
                    );

                    var oPropertyValue = new PropertyValue
                    {
                        PropertyDef = pd_EstadoProcesamiento
                    };

                    oLookup.Item = 1;

                    oPropertyValue.TypedValue.SetValueToLookup(oLookup);

                    env.Vault.ObjectPropertyOperations.SetProperty
                    (
                        ObjVer: env.ObjVer,
                        PropertyValue: oPropertyValue
                    );

                    // Actualizacion de estado del workflow
                    var oWorkflowState = new ObjectVersionWorkflowState();
                    var oObjVer = env.Vault.ObjectOperations.GetLatestObjVerEx(env.ObjVerEx.ObjID, true);

                    oWorkflowState.Workflow.TypedValue.SetValue(MFDataType.MFDatatypeLookup, workflow);
                    oWorkflowState.State.TypedValue.SetValue(MFDataType.MFDatatypeLookup, estado);
                    env.Vault.ObjectPropertyOperations.SetWorkflowStateEx(oObjVer, oWorkflowState);
                }
            }
            catch (Exception ex)
            {
                // Establecer el estatus "No Procesado (Error)" de la propiedad Estado de Procesamiento
                var oLookup = new Lookup();
                var oObjID = new ObjID();

                oObjID.SetIDs
                (
                    ObjType: (int)MFBuiltInObjectType.MFBuiltInObjectTypeDocument,
                    ID: env.ObjVer.ID
                );

                var oPropertyValue = new PropertyValue
                {
                    PropertyDef = pd_EstadoProcesamiento
                };

                oLookup.Item = 2;

                oPropertyValue.TypedValue.SetValueToLookup(oLookup);

                env.Vault.ObjectPropertyOperations.SetProperty
                (
                    ObjVer: env.ObjVer,
                    PropertyValue: oPropertyValue
                );

                // Cerrar objetos abiertos de Interop Excel


                SysUtils.ReportErrorMessageToEventLog("Error en proceso de creacion masiva de proveedores.", ex);
            }
            finally
            {
                foreach (var sFile in filesToDelete)
                {
                    File.Delete(sFile);
                }
            }
        }

        [EventHandler(MFEventHandlerType.MFEventHandlerBeforeCreateNewObjectFinalize, Class = "CL.OrdenDeCompraEmitidaProveedor")]
        [EventHandler(MFEventHandlerType.MFEventHandlerBeforeCreateNewObjectFinalize, Class = "CL.OrdenDeCompraRecibidaCliente")]
        [EventHandler(MFEventHandlerType.MFEventHandlerBeforeCreateNewObjectFinalize, Class = "CL.FacturaRecibidaProveedor")]
        [EventHandler(MFEventHandlerType.MFEventHandlerBeforeCreateNewObjectFinalize, Class = "CL.Contrato")]
        public void ProcesosDocumentosAdministrativos(EventHandlerEnvironment env)
        {
            var wf_ValidacionesChecklist = env.Vault.WorkflowOperations.GetWorkflowIDByAlias("WF.ValidcionesRepse");
            var wfs_DocumentoPorTraducir = env.Vault.WorkflowOperations.GetWorkflowStateIDByAlias("WFS.ValidacionesChecklist.DocumentoPorTraducir");
            var wf_CicloVidaContrato = env.Vault.WorkflowOperations.GetWorkflowIDByAlias("MF.WF.ContractLifecycle");
            var wfs_PresentacionPendiente = env.Vault.WorkflowOperations.GetWorkflowStateIDByAlias("M-Files.CLM.State.ContractLifecycle.PendingSubmission");
            var wf_FlujoExcepcionesProveedor = env.Vault.WorkflowOperations.GetWorkflowIDByAlias("WF.FlujoDeExcepcionesProveedor");
            var wfs_EstadoSolicitarFirma = env.Vault.WorkflowOperations.GetWorkflowStateIDByAlias("WFS.FlujoDeExcepcionesProveedor.SolicitarFirma");
            var ot_Proveedor = env.Vault.ObjectTypeOperations.GetObjectTypeIDByAlias("OT.Proveedor");
            var cl_ProveedorServicioEspecializado = env.Vault.ClassOperations.GetObjectClassIDByAlias("CL.ProveedorDeServicioEspecializado");
            var cl_OrdenCompraEmitidaProveedor = env.Vault.ClassOperations.GetObjectClassIDByAlias("CL.OrdenDeCompraEmitidaProveedor");
            var cl_OrdenCompraRecibidaCliente = env.Vault.ClassOperations.GetObjectClassIDByAlias("CL.OrdenDeCompraRecibidaCliente");
            var cl_FacturaRecibidaProveedor = env.Vault.ClassOperations.GetObjectClassIDByAlias("CL.FacturaRecibidaProveedor");
            var cl_EntregableRecibidoProveedor = env.Vault.ClassOperations.GetObjectClassIDByAlias("CL.EntregableRecibidoProveedor");
            var cl_EntregableEmitidoCliente = env.Vault.ClassOperations.GetObjectClassIDByAlias("CL.EntregableEmitidoCliente");
            var cl_ProyectoServicioEspecializado = env.Vault.ClassOperations.GetObjectClassIDByAlias("CL.ProyectoServicioEspecializado");
            var cl_ProyectoCliente = env.Vault.ClassOperations.GetObjectClassIDByAlias("MF.CL.Project");
            var cl_Contrato = env.Vault.ClassOperations.GetObjectClassIDByAlias("CL.Contrato");
            var pd_EmpresaInterna = env.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.EmpresaInterna");
            var pd_EntregablesRecibidos = env.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.EntregablesRecibidos");
            var pd_EntregablesEmitidos = env.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.EntregablesEmitidos");
            var pd_ProyectosRelacionados = env.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("MF.PD.Project");
            var pd_ContratosRelacionados = env.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("MF.PD.RelatedContract");
            var pd_ContratosRelacionadosSE = env.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.ContratosRelacionados");
            var pd_FacturasRelacionadasRecibidas = env.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.FacturasRelacionadasRecibidas");
            var pd_RfcEmpresa = env.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.RfcEmpresa");
            var pd_Proveedor = env.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("MF.PD.Proveedor");
            var pd_Severidad = env.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("MF.PD.Severity");
            var pd_EstatusFlujoExcepciones = env.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.EstatusFlujoDeExcepciones");
            var pd_OrdenesCompraEmitidas = env.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.OrdenesDeCompraEmitidas");
            var pd_EsConvenioModificatorio = env.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.EsConvenioModificatorio");
            var pd_Class = env.Vault.PropertyDefOperations.GetBuiltInPropertyDef(MFBuiltInPropertyDef.MFBuiltInPropertyDefClass);

            var oPropertyValues = new PropertyValues();
            oPropertyValues = env.Vault.ObjectPropertyOperations.GetProperties(env.ObjVer);

            var iClase = oPropertyValues.SearchForPropertyEx(pd_Class.ID, true).TypedValue.GetLookupID();

            if (cl_OrdenCompraEmitidaProveedor == iClase)
            {
                var oLookupsProveedor = oPropertyValues.SearchForPropertyEx(pd_Proveedor, true).TypedValue.GetValueAsLookups();

                // Relacionar a entregables recibidos
                SetPropertiesGenerico(env, cl_EntregableRecibidoProveedor, pd_Proveedor, pd_EntregablesRecibidos, oLookupsProveedor);

                // Relacionar a proyectos relacionados
                SetPropertiesGenerico(env, cl_ProyectoServicioEspecializado, pd_Proveedor, pd_ProyectosRelacionados, oLookupsProveedor);

                // Relacionar a contratos relacionados
                SetPropertiesGenerico(env, cl_Contrato, pd_Proveedor, pd_ContratosRelacionados, oLookupsProveedor, true);
                
                // Relacionar a facturas recibidas
                SetPropertiesGenerico(env, cl_FacturaRecibidaProveedor, pd_Proveedor, pd_FacturasRelacionadasRecibidas, oLookupsProveedor);
            }

            if (cl_OrdenCompraRecibidaCliente == iClase)
            {
                var oLookupsEmpresaInterna = oPropertyValues.SearchForPropertyEx(pd_EmpresaInterna, true).TypedValue.GetValueAsLookups();

                // Relacionar a entregables emitidos
                SetPropertiesGenerico(env, cl_EntregableEmitidoCliente, pd_EmpresaInterna, pd_EntregablesEmitidos, oLookupsEmpresaInterna);

                // Relacionar a proyectos relacionados
                SetPropertiesGenerico(env, cl_ProyectoCliente, pd_EmpresaInterna, pd_ProyectosRelacionados, oLookupsEmpresaInterna);

                // Relacionar a contratos relacionados
                SetPropertiesGenerico(env, cl_Contrato, pd_EmpresaInterna, pd_ContratosRelacionadosSE, oLookupsEmpresaInterna);
            }

            if (cl_FacturaRecibidaProveedor == iClase)
            {
                var oLookupsProveedor = oPropertyValues.SearchForPropertyEx(pd_Proveedor, true).TypedValue.GetValueAsLookups();

                // Relacionar a ordenes de compra emitidas
                SetPropertiesGenerico(env, cl_OrdenCompraEmitidaProveedor, pd_Proveedor, pd_OrdenesCompraEmitidas, oLookupsProveedor);

                // Relacionar a contratos relacionados
                SetPropertiesGenerico(env, cl_Contrato, pd_Proveedor, pd_ContratosRelacionados, oLookupsProveedor, true);
            }

            if (cl_Contrato == iClase)
            {
                var oLookupsProveedor = oPropertyValues.SearchForPropertyEx(pd_Proveedor, true).TypedValue.GetValueAsLookups();

                // Relacionar a entregables recibidos
                SetPropertiesGenerico(env, cl_EntregableRecibidoProveedor, pd_Proveedor, pd_EntregablesRecibidos, oLookupsProveedor);

                // Si el contrato es un convenio modificatorio, relacionar ordenes de compra emitidas y facturas recibidas
                if (!oPropertyValues.SearchForPropertyEx(pd_EsConvenioModificatorio, true).TypedValue.IsNULL() &&
                    Convert.ToBoolean(oPropertyValues.SearchForPropertyEx(pd_EsConvenioModificatorio, true).TypedValue.Value) == true)
                {
                    // Relacionar a ordenes de compra emitidas
                    SetPropertiesGenerico(env, cl_OrdenCompraEmitidaProveedor, pd_Proveedor, pd_OrdenesCompraEmitidas, oLookupsProveedor);

                    // Relacionar a facturas recibidas
                    SetPropertiesGenerico(env, cl_FacturaRecibidaProveedor, pd_Proveedor, pd_FacturasRelacionadasRecibidas, oLookupsProveedor);
                }

                // Relacionar a proyectos relacionados
                SetPropertiesGenerico(env, cl_ProyectoServicioEspecializado, pd_Proveedor, pd_ProyectosRelacionados, oLookupsProveedor);
            }

            if (cl_OrdenCompraEmitidaProveedor == iClase || cl_FacturaRecibidaProveedor == iClase || cl_Contrato == iClase)
            {
                var oWorkflowstate = new ObjectVersionWorkflowState();
                var oObjVer = env.Vault.ObjectOperations.GetLatestObjVerEx(env.ObjVerEx.ObjID, true);

                // Verificar estatus de la firma del flujo de excepciones
                UpdateEstatusDeFirmaFlujoExcepcionesProveedor(env);

                //var sRfcProveedor = oPropertyValues.SearchForPropertyEx(pd_RfcEmpresa, true).TypedValue.Value.ToString();
                if (!oPropertyValues.SearchForPropertyEx(pd_Proveedor, true).TypedValue.IsNULL())
                {
                    var oListProveedor = oPropertyValues.SearchForPropertyEx(pd_Proveedor, true).TypedValue.GetValueAsLookups().ToObjVerExs(env.Vault);
                    var objVerProveedor = oListProveedor[0];

                    var oPropertyValuesProveedor = new PropertyValues();                    
                    oPropertyValuesProveedor = objVerProveedor.Properties;                                       

                    var iIdSeveridad = oPropertyValuesProveedor.SearchForPropertyEx(pd_Severidad, true).TypedValue.GetLookupID();

                    if (iIdSeveridad == 3 || iIdSeveridad == 4)
                    {
                        // Verificar si el proveedor con severidad rojo o naranja ya tiene la firma para asociar documentos
                        if (oPropertyValuesProveedor.IndexOf(pd_EstatusFlujoExcepciones) != -1)
                        {
                            var iEstatusFlujoExcepciones = oPropertyValuesProveedor.SearchForPropertyEx(pd_EstatusFlujoExcepciones, true).TypedValue.GetLookupID();                            

                            // Si el estatus es diferente Firmado se ingresa el documento al flujo de excepciones
                            if (iEstatusFlujoExcepciones != 2)
                            {
                                // Asignar flujo de excepciones para solicitar firma a director                                
                                oWorkflowstate.Workflow.TypedValue.SetValue(MFDataType.MFDatatypeLookup, wf_FlujoExcepcionesProveedor);
                                oWorkflowstate.State.TypedValue.SetValue(MFDataType.MFDatatypeLookup, wfs_EstadoSolicitarFirma);
                                env.Vault.ObjectPropertyOperations.SetWorkflowStateEx(oObjVer, oWorkflowstate);

                                // Asignar propiedad estatus del flujo de excepciones en la metadata del proveedor
                                var oLookup = new Lookup();
                                var oObjID = new ObjID();

                                oObjID.SetIDs
                                (
                                    ObjType: ot_Proveedor,
                                    ID: objVerProveedor.ObjVer.ID
                                );

                                var checkedOutObjectVersion = env.Vault.ObjectOperations.CheckOut(oObjID);

                                var oPropertyValue = new PropertyValue
                                {
                                    PropertyDef = pd_EstatusFlujoExcepciones
                                };

                                oLookup.Item = 1;

                                oPropertyValue.TypedValue.SetValueToLookup(oLookup);

                                env.Vault.ObjectPropertyOperations.SetProperty
                                (
                                    ObjVer: checkedOutObjectVersion.ObjVer,
                                    PropertyValue: oPropertyValue
                                );

                                env.Vault.ObjectOperations.CheckIn(checkedOutObjectVersion.ObjVer);
                            }
                            else
                            {
                                // Mover el Contrato firmado a un nuevo workflow
                                if (cl_Contrato == iClase)
                                {
                                    oWorkflowstate.Workflow.TypedValue.SetValue(MFDataType.MFDatatypeLookup, wf_CicloVidaContrato);
                                    oWorkflowstate.State.TypedValue.SetValue(MFDataType.MFDatatypeLookup, wfs_PresentacionPendiente);
                                    env.Vault.ObjectPropertyOperations.SetWorkflowStateEx(oObjVer, oWorkflowstate);
                                }

                                // Despues de la firma, mover la Orden de Compra Emitida o Factura Recibida a un nuevo workflow 
                                if (cl_OrdenCompraEmitidaProveedor == iClase || cl_FacturaRecibidaProveedor == iClase)
                                {
                                    oWorkflowstate.Workflow.TypedValue.SetValue(MFDataType.MFDatatypeLookup, wf_ValidacionesChecklist);
                                    oWorkflowstate.State.TypedValue.SetValue(MFDataType.MFDatatypeLookup, wfs_DocumentoPorTraducir);
                                    env.Vault.ObjectPropertyOperations.SetWorkflowStateEx(oObjVer, oWorkflowstate);
                                }
                            }
                        }
                        
                        //throw new Exception("No es posible asignar el Proveedor seleccionado a este documento, debido a que el Proveedor tiene una severidad de nivel Rojo o Naranja, al dar clic en e boton Guardar se solicitara la autorizacion para asignar al proveedor.");
                    }
                    else
                    {
                        // Mover el Contrato firmado a un nuevo workflow
                        if (cl_Contrato == iClase)
                        {
                            oWorkflowstate.Workflow.TypedValue.SetValue(MFDataType.MFDatatypeLookup, wf_CicloVidaContrato);
                            oWorkflowstate.State.TypedValue.SetValue(MFDataType.MFDatatypeLookup, wfs_PresentacionPendiente);
                            env.Vault.ObjectPropertyOperations.SetWorkflowStateEx(oObjVer, oWorkflowstate);
                        }

                        // Despues de la firma, mover la Orden de Compra Emitida o Factura Recibida a un nuevo workflow 
                        if (cl_OrdenCompraEmitidaProveedor == iClase || cl_FacturaRecibidaProveedor == iClase)
                        {
                            oWorkflowstate.Workflow.TypedValue.SetValue(MFDataType.MFDatatypeLookup, wf_ValidacionesChecklist);
                            oWorkflowstate.State.TypedValue.SetValue(MFDataType.MFDatatypeLookup, wfs_DocumentoPorTraducir);
                            env.Vault.ObjectPropertyOperations.SetWorkflowStateEx(oObjVer, oWorkflowstate);
                        }
                    }
                }
                else
                {
                    // Mover el Contrato firmado a un nuevo workflow
                    if (cl_Contrato == iClase)
                    {
                        oWorkflowstate.Workflow.TypedValue.SetValue(MFDataType.MFDatatypeLookup, wf_CicloVidaContrato);
                        oWorkflowstate.State.TypedValue.SetValue(MFDataType.MFDatatypeLookup, wfs_PresentacionPendiente);
                        env.Vault.ObjectPropertyOperations.SetWorkflowStateEx(oObjVer, oWorkflowstate);
                    }

                    // Despues de la firma, mover la Orden de Compra Emitida o Factura Recibida a un nuevo workflow 
                    if (cl_OrdenCompraEmitidaProveedor == iClase || cl_FacturaRecibidaProveedor == iClase)
                    {
                        oWorkflowstate.Workflow.TypedValue.SetValue(MFDataType.MFDatatypeLookup, wf_ValidacionesChecklist);
                        oWorkflowstate.State.TypedValue.SetValue(MFDataType.MFDatatypeLookup, wfs_DocumentoPorTraducir);
                        env.Vault.ObjectPropertyOperations.SetWorkflowStateEx(oObjVer, oWorkflowstate);
                    }
                }
            }  
        }

        [EventHandler(MFEventHandlerType.MFEventHandlerBeforeCheckInChanges, Class = "CL.Proveedor")]
        [EventHandler(MFEventHandlerType.MFEventHandlerBeforeCheckInChanges, Class = "CL.ProveedorDeServicioEspecializado")]
        [EventHandler(MFEventHandlerType.MFEventHandlerBeforeCheckInChanges, Class = "CL.ProveedorDependencia")]
        [EventHandler(MFEventHandlerType.MFEventHandlerBeforeCheckInChanges, Class = "CL.ProveedorEstratgico")]
        [EventHandler(MFEventHandlerType.MFEventHandlerBeforeCheckInChanges, Class = "CL.ProveedorExtranjero")]
        [EventHandler(MFEventHandlerType.MFEventHandlerBeforeCheckInChanges, Class = "CL.ProveedorTransportista")]
        public void ProcesosProveedores(EventHandlerEnvironment env)
        {
            var ot_Proveedor = env.Vault.ObjectTypeOperations.GetObjectTypeIDByAlias("OT.Proveedor");
            var pd_Severidad = env.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("MF.PD.Severity");
            var pd_EstatusFlujoExcepciones = env.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.EstatusFlujoDeExcepciones");

            var oPropertyValues = new PropertyValues();
            oPropertyValues = env.Vault.ObjectPropertyOperations.GetProperties(env.ObjVer);

            var iIdSeveridad = oPropertyValues.SearchForPropertyEx(pd_Severidad, true).TypedValue.GetLookupID();

            if (iIdSeveridad == 1 || iIdSeveridad == 2)
            {
                // Asignar propiedad estatus del flujo de excepciones en la metadata del proveedor
                var oLookup = new Lookup();
                var oObjID = new ObjID();

                oObjID.SetIDs
                (
                    ObjType: ot_Proveedor,
                    ID: env.ObjVer.ID
                );

                var oPropertyValue = new PropertyValue
                {
                    PropertyDef = pd_EstatusFlujoExcepciones
                };

                oLookup.Item = 5;

                oPropertyValue.TypedValue.SetValueToLookup(oLookup);

                env.Vault.ObjectPropertyOperations.SetProperty
                (
                    ObjVer: env.ObjVer,
                    PropertyValue: oPropertyValue
                );
            }
        }

        [StateAction("WFS.FlujoDeExcepcionesProveedor.Firmado")]
        public void FlujoDeExcepcionesProveedor_Firmado(StateEnvironment env)
        {
            var wf_CicloVidaContrato = env.Vault.WorkflowOperations.GetWorkflowIDByAlias("MF.WF.ContractLifecycle");
            var wfs_PresentacionPendiente = env.Vault.WorkflowOperations.GetWorkflowStateIDByAlias("M-Files.CLM.State.ContractLifecycle.PendingSubmission");
            var wf_ValidacionesChecklist = env.Vault.WorkflowOperations.GetWorkflowIDByAlias("WF.ValidcionesRepse");
            var wfs_DocumentoPorTraducir = env.Vault.WorkflowOperations.GetWorkflowStateIDByAlias("WFS.ValidacionesChecklist.DocumentoPorTraducir");
            var ot_Proveedor = env.Vault.ObjectTypeOperations.GetObjectTypeIDByAlias("OT.Proveedor");
            var cl_OrdenCompraEmitidaProveedor = env.Vault.ClassOperations.GetObjectClassIDByAlias("CL.OrdenDeCompraEmitidaProveedor");
            var cl_FacturaRecibidaProveedor = env.Vault.ClassOperations.GetObjectClassIDByAlias("CL.FacturaRecibidaProveedor");
            var cl_Contrato = env.Vault.ClassOperations.GetObjectClassIDByAlias("CL.Contrato");
            var pd_Proveedor = env.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("MF.PD.Proveedor");
            var pd_EstatusFlujoExcepciones = env.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.EstatusFlujoDeExcepciones");

            var oWorkflowstate = new ObjectVersionWorkflowState();
            var oObjVer = env.Vault.ObjectOperations.GetLatestObjVerEx(env.ObjVerEx.ObjID, true);

            var oPropertyValues = new PropertyValues();
            oPropertyValues = env.Vault.ObjectPropertyOperations.GetProperties(env.ObjVer);

            var iClase = oPropertyValues
                .SearchForPropertyEx((int)MFBuiltInPropertyDef.MFBuiltInPropertyDefClass, true)
                .TypedValue
                .GetLookupID();

            var oListProveedor = oPropertyValues.SearchForPropertyEx(pd_Proveedor, true).TypedValue.GetValueAsLookups().ToObjVerExs(env.Vault);
            var objVerProveedor = oListProveedor[0];

            // Asignar propiedad estatus del flujo de excepciones en la metadata del proveedor
            var oLookup = new Lookup();
            var oObjID = new ObjID();

            oObjID.SetIDs
            (
                ObjType: ot_Proveedor,
                ID: objVerProveedor.ObjVer.ID
            );

            var checkedOutObjectVersion = env.Vault.ObjectOperations.CheckOut(oObjID);

            var oPropertyValue = new PropertyValue
            {
                PropertyDef = pd_EstatusFlujoExcepciones
            };

            oLookup.Item = 2;

            oPropertyValue.TypedValue.SetValueToLookup(oLookup);

            env.Vault.ObjectPropertyOperations.SetProperty
            (
                ObjVer: checkedOutObjectVersion.ObjVer,
                PropertyValue: oPropertyValue
            );

            env.Vault.ObjectOperations.CheckIn(checkedOutObjectVersion.ObjVer);

            // Mover el Contrato firmado a un nuevo workflow
            if (cl_Contrato == iClase)
            {                
                oWorkflowstate.Workflow.TypedValue.SetValue(MFDataType.MFDatatypeLookup, wf_CicloVidaContrato);
                oWorkflowstate.State.TypedValue.SetValue(MFDataType.MFDatatypeLookup, wfs_PresentacionPendiente);
                env.Vault.ObjectPropertyOperations.SetWorkflowStateEx(oObjVer, oWorkflowstate);
            }

            // Despues de la firma, mover la Orden de Compra Emitida o Factura Recibida a un nuevo workflow 
            if (cl_OrdenCompraEmitidaProveedor == iClase || cl_FacturaRecibidaProveedor == iClase)
            {
                oWorkflowstate.Workflow.TypedValue.SetValue(MFDataType.MFDatatypeLookup, wf_ValidacionesChecklist);
                oWorkflowstate.State.TypedValue.SetValue(MFDataType.MFDatatypeLookup, wfs_DocumentoPorTraducir);
                env.Vault.ObjectPropertyOperations.SetWorkflowStateEx(oObjVer, oWorkflowstate);
            }
        }

        [StateAction("WFS.FlujoDeExcepcionesProveedor.Rechazado")]
        public void FlujoDeExcepcionesProveedor_Rechazado(StateEnvironment env)
        {
            var ot_Proveedor = env.Vault.ObjectTypeOperations.GetObjectTypeIDByAlias("OT.Proveedor");
            var pd_Proveedor = env.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("MF.PD.Proveedor");
            var pd_EstatusFlujoExcepciones = env.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.EstatusFlujoDeExcepciones");

            var oPropertyValues = new PropertyValues();
            oPropertyValues = env.Vault.ObjectPropertyOperations.GetProperties(env.ObjVer);

            var oListProveedor = oPropertyValues.SearchForPropertyEx(pd_Proveedor, true).TypedValue.GetValueAsLookups().ToObjVerExs(env.Vault);
            var objVerProveedor = oListProveedor[0];

            // Asignar propiedad estatus del flujo de excepciones en la metadata del proveedor
            var oLookup = new Lookup();
            var oObjID = new ObjID();

            oObjID.SetIDs
            (
                ObjType: ot_Proveedor,
                ID: objVerProveedor.ObjVer.ID
            );

            var checkedOutObjectVersion = env.Vault.ObjectOperations.CheckOut(oObjID);

            var oPropertyValue = new PropertyValue
            {
                PropertyDef = pd_EstatusFlujoExcepciones
            };

            oLookup.Item = 3;

            oPropertyValue.TypedValue.SetValueToLookup(oLookup);

            env.Vault.ObjectPropertyOperations.SetProperty
            (
                ObjVer: checkedOutObjectVersion.ObjVer,
                PropertyValue: oPropertyValue
            );

            env.Vault.ObjectOperations.CheckIn(checkedOutObjectVersion.ObjVer);
        }

        [EventHandler(MFEventHandlerType.MFEventHandlerBeforeCreateNewObjectFinalize, Class = "CL.CargaMasivaSinClasificar")]
        [EventHandler(MFEventHandlerType.MFEventHandlerAfterFileUpload, Class = "CL.CargaMasivaSinClasificar")]
        public void Asignar_Workflow_CargaMasivaSinClasificar(EventHandlerEnvironment env)
        {
            var workflow = PermanentVault
                .WorkflowOperations
                .GetWorkflowIDByAlias("WF.ValidacionesDocumentosChronoscan");

            var estado = PermanentVault
                .WorkflowOperations
                .GetWorkflowStateIDByAlias("WFS.ValidacionesDocumentosChronoscan.Inicio");

            var pd_EstadoProcesamiento = PermanentVault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.EstadoDeProcesamiento");

            try
            {
                var oWorkflowState = new ObjectVersionWorkflowState();
                var oObjVer = env.Vault.ObjectOperations.GetLatestObjVerEx(env.ObjVerEx.ObjID, true);

                oWorkflowState.Workflow.TypedValue.SetValue(MFDataType.MFDatatypeLookup, workflow);
                oWorkflowState.State.TypedValue.SetValue(MFDataType.MFDatatypeLookup, estado);
                env.Vault.ObjectPropertyOperations.SetWorkflowStateEx(oObjVer, oWorkflowState);

                // Actualizacion de estado de procesamiento
                var oLookup = new Lookup();
                var oObjID = new ObjID();

                oObjID.SetIDs
                (
                    ObjType: (int)MFBuiltInObjectType.MFBuiltInObjectTypeDocument,
                    ID: env.ObjVer.ID
                );

                var oPropertyValue = new PropertyValue
                {
                    PropertyDef = pd_EstadoProcesamiento
                };

                oLookup.Item = 3;

                oPropertyValue.TypedValue.SetValueToLookup(oLookup);

                env.Vault.ObjectPropertyOperations.SetProperty
                (
                    ObjVer: env.ObjVer,
                    PropertyValue: oPropertyValue
                );
            }
            catch (Exception ex)
            {
                SysUtils.ReportErrorMessageToEventLog("Ocurrio un error al asignar workflow en Carga Masiva sin Clasificar", ex);
            }
        }

        async Task PutTaskDelay()
        {
            TimeSpan timeSpan = TimeSpan.FromSeconds(30);

            try
            {
                await Task.Delay(timeSpan);
            }
            catch (TaskCanceledException ex)
            {
                SysUtils.ReportErrorToEventLog("TaskCanceledException error: " + ex);
            }
            catch (Exception ex)
            {
                SysUtils.ReportErrorToEventLog("Exception error: " + ex);
            }
        }
              
        [StateAction("WFS.ValidacionesDocumentosChronoscan.EnviarAChronoscan", Class = "CL.CargaMasivaSinClasificar")]
        public void EnviarAChronoscan_DocumentosCargaMasivaSinClasificar(StateEnvironment env)
        {
            var workflow = PermanentVault
                .WorkflowOperations
                .GetWorkflowIDByAlias("WF.ValidacionesDocumentosChronoscan");

            var estado = PermanentVault
                .WorkflowOperations
                .GetWorkflowStateIDByAlias("WFS.ValidacionesDocumentosChronoscan.Enviado");

            var pd_Proveedor = PermanentVault.PropertyDefOperations.GetPropertyDefIDByAlias("MF.PD.Proveedor");
            var pd_RfcEmpresa = PermanentVault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.RfcEmpresa");
            var pd_EstadoProcesamiento = PermanentVault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.EstadoDeProcesamiento");
            var pd_TipoProveedor = PermanentVault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.TipoDeProveedor");            
            string sDirectorioChronoscan = "";
            int iEstadoProcesamiento = 0;

            // Files that we should clean up.
            var filesToDelete = new List<string>();            

            try
            {
                //await PutTaskDelay(); // Delay al exportar el documento a Chronoscan

                if (env.ObjVerEx.IsEnteringState)
                {
                    var oPropertyValues = new PropertyValues();

                    oPropertyValues = env.Vault.ObjectPropertyOperations.GetProperties(env.ObjVer);

                    var oLookupsProveedor = oPropertyValues.SearchForPropertyEx(pd_Proveedor, true).TypedValue.GetValueAsLookups();

                    var oObjVerExProveedor = oLookupsProveedor.ToObjVerExs(env.Vault);

                    oPropertyValues = env.Vault.ObjectPropertyOperations.GetProperties(oObjVerExProveedor[0].ObjVer);

                    var sRFCEmpresaValue = oPropertyValues.SearchForPropertyEx(pd_RfcEmpresa, true).TypedValue.GetValueAsLocalizedText();

                    var iIdTipoProveedor = oPropertyValues.SearchForPropertyEx(pd_TipoProveedor, true).TypedValue.GetLookupID();

                    if (iIdTipoProveedor == 1) // Persona Fisica
                    {
                        sDirectorioChronoscan = Configuration.ConfiguracionExportacionAChronoscan.DirectorioPersonaFisica;
                    }
                    else if (iIdTipoProveedor == 2) // Persona Moral
                    {
                        sDirectorioChronoscan = Configuration.ConfiguracionExportacionAChronoscan.DirectorioPersonaMoral;
                    }
                    else
                    {
                        string sMensajeError = "No fue posible determinar el tipo de persona del proveedor: " + oObjVerExProveedor[0].Title;

                        SysUtils.ReportErrorToEventLog(sMensajeError);
                        throw new Exception(sMensajeError);
                    }

                    // Validar carpeta chronoscan, si no existe se crea
                    if (!Directory.Exists(sDirectorioChronoscan))
                    {
                        Directory.CreateDirectory(sDirectorioChronoscan);
                    }

                    var oObjVerEx = env.ObjVerEx;
                    var oObjectFiles = oObjVerEx.Info.Files;
                    IEnumerator enumerator = oObjectFiles.GetEnumerator();

                    while (enumerator.MoveNext())
                    {
                        ObjectFile oFile = (ObjectFile)enumerator.Current;

                        var iObjectID = env.ObjVerEx.ID; //oFile.ID;

                        var sObjectGUID = oObjVerEx.Info.ObjectGUID; //oFile.FileGUID;

                        //var sDocumentoGUIDValue = sObjectGUID.Substring(1, sObjectGUID.LastIndexOf("}") - 1);

                        string sFilePath = SysUtils.GetTempFileName(".tmp");

                        // This must be generated from the temporary path and GetTempFileName. 
                        // It cannot contain the original file name.
                        filesToDelete.Add(sFilePath);

                        // Gets the latest version of the specified file
                        FileVer fileVer = oFile.FileVer;

                        // Download the file to a temporary location
                        env.Vault.ObjectFileOperations.DownloadFile(oFile.ID, fileVer.Version, sFilePath);

                        var sFileName = oFile.GetNameForFileSystem();

                        var sDelimitador = ".";

                        int iIndex = sFileName.LastIndexOf(sDelimitador);

                        //var sClassNameOrTitle = sFileName.Substring(0, iIndex);

                        var sExtension = sFileName.Substring(iIndex + 1);

                        // Directorio por RFC
                        string sFilePathByRFC = Path.Combine(sDirectorioChronoscan, sRFCEmpresaValue);

                        // Se crea el directorio por RFC si aun no existe 
                        if (!Directory.Exists(sFilePathByRFC))
                        {
                            Directory.CreateDirectory(sFilePathByRFC);
                        }

                        // Nombre concatenado para el archivo
                        var sFileNameConcatenado = iObjectID + " - " + sObjectGUID + "." + sExtension;

                        //SysUtils.ReportInfoToEventLog("Nombre de archivo exportado: " + sFileNameConcatenado);

                        // Directorio completo del documento
                        string sNewFilePath = Path.Combine(sFilePathByRFC, sFileNameConcatenado);

                        // Copiar el documento en el nuevo directorio
                        File.Copy(sFilePath, sNewFilePath);


                        if (File.Exists(sNewFilePath))
                        {
                            // Documento Procesado con exito
                            iEstadoProcesamiento = 1;
                        }
                        else
                        {
                            // No Procesado o Termino en Error
                            iEstadoProcesamiento = 2;
                        }

                        // Actualizacion de estado del workflow
                        var oWorkflowState = new ObjectVersionWorkflowState();
                        var oObjVer = env.Vault.ObjectOperations.GetLatestObjVerEx(env.ObjVerEx.ObjID, true);

                        oWorkflowState.Workflow.TypedValue.SetValue(MFDataType.MFDatatypeLookup, workflow);
                        oWorkflowState.State.TypedValue.SetValue(MFDataType.MFDatatypeLookup, estado);
                        env.Vault.ObjectPropertyOperations.SetWorkflowStateEx(oObjVer, oWorkflowState);

                        // Actualizacion de estado de procesamiento
                        var oLookup = new Lookup();
                        var oObjID = new ObjID();

                        oObjID.SetIDs
                        (
                            ObjType: (int)MFBuiltInObjectType.MFBuiltInObjectTypeDocument,
                            ID: env.ObjVer.ID
                        );

                        var oPropertyValue = new PropertyValue
                        {
                            PropertyDef = pd_EstadoProcesamiento
                        };

                        oLookup.Item = iEstadoProcesamiento;

                        oPropertyValue.TypedValue.SetValueToLookup(oLookup);

                        env.Vault.ObjectPropertyOperations.SetProperty
                        (
                            ObjVer: env.ObjVer,
                            PropertyValue: oPropertyValue
                        );

                        SysUtils.ReportInfoToEventLog("Exportacion completada. Documento: " + sFileNameConcatenado + ", Rfc: " + sRFCEmpresaValue);
                    }
                }                
            }
            catch (Exception ex)
            {
                SysUtils.ReportErrorToEventLog("Error al exportar el documento, ", ex);
            }
            finally
            {
                // Always clean up the files (whether it works or not).
                foreach (var sFile in filesToDelete)
                {
                    File.Delete(sFile);
                }
            }
        }

        [PropertyCustomValue("PD.HubGuid")]
        public TypedValue CalculatingHubGuidValue(PropertyEnvironment env)
        {
            var pd_Hubsharelink = PermanentVault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.Hubsharelink");
            var pd_HubGUID = PermanentVault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.HubGuid");

            var oPropertyValues = new PropertyValues();
            var oTypedValue = new TypedValue();

            oPropertyValues = env.Vault.ObjectPropertyOperations.GetProperties(env.ObjVer);

            if (oPropertyValues.IndexOf(pd_Hubsharelink) != -1)
            {
                var HubsharelinkValue = oPropertyValues
                    .SearchForPropertyEx(pd_Hubsharelink, true)
                    .TypedValue
                    .GetValueAsLocalizedText();

                // https://demo-usa.hubshare.com/#/Hub/29670d45-9172-4d95-955c-21d2f285e53c

                var delimitador = "/";

                int index = HubsharelinkValue.LastIndexOf(delimitador);

                var HubGuidValue = HubsharelinkValue.Substring(index + 1);

                if (oPropertyValues.IndexOf(pd_HubGUID) != -1)
                {
                    oTypedValue.SetValue(MFDataType.MFDatatypeText, HubGuidValue);
                }
            }

            return oTypedValue;
        }

        private void CreateNewObjectCreacionMasivaDeProveedores(EnvironmentBase env, string sFilePath)
        {
            var cl_CreacionMasivaDeProveedores = env.Vault.ClassOperations.GetObjectClassIDByAlias("CL.CreacionMasivaDeProveedores");
            //var pd_CrearListaProveedores = env.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.CrearListadoDeProveedores");

            string sFileName = Path.GetFileNameWithoutExtension(sFilePath);

            var createBuilder = new MFPropertyValuesBuilder(env.Vault);
            createBuilder.SetClass(cl_CreacionMasivaDeProveedores); // Clase issue
            createBuilder.Add
            (
                (int)MFBuiltInPropertyDef.MFBuiltInPropertyDefNameOrTitle,
                MFDataType.MFDatatypeText,
                sFileName // Name or title
            );
            //createBuilder.Add(pd_CrearListaProveedores, MFDataType.MFDatatypeBoolean, false);

            var oSourceObjetctFiles = new SourceObjectFiles();

            var oObjFile = new SourceObjectFile
            {
                SourceFilePath = sFilePath,
                Title = sFileName,
                Extension = "xls"
            };
            oSourceObjetctFiles.Add(-1, oObjFile);

            var objectTypeID = (int)MFBuiltInObjectType.MFBuiltInObjectTypeDocument;

            // Validate if the document is single-file or multi-file
            var isSingleFileDocument =
                objectTypeID == (int)MFBuiltInObjectType.MFBuiltInObjectTypeDocument &&
                oSourceObjetctFiles.Count == 1;

            // Create the new object and check it in.
            var objectVersion = env.Vault.ObjectOperations.CreateNewObjectEx
            (
                objectTypeID,
                createBuilder.Values,
                oSourceObjetctFiles,
                SFD: isSingleFileDocument,
                CheckIn: true
            );
        }

        private void UpdateEstatusDeFirmaFlujoExcepcionesProveedor(EnvironmentBase env)
        {
            bool bFirmaEsValida = false;
            var ot_Proveedor = env.Vault.ObjectTypeOperations.GetObjectTypeIDByAlias("OT.Proveedor");
            var cl_ProyectoServicioEspecializado = env.Vault.ClassOperations.GetObjectClassIDByAlias("CL.ProyectoServicioEspecializado");
            var cl_Contrato = env.Vault.ClassOperations.GetObjectClassIDByAlias("CL.Contrato");
            var pd_Proveedor = env.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("MF.PD.Proveedor");
            var pd_EstatusProyecto = env.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.EstatusProyecto");
            var pd_EstatusContrato = env.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("MF.PD.ContractStatus");
            var pd_EstatusFlujoExcepciones = env.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.EstatusFlujoDeExcepciones");

            var oPropertyValues = new PropertyValues();
            oPropertyValues = env.ObjVerEx.Properties;

            if (oPropertyValues.IndexOf(pd_Proveedor) != -1)
            {
                var oLookupsProveedor = oPropertyValues.SearchForPropertyEx(pd_Proveedor, true).TypedValue.GetValueAsLookups();

                // Crear arreglo de clases
                var arrClasesId = new[] { cl_Contrato, cl_ProyectoServicioEspecializado };

                var searchBuilder = new MFSearchBuilder(env.Vault);
                searchBuilder.Deleted(false);
                searchBuilder.Property
                (
                    (int)MFBuiltInPropertyDef.MFBuiltInPropertyDefClass,
                    MFDataType.MFDatatypeMultiSelectLookup,
                    arrClasesId
                );
                searchBuilder.Property(pd_Proveedor, MFDataType.MFDatatypeMultiSelectLookup, oLookupsProveedor);

                var searchResults = searchBuilder.FindEx();

                if (searchResults.Count > 0)
                {
                    foreach (var result in searchResults)
                    {
                        var oPropertyValuesResult = new PropertyValues();
                        oPropertyValuesResult = result.Properties;

                        if (oPropertyValuesResult.IndexOf(pd_EstatusContrato) != -1)
                        {
                            var iEstatusContrato = oPropertyValuesResult.SearchForPropertyEx(pd_EstatusContrato, true).TypedValue.GetLookupID();

                            if (iEstatusContrato == 2 || iEstatusContrato == 4)
                            {
                                bFirmaEsValida = true;
                            }
                        }

                        if (oPropertyValuesResult.IndexOf(pd_EstatusProyecto) != -1)
                        {
                            var iEstatusProyecto = oPropertyValuesResult.SearchForPropertyEx(pd_EstatusProyecto, true).TypedValue.GetLookupID();

                            if (iEstatusProyecto == 1 || iEstatusProyecto == 5)
                            {
                                bFirmaEsValida = true;
                            }
                        }
                    }

                    if (bFirmaEsValida == false)
                    {
                        var oListProveedor = oLookupsProveedor.ToObjVerExs(env.Vault);
                        var objVerProveedor = oListProveedor[0];

                        // Asignar propiedad estatus del flujo de excepciones en la metadata del proveedor
                        var oLookup = new Lookup();
                        var oObjID = new ObjID();

                        oObjID.SetIDs
                        (
                            ObjType: ot_Proveedor,
                            ID: objVerProveedor.ObjVer.ID
                        );

                        var checkedOutObjectVersion = env.Vault.ObjectOperations.CheckOut(oObjID);

                        var oPropertyValue = new PropertyValue
                        {
                            PropertyDef = pd_EstatusFlujoExcepciones
                        };

                        oLookup.Item = 4;

                        oPropertyValue.TypedValue.SetValueToLookup(oLookup);

                        env.Vault.ObjectPropertyOperations.SetProperty
                        (
                            ObjVer: checkedOutObjectVersion.ObjVer,
                            PropertyValue: oPropertyValue
                        );

                        env.Vault.ObjectOperations.CheckIn(checkedOutObjectVersion.ObjVer);
                    }
                }
            }
        }

        private bool SetPropertiesInProveedorAndContactoExterno(EnvironmentBase env, string sRfcProveedor)
        {
            bool bResult = false;
            var ot_ContactoExterno = env.Vault.ObjectTypeOperations.GetObjectTypeIDByAlias("MF.OT.ExternalContact");
            var ot_Proveedor = env.Vault.ObjectTypeOperations.GetObjectTypeIDByAlias("OT.Proveedor");
            var pd_RfcEmpresa = env.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.RfcEmpresa");
            var pd_ContactosExternosAdministradores = env.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.ContactosExternosAdministradores");
            var pd_Proveedor = env.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("MF.PD.Proveedor");

            var oPropertyValueProveedor = new PropertyValue();
            var oPropertyValueContacto = new PropertyValue();
            var oLookupsProveedor = new Lookups();
            var oLookupProveedor = new Lookup();
            var oLookupsContacto = new Lookups();
            var oLookupContacto = new Lookup();

            // Obtener el ObjVer del Proveedor
            var searchBuilderProveedor = new MFSearchBuilder(env.Vault);
            searchBuilderProveedor.Deleted(false);
            searchBuilderProveedor.ObjType(ot_Proveedor);
            searchBuilderProveedor.Property(pd_RfcEmpresa, MFDataType.MFDatatypeText, sRfcProveedor);

            var searchResultsProveedor = searchBuilderProveedor.FindEx();

            if (searchResultsProveedor.Count > 0)
            {
                var oProveedor = searchResultsProveedor[0].ObjVer;

                oLookupProveedor.Item = oProveedor.ID;
                oLookupsProveedor.Add(-1, oLookupProveedor);

                // Obtener los contactos externos del proveedor
                var searchBuilderContacto = new MFSearchBuilder(env.Vault);
                searchBuilderContacto.Deleted(false);
                searchBuilderContacto.ObjType(ot_ContactoExterno);
                searchBuilderContacto.Property(pd_RfcEmpresa, MFDataType.MFDatatypeText, sRfcProveedor);

                var searchResultsContacto = searchBuilderContacto.FindEx();

                if (searchResultsContacto.Count > 0)
                {
                    foreach (var contacto in searchResultsContacto)
                    {
                        oLookupContacto.Item = contacto.ObjVer.ID;
                        oLookupsContacto.Add(-1, oLookupContacto);

                        // Relacionar el proveedor en el contacto externo                        
                        var oObjVerContacto = env.Vault.ObjectOperations.GetLatestObjVerEx(contacto.ObjID, true);
                        oPropertyValueContacto.PropertyDef = pd_Proveedor;
                        oPropertyValueContacto.TypedValue.SetValueToMultiSelectLookup(oLookupsProveedor);
                        oObjVerContacto = env.Vault.ObjectOperations.CheckOut(contacto.ObjID).ObjVer;
                        env.Vault.ObjectPropertyOperations.SetProperty(oObjVerContacto, oPropertyValueContacto);
                        env.Vault.ObjectOperations.CheckIn(oObjVerContacto);
                    }

                    // Relacionar el o los contactos externos en el proveedor
                    var oObjVerProveedor = env.Vault.ObjectOperations.GetLatestObjVerEx(oProveedor.ObjID, true);
                    oPropertyValueProveedor.PropertyDef = pd_ContactosExternosAdministradores;
                    oPropertyValueProveedor.TypedValue.SetValueToMultiSelectLookup(oLookupsContacto);
                    oObjVerProveedor = env.Vault.ObjectOperations.CheckOut(oProveedor.ObjID).ObjVer;
                    env.Vault.ObjectPropertyOperations.SetProperty(oObjVerProveedor, oPropertyValueProveedor);
                    env.Vault.ObjectOperations.CheckIn(oObjVerProveedor);

                    bResult = true;
                }
            }

            return bResult;
        }

        private bool CreateOrUpdateContactoExternoAdmin(EnvironmentBase env, string sRfcProveedor, string sNombre, string sApellidoPaterno, string sApellidoMaterno, string sCurp, string sEmail)
        {
            bool bResult = false;

            var ot_ContactoExterno = env.Vault.ObjectTypeOperations.GetObjectTypeIDByAlias("MF.OT.ExternalContact");
            var cl_ContactoExterno = env.Vault.ClassOperations.GetObjectClassIDByAlias("CL.ContactoExternoServicioEspecializado");
            var cl_ProveedorServicioEspecializado = env.Vault.ClassOperations.GetObjectClassIDByAlias("CL.ProveedorDeServicioEspecializado");
            var pd_RfcEmpresa = env.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.RfcEmpresa");
            var pd_Curp = env.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.Curp");
            var pd_Nombre = env.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("MF.PD.FirstName");
            var pd_ApellidoPaterno = env.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("MF.PD.LastName.Paterno");
            var pd_ApellidoMaterno = env.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.ApellidoMaterno");
            var pd_Email = env.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("MF.PD.EmailAddress");
            var pd_EsAdministradorHubshare = env.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.EsAdministrador");

            var oPropertyValues = new PropertyValues();
            //var oLookupsProveedor = new Lookups();
            //var oLookupProveedor = new Lookup();
            var oObjID = new ObjID();

            // Validar si el contacto externo administrador ya existe en la boveda
            var searchBuilderContacto = new MFSearchBuilder(PermanentVault);
            searchBuilderContacto.Deleted(false);
            searchBuilderContacto.Class(cl_ContactoExterno);
            searchBuilderContacto.Property(pd_Curp, MFDataType.MFDatatypeText, sCurp);

            var searchResultsContacto = searchBuilderContacto.FindEx();

            if (searchResultsContacto.Count > 0)
            {
                // Si el contacto ya existe, actualizar
                var objVerContacto = searchResultsContacto[0].ObjVer;

                oObjID.SetIDs
                (
                    ObjType: ot_ContactoExterno,
                    ID: objVerContacto.ID
                );

                var checkedOutObjectVersion = env.Vault.ObjectOperations.CheckOut(oObjID);

                var propValNombre = new PropertyValue
                {
                    PropertyDef = pd_Nombre
                };
                propValNombre.TypedValue.SetValue(MFDataType.MFDatatypeText, sNombre);
                oPropertyValues.Add(-1, propValNombre);

                var propValApellidoPaterno = new PropertyValue
                {
                    PropertyDef = pd_ApellidoPaterno
                };
                propValApellidoPaterno.TypedValue.SetValue(MFDataType.MFDatatypeText, sApellidoPaterno);
                oPropertyValues.Add(-1, propValApellidoPaterno);

                var propValApellidoMaterno = new PropertyValue
                {
                    PropertyDef = pd_ApellidoMaterno
                };
                propValApellidoMaterno.TypedValue.SetValue(MFDataType.MFDatatypeText, sApellidoMaterno);
                oPropertyValues.Add(-1, propValApellidoMaterno);

                var propValEmail = new PropertyValue
                {
                    PropertyDef = pd_Email
                };
                propValEmail.TypedValue.SetValue(MFDataType.MFDatatypeText, sEmail);
                oPropertyValues.Add(-1, propValEmail);

                var propValRfcEmpresa = new PropertyValue
                {
                    PropertyDef = pd_RfcEmpresa
                };
                propValRfcEmpresa.TypedValue.SetValue(MFDataType.MFDatatypeText, sRfcProveedor);
                oPropertyValues.Add(-1, propValRfcEmpresa);

                env.Vault.ObjectPropertyOperations.SetPropertiesEx
                (
                    checkedOutObjectVersion.ObjVer,                    
                    oPropertyValues,
                    true
                );

                env.Vault.ObjectOperations.CheckIn(checkedOutObjectVersion.ObjVer);

                bResult = true;
            }
            else
            {
                // Si el contacto aun no existe, crear
                var createBuilderContacto = new MFPropertyValuesBuilder(PermanentVault);
                createBuilderContacto.SetClass(cl_ContactoExterno);
                createBuilderContacto.Add(pd_RfcEmpresa, MFDataType.MFDatatypeText, sRfcProveedor);
                createBuilderContacto.Add(pd_Curp, MFDataType.MFDatatypeText, sCurp);
                createBuilderContacto.Add(pd_Nombre, MFDataType.MFDatatypeText, sNombre);
                createBuilderContacto.Add(pd_ApellidoPaterno, MFDataType.MFDatatypeText, sApellidoPaterno);
                createBuilderContacto.Add(pd_ApellidoMaterno, MFDataType.MFDatatypeText, sApellidoMaterno);
                createBuilderContacto.Add(pd_Email, MFDataType.MFDatatypeText, sEmail);
                createBuilderContacto.Add(pd_EsAdministradorHubshare, MFDataType.MFDatatypeBoolean, true);
                //createBuilderContacto.Add(pd_Proveedor, MFDataType.MFDatatypeMultiSelectLookup, oLookupsProveedor);

                // Tipo de objeto a crear
                var objectTypeId = ot_ContactoExterno;

                var objectVersion = env.Vault.ObjectOperations.CreateNewObjectEx
                (
                    objectTypeId,
                    createBuilderContacto.Values,
                    CheckIn: true
                );

                bResult = true;
            }

            return bResult;
        }

        private bool CreateOrUpdateProveedor(EnvironmentBase env, 
            string sRfcProvedor, 
            string sNombreProveedor, 
            string sTipoProveedor, 
            string sTipoValidacionChecklist, 
            string sTipoPersona, 
            DateTime dtFechaInicioProveedor)
        {
            bool bResult = false;
            int iNuevaClaseAEstablecer = 0;
            int iTipoProveedor = 0;
            int iValidacionServiciosEspecializados = 0;

            var ot_Proveedor = env.Vault.ObjectTypeOperations.GetObjectTypeIDByAlias("OT.Proveedor");
            var cl_Proveedor = env.Vault.ClassOperations.GetObjectClassIDByAlias("CL.Proveedor");
            var cl_ProveedorServicioEspecializado = env.Vault.ClassOperations.GetObjectClassIDByAlias("CL.ProveedorDeServicioEspecializado");
            var cl_ProveedorDependencia = env.Vault.ClassOperations.GetObjectClassIDByAlias("CL.ProveedorDependencia");
            var cl_ProveedorEstrategico = env.Vault.ClassOperations.GetObjectClassIDByAlias("CL.ProveedorEstratgico");
            var cl_ProveedorExtranjero = env.Vault.ClassOperations.GetObjectClassIDByAlias("CL.ProveedorExtranjero");
            var cl_ProveedorTransportista = env.Vault.ClassOperations.GetObjectClassIDByAlias("CL.ProveedorTransportista");
            var cl_HubshareTemplate = env.Vault.ClassOperations.GetObjectClassIDByAlias("CL.HubshareTemplate");
            var pd_RfcEmpresa = env.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.RfcEmpresa");
            var pd_HubshareTemplate = env.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.HubshareTemplate");
            var pd_CrearHub = env.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.CrearHub");
            var pd_UsarPlantillaHubshare = env.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.UsarPlantillaDeHubshare");
            var pd_ContactosExternosAdministradores = env.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.ContactosExternosAdministradores");
            var pd_Severidad = env.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("MF.PD.Severity");
            var pd_ValidacionServiciosEspecializados = env.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.TipoDeValidacionLeyDeOutsourcing");
            var pd_TipoProveedor = env.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.TipoDeProveedor");
            var pd_FechaInicio = env.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("MF.PD.StartDate");

            var oPropertyValues = new PropertyValues();
            var oLookups = new Lookups();
            var oLookup = new Lookup();
            var oObjID = new ObjID();

            // Obtener el Hubshare Template correspondiente al tipo de proveedor
            var searchBuilderHubshareTemplate = new MFSearchBuilder(env.Vault);
            searchBuilderHubshareTemplate.Deleted(false);
            searchBuilderHubshareTemplate.Class(cl_HubshareTemplate);
            searchBuilderHubshareTemplate.Property
            (
                MFBuiltInPropertyDef.MFBuiltInPropertyDefNameOrTitle, 
                MFDataType.MFDatatypeText, 
                sTipoProveedor
            );

            var searchResultsHubshareTemplate = searchBuilderHubshareTemplate.FindEx();

            if (searchResultsHubshareTemplate.Count > 0)
            {
                var objVerHubshareTemplate = searchResultsHubshareTemplate[0].ObjVer;

                oLookup.Item = objVerHubshareTemplate.ID;
                oLookups.Add(-1, oLookup);
            }

            // Validar la clase del tipo de proveedor
            if (sTipoProveedor == "Proveedor de Servicios Especializados")
                iNuevaClaseAEstablecer = cl_ProveedorServicioEspecializado;
            else if (sTipoProveedor == "Proveedor Extranjero")
                iNuevaClaseAEstablecer = cl_ProveedorExtranjero;
            else if (sTipoProveedor == "Proveedor Dependencia")
                iNuevaClaseAEstablecer = cl_ProveedorDependencia;
            else if (sTipoProveedor == "Proveedor Estratégico")
                iNuevaClaseAEstablecer = cl_ProveedorEstrategico;
            else if (sTipoProveedor == "Proveedor Transportista")
                iNuevaClaseAEstablecer = cl_ProveedorTransportista;
            else //if (sTipoProveedor == "Proveedor General")
                iNuevaClaseAEstablecer = cl_Proveedor;

            // Validar el tipo de persona
            if (sTipoPersona == "Persona Moral")
                iTipoProveedor = 2;
            else
                iTipoProveedor = 1;

            // Tipo de validacion del checklist
            if (sTipoValidacionChecklist == "Por Proveedor")
                iValidacionServiciosEspecializados = 1;
            else if (sTipoValidacionChecklist == "Orden de Compra, Contrato y/o Proyecto")
                iValidacionServiciosEspecializados = 2;
            else
                iValidacionServiciosEspecializados = 3; // Por Empresa Interna 

            //// Conversion de fecha de inicio del proveedor
            //DateTime dtFechaInicio = DateTime.FromOADate(sFechaInicioProveedor);

            // Buscar Rfc en Proveedor
            var searchBuilderRfcProveedor = new MFSearchBuilder(env.Vault);
            searchBuilderRfcProveedor.Deleted(false); // No eliminados
            searchBuilderRfcProveedor.ObjType(ot_Proveedor);
            searchBuilderRfcProveedor.Property(pd_RfcEmpresa, MFDataType.MFDatatypeText, sRfcProvedor);            

            var searchResultsRfcProveedor = searchBuilderRfcProveedor.FindEx();

            if (searchResultsRfcProveedor.Count > 0)
            {
                // Si el proveedor ya existe, actualizar
                var oObjVerProveedor = searchResultsRfcProveedor[0];

                // Obtener las propiedades del proveedor
                var proveedorProperties = oObjVerProveedor.Properties;

                // Validar si la clase cambia, tomando como referencia el tipo de proveedor
                var iClaseActualDeProveedor = proveedorProperties
                    .SearchForPropertyEx((int)MFBuiltInPropertyDef.MFBuiltInPropertyDefClass, true)
                    .TypedValue
                    .GetLookupID();

                if (iClaseActualDeProveedor != iNuevaClaseAEstablecer)
                {
                    // Extraer las propiedades del proveedor que no fueron modificadas
                    var sRfcEmpresa = proveedorProperties.SearchForPropertyEx(pd_RfcEmpresa, true).TypedValue.Value.ToString();
                    var lookupsContactosExternosAdmin = proveedorProperties.SearchForPropertyEx(pd_ContactosExternosAdministradores, true).TypedValue.GetValueAsLookups();
                    var iSeveridad = proveedorProperties.SearchForPropertyEx(pd_Severidad, true).TypedValue.GetLookupID();

                    // Establecer el objetos y las propiedades a actualizar
                    oObjID.SetIDs
                    (
                        ObjType: ot_Proveedor,
                        ID: oObjVerProveedor.ObjVer.ID
                    );

                    var checkedOutObjectVersion = env.Vault.ObjectOperations.CheckOut(oObjID);

                    var propValClase = new PropertyValue
                    {
                        PropertyDef = (int)MFBuiltInPropertyDef.MFBuiltInPropertyDefClass
                    };
                    propValClase.TypedValue.SetValue(MFDataType.MFDatatypeLookup, iNuevaClaseAEstablecer); // Clase
                    oPropertyValues.Add(-1, propValClase);

                    var propValNombreProveedor = new PropertyValue
                    {
                        PropertyDef = (int)MFBuiltInPropertyDef.MFBuiltInPropertyDefNameOrTitle
                    };
                    propValNombreProveedor.TypedValue.SetValue(MFDataType.MFDatatypeText, sNombreProveedor); // Nombre o titulo
                    oPropertyValues.Add(-1, propValNombreProveedor);

                    var propValRfcEmpresa = new PropertyValue
                    {
                        PropertyDef = pd_RfcEmpresa
                    };
                    propValRfcEmpresa.TypedValue.SetValue(MFDataType.MFDatatypeText, sRfcEmpresa); // Rfc empresa
                    oPropertyValues.Add(-1, propValRfcEmpresa);

                    if (lookupsContactosExternosAdmin.Count > 0)
                    {
                        var propValContactosExternosAdmin = new PropertyValue
                        {
                            PropertyDef = pd_ContactosExternosAdministradores
                        };
                        propValContactosExternosAdmin.TypedValue.SetValueToMultiSelectLookup(lookupsContactosExternosAdmin); // Contactos externos administradores
                        oPropertyValues.Add(-1, propValContactosExternosAdmin);
                    }                    

                    if (iSeveridad > 0)
                    {
                        var propValSeveridad = new PropertyValue
                        {
                            PropertyDef = pd_Severidad
                        };
                        propValSeveridad.TypedValue.SetValue(MFDataType.MFDatatypeLookup, iSeveridad); // Severidad
                        oPropertyValues.Add(-1, propValSeveridad);
                    }

                    var propValTipoProveedor = new PropertyValue
                    {
                        PropertyDef = pd_TipoProveedor
                    };
                    propValTipoProveedor.TypedValue.SetValue(MFDataType.MFDatatypeLookup, iTipoProveedor); // Tipo de Proveedor
                    oPropertyValues.Add(-1, propValTipoProveedor);

                    var propValValidacionServiciosEspecializados = new PropertyValue
                    {
                        PropertyDef = pd_ValidacionServiciosEspecializados
                    };
                    propValValidacionServiciosEspecializados.TypedValue.SetValue(MFDataType.MFDatatypeLookup, iValidacionServiciosEspecializados); // Validacion Servicios Especializados
                    oPropertyValues.Add(-1, propValValidacionServiciosEspecializados);

                    var propValFechaInicio = new PropertyValue
                    {
                        PropertyDef = pd_FechaInicio
                    };
                    propValFechaInicio.TypedValue.SetValue(MFDataType.MFDatatypeDate, dtFechaInicioProveedor); //dtFechaInicio
                    oPropertyValues.Add(-1, propValFechaInicio);

                    var propValCrearHub = new PropertyValue
                    {
                        PropertyDef = pd_CrearHub
                    };
                    propValCrearHub.TypedValue.SetValue(MFDataType.MFDatatypeBoolean, true); // Crear hub
                    oPropertyValues.Add(-1, propValCrearHub);

                    var propValUsarPlantillaHubshare = new PropertyValue
                    {
                        PropertyDef = pd_UsarPlantillaHubshare
                    };
                    propValUsarPlantillaHubshare.TypedValue.SetValue(MFDataType.MFDatatypeBoolean, true); // Usar plantilla hubshare
                    oPropertyValues.Add(-1, propValUsarPlantillaHubshare);

                    var propValHubshareTemplate = new PropertyValue
                    {
                        PropertyDef = pd_HubshareTemplate
                    };
                    propValHubshareTemplate.TypedValue.SetValueToMultiSelectLookup(oLookups); // Hubshare template
                    oPropertyValues.Add(-1, propValHubshareTemplate);

                    var propValSingleFile = new PropertyValue
                    {
                        PropertyDef = (int)MFBuiltInPropertyDef.MFBuiltInPropertyDefSingleFileObject
                    };
                    propValSingleFile.TypedValue.SetValue(MFDataType.MFDatatypeBoolean, false); // Single file object
                    oPropertyValues.Add(-1, propValSingleFile);

                    env.Vault.ObjectPropertyOperations.SetAllProperties
                    (
                        checkedOutObjectVersion.ObjVer,
                        true,
                        oPropertyValues
                    );

                    env.Vault.ObjectOperations.CheckIn(checkedOutObjectVersion.ObjVer);

                    bResult = true;
                }
                else
                {
                    oObjID.SetIDs
                    (
                        ObjType: ot_Proveedor,
                        ID: oObjVerProveedor.ObjVer.ID
                    );

                    var checkedOutObjectVersion = env.Vault.ObjectOperations.CheckOut(oObjID);

                    // Actualizar el nombre del proveedor asociado a Rfc Empresa
                    var propValNombreProveedor = new PropertyValue
                    {
                        PropertyDef = (int)MFBuiltInPropertyDef.MFBuiltInPropertyDefNameOrTitle
                    };
                    propValNombreProveedor.TypedValue.SetValue(MFDataType.MFDatatypeText, sNombreProveedor); // Nombre o titulo

                    env.Vault.ObjectPropertyOperations.SetProperty
                    (
                        ObjVer: checkedOutObjectVersion.ObjVer,
                        PropertyValue: propValNombreProveedor
                    );

                    env.Vault.ObjectOperations.CheckIn(checkedOutObjectVersion.ObjVer);

                    bResult = true;
                }
            }
            else
            {
                // Si el proveedor aun no existe, crear
                var createBuilder = new MFPropertyValuesBuilder(env.Vault);
                createBuilder.SetClass(iNuevaClaseAEstablecer);
                createBuilder.Add
                (
                    (int)MFBuiltInPropertyDef.MFBuiltInPropertyDefNameOrTitle,
                    MFDataType.MFDatatypeText,
                    sNombreProveedor // Name or title
                );
                createBuilder.Add(pd_RfcEmpresa, MFDataType.MFDatatypeText, sRfcProvedor); // Rfc
                createBuilder.Add(pd_CrearHub, MFDataType.MFDatatypeBoolean, true);
                createBuilder.Add(pd_UsarPlantillaHubshare, MFDataType.MFDatatypeBoolean, true);
                createBuilder.Add(pd_HubshareTemplate, MFDataType.MFDatatypeMultiSelectLookup, oLookups);
                createBuilder.Add(pd_ValidacionServiciosEspecializados, MFDataType.MFDatatypeLookup, iValidacionServiciosEspecializados);
                createBuilder.Add(pd_TipoProveedor, MFDataType.MFDatatypeLookup, iTipoProveedor);
                createBuilder.Add(pd_FechaInicio, MFDataType.MFDatatypeDate, dtFechaInicioProveedor); //dtFechaInicio

                // Tipo de objeto a crear
                var objectTypeId = ot_Proveedor;

                var objectVersion = env.Vault.ObjectOperations.CreateNewObjectEx
                (
                    objectTypeId,
                    createBuilder.Values,
                    CheckIn: true
                );

                bResult = true;
            }        

            return bResult;
        }

        private void SetPropertiesGenerico(EnvironmentBase env, int iClaseBusqueda, int iPropiedadBusqueda, int iPropiedadRelacion, Lookups oValorBusqueda, bool bSeRequiereBuscarProyectosRelacionados = true, bool bValidarTipoDeContratoBusqueda = false)
        {
            var pd_EsConvenioModificatorio = env.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.EsConvenioModificatorio");
            var pd_ProyectosRelacionados = env.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("MF.PD.Project");

            var oPropertyValue = new PropertyValue();
            var oLookups = new Lookups();
            var oLookup = new Lookup();            

            var searchBuilder = new MFSearchBuilder(env.Vault);
            searchBuilder.Deleted(false); // No eliminados
            searchBuilder.Class(iClaseBusqueda);
            searchBuilder.Property(iPropiedadBusqueda, MFDataType.MFDatatypeMultiSelectLookup, oValorBusqueda);

            if (bSeRequiereBuscarProyectosRelacionados == true)
            {
                var oProyectosProperties = env.ObjVerEx.Properties;
                var oProyectosRelacionados = oProyectosProperties.SearchForPropertyEx(pd_ProyectosRelacionados, true).TypedValue.GetValueAsLookups();
                searchBuilder.Property(pd_ProyectosRelacionados, MFDataType.MFDatatypeMultiSelectLookup, oProyectosRelacionados);
            }                

            if (bValidarTipoDeContratoBusqueda == true)
                searchBuilder.Property(pd_EsConvenioModificatorio, MFDataType.MFDatatypeBoolean, true);

            var searchResults = searchBuilder.FindEx();

            if (searchResults.Count > 0)
            {
                foreach (var result in searchResults)
                {
                    oLookup.Item = result.ObjVer.ID;
                    oLookups.Add(-1, oLookup);
                }

                oPropertyValue.PropertyDef = iPropiedadRelacion;
                oPropertyValue.TypedValue.SetValueToMultiSelectLookup(oLookups);
                env.Vault.ObjectPropertyOperations.SetProperty(env.ObjVer, oPropertyValue);
            }
        }

        private void UpdateEstatusCFDIComp(ObjVerEx oCfdi, int iEstatus)
        {
            var pd_EstatusCompulsaCFDI = oCfdi.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.EstatusCompulsaCFDI");
            var oPropertyValues = new PropertyValues();
            var oLookupEstatus = new Lookup();
            var oObjID = new ObjID();
            
            oPropertyValues = oCfdi.Vault.ObjectPropertyOperations.GetProperties(oCfdi.ObjVer);

            if (oPropertyValues.SearchForPropertyEx(pd_EstatusCompulsaCFDI, true).TypedValue.IsNULL() || 
                oPropertyValues.SearchForPropertyEx(pd_EstatusCompulsaCFDI, true).TypedValue.GetLookupID() == 1)
            {
                oObjID.SetIDs
                (
                    ObjType: (int)MFBuiltInObjectType.MFBuiltInObjectTypeDocument,
                    ID: oCfdi.ObjVer.ID
                );

                var checkedOutObjectVersion = oCfdi.Vault.ObjectOperations.CheckOut(oObjID);

                var oPropertyValue = new PropertyValue
                {
                    PropertyDef = pd_EstatusCompulsaCFDI
                };

                oLookupEstatus.Item = 2;

                oPropertyValue.TypedValue.SetValueToLookup(oLookupEstatus);

                oCfdi.Vault.ObjectPropertyOperations.SetProperty
                (
                    ObjVer: checkedOutObjectVersion.ObjVer,
                    PropertyValue: oPropertyValue
                );

                oCfdi.Vault.ObjectOperations.CheckIn(checkedOutObjectVersion.ObjVer);
            }
        }

        private void UpdateEstatusCFDI(ObjVerEx oCfdi, int iEstatus)
        {
            var pd_EstatusCompulsaCFDI = oCfdi.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.EstatusCompulsaCFDI");
            var oLookup = new Lookup();
            var oObjID = new ObjID();
             
            oObjID.SetIDs
            (
                ObjType: (int)MFBuiltInObjectType.MFBuiltInObjectTypeDocument,
                ID: oCfdi.ObjVer.ID
            );

            var oPropertyValue = new PropertyValue
            {
                PropertyDef = pd_EstatusCompulsaCFDI
            };

            oLookup.Item = iEstatus;

            oPropertyValue.TypedValue.SetValueToLookup(oLookup);

            oCfdi.Vault.ObjectPropertyOperations.SetProperty
            (
                ObjVer: oCfdi.ObjVer,
                PropertyValue: oPropertyValue
            );
        }

        private void SetPropertiesCFDIComprobante(EnvironmentBase env, int iClaseComprobante, int iClaseBusqueda, int iPropiedadBusqueda, int iPropiedadRelacion, Lookups oValorBusqueda)
        {
            var oPropertyValue = new PropertyValue();
            var oLookups = new Lookups();
            var oLookup = new Lookup();
            var oLookupsComprobante = new Lookups();
            var oLookupComprobante = new Lookup();

            var searchBuilder = new MFSearchBuilder(env.Vault);
            searchBuilder.Deleted(false); // No eliminados
            searchBuilder.Class(iClaseBusqueda);
            searchBuilder.Property(iPropiedadBusqueda, MFDataType.MFDatatypeMultiSelectLookup, oValorBusqueda);

            var searchResults = searchBuilder.FindEx();

            if (searchResults.Count > 0)
            {
                foreach (var result in searchResults)
                {
                    oLookup.Item = result.ObjVer.ID;
                    oLookups.Add(-1, oLookup);
                }

                oPropertyValue.PropertyDef = iPropiedadRelacion;
                oPropertyValue.TypedValue.SetValueToMultiSelectLookup(oLookups);
                env.Vault.ObjectPropertyOperations.SetProperty(env.ObjVer, oPropertyValue);
            }

            // Agregar los comprobantes en el objeto vinculado para que la relacion sea bidireccional
            var pd_DocumentosRelacionados = env.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.DocumentosRelacionados");

            var searchBuilderComprobante = new MFSearchBuilder(env.Vault);
            searchBuilderComprobante.Deleted(false); // No eliminados
            searchBuilderComprobante.Class(iClaseComprobante);
            searchBuilderComprobante.Property(iPropiedadBusqueda, MFDataType.MFDatatypeMultiSelectLookup, oValorBusqueda);

            var searchResultsComprobante = searchBuilderComprobante.FindEx();

            foreach (var resultComprobante in searchResultsComprobante)
            {
                oLookupComprobante.Item = resultComprobante.ObjVer.ID;
                oLookupsComprobante.Add(-1, oLookupComprobante);
            }

            foreach (var result in searchResults)
            {
                var oObjVer = result.Vault.ObjectOperations.GetLatestObjVerEx(result.ObjID, true);
                oPropertyValue.PropertyDef = pd_DocumentosRelacionados;
                oPropertyValue.TypedValue.SetValueToMultiSelectLookup(oLookupsComprobante);
                oObjVer = result.Vault.ObjectOperations.CheckOut(result.ObjID).ObjVer;
                result.Vault.ObjectPropertyOperations.SetProperty(oObjVer, oPropertyValue);
                result.Vault.ObjectOperations.CheckIn(oObjVer);
            }
        }

        private void SetBindingProperties(EnvironmentBase env, int iClase, int iPropiedadBusqueda, int iPropiedadRelacion, string sValorBusqueda, int iTipoSeleccion)
        {
            var oPropertyValues = new PropertyValues();
            var oPropertyValue = new PropertyValue();
            var oLookups = new Lookups();
            var oLookup = new Lookup();
            var oLookupSeleccionSimple = new Lookup();

            oPropertyValues = env.Vault.ObjectPropertyOperations.GetProperties(env.ObjVer);

            var searchBuilder = new MFSearchBuilder(env.Vault);
            searchBuilder.Deleted(false); // No eliminados
            searchBuilder.Class(iClase);
            searchBuilder.Property(iPropiedadBusqueda, MFDataType.MFDatatypeText, sValorBusqueda);

            var searchResults = searchBuilder.FindEx();

            if (searchResults.Count > 0)
            {
                if (iTipoSeleccion == 2)
                {
                    foreach (var result in searchResults)
                    {
                        oLookup.Item = result.ObjVer.ID;
                        oLookups.Add(-1, oLookup);
                    }

                    oPropertyValue.PropertyDef = iPropiedadRelacion;
                    oPropertyValue.TypedValue.SetValueToMultiSelectLookup(oLookups);
                    env.Vault.ObjectPropertyOperations.SetProperty(env.ObjVer, oPropertyValue);                    
                }
                else
                {
                    oLookupSeleccionSimple.Item = searchResults[0].ObjVer.ID;

                    oPropertyValue.PropertyDef = iPropiedadRelacion;
                    oPropertyValue.TypedValue.SetValueToLookup(oLookupSeleccionSimple);
                    env.Vault.ObjectPropertyOperations.SetProperty(env.ObjVer, oPropertyValue);
                }
            }
        }

        private void CreateRfcObject(EnvironmentBase env, int iObjecto, int iClase, int iPropertyDef, string sRfc, string sNombreOTitulo)
        {
            var wf_ValidacionesDocumentoProveedor = env.Vault.WorkflowOperations.GetWorkflowIDByAlias("WF.ValidacionesDocumentoProveedor");
            var wfs_DocumentoVigente = env.Vault.WorkflowOperations.GetWorkflowStateIDByAlias("WFS.ValidacionesDocumentoProveedor.DocumentoVigente");
            var pd_EjecutivoCumplimiento = env.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.EjecutivoDeCumplimiento");

            var createBuilder = new MFPropertyValuesBuilder(env.Vault);
            createBuilder.SetClass(iClase); // Clase
            createBuilder.Add
            (
                (int)MFBuiltInPropertyDef.MFBuiltInPropertyDefNameOrTitle,
                MFDataType.MFDatatypeText,
                sNombreOTitulo // Name or title
            );
            createBuilder.Add(iPropertyDef, MFDataType.MFDatatypeText, sRfc); // Rfc
            createBuilder.SetWorkflowState(wf_ValidacionesDocumentoProveedor, wfs_DocumentoVigente); // Workflow

            // Si es proveedor, agregar el ejecutivo de cumplimiento
            if (iObjecto == 212)
                createBuilder.Add(pd_EjecutivoCumplimiento, MFDataType.MFDatatypeLookup, 1);

            // Tipo de objeto a crear
            var objectTypeId = iObjecto; // Tipo de objeto

            var objectVersion = env.Vault.ObjectOperations.CreateNewObjectEx
            (
                objectTypeId,
                createBuilder.Values,
                CheckIn: true
            );
        }

        private bool GetExistingRfc(EnvironmentBase env, int iTipoObjeto, int iPropertyDef, string sRfc)
        {
            bool result = false;

            // Buscar Rfc en Proveedor, Empresa Interna o Cliente
            var searchBuilder = new MFSearchBuilder(env.Vault);
            searchBuilder.Deleted(false); // No eliminados
            searchBuilder.ObjType(iTipoObjeto);
            searchBuilder.Property(iPropertyDef, MFDataType.MFDatatypeText, sRfc);

            var searchResults = searchBuilder.FindEx();

            if (searchResults.Count > 0)
            {
                result = true;
            }

            return result;
        }

        private List<ObjVerEx> GetListCFDICompulsa(EnvironmentBase env, int iClase, int pdOrigenCFDI, int iOrigenCFDI, int pdUuid, string sUuid)
        {
            var searchBuilder = new MFSearchBuilder(env.Vault);
            searchBuilder.Deleted(false); // No eliminados
            searchBuilder.ObjType((int)MFBuiltInObjectType.MFBuiltInObjectTypeDocument); // Tipo de objeto documento
            searchBuilder.Class(iClase);
            searchBuilder.Property
            (
                pdOrigenCFDI,
                MFDataType.MFDatatypeLookup,
                iOrigenCFDI
            );
            searchBuilder.Property
            (
                pdUuid,
                MFDataType.MFDatatypeText,
                sUuid
            );

            var oResultado = searchBuilder.FindEx();

            return oResultado;
        }

        private List<ObjVerEx> GetExistingIssues(EnvironmentBase env)
        {
            var ot_Issue = env.Vault.ObjectTypeOperations.GetObjectTypeIDByAlias("MF.OT.Issue");
            var cl_Issue = env.Vault.ClassOperations.GetObjectClassIDByAlias("MF.CL.Issue");

            // Busqueda de issues existentes en la boveda
            var searchBuilder = new MFSearchBuilder(env.Vault);
            searchBuilder.Deleted(false); // No eliminados
            searchBuilder.ObjType(ot_Issue);
            searchBuilder.Class(cl_Issue);

            var results = searchBuilder.FindEx();

            return results;
        }

        private void CreateIssue(EnvironmentBase env, Lookups lookupsProveedor, Lookups lookupsEmpresaInterna, Lookups lookupsDocumentos, string sNombreOTitulo, string sDescripcion)
        {
            var ot_Issue = env.Vault.ObjectTypeOperations.GetObjectTypeIDByAlias("MF.OT.Issue");
            var cl_Issue = env.Vault.ClassOperations.GetObjectClassIDByAlias("MF.CL.Issue");
            var pd_CategoriaIssue = env.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.CategoriaDeIssue");
            var pd_IssueType = env.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("MF.PD.IssueType");
            var pd_Document = env.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.Document");
            var pd_Descripcion = env.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("M-Files.CLM.Property.Description");
            var pd_Severidad = env.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("MF.PD.Severity");
            var pd_Proveedor = env.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("MF.PD.Proveedor");
            var pd_EmpresaInterna = env.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.EmpresaInterna");
            var wf_IssueProcessing = env.Vault.WorkflowOperations.GetWorkflowIDByAlias("MF.WF.IssueProcessing");
            var wfs_Submitted = env.Vault.WorkflowOperations.GetWorkflowStateIDByAlias("M-Files.CLM.State.IssueProcessing.Submitted");

            var oLookupsCategoria = new Lookups();
            var oLookupsIssueType = new Lookups();
            var oLookupCategoria = new Lookup();
            var oLookupIssueType = new Lookup();
            var oLookupSeveridad = new Lookup();

            // Verificar que no exista un issue ya generado para los CFDI validados
            var searchBuilder = new MFSearchBuilder(env.Vault);
            searchBuilder.Deleted(false); // No eliminados
            searchBuilder.Class(cl_Issue);
            searchBuilder.Property(pd_Document, MFDataType.MFDatatypeMultiSelectLookup, lookupsDocumentos);

            var searchResults = searchBuilder.FindEx();

            if (searchResults.Count == 0)
            {
                oLookupCategoria.Item = 4;
                oLookupsCategoria.Add(-1, oLookupCategoria);

                oLookupIssueType.Item = 9;
                oLookupsIssueType.Add(-1, oLookupIssueType);

                oLookupSeveridad.Item = 2;

                var createBuilder = new MFPropertyValuesBuilder(env.Vault);
                createBuilder.SetClass(cl_Issue); // Clase issue
                createBuilder.Add
                (
                    (int)MFBuiltInPropertyDef.MFBuiltInPropertyDefNameOrTitle,
                    MFDataType.MFDatatypeText,
                    sNombreOTitulo // Name or title
                );
                createBuilder.Add(pd_Document, MFDataType.MFDatatypeMultiSelectLookup, lookupsDocumentos); // Document
                createBuilder.Add(pd_CategoriaIssue, MFDataType.MFDatatypeMultiSelectLookup, oLookupsCategoria); // Categoria del issue
                createBuilder.Add(pd_IssueType, MFDataType.MFDatatypeMultiSelectLookup, oLookupsIssueType); // Tipo de incidencia
                createBuilder.Add(pd_Descripcion, MFDataType.MFDatatypeMultiLineText, sDescripcion); // Descripcion
                createBuilder.Add(pd_Severidad, MFDataType.MFDatatypeLookup, oLookupSeveridad); // Severidad
                createBuilder.Add(pd_Proveedor, MFDataType.MFDatatypeMultiSelectLookup, lookupsProveedor); // Proveedor
                createBuilder.Add(pd_EmpresaInterna, MFDataType.MFDatatypeMultiSelectLookup, lookupsEmpresaInterna); // Empresa Interna
                createBuilder.SetWorkflowState(wf_IssueProcessing, wfs_Submitted);

                // Tipo de objeto a crear
                var objectTypeId = ot_Issue;

                // Finaliza la creacion del issue
                var objectVersion = env.Vault.ObjectOperations.CreateNewObjectEx
                (
                    objectTypeId,
                    createBuilder.Values,
                    CheckIn: true
                );
            }            
        }

        private bool CreateCFDINomina(string sFilePath, string sFileName)
        {
            var ot_ContactoExternoSE = PermanentVault.ObjectTypeOperations.GetObjectTypeIDByAlias("OT.ContactoExterno");
            var cl_ContactoExternoSE = PermanentVault.ClassOperations.GetObjectClassIDByAlias("OT.ContactoExternoSE");
            var cl_CFDINomina = PermanentVault.ClassOperations.GetObjectClassIDByAlias("CL.CfdiNomina");
            var pd_RfcEmpleado = PermanentVault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.Rfc");
            var pd_Curp = PermanentVault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.Curp");
            var pd_NombreOTituloCFDINomina = PermanentVault.PropertyDefOperations.GetPropertyDefIDByAlias("CFDI.NombreotituloNomina.Texto");
            var pd_FechaCFDI = PermanentVault.PropertyDefOperations.GetPropertyDefIDByAlias("CFDI.FechadeCFDI.Texto");
            var pd_Moneda = PermanentVault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.MonedaTextCFDInomina");
            var pd_FormaDePago = PermanentVault.PropertyDefOperations.GetPropertyDefIDByAlias("CFDI.FormadePago.Texto");
            var pd_Sello = PermanentVault.PropertyDefOperations.GetPropertyDefIDByAlias("CFDI.Sello.Texto");
            var pd_TipoDeComprobante = PermanentVault.PropertyDefOperations.GetPropertyDefIDByAlias("CFDI.TipodeComprobante.TextoCFDInomina");
            var pd_Total = PermanentVault.PropertyDefOperations.GetPropertyDefIDByAlias("CFDI.Total.Texto");
            var pd_FechaDePago = PermanentVault.PropertyDefOperations.GetPropertyDefIDByAlias("CFDI.FechaPago.Texto");
            var pd_FechaInicialPago = PermanentVault.PropertyDefOperations.GetPropertyDefIDByAlias("CFDI.FechaInicialPago.Texto");
            var pd_TipoPercepcion = PermanentVault.PropertyDefOperations.GetPropertyDefIDByAlias("CFDI.TipoPercepcion.Texto");
            var pd_DiasPagados = PermanentVault.PropertyDefOperations.GetPropertyDefIDByAlias("CFDI.NumDiasPagados.Texto");
            var pd_FechaFinalPago = PermanentVault.PropertyDefOperations.GetPropertyDefIDByAlias("CFDI.FechaFinalPago.Texto");
            var pd_PeriodicidadPago = PermanentVault.PropertyDefOperations.GetPropertyDefIDByAlias("CFDI.PeriodicidadPago.Texto");
            var pd_SalarioIntegrado = PermanentVault.PropertyDefOperations.GetPropertyDefIDByAlias("CFDI.SalarioDiarioIntegrado.Texto");
            var pd_TotalOtrosPagos = PermanentVault.PropertyDefOperations.GetPropertyDefIDByAlias("CFDI.TotalOtrosPagos.Texto");
            var pd_TotalDeducciones = PermanentVault.PropertyDefOperations.GetPropertyDefIDByAlias("CFDI.TotalDeducciones.Texto");
            var pd_Norma = PermanentVault.PropertyDefOperations.GetPropertyDefIDByAlias("CFDI.Norma.Texto");
            var pd_DisposicionFiscal = PermanentVault.PropertyDefOperations.GetPropertyDefIDByAlias("CFDI.disposicionFiscal.Texto");
            var pd_Proveedor = PermanentVault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.ProveedorServiciosEspecializados");
            var pd_SelloSAT = PermanentVault.PropertyDefOperations.GetPropertyDefIDByAlias("CFDI.SelloSAT.Texto");
            var pd_TipoMoneda = PermanentVault.PropertyDefOperations.GetPropertyDefIDByAlias("CFDI.TipodeMoneda.Texto");
            var pd_ImporteNomina = PermanentVault.PropertyDefOperations.GetPropertyDefIDByAlias("CFDI.ImporteNomina.Texto");
            var pd_DeduccionIMSS = PermanentVault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.DeduccinImssCFDInomina");
            var pd_DeduccionISR = PermanentVault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.DeduccinIsrCFDInomina");
            var pd_ContactoExternoSE = PermanentVault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.ContactoExternoSE");

            bool bReturn = false;

            // Metadata de la clase CFDI Nomina
            string sFechaCFDI = "";                      
            string sMoneda = "";
            string sFormaPago = "";
            string sSello = "";
            string sTipoComprobante = "";
            string sTotal = "";
            string sFechaPago = "";
            string sFechaPagoInicial = "";
            string sTipoPercepcion = "";
            string sDiasPagados = "";
            string sFechaPagoFinal = "";
            string sPeriodicidadPago = "";
            string sSalarioIntegrado = "";
            string sTotalOtrosPagos = "";
            string sTotalDeducciones = "";
            //string sNorma = "";
            //string sDisposicionFiscal = "";
            string sSelloSAT = "";
            //string sTipoMoneda = ""; // Repetido
            string sImporteNomina = "";
            string sImporteDeduccionIMSS = "";
            string sImporteDeduccionISR = "";

            try
            {
                XmlDocument oDocumento = new XmlDocument();
                XmlNamespaceManager oManager = new XmlNamespaceManager(oDocumento.NameTable);

                oDocumento.Load(sFilePath);
                oManager.AddNamespace("cfdi", "http://www.sat.gob.mx/cfd/3");
                oManager.AddNamespace("nomina12", "http://www.sat.gob.mx/nomina12");
                oManager.AddNamespace("tfd", "http://www.sat.gob.mx/TimbreFiscalDigital");
                XmlElement singleNode = oDocumento.DocumentElement;

                XmlNodeList oNodoComprobante = oDocumento.SelectNodes("/cfdi:Comprobante", oManager);

                foreach (XmlElement atributo in oNodoComprobante)
                {
                    if (atributo.HasAttribute("Fecha"))
                    {
                        sFechaCFDI = atributo.Attributes["Fecha"].Value;
                    }

                    if (atributo.HasAttribute("Moneda"))
                    {
                        sMoneda = atributo.Attributes["Moneda"].Value;
                    }

                    if (atributo.HasAttribute("FormaPago"))
                    {
                        sFormaPago = atributo.Attributes["FormaPago"].Value;
                    }

                    if (atributo.HasAttribute("Sello"))
                    {
                        sSello = atributo.Attributes["Sello"].Value;
                    }

                    if (atributo.HasAttribute("TipoDeComprobante"))
                    {
                        sTipoComprobante = atributo.Attributes["TipoDeComprobante"].Value;
                    }

                    if (atributo.HasAttribute("Total"))
                    {
                        sTotal = atributo.Attributes["Total"].Value;
                    }
                }

                XmlNodeList oNodoNomina12 = oDocumento.SelectNodes("/cfdi:Comprobante/cfdi:Complemento/nomina12:Nomina", oManager);

                foreach (XmlElement atributo in oNodoNomina12)
                {
                    if (atributo.HasAttribute("FechaPago"))
                    {
                        sFechaPago = atributo.Attributes["FechaPago"].Value;
                    }

                    if (atributo.HasAttribute("FechaInicialPago"))
                    {
                        sFechaPagoInicial = atributo.Attributes["FechaInicialPago"].Value;
                    }

                    if (atributo.HasAttribute("NumDiasPagados"))
                    {
                        sDiasPagados = atributo.Attributes["NumDiasPagados"].Value;
                    }

                    if (atributo.HasAttribute("FechaFinalPago"))
                    {
                        sFechaPagoFinal = atributo.Attributes["FechaFinalPago"].Value;
                    }

                    if (atributo.HasAttribute("TotalOtrosPagos"))
                    {
                        sTotalOtrosPagos = atributo.Attributes["TotalOtrosPagos"].Value;
                    }

                    if (atributo.HasAttribute("TotalDeducciones"))
                    {
                        sTotalDeducciones = atributo.Attributes["TotalDeducciones"].Value;
                    }
                }

                sTipoPercepcion = oDocumento
                    .SelectSingleNode("/cfdi:Comprobante/cfdi:Complemento/nomina12:Nomina/nomina12:Percepciones/nomina12:Percepcion/@TipoPercepcion", oManager)
                    .InnerText;

                XmlNodeList oNodoNominaReceptor = oDocumento
                    .SelectNodes("/cfdi:Comprobante/cfdi:Complemento/nomina12:Nomina/nomina12:Receptor", oManager);

                foreach (XmlElement atributo in oNodoNominaReceptor)
                {
                    if (atributo.HasAttribute("PeriodicidadPago"))
                    {
                        sPeriodicidadPago = atributo.Attributes["PeriodicidadPago"].Value;
                    }

                    if (atributo.HasAttribute("SalarioDiarioIntegrado"))
                    {
                        sSalarioIntegrado = atributo.Attributes["SalarioDiarioIntegrado"].Value;
                    }
                }

                XmlNodeList oNodoNominaDeduccion = oDocumento
                    .SelectNodes("/cfdi:Comprobante/cfdi:Complemento/nomina12:Nomina/nomina12:Deducciones/nomina12:Deduccion", oManager);

                foreach (XmlElement atributo in oNodoNominaDeduccion)
                {
                    if (atributo.HasAttribute("Concepto") && atributo.Attributes["Concepto"].Value == "IMSS")
                    {
                        sImporteDeduccionIMSS = atributo.Attributes["Importe"].Value;
                    }

                    if (atributo.HasAttribute("Concepto") && atributo.Attributes["Concepto"].Value == "ISR mes")
                    {
                        sImporteDeduccionISR = atributo.Attributes["Importe"].Value;
                    }
                }

                sSelloSAT = oDocumento
                    .SelectSingleNode("/cfdi:Comprobante/cfdi:Complemento/tfd:TimbreFiscalDigital/@SelloSAT", oManager)
                    .InnerText;

                sImporteNomina = oDocumento
                    .SelectSingleNode("/cfdi:Comprobante/cfdi:Conceptos/cfdi:Concepto/@Importe", oManager)
                    .InnerText;

                // Conversion de datos
                DateTime dtFechaCFDI = Convert.ToDateTime(sFechaCFDI);

                // Nombre o titulo de la clase y Extension
                var sDelimitador = ".";
                int iIndex = sFileName.LastIndexOf(sDelimitador);
                var sNombreOTituloCFDINomina = sFileName.Substring(0, iIndex);
                var sExtension = sFileName.Substring(iIndex + 1);

                // Obtener el contacto externo SE asociado al RFC Receptor del CFDI
                var sRfcReceptor = oDocumento.SelectSingleNode("/cfdi:Comprobante/cfdi:Receptor/@Rfc", oManager).InnerText;
                var sCurp = oDocumento.SelectSingleNode("/cfdi:Comprobante/cfdi:Complemento/nomina12:Nomina/nomina12:Receptor/@Curp", oManager).InnerText;

                var searchBuilderContactoExterno = new MFSearchBuilder(PermanentVault);
                searchBuilderContactoExterno.Deleted(false);
                searchBuilderContactoExterno.ObjType(ot_ContactoExternoSE);
                searchBuilderContactoExterno.Property(pd_RfcEmpleado, MFDataType.MFDatatypeText, sRfcReceptor);
                searchBuilderContactoExterno.Property(pd_Curp, MFDataType.MFDatatypeText, sCurp);
                searchBuilderContactoExterno.Property
                (
                    (int)MFBuiltInPropertyDef.MFBuiltInPropertyDefClass,
                    MFDataType.MFDatatypeLookup,
                    cl_ContactoExternoSE
                );

                var searchResults = searchBuilderContactoExterno.Find();

                // Crear clase CFDI Nomina
                var builderCreateDocumento = new MFPropertyValuesBuilder(this.PermanentVault);
                builderCreateDocumento.SetClass(cl_CFDINomina);
                builderCreateDocumento.Add(pd_NombreOTituloCFDINomina, MFDataType.MFDatatypeText, sNombreOTituloCFDINomina);
                builderCreateDocumento.Add(pd_FechaCFDI, MFDataType.MFDatatypeDate, dtFechaCFDI);
                builderCreateDocumento.Add(pd_Moneda, MFDataType.MFDatatypeText, sMoneda);
                builderCreateDocumento.Add(pd_FormaDePago, MFDataType.MFDatatypeText, sFormaPago);
                builderCreateDocumento.Add(pd_Sello, MFDataType.MFDatatypeText, sSello);
                builderCreateDocumento.Add(pd_TipoDeComprobante, MFDataType.MFDatatypeText, sTipoComprobante);
                builderCreateDocumento.Add(pd_Total, MFDataType.MFDatatypeFloating, sTotal);
                builderCreateDocumento.Add(pd_FechaDePago, MFDataType.MFDatatypeDate, sFechaPago);
                builderCreateDocumento.Add(pd_FechaInicialPago, MFDataType.MFDatatypeDate, sFechaPagoInicial);
                builderCreateDocumento.Add(pd_TipoPercepcion, MFDataType.MFDatatypeText, sTipoPercepcion);
                builderCreateDocumento.Add(pd_DiasPagados, MFDataType.MFDatatypeInteger, sDiasPagados);
                builderCreateDocumento.Add(pd_FechaFinalPago, MFDataType.MFDatatypeDate, sFechaPagoFinal);
                builderCreateDocumento.Add(pd_PeriodicidadPago, MFDataType.MFDatatypeText, sPeriodicidadPago);
                builderCreateDocumento.Add(pd_SalarioIntegrado, MFDataType.MFDatatypeFloating, sSalarioIntegrado);
                builderCreateDocumento.Add(pd_TotalOtrosPagos, MFDataType.MFDatatypeFloating, sTotalOtrosPagos);
                builderCreateDocumento.Add(pd_TotalDeducciones, MFDataType.MFDatatypeFloating, sTotalDeducciones);
                //builderCreateDocumento.Add(pd_Proveedor, MFDataType.MFDatatypeMultiSelectLookup, sProveedorValue);
                builderCreateDocumento.Add(pd_SelloSAT, MFDataType.MFDatatypeText, sSelloSAT);
                builderCreateDocumento.Add(pd_ImporteNomina, MFDataType.MFDatatypeText, sImporteNomina);
                builderCreateDocumento.Add(pd_DeduccionIMSS, MFDataType.MFDatatypeText, sImporteDeduccionIMSS);
                builderCreateDocumento.Add(pd_DeduccionISR, MFDataType.MFDatatypeText, sImporteDeduccionISR);

                if (searchResults.Count > 0)
                {
                    var oLookups = new Lookups();
                    var oLookup = new Lookup();

                    //foreach (ObjectVersion result in searchResults)
                    //{
                    //    oLookup.Item = result.ObjVer.ID;
                    //    oLookups.Add(-1, oLookup);
                    //}

                    oLookup.Item = searchResults[1].ObjVer.ID;
                    oLookups.Add(-1, oLookup);

                    builderCreateDocumento.Add(pd_ContactoExternoSE, MFDataType.MFDatatypeMultiSelectLookup, oLookups);
                }

                var sourceFiles = new SourceObjectFiles();

                var oXmlFile = new SourceObjectFile
                {
                    SourceFilePath = sFilePath,
                    Title = sFileName,
                    Extension = sExtension
                };                

                sourceFiles.Add(-1, oXmlFile);

                var objectTypeID = (int)MFBuiltInObjectType.MFBuiltInObjectTypeDocument;

                var isSingleFileDocument =
                    objectTypeID == (int)MFBuiltInObjectType.MFBuiltInObjectTypeDocument &&
                    sourceFiles.Count == 1;

                var objectVersion = PermanentVault.ObjectOperations.CreateNewObjectEx
                (
                    objectTypeID,
                    builderCreateDocumento.Values,
                    sourceFiles,
                    SFD: isSingleFileDocument,
                    CheckIn: true
                );

                bReturn = true;
            }
            catch (Exception ex)
            {
                SysUtils.ReportErrorMessageToEventLog("Error en ProcesarMetadataCFDINomina...", ex);
            }

            return bReturn;
        }

        private bool CreateConceptoCFDI(string sFilePath)
        {
            bool bReturn = false;

            var ot_ConceptoCFDI = PermanentVault.ObjectTypeOperations.GetObjectTypeIDByAlias("OT.ConceptoCFDI");
            var cl_ConceptoCFDI = PermanentVault.ClassOperations.GetObjectClassIDByAlias("Arkiva.Class.ConceptoDeCFDI");
            var pd_NombreOTituloConceptoCFDI = PermanentVault.PropertyDefOperations.GetPropertyDefIDByAlias("CFDI.NombreConceptoCFDI.Texto");
            var pd_ClaveProductoOServicio = PermanentVault.PropertyDefOperations.GetPropertyDefIDByAlias("CFDI.ClaveProductooServicio.Texto");
            var pd_ClaveUnidad = PermanentVault.PropertyDefOperations.GetPropertyDefIDByAlias("CFDI.ClaveUnidad.Texto");
            var pd_DescripcionConcepto = PermanentVault.PropertyDefOperations.GetPropertyDefIDByAlias("CFDI.DescripciondeConcepto.Texto");
            var pd_Cantidad = PermanentVault.PropertyDefOperations.GetPropertyDefIDByAlias("CFDI.Cantidad.Texto");
            var pd_ValorUnitario = PermanentVault.PropertyDefOperations.GetPropertyDefIDByAlias("CFDI.ValorUnitario.Texto");
            var pd_Importe = PermanentVault.PropertyDefOperations.GetPropertyDefIDByAlias("CFDI.Importe.Texto");
            var pd_Uuid = PermanentVault.PropertyDefOperations.GetPropertyDefIDByAlias("CFDI.UUID.Texto");

            // Metadata de Concepto CFDI
            string sNombreOTituloConceptoCFDI = "";
            string sClaveProductoOServicio = "";
            string sClaveUnidad = "";
            string sDescripcionConcepto = "";
            string sCantidad = "";
            string sValorUnitario = "";
            string sImporte = "";
            string sUuid = "";

            try
            {
                XmlDocument oDocumento = new XmlDocument();
                XmlNamespaceManager oManager = new XmlNamespaceManager(oDocumento.NameTable);

                oDocumento.Load(sFilePath);
                oManager.AddNamespace("cfdi", "http://www.sat.gob.mx/cfd/3");
                oManager.AddNamespace("pago10", "http://www.sat.gob.mx/Pagos");
                oManager.AddNamespace("tfd", "http://www.sat.gob.mx/TimbreFiscalDigital");

                // Nodo Conceptos de CFDI
                XmlNodeList oNodoConceptos = oDocumento.SelectNodes("/cfdi:Comprobante/cfdi:Conceptos/cfdi:Concepto", oManager);

                // Nodo Timbre Fiscal Digital
                XmlNodeList xTimbreFiscalDigital = oDocumento.SelectNodes("/cfdi:Comprobante/cfdi:Complemento/tfd:TimbreFiscalDigital", oManager);

                // Obtener los valores de la metadata de la clase Concepto CFDI
                foreach (XmlElement element in oNodoConceptos)
                {
                    if (element.HasAttribute("Descripcion"))
                    {
                        sNombreOTituloConceptoCFDI = element.Attributes["Descripcion"].Value;
                    }

                    if (element.HasAttribute("ClaveProdServ"))
                    {
                        sClaveProductoOServicio = element.Attributes["ClaveProdServ"].Value;
                    }

                    if (element.HasAttribute("ClaveUnidad"))
                    {
                        sClaveUnidad = element.Attributes["ClaveUnidad"].Value;
                    }

                    if (element.HasAttribute("Descripcion"))
                    {
                        sDescripcionConcepto = element.Attributes["Descripcion"].Value;
                    }

                    if (element.HasAttribute("Cantidad"))
                    {
                        sCantidad = element.Attributes["Cantidad"].Value;
                    }

                    if (element.HasAttribute("ValorUnitario"))
                    {
                        sValorUnitario = element.Attributes["ValorUnitario"].Value;
                    }

                    if (element.HasAttribute("Importe"))
                    {
                        sImporte = element.Attributes["Importe"].Value;
                    }

                    if (xTimbreFiscalDigital.Count > 0)
                    {
                        foreach (XmlElement item in xTimbreFiscalDigital)
                        {
                            if (item.HasAttribute("UUID"))
                            {
                                sUuid = item.Attributes["UUID"].Value;
                            }

                            break;
                        }
                    }
                }

                // Crear el objeto en la boveda
                var builderCreateDocumento = new MFPropertyValuesBuilder(this.PermanentVault);
                builderCreateDocumento.SetClass(cl_ConceptoCFDI);
                builderCreateDocumento.Add(pd_NombreOTituloConceptoCFDI, MFDataType.MFDatatypeText, sNombreOTituloConceptoCFDI);
                builderCreateDocumento.Add(pd_ClaveProductoOServicio, MFDataType.MFDatatypeText, sClaveProductoOServicio);
                builderCreateDocumento.Add(pd_ClaveUnidad, MFDataType.MFDatatypeText, sClaveUnidad);
                builderCreateDocumento.Add(pd_DescripcionConcepto, MFDataType.MFDatatypeText, sDescripcionConcepto);
                builderCreateDocumento.Add(pd_Cantidad, MFDataType.MFDatatypeFloating, sCantidad);
                builderCreateDocumento.Add(pd_ValorUnitario, MFDataType.MFDatatypeFloating, sValorUnitario);
                builderCreateDocumento.Add(pd_Importe, MFDataType.MFDatatypeFloating, sImporte);
                builderCreateDocumento.Add(pd_Uuid, MFDataType.MFDatatypeText, sUuid);

                var sourceFiles = new SourceObjectFiles();

                var objectTypeId = ot_ConceptoCFDI;

                var isSingleFileDocument =
                    objectTypeId == (int)MFBuiltInObjectType.MFBuiltInObjectTypeDocument &&
                    sourceFiles.Count == 1;

                var objectVersion = PermanentVault.ObjectOperations.CreateNewObjectEx
                (
                    objectTypeId,
                    builderCreateDocumento.Values,
                    sourceFiles,
                    SFD: isSingleFileDocument,
                    CheckIn: true
                );

                bReturn = true;
            }
            catch (Exception ex)
            {
                SysUtils.ReportErrorMessageToEventLog("Error al intentar crear Concepto de CFDI...", ex);
            }

            return bReturn;
        }

        private bool CreateComplementoPago(string sFilePath, string sCfdiComprobante)
        {
            bool bReturn = false;

            var cl_ComplementoPagoEmitido = PermanentVault.ClassOperations.GetObjectClassIDByAlias("Arkiva.Class.CFDIComplementoPagoEmitido");
            var cl_ComplementoPagoRecibido = PermanentVault.ClassOperations.GetObjectClassIDByAlias("Arkiva.Class.CFDIComplementoPagoRecibido");
            var pd_ImporteSaldoInsoluto = PermanentVault.PropertyDefOperations.GetPropertyDefIDByAlias("CFDI.ImporteSaldoInsoluto.Texto");
            var pd_ImportePagado = PermanentVault.PropertyDefOperations.GetPropertyDefIDByAlias("CFDI.ImportePagado.Texto");
            var pd_ImporteSaldoAnterior = PermanentVault.PropertyDefOperations.GetPropertyDefIDByAlias("CFDI.ImporteSaldoAnterior.Texto");
            var pd_Folio = PermanentVault.PropertyDefOperations.GetPropertyDefIDByAlias("CFDI.FoliodeCFDI.TextoCFDInomina");
            var pd_Monto = PermanentVault.PropertyDefOperations.GetPropertyDefIDByAlias("CFDI.Monto.Texto");
            var pd_TipoMoneda = PermanentVault.PropertyDefOperations.GetPropertyDefIDByAlias("CFDI.TipodeMoneda.Texto");
            var pd_FormaPago = PermanentVault.PropertyDefOperations.GetPropertyDefIDByAlias("CFDI.FormadePago.Texto");
            var pd_FechaPago = PermanentVault.PropertyDefOperations.GetPropertyDefIDByAlias("CFDI.FechaPago.Texto");
            var pd_Version = PermanentVault.PropertyDefOperations.GetPropertyDefIDByAlias("CFDI.VersionCFDI.Texto");
            var pd_Uuid = PermanentVault.PropertyDefOperations.GetPropertyDefIDByAlias("CFDI.UUID.Texto");
            var pd_SelloSAT = PermanentVault.PropertyDefOperations.GetPropertyDefIDByAlias("CFDI.SelloSAT.Texto");
            var pd_SelloCFD = PermanentVault.PropertyDefOperations.GetPropertyDefIDByAlias("CFDI.SelloCFD.Texto");
            var pd_FechaTimbrado = PermanentVault.PropertyDefOperations.GetPropertyDefIDByAlias("CFDI.FechadeTimbrado.Texto");
            var pd_RfcEmisor = PermanentVault.PropertyDefOperations.GetPropertyDefIDByAlias("CFDI.RFCEmisor.Texto");

            // Variables para almacenar la metadata del complemento de pago
            string sImporteSaldoInsoluto = "";
            string sImportePagado = "";
            string sImporteSaldoAnterior = "";
            string sFolio = "";
            string sMonto = "";
            string sTipoMoneda = "";
            string sFormaPago = "";
            string sFechaPago = "";
            string sVersion = "";
            string sUuid = "";
            string sSelloSAT = "";
            string sSelloCFD = "";
            string sFechaTimbrado = "";
            string sRfcEmisor = "";
            int iClase = 0;

            // Si el comprobante es recibido, el complemento de pago debe ser emitido, de lo contrario es complemento de pago recibido
            if (sCfdiComprobante == "Recibido")
                iClase = cl_ComplementoPagoEmitido;
            else
                iClase = cl_ComplementoPagoRecibido;

            try
            {
                XmlDocument oDocumento = new XmlDocument();
                XmlNamespaceManager oManager = new XmlNamespaceManager(oDocumento.NameTable);

                oDocumento.Load(sFilePath);
                oManager.AddNamespace("cfdi", "http://www.sat.gob.mx/cfd/3");
                oManager.AddNamespace("pago10", "http://www.sat.gob.mx/Pagos");
                oManager.AddNamespace("tfd", "http://www.sat.gob.mx/TimbreFiscalDigital");

                //Complemento de pago
                XmlNodeList nodoPago = oDocumento.SelectNodes("/cfdi:Comprobante/cfdi:Complemento/pago10:Pagos/pago10:Pago", oManager);

                //Documentos relacionados del complemento de pago 
                XmlNodeList nodoDoctoRelacionado = oDocumento.SelectNodes("/cfdi:Comprobante/cfdi:Complemento/pago10:Pagos/pago10:Pago/pago10:DoctoRelacionado", oManager);

                //Timbrado Fiscal Digital
                XmlNodeList nodoTimbreFiscalDigital = oDocumento.SelectNodes("/cfdi:Comprobante/cfdi:Complemento/tfd:TimbreFiscalDigital", oManager);

                //Rfc Emisor
                XmlNodeList nodoEmisor = oDocumento.SelectNodes("/cfdi:Comprobante/cfdi:Emisor", oManager);

                foreach (XmlElement element in nodoDoctoRelacionado)
                {
                    if (element.HasAttribute("ImpSaldoInsoluto"))
                        sImporteSaldoInsoluto = element.Attributes["ImpSaldoInsoluto"].Value;

                    if (element.HasAttribute("ImpPagado"))
                        sImportePagado = element.Attributes["ImpPagado"].Value;

                    if (element.HasAttribute("ImpSaldoAnt"))
                        sImporteSaldoAnterior = element.Attributes["ImpSaldoAnt"].Value;

                    if (element.HasAttribute("Folio"))
                        sFolio = element.Attributes["Folio"].Value;

                    foreach (XmlElement ePago in nodoPago)
                    {
                        if (ePago.HasAttribute("Monto"))
                            sMonto = ePago.Attributes["Monto"].Value;

                        if (ePago.HasAttribute("MonedaP"))
                            sTipoMoneda = ePago.Attributes["MonedaP"].Value;

                        if (ePago.HasAttribute("FormaDePagoP"))
                            sFormaPago = ePago.Attributes["FormaDePagoP"].Value;

                        if (ePago.HasAttribute("FechaPago"))
                            sFechaPago = ePago.Attributes["FechaPago"].Value;

                        break;
                    }

                    foreach (XmlElement eTFD in nodoTimbreFiscalDigital)
                    {
                        if (eTFD.HasAttribute("Version"))
                            sVersion = eTFD.Attributes["Version"].Value;

                        if (eTFD.HasAttribute("UUID"))
                            sUuid = eTFD.Attributes["UUID"].Value;

                        if (eTFD.HasAttribute("SelloSAT"))
                            sSelloSAT = eTFD.Attributes["SelloSAT"].Value;

                        if (eTFD.HasAttribute("SelloCFD"))
                            sSelloCFD = eTFD.Attributes["SelloCFD"].Value;

                        if (eTFD.HasAttribute("FechaTimbrado"))
                            sFechaTimbrado = eTFD.Attributes["FechaTimbrado"].Value;

                        break;
                    }

                    foreach (XmlElement eEmisor in nodoEmisor)
                    {
                        if (eEmisor.HasAttribute("Rfc"))
                            sRfcEmisor = eEmisor.Attributes["Rfc"].Value;

                        break;
                    }
                }

                // Conversion a tipo de dato Fecha
                DateTime dtFechaPago = new DateTime();
                if (!string.IsNullOrEmpty(sFechaPago))
                    dtFechaPago = Convert.ToDateTime(sFechaPago);

                DateTime dtFechaTimbrado = new DateTime();
                if (!string.IsNullOrEmpty(sFechaTimbrado))
                    dtFechaTimbrado = Convert.ToDateTime(sFechaTimbrado);

                // Crear el objeto en la boveda
                var builderCreateDocumento = new MFPropertyValuesBuilder(this.PermanentVault);
                builderCreateDocumento.SetClass(cl_ComplementoPagoEmitido);
                builderCreateDocumento.Add(pd_FechaPago, MFDataType.MFDatatypeDate, dtFechaPago);
                builderCreateDocumento.Add(pd_FechaTimbrado, MFDataType.MFDatatypeDate, dtFechaTimbrado);
                builderCreateDocumento.Add(pd_Folio, MFDataType.MFDatatypeText, sFolio);
                builderCreateDocumento.Add(pd_ImportePagado, MFDataType.MFDatatypeFloating, sImportePagado);
                builderCreateDocumento.Add(pd_ImporteSaldoAnterior, MFDataType.MFDatatypeFloating, sImporteSaldoAnterior);
                builderCreateDocumento.Add(pd_ImporteSaldoInsoluto, MFDataType.MFDatatypeFloating, sImporteSaldoInsoluto);
                builderCreateDocumento.Add(pd_Monto, MFDataType.MFDatatypeFloating, sMonto);
                builderCreateDocumento.Add(pd_SelloSAT, MFDataType.MFDatatypeText, sSelloSAT);
                builderCreateDocumento.Add(pd_SelloCFD, MFDataType.MFDatatypeText, sSelloCFD);
                builderCreateDocumento.Add(pd_Uuid, MFDataType.MFDatatypeText, sUuid);
                builderCreateDocumento.Add(pd_Version, MFDataType.MFDatatypeText, sVersion);
                builderCreateDocumento.Add(pd_FormaPago, MFDataType.MFDatatypeText, sFormaPago);
                builderCreateDocumento.Add(pd_TipoMoneda, MFDataType.MFDatatypeText, sTipoMoneda);
                builderCreateDocumento.Add(pd_RfcEmisor, MFDataType.MFDatatypeText, sRfcEmisor);

                var sourceFiles = new SourceObjectFiles();

                var objectTypeID = (int)MFBuiltInObjectType.MFBuiltInObjectTypeDocument;

                var isSingleFileDocument =
                    objectTypeID == (int)MFBuiltInObjectType.MFBuiltInObjectTypeDocument &&
                    sourceFiles.Count == 1;

                var objectVersion = PermanentVault.ObjectOperations.CreateNewObjectEx
                (
                    objectTypeID,
                    builderCreateDocumento.Values,
                    sourceFiles,
                    SFD: isSingleFileDocument,
                    CheckIn: true
                );

                bReturn = true;
            }
            catch (Exception ex)
            {
                SysUtils.ReportErrorMessageToEventLog("Error al intentar crear Complemento de Pago...", ex);
            }

            return bReturn;
        }

        private bool CreateComprobanteCFDI(string sFilePath, string sFileName, string sCfdiComprobante)
        {
            bool bReturn = false;

            var cl_CFDIComprobanteRecibido = PermanentVault.ClassOperations.GetObjectClassIDByAlias("Arkiva.Class.CFDIComprobanteRecibido");
            var cl_CFDIComprobanteEmitido = PermanentVault.ClassOperations.GetObjectClassIDByAlias("Arkiva.Class.CFDIComprobanteEmitido");
            var pd_Version = PermanentVault.PropertyDefOperations.GetPropertyDefIDByAlias("CFDI.VersionCFDI.Texto");
            var pd_Total = PermanentVault.PropertyDefOperations.GetPropertyDefIDByAlias("CFDI.Total.Texto");
            var pd_TipoComprobante = PermanentVault.PropertyDefOperations.GetPropertyDefIDByAlias("CFDI.TipodeComprobante.TextoCFDInomina");
            var pd_TipoCambio = PermanentVault.PropertyDefOperations.GetPropertyDefIDByAlias("CFDI.TipodeCambio.Texto");
            var pd_Subtotal = PermanentVault.PropertyDefOperations.GetPropertyDefIDByAlias("CFDI.SubTotal.Texto");
            var pd_Serie = PermanentVault.PropertyDefOperations.GetPropertyDefIDByAlias("CFDI.Serie.TextoCFDInomina");
            var pd_Sello = PermanentVault.PropertyDefOperations.GetPropertyDefIDByAlias("CFDI.Sello.Texto");
            var pd_NoCertificado = PermanentVault.PropertyDefOperations.GetPropertyDefIDByAlias("CFDI.NumeroCertificado.TextoCFDInomina");
            var pd_Moneda = PermanentVault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.MonedaTextCFDInomina");
            var pd_MetodoPago = PermanentVault.PropertyDefOperations.GetPropertyDefIDByAlias("CFDI.MétododePago.TextoCFDInomina");
            var pd_LugarExpedicion = PermanentVault.PropertyDefOperations.GetPropertyDefIDByAlias("CFDI.LugardeExpedición.Texto");
            var pd_FormaPago = PermanentVault.PropertyDefOperations.GetPropertyDefIDByAlias("CFDI.FormadePago.Texto");
            var pd_Folio = PermanentVault.PropertyDefOperations.GetPropertyDefIDByAlias("CFDI.FoliodeCFDI.TextoCFDInomina");
            var pd_Fecha = PermanentVault.PropertyDefOperations.GetPropertyDefIDByAlias("CFDI.FechadeCFDI.Texto");
            var pd_CondicionesPago = PermanentVault.PropertyDefOperations.GetPropertyDefIDByAlias("CFDI.CondiciondePago.Texto");
            var pd_Certificado = PermanentVault.PropertyDefOperations.GetPropertyDefIDByAlias("CFDI.CertificadoCFDI.TextoCFDInomina");
            var pd_RfcEmisor = PermanentVault.PropertyDefOperations.GetPropertyDefIDByAlias("CFDI.RFCEmisor.Texto");
            var pd_NombreEmisor = PermanentVault.PropertyDefOperations.GetPropertyDefIDByAlias("CFDI.EmisorCFDI.Texto");
            var pd_RegimenFiscal = PermanentVault.PropertyDefOperations.GetPropertyDefIDByAlias("CFDI.RegimenFiscal.Texto");
            var pd_RfcReceptor = PermanentVault.PropertyDefOperations.GetPropertyDefIDByAlias("CFDI.RFCReceptor.Texto");
            var pd_NombreReceptor = PermanentVault.PropertyDefOperations.GetPropertyDefIDByAlias("CFDI.ReceptorCFDI.Texto");
            var pd_UsoCfdi = PermanentVault.PropertyDefOperations.GetPropertyDefIDByAlias("CFDI.UsodeCFDI.Texto");
            var pd_TotalImpuestosTrasladados = PermanentVault.PropertyDefOperations.GetPropertyDefIDByAlias("CFDI.TotalImpuestosTrasladados.Texto");
            var pd_Importe = PermanentVault.PropertyDefOperations.GetPropertyDefIDByAlias("CFDI.Importe.Texto");
            var pd_TipoFactor = PermanentVault.PropertyDefOperations.GetPropertyDefIDByAlias("CFDI.TipodeFactor.Texto");
            var pd_TasaOCuota = PermanentVault.PropertyDefOperations.GetPropertyDefIDByAlias("CFDI.TasaoCuota.Texto");
            var pd_Impuesto = PermanentVault.PropertyDefOperations.GetPropertyDefIDByAlias("CFDI.Impuestos.Texto");
            var pd_Uuid = PermanentVault.PropertyDefOperations.GetPropertyDefIDByAlias("CFDI.UUID.Texto");
            var pd_OrigenCFDI = PermanentVault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.OrigenDelCfdi");

            string sVersion = "";
            string sTotal = "";
            string sTipoComprobante = "";
            string sTipoCambio = "";
            string sSubtotal = "";
            string sSerie = "";
            string sSello = "";
            string sNoCertificado = "";
            string sMoneda = "";
            string sMetodoPago = "";
            string sLugarExpedicion = "";
            string sFormaPago = "";
            string sFolio = "";
            string sFecha = "";
            string sCondicionesPago = "";
            string sCertificado = "";
            string sRfcEmisor = "";
            string sNombreEmisor = "";
            string sRegimenFiscal = "";
            string sRfcReceptor = "";
            string sNombreReceptor = "";
            string sUsoCfdi = "";
            string sTotalImpuestosTrasladados = "";
            string sImporte = "";
            string sTipoFactor = "";
            string sTasaOCuota = "";
            string sImpuesto = "";
            string sUuid = "";
            int iClase = 0;

            // Generar los Conceptos del CFDI
            CreateConceptoCFDI(sFilePath);

            // Generar el Complemento de Pago
            CreateComplementoPago(sFilePath, sCfdiComprobante);

            // Identificar la clase Comprobante a generar, Emitido o Recibido
            if (sCfdiComprobante == "Recibido")
                iClase = cl_CFDIComprobanteRecibido;
            else
                iClase = cl_CFDIComprobanteEmitido;

            try
            {
                // Generar el Comprobante CFDI Emitido o Recibido
                XmlDocument oDocumento = new XmlDocument();
                XmlNamespaceManager oManager = new XmlNamespaceManager(oDocumento.NameTable);

                oDocumento.Load(sFilePath);
                oManager.AddNamespace("cfdi", "http://www.sat.gob.mx/cfd/3");
                oManager.AddNamespace("pago10", "http://www.sat.gob.mx/Pagos");
                oManager.AddNamespace("tfd", "http://www.sat.gob.mx/TimbreFiscalDigital");

                // Nodo Comprobante
                XmlNodeList nodoComprobante = oDocumento.SelectNodes("/cfdi:Comprobante", oManager);

                foreach (XmlElement element in nodoComprobante)
                {
                    if (element.HasAttribute("Version"))
                        sVersion = element.Attributes["Version"].Value;

                    if (element.HasAttribute("Total"))
                        sTotal = element.Attributes["Total"].Value;

                    if (element.HasAttribute("TipoDeComprobante"))
                        sTipoComprobante = element.Attributes["TipoDeComprobante"].Value;

                    if (element.HasAttribute("TipoCambio"))
                        sTipoCambio = element.Attributes["TipoCambio"].Value;

                    if (element.HasAttribute("SubTotal"))
                        sSubtotal = element.Attributes["SubTotal"].Value;

                    if (element.HasAttribute("Serie"))
                        sSerie = element.Attributes["Serie"].Value;

                    if (element.HasAttribute("Sello"))
                        sSello = element.Attributes["Sello"].Value;

                    if (element.HasAttribute("NoCertificado"))
                        sNoCertificado = element.Attributes["NoCertificado"].Value;

                    if (element.HasAttribute("Moneda"))
                        sMoneda = element.Attributes["Moneda"].Value;

                    if (element.HasAttribute("MetodoPago"))
                        sMetodoPago = element.Attributes["MetodoPago"].Value;

                    if (element.HasAttribute("LugarExpedicion"))
                        sLugarExpedicion = element.Attributes["LugarExpedicion"].Value;

                    if (element.HasAttribute("FormaPago"))
                        sFormaPago = element.Attributes["FormaPago"].Value;

                    if (element.HasAttribute("Folio"))
                        sFolio = element.Attributes["Folio"].Value;

                    if (element.HasAttribute("Fecha"))
                        sFecha = element.Attributes["Fecha"].Value;

                    if (element.HasAttribute("CondicionesDePago"))
                        sCondicionesPago = element.Attributes["CondicionesDePago"].Value;

                    if (element.HasAttribute("Certificado"))
                        sCertificado = element.Attributes["Certificado"].Value;
                }

                // Nodo Emisor
                XmlNodeList nodoEmisor = oDocumento.SelectNodes("/cfdi:Comprobante/cfdi:Emisor", oManager);

                foreach (XmlElement element in nodoEmisor)
                {
                    if (element.HasAttribute("Rfc"))
                        sRfcEmisor = element.Attributes["Rfc"].Value;

                    if (element.HasAttribute("Nombre"))
                        sNombreEmisor = element.Attributes["Nombre"].Value;

                    if (element.HasAttribute("RegimenFiscal"))
                        sRegimenFiscal = element.Attributes["RegimenFiscal"].Value;
                }

                // Nodo Receptor
                XmlNodeList nodoReceptor = oDocumento.SelectNodes("/cfdi:Comprobante/cfdi:Receptor", oManager);

                foreach (XmlElement element in nodoReceptor)
                {
                    if (element.HasAttribute("Rfc"))
                        sRfcReceptor = element.Attributes["Rfc"].Value;

                    if (element.HasAttribute("Nombre"))
                        sNombreReceptor = element.Attributes["Nombre"].Value;

                    if (element.HasAttribute("UsoCFDI"))
                        sUsoCfdi = element.Attributes["UsoCFDI"].Value;
                }

                // Nodo Total de Impuestos Trasladados
                XmlNodeList nodoImpuestos = oDocumento.SelectNodes("/cfdi:Comprobante/cfdi:Impuestos", oManager);

                foreach (XmlElement element in nodoImpuestos)
                {
                    if (element.HasAttribute("TotalImpuestosTrasladados"))
                        sTotalImpuestosTrasladados = element.Attributes["TotalImpuestosTrasladados"].Value;
                }

                // Nodo Traslados
                XmlNodeList nodoTraslados = oDocumento.SelectNodes("/cfdi:Comprobante/cfdi:Impuestos/cfdi:Traslados/cfdi:Traslado", oManager);

                foreach (XmlElement element in nodoTraslados)
                {
                    if (element.HasAttribute("Importe"))
                        sImporte = element.Attributes["Importe"].Value;

                    if (element.HasAttribute("TipoFactor"))
                        sTipoFactor = element.Attributes["TipoFactor"].Value;

                    if (element.HasAttribute("TasaOCuota"))
                        sTasaOCuota = element.Attributes["TasaOCuota"].Value;

                    if (element.HasAttribute("Impuesto"))
                        sImpuesto = element.Attributes["Impuesto"].Value;
                }

                // Nodo Timbre Fiscal Digital
                XmlNodeList nodoTimbreFiscalDigital = oDocumento.SelectNodes("/cfdi:Comprobante/cfdi:Complemento/tfd:TimbreFiscalDigital", oManager);

                foreach (XmlElement element in nodoTimbreFiscalDigital)
                {
                    if (element.HasAttribute("UUID"))
                        sUuid = element.Attributes["UUID"].Value;
                }

                // Conversion a tipo de dato Fecha
                DateTime dtFecha = Convert.ToDateTime(sFecha);

                // Crear el objeto en la boveda
                var builderCreateDocumento = new MFPropertyValuesBuilder(this.PermanentVault);
                builderCreateDocumento.SetClass(cl_CFDIComprobanteRecibido);
                builderCreateDocumento.Add(pd_Version, MFDataType.MFDatatypeText, sVersion);
                builderCreateDocumento.Add(pd_Total, MFDataType.MFDatatypeFloating, sTotal);
                builderCreateDocumento.Add(pd_TipoComprobante, MFDataType.MFDatatypeText, sTipoComprobante);
                builderCreateDocumento.Add(pd_TipoCambio, MFDataType.MFDatatypeFloating, sTipoCambio);
                builderCreateDocumento.Add(pd_Subtotal, MFDataType.MFDatatypeFloating, sSubtotal);
                builderCreateDocumento.Add(pd_Serie, MFDataType.MFDatatypeText, sSerie);
                builderCreateDocumento.Add(pd_Sello, MFDataType.MFDatatypeText, sSello);
                builderCreateDocumento.Add(pd_NoCertificado, MFDataType.MFDatatypeText, sNoCertificado);
                builderCreateDocumento.Add(pd_Moneda, MFDataType.MFDatatypeText, sMoneda);
                builderCreateDocumento.Add(pd_MetodoPago, MFDataType.MFDatatypeText, sMetodoPago);
                builderCreateDocumento.Add(pd_LugarExpedicion, MFDataType.MFDatatypeText, sLugarExpedicion);
                builderCreateDocumento.Add(pd_FormaPago, MFDataType.MFDatatypeText, sFormaPago);
                builderCreateDocumento.Add(pd_Folio, MFDataType.MFDatatypeText, sFolio);
                builderCreateDocumento.Add(pd_Fecha, MFDataType.MFDatatypeDate, dtFecha);
                builderCreateDocumento.Add(pd_CondicionesPago, MFDataType.MFDatatypeText, sCondicionesPago);
                builderCreateDocumento.Add(pd_Certificado, MFDataType.MFDatatypeText, sCertificado);
                builderCreateDocumento.Add(pd_RfcEmisor, MFDataType.MFDatatypeText, sRfcEmisor);
                builderCreateDocumento.Add(pd_NombreEmisor, MFDataType.MFDatatypeText, sNombreEmisor);
                builderCreateDocumento.Add(pd_RegimenFiscal, MFDataType.MFDatatypeText, sRegimenFiscal);
                builderCreateDocumento.Add(pd_RfcReceptor, MFDataType.MFDatatypeText, sRfcReceptor);
                builderCreateDocumento.Add(pd_NombreReceptor, MFDataType.MFDatatypeText, sNombreReceptor);
                builderCreateDocumento.Add(pd_UsoCfdi, MFDataType.MFDatatypeText, sUsoCfdi);
                builderCreateDocumento.Add(pd_TotalImpuestosTrasladados, MFDataType.MFDatatypeFloating, sTotalImpuestosTrasladados);
                builderCreateDocumento.Add(pd_Importe, MFDataType.MFDatatypeFloating, sImporte);
                builderCreateDocumento.Add(pd_TipoFactor, MFDataType.MFDatatypeText, sTipoFactor);
                builderCreateDocumento.Add(pd_TasaOCuota, MFDataType.MFDatatypeText, sTasaOCuota);
                builderCreateDocumento.Add(pd_Impuesto, MFDataType.MFDatatypeFloating, sImpuesto);
                builderCreateDocumento.Add(pd_Uuid, MFDataType.MFDatatypeText, sUuid);
                builderCreateDocumento.Add(pd_OrigenCFDI, MFDataType.MFDatatypeLookup, 2);

                var sourceFiles = new SourceObjectFiles();

                var oXmlFile = new SourceObjectFile
                {
                    SourceFilePath = sFilePath,
                    Title = sFileName,
                    Extension = ".xml"
                };

                sourceFiles.Add(-1, oXmlFile);

                var objectTypeID = (int)MFBuiltInObjectType.MFBuiltInObjectTypeDocument;

                var isSingleFileDocument =
                    objectTypeID == (int)MFBuiltInObjectType.MFBuiltInObjectTypeDocument &&
                    sourceFiles.Count == 1;

                var objectVersion = PermanentVault.ObjectOperations.CreateNewObjectEx
                (
                    objectTypeID,
                    builderCreateDocumento.Values,
                    sourceFiles,
                    SFD: isSingleFileDocument,
                    CheckIn: true
                );

                bReturn = true;
            }
            catch (Exception ex)
            {
                SysUtils.ReportErrorMessageToEventLog("Error al intentar crear Comprobante CFDI ...", ex);
            }

            return bReturn;
        }
    }
}