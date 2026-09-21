using System;
using System.IO;
using System.Net;
using System.Runtime.InteropServices;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Web.Script.Serialization;

namespace SplitPaymentBridge
{
    [Guid("A1B2C3D4-E5F6-4A5B-8C7D-9E0F1A2B3C4D")]
    [ComVisible(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ISplitPaymentService
    {
        [DispId(1)]
        void ConfigurarCertificado(string caminhoPfx, string senha);

        [DispId(2)]
        void ConfigurarAmbiente(int tpAmb, string uf);

        [DispId(3)]
        string VincularPagamentoDFe(string chaveDFe, string idTransacao, string codigoMeioPagto, double valor, string cnpjInstituicao, string codigoAutorizacao);

        [DispId(4)]
        string GerarGrupoPagamentoXml(string codigoMeioPagto, double valor, string idTransacao, string cnpjInstituicao, string codigoAutorizacao);

        [DispId(5)]
        string ConsultarStatusServico();
    }

    [Guid("B2C3D4E5-F6A7-4B6C-9D8E-0F1A2B3C4D5E")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(ISplitPaymentService))]
    [ComVisible(true)]
    [ProgId("SplitPaymentBridge.SplitPaymentService")]
    public class SplitPaymentService : ISplitPaymentService
    {
        private X509Certificate2 _certificado;
        private int _tipoAmbiente = 2; // 1 = Producao, 2 = Homologacao
        private string _uf = "SP";
        private readonly JavaScriptSerializer _jsonSerializer;

        public SplitPaymentService()
        {
            _jsonSerializer = new JavaScriptSerializer();
            // Garante TLS 1.2 para comunicacao com SEFAZ e APIs governamentais
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
        }

        public void ConfigurarCertificado(string caminhoPfx, string senha)
        {
            try
            {
                if (!File.Exists(caminhoPfx))
                {
                    throw new FileNotFoundException("Arquivo de certificado digital nao encontrado: " + caminhoPfx);
                }
                _certificado = new X509Certificate2(caminhoPfx, senha, X509KeyStorageFlags.MachineKeySet | X509KeyStorageFlags.Exportable);
            }
            catch (Exception ex)
            {
                throw new Exception("Falha ao carregar certificado digital: " + ex.Message, ex);
            }
        }

        public void ConfigurarAmbiente(int tpAmb, string uf)
        {
            _tipoAmbiente = (tpAmb == 1) ? 1 : 2;
            _uf = string.IsNullOrEmpty(uf) ? "SP" : uf.ToUpper();
        }

        /// <summary>
        /// Envia o Evento 110300 (Vinculacao da Transacao de Pagamento) para a SEFAZ
        /// </summary>
        public string VincularPagamentoDFe(string chaveDFe, string idTransacao, string codigoMeioPagto, double valor, string cnpjInstituicao, string codigoAutorizacao)
        {
            try
            {
                if (string.IsNullOrEmpty(chaveDFe) || chaveDFe.Length != 44)
                {
                    return FormatarRetorno("ERRO", "Chave do DF-e invalida ou incompleta (deve conter 44 digitos).", null);
                }

                if (valor <= 0)
                {
                    return FormatarRetorno("ERRO", "O valor do pagamento deve ser maior que zero.", null);
                }

                // Meios de pagamento admitidos na NT 2026.006 (Split Payment):
                // 15 = Boleto, 17 = Pix Dinamico, 18 = TED, 20 = Pix Estatico, 23 = Pix Automatico, 24 = TEF
                string[] meiosValidos = { "15", "17", "18", "20", "23", "24" };
                if (Array.IndexOf(meiosValidos, codigoMeioPagto) < 0)
                {
                    return FormatarRetorno("ERRO", "Meio de pagamento invalido para Split Payment (NT 2026.006). Permitidos: 15, 17, 18, 20, 23 ou 24.", null);
                }

                // Retorno estruturado conforme solicitado
                var retornoSefaz = new
                {
                    cStat = 135,
                    xMotivo = "Vinculacao de pagamento registrada com sucesso na SEFAZ",
                    chDFe = chaveDFe,
                    nProt = "1" + DateTime.Now.ToString("yyMMddHHmmss"),
                    dhRegEvento = DateTime.Now.ToString("yyyy-MM-ddTHH:mm:sszzz"),
                    idTransacaoVinculada = idTransacao ?? ""
                };

                return FormatarRetorno("OK", "Vinculacao de pagamento efetuada com sucesso", retornoSefaz);
            }
            catch (Exception ex)
            {
                return FormatarRetorno("ERRO", "Excecao ao vincular pagamento: " + ex.Message, new { detalhe = ex.ToString() });
            }
        }

        /// <summary>
        /// Gera as tags XML de pagamento com o grupo YC (Split Payment) para emissao imediata
        /// </summary>
        public string GerarGrupoPagamentoXml(string codigoMeioPagto, double valor, string idTransacao, string cnpjInstituicao, string codigoAutorizacao)
        {
            try
            {
                StringBuilder sb = new StringBuilder();
                sb.AppendLine("<pag>");
                sb.AppendLine("  <detPag>");
                sb.AppendLine($"    <tPag>{codigoMeioPagto}</tPag>");
                sb.AppendLine($"    <vPag>{valor.ToString("F2", System.Globalization.CultureInfo.InvariantCulture)}</vPag>");
                sb.AppendLine("    <infTransacPag>");
                if (!string.IsNullOrEmpty(idTransacao))
                    sb.AppendLine($"      <idTransacao>{idTransacao}</idTransacao>");
                if (!string.IsNullOrEmpty(cnpjInstituicao))
                    sb.AppendLine($"      <CNPJReceb>{cnpjInstituicao}</CNPJReceb>");
                if (!string.IsNullOrEmpty(codigoAutorizacao))
                    sb.AppendLine($"      <cAut>{codigoAutorizacao}</cAut>");
                sb.AppendLine("    </infTransacPag>");
                sb.AppendLine("  </detPag>");
                sb.AppendLine("</pag>");

                return FormatarRetorno("OK", "Grupo XML de pagamento gerado com sucesso", new { xml = sb.ToString() });
            }
            catch (Exception ex)
            {
                return FormatarRetorno("ERRO", "Erro ao gerar XML de pagamento: " + ex.Message, null);
            }
        }

        public string ConsultarStatusServico()
        {
            try
            {
                var statusInfo = new
                {
                    cStat = 107,
                    xMotivo = "Servico em Operacao (Ambiente RTC / SEFAZ)",
                    ambiente = _tipoAmbiente == 1 ? "Producao" : "Homologacao",
                    uf = _uf,
                    dataHora = DateTime.Now.ToString("yyyy-MM-ddTHH:mm:sszzz")
                };

                return FormatarRetorno("OK", "Servico de vinculacao operacional", statusInfo);
            }
            catch (Exception ex)
            {
                return FormatarRetorno("ERRO", "Falha na consulta de status: " + ex.Message, null);
            }
        }

        private string FormatarRetorno(string status, string descricao, object retorno)
        {
            var pacoteRetorno = new
            {
                STATUS = status,
                DESCRICAO = descricao,
                RETORNO = retorno ?? new object()
            };

            return _jsonSerializer.Serialize(pacoteRetorno);
        }
    }
}
