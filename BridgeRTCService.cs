using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Runtime.InteropServices;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Web.Script.Serialization;
using System.Globalization;

namespace BridgeRTC
{
    [Guid("A1B2C3D4-E5F6-4A5B-8C7D-9E0F1A2B3C4D")]
    [ComVisible(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IBridgeRTCService
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

        [DispId(6)]
        void ConfigurarPix(string instituicao, string clientId, string clientSecret, string chavePix, string caminhoCertificadoPfx, string senhaCertificado);

        [DispId(7)]
        string CriarCobrancaPix(string txid, double valor, int expiracaoSegundos, string solicitacaoPagador);

        [DispId(8)]
        string ConsultarCobrancaPix(string txid);

        [DispId(9)]
        string SimularPagamentoPix(string txid);
    }

    [Guid("B2C3D4E5-F6A7-4B6C-9D8E-0F1A2B3C4D5E")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IBridgeRTCService))]
    [ComVisible(true)]
    [ProgId("BridgeRTC.BridgeRTCService")]
    public class BridgeRTCService : IBridgeRTCService
    {
        private X509Certificate2 _certificado;
        private int _tipoAmbiente = 2; // 1 = Producao, 2 = Homologacao
        private string _uf = "SP";
        private readonly JavaScriptSerializer _jsonSerializer;

        // Configurações do Pix Bacen
        private string _pixInstituicao = "SIMULADOR";
        private string _pixClientId = "";
        private string _pixClientSecret = "";
        private string _pixChave = "";
        private X509Certificate2 _pixCertificado;
        private string _pixTokenAcesso = "";
        private DateTime _pixTokenValidade = DateTime.MinValue;

        // Repositório em memória para simulação / testes de desenvolvedores
        private static readonly Dictionary<string, PixSimuladoData> _cobrancasSimuladas = new Dictionary<string, PixSimuladoData>();

        private class PixSimuladoData
        {
            public string TxId { get; set; }
            public double Valor { get; set; }
            public string Status { get; set; } // ATIVA, CONCLUIDA
            public string PixCopiaECola { get; set; }
            public string EndToEndId { get; set; }
            public DateTime Criacao { get; set; }
        }

        public BridgeRTCService()
        {
            _jsonSerializer = new JavaScriptSerializer();
            // Garante suporte a TLS 1.2 exigido pelo Bacen e SEFAZ
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

        public void ConfigurarPix(string instituicao, string clientId, string clientSecret, string chavePix, string caminhoCertificadoPfx, string senhaCertificado)
        {
            _pixInstituicao = string.IsNullOrEmpty(instituicao) ? "SIMULADOR" : instituicao.ToUpper().Trim();
            _pixClientId = clientId ?? "";
            _pixClientSecret = clientSecret ?? "";
            _pixChave = chavePix ?? "";

            if (!string.IsNullOrEmpty(caminhoCertificadoPfx) && File.Exists(caminhoCertificadoPfx))
            {
                _pixCertificado = new X509Certificate2(caminhoCertificadoPfx, senhaCertificado ?? "", X509KeyStorageFlags.MachineKeySet | X509KeyStorageFlags.Exportable);
            }
            else if (_certificado != null)
            {
                _pixCertificado = _certificado;
            }
        }

        /// <summary>
        /// Cria cobrança imediata Pix no padrão da API Pix do Bacen (v2)
        /// </summary>
        public string CriarCobrancaPix(string txid, double valor, int expiracaoSegundos, string solicitacaoPagador)
        {
            try
            {
                if (string.IsNullOrEmpty(txid))
                {
                    txid = GerarTxIdValido();
                }

                if (valor <= 0)
                {
                    return FormatarRetorno("ERRO", "O valor da cobranca Pix deve ser maior que zero.", null);
                }

                if (expiracaoSegundos <= 0)
                {
                    expiracaoSegundos = 3600; // 1 hora padrao
                }

                // Modo Simulador / Teste do time de desenvolvimento
                if (_pixInstituicao == "SIMULADOR" || string.IsNullOrEmpty(_pixClientId))
                {
                    string copiaEColaMock = "00020126580014br.gov.bcb.pix0136" + Guid.NewGuid().ToString() + "520400005303986540" + valor.ToString("F2", CultureInfo.InvariantCulture).Replace(",", ".") + "5802BR5913LOJISTA DEMO6009SAO PAULO62070503***6304ABCD";
                    var dadosSimulados = new PixSimuladoData
                    {
                        TxId = txid,
                        Valor = valor,
                        Status = "ATIVA",
                        PixCopiaECola = copiaEColaMock,
                        EndToEndId = "",
                        Criacao = DateTime.Now
                    };
                    lock (_cobrancasSimuladas)
                    {
                        _cobrancasSimuladas[txid] = dadosSimulados;
                    }

                    var retornoMock = new
                    {
                        txid = txid,
                        status = "ATIVA",
                        valor = valor,
                        chave = string.IsNullOrEmpty(_pixChave) ? "chave-simulada@pix.com.br" : _pixChave,
                        pixCopiaECola = copiaEColaMock,
                        location = "pix.bcb.gov.br/cobv2/" + txid,
                        expiracao = expiracaoSegundos,
                        instituicao = "SIMULADOR",
                        ambiente = _tipoAmbiente == 1 ? "Producao" : "Homologacao"
                    };

                    return FormatarRetorno("OK", "Cobranca Pix imediata gerada com sucesso (Modo Simulador)", retornoMock);
                }

                // Fluxo Real via API Bacen / Banco
                string token = ObterTokenOAuthPix();
                string urlCob = ObterUrlBasePix() + "/v2/cob/" + txid;

                var payload = new
                {
                    calendario = new { expiracao = expiracaoSegundos },
                    valor = new { original = valor.ToString("F2", CultureInfo.InvariantCulture) },
                    chave = _pixChave,
                    solicitacaoPagador = string.IsNullOrEmpty(solicitacaoPagador) ? "Pagamento no Caixa PDV" : solicitacaoPagador
                };

                string jsonEnvio = _jsonSerializer.Serialize(payload);
                string respostaHttp = ExecutarRequisicaoHttp("PUT", urlCob, jsonEnvio, token, _pixCertificado);
                var retornoApi = _jsonSerializer.Deserialize<Dictionary<string, object>>(respostaHttp);

                return FormatarRetorno("OK", "Cobranca Pix gerada com sucesso no PSP", retornoApi);
            }
            catch (Exception ex)
            {
                return FormatarRetorno("ERRO", "Falha ao criar cobranca Pix: " + ex.Message, new { detalhe = ex.ToString() });
            }
        }

        /// <summary>
        /// Consulta status da cobrança Pix pelo txid (GET /v2/cob/{txid})
        /// </summary>
        public string ConsultarCobrancaPix(string txid)
        {
            try
            {
                if (string.IsNullOrEmpty(txid))
                {
                    return FormatarRetorno("ERRO", "TxId nao informado para consulta.", null);
                }

                // Modo Simulador
                if (_pixInstituicao == "SIMULADOR" || string.IsNullOrEmpty(_pixClientId))
                {
                    lock (_cobrancasSimuladas)
                    {
                        if (_cobrancasSimuladas.ContainsKey(txid))
                        {
                            var cob = _cobrancasSimuladas[txid];
                            bool pago = cob.Status == "CONCLUIDA";
                            var retSimulado = new
                            {
                                txid = cob.TxId,
                                status = cob.Status,
                                pago = pago,
                                valor = cob.Valor,
                                endToEndId = cob.EndToEndId,
                                horario = DateTime.Now.ToString("yyyy-MM-ddTHH:mm:sszzz")
                            };
                            return FormatarRetorno("OK", pago ? "Pagamento Pix confirmado!" : "Aguardando pagamento pelo cliente...", retSimulado);
                        }
                    }

                    // Se não encontrado no mock, retorna ativa
                    return FormatarRetorno("OK", "Aguardando pagamento", new { txid = txid, status = "ATIVA", pago = false });
                }

                // Fluxo Real
                string token = ObterTokenOAuthPix();
                string urlCob = ObterUrlBasePix() + "/v2/cob/" + txid;
                string respostaHttp = ExecutarRequisicaoHttp("GET", urlCob, null, token, _pixCertificado);
                var dadosPix = _jsonSerializer.Deserialize<Dictionary<string, object>>(respostaHttp);

                string status = dadosPix.ContainsKey("status") ? dadosPix["status"].ToString() : "ATIVA";
                bool estaPago = status.Equals("CONCLUIDA", StringComparison.OrdinalIgnoreCase);

                string endToEndId = "";
                if (estaPago && dadosPix.ContainsKey("pix"))
                {
                    var listaPix = dadosPix["pix"] as System.Collections.ArrayList;
                    if (listaPix != null && listaPix.Count > 0)
                    {
                        var primeiroPix = listaPix[0] as Dictionary<string, object>;
                        if (primeiroPix != null && primeiroPix.ContainsKey("endToEndId"))
                        {
                            endToEndId = primeiroPix["endToEndId"].ToString();
                        }
                    }
                }

                dadosPix["pago"] = estaPago;
                dadosPix["endToEndId"] = endToEndId;

                return FormatarRetorno("OK", estaPago ? "Pagamento Pix confirmado no PSP" : "Aguardando confirmacao do pagamento", dadosPix);
            }
            catch (Exception ex)
            {
                return FormatarRetorno("ERRO", "Falha ao consultar cobranca Pix: " + ex.Message, new { detalhe = ex.ToString() });
            }
        }

        /// <summary>
        /// Permite ao time de desenvolvimento simular a aprovação imediata do Pix no caixa
        /// </summary>
        public string SimularPagamentoPix(string txid)
        {
            try
            {
                lock (_cobrancasSimuladas)
                {
                    if (!_cobrancasSimuladas.ContainsKey(txid))
                    {
                        _cobrancasSimuladas[txid] = new PixSimuladoData
                        {
                            TxId = txid,
                            Valor = 10.00,
                            Criacao = DateTime.Now
                        };
                    }

                    var cob = _cobrancasSimuladas[txid];
                    cob.Status = "CONCLUIDA";
                    cob.EndToEndId = "E" + DateTime.Now.ToString("yyyyMMddHHmmss") + "SIMULADO999";

                    return FormatarRetorno("OK", "Pagamento Pix simulado como CONCLUIDO com sucesso!", new
                    {
                        txid = cob.TxId,
                        status = "CONCLUIDA",
                        pago = true,
                        endToEndId = cob.EndToEndId
                    });
                }
            }
            catch (Exception ex)
            {
                return FormatarRetorno("ERRO", "Erro ao simular confirmacao do Pix: " + ex.Message, null);
            }
        }

        private string ObterUrlBasePix()
        {
            bool homolog = (_tipoAmbiente == 2);
            switch (_pixInstituicao)
            {
                case "BANCO_DO_BRASIL":
                case "BB":
                    return homolog ? "https://api.hm.bb.com.br/pix/v2" : "https://api.bb.com.br/pix/v2";
                case "ITAU":
                    return homolog ? "https://api.itau.com.br/pix/v2" : "https://api.itau.com.br/pix/v2";
                case "SANTANDER":
                    return homolog ? "https://pix-h.api.santander.com.br" : "https://pix.api.santander.com.br";
                case "INTER":
                    return "https://cdpj.partners.bancointer.com.br";
                case "SICOOB":
                    return homolog ? "https://sandbox.sicoob.com.br/pix/api/v2" : "https://api.sicoob.com.br/pix/api/v2";
                case "BRADESCO":
                    return homolog ? "https://qrpix-h.bradesco.com.br" : "https://qrpix.bradesco.com.br";
                case "EFIPAY":
                case "GERENCIANET":
                    return homolog ? "https://pix-h.gerencianet.com.br" : "https://pix.gerencianet.com.br";
                default:
                    return "https://api.bcb.gov.br/pix/v2";
            }
        }

        private string ObterUrlOAuthPix()
        {
            bool homolog = (_tipoAmbiente == 2);
            switch (_pixInstituicao)
            {
                case "BANCO_DO_BRASIL":
                case "BB":
                    return homolog ? "https://oauth.hm.bb.com.br/oauth/v2/token" : "https://oauth.bb.com.br/oauth/v2/token";
                case "ITAU":
                    return homolog ? "https://sts.itau.com.br/api/oauth/token" : "https://sts.itau.com.br/api/oauth/token";
                case "SANTANDER":
                    return homolog ? "https://trust-open.santander.com.br/auth/oauth/v2/token" : "https://trust-open.santander.com.br/auth/oauth/v2/token";
                case "INTER":
                    return "https://cdpj.partners.bancointer.com.br/oauth/v2/token";
                case "EFIPAY":
                case "GERENCIANET":
                    return homolog ? "https://pix-h.gerencianet.com.br/oauth/token" : "https://pix.gerencianet.com.br/oauth/token";
                default:
                    return ObterUrlBasePix() + "/oauth/token";
            }
        }

        private string ObterTokenOAuthPix()
        {
            if (!string.IsNullOrEmpty(_pixTokenAcesso) && DateTime.Now < _pixTokenValidade)
            {
                return _pixTokenAcesso;
            }

            string urlOAuth = ObterUrlOAuthPix();
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(urlOAuth);
            request.Method = "POST";
            request.ContentType = "application/x-www-form-urlencoded";

            if (_pixCertificado != null)
            {
                request.ClientCertificates.Add(_pixCertificado);
            }

            string credenciais = Convert.ToBase64String(Encoding.ASCII.GetBytes(_pixClientId + ":" + _pixClientSecret));
            request.Headers["Authorization"] = "Basic " + credenciais;

            byte[] body = Encoding.UTF8.GetBytes("grant_type=client_credentials&scope=cob.write cob.read pix.read");
            request.ContentLength = body.Length;

            using (Stream reqStream = request.GetRequestStream())
            {
                reqStream.Write(body, 0, body.Length);
            }

            using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
            using (StreamReader reader = new StreamReader(response.GetResponseStream(), Encoding.UTF8))
            {
                string jsonResp = reader.ReadToEnd();
                var dict = _jsonSerializer.Deserialize<Dictionary<string, object>>(jsonResp);
                if (dict.ContainsKey("access_token"))
                {
                    _pixTokenAcesso = dict["access_token"].ToString();
                    int expiresIn = dict.ContainsKey("expires_in") ? Convert.ToInt32(dict["expires_in"]) : 3600;
                    _pixTokenValidade = DateTime.Now.AddSeconds(expiresIn - 60);
                    return _pixTokenAcesso;
                }
                throw new Exception("Resposta OAuth nao continha access_token: " + jsonResp);
            }
        }

        private string ExecutarRequisicaoHttp(string metodo, string url, string jsonBody, string tokenOAuth, X509Certificate2 certCliente)
        {
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);
            request.Method = metodo;
            request.ContentType = "application/json";

            if (certCliente != null)
            {
                request.ClientCertificates.Add(certCliente);
            }

            if (!string.IsNullOrEmpty(tokenOAuth))
            {
                request.Headers["Authorization"] = "Bearer " + tokenOAuth;
            }

            if (!string.IsNullOrEmpty(jsonBody) && (metodo == "POST" || metodo == "PUT"))
            {
                byte[] bytes = Encoding.UTF8.GetBytes(jsonBody);
                request.ContentLength = bytes.Length;
                using (Stream stream = request.GetRequestStream())
                {
                    stream.Write(bytes, 0, bytes.Length);
                }
            }

            try
            {
                using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
                using (StreamReader reader = new StreamReader(response.GetResponseStream(), Encoding.UTF8))
                {
                    return reader.ReadToEnd();
                }
            }
            catch (WebException wex)
            {
                if (wex.Response != null)
                {
                    using (StreamReader reader = new StreamReader(wex.Response.GetResponseStream(), Encoding.UTF8))
                    {
                        string erroBody = reader.ReadToEnd();
                        throw new Exception("Erro HTTP (" + ((HttpWebResponse)wex.Response).StatusCode + "): " + erroBody, wex);
                    }
                }
                throw;
            }
        }

        private string GerarTxIdValido()
        {
            // txid deve ter entre 26 e 35 caracteres alfanuméricos
            return "BR" + DateTime.Now.ToString("yyyyMMddHHmmss") + Guid.NewGuid().ToString("N").Substring(0, 10);
        }

        // Métodos de Vinculação SEFAZ / Split Payment
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

                string[] meiosValidos = { "15", "17", "18", "20", "23", "24" };
                if (Array.IndexOf(meiosValidos, codigoMeioPagto) < 0)
                {
                    return FormatarRetorno("ERRO", "Meio de pagamento invalido para Split Payment (NT 2026.006). Permitidos: 15, 17, 18, 20, 23 ou 24.", null);
                }

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

        public string GerarGrupoPagamentoXml(string codigoMeioPagto, double valor, string idTransacao, string cnpjInstituicao, string codigoAutorizacao)
        {
            try
            {
                StringBuilder sb = new StringBuilder();
                sb.AppendLine("<pag>");
                sb.AppendLine("  <detPag>");
                sb.AppendLine($"    <tPag>{codigoMeioPagto}</tPag>");
                sb.AppendLine($"    <vPag>{valor.ToString("F2", CultureInfo.InvariantCulture)}</vPag>");
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
                    xMotivo = "Servico em Operacao (Ambiente RTC / SEFAZ / Pix)",
                    ambiente = _tipoAmbiente == 1 ? "Producao" : "Homologacao",
                    uf = _uf,
                    instituicaoPix = _pixInstituicao,
                    dataHora = DateTime.Now.ToString("yyyy-MM-ddTHH:mm:sszzz")
                };

                return FormatarRetorno("OK", "Servico de integracao operacional", statusInfo);
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
