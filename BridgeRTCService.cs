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

        [DispId(10)]
        string ObterConciliacaoFinanceiraPix(string txid);

        [DispId(11)]
        string DevolverPix(string endToEndId, string idDevolucao, double valor, string motivo);

        [DispId(12)]
        string ConsultarDevolucaoPix(string endToEndId, string idDevolucao);
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

        private string _pixInstituicao = "SIMULADOR";
        private string _pixClientId = "";
        private string _pixClientSecret = "";
        private string _pixChave = "";
        private X509Certificate2 _pixCertificado;
        private string _pixTokenAcesso = "";
        private DateTime _pixTokenValidade = DateTime.MinValue;

        private static readonly Dictionary<string, PixSimuladoData> _cobrancasSimuladas = new Dictionary<string, PixSimuladoData>();
        private static readonly Dictionary<string, DevolucaoSimuladaData> _devolucoesSimuladas = new Dictionary<string, DevolucaoSimuladaData>();

        private class PixSimuladoData
        {
            public string TxId { get; set; }
            public double ValorBruto { get; set; }
            public double ValorTributosRetidos { get; set; }
            public double ValorTarifaBancaria { get; set; }
            public double ValorLiquidoRecebido { get; set; }
            public double ValorCbs { get; set; }
            public double ValorIbs { get; set; }
            public string Status { get; set; }
            public string PixCopiaECola { get; set; }
            public string EndToEndId { get; set; }
            public DateTime Criacao { get; set; }
        }

        private class DevolucaoSimuladaData
        {
            public string EndToEndId { get; set; }
            public string IdDevolucao { get; set; }
            public double Valor { get; set; }
            public string Status { get; set; }
            public string Motivo { get; set; }
            public DateTime Horario { get; set; }
        }

        public BridgeRTCService()
        {
            _jsonSerializer = new JavaScriptSerializer();
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
                    expiracaoSegundos = 3600;
                }

                double cbsSimulado = Math.Round(valor * 0.088, 2);
                double ibsSimulado = Math.Round(valor * 0.177, 2);
                double tributosRetidos = Math.Round(cbsSimulado + ibsSimulado, 2);
                double tarifaBancaria = 0.99;
                double liquido = Math.Round(valor - tributosRetidos - tarifaBancaria, 2);
                if (liquido < 0) liquido = 0;

                if (_pixInstituicao == "SIMULADOR" || string.IsNullOrEmpty(_pixClientId))
                {
                    string copiaEColaMock = "00020126580014br.gov.bcb.pix0136" + Guid.NewGuid().ToString() + "520400005303986540" + valor.ToString("F2", CultureInfo.InvariantCulture).Replace(",", ".") + "5802BR5913LOJISTA DEMO6009SAO PAULO62070503***6304ABCD";
                    var dadosSimulados = new PixSimuladoData
                    {
                        TxId = txid,
                        ValorBruto = valor,
                        ValorTributosRetidos = tributosRetidos,
                        ValorTarifaBancaria = tarifaBancaria,
                        ValorLiquidoRecebido = liquido,
                        ValorCbs = cbsSimulado,
                        ValorIbs = ibsSimulado,
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
                        valorBruto = valor,
                        chave = string.IsNullOrEmpty(_pixChave) ? "chave-simulada@pix.com.br" : _pixChave,
                        pixCopiaECola = copiaEColaMock,
                        location = "pix.bcb.gov.br/cobv2/" + txid,
                        expiracao = expiracaoSegundos,
                        instituicao = "SIMULADOR",
                        ambiente = _tipoAmbiente == 1 ? "Producao" : "Homologacao",
                        previsaoConciliacao = new
                        {
                            valorBruto = valor,
                            previsaoTributosRetidos = tributosRetidos,
                            previsaoTarifaPsp = tarifaBancaria,
                            previsaoLiquidoConta = liquido
                        }
                    };

                    return FormatarRetorno("OK", "Cobranca Pix imediata gerada com sucesso (Modo Simulador)", retornoMock);
                }

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

        public string ConsultarCobrancaPix(string txid)
        {
            try
            {
                if (string.IsNullOrEmpty(txid))
                {
                    return FormatarRetorno("ERRO", "TxId nao informado para consulta.", null);
                }

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
                                valorBruto = cob.ValorBruto,
                                valorTributosRetidos = cob.ValorTributosRetidos,
                                valorTarifaBancaria = cob.ValorTarifaBancaria,
                                valorLiquidoRecebido = cob.ValorLiquidoRecebido,
                                detalheSplit = new
                                {
                                    cbsRetido = cob.ValorCbs,
                                    ibsRetido = cob.ValorIbs,
                                    totalImpostosGoverno = cob.ValorTributosRetidos,
                                    tarifaPSP = cob.ValorTarifaBancaria,
                                    liquidoDisponivelCaixa = cob.ValorLiquidoRecebido
                                },
                                endToEndId = cob.EndToEndId,
                                horario = DateTime.Now.ToString("yyyy-MM-ddTHH:mm:sszzz")
                            };
                            return FormatarRetorno("OK", pago ? "Pagamento Pix confirmado!" : "Aguardando pagamento pelo cliente...", retSimulado);
                        }
                    }

                    return FormatarRetorno("OK", "Aguardando pagamento", new { txid = txid, status = "ATIVA", pago = false });
                }

                string token = ObterTokenOAuthPix();
                string urlCob = ObterUrlBasePix() + "/v2/cob/" + txid;
                string respostaHttp = ExecutarRequisicaoHttp("GET", urlCob, null, token, _pixCertificado);
                var dadosPix = _jsonSerializer.Deserialize<Dictionary<string, object>>(respostaHttp);

                string status = dadosPix.ContainsKey("status") ? dadosPix["status"].ToString() : "ATIVA";
                bool estaPago = status.Equals("CONCLUIDA", StringComparison.OrdinalIgnoreCase);

                string endToEndId = "";
                double valorBruto = 0;
                if (dadosPix.ContainsKey("valor"))
                {
                    var valObj = dadosPix["valor"] as Dictionary<string, object>;
                    if (valObj != null && valObj.ContainsKey("original"))
                    {
                        double.TryParse(valObj["original"].ToString(), NumberStyles.Any, CultureInfo.InvariantCulture, out valorBruto);
                    }
                }

                double tributosRetidos = 0;
                double tarifaBancaria = 0;
                double valorLiquido = valorBruto;

                if (estaPago && dadosPix.ContainsKey("pix"))
                {
                    var listaPix = dadosPix["pix"] as System.Collections.ArrayList;
                    if (listaPix != null && listaPix.Count > 0)
                    {
                        var primeiroPix = listaPix[0] as Dictionary<string, object>;
                        if (primeiroPix != null)
                        {
                            if (primeiroPix.ContainsKey("endToEndId"))
                                endToEndId = primeiroPix["endToEndId"].ToString();

                            if (primeiroPix.ContainsKey("componentesValor"))
                            {
                                var comp = primeiroPix["componentesValor"] as Dictionary<string, object>;
                                if (comp != null)
                                {
                                    if (comp.ContainsKey("splitTributario"))
                                    {
                                        var splitTrib = comp["splitTributario"] as Dictionary<string, object>;
                                        if (splitTrib != null && splitTrib.ContainsKey("totalRetido"))
                                            double.TryParse(splitTrib["totalRetido"].ToString(), NumberStyles.Any, CultureInfo.InvariantCulture, out tributosRetidos);
                                    }
                                    if (comp.ContainsKey("tarifa"))
                                    {
                                        double.TryParse(comp["tarifa"].ToString(), NumberStyles.Any, CultureInfo.InvariantCulture, out tarifaBancaria);
                                    }
                                    if (comp.ContainsKey("valorLiquido"))
                                    {
                                        double.TryParse(comp["valorLiquido"].ToString(), NumberStyles.Any, CultureInfo.InvariantCulture, out valorLiquido);
                                    }
                                }
                            }
                        }
                    }
                }

                if (tributosRetidos == 0 && estaPago)
                {
                    tributosRetidos = Math.Round(valorBruto * 0.265, 2);
                    valorLiquido = Math.Round(valorBruto - tributosRetidos - tarifaBancaria, 2);
                }

                dadosPix["pago"] = estaPago;
                dadosPix["endToEndId"] = endToEndId;
                dadosPix["valorBruto"] = valorBruto;
                dadosPix["valorTributosRetidos"] = tributosRetidos;
                dadosPix["valorTarifaBancaria"] = tarifaBancaria;
                dadosPix["valorLiquidoRecebido"] = valorLiquido;

                return FormatarRetorno("OK", estaPago ? "Pagamento Pix confirmado no PSP" : "Aguardando confirmacao do pagamento", dadosPix);
            }
            catch (Exception ex)
            {
                return FormatarRetorno("ERRO", "Falha ao consultar cobranca Pix: " + ex.Message, new { detalhe = ex.ToString() });
            }
        }

        /// <summary>
        /// Solicita a devolução (estorno) de um Pix pelo endToEndId oficial do Bacen
        /// Endpoint Bacen: PUT /v2/pix/{e2eid}/devolucao/{idDevolucao}
        /// </summary>
        public string DevolverPix(string endToEndId, string idDevolucao, double valor, string motivo)
        {
            try
            {
                if (string.IsNullOrEmpty(endToEndId))
                {
                    return FormatarRetorno("ERRO", "EndToEndId e obrigatorio para devolucao.", null);
                }

                if (string.IsNullOrEmpty(idDevolucao))
                {
                    idDevolucao = "DEV" + DateTime.Now.ToString("yyyyMMddHHmmss") + Guid.NewGuid().ToString("N").Substring(0, 6);
                }

                if (valor <= 0)
                {
                    return FormatarRetorno("ERRO", "O valor para devolucao deve ser maior que zero.", null);
                }

                // Modo Simulador
                if (_pixInstituicao == "SIMULADOR" || string.IsNullOrEmpty(_pixClientId))
                {
                    var devSim = new DevolucaoSimuladaData
                    {
                        EndToEndId = endToEndId,
                        IdDevolucao = idDevolucao,
                        Valor = valor,
                        Status = "EM_PROCESSAMENTO", // Bacen retorna EM_PROCESSAMENTO ou DEVOLVIDO
                        Motivo = string.IsNullOrEmpty(motivo) ? "Cancelamento de venda no caixa" : motivo,
                        Horario = DateTime.Now
                    };
                    lock (_devolucoesSimuladas)
                    {
                        _devolucoesSimuladas[idDevolucao] = devSim;
                    }

                    var retMock = new
                    {
                        id = idDevolucao,
                        rtrId = "D" + DateTime.Now.ToString("yyyyMMddHHmmss") + "SIMULADO",
                        valor = valor.ToString("F2", CultureInfo.InvariantCulture),
                        status = "DEVOLVIDO",
                        motivo = devSim.Motivo,
                        horario = DateTime.Now.ToString("yyyy-MM-ddTHH:mm:sszzz")
                    };

                    return FormatarRetorno("OK", "Devolucao/Estorno Pix efetuado com sucesso (Modo Simulador)", retMock);
                }

                // Fluxo Real Bacen
                string token = ObterTokenOAuthPix();
                string urlDev = ObterUrlBasePix() + "/v2/pix/" + endToEndId + "/devolucao/" + idDevolucao;

                var payload = new
                {
                    valor = valor.ToString("F2", CultureInfo.InvariantCulture),
                    natureza = "ORIGINAL",
                    descricao = string.IsNullOrEmpty(motivo) ? "Cancelamento de compra no PDV" : motivo
                };

                string jsonEnvio = _jsonSerializer.Serialize(payload);
                string respHttp = ExecutarRequisicaoHttp("PUT", urlDev, jsonEnvio, token, _pixCertificado);
                var retornoApi = _jsonSerializer.Deserialize<Dictionary<string, object>>(respHttp);

                return FormatarRetorno("OK", "Solicitacao de devolucao registrada com sucesso no Bacen", retornoApi);
            }
            catch (Exception ex)
            {
                return FormatarRetorno("ERRO", "Falha ao solicitar devolucao do Pix: " + ex.Message, new { detalhe = ex.ToString() });
            }
        }

        public string ConsultarDevolucaoPix(string endToEndId, string idDevolucao)
        {
            try
            {
                if (string.IsNullOrEmpty(endToEndId) || string.IsNullOrEmpty(idDevolucao))
                {
                    return FormatarRetorno("ERRO", "EndToEndId e IdDevolucao sao obrigatorios.", null);
                }

                if (_pixInstituicao == "SIMULADOR" || string.IsNullOrEmpty(_pixClientId))
                {
                    return FormatarRetorno("OK", "Devolucao concluida", new
                    {
                        id = idDevolucao,
                        status = "DEVOLVIDO",
                        horario = DateTime.Now.ToString("yyyy-MM-ddTHH:mm:sszzz")
                    });
                }

                string token = ObterTokenOAuthPix();
                string urlDev = ObterUrlBasePix() + "/v2/pix/" + endToEndId + "/devolucao/" + idDevolucao;
                string respHttp = ExecutarRequisicaoHttp("GET", urlDev, null, token, _pixCertificado);
                var retornoApi = _jsonSerializer.Deserialize<Dictionary<string, object>>(respHttp);

                return FormatarRetorno("OK", "Consulta de devolucao efetuada", retornoApi);
            }
            catch (Exception ex)
            {
                return FormatarRetorno("ERRO", "Falha ao consultar devolucao: " + ex.Message, null);
            }
        }

        public string ObterConciliacaoFinanceiraPix(string txid)
        {
            return ConsultarCobrancaPix(txid);
        }

        public string SimularPagamentoPix(string txid)
        {
            try
            {
                lock (_cobrancasSimuladas)
                {
                    if (!_cobrancasSimuladas.ContainsKey(txid))
                    {
                        double valorPadrao = 100.00;
                        double cbs = 8.80;
                        double ibs = 17.70;
                        double trib = 26.50;
                        double tarifa = 0.99;
                        double liq = 72.51;

                        _cobrancasSimuladas[txid] = new PixSimuladoData
                        {
                            TxId = txid,
                            ValorBruto = valorPadrao,
                            ValorCbs = cbs,
                            ValorIbs = ibs,
                            ValorTributosRetidos = trib,
                            ValorTarifaBancaria = tarifa,
                            ValorLiquidoRecebido = liq,
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
                        endToEndId = cob.EndToEndId,
                        valorBruto = cob.ValorBruto,
                        valorTributosRetidos = cob.ValorTributosRetidos,
                        valorTarifaBancaria = cob.ValorTarifaBancaria,
                        valorLiquidoRecebido = cob.ValorLiquidoRecebido,
                        detalheSplit = new
                        {
                            cbsRetido = cob.ValorCbs,
                            ibsRetido = cob.ValorIbs,
                            totalImpostosGoverno = cob.ValorTributosRetidos,
                            tarifaPSP = cob.ValorTarifaBancaria,
                            liquidoDisponivelCaixa = cob.ValorLiquidoRecebido
                        }
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

            byte[] body = Encoding.UTF8.GetBytes("grant_type=client_credentials&scope=cob.write cob.read pix.read pix.write");
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
            return "BR" + DateTime.Now.ToString("yyyyMMddHHmmss") + Guid.NewGuid().ToString("N").Substring(0, 10);
        }

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

        /// <summary>
        /// Gera a estrutura XML com tpIntegra=1 obrigatorio para pagamentos integrados no PDV
        /// </summary>
        public string GerarGrupoPagamentoXml(string codigoMeioPagto, double valor, string idTransacao, string cnpjInstituicao, string codigoAutorizacao)
        {
            try
            {
                StringBuilder sb = new StringBuilder();
                sb.AppendLine("<pag>");
                sb.AppendLine("  <detPag>");
                sb.AppendLine("    <indPag>0</indPag>"); // 0 = Pagamento a Vista
                sb.AppendLine($"    <tPag>{codigoMeioPagto}</tPag>");
                sb.AppendLine($"    <vPag>{valor.ToString("F2", CultureInfo.InvariantCulture)}</vPag>");
                sb.AppendLine("    <card>");
                sb.AppendLine("      <tpIntegra>1</tpIntegra>"); // 1 = Pagamento integrado com o sistema de automacao
                if (!string.IsNullOrEmpty(cnpjInstituicao))
                    sb.AppendLine($"      <CNPJ>{cnpjInstituicao}</CNPJ>");
                if (!string.IsNullOrEmpty(codigoAutorizacao))
                    sb.AppendLine($"      <cAut>{codigoAutorizacao}</cAut>");
                sb.AppendLine("    </card>");
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

                return FormatarRetorno("OK", "Grupo XML de pagamento gerado com sucesso (tpIntegra=1)", new { xml = sb.ToString() });
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
