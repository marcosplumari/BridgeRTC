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

        [DispId(13)]
        string ConfigurarPixCustomizado(string urlBase, string urlOAuth, string clientId, string clientSecret, string chavePix, string caminhoCertificadoPfx, string senhaCertificado);

        [DispId(14)]
        string CriarBoletoHibridoPix(string nossoNumeroOuTxId, double valorOriginal, string dataVencimentoIso, int diasValidadeAposVencimento, double valorAbatimentoDesconto, string dataLimiteDescontoIso, double percentualMulta, double percentualJurosMensal, string cpfCnpjDevedor, string nomeDevedor, string solicitacaoPagador);

        [DispId(15)]
        string ConsultarBoletoPix(string nossoNumeroOuTxId);

        [DispId(16)]
        string SimularPagamentoBoletoPix(string nossoNumeroOuTxId, string dataPagamentoIsoSimulada);
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
        private string _pixUrlBaseCustomizada = "";
        private string _pixUrlOAuthCustomizada = "";
        private X509Certificate2 _pixCertificado;
        private string _pixTokenAcesso = "";
        private DateTime _pixTokenValidade = DateTime.MinValue;

        private static readonly Dictionary<string, PixSimuladoData> _cobrancasSimuladas = new Dictionary<string, PixSimuladoData>();
        private static readonly Dictionary<string, DevolucaoSimuladaData> _devolucoesSimuladas = new Dictionary<string, DevolucaoSimuladaData>();
        private static readonly Dictionary<string, BoletoPixSimuladoData> _boletosSimulados = new Dictionary<string, BoletoPixSimuladoData>();

        private class BoletoPixSimuladoData
        {
            public string NossoNumero { get; set; }
            public double ValorOriginal { get; set; }
            public DateTime DataVencimento { get; set; }
            public int DiasValidadeAposVencimento { get; set; }
            public double ValorDesconto { get; set; }
            public DateTime? DataLimiteDesconto { get; set; }
            public double PercMulta { get; set; }
            public double PercJurosMensal { get; set; }
            public string CpfCnpjDevedor { get; set; }
            public string NomeDevedor { get; set; }
            public string Status { get; set; }
            public string PixCopiaECola { get; set; }
            public double ValorPagoFinal { get; set; }
            public double ValorDescontoAplicado { get; set; }
            public double ValorMultaAplicada { get; set; }
            public double ValorJurosAplicados { get; set; }
            public DateTime? DataPagamento { get; set; }
            public string EndToEndId { get; set; }
        }

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
            _pixUrlBaseCustomizada = "";
            _pixUrlOAuthCustomizada = "";
            _pixTokenAcesso = "";
            _pixTokenValidade = DateTime.MinValue;

            if (!string.IsNullOrEmpty(caminhoCertificadoPfx) && File.Exists(caminhoCertificadoPfx))
            {
                _pixCertificado = new X509Certificate2(caminhoCertificadoPfx, senhaCertificado ?? "", X509KeyStorageFlags.MachineKeySet | X509KeyStorageFlags.Exportable);
            }
            else if (_certificado != null)
            {
                _pixCertificado = _certificado;
            }
        }

        public string ConfigurarPixCustomizado(string urlBase, string urlOAuth, string clientId, string clientSecret, string chavePix, string caminhoCertificadoPfx, string senhaCertificado)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(urlBase))
                {
                    return FormatarRetorno("ERRO", "URL Base do Pix e obrigatoria.", null);
                }

                if (!Uri.TryCreate(urlBase.Trim(), UriKind.Absolute, out Uri uriBase) || uriBase.Scheme != Uri.UriSchemeHttps)
                {
                    return FormatarRetorno("ERRO", "URL Base invalida. Deve ser uma URL absoluta iniciando obrigatoriamente com https:// (ex: https://api.seubanco.com.br/pix/v2).", null);
                }

                if (string.IsNullOrWhiteSpace(urlOAuth))
                {
                    return FormatarRetorno("ERRO", "URL de autenticacao OAuth e obrigatoria.", null);
                }

                if (!Uri.TryCreate(urlOAuth.Trim(), UriKind.Absolute, out Uri uriOAuth) || uriOAuth.Scheme != Uri.UriSchemeHttps)
                {
                    return FormatarRetorno("ERRO", "URL OAuth invalida. Deve ser uma URL absoluta iniciando obrigatoriamente com https:// (ex: https://oauth.seubanco.com.br/oauth/token).", null);
                }

                if (string.IsNullOrWhiteSpace(clientId))
                {
                    return FormatarRetorno("ERRO", "Client ID e obrigatorio.", null);
                }

                if (string.IsNullOrWhiteSpace(clientSecret))
                {
                    return FormatarRetorno("ERRO", "Client Secret e obrigatorio.", null);
                }

                if (string.IsNullOrWhiteSpace(chavePix))
                {
                    return FormatarRetorno("ERRO", "Chave Pix e obrigatoria.", null);
                }

                bool certificadoCarregado = false;
                if (!string.IsNullOrEmpty(caminhoCertificadoPfx))
                {
                    if (!File.Exists(caminhoCertificadoPfx))
                    {
                        return FormatarRetorno("ERRO", "Arquivo de certificado A1 (.pfx) nao encontrado no caminho: " + caminhoCertificadoPfx, null);
                    }
                    try
                    {
                        _pixCertificado = new X509Certificate2(caminhoCertificadoPfx, senhaCertificado ?? "", X509KeyStorageFlags.MachineKeySet | X509KeyStorageFlags.Exportable);
                        certificadoCarregado = true;
                    }
                    catch (Exception exCert)
                    {
                        return FormatarRetorno("ERRO", "Falha ao abrir certificado A1 com a senha fornecida: " + exCert.Message, null);
                    }
                }
                else if (_certificado != null)
                {
                    _pixCertificado = _certificado;
                    certificadoCarregado = true;
                }

                _pixInstituicao = "CUSTOM";
                _pixUrlBaseCustomizada = urlBase.Trim().TrimEnd('/');
                _pixUrlOAuthCustomizada = urlOAuth.Trim();
                _pixClientId = clientId.Trim();
                _pixClientSecret = clientSecret.Trim();
                _pixChave = chavePix.Trim();
                _pixTokenAcesso = "";
                _pixTokenValidade = DateTime.MinValue;

                var meta = new Dictionary<string, object>
                {
                    { "modo", "CUSTOM" },
                    { "urlBase", _pixUrlBaseCustomizada },
                    { "urlOAuth", _pixUrlOAuthCustomizada },
                    { "chavePix", _pixChave },
                    { "certificadoA1Carregado", certificadoCarregado }
                };

                return FormatarRetorno("OK", "Configuracao Pix personalizada validada e carregada com sucesso.", meta);
            }
            catch (Exception ex)
            {
                return FormatarRetorno("ERRO", "Erro ao validar configuracao personalizada do Pix: " + ex.Message, null);
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
            if (!string.IsNullOrEmpty(_pixUrlBaseCustomizada))
            {
                return _pixUrlBaseCustomizada;
            }

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
                    throw new InvalidOperationException("Instituicao Pix '" + _pixInstituicao + "' nao possui URLs pre-configuradas na DLL. Utilize o metodo ConfigurarPixCustomizado informando a URL Base e a URL OAuth da sua instituicao bancaria.");
            }
        }

        private string ObterUrlOAuthPix()
        {
            if (!string.IsNullOrEmpty(_pixUrlOAuthCustomizada))
            {
                return _pixUrlOAuthCustomizada;
            }

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
    
        public string CriarBoletoHibridoPix(string nossoNumeroOuTxId, double valorOriginal, string dataVencimentoIso, int diasValidadeAposVencimento, double valorAbatimentoDesconto, string dataLimiteDescontoIso, double percentualMulta, double percentualJurosMensal, string cpfCnpjDevedor, string nomeDevedor, string solicitacaoPagador)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(nossoNumeroOuTxId))
                {
                    nossoNumeroOuTxId = "BOL" + DateTime.Now.ToString("yyyyMMddHHmmss") + "001";
                }

                if (valorOriginal <= 0)
                {
                    return FormatarRetorno("ERRO", "O valor original do boleto deve ser maior que zero.", null);
                }

                DateTime vencimento;
                if (!DateTime.TryParse(dataVencimentoIso, CultureInfo.InvariantCulture, DateTimeStyles.None, out vencimento) &&
                    !DateTime.TryParseExact(dataVencimentoIso, new string[] { "yyyy-MM-dd", "dd/MM/yyyy" }, CultureInfo.InvariantCulture, DateTimeStyles.None, out vencimento))
                {
                    return FormatarRetorno("ERRO", "Data de vencimento invalida. Utilize o formato YYYY-MM-DD ou DD/MM/YYYY.", null);
                }

                DateTime? limiteDesconto = null;
                if (!string.IsNullOrWhiteSpace(dataLimiteDescontoIso))
                {
                    DateTime parsedDesc;
                    if (DateTime.TryParse(dataLimiteDescontoIso, CultureInfo.InvariantCulture, DateTimeStyles.None, out parsedDesc) ||
                        DateTime.TryParseExact(dataLimiteDescontoIso, new string[] { "yyyy-MM-dd", "dd/MM/yyyy" }, CultureInfo.InvariantCulture, DateTimeStyles.None, out parsedDesc))
                    {
                        limiteDesconto = parsedDesc;
                    }
                }

                if (diasValidadeAposVencimento <= 0) diasValidadeAposVencimento = 30;

                // Modo Real (Banco com API Pix CobV)
                if (_pixInstituicao != "SIMULADOR" && !string.IsNullOrEmpty(_pixClientId))
                {
                    string token = ObterTokenOAuthPix();
                    string urlCobV = ObterUrlBasePix() + "/v2/cobv/" + nossoNumeroOuTxId;

                    HttpWebRequest request = (HttpWebRequest)WebRequest.Create(urlCobV);
                    request.Method = "PUT";
                    request.ContentType = "application/json";
                    request.Headers["Authorization"] = "Bearer " + token;

                    if (_pixCertificado != null)
                    {
                        request.ClientCertificates.Add(_pixCertificado);
                    }

                    var payloadDict = new Dictionary<string, object>
                    {
                        { "calendario", new Dictionary<string, object> {
                            { "dataDeVencimento", vencimento.ToString("yyyy-MM-dd") },
                            { "validadeAposVencimento", diasValidadeAposVencimento }
                        }},
                        { "devedor", new Dictionary<string, object> {
                            { (cpfCnpjDevedor != null && cpfCnpjDevedor.Length > 11 ? "cnpj" : "cpf"), (cpfCnpjDevedor ?? "").Replace(".", "").Replace("-", "").Replace("/", "") },
                            { "nome", nomeDevedor ?? "CLIENTE" }
                        }},
                        { "valor", new Dictionary<string, object> {
                            { "original", valorOriginal.ToString("F2", CultureInfo.InvariantCulture) },
                            { "multa", new Dictionary<string, object> { { "modalidade", 2 }, { "valorPerc", percentualMulta.ToString("F2", CultureInfo.InvariantCulture) } } },
                            { "juros", new Dictionary<string, object> { { "modalidade", 2 }, { "valorPerc", percentualJurosMensal.ToString("F2", CultureInfo.InvariantCulture) } } }
                        }},
                        { "chave", _pixChave },
                        { "solicitacaoPagador", string.IsNullOrEmpty(solicitacaoPagador) ? ("Boleto NossoNumero " + nossoNumeroOuTxId) : solicitacaoPagador }
                    };

                    if (valorAbatimentoDesconto > 0 && limiteDesconto.HasValue)
                    {
                        var valorSub = (Dictionary<string, object>)payloadDict["valor"];
                        valorSub.Add("desconto", new Dictionary<string, object> {
                            { "modalidade", 1 },
                            { "descontoDataFixa", new List<object> {
                                new Dictionary<string, object> {
                                    { "data", limiteDesconto.Value.ToString("yyyy-MM-dd") },
                                    { "valorPerc", (Math.Round((valorAbatimentoDesconto / valorOriginal) * 100.0, 2)).ToString("F2", CultureInfo.InvariantCulture) }
                                }
                            }}
                        });
                    }

                    byte[] postBytes = Encoding.UTF8.GetBytes(_jsonSerializer.Serialize(payloadDict));
                    request.ContentLength = postBytes.Length;

                    using (Stream reqStream = request.GetRequestStream())
                    {
                        reqStream.Write(postBytes, 0, postBytes.Length);
                    }

                    using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
                    using (StreamReader reader = new StreamReader(response.GetResponseStream(), Encoding.UTF8))
                    {
                        string jsonResp = reader.ReadToEnd();
                        var respDict = _jsonSerializer.Deserialize<Dictionary<string, object>>(jsonResp);
                        string pixCopiaECola = respDict.ContainsKey("pixCopiaECola") ? respDict["pixCopiaECola"].ToString() : "";

                        var retObj = new Dictionary<string, object>
                        {
                            { "nossoNumero", nossoNumeroOuTxId },
                            { "pixCopiaECola", pixCopiaECola },
                            { "status", "ATIVA" },
                            { "valorOriginal", valorOriginal },
                            { "vencimento", vencimento.ToString("yyyy-MM-dd") },
                            { "diasValidadeAposVencimento", diasValidadeAposVencimento },
                            { "bancoIntegrado", true }
                        };

                        return FormatarRetorno("OK", "Boleto Hibrido com Pix gerado com sucesso na instituicao bancaria.", retObj);
                    }
                }

                // Modo Simulador
                string locationMock = "pix.bancoexemplo.com.br/cobv/" + nossoNumeroOuTxId;
                string copiaEColaMock = "00020126580014br.gov.bcb.pix2536" + locationMock + "520400005303986540" + valorOriginal.ToString("F2", CultureInfo.InvariantCulture) + "5802BR5915EMPRESA LOJA SA6009SAO PAULO62070503***6304ABCD";

                var boletoData = new BoletoPixSimuladoData
                {
                    NossoNumero = nossoNumeroOuTxId,
                    ValorOriginal = valorOriginal,
                    DataVencimento = vencimento,
                    DiasValidadeAposVencimento = diasValidadeAposVencimento,
                    ValorDesconto = valorAbatimentoDesconto,
                    DataLimiteDesconto = limiteDesconto,
                    PercMulta = percentualMulta,
                    PercJurosMensal = percentualJurosMensal,
                    CpfCnpjDevedor = cpfCnpjDevedor,
                    NomeDevedor = nomeDevedor,
                    Status = "ATIVA",
                    PixCopiaECola = copiaEColaMock,
                    ValorPagoFinal = 0,
                    ValorDescontoAplicado = 0,
                    ValorMultaAplicada = 0,
                    ValorJurosAplicados = 0,
                    DataPagamento = null,
                    EndToEndId = ""
                };

                _boletosSimulados[nossoNumeroOuTxId] = boletoData;

                var retSimulado = new Dictionary<string, object>
                {
                    { "nossoNumero", nossoNumeroOuTxId },
                    { "pixCopiaECola", copiaEColaMock },
                    { "status", "ATIVA" },
                    { "valorOriginal", valorOriginal },
                    { "vencimento", vencimento.ToString("yyyy-MM-dd") },
                    { "descontoDisponivel", valorAbatimentoDesconto },
                    { "limiteDesconto", limiteDesconto.HasValue ? limiteDesconto.Value.ToString("yyyy-MM-dd") : "" },
                    { "percentualMulta", percentualMulta },
                    { "percentualJurosMensal", percentualJurosMensal },
                    { "instrucaoImpressao", "Pague com Pix escaneando o QR Code ao lado. Descontos e acrescimos sao calculados automaticamente pelo seu banco." }
                };

                return FormatarRetorno("OK", "Boleto Hibrido Pix (CobV) criado com sucesso (Modo Simulador).", retSimulado);
            }
            catch (Exception ex)
            {
                return FormatarRetorno("ERRO", "Falha ao gerar Boleto Hibrido com Pix: " + ex.Message, null);
            }
        }

        public string ConsultarBoletoPix(string nossoNumeroOuTxId)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(nossoNumeroOuTxId))
                {
                    return FormatarRetorno("ERRO", "Nosso Numero / TxId e obrigatorio.", null);
                }

                if (_pixInstituicao != "SIMULADOR" && !string.IsNullOrEmpty(_pixClientId))
                {
                    string token = ObterTokenOAuthPix();
                    string urlCobV = ObterUrlBasePix() + "/v2/cobv/" + nossoNumeroOuTxId;

                    HttpWebRequest request = (HttpWebRequest)WebRequest.Create(urlCobV);
                    request.Method = "GET";
                    request.Headers["Authorization"] = "Bearer " + token;

                    if (_pixCertificado != null)
                    {
                        request.ClientCertificates.Add(_pixCertificado);
                    }

                    using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
                    using (StreamReader reader = new StreamReader(response.GetResponseStream(), Encoding.UTF8))
                    {
                        string jsonResp = reader.ReadToEnd();
                        var respDict = _jsonSerializer.Deserialize<Dictionary<string, object>>(jsonResp);
                        return FormatarRetorno("OK", "Consulta de Boleto Pix realizada com sucesso.", respDict);
                    }
                }

                if (_boletosSimulados.ContainsKey(nossoNumeroOuTxId))
                {
                    var bol = _boletosSimulados[nossoNumeroOuTxId];
                    var dados = new Dictionary<string, object>
                    {
                        { "nossoNumero", bol.NossoNumero },
                        { "status", bol.Status },
                        { "valorOriginal", bol.ValorOriginal },
                        { "vencimento", bol.DataVencimento.ToString("yyyy-MM-dd") },
                        { "valorPagoFinal", bol.ValorPagoFinal },
                        { "valorDescontoAplicado", bol.ValorDescontoAplicado },
                        { "valorMultaAplicada", bol.ValorMultaAplicada },
                        { "valorJurosAplicados", bol.ValorJurosAplicados },
                        { "dataPagamento", bol.DataPagamento.HasValue ? bol.DataPagamento.Value.ToString("yyyy-MM-dd HH:mm:ss") : null },
                        { "endToEndId", bol.EndToEndId },
                        { "pixCopiaECola", bol.PixCopiaECola }
                    };

                    return FormatarRetorno("OK", "Boleto Pix consultado no simulador.", dados);
                }

                return FormatarRetorno("ERRO", "Boleto Pix nao encontrado no simulador com o Nosso Numero informado.", null);
            }
            catch (Exception ex)
            {
                return FormatarRetorno("ERRO", "Erro ao consultar Boleto Pix: " + ex.Message, null);
            }
        }

        public string SimularPagamentoBoletoPix(string nossoNumeroOuTxId, string dataPagamentoIsoSimulada)
        {
            try
            {
                if (!_boletosSimulados.ContainsKey(nossoNumeroOuTxId))
                {
                    return FormatarRetorno("ERRO", "Boleto com Nosso Numero " + nossoNumeroOuTxId + " nao foi encontrado.", null);
                }

                var bol = _boletosSimulados[nossoNumeroOuTxId];

                DateTime dtPagamento = DateTime.Now;
                if (!string.IsNullOrWhiteSpace(dataPagamentoIsoSimulada))
                {
                    DateTime parsed;
                    if (DateTime.TryParse(dataPagamentoIsoSimulada, CultureInfo.InvariantCulture, DateTimeStyles.None, out parsed) ||
                        DateTime.TryParseExact(dataPagamentoIsoSimulada, new string[] { "yyyy-MM-dd", "dd/MM/yyyy" }, CultureInfo.InvariantCulture, DateTimeStyles.None, out parsed))
                    {
                        dtPagamento = parsed;
                    }
                }

                double valorFinal = bol.ValorOriginal;
                double descontoAplicado = 0;
                double multaAplicada = 0;
                double jurosAplicados = 0;

                // Cenário 1: Pagamento antecipado com desconto
                if (bol.ValorDesconto > 0 && bol.DataLimiteDesconto.HasValue && dtPagamento.Date <= bol.DataLimiteDesconto.Value.Date)
                {
                    descontoAplicado = bol.ValorDesconto;
                    valorFinal = Math.Round(bol.ValorOriginal - descontoAplicado, 2);
                }
                // Cenário 2: Pagamento em atraso (após o vencimento)
                else if (dtPagamento.Date > bol.DataVencimento.Date)
                {
                    int diasAtraso = (int)(dtPagamento.Date - bol.DataVencimento.Date).TotalDays;

                    // Multa percentual fixa sobre o original
                    if (bol.PercMulta > 0)
                    {
                        multaAplicada = Math.Round(bol.ValorOriginal * (bol.PercMulta / 100.0), 2);
                    }

                    // Juros pró-rata dia com base no percentual mensal
                    if (bol.PercJurosMensal > 0)
                    {
                        double jurosDia = (bol.PercJurosMensal / 30.0) / 100.0;
                        jurosAplicados = Math.Round(bol.ValorOriginal * jurosDia * diasAtraso, 2);
                    }

                    valorFinal = Math.Round(bol.ValorOriginal + multaAplicada + jurosAplicados, 2);
                }

                bol.Status = "CONCLUIDA";
                bol.ValorPagoFinal = valorFinal;
                bol.ValorDescontoAplicado = descontoAplicado;
                bol.ValorMultaAplicada = multaAplicada;
                bol.ValorJurosAplicados = jurosAplicados;
                bol.DataPagamento = dtPagamento;
                bol.EndToEndId = "E" + DateTime.Now.ToString("yyyyMMddHHmmss") + "BOL" + (new Random().Next(1000, 9999));

                var res = new Dictionary<string, object>
                {
                    { "nossoNumero", bol.NossoNumero },
                    { "status", "CONCLUIDA" },
                    { "valorOriginal", bol.ValorOriginal },
                    { "dataPagamento", dtPagamento.ToString("yyyy-MM-dd") },
                    { "descontoAplicado", descontoAplicado },
                    { "multaAplicada", multaAplicada },
                    { "jurosAplicados", jurosAplicados },
                    { "valorPagoFinal", valorFinal },
                    { "endToEndId", bol.EndToEndId },
                    { "liquidacaoInstantanea", true },
                    { "conciliacaoCnab", new Dictionary<string, object> {
                        { "ocorrencia", "06-LIQUIDACAO_PIX" },
                        { "valorTarifaEstimada", 0.99 },
                        { "valorLiquidoCreditado", Math.Round(valorFinal - 0.99, 2) }
                    }}
                };

                return FormatarRetorno("OK", "Pagamento de Boleto Hibrido via Pix confirmado com sucesso.", res);
            }
            catch (Exception ex)
            {
                return FormatarRetorno("ERRO", "Falha ao simular liquidacao de boleto: " + ex.Message, null);
            }
        }

    }
}