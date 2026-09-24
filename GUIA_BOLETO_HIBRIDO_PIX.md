# Guia Técnico: Boleto Híbrido com Pix (CobV) no ERP em VB6

Este documento orienta a equipe de desenvolvimento sobre a emissão, impressão em ActiveReports e conciliação bancária de **Boletos Híbridos** (boletos bancários tradicionais com QR Code Pix integrado).

---

## 1. O que é o Boleto Híbrido?

O **Boleto Híbrido** é a junção do boleto bancário tradicional com a agilidade do Pix:

```text
┌────────────────────────────────────────────────────────────────────────┐
│                        BOLETO BANCÁRIO HÍBRIDO                         │
├────────────────────────────────────────────────────────────────────────┤
│ Beneficiário: SUA EMPRESA LTDA                 Vencimento: 15/10/2026  │
│ Pagador: MERCADO EXEMPLO LTDA                  Valor: R$ 150,00        │
├──────────────────────────────────────┬─────────────────────────────────┤
│                                      │                                 │
│  [ QR CODE PIX DINÂMICO ]            │   Linha Digitável / Barras:     │
│                                      │   34191.79001 01043.510047 ...  │
│  "Escaneie para pagar com Pix.       │                                 │
│   Descontos e juros são calculados   │   Código de barras tradicional  │
│   automaticamente pelo seu banco."   │   ||| | ||||| || |||| |||||     │
│                                      │                                 │
└──────────────────────────────────────┴─────────────────────────────────┘
```

Quando o cliente escolhe pagar pelo **Pix**, o banco liquida o dinheiro em poucos segundos na conta da empresa e cancela o código de barras tradicional na câmara de compensação (CIP/NPC) para impedir que o boleto seja pago duas vezes.

---

## 2. As Dúvidas Frequentes da Equipe Técnica

### 1. Como montar o código Copia e Cola para o ActiveReports?
Você **não** monta o texto do Pix no código VB6 nem inventa um Pix estático. Quem gera a string oficial do Copia e Cola é o **banco**, por meio da API de Boleto/CobV (`CriarBoletoHibridoPix` da DLL).
A DLL devolve o campo `pixCopiaECola`. No ActiveReports, no evento `Detail_Format` ou antes de exibir o relatório, basta passar essa string para a sua função de desenho de imagem:
```vb
Me.ImgQrCode.Picture = SuaFuncaoDesenharQrCode(sPixCopiaECola)
```

### 2. Validade, Vencimento, Multa e Juros
No padrão do Banco Central (**Pix CobV - Cobrança com Vencimento**):
- O QR Code carrega um endereço da Web seguro do banco emissor.
- Quando o cliente aponta o celular para o QR Code, o aplicativo do banco dele consulta esse endereço **naquele exato segundo**.
- **Até o vencimento:** o banco cobra o valor normal.
- **Após o vencimento:** o próprio banco calcula os dias de atraso, soma a **multa percentual** e os **juros pró-rata dia**, e já mostra na tela do celular o valor final atualizado. O pagador não pode alterar o valor nem burlar os juros.

### 3. Descontos por Antecipação
O funcionamento é 100% automático. Se o boleto tiver instrução de *"R$ 10,00 de desconto até dia 10"*, essa regra fica gravada no banco. Se o cliente escanear o QR Code no dia 09, o app do banco dele exibirá:
- **Valor original:** R$ 150,00
- **Desconto:** R$ 10,00
- **Total a pagar:** R$ 140,00

### 4. Como o Banco Devolve os Pagamentos (Conciliação)?
Você tem dois caminhos que podem funcionar juntos:
1. **Pelo arquivo de retorno CNAB (240 ou 400):** O banco continua enviando o arquivo `.ret` diário normalmente. Nele, a ocorrência vem identificada (ex: `06 - Liquidação via Pix`), trazendo o `Nosso Número`, o valor pago, o desconto concedido, a multa/juros e a tarifa bancária. O seu leitor atual de CNAB processa sem grandes alterações.
2. **Pela DLL / API em Tempo Real:** Chamando `ConsultarBoletoPix(nossoNumero)`, o ERP sabe no mesmo minuto se o cliente pagou, com a discriminação exata de valores.

### 5. Precisa de Arquivo de Remessa CNAB?
- **Pelo modelo moderno (API):** Não precisa esperar o fechamento da remessa. O boleto já nasce registrado no banco na hora em que o operador emite no ERP.
- **Pelo modelo híbrido tradicional:** Alguns bancos permitem que você gere a remessa CNAB tradicional e, no arquivo de retorno do dia seguinte, eles devolvem o código Pix Copia e Cola para você imprimir. Porém, para imprimir o boleto na hora para o cliente, a **API online** é a solução padrão adotada pelo mercado.

### 6. Identificador Único: Como Fazer a Baixa Automática no Caixa?
O elo de ligação entre o mundo tradicional e o Pix é o **`Nosso Número`**:
- O mesmo `Nosso Número` impresso na linha digitável é enviado como `txid` na API Pix.
- No arquivo de retorno CNAB ou no JSON da API, o banco sempre devolve esse `Nosso Número`. O seu sistema localiza a duplicata no Contas a Receber e baixa na hora.

---

## 3. O Mistério do Aplicativo do Itaú que Não Leu o QR Code

No seu teste anterior, 4 aplicativos leram e apenas o Itaú rejeitou. Por que isso aconteceu?

1. **Validação Rigorosa de Payload pelo Itaú:** O Itaú é um dos bancos mais rígidos na validação das normas do Banco Central. Se o QR Code impresso no papel for um Pix estático simples (gerado com chave avulsa) ou faltar qualquer tag obrigatória da norma EMVCo do Bacen, o app do Itaú descarta e não abre a tela de pagamento.
2. **Registro Obrigatório na CIP/NPC:** Em boletos híbridos, o Itaú exige que o `txid` corresponda a um boleto registrado oficialmente no sistema de compensação. Se alguém tentar "improvisar" um QR Code no boleto sem registrar previamente no banco, o Itaú recusa a leitura por segurança contra fraudes.
3. **Solução Definitiva:** Gerando o Pix CobV através da API do banco emissor (como implementado no `CriarBoletoHibridoPix`), a string Copia e Cola gerada é 100% homologada pelo Bacen e aceita pelo app do Itaú, Bradesco, Santander, Nubank, Caixa e todos os demais.

---

## 4. O que o Lojista Deve Solicitar ao Gerente do Banco?

Ao contratar ou ativar o serviço no banco da empresa, peça o seguinte pacote:
1. **Habilitação de Boleto Híbrido com Pix Cobrança (API V2 CobV).**
2. **Acesso ao Portal de Desenvolvedores / Developers do Banco.**
3. **Criação das credenciais de aplicação:** `Client ID` e `Client Secret`.
4. **Instalação do Certificado A1 (.pfx)** para autenticação mTLS.
5. **Configuração da Chave Pix** vinculada à conta corrente de cobrança da empresa.

---

## 5. Como Testar no VB6 com o Projeto BridgeRTC

No formulário `frmBoletoPix.frm` disponível na pasta `ExemploVB6`, o time pode testar:
1. Preenchimento de valor, vencimento, desconto com data limite, multa e juros.
2. Clique em **"Gerar Pix Copia e Cola para ActiveReports"** (gera a string oficial).
3. Teste das 3 simulações reais:
   - **Pagamento Antecipado:** confere o desconto aplicado sozinho pelo banco.
   - **Pagamento no Vencimento:** confere o valor normal.
   - **Pagamento em Atraso:** confere a soma exata de multa + juros pró-rata.
