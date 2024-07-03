using Microsoft.EntityFrameworkCore.Metadata.Builders;
using Microsoft.EntityFrameworkCore;
using Microsoft.VisualBasic;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Fiscal.Classe {
    public class Estoque {
        [DisplayName("Código")]
        public int Controle {
            get; set;
        }

        [DisplayName("Produto")]
        public string Produto {
            get; set;
        }

        [DisplayName("Tipo Barra")]
        public string TipoBarra {
            get; set;
        }

        [DisplayName("Código de Barras")]
        public string CodBarras {
            get; set;
        }

        [DisplayName("Unidade")]
        public string Unidade {
            get; set;
        }

        [DisplayName("Preço de Custo")]
        public decimal? PrecoCusto {
            get; set;
        }

        [DisplayName("Percentual de Lucro")]
        public decimal? PercLucro {
            get; set;
        }

        [DisplayName("Preço de Venda")]
        public decimal? PrecoVenda {
            get; set;
        }

        [DisplayName("Código Grupo")]
        public int? CodGrupo {
            get; set;
        }

        [DisplayName("Grupo")]
        public string Grupo {
            get; set;
        }

        [DisplayName("OBS")]
        public string OBS {
            get; set;
        }

        [DisplayName("Código Fornecedor")]
        public string CODFORNECEDOR {
            get; set;
        }

        [DisplayName("Fornecedor")]
        public string Fornecedor {
            get; set;
        }

        [DisplayName("Tamanho")]
        public string Tamanho {
            get; set;
        }

        [DisplayName("Peso")]
        public decimal? Peso {
            get; set;
        }

        [DisplayName("Custo Ultima Compra")]
        public decimal? CustoUltimaCompra {
            get; set;
        }

        [DisplayName("Custo Médio")]
        public decimal? CustoMedio {
            get; set;
        }

        [DisplayName("QTDE Total Comprada")]
        public decimal? QTDETotalComprada {
            get; set;
        }

        [DisplayName("QTDE Total Vendida")]
        public decimal? QTDETotalVendida {
            get; set;
        }

        [DisplayName("Data Ultima Compra")]
        public DateTime? DataUltimaCompra {
            get; set;
        }

        [DisplayName("Data Ultima Venda")]
        public DateTime? DataUltimaVenda {
            get; set;
        }

        [DisplayName("Data Hora Cadastro")]
        public DateTime? DataHoraCadastro {
            get; set;
        }

        [DisplayName("QTDE Minima")]
        public decimal? QTDEMinima {
            get; set;
        }

        [DisplayName("QTDE Maxima")]
        public decimal? QTDEMaxima {
            get; set;
        }

        [DisplayName("QTDE")]
        public decimal? QTDE {
            get; set;
        }

        [DisplayName("Ativo")]
        public string Ativo {
            get; set;
        }

        [DisplayName("NCM")]
        public string NCM {
            get; set;
        }

        [DisplayName("Percentual Credito ICMS")]
        public decimal? Percreditocicms {
            get; set;
        }

        [DisplayName("Usa Grade")]
        public string UsaGrade {
            get; set;
        }

        [DisplayName("Usa Serial")]
        public string UsaSerial {
            get; set;
        }

        [DisplayName("Origem")]
        public string Origem {
            get; set;
        }

        [DisplayName("Código Tributação IPI")]
        public string CodTributacaoIPI {
            get; set;
        }

        [DisplayName("Tributação IPI")]
        public string TributacaoIPI {
            get; set;
        }

        [DisplayName("Código Tributação Pis")]
        public string CodTributacaoPIS {
            get; set;
        }

        [DisplayName("Código Tributação Cofins")]
        public string CodTributacaoCofins {
            get; set;
        }

        [DisplayName("Tributação Pis")]
        public string TributacaoPis {
            get; set;
        }

        [DisplayName("Tributação Cofins")]
        public string TributacaoCofins {
            get; set;
        }

        [DisplayName("Tipo Tributação")]
        public string TipoTributação {
            get; set;
        }

        [DisplayName("Percentual ICMS ST Interna")]
        public decimal? PercIcmsStInterna {
            get; set;
        }

        [DisplayName("Percentual MVA Original")]
        public decimal? PercMvaOriginal {
            get; set;
        }

        [DisplayName("Possui ICMS ST")]
        public string PossuiIcmsSt {
            get; set;
        }

        [DisplayName("Isento")]
        public string Isento {
            get; set;
        }

        [DisplayName("PercIcmsProprioSt")]
        public decimal? PercIcmsProprioSt {
            get; set;
        }

        [DisplayName("IAT")]
        public string IAT {
            get; set;
        }

        [DisplayName("IPPT")]
        public string IPPT {
            get; set;
        }

        [DisplayName("Tributado")]
        public string Tributado {
            get; set;
        }

        [DisplayName("CSOSN")]
        public string CSOSN {
            get; set;
        }

        [DisplayName("Descrição CSOSN")]
        public string DescricaoCSOSN {
            get; set;
        }

        [DisplayName("Pesado")]
        public string Pesado {
            get; set;
        }

        [DisplayName("Base Calculo Icms ST Retido")]
        public decimal? BaseCalculoIcmsStRetido {
            get; set;
        }

        [DisplayName("VALOR ICMS ST RETIDO")]
        public decimal? VALORICMSSTRETIDO {
            get; set;
        }

        [DisplayName("Valor Icms ST Retido")]
        public decimal? VALORBCICMSSTRET {
            get; set;
        }

        [DisplayName("Alíquota Icms Ecf")]
        public decimal? AliquotaIcmsEcf {
            get; set;
        }

        [DisplayName("Foto")]
        public string Foto {
            get; set;
        }

        [DisplayName("MD5")]
        public string MD5 {
            get; set;
        }

        [DisplayName("Mensagem NF-e")]
        public string MensagemNfe {
            get; set;
        }

        [DisplayName("Código Mensagem NF-e")]
        public int? CodMensagemNfe {
            get; set;
        }

        [DisplayName("Código Unidade De Medida")]
        public int? CodUnidadeMedida {
            get; set;
        }

        [DisplayName("Código Aplicação Produto")]
        public string CodAplicacaoProduto {
            get; set;
        }

        [DisplayName("Aplicação Produto")]
        public string AplicacaoProduto {
            get; set;
        }

        [DisplayName("MD5E2")]
        public string MD5E2 {
            get; set;
        }

        [DisplayName("QTDE Inicial")]
        public decimal? QTDEInicial {
            get; set;
        }

        [DisplayName("Aliquota ISS")]
        public decimal? AliquotaISS {
            get; set;
        }

        [DisplayName("Código imposto médio")]
        public int? CODIMPOSTOMEDIO {
            get; set;
        }

        [DisplayName("Percentual Imposto Medio")]
        public decimal? PercImpostoMedio {
            get; set;
        }

        [DisplayName("Código CST Origem")]
        public int? CodCstOrigem {
            get; set;
        }

        [DisplayName("Codigo cst Origem")]
        public int? CodigoCstOrigem {
            get; set;
        }

        [DisplayName("Código Emitente")]
        public int? CodEmitente {
            get; set;
        }

        [DisplayName("Dias Validade Produto")]
        public int? DiasValidadeProduto {
            get; set;
        }

        [DisplayName("Fator de Conversão")]
        public string FatorConversao {
            get; set;
        }

        [DisplayName("Tributação Serviço")]
        public string TributaçãoServiço {
            get; set;
        }

        [DisplayName("Redução Base Calculo Serviço")]
        public decimal? ReduçãoBaseCalculoServiço {
            get; set;
        }

        [DisplayName("Status")]
        public string Status {
            get; set;
        }

        [DisplayName("Descrição complementar")]
        public string DescriçãoComplementar {
            get; set;
        }

        [DisplayName("Referenciar")]
        public string Referenciar {
            get; set;
        }

        [DisplayName("Preço Venda USS")]
        public decimal? PrecoVendaUss {
            get; set;
        }

        [DisplayName("Percentual Maximo Desconto")]
        public decimal? PercentualMaximoDesconto {
            get; set;
        }

        [DisplayName("Valor Comissao Fixo")]
        public decimal? ValorComissaoFixo {
            get; set;
        }

        [DisplayName("Perc Comissao")]
        public decimal? PercComissao {
            get; set;
        }

        [DisplayName("Preço Minimo USS")]
        public decimal? PrecoMinimoUss {
            get; set;
        }

        [DisplayName("Preço Minimo")]
        public decimal? PrecoMinimo {
            get; set;
        }

        [DisplayName("Código Compra")]
        public int? CodCompra {
            get; set;
        }

        [DisplayName("Valor Frete")]
        public decimal? ValorFrete {
            get; set;
        }

        [DisplayName("Valor Outros")]
        public decimal? ValorOutros {
            get; set;
        }

        [DisplayName("Valor Icms ST")]
        public decimal? ValorIcmsSt {
            get; set;
        }

        [DisplayName("Valor Ipi")]
        public decimal? ValorIpi {
            get; set;
        }

        [DisplayName("Valor Unitario Compra")]
        public decimal? ValorUnitarioCompra {
            get; set;
        }

        [DisplayName("Perc Pis")]
        public decimal? PercPis {
            get; set;
        }

        [DisplayName("Perc Cofins")]
        public decimal? PercCofins {
            get; set;
        }

        [DisplayName("Valor Conversão")]
        public decimal? ValorConversao {
            get; set;
        }

        [DisplayName("Campo 1")]
        public string Campo1 {
            get; set;
        }

        [DisplayName("Campo 2")]
        public string Campo2 {
            get; set;
        }

        [DisplayName("Campo 3")]
        public string Campo3 {
            get; set;
        }

        [DisplayName("Campo 4")]
        public string Campo4 {
            get; set;
        }

        [DisplayName("Campo 5")]
        public string Campo5 {
            get; set;
        }

        [DisplayName("Campo 6")]
        public string Campo6 {
            get; set;
        }

        [DisplayName("Campo 7")]
        public string Campo7 {
            get; set;
        }

        [DisplayName("Campo 8")]
        public string Campo8 {
            get; set;
        }

        [DisplayName("Campo 9")]
        public string Campo9 {
            get; set;
        }

        [DisplayName("Campo 10")]
        public string Campo10 {
            get; set;
        }

        [DisplayName("CodigoAnp")]
        public string CodigoAnp {
            get; set;
        }

        [DisplayName("Marca")]
        public string Marca {
            get; set;
        }

        [DisplayName("Perc Redução BC")]
        public decimal? PercReducaoBC {
            get; set;
        }

        [DisplayName("Perc Reducao BCST")]
        public decimal? PercReducaoBCST {
            get; set;
        }

        [DisplayName("Código Sub Grupo")]
        public string CodSubGrupo {
            get; set;
        }

        [DisplayName("Controlar Validade")]
        public string ControlarValidade {
            get; set;
        }

        [DisplayName("Código Marca")]
        public int? CodMarca {
            get; set;
        }

        [DisplayName("Preço Revenda")]
        public decimal? PrecoRevenda {
            get; set;
        }

        [DisplayName("Perc Ipi")]
        public decimal? PercIpi {
            get; set;
        }

        [DisplayName("Cfop")]
        public string Cfop {
            get; set;
        }

        [DisplayName("Unidade Conversao Venda")]
        public string UnidadeConversaoVenda {
            get; set;
        }

        [DisplayName("Valor Conversao Venda")]
        public decimal? ValorConversaoVenda {
            get; set;
        }

        [DisplayName("Código Tabela Preço")]
        public int? CodTabelaPreco {
            get; set;
        }

        [DisplayName("Nome Tabela de Preço")]
        public string NomeTabelaPreco {
            get; set;
        }

        [DisplayName("Perc Imposto Medio Estadual")]
        public decimal? PercImpostoMedioEstadual {
            get; set;
        }

        [DisplayName("Perc Imposto Medio Municipal")]
        public decimal? PercImpostoMedioMunicipal {
            get; set;
        }

        [DisplayName("Código Enquadramento")]
        public string CódigoEnquadramento {
            get; set;
        }

        [DisplayName("Cest")]
        public string Cest {
            get; set;
        }

        [DisplayName("Qtde Em Produção")]
        public decimal? QtdeEmProducao {
            get; set;
        }

        [DisplayName("Qtde Pedido Venda")]
        public decimal? QtdePedidoVenda {
            get; set;
        }

        [DisplayName("Qtde Pedido Compra")]
        public decimal? QtdePedidoCompra {
            get; set;
        }

        [DisplayName("Qtde Reservada")]
        public decimal? QtdeReservada {
            get; set;
        }

        [DisplayName("Qtde Real")]
        public decimal? QtdeReal {
            get; set;
        }

        [DisplayName("Qtde em Produção MP")]
        public string QTDEEMPRODUCAOMP {
            get; set;
        }

        [DisplayName("Unidade Medida Etiqueta")]
        public string UnidadeMedidaEtiqueta {
            get; set;
        }

        [DisplayName("Fator de Conversão Etiqueta")]
        public string FatorConversaoEtiqueta {
            get; set;
        }

        [DisplayName("Selo IPI")]
        public string SeloIpi {
            get; set;
        }

        [DisplayName("ALIQUOTAFCP")]
        public decimal? ALIQUOTAFCP {
            get; set;
        }

        [DisplayName("PERCFCPST")]
        public decimal? PERCFCPST {
            get; set;
        }

        [DisplayName("DESCRICAOANP")]
        public string DESCRICAOANP {
            get; set;
        }

        [DisplayName("PERCGLPCOMB")]
        public decimal? PERCGLPCOMB {
            get; set;
        }

        [DisplayName("PERCGNNCOMB")]
        public decimal? PERCGNNCOMB {
            get; set;
        }

        [DisplayName("PERCGNICOMB")]
        public decimal? PERCGNICOMB {
            get; set;
        }

        [DisplayName("VALORPARTIDACOMB")]
        public decimal? VALORPARTIDACOMB {
            get; set;
        }

        [DisplayName("CNPJFABRICANTE")]
        public string CNPJFABRICANTE {
            get; set;
        }

        [DisplayName("CODBENEFICIOFISCAL")]
        public int? CODBENEFICIOFISCAL {
            get; set;
        }

        [DisplayName("CODIGOANVISA")]
        public string CODIGOANVISA {
            get; set;
        }

        [DisplayName("CODUNIDADETRIBUTAVEL")]
        public int? CODUNIDADETRIBUTAVEL {
            get; set;
        }

        [DisplayName("UNIDADETRIBUTAVEL")]
        public string UNIDADETRIBUTAVEL {
            get; set;
        }

        [DisplayName("QTDETRIBUTAVEL")]
        public decimal? QTDETRIBUTAVEL {
            get; set;
        }

        [DisplayName("DESMONTARNAVENDA")]
        public string DESMONTARNAVENDA {
            get; set;
        }

        [DisplayName("ALIQUOTAICMSSTRET")]
        public decimal? ALIQUOTAICMSSTRET {
            get; set;
        }

        [DisplayName("ALIQUOTAICMSEFET")]
        public decimal? ALIQUOTAICMSEFET {
            get; set;
        }

        [DisplayName("PERCREDUCAOICMSEFET")]
        public decimal? PERCREDUCAOICMSEFET {
            get; set;
        }

        [DisplayName("VALORBCICMSEFET")]
        public decimal? VALORBCICMSEFET {
            get; set;
        }

        [DisplayName("VALORICMSEFET")]
        public decimal? VALORICMSEFET {
            get; set;
        }

        [DisplayName("VALORICMSSUBSTITUTO")]
        public decimal? VALORICMSSUBSTITUTO {
            get; set;
        }

        [DisplayName("CODBARRASCAIXA")]
        public string CODBARRASCAIXA {
            get; set;
        }

        [DisplayName("ENVIARDADOS")]
        public string ENVIARDADOS {
            get; set;
        }

        [DisplayName("VALORPMC")]
        public decimal? VALORPMC {
            get; set;
        }

        [DisplayName("MD5O")]
        public string MD5O {
            get; set;
        }

        [DisplayName("MD5S")]
        public string MD5S {
            get; set;
        }

        [DisplayName("VALORFCPST")]
        public decimal? VALORFCPST {
            get; set;
        }

        [DisplayName("LOCALIZACAO")]
        public string LOCALIZACAO {
            get; set;
        }

        [DisplayName("VOLUME")]
        public decimal? VOLUME {
            get; set;
        }

        [DisplayName("MOSTRARCOZINHA")]
        public string MOSTRARCOZINHA {
            get; set;
        }

        [DisplayName("CODBARRASINTERNO")]
        public string CODBARRASINTERNO {
            get; set;
        }

        [DisplayName("CODBARRASTRIB")]
        public string CODBARRASTRIB {
            get; set;
        }

        [DisplayName("VACINA")]
        public string VACINA {
            get; set;
        }

        [DisplayName("PERCDESCONTOCAIXA")]
        public decimal? PERCDESCONTOCAIXA {
            get; set;
        }

        [DisplayName("SINCRONIZADO")]
        public string SINCRONIZADO {
            get; set;
        }

        [DisplayName("PERCCASHBACK")]
        public decimal? PERCCASHBACK {
            get; set;
        }
    }
    public class EstoqueEntityTypeConfiguration : IEntityTypeConfiguration<Estoque> {
        public void Configure(EntityTypeBuilder<Estoque> builder) {
            builder.ToTable("TESTOQUE");

            builder.HasKey(e => e.Controle);

            builder.Property(e => e.Controle).HasColumnName("CONTROLE").IsRequired();

            builder.Property(e => e.Produto).HasColumnName("PRODUTO");

            builder.Property(e => e.TipoBarra).HasColumnName("TIPOBARRA");

            builder.Property(e => e.CodBarras).HasColumnName("CODBARRAS");

            builder.Property(e => e.Unidade).HasColumnName("UNIDADE");

            builder.Property(e => e.PrecoCusto).HasColumnName("PRECOCUSTO");

            builder.Property(e => e.PercLucro).HasColumnName("PERCLUCRO");

            builder.Property(e => e.PrecoVenda).HasColumnName("PRECOVENDA");

            builder.Property(e => e.CodGrupo).HasColumnName("CODGRUPO");

            builder.Property(e => e.Grupo).HasColumnName("GRUPO");

            builder.Property(e => e.CODFORNECEDOR).HasColumnName("CODFORNECEDOR");

            builder.Property(e => e.Fornecedor).HasColumnName("FORNECEDOR");

            builder.Property(e => e.Tamanho).HasColumnName("TAMANHO");

            builder.Property(e => e.Peso).HasColumnName("PESO");

            builder.Property(e => e.CustoUltimaCompra).HasColumnName("CUSTOULTIMACOMPRA");

            builder.Property(e => e.CustoMedio).HasColumnName("CUSTOMEDIO");

            builder.Property(e => e.QTDETotalComprada).HasColumnName("QTDETOTALCOMPRADA");

            builder.Property(e => e.QTDETotalVendida).HasColumnName("QTDETOTALVENDIDA");

            builder.Property(e => e.DataUltimaCompra).HasColumnName("DATAULTIMACOMPRA");

            builder.Property(e => e.DataUltimaVenda).HasColumnName("DATAULTIMAVENDA");

            builder.Property(e => e.DataHoraCadastro).HasColumnName("DATAHORACADASTRO");

            builder.Property(e => e.QTDEMinima).HasColumnName("QTDEMINIMA");

            builder.Property(e => e.QTDEMaxima).HasColumnName("QTDEMAXIMA");

            builder.Property(e => e.QTDE).HasColumnName("QTDE");

            builder.Property(e => e.Ativo).HasColumnName("ATIVO");

            builder.Property(e => e.NCM).HasColumnName("NCM");

            builder.Property(e => e.Percreditocicms).HasColumnName("PERCREDITOCICMS");

            builder.Property(e => e.UsaGrade).HasColumnName("USAGRADE");

            builder.Property(e => e.UsaSerial).HasColumnName("USASERIAL");

            builder.Property(e => e.Origem).HasColumnName("ORIGEM");

            builder.Property(e => e.CodTributacaoIPI).HasColumnName("CODTRIBUTACAOIPI");

            builder.Property(e => e.TributacaoIPI).HasColumnName("TRIBUTACAOIPI");

            builder.Property(e => e.CodTributacaoPIS).HasColumnName("CODTRIBUTACAOPIS");

            builder.Property(e => e.CodTributacaoCofins).HasColumnName("CODTRIBUTACAOCOFINS");

            builder.Property(e => e.TributacaoPis).HasColumnName("TRIBUTACAOPIS");

            builder.Property(e => e.TributacaoCofins).HasColumnName("TRIBUTACAOCOFINS");

            builder.Property(e => e.TipoTributação).HasColumnName("TIPOTRIBUTACAO");

            builder.Property(e => e.PercIcmsStInterna).HasColumnName("PERCICMSSTINTERNA");

            builder.Property(e => e.PercMvaOriginal).HasColumnName("PERCMVAORIGINAL");

            builder.Property(e => e.PossuiIcmsSt).HasColumnName("POSSUIICMSST");

            builder.Property(e => e.Isento).HasColumnName("ISENTO");

            builder.Property(e => e.PercIcmsProprioSt).HasColumnName("PERCICMSPROPRIOST");

            builder.Property(e => e.IAT).HasColumnName("IAT");

            builder.Property(e => e.IPPT).HasColumnName("IPPT");

            builder.Property(e => e.Tributado).HasColumnName("TRIBUTADO");

            builder.Property(e => e.CSOSN).HasColumnName("CSOSN");

            builder.Property(e => e.DescricaoCSOSN).HasColumnName("DESCRICAOCSOSN");

            builder.Property(e => e.Pesado).HasColumnName("PESADO");

            builder.Property(e => e.BaseCalculoIcmsStRetido).HasColumnName("BASECALCULOICMSSTRETIDO");

            builder.Property(e => e.VALORICMSSTRETIDO).HasColumnName("VALORICMSSTRETIDO");

            builder.Property(e => e.AliquotaIcmsEcf).HasColumnName("ALIQUOTAICMSECF");

            builder.Property(e => e.Foto).HasColumnName("FOTO");

            builder.Property(e => e.MD5).HasColumnName("MD5");

            builder.Property(e => e.MensagemNfe).HasColumnName("MENSAGEMNFE");

            builder.Property(e => e.CodMensagemNfe).HasColumnName("CODMENSAGEMNFE");

            builder.Property(e => e.CodUnidadeMedida).HasColumnName("CODUNIDADEMEDIDA");

            builder.Property(e => e.CodAplicacaoProduto).HasColumnName("CODAPLICACAOPRODUTO");

            builder.Property(e => e.AplicacaoProduto).HasColumnName("APLICACAOPRODUTO");

            builder.Property(e => e.MD5E2).HasColumnName("MD5E2");

            builder.Property(e => e.QTDEInicial).HasColumnName("QTDEINICIAL");

            builder.Property(e => e.AliquotaISS).HasColumnName("ALIQUOTAISS");

            builder.Property(e => e.CODIMPOSTOMEDIO).HasColumnName("CODIMPOSTOMEDIO");

            builder.Property(e => e.PercImpostoMedio).HasColumnName("PERCIMPOSTOMEDIO");

            builder.Property(e => e.CodCstOrigem).HasColumnName("CODCSTORIGEM");

            builder.Property(e => e.CodigoCstOrigem).HasColumnName("CODIGOCSTORIGEM");

            builder.Property(e => e.CodEmitente).HasColumnName("CODEMITENTE");

            builder.Property(e => e.DiasValidadeProduto).HasColumnName("DIASVALIDADEPRODUTO");

            builder.Property(e => e.FatorConversao).HasColumnName("FATORCONVERSAO");

            builder.Property(e => e.TributaçãoServiço).HasColumnName("TRIBUTACAOSERVICO");

            builder.Property(e => e.ReduçãoBaseCalculoServiço).HasColumnName("REDUCAOBASECALCULOSERVICO");

            builder.Property(e => e.Status).HasColumnName("STATUS");

            builder.Property(e => e.DescriçãoComplementar).HasColumnName("DESCRICAOCOMPLEMENTAR");

            builder.Property(e => e.Referenciar).HasColumnName("REFERENCIA");

            builder.Property(e => e.PrecoVendaUss).HasColumnName("PRECOVENDAUSS");

            builder.Property(e => e.PercentualMaximoDesconto).HasColumnName("PERCMAXIMODESCONTO");

            builder.Property(e => e.ValorComissaoFixo).HasColumnName("VALORCOMISSAOFIXO");

            builder.Property(e => e.PercComissao).HasColumnName("PERCCOMISSAO");

            builder.Property(e => e.PrecoMinimoUss).HasColumnName("PRECOMINIMOUSS");

            builder.Property(e => e.PrecoMinimo).HasColumnName("PRECOMINIMO");

            builder.Property(e => e.CodCompra).HasColumnName("CODCOMPRA");

            builder.Property(e => e.ValorFrete).HasColumnName("VALORFRETE");

            builder.Property(e => e.ValorOutros).HasColumnName("VALOROUTROS");

            builder.Property(e => e.ValorIcmsSt).HasColumnName("VALORICMSST");

            builder.Property(e => e.ValorIpi).HasColumnName("VALORIPI");

            builder.Property(e => e.ValorUnitarioCompra).HasColumnName("VALORUNITARIOCOMPRA");

            builder.Property(e => e.PercPis).HasColumnName("PERCPIS");

            builder.Property(e => e.PercCofins).HasColumnName("PERCCOFINS");

            builder.Property(e => e.ValorConversao).HasColumnName("VALORCONVERSAO");

            builder.Property(e => e.Campo1).HasColumnName("CAMPO1");

            builder.Property(e => e.Campo2).HasColumnName("CAMPO2");

            builder.Property(e => e.Campo3).HasColumnName("CAMPO3");

            builder.Property(e => e.Campo4).HasColumnName("CAMPO4");

            builder.Property(e => e.Campo5).HasColumnName("CAMPO5");

            builder.Property(e => e.Campo6).HasColumnName("CAMPO6");

            builder.Property(e => e.Campo7).HasColumnName("CAMPO7");

            builder.Property(e => e.Campo8).HasColumnName("CAMPO8");

            builder.Property(e => e.Campo9).HasColumnName("CAMPO9");

            builder.Property(e => e.Campo10).HasColumnName("CAMPO10");

            builder.Property(e => e.CodigoAnp).HasColumnName("CODIGOANP");

            builder.Property(e => e.Marca).HasColumnName("MARCA");

            builder.Property(e => e.PercReducaoBC).HasColumnName("PERCREDUCAOBC");

            builder.Property(e => e.PercReducaoBCST).HasColumnName("PERCREDUCAOBCST");

            builder.Property(e => e.CodSubGrupo).HasColumnName("CODSUBGRUPO");

            builder.Property(e => e.CodSubGrupo).HasColumnName("SUBGRUPO");

            builder.Property(e => e.ControlarValidade).HasColumnName("CONTROLARVALIDADE");

            builder.Property(e => e.CodMarca).HasColumnName("CODMARCA");

            builder.Property(e => e.PrecoRevenda).HasColumnName("PRECOREVENDA");

            builder.Property(e => e.PercIpi).HasColumnName("PERCIPI");

            builder.Property(e => e.Cfop).HasColumnName("CFOP");

            builder.Property(e => e.UnidadeConversaoVenda).HasColumnName("UNIDADECONVERSAOVENDA");

            builder.Property(e => e.ValorConversaoVenda).HasColumnName("VALORCONVERSAOVENDA");

            builder.Property(e => e.CodTabelaPreco).HasColumnName("CODTABELAPRECO");

            builder.Property(e => e.NomeTabelaPreco).HasColumnName("NOMETABELAPRECO");

            builder.Property(e => e.PercImpostoMedioEstadual).HasColumnName("PERCIMPOSTOMEDIOESTADUAL");

            builder.Property(e => e.PercImpostoMedioMunicipal).HasColumnName("PERCIMPOSTOMEDIOMUNICIPAL");

            builder.Property(e => e.CódigoEnquadramento).HasColumnName("CODIGOENQUADRAMENTOIPI");

            builder.Property(e => e.Cest).HasColumnName("CEST");

            builder.Property(e => e.QtdeEmProducao).HasColumnName("QTDEEMPRODUCAO");

            builder.Property(e => e.QtdePedidoVenda).HasColumnName("QTDEPEDIDOVENDA");

            builder.Property(e => e.QtdePedidoCompra).HasColumnName("QTDEPEDIDOCOMPRA");

            builder.Property(e => e.QtdeReservada).HasColumnName("QTDERESERVADA");

            builder.Property(e => e.QtdeReal).HasColumnName("QTDEREAL");

            builder.Property(e => e.QTDEEMPRODUCAOMP).HasColumnName("QTDEEMPRODUCAOMP");

            builder.Property(e => e.UnidadeMedidaEtiqueta).HasColumnName("UNIDADEMEDIDAETIQUETA");

            builder.Property(e => e.FatorConversaoEtiqueta).HasColumnName("FATORCONVERSAOETIQUETA");

            builder.Property(e => e.SeloIpi).HasColumnName("SELOIPI");

            builder.Property(e => e.ALIQUOTAFCP).HasColumnName("ALIQUOTAFCP");

            builder.Property(e => e.PERCFCPST).HasColumnName("PERCFCPST");

            builder.Property(e => e.DESCRICAOANP).HasColumnName("DESCRICAOANP");

            builder.Property(e => e.PERCGLPCOMB).HasColumnName("PERCGLPCOMB");

            builder.Property(e => e.PERCGNNCOMB).HasColumnName("PERCGNNCOMB");

            builder.Property(e => e.PERCGNICOMB).HasColumnName("PERCGNICOMB");

            builder.Property(e => e.VALORPARTIDACOMB).HasColumnName("VALORPARTIDACOMB");

            builder.Property(e => e.CNPJFABRICANTE).HasColumnName("CNPJFABRICANTE");

            builder.Property(e => e.CODBENEFICIOFISCAL).HasColumnName("CODBENEFICIOFISCAL");

            builder.Property(e => e.CODIGOANVISA).HasColumnName("CODIGOANVISA");

            builder.Property(e => e.CODUNIDADETRIBUTAVEL).HasColumnName("CODUNIDADETRIBUTAVEL");

            builder.Property(e => e.UNIDADETRIBUTAVEL).HasColumnName("UNIDADETRIBUTAVEL");

            builder.Property(e => e.QTDETRIBUTAVEL).HasColumnName("QTDETRIBUTAVEL");

            builder.Property(e => e.DESMONTARNAVENDA).HasColumnName("DESMONTARNAVENDA");

            builder.Property(e => e.ALIQUOTAICMSSTRET).HasColumnName("ALIQUOTAICMSSTRET");

            builder.Property(e => e.VALORBCICMSSTRET).HasColumnName("VALORBCICMSSTRET");

            builder.Property(e => e.VALORICMSEFET).HasColumnName("VALORICMSEFET");

            builder.Property(e => e.ALIQUOTAICMSEFET).HasColumnName("ALIQUOTAICMSEFET");

            builder.Property(e => e.PERCREDUCAOICMSEFET).HasColumnName("PERCREDUCAOICMSEFET");

            builder.Property(e => e.VALORICMSEFET).HasColumnName("VALORICMSEFET");

            builder.Property(e => e.VALORICMSSUBSTITUTO).HasColumnName("VALORICMSSUBSTITUTO");

            builder.Property(e => e.CODBARRASCAIXA).HasColumnName("CODBARRASCAIXA");

            builder.Property(e => e.ENVIARDADOS).HasColumnName("ENVIARDADOS");

            builder.Property(e => e.VALORPMC).HasColumnName("VALORPMC");

            builder.Property(e => e.MD5O).HasColumnName("MD5O");

            builder.Property(e => e.MD5S).HasColumnName("MD5S");

            builder.Property(e => e.VALORFCPST).HasColumnName("VALORFCPST");

            builder.Property(e => e.LOCALIZACAO).HasColumnName("LOCALIZACAO");

            builder.Property(e => e.VOLUME).HasColumnName("VOLUME");

            builder.Property(e => e.MOSTRARCOZINHA).HasColumnName("MOSTRARCOZINHA");

            builder.Property(e => e.CODBARRASINTERNO).HasColumnName("CODBARRASINTERNO");

            builder.Property(e => e.CODBARRASTRIB).HasColumnName("CODBARRASTRIB");

            builder.Property(e => e.VACINA).HasColumnName("VACINA");

            builder.Property(e => e.PERCDESCONTOCAIXA).HasColumnName("PERCDESCONTOCAIXA");

            builder.Property(e => e.SINCRONIZADO).HasColumnName("SINCRONIZADO");

            builder.Property(e => e.PERCCASHBACK).HasColumnName("PERCCASHBACK");
        }
    }
}
