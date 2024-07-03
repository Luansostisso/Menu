using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.EntityFrameworkCore.Metadata.Builders;
using Microsoft.EntityFrameworkCore;

namespace Fiscal.Classe {
    public class Cliente {

        [DisplayName("Controle")]
        public int Controle {
            get; set;
        }

        [DisplayName("Cliente")]
        public string RazaoSocial {
            get; set;
        }

        [DisplayName("Endereço")]
        public string Endereco {
            get; set;
        }

        [DisplayName("Complemento")]
        public string Complemento {
            get; set;
        }

        [DisplayName("Bairro")]
        public string Bairro {
            get; set;
        }

        [DisplayName("Código da Cidade")]
        public int? CodCidade {
            get; set;
        }

        [DisplayName("Cidade")]
        public string Cidade {
            get; set;
        }

        [DisplayName("UF")]
        public string UF {
            get; set;
        }

        [DisplayName("País")]
        public string Pais {
            get; set;
        }

        [DisplayName("CEP")]
        public string CEP {
            get; set;
        }

        [DisplayName("Naturalidade")]
        public string Naturalidade {
            get; set;
        }

        [DisplayName("Tipo de Cliente")]
        public string TipoCliente {
            get; set;
        }

        [DisplayName("RG")]
        public string RG {
            get; set;
        }

        [DisplayName("CPF")]
        public string CPF {
            get; set;
        }

        [DisplayName("CNPJ")]
        public string CNPJ {
            get; set;
        }

        [DisplayName("IE")]
        public string IE {
            get; set;
        }

        [DisplayName("IM")]
        public string IM {
            get; set;
        }

        [DisplayName("Data de Nascimento")]
        public DateTime? DataNascimento {
            get; set;
        }

        [DisplayName("Data e Hora de Cadastro")]
        public DateTime? DataHoraCadastro {
            get; set;
        }

        [DisplayName("Pai")]
        public string Pai {
            get; set;
        }

        [DisplayName("Mãe")]
        public string Mae {
            get; set;
        }

        [DisplayName("Telefone")]
        public string Telefone {
            get; set;
        }

        [DisplayName("Celular")]
        public string Celular {
            get; set;
        }

        [DisplayName("Email")]
        public string Email {
            get; set;
        }

        [DisplayName("Estado Civil")]
        public string EstadoCivil {
            get; set;
        }

        [DisplayName("Ativo")]
        public string Ativo {
            get; set;
        }

        [DisplayName("Fantasia")]
        public string Fantasia {
            get; set;
        }

        [DisplayName("Nome do Cônjuge")]
        public string NomeConjuge {
            get; set;
        }

        [DisplayName("Data da Última Venda")]
        public DateTime? DataUltimaVenda {
            get; set;
        }

        [DisplayName("Dias sem Comprar")]
        public int? DiasSemComprar {
            get; set;
        }

        [DisplayName("Código do Convênio")]
        public int? CodConvenio {
            get; set;
        }

        [DisplayName("Convênio")]
        public string Convenio {
            get; set;
        }

        [DisplayName("Profissão")]
        public string Profissao {
            get; set;
        }

        [DisplayName("Empresa que Trabalha")]
        public string EmpresaQueTrabalha {
            get; set;
        }

        [DisplayName("Telefone do Trabalho")]
        public string TelefoneTrabalho {
            get; set;
        }

        [DisplayName("Renda Mensal")]
        public decimal? RendaMensal {
            get; set;
        }

        [DisplayName("Total Vendido")]
        public decimal? TotalVendido {
            get; set;
        }

        [DisplayName("Nacionalidade")]
        public string Nacionalidade {
            get; set;
        }

        [DisplayName("Número")]
        public string Numero {
            get; set;
        }

        [DisplayName("Código de Consultas no SPC")]
        public int? CodConsultasSPC {
            get; set;
        }

        [DisplayName("Sexo")]
        public string Sexo {
            get; set;
        }

        [DisplayName("Código da Cidade IBGE")]
        public string CodigoCidadeIBGE {
            get; set;
        }

        [DisplayName("Código do Emitente")]
        public int? CodEmitente {
            get; set;
        }

        [DisplayName("Nome do Contato Jurídica")]
        public string NomeContatoJuridica {
            get; set;
        }

        [DisplayName("Email 2")]
        public string Email2 {
            get; set;
        }

        [DisplayName("Status")]
        public string Status {
            get; set;
        }

        [DisplayName("MD5")]
        public string MD5 {
            get; set;
        }

        [DisplayName("Serial")]
        public string Serial {
            get; set;
        }

        [DisplayName("Código do Grupo")]
        public int? CodGrupo {
            get; set;
        }

        [DisplayName("Grupo")]
        public string Grupo {
            get; set;
        }

        [DisplayName("Tributação")]
        public string Tributacao {
            get; set;
        }

        [DisplayName("Campo1")]
        public string Campo1 {
            get; set;
        }

        [DisplayName("Campo2")]
        public string Campo2 {
            get; set;
        }

        [DisplayName("Campo3")]
        public string Campo3 {
            get; set;
        }

        [DisplayName("Campo4")]
        public string Campo4 {
            get; set;
        }

        [DisplayName("Campo5")]
        public string Campo5 {
            get; set;
        }

        [DisplayName("Campo6")]
        public string Campo6 {
            get; set;
        }

        [DisplayName("Campo7")]
        public string Campo7 {
            get; set;
        }

        [DisplayName("Campo8")]
        public string Campo8 {
            get; set;
        }

        [DisplayName("Campo9")]
        public string Campo9 {
            get; set;
        }

        [DisplayName("Campo10")]
        public string Campo10 {
            get; set;
        }

        [DisplayName("ISS Município")]
        public int? ISSMunicipio {
            get; set;
        }

        [DisplayName("Site")]
        public string Site {
            get; set;
        }

        [DisplayName("Código do País")]
        public string CodigoPais {
            get; set;
        }

        [DisplayName("CSOSN")]
        public string CSOSN {
            get; set;
        }

        [DisplayName("Percentual de ICMS Próprio Substituição Tributária")]
        public decimal? PercICMSProprioST {
            get; set;
        }

        [DisplayName("Percentual MVA Original")]
        public decimal? PercMVAOriginal {
            get; set;
        }

        [DisplayName("Percentual de ICMS ST Interna")]
        public decimal? PercICMSSTInterna {
            get; set;
        }

        [DisplayName("Percentual de Redução na Base de Cálculo ST")]
        public decimal? PercReducaoBCST {
            get; set;
        }

        [DisplayName("Descrição CSOSN")]
        public string DescricaoCSOSN {
            get; set;
        }

        [DisplayName("Código do CST Origem")]
        public int? CodigoCSTOrigem {
            get; set;
        }

        [DisplayName("Origem")]
        public string Origem {
            get; set;
        }

        [DisplayName("Possui ICMS ST")]
        public string PossuiICMSST {
            get; set;
        }

        [DisplayName("Isento")]
        public string Isento {
            get; set;
        }

        [DisplayName("Tributado")]
        public string Tributado {
            get; set;
        }

        [DisplayName("Código CST Origem")]
        public int? CodcstOrigem {
            get; set;
        }

        [DisplayName("Mensagem Fiscal")]
        public string MensagemFiscal {
            get; set;
        }

        [DisplayName("Código Tabela de Preço")]
        public int? CodTabelaPreco {
            get; set;
        }

        [DisplayName("Tabela de Preço")]
        public string TabelaPreco {
            get; set;
        }

        [DisplayName("Código do Vendedor")]
        public int? CodVendedor {
            get; set;
        }

        [DisplayName("Vendedor")]
        public string Vendedor {
            get; set;
        }

        [DisplayName("Limite de Crédito")]
        public decimal? LimiteCredito {
            get; set;
        }

        [DisplayName("Código da Parcela Predefinida")]
        public int? CodParcelaPredefinida {
            get; set;
        }

        [DisplayName("Parcela Predefinida")]
        public string ParcelaPredefinida {
            get; set;
        }

        [DisplayName("Código da Transportadora")]
        public int? CodTransportadora {
            get; set;
        }

        [DisplayName("Transportadora")]
        public string Transportadora {
            get; set;
        }

        [DisplayName("Valor do Frete")]
        public decimal? ValorFrete {
            get; set;
        }

        [DisplayName("MD5 Original")]
        public string MD5Original {
            get; set;
        }

        [DisplayName("Suframa")]
        public string Suframa {
            get; set;
        }

        [DisplayName("Percentual FCP")]
        public decimal? PercFCP {
            get; set;
        }

        [DisplayName("Percentual de Redução na Base de Cálculo")]
        public decimal? PercReducaoBC {
            get; set;
        }

        [DisplayName("CST/CSOSN a Aplicar")]
        public string CSTCSOSNaplicar {
            get; set;
        }

        [DisplayName("Sincronizado")]
        public string Sincronizado {
            get; set;
        }

        [DisplayName("Utilizar Cashback")]
        public string UtilizarCashback {
            get; set;
        }
    }
    public class ClienteEntityTypeConfiguration : IEntityTypeConfiguration<Cliente> {
        public void Configure(EntityTypeBuilder<Cliente> builder) {
            builder.ToTable("TCLIENTE");

            builder.HasKey(e => e.Controle);

            builder.Property(e => e.Controle).HasColumnName("CONTROLE").IsRequired(); ;

            builder.Property(e => e.RazaoSocial).HasColumnName("CLIENTE");

            builder.Property(e => e.Endereco).HasColumnName("ENDERECO");

            builder.Property(e => e.Complemento).HasColumnName("COMPLEMENTO");

            builder.Property(e => e.Bairro).HasColumnName("BAIRRO");

            builder.Property(e => e.CodCidade).HasColumnName("CODCIDADE");

            builder.Property(e => e.Cidade).HasColumnName("CIDADE");

            builder.Property(e => e.UF).HasColumnName("UF");

            builder.Property(e => e.Pais).HasColumnName("PAIS");

            builder.Property(e => e.CEP).HasColumnName("CEP");

            builder.Property(e => e.Naturalidade).HasColumnName("NATURALIDADE");

            builder.Property(e => e.TipoCliente).HasColumnName("TIPOCLIENTE");

            builder.Property(e => e.RG).HasColumnName("RG");

            builder.Property(e => e.CPF).HasColumnName("CPF");

            builder.Property(e => e.CNPJ).HasColumnName("CNPJ");

            builder.Property(e => e.IE).HasColumnName("IE");

            builder.Property(e => e.IM).HasColumnName("IM");

            builder.Property(e => e.DataNascimento).HasColumnName("DATANASCIMENTO");

            builder.Property(e => e.DataHoraCadastro).HasColumnName("DATAHORACADASTRO");

            builder.Property(e => e.Pai).HasColumnName("PAI");

            builder.Property(e => e.Mae).HasColumnName("MAE");

            builder.Property(e => e.Telefone).HasColumnName("TELEFONE");

            builder.Property(e => e.Celular).HasColumnName("CELULAR");

            builder.Property(e => e.Email).HasColumnName("EMAIL");

            builder.Property(e => e.EstadoCivil).HasColumnName("ESTADOCIVIL");

            builder.Property(e => e.Ativo).HasColumnName("ATIVO");

            builder.Property(e => e.Fantasia).HasColumnName("FANTASIA");

            builder.Property(e => e.NomeConjuge).HasColumnName("NOMECONJUGE");

            builder.Property(e => e.DataUltimaVenda).HasColumnName("DATAULTIMAVENDA");

            builder.Property(e => e.DiasSemComprar).HasColumnName("DIASSEMCOMPRAR");

            builder.Property(e => e.CodConvenio).HasColumnName("CODCONVENIO");

            builder.Property(e => e.Convenio).HasColumnName("CONVENIO");

            builder.Property(e => e.Profissao).HasColumnName("PROFISSAO");

            builder.Property(e => e.EmpresaQueTrabalha).HasColumnName("EMPRESAQUETRABALHA");

            builder.Property(e => e.TelefoneTrabalho).HasColumnName("FONETRABALHO");

            builder.Property(e => e.RendaMensal).HasColumnName("RENDAMENSAL");

            builder.Property(e => e.TotalVendido).HasColumnName("TOTALVENDIDO");

            builder.Property(e => e.Nacionalidade).HasColumnName("NACIONALIDADE");

            builder.Property(e => e.Numero).HasColumnName("NUMERO");

            builder.Property(e => e.CodConsultasSPC).HasColumnName("CODCONSULTASPC");

            builder.Property(e => e.Sexo).HasColumnName("SEXO");

            builder.Property(e => e.CodigoCidadeIBGE).HasColumnName("CODIGOCIDADEIBGE");

            builder.Property(e => e.CodEmitente).HasColumnName("CODEMITENTE");

            builder.Property(e => e.NomeContatoJuridica).HasColumnName("NOMECONTATOJURIDICA");

            builder.Property(e => e.Email2).HasColumnName("EMAIL2");

            builder.Property(e => e.Status).HasColumnName("STATUS");

            builder.Property(e => e.MD5).HasColumnName("MD5L");

            builder.Property(e => e.Serial).HasColumnName("SERIAL");

            builder.Property(e => e.CodGrupo).HasColumnName("CODGRUPO");

            builder.Property(e => e.Grupo).HasColumnName("GRUPO");

            builder.Property(e => e.Tributacao).HasColumnName("TRIBUTACAO");

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

            builder.Property(e => e.ISSMunicipio).HasColumnName("ISSMUNICIPIO");

            builder.Property(e => e.Site).HasColumnName("SITE");

            builder.Property(e => e.CodigoPais).HasColumnName("CODIGOPAIS");

            builder.Property(e => e.CSOSN).HasColumnName("CSOSN");

            builder.Property(e => e.PercICMSProprioST).HasColumnName("PERCICMSPROPRIOST");

            builder.Property(e => e.PercMVAOriginal).HasColumnName("PERCMVAORIGINAL");

            builder.Property(e => e.PercICMSSTInterna).HasColumnName("PERCICMSSTINTERNA");

            builder.Property(e => e.PercReducaoBCST).HasColumnName("PERCREDUCAOBCST");

            builder.Property(e => e.DescricaoCSOSN).HasColumnName("DESCRICAOCSOSN");

            builder.Property(e => e.CodigoCSTOrigem).HasColumnName("CODIGOCSTORIGEM");

            builder.Property(e => e.Origem).HasColumnName("ORIGEM");

            builder.Property(e => e.PossuiICMSST).HasColumnName("POSSUIICMSST");

            builder.Property(e => e.Isento).HasColumnName("ISENTO");

            builder.Property(e => e.Tributado).HasColumnName("TRIBUTADO");

            builder.Property(e => e.CodcstOrigem).HasColumnName("CODCSTORIGEM");

            builder.Property(e => e.MensagemFiscal).HasColumnName("MENSAGEMFISCAL");

            builder.Property(e => e.CodTabelaPreco).HasColumnName("CODTABELAPRECO");

            builder.Property(e => e.TabelaPreco).HasColumnName("TABELAPRECO");

            builder.Property(e => e.CodVendedor).HasColumnName("CODVENDEDOR");

            builder.Property(e => e.Vendedor).HasColumnName("VENDEDOR");

            builder.Property(e => e.LimiteCredito).HasColumnName("LIMITECREDITO");

            builder.Property(e => e.CodParcelaPredefinida).HasColumnName("CODPARCELAPREDEFINIDA");

            builder.Property(e => e.ParcelaPredefinida).HasColumnName("PARCELAPREDEFINIDA");

            builder.Property(e => e.CodTransportadora).HasColumnName("CODTRANSPORTADORA");

            builder.Property(e => e.Transportadora).HasColumnName("TRANSPORTADORA");

            builder.Property(e => e.ValorFrete).HasColumnName("VALORFRETE");

            builder.Property(e => e.MD5Original).HasColumnName("MD5O");

            builder.Property(e => e.Suframa).HasColumnName("SUFRAMA");

            builder.Property(e => e.PercFCP).HasColumnName("PERCFCP");

            builder.Property(e => e.PercReducaoBC).HasColumnName("PERCREDUCAOBC");

            builder.Property(e => e.CSTCSOSNaplicar).HasColumnName("CSTCSOSNAPLICAR");

            builder.Property(e => e.Sincronizado).HasColumnName("SINCRONIZADO");

            builder.Property(e => e.UtilizarCashback).HasColumnName("UTILIZARCASHBACK");


        }
    }
}
