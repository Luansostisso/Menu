
using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Metadata.Builders;
using Microsoft.EntityFrameworkCore.Storage.ValueConversion.Internal;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Fiscal.Classe {
    public class Fornecedor {
        [DisplayName("Código")]
        public int Controle {
            get; set;
        }

        [DisplayName("Razão social")]
        public string RazaoSocial {
            get; set;
        }

        [DisplayName("Nome fantasia")]
        public string NomeFantasia {
            get; set;
        }

        [DisplayName("CNPJ")]
        public string CNPJ {
            get; set;
        }

        [DisplayName("Endereço")]
        public string Endereco {
            get; set;
        }

        [DisplayName("Bairro")]
        public string Bairro {
            get; set;
        }

        [DisplayName("Cód. cidade")]
        public int? CodCidade {
            get; set;
        }

        [DisplayName("Cód. cidade IBGE")]
        public int? CodigoCidadeIBGE {
            get; set;
        }

        [DisplayName("Cidade")]
        public string Cidade {
            get; set;
        }

        [DisplayName("Cód. país")]
        public string CodigoPais {
            get; set;
        }

        [DisplayName("UF")]
        public string UF {
            get; set;
        }

        [DisplayName("CEP")]
        public string CEP {
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

        [DisplayName("IE")]
        public string IE {
            get; set;
        }

        [DisplayName("IM")]
        public string IM {
            get; set;
        }

        [DisplayName("Complemento")]
        public string Complemento {
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

        [DisplayName("SAC")]
        public string SAC {
            get; set;
        }

        [DisplayName("E-mail")]
        public string Email {
            get; set;
        }

        [DisplayName("Site")]
        public string Site {
            get; set;
        }

        [DisplayName("Obs")]
        public string Obs {
            get; set;
        }

        [DisplayName("Forma de pagamento")]
        public string FormaPagamento {
            get; set;
        }

        [DisplayName("País")]
        public string Pais {
            get; set;
        }

        [DisplayName("Número")]
        public string Numero {
            get; set;
        }

        [DisplayName("Contato")]
        public string NomeContatoJuridica {
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

        [DisplayName("#")]
        public string Marcado {
            get; set;
        }

        [DisplayName("Data hora cadastro")]
        public DateTime DataHoraCadastro {
            get; set;
        }

        [DisplayName("Ativo")]
        public string Ativo {
            get; set;
        }

        [DisplayName("Cód. emitente")]
        public int? CodEmitente {
            get; set;
        }

        [DisplayName("Contribuinte IPI")]
        public string ContribuinteIPI {
            get; set;
        }

        [DisplayName("Suframa")]
        public string Suframa {
            get; set;
        }
    }
    public class FornecedorEntityTypeConfiguration : IEntityTypeConfiguration<Fornecedor> {
        public void Configure(EntityTypeBuilder<Fornecedor> builder) {
            builder.ToTable("TFORNECEDOR");

            builder.HasKey(e => e.Controle);

            builder.Property(e => e.Controle).HasColumnName("CONTROLE").IsRequired();

            builder.Property(e => e.DataHoraCadastro).HasColumnName("DATAHORACADASTRO").IsRequired();

            builder.Property(e => e.RazaoSocial).HasColumnName("RAZAOSOCIAL").IsRequired();

            builder.Property(e => e.NomeFantasia).HasColumnName("NOMEFANTASIA");

            builder.Property(e => e.Endereco).HasColumnName("ENDERECO");

            builder.Property(e => e.Bairro).HasColumnName("BAIRRO");

            builder.Property(e => e.CodCidade).HasColumnName("CODCIDADE");

            builder.Property(e => e.CodigoCidadeIBGE).HasColumnName("CODIGOCIDADEIBGE");

            builder.Property(e => e.Cidade).HasColumnName("CIDADE");

            builder.Property(e => e.CodigoPais).HasColumnName("CODIGOPAIS");

            builder.Property(e => e.UF).HasColumnName("UF");

            builder.Property(e => e.CEP).HasColumnName("CEP");

            builder.Property(e => e.RG).HasColumnName("RG");

            builder.Property(c => c.CPF).HasColumnName("CPF").HasMaxLength(20);

            builder.Property(c => c.CNPJ).HasColumnName("CNPJ").HasMaxLength(20);

            builder.Property(e => e.IE).HasColumnName("IE");

            builder.Property(e => e.IM).HasColumnName("IM");

            builder.Property(e => e.Complemento).HasColumnName("COMPLEMENTO");

            builder.Property(e => e.Telefone).HasColumnName("TELEFONE");

            builder.Property(e => e.Celular).HasColumnName("CELULAR");

            builder.Property(e => e.SAC).HasColumnName("SAC");

            builder.Property(e => e.Email).HasColumnName("EMAIL");

            builder.Property(e => e.Site).HasColumnName("SITE");

            builder.Property(e => e.Obs).HasColumnName("OBS");

            builder.Property(e => e.FormaPagamento).HasColumnName("FORMAPAGAMENTO");

            builder.Property(e => e.Pais).HasColumnName("PAIS");

            builder.Property(e => e.Numero).HasColumnName("NUMERO");

            builder.Property(e => e.NomeContatoJuridica).HasColumnName("NOMECONTATOJURIDICA");

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

            builder.Property(e => e.Marcado).HasColumnName("MARCADO");

            builder.Property(e => e.Ativo).HasColumnName("ATIVO");

            builder.Property(e => e.CodEmitente).HasColumnName("CODEMITENTE");

            builder.Property(e => e.ContribuinteIPI).HasColumnName("CONTRIBUINTEIPI");

            builder.Property(e => e.Suframa).HasColumnName("SUFRAMA");
        }
    }
}
