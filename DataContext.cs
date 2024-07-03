using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection.Emit;
using System.Runtime.Remoting.Contexts;
using System.Text;
using System.Threading.Tasks;


namespace Fiscal.Classe {
    public class DataContext {
        public class ConnectionParams {
            public string ConnectionString() {
                return $"DataSource=localhost;Database=C:\\SGBR\\Master\\BD\\BASESGMASTER populada.FDB;Port=3045;User=SYSDBA;Password=masterkey;Charset=UTF8;Dialect=3;Connection lifetime=15;PacketSize=8192;ServerType=0;Unicode=True;Max Pool Size=1000";
            }
        }

        public class Contexto : DbContext {
            public DbSet<Fornecedor> fornecedor {
                get; set;
            }
            public DbSet<Cliente> cliente {
                get; set;
            }
            public DbSet<Estoque> Estoque {
                get; set;
            }

            protected override void OnModelCreating(ModelBuilder modelBuilder) {
                base.OnModelCreating(modelBuilder);
                new FornecedorEntityTypeConfiguration().Configure(modelBuilder.Entity<Fornecedor>());
                new ClienteEntityTypeConfiguration().Configure(modelBuilder.Entity<Cliente>());
                new EstoqueEntityTypeConfiguration().Configure(modelBuilder.Entity<Estoque>());
            }

            protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
                => optionsBuilder.UseFirebird($"DataSource=localhost;Database=C:\\SGBR\\Master\\BD\\BASESGMASTER populada.FDB;Port=3050;User=SYSDBA;Password=masterkey;Charset=UTF8;Dialect=3;Connection lifetime=15;PacketSize=8192;ServerType=0;Unicode=True;Max Pool Size=1000");
        }
    }
}