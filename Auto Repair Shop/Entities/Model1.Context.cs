﻿//------------------------------------------------------------------------------
// <auto-generated>
//     Этот код создан по шаблону.
//
//     Изменения, вносимые в этот файл вручную, могут привести к непредвиденной работе приложения.
//     Изменения, вносимые в этот файл вручную, будут перезаписаны при повторном создании кода.
// </auto-generated>
//------------------------------------------------------------------------------

namespace Auto_Repair_Shop.Entities { 

    using System.Data.Entity;
    using System.Data.Entity.Infrastructure;
    
    /// <summary>
    /// Класс, содержащий контекст данных для работы с данными.
    /// </summary>
    public partial class DBEntities : DbContext {

        /// <summary>
        /// Свойство, содержащее контекст данных для работы с данными.
        /// </summary>
        public static DBEntities Instance { get; private set; } = new DBEntities();

        /// <summary>
        /// Конструктор класса.
        /// <br/>
        /// Он создан приватным, чтобы обеспечить взаимодействие с контекстом данных только через единый интерфейс взаимодействия.
        /// То есть, через свойство "Instance".
        /// </summary>
        private DBEntities() : base("name=DBEntities") {

        }
    
        protected override void OnModelCreating(DbModelBuilder modelBuilder) {
            throw new UnintentionalCodeFirstException();
        }
    
        public virtual DbSet<Part> Parts { get; set; }
        public virtual DbSet<Parts_To_Request> Parts_To_Request { get; set; }
        public virtual DbSet<Person> People { get; set; }
        public virtual DbSet<Service_Request> Service_Request { get; set; }
        public virtual DbSet<Service_Type> Service_Type { get; set; }
        public virtual DbSet<Vehicle> Vehicles { get; set; }
        public virtual DbSet<Vehicle_Brand> Vehicle_Brand { get; set; }
    }
}
