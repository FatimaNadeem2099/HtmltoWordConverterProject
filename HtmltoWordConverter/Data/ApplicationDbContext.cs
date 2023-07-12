using HtmltoWordConverter.Models;
using Microsoft.AspNetCore.Identity.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore;
using System.Collections.Generic;

namespace HtmltoWordConverter.Data
{
   
        public class ApplicationDbContext : IdentityDbContext
        {
            public ApplicationDbContext(DbContextOptions<ApplicationDbContext> options)
                : base(options)
            {
            }
            public DbSet<Admin> Admins { get; set; }

            public DbSet<Companies> Company { get; set; }

            public DbSet<Documents_Types> Documents_Type { get; set; }
            public DbSet<Email_Template> Email_Templates { get; set; }

            public DbSet<Employees> Employee { get; set; }

            public DbSet<Failed_Jobs> Failed_Job { get; set; }
            public DbSet<Guidelines> Guideline { get; set; }
            public DbSet<Kind_Of_Documents> Kind_Of_Document { get; set; }
            public DbSet<Kind_Of_Products> Kind_Of_Product { get; set; }
            public DbSet<languages> language { get; set; }
            public DbSet<migrations> migration { get; set; }
            public DbSet<Observation> Observations { get; set; }
            public DbSet<ObservationsHistory> ObservationsHistories { get; set; }
            public DbSet<PasswordReset> PasswordResets { get; set; }
            public DbSet<PersonalAccessToken> PersonalAccessTokens { get; set; }
            public DbSet<ProductClassification> ProductClassifications { get; set; }
            public DbSet<ProductDocument> ProductDocuments { get; set; }
            public DbSet<Project> Projects { get; set; }
            public DbSet<ProjectAudit> ProjectAudits { get; set; }
            public DbSet<ProjectCertificate> ProjectCertificates { get; set; }
            public DbSet<ProjectsDocuments> ProjectsDocument { get; set; }
            public DbSet<ProjectSetting> ProjectSettings { get; set; }
            public DbSet<ProjectsHistory> ProjectsHistories { get; set; }
            public DbSet<Scholars> Scholar { get; set; }
            public DbSet<Sector> Sectors { get; set; }
            public DbSet<SessionAdmin> SessionAdmins { get; set; }
            public DbSet<SessionUsers> SessionUser { get; set; }
            public DbSet<ShariyahCompanies> ShariyahCompany { get; set; }
            public DbSet<Signatures> Signature { get; set; }
            public DbSet<StatusOfFunds> StatusOfFund { get; set; }
            public DbSet<TypeClassification> TypeClassifications { get; set; }
            public DbSet<TypeOfFunds> TypeOfFund { get; set; }
            public DbSet<Users> User { get; set; }
            public DbSet<UsersSectors> UsersSectors { get; set; }


    }
    }

