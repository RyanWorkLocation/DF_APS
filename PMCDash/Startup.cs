using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using Microsoft.AspNetCore.SpaServices.Webpack;
using NSwag;
using NSwag.Generation.Processors.Security;
using PMCDash.Repos;
using PMCDash.Services;
using PMCDash.Helper;
using PMCDash.Formatter;
using Microsoft.AspNetCore.Authentication.JwtBearer;
using Microsoft.IdentityModel.Tokens;
using System.Text;
using System.IO;
using Microsoft.Extensions.FileProviders;
namespace PMCDash
{
    public class Startup
    {

        public Startup(IConfiguration configuration)
        {
            Configuration = configuration;
        }

        public IConfiguration Configuration { get; }


        // This method gets called by the runtime. Use this method to add services to the container.
        public void ConfigureServices(IServiceCollection services)
        {
            services.AddAuthentication(JwtBearerDefaults.AuthenticationScheme)
                    .AddJwtBearer(options =>
                    {
                        // �����ҥ��ѮɡA�^�����Y�|�]�t WWW-Authenticate ���Y�A�o�̷|��ܥ��Ѫ��Բӿ��~��]
                        options.IncludeErrorDetails = true; // �w�]�Ȭ� true�A���ɷ|�S�O����

                        options.TokenValidationParameters = new TokenValidationParameters
                        {
                            // �z�L�o���ŧi�A�N�i�H�q "sub" ���Ȩó]�w�� User.Identity.Name
                            NameClaimType = "http://schemas.xmlsoap.org/ws/2005/05/identity/claims/nameidentifier",
                            // �z�L�o���ŧi�A�N�i�H�q "roles" ���ȡA�åi�� [Authorize] �P�_����
                            RoleClaimType = "http://schemas.microsoft.com/ws/2008/06/identity/claims/role",

                            // �@��ڭ̳��|���� Issuer
                            ValidateIssuer = true,
                            ValidIssuer = Configuration.GetValue<string>("JwtBearerSettings:Issuer"),

                            // �q�`���ӻݭn���� Audience
                            ValidateAudience = false,

                            // �@��ڭ̳��|���� Token �����Ĵ���
                            ValidateLifetime = true,

                            // �p�G Token ���]�t key �~�ݭn���ҡA�@�볣�u��ñ���Ӥw
                            ValidateIssuerSigningKey = false,

                            // "1234567890123456" ���ӱq IConfiguration ���o
                            IssuerSigningKey = new SymmetricSecurityKey(Encoding.UTF8.GetBytes(Configuration.GetValue<string>("JwtBearerSettings:SignKey")))
                        };
                    });
            services.AddCors(options =>
                options.AddDefaultPolicy(x => x.AllowAnyOrigin()
                                               .AllowAnyMethod()
                                               .AllowAnyHeader())
            );
            services.AddControllers(options => options.OutputFormatters.Insert(0, new JsonOutputFormatter()));
            services.AddOpenApiDocument(document =>
            {
                document.Title = "PMC API";
                var openApiSecurityScheme = new OpenApiSecurityScheme
                {
                    Type = OpenApiSecuritySchemeType.ApiKey,
                    Name = "Authorization",
                    Description = "請先輸入【bearer】加一個空格，再貼上你的token",
                    In = OpenApiSecurityApiKeyLocation.Header
                };
                document.AddSecurity("JwtBearer",
                                     Enumerable.Empty<string>(),
                                     openApiSecurityScheme);

                document.OperationProcessors
                        .Add(new AspNetCoreOperationSecurityScopeProcessor("JwtBearer"));
            });
            

            services.AddSingleton<DistributionRepo>();

            services.AddScoped<DeviceDistributionService>();
            services.AddScoped<AccountService>();
            services.AddScoped<AlarmService>();
            services.AddSingleton<JwtProviderHelper>();
        }

        // This method gets called by the runtime. Use this method to configure the HTTP request pipeline.
        public void Configure(IApplicationBuilder app, IWebHostEnvironment env)
        {
            if (env.IsDevelopment())
            {
                app.UseDeveloperExceptionPage();
            }
            app.UseHsts();
            app.UseAuthentication();
            app.UseOpenApi(settings =>
            {
                settings.PostProcess = (document, request) => { /*document.Host = @"ab902b19b9e8.ngrok.io";*/ };
            });
            app.UseSwaggerUi3();
            app.UseCors();
            app.UseRouting();
            app.UseStaticFiles();         
            app.UseAuthorization();
            app.UseEndpoints(endpoints =>
            {
                endpoints.MapControllers();
            });
            app.ConfigSpaHost(env, "ClientApp", 8070);
        }       
    }
}
