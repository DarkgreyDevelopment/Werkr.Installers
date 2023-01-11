using System;
using System.Collections.Generic;
using System.IO;
using Werkr.Common.Wix.Configuration;
using Werkr.Common.Wix.Configuration.Kestrel;
using WixToolset.Dtf.WindowsInstaller;

// This custom action code came from https://localjoost.github.io/wix-configurable-search-replace-custom/ originally,
// was combined with logic from
// https://stackoverflow.com/questions/57294132/how-to-update-appsettings-json-from-wix-custom-actions-with-the-values-passed-as
// and was enhanced as needed.
namespace Werkr.Installers.Wix {
    public class CustomActions {

        private static string GetSessionProperty( Session session, string key, bool deferred = false ) {
            try {
                string result = deferred ? session.CustomActionData[key] : session[key];
                if (string.IsNullOrEmpty( result )) { session.Log( $"Install key '{key}' is null or empty." ); }
                return result;
            } catch (KeyNotFoundException) {
                session.Log( $"Install key '{key}' not found." );
                return string.Empty;
            }
        }

        [CustomAction]
        public static ActionResult ConvertPropsToSessionString( Session session ) {

            session.Log( "Begin ConvertPropsToSessionString" );

            // Set string variables
            string productName                       = GetSessionProperty( session, "ProductName" );
            string allowedHosts                      = GetSessionProperty( session, "ALLOWEDHOSTS" );
            string oTelCollectorAddress              = GetSessionProperty( session, "TELEMETRYCOLLECTORADDRESS" );
            string httpUrl                           = GetSessionProperty( session, "HTTP.URL" );
            string httpsDefaultCertURL               = GetSessionProperty( session, "HTTPSDEFAULTCERT.URL" );
            string httpsInlineCertFileURL            = GetSessionProperty( session, "HTTPSINLINECERTFILE.URL" );
            string httpsInlineCertFileCERTPATH       = GetSessionProperty( session, "HTTPSINLINECERTFILE.CERTPATH" );
            string httpsInlineCertFilePASSWORD       = GetSessionProperty( session, "HTTPSINLINECERTFILE.PASSWORD" );
            string httpsInlineCertAndKeyfileURL      = GetSessionProperty( session, "HTTPSINLINECERTANDKEYFILE.URL" );
            string httpsInlineCertAndKeyfileCERTPATH = GetSessionProperty( session, "HTTPSINLINECERTANDKEYFILE.CERTPATH" );
            string httpsInlineCertAndKeyfilePASSWORD = GetSessionProperty( session, "HTTPSINLINECERTANDKEYFILE.PASSWORD" );
            string httpsInlineCertAndKeyfileKEYPATH  = GetSessionProperty( session, "HTTPSINLINECERTANDKEYFILE.KEYPATH" );
            string httpsInlineCertStoreURL           = GetSessionProperty( session, "HTTPSINLINECERTSTORE.URL" );
            string httpsInlineCertStoreSUBJECT       = GetSessionProperty( session, "HTTPSINLINECERTSTORE.SUBJECT" );
            string httpsInlineCertStoreSTORE         = GetSessionProperty( session, "HTTPSINLINECERTSTORE.STORE" );
            string httpsInlineCertStoreLOCATION      = GetSessionProperty( session, "HTTPSINLINECERTSTORE.LOCATION" );
            string shellWorkingDir                   = GetSessionProperty( session, "SHELLWORKINGDIR" );

            // Parse out non-string values.
            string loglevel1     = GetSessionProperty( session, "LOGLEVEL.DEFAULT" );
            string loglevel2     = GetSessionProperty( session, "LOGLEVEL.LIFETIME" );
            string loglevel3     = GetSessionProperty( session, "LOGLEVEL.ASPNETCORE" );
            string enableLogging = GetSessionProperty( session, "ENABLETELEMETRYLOGGING" );
            string pwshParse     = GetSessionProperty( session, "ENABLEPWSH" );
            string shellParse    = GetSessionProperty( session, "ENABLESHELL" );

            // Enums should be passed via the cmdline via # instead of name.
            bool loglevel1parse = Enum.TryParse( loglevel1, out Microsoft.Extensions.Logging.LogLevel defaultLogLevel );
            bool loglevel2parse = Enum.TryParse( loglevel2, out Microsoft.Extensions.Logging.LogLevel lifetimeLogLevel );
            bool loglevel3parse = Enum.TryParse( loglevel3, out Microsoft.Extensions.Logging.LogLevel aspNetCoreLogLevel );

            // Bool works with true/false or 1/0
            bool otelLoggingParse = bool.TryParse( enableLogging, out bool enableOTelLogging );
            bool enablePwshParse  = bool.TryParse( pwshParse, out bool enablePwsh );
            bool enableShellParse = bool.TryParse( shellParse, out bool enableShell );

            // Log out parse status.
            session.Log( $"Non-String Install Parameters Parse Status: " );
            session.Log( $"LOGLEVEL.DEFAULT       - input: '{loglevel1}'; Parsed: {loglevel1parse}; Value: {defaultLogLevel}" );
            session.Log( $"LOGLEVEL.LIFETIME      - input: '{loglevel2}'; Parsed: {loglevel2parse}; Value: {lifetimeLogLevel}" );
            session.Log( $"LOGLEVEL.ASPNETCORE    - input: '{loglevel3}'; Parsed: {loglevel3parse}; Value: {aspNetCoreLogLevel}" );
            session.Log( $"ENABLETELEMETRYLOGGING - input: '{enableLogging}'; Parsed: {otelLoggingParse}; Value: {enableOTelLogging}" );
            session.Log( $"ENABLEPWSH             - input: '{pwshParse}'; Parsed: {enablePwshParse}; Value: {enablePwsh}" );
            session.Log( $"ENABLESHELL            - input: '{shellParse}'; Parsed: {enableShellParse}; Value: {enableShell}" );

            // Agent uses some extra appsettings.
            bool showOperators = false;
            Operators ops = new Operators();
            Endpoints defaultEndpoint;
            if (productName == "Werkr Agent") {
                showOperators = true;
                ops.WorkingDirectory = string.IsNullOrWhiteSpace( shellWorkingDir )
                    ? "/"
                    : shellWorkingDir;
                ops.EnablePowerShell = enablePwshParse
                    ? enablePwsh
                    : ops.EnablePowerShell;
                ops.EnableSystemShell = enableShellParse
                    ? enableShell
                    : ops.EnableSystemShell;

                defaultEndpoint = new HttpsInlineCertFile( "https://localhost:12345", string.Empty, string.Empty );
            } else {
                defaultEndpoint = new HttpsInlineCertFile( "https://localhost:5001", string.Empty, string.Empty );
            }

            // Build AppSettings based on installer configuration.
            AppSettings installerAppsettings = new AppSettings {
                Operators = ops,
                AllowedHosts = string.IsNullOrEmpty(allowedHosts) ? "*" : allowedHosts
            };

            // Logging Settings - OTLP:
            installerAppsettings.Logging.OTLP.EnableTelemetry = otelLoggingParse
                ? enableOTelLogging
                : installerAppsettings.Logging.OTLP.EnableTelemetry;
            installerAppsettings.Logging.OTLP.CollectorAddress = oTelCollectorAddress;
            // Logging Settings - LogLevel:
            installerAppsettings.Logging.LogLevel.Default = loglevel1parse
                ? defaultLogLevel
                : installerAppsettings.Logging.LogLevel.Default;
            installerAppsettings.Logging.LogLevel.MicrosoftHostingLifetime = loglevel2parse
                ? lifetimeLogLevel
                : installerAppsettings.Logging.LogLevel.MicrosoftHostingLifetime;
            installerAppsettings.Logging.LogLevel.MicrosoftAspNetCore = loglevel3parse
                ? aspNetCoreLogLevel
                : installerAppsettings.Logging.LogLevel.MicrosoftAspNetCore;

            // Kestrel Settings - Endpoints
            List<Endpoints> endpoints = new List<Endpoints>( );
            // Kestrel Settings - Endpoints - http
            if (string.IsNullOrWhiteSpace( httpUrl ) == false) {
                session.Log( "Adding http kestrel endpoint" );
                endpoints.Add( new Http( httpUrl ) );
            }
            // Kestrel Settings - Endpoints - httpsDefaultCert
            if (string.IsNullOrWhiteSpace( httpsDefaultCertURL ) == false) {
                session.Log( "Adding httpsDefaultCert kestrel endpoint" );
                endpoints.Add( new HttpsDefaultCert( httpsDefaultCertURL ) );
            }
            // Kestrel Settings - Endpoints - httpsInlineCertFile
            if (string.IsNullOrWhiteSpace( httpsInlineCertFileURL ) == false) {
                session.Log( "Adding httpsInlineCertFile kestrel endpoint" );
                endpoints.Add(
                    new HttpsInlineCertFile(
                        httpsInlineCertFileURL,
                        httpsInlineCertFileCERTPATH,
                        httpsInlineCertFilePASSWORD
                    )
                );
            }
            // Kestrel Settings - Endpoints - httpsInlineCertAndKeyfile
            if (string.IsNullOrWhiteSpace( httpsInlineCertAndKeyfileURL ) == false) {
                session.Log( "Adding httpsInlineCertAndKeyfile kestrel endpoint" );
                endpoints.Add(
                    new HttpsInlineCertAndKeyFile(
                        httpsInlineCertAndKeyfileURL,
                        httpsInlineCertAndKeyfileCERTPATH,
                        httpsInlineCertAndKeyfilePASSWORD,
                        httpsInlineCertAndKeyfileKEYPATH
                    )
                );
            }
            // Kestrel Settings - Endpoints - httpsInlineCertStore
            if (string.IsNullOrWhiteSpace( httpsInlineCertStoreURL ) == false) {
                session.Log( "Adding httpsInlineCertStore kestrel endpoint" );
                endpoints.Add(
                    new HttpsInlineCertStore(
                        httpsInlineCertStoreURL,
                        httpsInlineCertStoreSUBJECT,
                        httpsInlineCertStoreSTORE,
                        httpsInlineCertStoreLOCATION
                    )
                );
            }
            // Set Kestrel Settings
            if (endpoints.Count == 0) { endpoints.Add( defaultEndpoint ); }
            installerAppsettings.Kestrel = new KestrelConfig( endpoints.ToArray( ) );

            // Convert object to appsettings json
            string jsonString = installerAppsettings.ToString( showOperators );
            // Save appsettings json for later processing in ConfigSaveExec
            session["CompletedAppSettingsJson"] = jsonString;

            session.Log( "End ConvertPropsToSessionString" );
            return ActionResult.Success;
        }


        [CustomAction]
        public static ActionResult ConfigSaveExec( Session session ) {
            session.Log( "Begin ConfigSaveExec" );

            string completedAppSettingsJson = GetSessionProperty( session, "APPSETTINGS", true );

            string appSettingsPath = Path.Combine( GetSessionProperty( session, "INSTALLFOLDER", true ), "appsettings.json" );
            if (File.Exists( appSettingsPath )) {
                File.Delete( appSettingsPath );
                session.Log( $"Deleted initial AppSettings file '{appSettingsPath}'." );
            }

            session.Log( "Building AppSettings file based on installer configuration." );
            File.WriteAllText( appSettingsPath, completedAppSettingsJson );

            session.Log( $"Saved updated appsettings file to '{appSettingsPath}'." );

            session.Log( "End ConfigSaveExec" );
            return ActionResult.Success;
        }
    }
}
