using MLogging = Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using System.Threading;
using System.Windows.Forms;
using Werkr.Common.Configuration;
using Werkr.Common.Configuration.Kestrel;
using WixToolset.Dtf.WindowsInstaller;

// This custom action code came from https://localjoost.github.io/wix-configurable-search-replace-custom/ originally, was combined with logic from
// https://stackoverflow.com/questions/57294132/how-to-update-appsettings-json-from-wix-custom-actions-with-the-values-passed-as and was enhanced as needed.
namespace Werkr.Installers.Wix {
    public class CustomActions {

        /// <summary>
        /// Creates an <see cref="OpenFileDialog"/> to select a certificate file and 
        /// sets the HTTPSINLINECERTFILE.CERTPATH property.
        /// </summary>
        [CustomAction]
        public static ActionResult OpenCertDetails( Session session ) {
            try {
                OpenCertStoreDetailsTable( session );
            } catch (Exception e) {
                session.Log( $"An exception has occurred while opening the OpenCertStoreDetailsTable form. Error: {e.Message}" );
                return ActionResult.Failure;
            }
            return ActionResult.Success;
        }

        /// <summary>
        /// Creates an <see cref="OpenFileDialog"/> to select a certificate file and 
        /// sets the HTTPSINLINECERTFILE.CERTPATH property.
        /// </summary>
        [CustomAction]
        public static ActionResult OpenCertFileCertPath( Session session ) {
            try {
                session["HTTPSINLINECERTFILE.CERTPATH"] = OpenCertificateFile( );
            } catch (Exception e) {
                session.Log( $"An exception has occurred while opening the OpenCertFileCertPath file prompt. Error: {e.Message}" );
                return ActionResult.Failure;
            }
            return ActionResult.Success;
        }

        /// <summary>
        /// Creates an <see cref="OpenFileDialog"/> to select a certificate file and 
        /// sets the HTTPSINLINECERTANDKEYFILE.CERTPATH property.
        /// </summary>
        [CustomAction]
        public static ActionResult OpenCertAndKeyFileCertPath( Session session ) {
            try {
                session["HTTPSINLINECERTANDKEYFILE.CERTPATH"] = OpenCertificateFile( );
            } catch (Exception e) {
                session.Log( $"An exception has occurred while opening the OpenCertAndKeyFileCertPath file prompt. Error: {e.Message}" );
                return ActionResult.Failure;
            }
            return ActionResult.Success;
        }

        /// <summary>
        /// Creates an <see cref="OpenFileDialog"/> to select a certificate key file and 
        /// sets the HTTPSINLINECERTANDKEYFILE.KEYPATH property.
        /// </summary>
        [CustomAction]
        public static ActionResult OpenCertAndKeyFileKeyPath( Session session ) {
            try {
                session["HTTPSINLINECERTANDKEYFILE.KEYPATH"] = OpenKeyFile( );
            } catch (Exception e) {
                session.Log( $"An exception has occurred while opening the CertAndKeyFileKeyPath file prompt. Error: {e.Message}" );
                return ActionResult.Failure;
            }
            return ActionResult.Success;
        }

        /// <summary>
        /// Creates an <see cref="FolderBrowserDialog"/> to select a the SHELLWORKINGDIR property.
        /// </summary>
        [CustomAction]
        public static ActionResult OpenShellDir( Session session ) {
            try {
                session["SHELLWORKINGDIR"] = OpenFolderDialog( "Shell Working Directory:" );
            } catch (Exception e) {
                session.Log( $"An exception has occurred while opening the shell working directory prompt. Error: {e.Message}" );
                return ActionResult.Failure;
            }
            return ActionResult.Success;
        }

        /// <summary>
        /// Combines the installer properties into an <see cref="AppSettings"/> object and sets the CompletedAppSettingsJson installer property.
        /// The CompletedAppSettingsJson property will be pulled into the <see cref="ConfigSaveExec"/> method later.
        /// </summary>
        /// <param name="session">The current Windows installer session.</param>
        [CustomAction]
        public static ActionResult ConvertPropertiesToCompletedAppSettingsJson( Session session ) {
            try {
                session.Log( "Begin ConvertPropertiesToCompletedAppSettingsJson" );

                // Set string variables
                string productName  = GetSessionProperty( session, "ProductName" );
                string allowedHosts = GetSessionProperty( session, "ALLOWEDHOSTS" );

                // Configure Initial Defaults:
                Endpoints defaultEndpoint = new Http(); // The Http default will be replaced by the ConfigureDefaults method.
                Operators ops = new Operators(); // The Agent uses some extra appsettings that are configured in ConfigureDefaults, and shown if showOperators is true.
                bool showOperators = ConfigureDefaults( session, productName, ref ops, ref defaultEndpoint );

                // Build AppSettings based on installer configuration.
                AppSettings installerAppsettings = new AppSettings {
                    Operators    = ops,
                    AllowedHosts = allowedHosts
                };

                ConfigureLogging( session, installerAppsettings );
                ConfigureCertificates( session, installerAppsettings, defaultEndpoint );

                // Convert object to appsettings json
                string jsonString = installerAppsettings.ToString( showOperators );
                // Save appsettings json for later processing in ConfigSaveExec
                session["CompletedAppSettingsJson"] = jsonString;

                session.Log( "End ConvertPropertiesToCompletedAppSettingsJson" );
            } catch (Exception e) {
                session.Log( $"An exception has occurred while converting msi properties to appsettings json. Error: {e.Message}" );
                return ActionResult.Failure;
            }
            return ActionResult.Success;
        }

        /// <summary>
        /// Overwrites the applications default appsettings.json file with the contents of the
        /// "CompletedAppSettingsJson" installer property.
        /// This should be run as a deferred action.
        /// </summary>
        /// <param name="session">The current Windows installer session.</param>
        [CustomAction]
        public static ActionResult ConfigSaveExec( Session session ) {
            try {
                session.Log( "Begin ConfigSaveExec" );

                string completedAppSettingsJson = GetSessionProperty( session, "CompletedAppSettingsJson", true );

                string appSettingsPath = Path.Combine( GetSessionProperty( session, "INSTALLDIRECTORY", true ), "appsettings.json" );
                if (File.Exists( appSettingsPath )) {
                    File.Delete( appSettingsPath );
                    session.Log( $"Deleted initial AppSettings file '{appSettingsPath}'." );
                }

                session.Log( "Building AppSettings file based on installer configuration." );
                File.WriteAllText( appSettingsPath, completedAppSettingsJson );

                session.Log( $"Saved updated appsettings file to '{appSettingsPath}'." );

                session.Log( "End ConfigSaveExec" );
            } catch (Exception e) {
                session.Log( $"An exception has occurred while saving configuration details. Error: {e.Message}" );
                return ActionResult.Failure;
            }
            return ActionResult.Success;
        }

        #region Private Methods

        /// <summary>
        /// Retrieves properties from the current Windows installer <paramref name="session"/>.
        /// </summary>
        /// <param name="session">The current Windows installer session.</param>
        /// <param name="key">The parameter name to retrieve. Parameters are stored in a Key/Value format.</param>
        /// <param name="deferred">If deferred is true then this function will retrieve the property from the CustomActionData property table instead of the default property table.</param>
        /// <returns>The property value if set. Otherwise an empty string.</returns>
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

        /// <summary>
        /// Configures the <paramref name="ops"/> and <paramref name="defaultEndpoint"/> parameters.
        /// The Agent uses the <paramref name="ops"/> settings. They will be configured if <paramref name="productName"/> is "Werkr Agent"
        /// </summary>
        /// <param name="session">The current Windows installer session.</param>
        /// <param name="productName">Configures the <paramref name="ops"/> object if productName is "Werkr Agent"</param>
        /// <param name="ops">Settings used by the "Werkr Agent" application.</param>
        /// <param name="defaultEndpoint">Sets the default endpoint for both the Werkr Agent and Server.</param>
        /// <returns>True if <paramref name="productName"/> is "Werkr Agent". Otherwise False.</returns>
        private static bool ConfigureDefaults( Session session, string productName, ref Operators ops, ref Endpoints defaultEndpoint ) {
            bool showOperators = false;
            if (productName == "Werkr Agent") {
                string shellWorkingDir = GetSessionProperty( session, "SHELLWORKINGDIR" );

                // Parse out non-string operator values.
                string pwshParse  = GetSessionProperty( session, "ENABLEPWSH" );
                string shellParse = GetSessionProperty( session, "ENABLESHELL" );

                // Bool works with true/false or 1/0
                bool enablePwshParse  = bool.TryParse( pwshParse,  out bool enablePwsh );
                bool enableShellParse = bool.TryParse( shellParse, out bool enableShell );

                // Log out operator parse status.
                session.Log( $"Non-String Install 'Operator Parameters' Parse Status: " );
                session.Log( $"ENABLEPWSH  - input: '{pwshParse}'; Parsed: {enablePwshParse}; Value: {enablePwsh}" );
                session.Log( $"ENABLESHELL - input: '{shellParse}'; Parsed: {enableShellParse}; Value: {enableShell}" );

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

            return showOperators;
        }

        /// <summary>
        /// Configures the <see cref="Logging"/> OTLP and LogLevel settings for the provided <paramref name="installerAppsettings"/>.
        /// </summary>
        /// <param name="session">The current Windows installer session.</param>
        /// <param name="installerAppsettings">Will have its <see cref="Logging"/> settings configured.</param>
        private static void ConfigureLogging( Session session, AppSettings installerAppsettings ) {
            string oTelCollectorAddress = GetSessionProperty( session, "TELEMETRYCOLLECTORADDRESS" );

            // Parse out non-string values.
            string enableLogging = GetSessionProperty( session, "ENABLETELEMETRYLOGGING" );
            bool otelLoggingParse = bool.TryParse( enableLogging, out bool enableOTelLogging ); // Bool works with true/false or 1/0

            // Enums should be passed via the cmdline via # instead of name.
            string loglevel1 = GetSessionProperty( session, "LOGLEVEL.DEFAULT" );
            string loglevel2 = GetSessionProperty( session, "LOGLEVEL.LIFETIME" );
            string loglevel3 = GetSessionProperty( session, "LOGLEVEL.ASPNETCORE" );

            bool loglevel1parse = Enum.TryParse( loglevel1, out MLogging.LogLevel defaultLogLevel );
            bool loglevel2parse = Enum.TryParse( loglevel2, out MLogging.LogLevel lifetimeLogLevel );
            bool loglevel3parse = Enum.TryParse( loglevel3, out MLogging.LogLevel aspNetCoreLogLevel );

            // Log out logging parse status.
            session.Log( $"Non-String Install 'Logging Parameters' Parse Status: " );
            session.Log( $"ENABLETELEMETRYLOGGING - input: '{enableLogging}'; Parsed: {otelLoggingParse}; Value: {enableOTelLogging}" );
            session.Log( $"LOGLEVEL.DEFAULT       - input: '{loglevel1}'; Parsed: {loglevel1parse}; Value: {defaultLogLevel}" );
            session.Log( $"LOGLEVEL.LIFETIME      - input: '{loglevel2}'; Parsed: {loglevel2parse}; Value: {lifetimeLogLevel}" );
            session.Log( $"LOGLEVEL.ASPNETCORE    - input: '{loglevel3}'; Parsed: {loglevel3parse}; Value: {aspNetCoreLogLevel}" );

            // Set OTLP Settings:
            installerAppsettings.Logging.OTLP.EnableTelemetry = otelLoggingParse
                ? enableOTelLogging
                : installerAppsettings.Logging.OTLP.EnableTelemetry;
            installerAppsettings.Logging.OTLP.CollectorAddress = oTelCollectorAddress;

            // Set LogLevel Settings:
            installerAppsettings.Logging.LogLevel.Default = loglevel1parse
                ? defaultLogLevel
                : installerAppsettings.Logging.LogLevel.Default;
            installerAppsettings.Logging.LogLevel.MicrosoftHostingLifetime = loglevel2parse
                ? lifetimeLogLevel
                : installerAppsettings.Logging.LogLevel.MicrosoftHostingLifetime;
            installerAppsettings.Logging.LogLevel.MicrosoftAspNetCore = loglevel3parse
                ? aspNetCoreLogLevel
                : installerAppsettings.Logging.LogLevel.MicrosoftAspNetCore;
        }

        /// <summary>
        /// Configures the <see cref="KestrelConfig"/> settings for the provided <paramref name="installerAppsettings"/>. 
        /// </summary>
        /// <param name="session">The current Windows installer session.</param>
        /// <param name="installerAppsettings"></param>
        /// <param name="defaultEndpoint"></param>
        private static void ConfigureCertificates( Session session, AppSettings installerAppsettings, Endpoints defaultEndpoint ) {
            string certificateType              = GetSessionProperty( session, "CertificateType" );
            string httpsInlineCertFileURL       = GetSessionProperty( session, "HTTPSINLINECERTFILE.URL" );
            string httpsInlineCertAndKeyfileURL = GetSessionProperty( session, "HTTPSINLINECERTANDKEYFILE.URL" );
            string httpsInlineCertStoreURL      = GetSessionProperty( session, "HTTPSINLINECERTSTORE.URL" );

            // Kestrel Settings - Endpoints
            List<Endpoints> endpoints = new List<Endpoints>( );
            switch (certificateType) {
                case "CERTFILE":
                    AddHttpsInlineCertFile( session, endpoints, httpsInlineCertFileURL );
                    break;
                case "CERTANDKEYFILE":
                    AddHttpsInlineCertAndKeyFile( session, endpoints, httpsInlineCertAndKeyfileURL );
                    break;
                case "CERTSTORE":
                    AddHttpsInlineCertStore( session, endpoints, httpsInlineCertStoreURL );
                    break;
                default:
                    AddHttpsInlineCertFile( session, endpoints, httpsInlineCertFileURL );
                    AddHttpsInlineCertAndKeyFile( session, endpoints, httpsInlineCertAndKeyfileURL );
                    AddHttpsInlineCertStore( session, endpoints, httpsInlineCertStoreURL );
                    break;
            }

            // Set Kestrel Settings
            if (endpoints.Count == 0) { endpoints.Add( defaultEndpoint ); }
            installerAppsettings.Kestrel = new KestrelConfig( endpoints.ToArray( ) );
        }

        /// <summary>
        /// Adds an <see cref="HttpsInlineCertFile"/> object to the <paramref name="endpoints"/> list if <paramref name="httpsInlineCertFileURL"/> has content.
        /// </summary>
        /// <param name="session">The current Windows installer session.</param>
        /// <param name="endpoints">The list to add the <see cref="HttpsInlineCertFile"/> object to if <paramref name="httpsInlineCertFileURL"/> has content.</param>
        /// <param name="httpsInlineCertFileURL">The certificate url, or an empty string.</param>
        private static void AddHttpsInlineCertFile( Session session, List<Endpoints> endpoints, string httpsInlineCertFileURL ) {
            if (string.IsNullOrWhiteSpace( httpsInlineCertFileURL ) == false) {
                session.Log( "Adding httpsInlineCertFile kestrel endpoint" );
                endpoints.Add(
                    new HttpsInlineCertFile(
                        httpsInlineCertFileURL,
                        GetSessionProperty( session, "HTTPSINLINECERTFILE.CERTPATH" ),
                        GetSessionProperty( session, "HTTPSINLINECERTFILE.PASSWORD" )
                    )
                );
            }
        }

        /// <summary>
        /// Adds an <see cref="HttpsInlineCertAndKeyFile"/> object to the <paramref name="endpoints"/> list if <paramref name="httpsInlineCertAndKeyfileURL"/> has content.
        /// </summary>
        /// <param name="session">The current Windows installer session.</param>
        /// <param name="endpoints">The list to add the <see cref="HttpsInlineCertAndKeyFile"/> object to if <paramref name="httpsInlineCertAndKeyfileURL"/> has content.</param>
        /// <param name="httpsInlineCertAndKeyfileURL">The certificate url, or an empty string.</param>
        private static void AddHttpsInlineCertAndKeyFile( Session session, List<Endpoints> endpoints, string httpsInlineCertAndKeyfileURL ) {
            if (string.IsNullOrWhiteSpace( httpsInlineCertAndKeyfileURL ) == false) {
                session.Log( "Adding httpsInlineCertAndKeyfile kestrel endpoint" );
                endpoints.Add(
                    new HttpsInlineCertAndKeyFile(
                        httpsInlineCertAndKeyfileURL,
                        GetSessionProperty( session, "HTTPSINLINECERTANDKEYFILE.CERTPATH" ),
                        GetSessionProperty( session, "HTTPSINLINECERTANDKEYFILE.PASSWORD" ),
                        GetSessionProperty( session, "HTTPSINLINECERTANDKEYFILE.KEYPATH" )
                    )
                );
            }
        }

        /// <summary>
        /// Adds an <see cref="HttpsInlineCertStore"/> object to the <paramref name="endpoints"/> list if <paramref name="httpsInlineCertStoreURL"/> has content.
        /// </summary>
        /// <param name="session">The current Windows installer session.</param>
        /// <param name="endpoints">The list to add the <see cref="HttpsInlineCertStore"/> object to if <paramref name="httpsInlineCertStoreURL"/> has content.</param>
        /// <param name="httpsInlineCertStoreURL">The certificate url, or an empty string.</param>
        private static void AddHttpsInlineCertStore( Session session, List<Endpoints> endpoints, string httpsInlineCertStoreURL ) {
            if (string.IsNullOrWhiteSpace( httpsInlineCertStoreURL ) == false) {
                session.Log( "Adding httpsInlineCertStore kestrel endpoint" );
                endpoints.Add(
                    new HttpsInlineCertStore(
                        httpsInlineCertStoreURL,
                        GetSessionProperty( session, "HTTPSINLINECERTSTORE.SUBJECT" ),
                        GetSessionProperty( session, "HTTPSINLINECERTSTORE.STORE" ),
                        GetSessionProperty( session, "HTTPSINLINECERTSTORE.LOCATION" )
                    )
                );
            }
        }

        /// <summary>
        /// Creates an <see cref="OpenFileDialog"/> to select a certificate file.
        /// </summary>
        private static string OpenCertificateFile( ) => OpenDialog( "Select a certificate file:", ".pfx", "Certificate File|*.pfx|All files|*.*" );

        /// <summary>
        /// Creates an <see cref="OpenFileDialog"/> to select a certificate key file.
        /// </summary>
        private static string OpenKeyFile( ) => OpenDialog( "Select a certificate key file:", ".key", "Certificate Key File|*.key|All files|*.*" );

        /// <summary>
        /// Creates an <see cref="OpenFileDialog"/>.
        /// </summary>
        private static string OpenDialog( string title, string defaultExt, string filter ) {
            string filePath = string.Empty;
            Thread task = new Thread(() => {
                using (OpenFileDialog openFileDialog = new OpenFileDialog( )) {
                    openFileDialog.InitialDirectory = "c:\\";
                    openFileDialog.Title = title;
                    openFileDialog.DefaultExt = defaultExt;
                    openFileDialog.Filter = filter;
                    openFileDialog.FilterIndex = 0;
                    openFileDialog.CheckFileExists = true;
                    openFileDialog.SupportMultiDottedExtensions = true;
                    openFileDialog.Multiselect = false;
                    if (openFileDialog.ShowDialog( ) == DialogResult.OK) {
                        //Get the path of specified file
                        filePath = openFileDialog.FileName;
                    }
                }
            });
            task.SetApartmentState( ApartmentState.STA );
            task.Start( );
            task.Join( );
            return filePath;
        }

        /// <summary>
        /// Creates an <see cref="FolderBrowserDialog"/> and returns the SelectedPath.
        /// </summary>
        private static string OpenFolderDialog( string description ) {
            string folderPath = string.Empty;
            Thread task = new Thread( ( ) => {
                using (FolderBrowserDialog openFolder = new FolderBrowserDialog( )) {
                    openFolder.RootFolder = Environment.SpecialFolder.MyComputer;
                    openFolder.Description = description;
                    if (openFolder.ShowDialog( ) == DialogResult.OK) {
                        //Get the path of specified file
                        folderPath = openFolder.SelectedPath;
                    }
                }
            } ) {
                IsBackground = false
            };
            task.SetApartmentState( ApartmentState.STA );
            task.Start( );
            task.Join( );
            return folderPath;
        }

        private static IEnumerable<StoreDetails> GetCertificateStoreDetails( Session session ) {
            StoreLocation[] storeLocations = (StoreLocation[])Enum.GetValues( typeof( StoreLocation ) );
            List<StoreDetails> details = new List<StoreDetails>();

            foreach (StoreLocation storeLocation in storeLocations) {
                foreach (StoreName storeName in (StoreName[])Enum.GetValues( typeof( StoreName ) )) {
                    using (X509Store store = new X509Store( storeName, storeLocation )) {
                        try {
                            store.Open( OpenFlags.OpenExistingOnly );
                            details.Add(
                                new StoreDetails( ) {
                                    Name = store.Name,
                                    CertificateCount = store.Certificates.Count,
                                    Location = store.Location
                                }
                            );
                        } catch (CryptographicException e) {
                            // Store Location likely does not exist.
                            session.Log( $"An exception has occurred while retrieving certificate store details. Error: {e.Message}" );
                        }
                    }
                }
            }
            return details;
        }

        private static IEnumerable<CertDetails> GetCertificateDetails( IEnumerable<StoreDetails> storeDetails, Session session ) {
            List<CertDetails> details = new List<CertDetails>();
            foreach (StoreDetails storeDetailItem in storeDetails) {
                string storeName = storeDetailItem.Name;
                string location = Enum.GetName( typeof(StoreLocation), storeDetailItem.Location);
                try {
                    X509Store store = new X509Store(storeName, storeDetailItem.Location);
                    store.Open( OpenFlags.ReadOnly | OpenFlags.OpenExistingOnly );

                    X509Certificate2Collection collection = store.Certificates.Find(X509FindType.FindByTimeValid,DateTime.Now,false);
                    foreach (X509Certificate2 x509 in collection) {
                        CertDetails certDetails = new CertDetails(){
                            Store              = storeName,
                            Location           = location,
                            Subject            = x509.Subject,
                            SimpleName         = x509.GetNameInfo( X509NameType.SimpleName, true ),
                            SignatureAlgorithm = x509.SignatureAlgorithm.FriendlyName,
                        };
                        details.Add( certDetails );
                        x509.Reset( );
                    }
                    store.Close( );
                } catch (Exception ex) {
                    session.Log(
                        "An error occurred while populating details about the Certificate Store.\r\n"
                        + $"ExceptionType: {ex.GetType( ).FullName}\r\n"
                        + $"Store: {storeName}, Location: {location}\r\n"
                        + $"Message: {ex.Message}\r\n"
                        + $"StackTrace: {ex.StackTrace}"
                    );
                }
            }
            return details;
        }

        private static DataGridView CreateCertificateTable( IEnumerable<CertDetails> certDetails ) =>
            new DataGridView {
                Name = "Certificate Details",
                DataSource = certDetails,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.Single,
                CellBorderStyle = DataGridViewCellBorderStyle.Single,
                GridColor = Color.Black,
                Dock = DockStyle.Fill,
                AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells,
                AutoGenerateColumns = true,
                AllowUserToOrderColumns = true,
                ColumnHeadersVisible = true,
                ReadOnly = true,
                MultiSelect = false,
            };

        private static void SaveCertificateStoreDetails( Session session, DataGridView certGrid ) {
            using (Form form = new Form( )) {
                Panel panel = CreateAcceptPanel( form );
                form.Controls.Add( certGrid );
                form.Controls.Add( panel );
                form.Height = 600;
                form.Width = 1400;

                if (form.ShowDialog( ) == DialogResult.OK) {
                    int selectedRowCount = certGrid.Rows.GetRowCount(DataGridViewElementStates.Selected);
                    if (selectedRowCount > 0) {
                        session["HTTPSINLINECERTSTORE.SUBJECT"] = certGrid.Rows[0].Cells[3].Value as string ?? string.Empty;
                        session["HTTPSINLINECERTSTORE.STORE"] = certGrid.Rows[0].Cells[0].Value as string ?? string.Empty;
                        session["HTTPSINLINECERTSTORE.LOCATION"] = certGrid.Rows[0].Cells[1].Value as string ?? string.Empty;
                    }
                }
            }
        }

        private static Panel CreateAcceptPanel( Form form ) {
            Button okButton = new Button() { Text = "OK" };
            okButton.Click += ( sender, e ) => { form.DialogResult = DialogResult.OK; };
            okButton.Dock = DockStyle.Right;

            Button cancelButton = new Button() { Text = "Cancel" };
            cancelButton.Click += ( sender, e ) => { form.DialogResult = DialogResult.Cancel; };
            cancelButton.Dock = DockStyle.Right;

            Panel panel = new Panel();
            panel.Controls.Add( okButton );
            panel.Controls.Add( cancelButton );
            panel.Height = 30;
            panel.Dock = DockStyle.Bottom;
            panel.BorderStyle = BorderStyle.FixedSingle;
            return panel;
        }

        private static void OpenCertStoreDetailsTable( Session session ) {
            Thread task = new Thread(() => {
                IEnumerable<StoreDetails> storeDetails = GetCertificateStoreDetails( session );
                IEnumerable<CertDetails> certDetails  = GetCertificateDetails( storeDetails, session );
                DataGridView certGrid = CreateCertificateTable( certDetails );
                SaveCertificateStoreDetails(session, certGrid );
            });
            task.SetApartmentState( ApartmentState.STA );
            task.Start( );
            task.Join( );
        }

        #endregion Private Methods

        private class CertDetails {
            public string Store { get; set; }
            public string Location { get; set; }
            public string SimpleName { get; set; }
            public string Subject { get; set; }
            public string SignatureAlgorithm { get; set; }
        }

        private class StoreDetails {
            public string Name { get; set; }
            public int CertificateCount { get; set; }
            public StoreLocation Location { get; set; }
        }
    }
}
