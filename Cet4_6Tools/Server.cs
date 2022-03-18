namespace Cet4_6Tools
{
    public static class Server
    {
        private static IWebHostEnvironment? _webHostEnvironment;
        internal static void Configure(IWebHostEnvironment webHostEnvironment)
        {
            _webHostEnvironment = webHostEnvironment;
        }

        public static string? WebRootPath => _webHostEnvironment?.WebRootPath;
        public static string? ContentRootPath => _webHostEnvironment?.ContentRootPath;

        public static string MapPath(string path)
        {
            if(string.IsNullOrEmpty(WebRootPath))
            {
                throw new ArgumentNullException(nameof(_webHostEnvironment));
            }
            return Path.Combine(WebRootPath, path.Replace("~/",""));
        }
    }

    public static class WebHostEnvironmentExtensions
    {
        public static IApplicationBuilder UseStaticWebHostEnviroment(this IApplicationBuilder app)
        {
            var webHostEnvironment = app.ApplicationServices.GetRequiredService<IWebHostEnvironment>();
            Server.Configure(webHostEnvironment);
            return app;
        }
    }
}
