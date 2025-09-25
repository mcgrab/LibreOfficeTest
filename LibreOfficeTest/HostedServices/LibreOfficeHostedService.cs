namespace LibreOfficeTest.HostedServices
{
    using Microsoft.Extensions.Hosting;
    using System.Diagnostics;

    public class LibreOfficeHostedService : IHostedService
    {
        private Process? _process;

        public Task StartAsync(CancellationToken cancellationToken)
        {
            // Detect correct soffice path
            string sofficePath;
            if (OperatingSystem.IsWindows())
            {
                sofficePath = @"C:\Program Files\LibreOffice\program\soffice.exe";
            }
            else
            {
                // In Linux containers soffice is usually symlinked into /usr/bin by our Dockerfile
                sofficePath = DetectSofficePath();
            }

            var psi = new ProcessStartInfo
            {
                FileName = sofficePath,
                Arguments = "--headless --accept=\"socket,host=localhost,port=2002;urp;StarOffice.ServiceManager\" " +
                            "--nologo --nodefault --nofirststartwizard",
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };

            _process = Process.Start(psi);

            if (_process == null)
                throw new InvalidOperationException($"Failed to start LibreOffice at: {sofficePath}");

            return Task.CompletedTask;
        }

        public Task StopAsync(CancellationToken cancellationToken)
        {
            try
            {
                if (_process != null && !_process.HasExited)
                {
                    _process.Kill(true);
                }
            }
            catch { }
            return Task.CompletedTask;
        }

        private string DetectSofficePath()
        {
            if (System.IO.File.Exists("/usr/bin/soffice"))
                return "/usr/bin/soffice";

            if (System.IO.File.Exists("/usr/bin/libreoffice"))
                return "/usr/bin/libreoffice";

            // Fallback: rely on PATH
            return "soffice";
        }
    }
}
