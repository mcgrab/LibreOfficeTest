using Microsoft.Extensions.Diagnostics.HealthChecks;
using System.Diagnostics;

namespace LibreOfficeTest.HealthChecks
{
    public class LibreOfficeUnoHealthCheck : IHealthCheck
    {
        public async Task<HealthCheckResult> CheckHealthAsync(
            HealthCheckContext context,
            CancellationToken cancellationToken = default)
        {
            try
            {
                var psi = new ProcessStartInfo
                {
                    FileName = "python3",
                    Arguments = "/app/python/uno_healthcheck.py",
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    UseShellExecute = false,
                    CreateNoWindow = true
                };

                using var proc = Process.Start(psi);
                string stdout = await proc.StandardOutput.ReadToEndAsync();
                string stderr = await proc.StandardError.ReadToEndAsync();
                proc.WaitForExit(5000);

                return proc.ExitCode == 0
                    ? HealthCheckResult.Healthy("UNO is responsive")
                    : HealthCheckResult.Unhealthy($"UNO failed: {stdout} {stderr}");
            }
            catch (Exception ex)
            {
                return HealthCheckResult.Unhealthy("LibreOffice UNO error", ex);
            }
        }
    }

}
