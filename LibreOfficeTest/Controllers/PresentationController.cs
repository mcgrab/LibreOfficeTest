using Microsoft.AspNetCore.Mvc;
using System.Diagnostics;
using System.IO.Compression;

namespace LibreOfficeTest.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class PresentationController : ControllerBase
    {
        private readonly string scriptDir = "/app/python";

        // ---------- Split ----------
        [HttpPost("split")]
        public async Task<IActionResult> Split(IFormFile file, [FromQuery] string format = "odp")
        {
            if (file == null || file.Length == 0)
                return BadRequest("No file uploaded.");

            string tempDir = Path.Combine(Path.GetTempPath(), $"split_{Guid.NewGuid()}");
            Directory.CreateDirectory(tempDir);

            string inputPath = Path.Combine(tempDir, file.FileName);
            using (var stream = new FileStream(inputPath, FileMode.Create))
                await file.CopyToAsync(stream);

            string outputDir = Path.Combine(tempDir, "out");
            Directory.CreateDirectory(outputDir);

            string scriptPath = Path.Combine(scriptDir, "split.py");
            string args = $"\"{scriptPath}\" \"{inputPath}\" \"{outputDir}\" --format {format}";
            var result = RunPython("python3", args);

            if (!result.Success)
                return BadRequest(result.Error);

            // Zip all split outputs
            string zipPath = Path.Combine(Path.GetTempPath(), $"split_{Guid.NewGuid()}.zip");
            ZipFile.CreateFromDirectory(outputDir, zipPath);
            var zipBytes = System.IO.File.ReadAllBytes(zipPath);

            // Cleanup
            try { Directory.Delete(tempDir, true); } catch { }
            try { System.IO.File.Delete(zipPath); } catch { }

            return File(zipBytes, "application/zip", "slides_split.zip");
        }

        // ---------- Merge ----------
        [HttpPost("merge")]
        public async Task<IActionResult> Merge([FromForm] List<IFormFile> files, [FromQuery] string format = "odp")
        {
            if (files == null || files.Count < 2)
                return BadRequest("At least 2 files are required.");

            string tempDir = Path.Combine(Path.GetTempPath(), $"merge_{Guid.NewGuid()}");
            Directory.CreateDirectory(tempDir);

            List<string> inputPaths = new();
            foreach (var file in files)
            {
                string path = Path.Combine(tempDir, file.FileName);
                using (var stream = new FileStream(path, FileMode.Create))
                    await file.CopyToAsync(stream);
                inputPaths.Add(path);
            }

            string outputPath = Path.Combine(tempDir, $"merged.{format}");
            string scriptPath = Path.Combine(scriptDir, "merge.py");
            string inputs = string.Join(" ", inputPaths.Select(i => $"\"{i}\""));
            string args = $"\"{scriptPath}\" {inputs} -o \"{outputPath}\" --format {format}";

            var result = RunPython("python3", args);

            if (!result.Success)
                return BadRequest(result.Error);

            if (!System.IO.File.Exists(outputPath))
                return NotFound("Merged file not found.");

            var bytes = System.IO.File.ReadAllBytes(outputPath);
            string contentType = format switch
            {
                "pptx" => "application/vnd.openxmlformats-officedocument.presentationml.presentation",
                "pdf" => "application/pdf",
                _ => "application/vnd.oasis.opendocument.presentation"
            };

            string fileName = $"merged.{format}";
            try { Directory.Delete(tempDir, true); } catch { }

            return File(bytes, contentType, fileName);
        }

        // ---------- Convert ----------
        [HttpPost("convert")]
        public async Task<IActionResult> Convert(IFormFile file, [FromQuery] string format = "pdf")
        {
            if (file == null || file.Length == 0)
                return BadRequest("No file uploaded.");

            string tempDir = Path.Combine(Path.GetTempPath(), $"convert_{Guid.NewGuid()}");
            Directory.CreateDirectory(tempDir);

            string inputPath = Path.Combine(tempDir, file.FileName);
            using (var stream = new FileStream(inputPath, FileMode.Create))
                await file.CopyToAsync(stream);

            string outputPath = Path.Combine(tempDir, $"converted.{format}");
            string scriptPath = Path.Combine(scriptDir, "convert.py");
            string args = $"\"{scriptPath}\" \"{inputPath}\" \"{outputPath}\" {format}";

            var result = RunPython("python3", args);

            if (!result.Success)
                return BadRequest(result.Error);

            if (!System.IO.File.Exists(outputPath))
                return NotFound("Converted file not found.");

            var bytes = System.IO.File.ReadAllBytes(outputPath);
            string contentType = format switch
            {
                "pptx" => "application/vnd.openxmlformats-officedocument.presentationml.presentation",
                "pdf" => "application/pdf",
                _ => "application/vnd.oasis.opendocument.presentation"
            };

            string fileName = $"converted.{format}";
            try { Directory.Delete(tempDir, true); } catch { }

            return File(bytes, contentType, fileName);
        }

        // ---------- Helper ----------
        private (bool Success, string Output, string Error) RunPython(string exe, string args)
        {
            var psi = new ProcessStartInfo
            {
                FileName = exe,
                Arguments = args,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };

            using var proc = Process.Start(psi);
            string output = proc.StandardOutput.ReadToEnd();
            string error = proc.StandardError.ReadToEnd();
            proc.WaitForExit();

            return (proc.ExitCode == 0, output, error);
        }
    }
}
