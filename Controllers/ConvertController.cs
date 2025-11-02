using Microsoft.AspNetCore.Mvc;
using System.Diagnostics;
using System.IO;
using System.Threading.Tasks;

namespace Soffice.Controllers
{
    public class ConvertController : Controller
    {

        private const string SofficePath = "soffice";

        [HttpGet("/convert")]
        public IActionResult Index()
        {
            return View();
        }

        [HttpPost("/convert")]
        public async Task<IActionResult> ConvertFlexible(IFormFile file, string format)
        {
            // === Validation ===
            if (file == null || file.Length == 0)
                return BadRequest("Aucun fichier envoyé.");

            if (string.IsNullOrWhiteSpace(format) || !new[] { "ods", "xlsx" }.Contains(format.ToLower()))
                return BadRequest("Format invalide : 'ods' ou 'xlsx'.");

            string inputExt = Path.GetExtension(file.FileName).ToLower();
            if (!new[] { ".xlsx", ".xls", ".ods" }.Contains(inputExt))
                return BadRequest("Format non supporté : .xlsx, .xls ou .ods.");

            string targetFormat = format.ToLower();
            bool isXlsxToXlsx = inputExt == ".xlsx" && targetFormat == "xlsx";

            // === Choix du traitement ===
            return isXlsxToXlsx
                ? await ConvertFile(file)
                : await ConvertGeneric(file, targetFormat);
        }
        
        private async Task<IActionResult> ConvertGeneric(IFormFile file, string targetFormat)
        {
            string inputExt = Path.GetExtension(file.FileName).ToLower();
            string tempInput = Path.GetTempFileName() + inputExt;
            string workingDir = Path.GetDirectoryName(tempInput)!;

            try
            {
                using (var stream = System.IO.File.Create(tempInput))
                    await file.CopyToAsync(stream);

                string expectedOutput = Path.Combine(workingDir, Path.GetFileNameWithoutExtension(tempInput) + $".{targetFormat}");
                await RunSoffice(SofficePath, tempInput, workingDir, targetFormat);

                if (!System.IO.File.Exists(expectedOutput))
                    return BadRequest($"Conversion échouée vers {targetFormat}.");

                byte[] data = await System.IO.File.ReadAllBytesAsync(expectedOutput);
                string fileName = Path.GetFileNameWithoutExtension(file.FileName) + $".{targetFormat}";

                var contentType = targetFormat == "xlsx"
                    ? "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    : "application/vnd.oasis.opendocument.spreadsheet";

                return File(data, contentType, fileName);
            }
            finally
            {
                CleanupFiles(tempInput);
            }
        }
        
        public async Task<IActionResult> ConvertFile(IFormFile file)
        {
            if (file == null || file.Length == 0)
                return BadRequest("Aucun fichier envoyé.");

            string tempInput = Path.GetTempFileName() + Path.GetExtension(file.FileName);
            string workingDir = Path.GetDirectoryName(tempInput)!;

            try
            {
                using (var stream = System.IO.File.Create(tempInput))
                    await file.CopyToAsync(stream);

                // Étape 1 : XLSX → ODS
                string tempOds = Path.Combine(workingDir, Path.GetFileNameWithoutExtension(tempInput) + ".ods");
                await RunSoffice("soffice", tempInput, workingDir, "ods");

                if (!System.IO.File.Exists(tempOds))
                    return BadRequest("Échec : ODS intermédiaire non généré.");

                // Étape 2 : ODS → XLSX
                string finalXlsx = Path.Combine(workingDir, Path.GetFileNameWithoutExtension(tempOds) + ".xlsx");
                await RunSoffice("soffice", tempOds, workingDir, "xlsx");

                if (!System.IO.File.Exists(finalXlsx))
                    return BadRequest("Échec : XLSX final non généré.");

                byte[] data = await System.IO.File.ReadAllBytesAsync(finalXlsx);
                string fileName = Path.GetFileNameWithoutExtension(file.FileName) + "_nettoye.xlsx";

                return File(data, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
            }
            finally
            {
                CleanupFiles(tempInput, 
                    Path.Combine(workingDir, Path.GetFileNameWithoutExtension(tempInput) + ".ods"),
                    Path.Combine(workingDir, Path.GetFileNameWithoutExtension(tempInput) + ".xlsx"));
            }
        }

        private async Task RunSoffice(string sofficePath, string inputFile, string outputDir, string outputFormat)
        {
            ProcessStartInfo psi = new ProcessStartInfo
            {
                FileName = sofficePath,
                Arguments = $"--headless --convert-to {outputFormat} --outdir \"{outputDir}\" \"{inputFile}\"",
                RedirectStandardOutput = true, // Rediriger la sortie standard
                RedirectStandardError = true, //   Rediriger la sortie d'erreur
                UseShellExecute = false,        // Nécessaire pour rediriger les flux
                CreateNoWindow = true       // Ne pas créer de fenêtre
            };

            using var process = new Process { StartInfo = psi };
            process.Start();
            await process.StandardOutput.ReadToEndAsync();  // Lire la sortie standard (utile pour le débogage)
            await process.StandardError.ReadToEndAsync();   // Lire la sortie d'erreur (utile pour le débogage)
            await process.WaitForExitAsync();           // Attendre la fin du processus
        }

        // === Nettoyage ===
        private void CleanupFiles(params string[] files)
        {
            foreach (var file in files)
                try { if (System.IO.File.Exists(file)) System.IO.File.Delete(file); } catch { }
        }
    }
    
}