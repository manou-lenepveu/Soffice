using Microsoft.AspNetCore.Mvc;
using System.Diagnostics;
using System.IO;
using System.Threading.Tasks;

namespace Soffice.Controllers
{
    public class ConvertController : Controller
    {
        [HttpGet("/convert")]
        public IActionResult Index()
        {
            return View();
        }

        [HttpPost("/convert")]
        public async Task<IActionResult> ConvertFile(IFormFile file)
        {
            if (file == null || file.Length == 0)
                return BadRequest("Aucun fichier envoyé.");

            // Enregistrement temporaire du fichier source
            // Path.GetTempFileName crée un fichier vide, on ajoute l'extension du fichier original
            // On utilise Path.GetExtension pour conserver l'extension correcte
            string tempInput = Path.GetTempFileName() + Path.GetExtension(file.FileName);
            using (var stream = System.IO.File.Create(tempInput)) // Crée le fichier, donc c'est un fichier vide
            {
                await file.CopyToAsync(stream);  // Copie le contenu du fichier uploadé dans le fichier temporaire
            }

            string workingDir = Path.GetDirectoryName(tempInput)!;  // Dossier de travail pour les fichiers temporaires
            string sofficePath = "soffice"; // ou le chemin complet, /usr/bin/soffice

            // Étape 1 : XLSX → ODS
            string tempOds = Path.ChangeExtension(tempInput, ".ods");
            await RunSoffice(sofficePath, tempInput, workingDir, "ods");

            // Étape 2 : ODS → XLSX
            string finalXlsx = Path.ChangeExtension(tempInput, ".converted.xlsx");
            await RunSoffice(sofficePath, tempOds, workingDir, "xlsx");

            // Lecture du fichier final
            string convertedFile = Path.ChangeExtension(tempOds, ".xlsx");
            if (!System.IO.File.Exists(convertedFile))
                return BadRequest("Fichier final introuvable.");


            // Lire le fichier et le retourner
            byte[] data = await System.IO.File.ReadAllBytesAsync(convertedFile);
            string fileName = Path.GetFileNameWithoutExtension(file.FileName) + "_reconverti.xlsx";

            // Nettoyage
            System.IO.File.Delete(tempInput);
            System.IO.File.Delete(tempOds);
            System.IO.File.Delete(convertedFile);

            // Retourner le fichier converti
            return File(data, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
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
    }
    
}