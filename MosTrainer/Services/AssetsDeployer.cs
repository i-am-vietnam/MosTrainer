using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MosTrainer.Services
{
    public static class AssetsDeployer
    {
        public static string RootTempDir =>
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                "MOS Trainer", "AssetsTemp");

        public static string DeployProjectAssets(string projectId, string projectFolderPath)
        {
            string src = Path.Combine(projectFolderPath, "assets");
            string dest = Path.Combine(RootTempDir, projectId);

            Directory.CreateDirectory(dest);

            if (Directory.Exists(src))
            {
                foreach (var file in Directory.GetFiles(src, "*.*", SearchOption.AllDirectories))
                {
                    string rel = file.Substring(src.Length)
                        .TrimStart(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);

                    string outPath = Path.Combine(dest, rel);
                    var dir = Path.GetDirectoryName(outPath);
                    if (!string.IsNullOrEmpty(dir)) Directory.CreateDirectory(dir);

                    File.Copy(file, outPath, true);
                }
            }

            return dest;
        }

        public static void CleanupAll()
        {
            try
            {
                if (Directory.Exists(RootTempDir))
                    Directory.Delete(RootTempDir, true);
            }
            catch { }
        }
    }
}
