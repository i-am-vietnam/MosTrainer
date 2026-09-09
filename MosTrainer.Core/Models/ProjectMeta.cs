using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MosTrainer.Core.Models
{
    public class ProjectMeta
    {
        public string ProjectId { get; set; } = "";
        public string OfficeVersion { get; set; } = "";
        public string Version { get; set; } = "";
        public string Starter { get; set; } = "starter.xlsx";
    }
}
