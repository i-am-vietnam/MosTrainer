using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MosTrainer.Core.Models
{
    public class ProjectPackage
    {
        //Vua them
        public string DisplayName
        {
            get
            {
                if (Meta != null && !string.IsNullOrEmpty(Meta.ProjectId))
                {
                    var id = Meta.ProjectId;

                    if (id.StartsWith("Excel2016_P") || id.StartsWith("Excel2019_P"))
                    {
                        var n = id.Substring(id.LastIndexOf("_P") + 2);
                        int num;
                        if (int.TryParse(n, out num))
                            return "Project " + num;
                    }

                    return id;
                }

                return "Project";
            }
        }
        //vua them doan tren
        public ProjectMeta Meta { get; set; } = new ProjectMeta();
        public List<TaskDefinition> Tasks { get; set; } = new List<TaskDefinition>();
        public Dictionary<string, string> Lang { get; set; } = new Dictionary<string, string>();
        public string ProjectFolderPath { get; set; } = "";
        
    }
}
